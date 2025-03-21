from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import PyPDF2
import re
import io
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB

def extrair_patrimonios_pdf(file_stream):
    try:
        # Criar um buffer seekable
        pdf_buffer = io.BytesIO()
        file_stream.save(pdf_buffer)
        pdf_buffer.seek(0)
        
        patrimonios = set()
        pdf = PyPDF2.PdfReader(pdf_buffer)
        for page in pdf.pages:
            text = page.extract_text()
            # Padrão para encontrar números de patrimônio (9 dígitos)
            encontrados = re.findall(r'\b(\d{9})\b', text)
            patrimonios.update(encontrados)
        
        if not patrimonios:
            app.logger.warning("Nenhum patrimônio encontrado no PDF. Verifique se o formato está correto.")
        
        return patrimonios
    except Exception as e:
        app.logger.error(f"Erro ao processar PDF: {str(e)}")
        raise ValueError(f"Erro no PDF: {str(e)}")

def extrair_patrimonios_excel(file_stream):
    try:
        # Converter para buffer seekable
        excel_buffer = io.BytesIO()
        file_stream.save(excel_buffer)
        excel_buffer.seek(0)
        
        # Identificar o engine adequado baseado na extensão
        file_ext = os.path.splitext(file_stream.filename)[1].lower()
        engine = 'odf' if file_ext == '.ods' else 'openpyxl'
        
        try:
            # Tenta ler diretamente especificando o engine
            df = pd.read_excel(excel_buffer, engine=engine)
        except:
            # Fallback: tenta usar o pandas para detectar o formato
            excel_buffer.seek(0)
            df = pd.read_excel(excel_buffer)
        
        # Verifica se o DataFrame está vazio
        if df.empty:
            raise ValueError("A planilha está vazia ou não pôde ser lida corretamente")
            
        if len(df.columns) < 2:
            # Tenta usar a primeira coluna se houver apenas uma
            if len(df.columns) == 1:
                coluna = df.columns[0]
            else:
                raise ValueError("O arquivo precisa ter pelo menos 1 coluna com os números de patrimônio")
        else:
            # Usa a segunda coluna por padrão
            coluna = df.columns[1]
        
        # Log para debug
        app.logger.info(f"Colunas encontradas: {df.columns.tolist()}")
        app.logger.info(f"Usando coluna: {coluna}")
        
        # Converte para string e filtra os patrimônios válidos (9 dígitos)
        df[coluna] = df[coluna].astype(str)
        patrimonios = set(p.strip() for p in df[coluna].dropna() if re.match(r'^\d{9}$', p.strip()))
        
        if not patrimonios:
            app.logger.warning(f"Nenhum patrimônio (formato de 9 dígitos) encontrado na coluna '{coluna}'")
            
        return patrimonios
    except Exception as e:
        app.logger.error(f"Erro ao processar Excel: {str(e)}")
        raise ValueError(f"Erro no Excel: {str(e)}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            # Verifica se os arquivos foram enviados
            if 'excel' not in request.files or 'pdf' not in request.files:
                raise ValueError("Ambos os arquivos (Excel e PDF) são obrigatórios")
                
            excel_file = request.files['excel']
            pdf_file = request.files['pdf']
            
            # Verifica se os arquivos têm nomes válidos
            if excel_file.filename == '' or pdf_file.filename == '':
                raise ValueError("Nenhum arquivo selecionado")
            
            # Processa arquivos para extrair patrimônios
            patrimonios_excel = extrair_patrimonios_excel(excel_file)
            patrimonios_pdf = extrair_patrimonios_pdf(pdf_file)
            
            # Calcular interseções e diferenças
            patrimonios_ambos = patrimonios_excel & patrimonios_pdf
            somente_excel = patrimonios_excel - patrimonios_pdf
            somente_pdf = patrimonios_pdf - patrimonios_excel
            
            # Preparar resultados
            resultados = {
                'ambos': sorted(list(patrimonios_ambos)),
                'somente_excel': sorted(list(somente_excel)),
                'somente_pdf': sorted(list(somente_pdf)),
                'total_excel': len(patrimonios_excel),
                'total_pdf': len(patrimonios_pdf),
                'total_ambos': len(patrimonios_ambos)
            }
            
            # Gerar DataFrames
            df_ambos = pd.DataFrame({'Patrimônios em Ambos': resultados['ambos']}) if resultados['ambos'] else pd.DataFrame({'Info': ['Nenhum patrimônio em comum']})
            df_excel = pd.DataFrame({'Somente no Excel': resultados['somente_excel']}) if resultados['somente_excel'] else pd.DataFrame({'Info': ['Nenhum patrimônio exclusivo']})
            df_pdf = pd.DataFrame({'Somente no PDF': resultados['somente_pdf']}) if resultados['somente_pdf'] else pd.DataFrame({'Info': ['Nenhum patrimônio exclusivo']})
            
            # Criar uma planilha de resumo
            resumo_data = {
                'Descrição': [
                    'Total de patrimônios no Excel',
                    'Total de patrimônios no PDF',
                    'Patrimônios em ambas as fontes',
                    'Patrimônios apenas no Excel',
                    'Patrimônios apenas no PDF'
                ],
                'Quantidade': [
                    len(patrimonios_excel),
                    len(patrimonios_pdf),
                    len(patrimonios_ambos),
                    len(somente_excel),
                    len(somente_pdf)
                ]
            }
            df_resumo = pd.DataFrame(resumo_data)
            
            # Criar DataFrame com exemplos
            exemplos_excel = sorted(list(somente_excel))[:5]  # Primeiros 5 exemplos
            exemplos_data = {
                'Exemplos de patrimônios que estão apenas no Excel': 
                exemplos_excel + [''] * (5 - len(exemplos_excel))  # Preencher até 5 itens
            }
            df_exemplos = pd.DataFrame(exemplos_data)
            
            # Gerar Excel com os resultados
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
                df_exemplos.to_excel(writer, sheet_name='Exemplos', index=False)
                df_ambos.to_excel(writer, sheet_name='Em Ambos', index=False)
                df_excel.to_excel(writer, sheet_name='Somente Excel', index=False)
                df_pdf.to_excel(writer, sheet_name='Somente PDF', index=False)
            
            output.seek(0)
            
            # Retornar também um JSON com o resumo para exibir na interface
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                # Preparar exemplos somente se houver patrimônios apenas no PDF
                exemplos_para_exibir = []
                if len(somente_pdf) > 0:
                    exemplos_para_exibir = sorted(list(somente_excel))[:4]  # Primeiros 4 exemplos
                
                return jsonify({
                    'success': True,
                    'resumo': {
                        'total_excel': len(patrimonios_excel),
                        'total_pdf': len(patrimonios_pdf),
                        'total_ambos': len(patrimonios_ambos),
                        'total_somente_excel': len(somente_excel),
                        'total_somente_pdf': len(somente_pdf),
                        'exemplos_excel': exemplos_para_exibir,
                        'mostrar_exemplos': len(somente_pdf) > 0
                    }
                })
            
            # Retornar o arquivo Excel
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                download_name='resultado_validacao.xlsx',
                as_attachment=True
            )
            
        except Exception as e:
            app.logger.error(f"Erro durante o processamento: {str(e)}")
            
            # Responder com JSON se for uma requisição AJAX
            if request.headers.get('X-Requested-With') == 'XMLHttpRequest':
                return jsonify({
                    'success': False,
                    'erro': str(e)
                })
            
            # Gerar arquivo de erro
            error_df = pd.DataFrame({
                'Erro': [str(e)],
                'Dica': ['Verifique se os arquivos estão no formato correto e possuem números de patrimônio com 9 dígitos.']
            })
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                error_df.to_excel(writer, index=False)
            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                download_name='erro_validacao.xlsx',
                as_attachment=True
            )
    
    return render_template('index.html')

# Para debug local
if __name__ == '__main__':
    # Configuração de log para debug
    import logging
    logging.basicConfig(level=logging.DEBUG)
    app.logger.setLevel(logging.DEBUG)
    
    app.run(debug=True)
