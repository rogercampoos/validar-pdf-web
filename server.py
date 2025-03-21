from flask import Flask, render_template, request, send_file
import pandas as pd
import PyPDF2
import re
import io
import os

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB

def extrair_patrimonios_pdf(file_stream):
    patrimonios = set()
    try:
        pdf = PyPDF2.PdfReader(file_stream)
        for page in pdf.pages:
            text = page.extract_text()
            encontrados = re.findall(r'\b(\d{9})\b', text)
            patrimonios.update(encontrados)
    except Exception as e:
        raise ValueError(f"Erro no PDF: {str(e)}")
    return patrimonios

def extrair_patrimonios_excel(file_stream):
    try:
        df = pd.read_excel(file_stream)
        if len(df.columns) < 2:
            raise ValueError("O arquivo precisa ter pelo menos 2 colunas")
            
        coluna = df.columns[1]
        patrimonios = set(df[coluna].dropna().astype(str))
        return {p for p in patrimonios if re.match(r'^\d{9}$', p)}
    except Exception as e:
        raise ValueError(f"Erro no Excel: {str(e)}")

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            # Processar arquivos
            excel_file = request.files['excel']
            pdf_file = request.files['pdf']
            
            # Extrair dados
            patrimonios_excel = extrair_patrimonios_excel(excel_file.stream)
            patrimonios_pdf = extrair_patrimonios_pdf(pdf_file.stream)
            
            # Gerar resultados
            resultados = {
                'ambos': list(patrimonios_excel & patrimonios_pdf),
                'somente_excel': list(patrimonios_excel - patrimonios_pdf),
                'somente_pdf': list(patrimonios_pdf - patrimonios_excel)
            }
            
            # Gerar Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pd.DataFrame({'Patrimônios em Ambos': resultados['ambos']}).to_excel(writer, sheet_name='Em Ambos', index=False)
                pd.DataFrame({'Somente Excel': resultados['somente_excel']}).to_excel(writer, sheet_name='Somente Excel', index=False)
                pd.DataFrame({'Somente PDF': resultados['somente_pdf']}).to_excel(writer, sheet_name='Somente PDF', index=False)
            
            output.seek(0)
            return send_file(
                output,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                download_name='resultado_validacao.xlsx',
                as_attachment=True
            )
            
        except Exception as e:
            return render_template('index.html', error=str(e))
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)