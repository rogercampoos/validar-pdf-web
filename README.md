# Validador de Patrimônios

Sistema web para validação e comparação de patrimônios entre planilhas Excel/ODS e documentos PDF, permitindo identificar rapidamente divergências entre as fontes e gerar relatórios detalhados.

🌐 Acesso ao Sistema

O sistema está disponível online em:

https://validarpdf2.pythonanywhere.com/

## 🌟 Funcionalidades

- ✅ Upload de arquivos Excel/ODS e PDF
- ✅ Extração automática de números de patrimônio
- ✅ Comparação detalhada entre as fontes
- ✅ Geração de relatório em Excel
- ✅ Interface intuitiva com drag & drop
- ✅ Validação em tempo real
- ✅ Feedback visual do processamento

## 🚀 Tecnologias Utilizadas

- **Frontend:**
  - HTML5
  - CSS3 (Design responsivo)
  - JavaScript (Vanilla JS)
  - Fetch API

- **Backend:**
  - Python
  - Flask
  - Pandas
  - PyPDF2
  - OpenPyXL

## 💻 Funcionalidades Detalhadas

### Processamento de Arquivos
- Suporte para arquivos Excel (.xlsx, .xls) e OpenDocument (.ods)
- Extração de texto de arquivos PDF
- Validação de formato dos números de patrimônio (9 dígitos)

### Análise Comparativa
- Identificação de patrimônios presentes em ambas as fontes
- Detecção de patrimônios exclusivos do Excel
- Detecção de patrimônios exclusivos do PDF
- Geração de estatísticas detalhadas

### Relatório de Resultados
- Planilha Excel com múltiplas abas:
  - Resumo geral
  - Lista de exemplos
  - Patrimônios em ambas as fontes
  - Patrimônios exclusivos do Excel
  - Patrimônios exclusivos do PDF

## 📊 Estrutura do Relatório

O relatório gerado contém as seguintes informações:

1. **Aba Resumo:**
   - Total de patrimônios no Excel
   - Total de patrimônios no PDF
   - Quantidade de patrimônios em comum
   - Quantidade de patrimônios exclusivos

2. **Aba Exemplos:**
   - Amostra dos primeiros patrimônios encontrados apenas no Excel

3. **Abas Detalhadas:**
   - Listagem completa dos patrimônios por categoria

## 🔧 Requisitos do Sistema

- Navegador web moderno
- Arquivos Excel/ODS com números de patrimônio na segunda coluna
- Arquivos PDF com números de patrimônio legíveis
- Limite máximo de arquivo: 32MB

## 🚨 Limitações

- Os números de patrimônio devem ter exatamente 9 dígitos
- A planilha Excel/ODS deve conter pelo menos uma coluna com os números
- O PDF deve ter o texto extraível (não pode ser uma imagem)

## 💡 Dicas de Uso

1. **Preparação dos Arquivos:**
   - Certifique-se que os números de patrimônio estão no formato correto
   - Verifique se o PDF não está protegido ou danificado
   - Organize a planilha com os números de patrimônio preferencialmente na segunda coluna

2. **Upload de Arquivos:**
   - Arraste os arquivos para as áreas designadas ou use os botões de seleção
   - Aguarde o processamento completo
   - Verifique o relatório gerado para análise detalhada

3. **Análise dos Resultados:**
   - Revise o resumo na interface
   - Baixe o relatório Excel para análise completa
   - Verifique os exemplos de divergências

## 🆘 Solução de Problemas

Se encontrar problemas:

1. Verifique se os arquivos estão no formato correto
2. Confirme se os números de patrimônio têm 9 dígitos
3. Certifique-se que o PDF permite extração de texto
4. Verifique se a planilha tem pelo menos uma coluna com dados

## 👥 Suporte

Para suporte e dúvidas:
- Consulte a documentação
- Entre em contato com o administrador do sistema
- Verifique se está usando a versão mais recente

## 📝 Notas de Versão

### Versão 1.3
- Interface responsiva aprimorada
- Melhor feedback visual durante o processamento
- Suporte a arquivos OpenDocument (.ods)
- Relatório Excel mais detalhado
