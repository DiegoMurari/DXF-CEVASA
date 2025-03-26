# pdf_converter.py
import os
from xlsx2html import xlsx2html
from weasyprint import HTML

def convert_excel_to_pdf(excel_path, pdf_path, html_path="temp.html"):
    """
    Converte um arquivo Excel para PDF.
    
    Passos:
      1. Converte o arquivo Excel para um arquivo HTML utilizando xlsx2html.
      2. Converte o HTML para PDF usando WeasyPrint.
    
    :param excel_path: Caminho do arquivo Excel (ex.: "Planilha_Final.xlsx").
    :param pdf_path: Caminho do arquivo PDF de saída (ex.: "Planilha_Final.pdf").
    :param html_path: Caminho temporário para salvar o HTML (padrão "temp.html").
    """
    try:
        # Converter o Excel para HTML
        xlsx2html(excel_path, html_path)
        print(f"✅ Excel convertido para HTML: {html_path}")
        
        # Converter o HTML para PDF
        HTML(html_path).write_pdf(pdf_path)
        print(f"✅ PDF gerado com sucesso: {pdf_path}")
        
        # Opcional: remover o arquivo HTML temporário
        os.remove(html_path)
    except Exception as e:
        print(f"❌ Erro na conversão de Excel para PDF: {e}")

# Exemplo de uso quando o módulo é executado diretamente
if __name__ == "__main__":
    excel_file = "Planilha_Final.xlsx"
    pdf_file = "Planilha_Final.pdf"
    convert_excel_to_pdf(excel_file, pdf_file)
