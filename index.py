import os
import sys
import urllib.request
import win32com.client

def baixar_pdf(url, caminho_destino):
    try:
        urllib.request.urlretrieve(url, caminho_destino)
        print(f"PDF baixado com sucesso: {caminho_destino}")
        return True
    except Exception as e:
        print(f"Erro ao baixar PDF: {e}")
        return False

def converter_pdf_para_docx(caminho_pdf, caminho_docx):
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(caminho_pdf)
        doc.SaveAs2(caminho_docx, FileFormat=16)  # .docx
        doc.Close()
        word.Quit()
        print(f"Conversão concluída: {caminho_docx}")
    except Exception as e:
        print(f"Erro ao converter PDF: {e}")

if __name__ == "__main__":
    url = "https://exemplo.com/meuarquivo.pdf"
    nome_pdf = url.split('/')[-1]
    pasta_destino = os.path.expanduser("~/Downloads/")

    caminho_pdf = os.path.join(pasta_destino, nome_pdf)
    caminho_docx = os.path.join(pasta_destino, nome_pdf.replace(".pdf", ".docx"))

    if baixar_pdf(url, caminho_pdf):
        converter_pdf_para_docx(caminho_pdf, caminho_docx)
