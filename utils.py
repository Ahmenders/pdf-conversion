import os
import fitz  # PyMuPDF
import pdftables_api
from docx import Document
from config import apiKey
from pdf2docx import Converter
from docxcompose.composer import Composer



def count_pdf_pages(pdf_file_path):
    try:
        pdf_document = fitz.open(pdf_file_path)
        num_pages = pdf_document.page_count
        return num_pages
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return None


def getFilename(path):
    file = path.split("/")[-1]
    file = file.split(".")[0]
    return file
    

def pdf2doc(path, outputPath):
    pdf_file = path
    filename = getFilename(path)
    # docx_file = f"{outputPath}/{filename}.docx"

    num_pages = count_pdf_pages(pdf_file)

    for i in range(num_pages):
        docx_file = f'{outputPath}/{filename}{i}.docx'

        # convert pdf to docx
        cv = Converter(pdf_file)
        cv.convert(docx_file, start=i, end= i + 1)      # all pages by default
        cv.close()
    
    master = Document(f"{outputPath}/{filename}{0}.docx")
    composer = Composer(master)
    os.remove(f"{outputPath}/{filename}{0}.docx")
        
    for i in range(1, num_pages):
        doc = Document(f"{outputPath}/{filename}{i}.docx")
        composer.append(doc)
        os.remove(f"{outputPath}/{filename}{i}.docx")
        
    docx_file = f"{outputPath}/{filename}.docx"
    composer.save(docx_file)
    
    return docx_file

def pdf2ppt(path, outputPath):
    os.system(f"pdf2pptx {path}")
    output = path.replace("pdf", "pptx")
    return output


# pip install git+https://github.com/pdftables/python-pdftables-api.git
# pip setup.py install 
def pdf2csv(path, outputPath):
    c = pdftables_api.Client(apiKey)
    filename = getFilename(path)
    c.xlsx(path, f'{outputPath}/{filename}.xlsx') 
    return f'{outputPath}/{filename}.xlsx'

def pdf2html(path, outputPath):
    pass