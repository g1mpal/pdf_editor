from pdf2docx import Converter
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import pdfplumber

def merge_pdfs(file_paths, output_path):
    merger = PdfMerger()
    for pdf in file_paths:
        merger.append(pdf)
    merger.write(output_path)
    merger.close()

    remove_blank_pages(output_path)

def remove_blank_pages(file_path):
    reader = PdfReader(file_path)
    writer = PdfWriter()
    
    for page in reader.pages:
        if not page_is_blank(page):
            writer.add_page(page)
    
    with open(file_path, 'wb') as out_file:
        writer.write(out_file)

def page_is_blank(page):
    return len(page.extract_text().strip()) == 0

def delete_page(input_pdf, page_number, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    
    num_pages = len(reader.pages)
    if page_number < 1 or page_number > num_pages:
        raise ValueError("Invalid page number")
    
    for i in range(num_pages):
        if i != page_number - 1:
            writer.add_page(reader.pages[i])
    
    with open(output_pdf, 'wb') as out_file:
        writer.write(out_file)

def split_pdf(input_pdf, split_page, output_pdf1, output_pdf2):
    reader = PdfReader(input_pdf)
    writer1 = PdfWriter()
    writer2 = PdfWriter()
    
    num_pages = len(reader.pages)
    if split_page < 1 or split_page >= num_pages:
        raise ValueError("Invalid split page number")
    
    for i in range(split_page):
        writer1.add_page(reader.pages[i])
    
    for i in range(split_page, num_pages):
        writer2.add_page(reader.pages[i])
    
    with open(output_pdf1, 'wb') as out_file1:
        writer1.write(out_file1)
        
    with open(output_pdf2, 'wb') as out_file2:
        writer2.write(out_file2)

def rotate_page(input_pdf, page_number, rotation_angle, output_pdf):
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    
    num_pages = len(reader.pages)
    if page_number < 1 or page_number > num_pages:
        raise ValueError("Invalid page number")
    
    for i in range(num_pages):
        page = reader.pages[i]
        if i == page_number - 1:
            page.rotate(rotation_angle)
        writer.add_page(page)
    
    with open(output_pdf, 'wb') as out_file:
        writer.write(out_file)

def pdf_to_word(input_pdf, output_docx):
    cv = Converter(input_pdf)
    cv.convert(output_docx, start=0, end=None, 
               layout=True, 
               font_path=None, 
               output_font_size=None, 
               accuracy=2)  # Higher accuracy for better layout preservation
    cv.close()

    # Post-process the Word document to handle images better
    post_process_word(output_docx)

def post_process_word(docx_path):
    from docx import Document
    from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    doc = Document(docx_path)

    for paragraph in doc.paragraphs:
        # Adjust paragraph formatting
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
        for run in paragraph.runs:
            run.font.size = Pt(12)
            run.font.color.rgb = RGBColor(0, 0, 0)
    
    # Save the post-processed document
    doc.save(docx_path)
