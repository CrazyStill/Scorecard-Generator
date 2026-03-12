import os, csv, shutil, tempfile
from docx import Document
import docx2pdf
from PyPDF2 import PdfMerger
import comtypes.client  
import pythoncom       

def merge_runs_in_paragraph(paragraph):
    if len(paragraph.runs) <= 1:
        return
    full_text = "".join(run.text for run in paragraph.runs)
    for run in paragraph.runs:
        run.text = ""
    paragraph.runs[0].text = full_text

def replace_text_in_paragraph(paragraph, placeholder, replacement):
    if placeholder in paragraph.text:
        new_text = paragraph.text.replace(placeholder, replacement)
        paragraph.runs[0].text = new_text

def replace_placeholders_in_paragraphs(paragraphs, placeholders):
    for paragraph in paragraphs:
        merge_runs_in_paragraph(paragraph)
        for placeholder, replacement in placeholders:
            replace_text_in_paragraph(paragraph, placeholder, replacement)

def replace_text_in_doc(doc, rows, mapping, cards_per_page=4):
    all_placeholders = []
    for i in range(1, cards_per_page + 1):
        row = rows[i - 1] if i - 1 < len(rows) else None
        for csv_header, placeholder in mapping.items():
            value = row.get(csv_header, "") if row else ""
            all_placeholders.append((f"{placeholder}_{i}", value))
    replace_placeholders_in_paragraphs(doc.paragraphs, all_placeholders)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                replace_placeholders_in_paragraphs(cell.paragraphs, all_placeholders)

def merge_two_pdfs(pdf1, pdf2, merged_pdf):
    merger = PdfMerger()
    merger.append(pdf1)
    merger.append(pdf2)
    merger.write(merged_pdf)
    merger.close()

def convert_docx_to_pdf(docx_path, pdf_path):
    pythoncom.CoInitialize()
    try:
        docx2pdf.convert(docx_path, pdf_path)
    except Exception:
        wdFormatPDF = 17
        word = comtypes.client.CreateObject("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        try:
            doc.SaveAs2(pdf_path, FileFormat=wdFormatPDF)
        finally:
            doc.Close()
            word.Quit()
    finally:
        pythoncom.CoUninitialize()

def generate_scorecard(template_path, csv_path, mapping,
                       cards_per_page=4, back_pdf_path=None, temp_dir=None):
    final_pdf_list = []
    with open(csv_path, newline='', encoding='latin-1') as f:
        sample = f.read(1024)
        f.seek(0)
        dialect = csv.Sniffer().sniff(sample)
        reader = csv.DictReader(f, dialect=dialect)
        rows = [row for row in reader if any(value.strip() for value in row.values())]

    for i in range(0, len(rows), cards_per_page):
        group = rows[i : i + cards_per_page]
        front_doc = Document(template_path)
        replace_text_in_doc(front_doc, group, mapping, cards_per_page=cards_per_page)

        idx = i // cards_per_page
        temp_front_docx = os.path.join(temp_dir, f"temp_front_{idx}.docx")
        temp_front_pdf  = os.path.join(temp_dir, f"temp_front_{idx}.pdf")
        page_pdf        = os.path.join(temp_dir, f"page_{idx}.pdf")

        front_doc.save(temp_front_docx)
        convert_docx_to_pdf(temp_front_docx, temp_front_pdf)

        if back_pdf_path and os.path.exists(back_pdf_path):
            merge_two_pdfs(temp_front_pdf, back_pdf_path, page_pdf)
        else:
            shutil.copy(temp_front_pdf, page_pdf)

        final_pdf_list.append(page_pdf)
        os.remove(temp_front_docx)
        os.remove(temp_front_pdf)

    output_pdf = os.path.join(temp_dir, "Final_Scorecards.pdf")
    merger = PdfMerger()
    for pdf in sorted(final_pdf_list, key=lambda x: int(x.split('_')[-1].split('.')[0])):
        merger.append(pdf)
    merger.write(output_pdf)
    merger.close()

    for pdf in final_pdf_list:
        os.remove(pdf)

    return output_pdf
