from PyPDF2 import PdfReader, PdfWriter
from docx import Document
from docx2pdf import convert  # uses ms word to convert docx to pdf
import os


def convert_docx_to_pdf(input_path, output_path):
    convert(input_path, output_path)


def count_pdf_pages(pdf_path):
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        print(pdf_reader.pages[0].extract_text())
        return len(pdf_reader.pages)


def split_pdf_by_pages(pdf_path, output_dir):
    with open(pdf_path, "rb") as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        for page_num in range(len(pdf_reader.pages)):
            pdf_writer = PdfWriter()
            pdf_writer.add_page(pdf_reader.pages[page_num])
            output_path = os.path.join(output_dir, f"page_{page_num + 1}.pdf")
            with open(output_path, "wb") as output_pdf_file:
                pdf_writer.write(output_pdf_file)


def split_pdf_by_content(input_pdf, split_keywords, output_dir):
    reader = PdfReader(input_pdf)
    total_pages = len(reader.pages)
    split_points = []

    # Identify split points based on keywords
    for i in range(total_pages):
        page = reader.pages[i]
        text = page.extract_text()
        if any(keyword in text for keyword in split_keywords):
            split_points.append(i)

    # Split the PDF at identified points
    split_points.append(total_pages)  # Ensure the last section is included
    start = 0
    for idx, end in enumerate(split_points):
        content_added_flag = False
        writer = PdfWriter()

        for j in range(start, end):
            writer.add_page(reader.pages[j])
            content_added_flag = True

        if content_added_flag == True:
            output_pdf = output_dir + "/" + f"output_section_{idx}.pdf"
            with open(output_pdf, "wb") as output_file:
                writer.write(output_file)
            print(f"Created: {output_pdf}")
            start = end


# ==================================================================================================
if _name_ == "_main_":
    base_dir = "/home/maciej/projects/docx2pdfsplitter"
    input_path = os.path.join(base_dir, "document.docx")
    output_pdf_path = os.path.join(base_dir, "document.pdf")
    output_dir = base_dir

    # Convert DOCX to PDF
    convert_docx_to_pdf(input_path, output_pdf_path)

    # Count the number of pages in the PDF
    page_count = count_pdf_pages(output_pdf_path)
    print(f"Number of pages: {page_count}")

    # Split the PDF into separate files by pages
    # split_pdf_by_pages(output_pdf_path, output_dir)

    # split pdf by content
    split_keywords = ["CHAPTER 1", "CHAPTER 2",
                      "CHAPTER 4"]  # Keywords to split by
    split_pdf_by_content(output_pdf_path, split_keywords, output_dir)