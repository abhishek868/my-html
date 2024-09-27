import os
from flask import Flask, request, render_template, send_file, redirect, url_for, flash
from pdf2docx import Converter
from pptx import Presentation
from docx import Document
import fitz  # PyMuPDF for extracting text from PDFs
import json
import time
from pathlib import Path
import shutil

# Initialize Flask app
app = Flask(__name__)
app.secret_key = "supersecretkey"  # For handling flash messages

CONVERTED_FOLDER = Path('converted')

# Ensure the converted directory exists
CONVERTED_FOLDER.mkdir(exist_ok=True)

def save_metadata(file_path, output_json_path):
    """Save file metadata to a JSON file."""
    file_metadata = {
        "file_path": str(file_path.resolve()),  # Absolute path
        "file_name": file_path.name,
        "file_size": file_path.stat().st_size,
        "created_time": time.ctime(file_path.stat().st_ctime),
        "modified_time": time.ctime(file_path.stat().st_mtime)
    }

    # Save metadata to JSON
    with open(output_json_path, 'w') as json_file:
        json.dump(file_metadata, json_file, indent=4)

def convert_pdf_to_docx(pdf_file, output_docx_file):
    """Convert PDF to DOCX using pdf2docx."""
    cv = Converter(str(pdf_file))
    cv.convert(str(output_docx_file), start=0, end=None)
    cv.close()

def extract_text_from_pdf(pdf_path, output_txt_file):
    """Extract plain text from a PDF using PyMuPDF (fitz)."""
    doc = fitz.open(str(pdf_path))
    extracted_text = ""

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        text = page.get_text("text")
        extracted_text += text + "\n\n"

    with open(output_txt_file, 'w', encoding='utf-8') as txt_file:
        txt_file.write(extracted_text)

def extract_text_from_pptx(pptx_path, output_txt_file):
    """Extract text from a PowerPoint file."""
    prs = Presentation(str(pptx_path))
    extracted_text = ""

    for slide_num, slide in enumerate(prs.slides):
        slide_header = f"Slide {slide_num + 1}\n{'=' * 10}\n"
        extracted_text += slide_header

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                extracted_text += shape.text + "\n"

        extracted_text += "\n\n"

    with open(output_txt_file, 'w', encoding='utf-8') as txt_file:
        txt_file.write(extracted_text)

def save_pptx_as_docx(pptx_path, output_docx_file):
    """Convert PowerPoint to DOCX format."""
    prs = Presentation(str(pptx_path))
    doc = Document()

    for slide_num, slide in enumerate(prs.slides):
        doc.add_heading(f'Slide {slide_num + 1}', level=1)

        for shape in slide.shapes:
            if hasattr(shape, "text"):
                doc.add_paragraph(shape.text)

    doc.save(output_docx_file)

def process_file(file_path):
    """Process an individual PDF or PPTX file."""
    if file_path.suffix == '.pdf':
        docx_filename = file_path.stem + '.docx'
        docx_path = CONVERTED_FOLDER / docx_filename
        convert_pdf_to_docx(file_path, docx_path)

        txt_filename = file_path.stem + '.txt'
        txt_path = CONVERTED_FOLDER / txt_filename
        extract_text_from_pdf(file_path, txt_path)

        json_filename = file_path.stem + '.json'
        json_path = CONVERTED_FOLDER / json_filename
        save_metadata(file_path, json_path)

    elif file_path.suffix == '.pptx':
        docx_filename = file_path.stem + '.docx'
        docx_path = CONVERTED_FOLDER / docx_filename
        save_pptx_as_docx(file_path, docx_path)

        txt_filename = file_path.stem + '.txt'
        txt_path = CONVERTED_FOLDER / txt_filename
        extract_text_from_pptx(file_path, txt_path)

        json_filename = file_path.stem + '.json'
        json_path = CONVERTED_FOLDER / json_filename
        save_metadata(file_path, json_path)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if file is uploaded
        if 'file' in request.files:
            uploaded_file = request.files['file']

            if uploaded_file.filename == '':
                flash('No file selected', 'error')
                return redirect(request.url)

            # Save the uploaded file
            file_path = Path('uploads') / uploaded_file.filename
            file_path.parent.mkdir(exist_ok=True)
            uploaded_file.save(file_path)

            # Process the uploaded file
            process_file(file_path)

            return redirect(url_for('download_all'))

        # Check if folder is uploaded
        if 'folder' in request.files:
            uploaded_files = request.files.getlist('folder')
            for uploaded_file in uploaded_files:
                if uploaded_file.filename == '':
                    flash('No file selected in folder', 'error')
                    continue

                # Save each file from the folder
                file_path = Path('uploads') / uploaded_file.filename
                file_path.parent.mkdir(exist_ok=True)
                uploaded_file.save(file_path)

                # Process the uploaded file
                process_file(file_path)

            return redirect(url_for('download_all'))

    return render_template('upload.html')

@app.route('/download_all')
def download_all():
    # Zip the converted folder and allow download
    shutil.make_archive('converted_files', 'zip', CONVERTED_FOLDER)
    return send_file('converted_files.zip', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
