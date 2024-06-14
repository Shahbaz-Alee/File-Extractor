from flask import Flask, request, render_template
import os
import fitz  # PyMuPDF
import docx
import pandas as pd
import re

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def extract_text_from_pdf(filepath):
    text = ""
    document = fitz.open(filepath)
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        page_text = page.get_text()
        # Remove <br> tags and extra spaces
        page_text = re.sub(r'<br>\s*', '\n', page_text)
        # Remove other HTML tags if any
        page_text = re.sub(r'<.*?>', '', page_text)
        text += page_text + "\n\n"
    return text.strip()

def extract_text_from_docx(filepath):
    doc = docx.Document(filepath)
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

def extract_text_from_xlsx(filepath):
    text = ""
    xlsx = pd.ExcelFile(filepath)
    for sheet in xlsx.sheet_names:
        df = pd.read_excel(filepath, sheet_name=sheet)
        text += df.to_string()
    return text

def extract_text(filepath, filename):
    if filename.endswith('.pdf'):
        return extract_text_from_pdf(filepath)
    elif filename.endswith('.docx'):
        return extract_text_from_docx(filepath)
    elif filename.endswith('.xlsx'):
        return extract_text_from_xlsx(filepath)
    elif filename.endswith('.txt'):
        with open(filepath, 'r') as file:
            return file.read()
    return "Unsupported file type"

@app.route('/')
def home():
    return render_template('index.html', text=None)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        text = extract_text(filepath, file.filename)
        return render_template('index.html', text=text)

if __name__ == '__main__':
    app.run(debug=True)
