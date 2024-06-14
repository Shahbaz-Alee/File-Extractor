from flask import Flask, request, jsonify, render_template_string, render_template
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
    return render_template('index.html', extracted_text=None)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        error_message = "No file part"
        return render_template_string('''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Error</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    padding: 20px;
                }
                .container {
                    max-width: 800px;
                    margin: 0 auto;
                }
                .error {
                    color: red;
                    font-weight: bold;
                }
            </style>
            <script>
                window.onload = function() {
                    alert("''' + error_message + '''");
                    window.history.back();
                }
            </script>
        </head>
        <body>
            <div class="container">
                <h1>Error</h1>
                <p class="error">An error occurred: {{ error_message }}</p>
            </div>
        </body>
        </html>
        ''', error_message=error_message)

    file = request.files['file']
    if file.filename == '':
        error_message = "No selected file"
        return render_template_string('''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Error</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    line-height: 1.6;
                    padding: 20px;
                }
                .button {
                    display: inline-block;
                    padding: 10px 20px;
                    background-color: #4CAF50;
                    color: white;
                    text-align: center;
                    text-decoration: none;
                    font-size: 16px;
                    margin-top: 10px;
                    cursor: pointer;
                    border: none;
                    border-radius: 5px;
                    transition: background-color 0.3s; /* Add transition for smooth color change */
                }

                .button:hover {
                    background-color: #45a049; /* New background color on hover */
                }
                .container {
                    max-width: 800px;
                    margin: 0 auto;
                }
                .error {
                    color: red;
                    font-weight: bold;
                }
            </style>
            <script>
                window.onload = function() {
                    alert("''' + error_message + '''");
                    window.history.back();
                }
            </script>
        </head>
        <body>
            <div class="container">
                <h1>Error</h1>
                <p class="error">An error occurred: {{ error_message }}</p>
            </div>
        </body>
        </html>
        ''', error_message=error_message)

    if file:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filepath)
        text = extract_text(filepath, file.filename)
        return render_template_string('''
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>File Extractor</title>
            <style>
                body { 
                    font-family: Arial, sans-serif; 
                    line-height: 1.6; 
                    padding: 20px; 
                    margin: 0; 
                    overflow: hidden; /* Prevents scrollbars from appearing */
                }
                .container { 
                    max-width: 800px; 
                    margin: 0 auto; 
                    position: relative; 
                    z-index: 2; 
                    background: rgba(255, 255, 255, 0.8); /* Semi-transparent background for better readability */
                    padding: 20px; 
                    border-radius: 10px; 
                }
                .textbox { 
                    width: 100%; 
                    height: 300px; 
                    border: 1px solid #ccc; 
                    padding: 10px; 
                    margin-bottom: 10px; 
                    resize: vertical; 
                    font-family: 'Times New Roman', Times, serif; 
                    font-size: 14px; 
                }
                .button { 
                    display: inline-block; 
                    padding: 10px 20px; 
                    background-color: #4CAF50; 
                    color: white; 
                    text-align: center; 
                    text-decoration: none; 
                    font-size: 16px; 
                    margin-top: 10px; 
                    cursor: pointer; 
                    border: none; 
                    border-radius: 5px; 
                }
                .button:hover { 
                    background-color: #45a049; 
                }
                .bgAnimation {
                    position: absolute;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 100vh;
                    display: grid;
                    grid-template-columns: repeat(20, 1fr);
                    grid-template-rows: repeat(20, 1fr);
                    background: #1d1d1d;
                    filter: saturate(2);
                    overflow: hidden;
                    z-index: 1; /* Lower z-index than the container */
                }
                .colorBox {
                    z-index: 2;
                    filter: brightness(1.1);
                    transition: 2s ease;
                    position: relative;
                    margin: 2px;
                    background: #1d1d1d;
                }
                .colorBox:hover {
                    background: #00bfff;
                    transition-duration: 0s;
                }
                .backgroundAmim {
                    position: absolute;
                    top: 0;
                    left: 0;
                    width: 100%;
                    height: 40px;
                    background: #00bfff;
                    filter: blur(60px);
                    animation: animBack 6s linear infinite;
                }
                @keyframes animBack {
                    0% { top: 0; }
                    100% { top: 100vh; }
                }
            </style>
            <script>
                function copyToClipboard() {
                    var copyText = document.getElementById("extracted-text");
                    copyText.select();
                    document.execCommand("copy");
                }
            </script>
        </head>
        <body>
            <div class="bgAnimation" id="bgAnimation">
                <div class="backgroundAmim"></div>
            </div>
            <div class="container">
                <h1>Upload your document</h1>
                <form action="/upload" method="post" enctype="multipart/form-data"> 
                    <input type="file" name="file">
                    <input type="submit" value="Upload" class="button">
                </form>
                <hr>
                {% if text %}
                <h2>Extracted Text</h2>
                <textarea id="extracted-text" class="textbox" rows="10" readonly>{{ text }}</textarea><br>
                <button onclick="copyToClipboard()" class="button">Copy Text</button>
                {% endif %}
            </div>
            <script src="{{ url_for('static', filename='script.js') }}"></script>
        </body>
        </html>
        ''', text=text)

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
