import os
import platform
from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename  # Import secure_filename here
from docx import Document
from reportlab.pdfgen import canvas
from comtypes.client import CreateObject
import importlib.util
import comtypes

# Check for module availability
COMTYPES_AVAILABLE = importlib.util.find_spec("comtypes") is not None
PANDOC_AVAILABLE = importlib.util.find_spec("pypandoc") is not None

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'

# Ensure upload and output folders exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def convert_docx_to_pdf(input_path, output_path, method="auto"):
    """
    Convert a DOCX file to PDF using the specified method.
    """
    try:
        if method == "comtypes" or (method == "auto" and COMTYPES_AVAILABLE):
            # Initialize COM
            comtypes.CoInitialize()

            # Use comtypes for Windows
            word = CreateObject("Word.Application")
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.SaveAs(os.path.abspath(output_path), FileFormat=17)  # 17 represents PDF format
            doc.Close()
            word.Quit()

            # Uninitialize COM
            comtypes.CoUninitialize()

        elif method == "pypandoc" or (method == "auto" and PANDOC_AVAILABLE):
            # Use pypandoc for cross-platform
            import pypandoc  # Import only when required
            pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)

        elif method == "reportlab" or method == "auto":
            # Use python-docx + reportlab for fallback
            doc = Document(input_path)
            pdf = canvas.Canvas(output_path)
            x, y = 50, 800
            for paragraph in doc.paragraphs:
                text = paragraph.text
                pdf.drawString(x, y, text)
                y -= 20
                if y < 50:
                    pdf.showPage()
                    y = 800
            pdf.save()
        else:
            raise ValueError("No valid method available or specified for conversion.")
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False

@app.route('/')
def upload_page():
    return render_template('index.html')

@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        return "No file uploaded", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    if file and file.filename.endswith('.docx'):
        method = request.form.get('method', 'auto')  # Get conversion method
        input_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], secure_filename(file.filename.replace('.docx', '.pdf')))

        # Save uploaded file
        file.save(input_path)

        # Convert DOCX to PDF
        success = convert_docx_to_pdf(input_path, output_path, method)
        if success:
            return send_file(output_path, as_attachment=True)
        else:
            return "Conversion failed", 500
    else:
        return "Invalid file type. Please upload a DOCX file.", 400

if __name__ == '__main__':
    app.run(debug=True)
