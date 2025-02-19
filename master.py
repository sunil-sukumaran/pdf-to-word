from pypdf import PdfReader
from docx import Document
from google.colab import files

# Upload PDF file
uploaded = files.upload()
pdf_file = list(uploaded.keys())[0]

# Read PDF
pdf_reader = PdfReader(pdf_file)
doc = Document()

for page in pdf_reader.pages:
    text = page.extract_text()
    if text:
        doc.add_paragraph(text)

# Save to Word
docx_file = pdf_file.replace('.pdf', '.docx')
doc.save(docx_file)

# Download the Word file
files.download(docx_file)
print(f'Conversion complete. Download your file: {docx_file}')
