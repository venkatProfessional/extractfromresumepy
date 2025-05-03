import pdfplumber
import docx
import os
import re
import warnings

warnings.filterwarnings("ignore", message="CropBox missing from /Page, defaulting to MediaBox")

# Extract name: try "Name:" label, capitalized lines, or fallback regex
def extract_name(text):
    lines = text.strip().split('\n')
    for line in lines[:10]:
        if re.search(r'\bName[:\-]?', line, re.IGNORECASE):
            possible_name = re.sub(r'(?i)Name[:\-]?', '', line).strip()
            if 2 <= len(possible_name.split()) <= 4:
                return possible_name
    for line in lines[:10]:
        words = line.strip().split()
        if 1 < len(words) <= 4 and all(w[0].isupper() for w in words if w.isalpha()):
            return ' '.join(words)
    matches = re.findall(r'\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+)+)\b', text)
    return matches[0] if matches else None

# Extract general location: keywords or pincode
def extract_location(text):
    for line in text.split('\n'):
        if re.search(r'\b(Address|Location|City|Resident of)\b', line, re.IGNORECASE):
            return line.strip()
        if re.search(r'\b\d{6}\b', line):  # Indian pincode
            return line.strip()
    return None

# Read text from PDF
def read_pdf(file_path):
    text = ''
    with pdfplumber.open(file_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + '\n'
    return text

# Read text from DOCX
def read_docx(file_path):
    doc = docx.Document(file_path)
    return '\n'.join(para.text for para in doc.paragraphs)

# Rename resumes using name and location
def rename_resumes(folder_path):
    for file_name in os.listdir(folder_path):
        if not (file_name.endswith('.pdf') or file_name.endswith('.docx')):
            continue

        full_path = os.path.join(folder_path, file_name)
        ext = os.path.splitext(file_name)[1].lower()

        try:
            text = read_pdf(full_path) if ext == '.pdf' else read_docx(full_path)
        except Exception as e:
            print(f"⚠️ Error reading {file_name}: {e}")
            continue

        name = extract_name(text) or "UnknownName"
        location = extract_location(text) or "UnknownLocation"

        safe_name = re.sub(r'[^a-zA-Z0-9]+', '_', name).strip('_')
        safe_location = re.sub(r'[^a-zA-Z0-9]+', '_', location).strip('_')

        new_name = f"{safe_name}_{safe_location}_Resume{ext}"
        new_path = os.path.join(folder_path, new_name)

        try:
            os.rename(full_path, new_path)
            print(f"✅ Renamed: {file_name} → {new_name}")
        except Exception as e:
            print(f"❌ Failed to rename {file_name}: {e}")

# Entry point
if __name__ == "__main__":
    folder_path = r"E:\Venkat\resumebank"  # Update your folder path
    rename_resumes(folder_path)
