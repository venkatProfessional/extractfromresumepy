import pdfplumber
import docx
import os
import re
import openpyxl
import warnings
import requests
import shutil
from openpyxl.styles import Alignment

warnings.filterwarnings("ignore", message="CropBox missing from /Page, defaulting to MediaBox")

# --- Tamil Nadu city list ---
tamil_nadu_cities = {
    "Chennai", "Coimbatore", "Madurai", "Tiruchirappalli", "Salem", "Tirunelveli", "Tiruppur",
    "Erode", "Vellore", "Thoothukudi", "Dindigul", "Thanjavur", "Nagercoil", "Karur", "Kanchipuram",
    "Virudhunagar", "Sivakasi", "Cuddalore", "Nagapattinam", "Ariyalur", "Perambalur", "Namakkal",
    "Dharmapuri", "Krishnagiri", "Pudukkottai", "Ramanathapuram", "Tenkasi", "Villupuram",
    "Theni", "Sivaganga", "Tiruvarur", "Nilgiris"
}

COMMON_JOB_TITLES = {
    "Software Developer", "Project Manager", "Data Analyst", "HR Manager",
    "Software Engineer", "Business Analyst", "Technical Lead", "DevOps Engineer",
    "UI Designer", "Backend Developer", "Frontend Developer", "Web Developer",
    "C Programming", "Front End Developer", "An Electronics And Communication Under Graduate", "About Me"
}


# --- Extraction Functions ---
def extract_email(text):
    match = re.search(r'[\w\.-]+@[\w\.-]+', text)
    return match.group(0) if match else None


def extract_phone(text):
    match = re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', text)
    return match.group(0) if match else None


def extract_degrees(text):
    degrees = ['Bachelor', 'Master', 'PhD', 'BSc', 'MSc', 'MBA']
    found = [deg for deg in degrees if deg.lower() in text.lower()]
    return ', '.join(found) if found else None


def validate_location_api(location):
    try:
        response = requests.get(f"http://ip-api.com/json/{location}")
        data = response.json()
        if data.get('status') == 'success':
            return data.get('city') or location
    except:
        pass
    return location


def extract_location(text):
    """
    Extract location information from resume text with improved accuracy.
    Returns the most likely location as a string, focusing on Tamil Nadu cities when possible.
    """
    # Known Tamil Nadu cities for validation
    tamil_nadu_cities = {
        "Ariyalur", "Chengalpattu", "Chennai", "Coimbatore", "Cuddalore", "Dharmapuri", "Dindigul",
        "Erode", "Kallakurichi", "Kanchipuram", "Kanyakumari", "Karur", "Krishnagiri", "Madurai",
        "Mayiladuthurai", "Nagapattinam", "Namakkal", "Nilgiris", "Perambalur", "Pudukkottai",
        "Ramanathapuram", "Ranipet", "Salem", "Sivaganga", "Sivakasi", "Tenkasi", "Thanjavur",
        "Theni", "Thoothukudi", "Tiruchirappalli", "Tirunelveli", "Tirupathur", "Tiruppur",
        "Tiruvallur", "Tiruvannamalai", "Tiruvarur", "Vellore", "Viluppuram", "Virudhunagar",
        "Ambur", "Arakkonam", "Avadi", "Chidambaram", "Gobichettipalayam", "Hosur",
        "Karaikudi", "Kumbakonam", "Mettupalayam", "Nagercoil", "Pattukkottai", "Pollachi",
        "Rajapalayam", "Sankagiri", "Srivilliputhur", "Tambaram", "Thiruvarur", "Thiruvannamalai",
        "Tuticorin", "Udumalaipettai", "Vaniyambadi", "Vellakoil", "Wellington"
    }








    # Additional Indian cities for better coverage
    major_indian_cities = {
        "Mumbai", "Delhi", "Bangalore", "Hyderabad", "Ahmedabad", "Pune", "Surat", "Jaipur",
        "Kolkata", "Lucknow", "Kanpur", "Nagpur", "Indore", "Thane", "Bhopal", "Visakhapatnam",
        "Patna", "Vadodara", "Ghaziabad", "Ludhiana", "Agra", "Nashik", "Faridabad", "Meerut",
        "Rajkot", "Varanasi", "Srinagar", "Aurangabad", "Dhanbad", "Amritsar", "Allahabad",
        "Ranchi", "Howrah", "Jabalpur", "Gwalior", "Vijayawada", "Jodhpur", "Raipur", "Kota"
    }

    # All valid cities to check
    all_cities = tamil_nadu_cities.union(major_indian_cities)

    import re

    # Function to clean and standardize text for better matching
    def clean_text(text):
        # Convert to lowercase
        text = text.lower()
        # Replace multiple spaces with single space
        text = re.sub(r'\s+', ' ', text)
        # Remove special characters
        text = re.sub(r'[^\w\s,.-]', '', text)
        return text

    # Clean the input text
    cleaned_text = clean_text(text)

    # PHASE 1: Check for explicit location markers
    explicit_patterns = [
        # Location headers with various formats
        r'(?:location|address|residence|residing at|based in|residing in|current location|permanent address)[:\s]+([A-Za-z0-9\s,.()-]+)(?:\n|$|,)',
        # Contact information section
        r'(?:contact|personal info|details)[:\s]+(?:[^,]*?,\s*)?([A-Za-z\s]+(?:,\s*[A-Za-z\s]+)+)',
        # Address with postal code format
        r'([A-Za-z\s]+,\s*[A-Za-z\s]+\s*[-,]\s*\d{5,6})',
        # City, State format
        r'([A-Za-z\s]+,\s*[A-Za-z\s]+\s*[-,]\s*[A-Za-z\s]+)',
        # Standalone city-state combinations
        r'(?<=\n|\s)([A-Za-z\s]+,\s*[A-Za-z]{2})(?=\n|\s|$)'
    ]

    for pattern in explicit_patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            # Process match to extract just the city
            location_text = match.strip() if isinstance(match, str) else match[0].strip()

            # Try to extract Tamil Nadu cities from the match
            for city in tamil_nadu_cities:
                if city.lower() in location_text.lower():
                    return city

            # Try other Indian cities if no Tamil Nadu city found
            for city in major_indian_cities:
                if city.lower() in location_text.lower():
                    return city

            # If no known city found but we have location text, parse it for city
            components = re.split(r'[,\s-]+', location_text)
            for component in components:
                component = component.strip()
                # Check if component matches a known city (case insensitive)
                city_match = next((city for city in all_cities if city.lower() == component.lower()), None)
                if city_match:
                    return city_match

    # PHASE 2: Look for city names within the entire text
    # First check for Tamil Nadu cities (preference)
    text_words = set(re.findall(r'\b[A-Za-z]+\b', text))
    for city in tamil_nadu_cities:
        if city in text_words:
            return city

    # Check for other Indian cities
    for city in major_indian_cities:
        if city in text_words:
            return city

    # PHASE 3: Extract from email domains if available (last resort)
    email_match = re.search(r'[\w\.-]+@([\w\.-]+)', text)
    if email_match:
        domain = email_match.group(1).lower()
        if 'chennai' in domain:
            return 'Chennai'
        elif 'madras' in domain:
            return 'Chennai'
        elif 'coimbatore' in domain:
            return 'Coimbatore'
        # Add more domain-to-city mappings as needed

    # PHASE 4: Smart context analysis for implicit locations
    # Look for educational institutions and map them to cities
    education_markers = [
        ('Anna University', 'Chennai'),
        ('IIT Madras', 'Chennai'),
        ('Madras University', 'Chennai'),
        ('PSG', 'Coimbatore'),
        ('Amrita', 'Coimbatore'),
        ('VIT', 'Vellore'),
        ('NIT Trichy', 'Tiruchirappalli'),
        ('Madurai Kamaraj University', 'Madurai'),
        ('Annamalai University', 'Chidambaram'),
        # Add more institution-city mappings
    ]

    for institution, city in education_markers:
        if institution.lower() in cleaned_text:
            return city

    # PHASE 5: Postal code analysis
    # Extract Tamil Nadu postal codes
    tn_postal_codes = re.findall(r'\b6[0-4][0-9]{4}\b', text)
    if tn_postal_codes:
        # Map postal code ranges to cities (simplified mapping)
        postal_ranges = {
            '60': 'Chennai',  # 600001-609999
            '61': 'Kanchipuram',  # 610000-619999
            '62': 'Coimbatore',  # 620000-629999
            '63': 'Madurai',  # 630000-639999
            '64': 'Salem'  # 640000-649999
            # Add more mappings as needed
        }

        for code in tn_postal_codes:
            prefix = code[:2]
            if prefix in postal_ranges:
                return postal_ranges[prefix]

    # If no location found after all attempts
    return None


def extract_name(text):
    lines = text.strip().split('\n')
    first_five_lines = [line.strip() for line in lines if line.strip()][:3]

    def is_valid_name(line):
        clean_line = line.strip()
        return clean_line and clean_line not in COMMON_JOB_TITLES

    for line in first_five_lines:
        if is_valid_name(line) and re.match(r'^([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)$', line):
            return line

    indian_name_patterns = [
        r'^([A-Z]\.\s?[A-Z][a-z]+)$',
        r'^([A-Z]{1,2}\s?[A-Z][a-z]+)$',
        r'^([A-Z][a-z]+\s?[A-Z]{1,2})$',
        r'^([A-Z][a-z]+\s[A-Z][a-z]+)$',
        r'^([A-Z][a-z]+\s[A-Z][a-z]+\s[A-Z][a-z]+)$',
        r'^([A-Z]\.\s?[A-Z][a-z]+\s[A-Z][a-z]+)$'
    ]

    for line in first_five_lines:
        if not is_valid_name(line):
            continue
        for pattern in indian_name_patterns:
            if re.match(pattern, line):
                return line

    email = extract_email(text)
    if email:
        name_part = email.split('@')[0]
        name_part = name_part.replace('.', ' ').replace('_', ' ')
        return ' '.join(word.capitalize() for word in name_part.split())

    return None


# --- File Reading ---
def read_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        return '\n'.join([page.extract_text() or '' for page in pdf.pages])


def read_docx(file_path):
    doc = docx.Document(file_path)
    return '\n'.join([para.text for para in doc.paragraphs])


# --- Filename Sanitization ---
def sanitize_filename(name):
    """Sanitize the name to be used in a filename"""
    if not name:
        return "Unknown"

    # Remove invalid filename characters
    name = re.sub(r'[\\/*?:"<>|]', '', name)
    # Replace spaces with underscores
    name = name.replace(' ', '_')
    # Limit filename length (Windows has a limit of 260 chars for full path)
    name = name[:50]  # Safe limit for a name part
    return name


# --- Resume Processing ---
def process_resumes(folder_path, output_folder):
    # Create output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    resume_data = []
    renamed_count = 0
    for file_name in os.listdir(folder_path):
        path = os.path.join(folder_path, file_name)
        if not os.path.isfile(path):
            continue

        # Check if file is supported
        file_ext = os.path.splitext(file_name)[1].lower()
        if file_ext not in ['.pdf', '.docx']:
            print(f"⚠️ Skipping unsupported file: {file_name}")
            continue

        try:
            # Read the file content
            if file_ext == '.pdf':
                text = read_pdf(path)
            elif file_ext == '.docx':
                text = read_docx(path)

            # Extract information
            name = extract_name(text)
            email = extract_email(text)
            phone = extract_phone(text)
            degrees = extract_degrees(text)
            location = extract_location(text)

            # Clean name: remove digits and sanitize for filename
            if name:
                name = re.sub(r'\d+', '', name).strip()
                sanitized_name = sanitize_filename(name)
            else:
                sanitized_name = "Unknown"

            # Create new filename based on name and location
            if location:
                new_filename = f"{sanitized_name}_{location}{file_ext}"
            else:
                new_filename = f"{sanitized_name}_resume{file_ext}"

            # Ensure the new filename is unique
            counter = 1
            base_name = os.path.splitext(new_filename)[0]
            while os.path.exists(os.path.join(output_folder, new_filename)):
                new_filename = f"{base_name}_{counter}{file_ext}"
                counter += 1

            # Copy the file to output folder with new name
            output_path = os.path.join(output_folder, new_filename)
            shutil.copy2(path, output_path)
            renamed_count += 1

            # Add data to list for Excel
            resume_data.append([name, email, phone, degrees, location, file_name, new_filename])

            print(f"✅ Renamed: {file_name} → {new_filename}")

        except Exception as e:
            print(f"❌ Error processing {file_name}: {str(e)}")

    return resume_data, renamed_count


# --- Save to Excel ---


def create_excel(data, output_file):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resume Data"

    # Header row
    headers = ["Name", "Email", "Phone", "Location"]
    ws.append(headers)

    # Set column widths manually
    column_widths = [25, 38, 20, 40]  # Adjust as needed
    for i, width in enumerate(column_widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = width

    for row_idx, row in enumerate(data, start=2):
        name = row[0].strip() if isinstance(row[0], str) else row[0]
        email = row[1].strip() if isinstance(row[1], str) else row[1]
        phone = row[2].strip() if isinstance(row[2], str) else row[2]
        location = row[4].strip() if isinstance(row[4], str) and row[4].strip() else "Unknown Location"

        cleaned_row = [name, email, phone, location]
        ws.append(cleaned_row)

        # Optional: set row height
        ws.row_dimensions[row_idx].height = 25  # Adjust height as needed

    # Set header row height
    ws.row_dimensions[1].height = 30

    wb.save(output_file)
    print(f"✅ Excel file created with increased cell sizes: {output_file}")




# --- Main Function ---
def main():
    folder_path = r"E:\\Venkat\\resumebank"  # Change to your folder path
    output_folder = r"E:\\Venkat\\renamed_resumes"  # Change to your desired output folder
    output_excel = "Resume_Cleaned_data_7.xlsx"

    print(f"📂 Processing resumes from: {folder_path}")
    print(f"📁 Output folder: {output_folder}")

    # Process and rename the resumes
    data, renamed_count = process_resumes(folder_path, output_folder)

    # Count total processed and locations found
    total_processed = len(data)
    locations_found = sum(1 for row in data if row[4] is not None)

    print(f"\n📊 Summary:")
    print(f"   - Total resumes processed: {total_processed}")
    print(f"   - Resumes successfully renamed: {renamed_count}")
    print(
        f"   - Locations successfully extracted: {locations_found} ({int(locations_found / total_processed * 100) if total_processed > 0 else 0}%)")

    # Create the Excel file
    create_excel(data, output_excel)

    # Provide summary of top locations found
    if locations_found > 0:
        location_counts = {}
        for row in data:
            location = row[4]
            if location:
                location_counts[location] = location_counts.get(location, 0) + 1

        print(f"\n📍 Top locations found:")
        for location, count in sorted(location_counts.items(), key=lambda x: x[1], reverse=True)[:5]:
            print(f"   - {location}: {count} resumes")


if __name__ == "__main__":
    main()