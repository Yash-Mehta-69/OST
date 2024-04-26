import os
import re
import pandas as pd
from pdfminer.high_level import extract_text
from docx import Document

def extract_info_from_pdf(pdf_file):
    try:
        text = extract_text(pdf_file)
    except Exception as e:
        print(f"Error extracting text from PDF file '{pdf_file}': {e}")
        return [], [], ''

    # Regular expressions to find email and phone number
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'(?:(?:\+|0{0,2})91[\s-]?)?[6789]\d{9}\b'

    # Extracting email and phone number
    email = re.findall(email_pattern, text)
    phone = re.findall(phone_pattern, text)
    
    # Remove duplicates
    email = list(set(email))
    phone = list(set(phone))
    
    # Extracting overall text (excluding email and phone number)
    overall_text = re.sub(email_pattern, '', text)
    overall_text = re.sub(phone_pattern, '', overall_text)
    
    return email, phone, overall_text

def extract_info_from_docx(docx_file):
    try:
        doc = Document(docx_file)
    except Exception as e:
        print(f"Error opening DOCX file '{docx_file}': {e}")
        return [], [], ''

    # Regular expressions to find email and phone number
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    phone_pattern = r'(?:(?:\+|0{0,2})91[\s-]?)?[6789]\d{9}\b'
    
    # Extracting email and phone number
    email = []
    phone = []
    text = ''
    
    for paragraph in doc.paragraphs:
        text += paragraph.text + '\n'
    
    email = re.findall(email_pattern, text)
    phone = re.findall(phone_pattern, text)
    
    # Remove duplicates
    email = list(set(email))
    phone = list(set(phone))
    
    # Extracting overall text (excluding email and phone number)
    overall_text = re.sub(email_pattern, '', text)
    overall_text = re.sub(phone_pattern, '', overall_text)
    
    return email, phone, overall_text

def extract_info_from_cv(cv_file):
    if cv_file.lower().endswith('.pdf'):
        return extract_info_from_pdf(cv_file)
    elif cv_file.lower().endswith('.docx'):
        return extract_info_from_docx(cv_file)
    else:
        print("Unsupported file format:", cv_file)
        return [], [], ''

def create_excel(emails, phones, texts, output_file='cv_info.xlsx'):
    # Check if all arrays have the same length
    if len(emails) != len(texts):
        print("Arrays must be of the same length.")
        print("Length of emails:", len(emails))
        print("Length of texts:", len(texts))
        return
    
    # If phone list is shorter than email list, pad it with empty strings
    if len(phones) < len(emails):
        phones += [''] * (len(emails) - len(phones))
    # If phone list is longer than email list, truncate it
    elif len(phones) > len(emails):
        phones = phones[:len(emails)]
    
    # Create a Pandas DataFrame
    df = pd.DataFrame({'Email': emails, 'Phone': phones, 'Text': texts})
    
    # Write DataFrame to Excel
    df.to_excel(output_file, index=False)

# Example usage
if __name__ == "__main__":
    # Folder containing CV files
    cv_folder = 'Sample2'
    
    # Lists to store extracted information
    emails = []
    phones = []
    texts = []
    
    # Iterate over all files in the folder
    for filename in os.listdir(cv_folder):
        cv_file = os.path.join(cv_folder, filename)
        if os.path.isfile(cv_file):
            email, phone, overall_text = extract_info_from_cv(cv_file)
            print(f"Filename: {filename}, Email count: {len(email)}, Phone count: {len(phone)}, Text length: {len(overall_text)}")
            if email or phone or overall_text:  # Only append if at least one type of information is found
                emails.extend(email)
                phones.extend(phone[:1])  # Select only the first phone number
                texts.append(overall_text)
    
    # Creating Excel file
    create_excel(emails, phones, texts)
