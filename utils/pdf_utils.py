from PyPDF2 import PdfMerger
import os
from datetime import datetime

def is_valid_pdf(file_path):
    try:
        with open(file_path, 'rb'):
            return True
    except:
        return False

def merge_letters_to_pdf(condonation_letters, detention_letters, attendance_letters):
    merger = PdfMerger()
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    output_path = f"generated_letters/combined_letters_{timestamp}.pdf"

    for letter in condonation_letters + detention_letters + attendance_letters:
        if is_valid_pdf(letter):
            merger.append(letter)

    with open(output_path, 'wb') as f:
        merger.write(f)
    return output_path
