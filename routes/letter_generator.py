from flask import Blueprint, request, jsonify
from database import get_db_connection
from utils.docx_utils import create_letter
from utils.pdf_utils import merge_letters_to_pdf
import os
import pandas as pd
from io import BytesIO
import base64

bp = Blueprint('letter_generator', __name__)

@bp.route('/generate-letter', methods=['POST'])
def generate_letter_route():
    file = request.files['file']
    month = request.form['month']
    include_monthly = request.form['includeMonthlyReport'] == 'true'

    file_path = os.path.join("uploads", file.filename)
    file.save(file_path)

    df = pd.read_excel(file_path)
    conn = get_db_connection()
    cursor = conn.cursor()

    condonation, detention, monthly = [], [], []

    for _, row in df.iterrows():
        roll = row['Roll_No']
        att = row['Attendance']
        cursor.execute("SELECT student_name, parent_name, address, branch FROM students WHERE Roll_no=?", (roll,))
        s = cursor.fetchone()
        if s:
            letter = None
            if include_monthly and att < 75:
                letter = create_letter('templates/Attendance_Report.docx', *s, att, roll)
                monthly.append(letter)
            elif not include_monthly:
                if att < 65:
                    letter = create_letter('templates/Detention_Report.docx', *s, att, roll)
                    detention.append(letter)
                elif att < 75:
                    letter = create_letter('templates/Condonation_Report.docx', *s, att, roll)
                    condonation.append(letter)

    conn.close()

    combined_pdf = merge_letters_to_pdf(condonation, detention, monthly)
    with open(combined_pdf, 'rb') as f:
        encoded_pdf = base64.b64encode(BytesIO(f.read()).getvalue()).decode()

    return jsonify({
        "message": "Letter generation successful!",
        "generated_students": {
            "condonation_count": len(condonation),
            "detention_count": len(detention),
            "attendance_count": len(monthly),
            "pdf_preview": encoded_pdf
        }
    })