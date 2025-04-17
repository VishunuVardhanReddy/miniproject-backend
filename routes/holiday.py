from flask import Blueprint, request, send_file
from docx import Document
from datetime import datetime
import pythoncom
import comtypes.client
import os

bp = Blueprint('holiday', __name__)

@bp.route('/generate-holiday-circular', methods=['POST'])
def generate_holiday_circular():
    data = request.get_json()
    date = data.get('date')
    reason = data.get('reason')
    is_holiday = data.get('isHoliday')

    try:
        holiday_date = datetime.strptime(date, "%Y-%m-%d")
        day_name = holiday_date.strftime("%A")
    except ValueError:
        return "Invalid date format", 400

    today = datetime.now().strftime("%B %d, %Y")
    doc = Document('templates/holiday_template.docx')

    replacements = {
        '{DATE}': date,
        '{REASON}': reason,
        '{IS_HOLIDAY}': 'holiday' if is_holiday else 'working day',
        '{CURRENT_DATE}': today,
        '{DAY_OF_WEEK}': day_name
    }

    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    word_path = "generated_letters/temp_holiday_notice.docx"
    pdf_path = word_path.replace(".docx", ".pdf")
    doc.save(word_path)

    pythoncom.CoInitialize()
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(word_path))
    doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
    doc.Close()
    word.Quit()

    return send_file(pdf_path, as_attachment=True, download_name="holiday_notice.pdf", mimetype="application/pdf")