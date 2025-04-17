from flask import Blueprint, request, jsonify
from database import get_db_connection
from datetime import datetime, timedelta

bp = Blueprint('attendance', __name__)

@bp.route('/get-attendance', methods=['POST'])
def get_attendance():
    try:
        data = request.get_json()
        roll_no = data.get('rollNo')
        from_date = datetime.strptime(data.get('fromDate'), '%Y-%m-%d')
        to_date = datetime.strptime(data.get('toDate'), '%Y-%m-%d')

        attendance_data = []
        total_classes = 0
        attended_classes = 0

        for i in range((to_date - from_date).days + 1):
            current_date = from_date + timedelta(days=i)
            table_name = f"Attendance_Details_{current_date.strftime('%d_%b_%Y')}"

            conn = get_db_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
            if cursor.fetchone():
                cursor.execute(f"SELECT total_classes, attended_classes FROM {table_name} WHERE roll_no = ?", (roll_no,))
                row = cursor.fetchone()
                if row:
                    total, attended = row
                    perc = (attended / total * 100) if total else 0
                    attendance_data.append({
                        "date": current_date.strftime('%d %b %Y'),
                        "total_classes": total,
                        "attended_classes": attended,
                        "attendance_percentage": round(perc, 2)
                    })
                    total_classes += total
                    attended_classes += attended

            conn.close()

        summary = {
            "total_classes": total_classes,
            "attended_classes": attended_classes,
            "attendance_percentage": round((attended_classes / total_classes * 100), 2) if total_classes else 0
        }

        return jsonify({
            "status": "success",
            "attendance_data": attendance_data,
            "combined_summary": summary
        })

    except Exception as e:
        print(f"Error: {e}")
        return jsonify({"status": "error", "message": "Failed to fetch attendance."})