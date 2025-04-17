from flask import Blueprint, request, jsonify
from database import get_db_connection

bp = Blueprint('track_attendance', __name__)

@bp.route('/track-attendance', methods=['GET'])
def track_attendance():
    branch = request.args.get('branch')
    roll_no = request.args.get('roll_no')
    semester = request.args.get('semester', 'ALL Semesters')

    if not branch or not roll_no:
        return jsonify({"message": "Branch and Roll No are required"}), 400

    try:
        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tables = cursor.fetchall()
        relevant_tables = [t['name'] for t in tables if branch in t['name']]

        attendance_summary = {}

        for table in relevant_tables:
            query = f"""
                SELECT Month_Year, Attendance, Semester 
                FROM {table} 
                WHERE Roll_No = ?
            """
            params = (roll_no,)
            if semester != 'ALL Semesters':
                query += " AND Semester = ?"
                params += (semester,)

            cursor.execute(query, params)
            rows = cursor.fetchall()

            for row in rows:
                sem = row['Semester']
                if sem not in attendance_summary:
                    attendance_summary[sem] = {
                        "attendance": [],
                        "total_attendance_percentage": 0,
                        "record_count": 0
                    }

                attendance_summary[sem]["attendance"].append({
                    "Month_Year": row["Month_Year"],
                    "Attendance_Percentage": row["Attendance"]
                })
                attendance_summary[sem]["total_attendance_percentage"] += row["Attendance"]
                attendance_summary[sem]["record_count"] += 1

        # Calculate average attendance percentage for each semester
        for sem, data in attendance_summary.items():
            if data["record_count"]:
                avg = data["total_attendance_percentage"] / data["record_count"]
                data["average_attendance_percentage"] = round(avg, 2)

        if attendance_summary:
            return jsonify({"attendance": attendance_summary}), 200
        else:
            return jsonify({"message": "No attendance data found"}), 404

    except Exception as e:
        return jsonify({"message": f"Error: {str(e)}"}), 500
    finally:
        conn.close()