from flask import Blueprint, request, jsonify
from database import get_db_connection

bp = Blueprint('upload_attendance', __name__)

@bp.route('/upload-attendance', methods=['POST'])
def upload_attendance():
    data = request.json.get('data')
    branch = request.json.get('branch')
    batch = request.json.get('batch')
    month_year = request.json.get('monthYear')

    if not data or not branch or not batch or not month_year:
        return jsonify({'message': 'Missing required fields'}), 400

    conn = get_db_connection()
    cursor = conn.cursor()

    table_name = f'Attendance_{batch}_{branch}_{month_year}'
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    if not cursor.fetchone():
        cursor.execute(f"""
            CREATE TABLE {table_name} (
                roll_no TEXT NOT NULL,
                branch TEXT NOT NULL,
                batch TEXT NOT NULL,
                semester TEXT NOT NULL,
                month_year TEXT NOT NULL,
                attendance INTEGER NOT NULL,
                UNIQUE(roll_no, month_year)
            )
        """)
        conn.commit()

    try:
        new_entries = []
        for record in data:
            roll_no = record.get('Roll_No')
            semester = record.get('Semester')
            attendance = record.get('Attendance')
            if roll_no and attendance is not None:
                cursor.execute(f"""
                    SELECT * FROM {table_name}
                    WHERE roll_no = ? AND month_year = ?
                """, (roll_no, month_year))
                if cursor.fetchone():
                    cursor.execute(f"""
                        UPDATE {table_name}
                        SET attendance = ?
                        WHERE roll_no = ? AND month_year = ?
                    """, (attendance, roll_no, month_year))
                else:
                    new_entries.append((roll_no, branch, batch, semester, month_year, attendance))

        if new_entries:
            cursor.executemany(f"""
                INSERT INTO {table_name} (roll_no, branch, batch, semester, month_year, attendance)
                VALUES (?, ?, ?, ?, ?, ?)
            """, new_entries)

        conn.commit()
        return jsonify({'message': 'Data uploaded successfully'}), 200

    except Exception as e:
        return jsonify({'message': f'Error uploading data: {str(e)}'}), 500
    finally:
        conn.close()
