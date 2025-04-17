from flask import Blueprint, request, jsonify
from flask_jwt_extended import create_access_token, jwt_required, get_jwt_identity
from database import get_db_connection
import bcrypt

bp = Blueprint('auth', __name__)

@bp.route('/login', methods=['POST'])
def login():
    data = request.json
    username = data.get('username', '')
    password = data.get('password', '')

    conn = get_db_connection()
    user = conn.execute('SELECT * FROM Login WHERE Username = ?', (username,)).fetchone()
    conn.close()

    if not user:
        return jsonify({"message": "Invalid username"}), 401

    if bcrypt.checkpw(password.encode('utf-8'), user['Password'].encode('utf-8')):
        token = create_access_token(identity={
            'userId': user['Username'],
            'role': user['Role'],
            'firstName': user['First_name'],
            'lastName': user['Last_name']
        })
        return jsonify({
            'token': token,
            'user': {
                'firstName': user['First_name'],
                'lastName': user['Last_name'],
                'role': user['Role']
            }
        })
    return jsonify({"message": "Invalid password"}), 401

@bp.route('/user', methods=['GET'])
@jwt_required()
def get_user():
    return jsonify(get_jwt_identity())