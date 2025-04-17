from flask import Flask
from flask_cors import CORS
from flask_jwt_extended import JWTManager
from config import Config

# Import blueprints
from routes import auth, attendance, letter_generator, holiday, upload_attendance, track_attendance, fee_letter, hostel_fee_letter, transport_fee_letter
from routes import warning, probation, academic_misconduct, suspension
from routes import placement, health, custom, administrative, extracurricular, academic

app = Flask(__name__)
app.config.from_object(Config)

CORS(app, supports_credentials=True)
jwt = JWTManager(app)

# Register blueprints
app.register_blueprint(auth.bp)
app.register_blueprint(attendance.bp)
app.register_blueprint(letter_generator.bp)
app.register_blueprint(holiday.bp)
app.register_blueprint(upload_attendance.bp)
app.register_blueprint(track_attendance.bp)
app.register_blueprint(fee_letter.bp)
app.register_blueprint(hostel_fee_letter.bp)
app.register_blueprint(transport_fee_letter.bp)
app.register_blueprint(warning.bp)
app.register_blueprint(probation.bp)
app.register_blueprint(suspension.bp)
app.register_blueprint(academic_misconduct.bp)
app.register_blueprint(placement.bp)
app.register_blueprint(health.bp)
app.register_blueprint(custom.bp)
app.register_blueprint(administrative.bp)
app.register_blueprint(extracurricular.bp)
app.register_blueprint(academic.bp)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5001)