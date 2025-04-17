import secrets

class Config:
    JWT_SECRET_KEY = secrets.token_hex(32)
    DATABASE = 'students_database_unique_names.db'