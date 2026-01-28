import sqlite3
import os
import bcrypt

DB_NAME = "contabilidade.db"

def initialize_db():
    """Initializes the database with required tables."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # Create Users table if not exists
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            permissions TEXT,
            is_admin BOOLEAN DEFAULT 0
        )
    ''')

    # Check if admin user exists, if not create one
    cursor.execute('SELECT count(*) FROM users')
    if cursor.fetchone()[0] == 0:
        # Default admin: admin/admin
        hashed_password = bcrypt.hashpw("admin".encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        cursor.execute(
            'INSERT INTO users (username, password, permissions, is_admin) VALUES (?, ?, ?, ?)',
            ("admin", hashed_password, "all", True)
        )
        print("Default admin user created.")

    conn.commit()
    conn.close()

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn
