import sqlite3
import os
import bcrypt

DB_NAME = "contabilidade.db"

def initialize_db():
    """Initializes the database with required tables."""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # Create Users table if not exists (matching existing schema)
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash BLOB NOT NULL,
            is_admin BOOLEAN DEFAULT 0,
            permissions TEXT
        )
    ''')

    # Check if admin user exists, if not create one
    cursor.execute('SELECT count(*) FROM users')
    if cursor.fetchone()[0] == 0:
        # Default admin: admin/admin
        hashed_password = bcrypt.hashpw("admin".encode('utf-8'), bcrypt.gensalt()) # Keep as bytes for BLOB
        cursor.execute(
            'INSERT INTO users (username, password_hash, permissions, is_admin) VALUES (?, ?, ?, ?)',
            ("admin", hashed_password, "all", True)
        )
        print("Default admin user created.")

    conn.commit()
    conn.close()

def get_db_connection():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn
