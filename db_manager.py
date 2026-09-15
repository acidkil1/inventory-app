import sqlite3
from datetime import datetime
import openpyxl
from openpyxl import load_workbook
DB_NAME = "inventory.db"
def get_connection():
    conn = sqlite3.connect(DB_NAME)
    return conn
def init_db():
    conn = get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS categories (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS equipment (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            inv_number TEXT UNIQUE NOT NULL,
            name TEXT NOT NULL,
            model TEXT,
            category_id INTEGER,
            location TEXT,
            responsible TEXT,
            date_in DATE,
            status TEXT DEFAULT 'Активно',
            description TEXT,
            FOREIGN KEY (category_id) REFERENCES categories(id)
        )
    ''')
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            equipment_id INTEGER,
            action_type TEXT,
            date_action DATE,
            old_location TEXT,
            new_location TEXT,
            comment TEXT,
            FOREIGN KEY (equipment_id) REFERENCES equipment(id)
        )
    ''')
    default_cats = ['ПК', 'Периферия', 'Сетевое оборудование', 'Криптография', 'Сервер']
    for cat in default_cats:
        try:
            cursor.execute("INSERT INTO categories (name) VALUES (?)", (cat,))
        except sqlite3.IntegrityError:
            pass
    conn.commit()
    conn.close()
    print("База данных инициализирована.")
if __name__ == "__main__":
    init_db()