import sqlite3
import os
import hashlib
from datetime import datetime

DB_DIR = os.path.join(os.path.dirname(__file__), 'data')
DB_PATH = os.path.join(DB_DIR, 'social_assistance.db')

def get_db_connection():
    os.makedirs(DB_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def hash_password(password):
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

KHMER_NUMS = {'0': '០', '1': '១', '2': '២', '3': '៣', '4': '៤', '5': '៥', '6': '៦', '7': '៧', '8': '៨', '9': '៩'}
KHMER_MONTHS_GREGORIAN = {
    1: 'មករា', 2: 'កុម្ភៈ', 3: 'មីនា', 4: 'មេសា',
    5: 'ឧសភា', 6: 'មិថុនា', 7: 'កក្កដា', 8: 'សីហា',
    9: 'កញ្ញា', 10: 'តុលា', 11: 'វិច្ឆិកា', 12: 'ធ្នូ'
}

def to_khmer_number(num):
    return ''.join(KHMER_NUMS.get(c, c) for c in str(num))

def format_khmer_dob(dob_str):
    """
    Converts formats like '11-04-1984', '1984-04-11', '11/04/1984'
    into authentic Khmer: 'ថ្ងៃទី១១ ខែមេសា ឆ្នាំ១៩៨៤'
    """
    if not dob_str:
        return ''
    dob_str = str(dob_str).strip()
    if 'ថ្ងៃទី' in dob_str and 'ខែ' in dob_str:
        return dob_str

    import re
    # Match dd-mm-yyyy or dd/mm/yyyy
    m1 = re.match(r'^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$', dob_str)
    if m1:
        d, m, y = int(m1.group(1)), int(m1.group(2)), int(m1.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            month_name = KHMER_MONTHS_GREGORIAN.get(m, str(m))
            return f"ថ្ងៃទី{to_khmer_number(d)} ខែ{month_name} ឆ្នាំ{to_khmer_number(y)}"

    # Match yyyy-mm-dd or yyyy/mm/dd
    m2 = re.match(r'^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$', dob_str)
    if m2:
        y, m, d = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        if 1 <= m <= 12 and 1 <= d <= 31:
            month_name = KHMER_MONTHS_GREGORIAN.get(m, str(m))
            return f"ថ្ងៃទី{to_khmer_number(d)} ខែ{month_name} ឆ្នាំ{to_khmer_number(y)}"

    return dob_str

def init_db():
    os.makedirs(DB_DIR, exist_ok=True)
    os.makedirs(os.path.join(DB_DIR, 'exports'), exist_ok=True)
    
    conn = get_db_connection()
    cursor = conn.cursor()

    # Table: Users (គណនីគ្រប់គ្រង)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL, -- 'SuperAdmin', 'Admin', 'Staff'
        phone TEXT,
        is_active INTEGER DEFAULT 1,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # Table: Academic Years (ឆ្នាំសិក្សា)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS academic_years (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year_name TEXT UNIQUE NOT NULL, -- e.g. '២០២៥-២០២៦', '២០២៦-២០២៧'
        is_current INTEGER DEFAULT 0,
        notes TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # Table: School Settings (ការកំណត់ព័ត៌មានសាលា)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS school_settings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        province_dept TEXT DEFAULT 'មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្តតាកែវ',
        school_name TEXT DEFAULT 'វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍',
        school_short_name TEXT DEFAULT 'វិ.ហ.ស អង្គប្រីយ៍',
        principal_title TEXT DEFAULT 'នាយកសាលា',
        principal_name TEXT DEFAULT '',
        default_lunar_day TEXT DEFAULT 'ថ្ងៃសុក្រ ៥កើត ខែបឋមាសាឍ ឆ្នាំមមី អដ្ឋស័ក ព.ស.២៥៧០',
        google_sheet_url TEXT DEFAULT '',
        auto_sync_google INTEGER DEFAULT 0,
        logo_url TEXT DEFAULT '/static/images/logo.png'
    )
    ''')
    try:
        cursor.execute("ALTER TABLE school_settings ADD COLUMN logo_url TEXT DEFAULT '/static/images/logo.png'")
    except Exception:
        pass
    cursor.execute("UPDATE school_settings SET logo_url = '/static/images/logo.png' WHERE logo_url IS NULL OR logo_url = ''")

    # Table: Students (សិស្សទទួលជំនួយសង្គម)
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS students (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        student_code TEXT,
        full_name TEXT NOT NULL,
        gender TEXT NOT NULL, -- 'ប្រុស' ឬ 'ស្រី'
        dob TEXT NOT NULL,
        pob TEXT, -- ទីកន្លែងកំណើត
        father_name TEXT,
        father_job TEXT DEFAULT 'កសិករ',
        mother_name TEXT,
        mother_job TEXT DEFAULT 'កសិករ',
        current_address TEXT,
        guardian_name TEXT,
        guardian_job TEXT DEFAULT 'កសិករ',
        guardian_relation TEXT DEFAULT 'ឪពុក',
        grade TEXT NOT NULL, -- '៧', '៨', '៩', '១០', '១១', '១២'
        academic_year TEXT NOT NULL, -- '២០២៥-២០២៦'
        equity_card_type TEXT NOT NULL, -- 'បណ្ណសមធម៌ក្រីក្រ កម្រិត១' ឬ 'បណ្ណសមធម៌ក្រីក្រ កម្រិត២'
        equity_card_no TEXT, -- លេខប័ណ្ណ
        registration_status TEXT DEFAULT 'ចុះឈ្មោះថ្មី', -- 'ចុះឈ្មោះថ្មី' ឬ 'បច្ចុប្បន្នភាព'
        phone TEXT,
        note TEXT,
        photo_filename TEXT,
        created_by TEXT,
        cert_solar_date TEXT,
        cert_lunar_date TEXT,
        cert_solar_raw TEXT,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
        updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    ''')

    # Migration: Add columns if they do not exist
    cursor.execute("PRAGMA table_info(students)")
    existing_cols = [row[1] for row in cursor.fetchall()]
    if 'cert_solar_date' not in existing_cols:
        cursor.execute("ALTER TABLE students ADD COLUMN cert_solar_date TEXT")
    if 'cert_lunar_date' not in existing_cols:
        cursor.execute("ALTER TABLE students ADD COLUMN cert_lunar_date TEXT")
    if 'cert_solar_raw' not in existing_cols:
        cursor.execute("ALTER TABLE students ADD COLUMN cert_solar_raw TEXT")
    if 'registration_status' not in existing_cols:
        cursor.execute("ALTER TABLE students ADD COLUMN registration_status TEXT DEFAULT 'ចុះឈ្មោះថ្មី'")

    # Check and insert default school settings
    cursor.execute('SELECT COUNT(*) FROM school_settings')
    if cursor.fetchone()[0] == 0:
        cursor.execute('''
        INSERT INTO school_settings (province_dept, school_name, school_short_name, principal_title, default_lunar_day)
        VALUES (?, ?, ?, ?, ?)
        ''', (
            'មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្តតាកែវ',
            'វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍',
            'វិ.ហ.ស អង្គប្រីយ៍',
            'នាយកសាលា',
            'ថ្ងៃសុក្រ ៥កើត ខែបឋមាសាឍ ឆ្នាំមមី អដ្ឋស័ក ព.ស.២៥៧០'
        ))

    # Check and insert default academic years
    cursor.execute('SELECT COUNT(*) FROM academic_years')
    if cursor.fetchone()[0] == 0:
        years = [
            ('២០២៤-២០២៥', 0),
            ('២០២៥-២០២៦', 1),
            ('២០២៦-២០២៧', 0),
            ('២០២៧-២០២៨', 0)
        ]
        for y, is_cur in years:
            cursor.execute('INSERT INTO academic_years (year_name, is_current) VALUES (?, ?)', (y, is_cur))

    # Check and insert default users: SuperAdmin, Admin, Staff (បុគ្គលិក)
    cursor.execute('SELECT COUNT(*) FROM users')
    if cursor.fetchone()[0] == 0:
        default_users = [
            ('superadmin', hash_password('admin123'), 'អភិបាលជាន់ខ្ពស់ (SuperAdmin)', 'SuperAdmin', '012 345 678'),
            ('admin', hash_password('admin123'), 'អ្នកគ្រប់គ្រងទូទៅ (Admin)', 'Admin', '098 765 432'),
            ('staff', hash_password('staff123'), 'បុគ្គលិកបញ្ចូលទិន្នន័យ (Staff)', 'Staff', '088 123 456')
        ]
        for u in default_users:
            cursor.execute('''
            INSERT INTO users (username, password, full_name, role, phone)
            VALUES (?, ?, ?, ?, ?)
            ''', u)

    # Check and insert sample student matching the provided PDF document
    cursor.execute('SELECT COUNT(*) FROM students')
    if cursor.fetchone()[0] == 0:
        sample_student = (
            'STU-2025-001',
            'កង ស្រីគីន',
            'ស្រី',
            'ថ្ងៃទី៣ ខែតុលា ឆ្នាំ២០១០',
            'ភូមិមហារាជ្យ ឃុំបឹងត្រាញ់ខាងត្បូង ស្រុកសំរោង ខេត្តតាកែវ',
            'មុី សុគន្ធ',
            'កសិករ',
            'សំ កាន',
            'កសិករ',
            'ភូមិមហារាជ្យ ឃុំបឹងត្រាញ់ខាងត្បូង ស្រុកសំរោង ខេត្តតាកែវ',
            'មុី សុគន្ធ',
            'កសិករ',
            'ឪពុក',
            '១០',
            '២០២៥-២០២៦',
            'បណ្ណសមធម៌ក្រីក្រ កម្រិត២',
            'IDP-TK-2025-089',
            '097 555 1234',
            'ទិន្នន័យគំរូនាំចូលពីលិខិតបញ្ជាក់គំរូ',
            'superadmin'
        )
        cursor.execute('''
        INSERT INTO students (
            student_code, full_name, gender, dob, pob, 
            father_name, father_job, mother_name, mother_job,
            current_address, guardian_name, guardian_job, guardian_relation,
            grade, academic_year, equity_card_type, equity_card_no, phone, note, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', sample_student)

    conn.commit()
    conn.close()
    print("Database initialized successfully.")

if __name__ == '__main__':
    init_db()
