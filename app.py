import os
import sys
from functools import wraps
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_file
import database
import excel_service
import word_service
import google_sheet_service

app = Flask(__name__)
app.secret_key = 'hunsen_angprey_social_assistance_secret_key_2026'

# Ensure database is initialized
database.init_db()

def get_current_school():
    conn = database.get_db_connection()
    school = conn.cursor().execute('SELECT * FROM school_settings LIMIT 1').fetchone()
    conn.close()
    return dict(school) if school else {}

def get_academic_years():
    conn = database.get_db_connection()
    years = [dict(r) for r in conn.cursor().execute('SELECT * FROM academic_years ORDER BY id DESC').fetchall()]
    conn.close()
    return years

def get_active_year():
    if 'selected_year' in session and session['selected_year']:
        return session['selected_year']
    years = get_academic_years()
    for y in years:
        if y.get('is_current'):
            return y['year_name']
    return years[0]['year_name'] if years else '២០២៥-២០២៦'

@app.context_processor
def inject_globals():
    return {
        'current_user': session.get('user'),
        'school_settings': get_current_school(),
        'academic_years': get_academic_years(),
        'active_year': get_active_year()
    }

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user' not in session:
            flash('សូមចូលគណនីជាមុនសិន!', 'warning')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def role_required(allowed_roles):
    def decorator(f):
        @wraps(f)
        def decorated_function(*args, **kwargs):
            if 'user' not in session:
                return redirect(url_for('login'))
            user_role = session['user'].get('role')
            if user_role not in allowed_roles:
                flash('លោកអ្នកមិនមានសិទ្ធិចូលដំណើរការផ្នែកនេះទេ!', 'danger')
                return redirect(url_for('index'))
            return f(*args, **kwargs)
        return decorated_function
    return decorator

# --- AUTH ROUTES ---
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '').strip()
        hashed = database.hash_password(password)

        conn = database.get_db_connection()
        user = conn.cursor().execute(
            'SELECT * FROM users WHERE username = ? AND password = ? AND is_active = 1',
            (username, hashed)
        ).fetchone()
        conn.close()

        if user:
            user_dict = dict(user)
            user_dict.pop('password', None)
            session['user'] = user_dict
            session['selected_year'] = get_active_year()
            flash(f'ស្វាគមន៍ការវិលត្រឡប់មកវិញ {user_dict["full_name"]}!', 'success')
            return redirect(url_for('index'))
        else:
            flash('ឈ្មោះគណនី ឬ ពាក្យសម្ងាត់មិនត្រឹមត្រូវឡើយ!', 'danger')

    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    flash('អ្នកបានចាកចេញដោយជោគជ័យ។', 'info')
    return redirect(url_for('login'))

@app.route('/set-year/<year_name>')
@login_required
def set_year(year_name):
    session['selected_year'] = year_name
    flash(f'បានប្តូរទៅឆ្នាំសិក្សា៖ {year_name}', 'info')
    return redirect(request.referrer or url_for('index'))

# --- MAIN DASHBOARD / STUDENTS ---
@app.route('/')
@login_required
def index():
    cur_year = get_active_year()
    grade = request.args.get('grade', '')
    card_type = request.args.get('card_type', '')
    gender = request.args.get('gender', '')
    reg_status = request.args.get('reg_status', '')
    search = request.args.get('q', '').strip()

    conn = database.get_db_connection()
    # Fetch all students for the active year so the browser can perform instant live filtering without reloads
    query = 'SELECT * FROM students WHERE academic_year = ? ORDER BY id DESC'
    students = [dict(r) for r in conn.cursor().execute(query, (cur_year,)).fetchall()]

    # Stats for current active year
    total_count = len(students)
    level1_count = sum(1 for s in students if 'កម្រិត១' in s.get('equity_card_type', ''))
    level2_count = sum(1 for s in students if 'កម្រិត២' in s.get('equity_card_type', ''))
    female_count = sum(1 for s in students if s.get('gender') == 'ស្រី')
    male_count = sum(1 for s in students if s.get('gender') == 'ប្រុស')

    conn.close()

    return render_template('dashboard.html',
                           students=students,
                           total_count=total_count,
                           level1_count=level1_count,
                           level2_count=level2_count,
                           female_count=female_count,
                           male_count=male_count,
                           selected_grade=grade,
                           selected_card=card_type,
                           selected_gender=gender,
                           selected_reg_status=reg_status,
                           search_query=search)

@app.route('/api/search-students')
@login_required
def api_search_students():
    q = request.args.get('q', '').strip()
    if not q:
        return jsonify([])
    conn = database.get_db_connection()
    cursor = conn.cursor()
    query = f"%{q}%"
    rows = cursor.execute('''
        SELECT id, student_code, full_name, gender, dob, pob,
               father_name, father_job, mother_name, mother_job,
               current_address, guardian_name, guardian_job, guardian_relation,
               grade, academic_year, equity_card_type, equity_card_no,
               phone, note, registration_status
        FROM students
        WHERE full_name LIKE ? OR student_code LIKE ? OR equity_card_no LIKE ? OR phone LIKE ?
        ORDER BY academic_year DESC, id DESC LIMIT 25
    ''', (query, query, query, query)).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])

@app.route('/api/check-duplicate', methods=['POST'])
@login_required
def api_check_duplicate():
    data = request.get_json() or request.form
    full_name = data.get('full_name', '').strip()
    gender = data.get('gender', '').strip()
    dob_raw = data.get('dob', '').strip()
    formatted_dob = database.format_khmer_dob(dob_raw)
    card_no = data.get('equity_card_no', '').strip()
    cur_year = get_active_year()

    if not full_name:
        return jsonify({'duplicate': False})

    conn = database.get_db_connection()
    cursor = conn.cursor()

    # Search duplicate matching all 4 fields simultaneously: full_name, gender, dob, equity_card_no
    # 1. Match in current academic year
    row_same = cursor.execute('''
        SELECT id, student_code, full_name, gender, dob, grade, academic_year, equity_card_no
        FROM students
        WHERE TRIM(full_name) = ? AND gender = ? AND dob = ? AND TRIM(equity_card_no) = ? AND academic_year = ?
    ''', (full_name, gender, formatted_dob, card_no, cur_year)).fetchone()

    # 2. Match in any academic year
    row_any = cursor.execute('''
        SELECT id, student_code, full_name, gender, dob, grade, academic_year, equity_card_no
        FROM students
        WHERE TRIM(full_name) = ? AND gender = ? AND dob = ? AND TRIM(equity_card_no) = ?
        ORDER BY academic_year DESC, id DESC LIMIT 1
    ''', (full_name, gender, formatted_dob, card_no)).fetchone()

    conn.close()

    if row_same:
        st = dict(row_same)
        return jsonify({
            'duplicate': True,
            'scope': 'same_year',
            'student': st,
            'message': f"រកឃើញទិន្នន័យស្ទួនទាំងស្រុង! សិស្សឈ្មោះ «{full_name}» ភេទ {gender} ថ្ងៃខែឆ្នាំកំណើត {formatted_dob} និងលេខប័ណ្ណ «{card_no or 'គ្មាន'}» បានចុះឈ្មោះក្នុងឆ្នាំសិក្សា {cur_year} (ថ្នាក់ទី {st.get('grade')}) នេះរួចរាល់ហើយ មិនអាចបញ្ចូលស្ទួនបានទេ!"
        })
    elif row_any:
        st = dict(row_any)
        return jsonify({
            'duplicate': True,
            'scope': 'other_year',
            'student': st,
            'message': f"សិស្សឈ្មោះ «{full_name}» ភេទ {gender} ថ្ងៃខែឆ្នាំកំណើត {formatted_dob} និងលេខប័ណ្ណ «{card_no or 'គ្មាន'}» ធ្លាប់មានរួចហើយក្នុងប្រព័ន្ធកាលពីឆ្នាំសិក្សា {st.get('academic_year')} (ថ្នាក់ទី {st.get('grade')})។ អ្នកអាចជ្រើសរើសជា «បច្ចុប្បន្នភាព» ដើម្បីទាញយកទិន្នន័យមកប្រើប្រាស់បន្ត។"
        })

    return jsonify({'duplicate': False})

@app.route('/student/add', methods=['POST'])
@login_required
def add_student():
    data = request.form
    is_ajax = request.headers.get('X-Requested-With') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', '')
    conn = database.get_db_connection()
    cursor = conn.cursor()

    full_name = data.get('full_name', '').strip()
    gender = data.get('gender', 'ស្រី').strip()
    formatted_dob = database.format_khmer_dob(data.get('dob', '').strip())
    card_no = data.get('equity_card_no', '').strip()
    academic_year = data.get('academic_year', get_active_year()).strip() or get_active_year()

    # Check duplicate matching all 4 fields simultaneously in the target academic year
    dup = cursor.execute('''
        SELECT id, student_code, full_name, gender, dob, grade, academic_year, equity_card_no
        FROM students
        WHERE TRIM(full_name) = ? AND gender = ? AND dob = ? AND TRIM(equity_card_no) = ? AND academic_year = ?
    ''', (full_name, gender, formatted_dob, card_no, academic_year)).fetchone()

    if dup:
        conn.close()
        dup_dict = dict(dup)
        err_msg = f"មិនអាចរក្សាទុកបានទេ! រកឃើញទិន្នន័យស្ទួនទាំងស្រុង (ឈ្មោះ: «{full_name}», ភេទ: {gender}, ថ្ងៃខែកំណើត: {formatted_dob}, លេខប័ណ្ណ: «{card_no or 'គ្មាន'}») បានចុះឈ្មោះក្នុងឆ្នាំសិក្សា {academic_year} (ថ្នាក់ទី {dup_dict.get('grade')}) រួចរាល់ហើយ!"
        if is_ajax:
            return jsonify({'success': False, 'duplicate': True, 'message': err_msg}), 409
        flash(err_msg, 'danger')
        return redirect(url_for('index'))

    student_code = data.get('student_code', '').strip()
    if not student_code:
        # Auto generate code
        count = cursor.execute('SELECT COUNT(*) FROM students').fetchone()[0]
        student_code = f"STU-{get_active_year()[-4:]}-{count + 1:04d}"

    registration_status = data.get('registration_status', 'ចុះឈ្មោះថ្មី').strip()
    if not registration_status:
        registration_status = 'ចុះឈ្មោះថ្មី'

    cursor.execute('''
    INSERT INTO students (
        student_code, full_name, gender, dob, pob,
        father_name, father_job, mother_name, mother_job,
        current_address, guardian_name, guardian_job, guardian_relation,
        grade, academic_year, equity_card_type, equity_card_no, phone, note,
        cert_solar_date, cert_lunar_date, cert_solar_raw, registration_status, created_by
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ''', (
        student_code,
        full_name,
        gender,
        formatted_dob,
        data.get('pob', '').strip(),
        data.get('father_name', '').strip(),
        data.get('father_job', 'កសិករ').strip(),
        data.get('mother_name', '').strip(),
        data.get('mother_job', 'កសិករ').strip(),
        data.get('current_address', '').strip(),
        data.get('guardian_name', '').strip(),
        data.get('guardian_job', 'កសិករ').strip(),
        data.get('guardian_relation', 'ឪពុក').strip(),
        data.get('grade', '១០').strip(),
        academic_year,
        data.get('equity_card_type', 'បណ្ណសមធម៌ក្រីក្រ កម្រិត២'),
        card_no,
        data.get('phone', '').strip(),
        data.get('note', '').strip(),
        data.get('cert_solar_date', '').strip(),
        data.get('cert_lunar_date', '').strip(),
        data.get('cert_solar_raw', '').strip(),
        registration_status,
        session['user']['username']
    ))
    new_id = cursor.lastrowid
    conn.commit()

    # Auto sync to local Excel file
    all_students = [dict(r) for r in cursor.execute('SELECT * FROM students').fetchall()]
    conn.close()
    try:
        excel_service.export_students_to_excel(all_students)
    except Exception as e:
        print("Excel auto-sync error:", e)

    success_msg = f"បានបញ្ចូលទិន្នន័យសិស្ស «{full_name}» ដោយជោគជ័យ!"

    if is_ajax:
        # Calculate next student code for fast consecutive entries
        conn = database.get_db_connection()
        total_st = conn.cursor().execute('SELECT COUNT(*) FROM students').fetchone()[0]
        next_code = f"STU-{get_active_year()[-4:]}-{total_st + 1:04d}"
        conn.close()
        return jsonify({
            'success': True,
            'message': success_msg,
            'student_id': new_id,
            'student_name': full_name,
            'next_student_code': next_code
        })

    flash(success_msg, 'success')
    return redirect(url_for('index'))

@app.route('/student/<int:student_id>/edit', methods=['POST'])
@login_required
def edit_student(student_id):
    data = request.form
    conn = database.get_db_connection()
    cursor = conn.cursor()

    formatted_dob = database.format_khmer_dob(data.get('dob', '').strip())
    registration_status = data.get('registration_status', 'ចុះឈ្មោះថ្មី').strip()
    if not registration_status:
        registration_status = 'ចុះឈ្មោះថ្មី'

    cursor.execute('''
    UPDATE students SET
        student_code = ?, full_name = ?, gender = ?, dob = ?, pob = ?,
        father_name = ?, father_job = ?, mother_name = ?, mother_job = ?,
        current_address = ?, guardian_name = ?, guardian_job = ?, guardian_relation = ?,
        grade = ?, academic_year = ?, equity_card_type = ?, equity_card_no = ?,
        phone = ?, note = ?, cert_solar_date = ?, cert_lunar_date = ?, cert_solar_raw = ?,
        registration_status = ?, updated_at = CURRENT_TIMESTAMP
    WHERE id = ?
    ''', (
        data.get('student_code', '').strip(),
        data.get('full_name', '').strip(),
        data.get('gender', 'ស្រី'),
        formatted_dob,
        data.get('pob', '').strip(),
        data.get('father_name', '').strip(),
        data.get('father_job', 'កសិករ').strip(),
        data.get('mother_name', '').strip(),
        data.get('mother_job', 'កសិករ').strip(),
        data.get('current_address', '').strip(),
        data.get('guardian_name', '').strip(),
        data.get('guardian_job', 'កសិករ').strip(),
        data.get('guardian_relation', 'ឪពុក').strip(),
        data.get('grade', '១០').strip(),
        data.get('academic_year', get_active_year()),
        data.get('equity_card_type', 'បណ្ណសមធម៌ក្រីក្រ កម្រិត២'),
        data.get('equity_card_no', '').strip(),
        data.get('phone', '').strip(),
        data.get('note', '').strip(),
        data.get('cert_solar_date', '').strip(),
        data.get('cert_lunar_date', '').strip(),
        data.get('cert_solar_raw', '').strip(),
        registration_status,
        student_id
    ))
    conn.commit()

    # Auto sync to local Excel file
    all_students = [dict(r) for r in cursor.execute('SELECT * FROM students').fetchall()]
    conn.close()
    try:
        excel_service.export_students_to_excel(all_students)
    except Exception as e:
        print("Excel auto-sync error:", e)

    flash('បានកែប្រែទិន្នន័យសិស្សដោយជោគជ័យ!', 'success')
    return redirect(url_for('index'))

@app.route('/student/<int:student_id>/delete', methods=['POST'])
@login_required
@role_required(['SuperAdmin', 'Admin'])
def delete_student(student_id):
    conn = database.get_db_connection()
    cursor = conn.cursor()
    cursor.execute('DELETE FROM students WHERE id = ?', (student_id,))
    conn.commit()

    all_students = [dict(r) for r in cursor.execute('SELECT * FROM students').fetchall()]
    conn.close()
    try:
        excel_service.export_students_to_excel(all_students)
    except Exception as e:
        print("Excel auto-sync error:", e)

    flash('បានលុបទិន្នន័យសិស្សដោយជោគជ័យ!', 'info')
    return redirect(url_for('index'))

# --- CERTIFICATE VIEW & WORD EXPORT ---
@app.route('/student/<int:student_id>/certificate')
@login_required
def view_certificate(student_id):
    conn = database.get_db_connection()
    student = conn.cursor().execute('SELECT * FROM students WHERE id = ?', (student_id,)).fetchone()
    conn.close()
    if not student:
        flash('រកមិនឃើញទិន្នន័យសិស្សឡើយ!', 'danger')
        return redirect(url_for('index'))

    return render_template('certificate_view.html', student=dict(student))

@app.route('/student/<int:student_id>/download-docx')
@login_required
def download_certificate_docx(student_id):
    conn = database.get_db_connection()
    student = conn.cursor().execute('SELECT * FROM students WHERE id = ?', (student_id,)).fetchone()
    school = conn.cursor().execute('SELECT * FROM school_settings LIMIT 1').fetchone()
    conn.close()

    if not student:
        flash('រកមិនឃើញទិន្នន័យសិស្សឡើយ!', 'danger')
        return redirect(url_for('index'))

    custom_lunar = request.args.get('lunar')
    custom_date = request.args.get('solar')

    filename = f"លិខិតបញ្ជាក់_{dict(student)['full_name']}_{dict(student)['academic_year']}.docx"
    output_path = os.path.join(database.DB_DIR, 'exports', filename)

    word_service.generate_certificate_docx(
        dict(student),
        dict(school) if school else {},
        output_path,
        custom_date=custom_date,
        custom_lunar=custom_lunar
    )

    return send_file(output_path, as_attachment=True, download_name=filename)

# --- REPORTS & STATISTICS ---
@app.route('/reports')
@login_required
def reports():
    cur_year = get_active_year()
    conn = database.get_db_connection()
    students = [dict(r) for r in conn.cursor().execute(
        'SELECT * FROM students WHERE academic_year = ? ORDER BY grade ASC, full_name ASC',
        (cur_year,)
    ).fetchall()]
    conn.close()

    # Calculations
    total = len(students)
    female = sum(1 for s in students if s.get('gender') == 'ស្រី')
    male = total - female
    level1 = sum(1 for s in students if 'កម្រិត១' in s.get('equity_card_type', ''))
    level2 = sum(1 for s in students if 'កម្រិត២' in s.get('equity_card_type', ''))

    # Grade breakdown
    grade_stats = {}
    for g in ['៧', '៨', '៩', '១០', '១១', '១២']:
        g_stus = [s for s in students if s.get('grade') == g]
        grade_stats[g] = {
            'total': len(g_stus),
            'female': sum(1 for s in g_stus if s.get('gender') == 'ស្រី'),
            'lvl1': sum(1 for s in g_stus if 'កម្រិត១' in s.get('equity_card_type', '')),
            'lvl2': sum(1 for s in g_stus if 'កម្រិត២' in s.get('equity_card_type', ''))
        }

    return render_template('reports.html',
                           students=students,
                           total=total,
                           female=female,
                           male=male,
                           level1=level1,
                           level2=level2,
                           grade_stats=grade_stats,
                           active_year=cur_year)

# --- EXCEL EXPORT & IMPORT ---
@app.route('/export/excel')
@login_required
def export_excel():
    cur_year = get_active_year()
    all_years = request.args.get('all', '0') == '1'

    conn = database.get_db_connection()
    if all_years:
        students = [dict(r) for r in conn.cursor().execute('SELECT * FROM students ORDER BY academic_year DESC, grade ASC, id DESC').fetchall()]
        filename = "ទិន្នន័យសិស្សជំនួយសង្គម_គ្រប់ឆ្នាំសិក្សា.xlsx"
        title = "បញ្ជីរាយនាមសិស្សទទួលកម្មវិធីជំនួយសង្គម (គ្រប់ឆ្នាំសិក្សា)"
    else:
        students = [dict(r) for r in conn.cursor().execute('SELECT * FROM students WHERE academic_year = ? ORDER BY grade ASC, id DESC', (cur_year,)).fetchall()]
        filename = f"ទិន្នន័យសិស្សជំនួយសង្គម_ឆ្នាំ_{cur_year}.xlsx"
        title = f"បញ្ជីរាយនាមសិស្សទទួលកម្មវិធីជំនួយសង្គម ឆ្នាំសិក្សា {cur_year}"
    conn.close()

    output_path = os.path.join(database.DB_DIR, 'exports', filename)
    excel_service.export_students_to_excel(students, output_path, title=title)
    return send_file(output_path, as_attachment=True, download_name=filename)

@app.route('/import/excel', methods=['POST'])
@login_required
@role_required(['SuperAdmin', 'Admin'])
def import_excel():
    if 'excel_file' not in request.files:
        flash('សូមជ្រើសរើសឯកសារ Excel!', 'warning')
        return redirect(url_for('index'))

    file = request.files['excel_file']
    if file.filename == '':
        flash('មិនបានជ្រើសរើសឯកសារឡើយ!', 'warning')
        return redirect(url_for('index'))

    upload_path = os.path.join(database.DB_DIR, 'exports', 'uploaded_import.xlsx')
    file.save(upload_path)

    try:
        students = excel_service.import_students_from_excel(
            upload_path,
            academic_year_default=get_active_year(),
            created_by=session['user']['username']
        )
        if not students:
            flash('មិនមានទិន្នន័យសិស្សត្រឹមត្រូវនៅក្នុងឯកសារ Excel នោះទេ!', 'danger')
            return redirect(url_for('index'))

        conn = database.get_db_connection()
        cursor = conn.cursor()
        inserted_count = 0
        for s in students:
            cursor.execute('''
            INSERT INTO students (
                student_code, full_name, gender, dob, pob,
                father_name, father_job, mother_name, mother_job,
                current_address, guardian_name, guardian_job, guardian_relation,
                grade, academic_year, equity_card_type, equity_card_no, phone, note, created_by
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                s['student_code'], s['full_name'], s['gender'], s['dob'], s['pob'],
                s['father_name'], s['father_job'], s['mother_name'], s['mother_job'],
                s['current_address'], s['guardian_name'], s['guardian_job'], s['guardian_relation'],
                s['grade'], s['academic_year'], s['equity_card_type'], s['equity_card_no'],
                s['phone'], s['note'], s['created_by']
            ))
            inserted_count += 1
        conn.commit()

        # Update local Excel store
        all_students = [dict(r) for r in cursor.execute('SELECT * FROM students').fetchall()]
        conn.close()
        excel_service.export_students_to_excel(all_students)

        flash(f'បាននាំចូលទិន្នន័យសិស្សចំនួន {inserted_count} នាក់ពី Excel ដោយជោគជ័យ!', 'success')
    except Exception as e:
        flash(f'មានបញ្ហាក្នុងការនាំចូល៖ {str(e)}', 'danger')

    return redirect(url_for('index'))

# --- GOOGLE SHEETS SYNC ---
@app.route('/sync/google-sheet', methods=['POST'])
@login_required
def sync_google():
    school = get_current_school()
    webhook_url = school.get('google_sheet_url', '').strip()
    if not webhook_url:
        return jsonify({'success': False, 'message': 'សូមកំណត់ Google Apps Script Webhook URL នៅក្នុងទំព័រ Settings ជាមុនសិន!'})

    conn = database.get_db_connection()
    students = [dict(r) for r in conn.cursor().execute('SELECT * FROM students ORDER BY academic_year DESC, id ASC').fetchall()]
    conn.close()

    result = google_sheet_service.sync_to_google_sheet(webhook_url, students, school.get('school_name'))
    return jsonify(result)

@app.route('/pull/google-sheet', methods=['POST'])
@login_required
@role_required(['SuperAdmin', 'Admin'])
def pull_google():
    school = get_current_school()
    webhook_url = school.get('google_sheet_url', '').strip()
    if not webhook_url:
        return jsonify({'success': False, 'message': 'សូមកំណត់ Google Apps Script Webhook URL ជាមុនសិន!'})

    res = google_sheet_service.pull_from_google_sheet(webhook_url)
    if not res.get('success'):
        return jsonify(res)

    students = res.get('students', [])
    if not students:
        return jsonify({'success': False, 'message': 'ពុំមានទិន្នន័យសិស្សនៅលើ Google Sheet ឡើយ!'})

    conn = database.get_db_connection()
    cursor = conn.cursor()
    # Insert or update
    synced_count = 0
    for s in students:
        if not s.get('full_name'):
            continue
        cursor.execute('''
        INSERT INTO students (
            student_code, full_name, gender, dob, pob,
            father_name, father_job, mother_name, mother_job,
            current_address, guardian_name, guardian_job, guardian_relation,
            grade, academic_year, equity_card_type, equity_card_no, phone, note, created_by
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            s.get('student_code', ''), s.get('full_name'), s.get('gender', 'ស្រី'), s.get('dob', ''), s.get('pob', ''),
            s.get('father_name', ''), s.get('father_job', 'កសិករ'), s.get('mother_name', ''), s.get('mother_job', 'កសិករ'),
            s.get('current_address', ''), s.get('guardian_name', ''), s.get('guardian_job', 'កសិករ'), s.get('guardian_relation', 'ឪពុក'),
            s.get('grade', '១០'), s.get('academic_year', get_active_year()), s.get('equity_card_type', 'បណ្ណសមធម៌ក្រីក្រ កម្រិត២'),
            s.get('equity_card_no', ''), s.get('phone', ''), s.get('note', ''), 'GoogleSheetSync'
        ))
        synced_count += 1
    conn.commit()

    all_stus = [dict(r) for r in cursor.execute('SELECT * FROM students').fetchall()]
    conn.close()
    excel_service.export_students_to_excel(all_stus)

    return jsonify({'success': True, 'message': f'បានទាញទិន្នន័យចំនួន {synced_count} នាក់ពី Google Sheet ចូល Local Database ដោយជោគជ័យ!'})

# --- SETTINGS & ACADEMIC YEARS ---
@app.route('/settings', methods=['GET', 'POST'])
@login_required
@role_required(['SuperAdmin', 'Admin'])
def settings():
    conn = database.get_db_connection()
    cursor = conn.cursor()

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'update_school':
            cursor.execute('''
            UPDATE school_settings SET
                province_dept = ?, school_name = ?, school_short_name = ?,
                principal_title = ?, principal_name = ?, default_lunar_day = ?,
                google_sheet_url = ?
            WHERE id = 1
            ''', (
                request.form.get('province_dept', '').strip(),
                request.form.get('school_name', '').strip(),
                request.form.get('school_short_name', '').strip(),
                request.form.get('principal_title', 'នាយកសាលា').strip(),
                request.form.get('principal_name', '').strip(),
                request.form.get('default_lunar_day', '').strip(),
                request.form.get('google_sheet_url', '').strip()
            ))
            conn.commit()
            flash('បានរក្សាទុកការកំណត់ព័ត៌មានសាលាដោយជោគជ័យ!', 'success')

        elif action == 'add_year':
            year_name = request.form.get('year_name', '').strip()
            if year_name:
                try:
                    cursor.execute('INSERT INTO academic_years (year_name, is_current) VALUES (?, 0)', (year_name,))
                    conn.commit()
                    flash(f'បានបង្កើតឆ្នាំសិក្សាថ្មី៖ {year_name} ដោយជោគជ័យ!', 'success')
                except Exception:
                    flash('ឆ្នាំសិក្សានេះមានរួចហើយ!', 'warning')

        elif action == 'set_current_year':
            year_id = request.form.get('year_id')
            cursor.execute('UPDATE academic_years SET is_current = 0')
            cursor.execute('UPDATE academic_years SET is_current = 1 WHERE id = ?', (year_id,))
            conn.commit()
            flash('បានកំណត់ឆ្នាំសិក្សាលំនាំដើម!', 'info')

        elif action == 'update_logo':
            logo_file = request.files.get('logo_file')
            if logo_file and logo_file.filename:
                ext = os.path.splitext(logo_file.filename)[1].lower()
                if ext in ['.png', '.jpg', '.jpeg', '.svg', '.webp', '.ico']:
                    upload_dir = os.path.join(app.root_path, 'static', 'uploads')
                    os.makedirs(upload_dir, exist_ok=True)
                    filename = f"logo_{int(datetime.datetime.now().timestamp())}{ext}"
                    filepath = os.path.join(upload_dir, filename)
                    logo_file.save(filepath)
                    logo_url = f"/static/uploads/{filename}"
                    cursor.execute('UPDATE school_settings SET logo_url = ? WHERE id = 1', (logo_url,))
                    conn.commit()
                    flash('បានផ្លាស់ប្តូរ Logo របស់ប្រព័ន្ធដោយជោគជ័យ!', 'success')
                else:
                    flash('ប្រភេទរូបភាពមិនត្រឹមត្រូវ! សូមជ្រើសរើសឯកសារ PNG, JPG, WEBP ឬ SVG', 'danger')
            else:
                flash('សូមជ្រើសរើសឯកសាររូបភាព Logo ជាមុនសិន!', 'warning')

        elif action == 'reset_logo':
            cursor.execute("UPDATE school_settings SET logo_url = '/static/images/logo.png' WHERE id = 1")
            conn.commit()
            flash('បានកំណត់ Logo ដើមឡើងវិញដោយជោគជ័យ!', 'info')

        return redirect(url_for('settings'))

    school = conn.cursor().execute('SELECT * FROM school_settings LIMIT 1').fetchone()
    years = [dict(r) for r in conn.cursor().execute('SELECT * FROM academic_years ORDER BY id DESC').fetchall()]
    conn.close()

    # Read google_apps_script.js content to display for copy-pasting
    script_path = os.path.join(os.path.dirname(__file__), 'google_apps_script.js')
    script_code = ""
    if os.path.exists(script_path):
        with open(script_path, 'r', encoding='utf-8') as f:
            script_code = f.read()

    return render_template('settings.html', school=dict(school) if school else {}, years=years, script_code=script_code)

# --- USERS MANAGEMENT (SUPERADMIN ONLY) ---
@app.route('/users', methods=['GET', 'POST'])
@login_required
@role_required(['SuperAdmin'])
def users_management():
    conn = database.get_db_connection()
    cursor = conn.cursor()

    if request.method == 'POST':
        action = request.form.get('action')
        if action == 'add_user':
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '').strip()
            full_name = request.form.get('full_name', '').strip()
            role = request.form.get('role', 'Staff')
            phone = request.form.get('phone', '').strip()

            if username and password and full_name:
                try:
                    cursor.execute('''
                    INSERT INTO users (username, password, full_name, role, phone)
                    VALUES (?, ?, ?, ?, ?)
                    ''', (username, database.hash_password(password), full_name, role, phone))
                    conn.commit()
                    flash(f'បានបង្កើតគណនីថ្មី ({username}) ដោយជោគជ័យ!', 'success')
                except Exception as e:
                    flash(f'មិនអាចបង្កើតគណនីបានទេ៖ ឈ្មោះគណនីនេះអាចមានរួចហើយ!', 'danger')

        elif action == 'edit_user':
            user_id = request.form.get('user_id')
            username = request.form.get('username', '').strip()
            full_name = request.form.get('full_name', '').strip()
            role = request.form.get('role', 'Staff')
            phone = request.form.get('phone', '').strip()
            new_password = request.form.get('new_password', '').strip()

            if not username:
                flash('ឈ្មោះគណនី (Username) មិនអាចទទេបានឡើយ!', 'warning')
                return redirect(url_for('users_management'))

            # Check if another user already has this username
            existing = cursor.execute('SELECT id FROM users WHERE username = ? AND id != ?', (username, user_id)).fetchone()
            if existing:
                flash(f'ឈ្មោះគណនី «{username}» នេះមានអ្នកប្រើប្រាស់រួចហើយ! សូមជ្រើសរើសឈ្មោះផ្សេង។', 'danger')
                return redirect(url_for('users_management'))

            if new_password:
                cursor.execute('''
                UPDATE users SET username = ?, full_name = ?, role = ?, phone = ?, password = ? WHERE id = ?
                ''', (username, full_name, role, phone, database.hash_password(new_password), user_id))
            else:
                cursor.execute('''
                UPDATE users SET username = ?, full_name = ?, role = ?, phone = ? WHERE id = ?
                ''', (username, full_name, role, phone, user_id))
            conn.commit()

            # If the user edited their own logged-in account, update current session info
            if session.get('user') and int(user_id) == session['user']['id']:
                session['user']['username'] = username
                session['user']['full_name'] = full_name
                session['user']['role'] = role
                session.modified = True

            flash('បានកែប្រែព័ត៌មានគណនីដោយជោគជ័យ!', 'success')

        elif action == 'toggle_active':
            user_id = request.form.get('user_id')
            cursor.execute('UPDATE users SET is_active = 1 - is_active WHERE id = ?', (user_id,))
            conn.commit()
            flash('បានប្តូរស្ថានភាពគណនី!', 'info')

        elif action == 'delete_user':
            user_id = request.form.get('user_id')
            if int(user_id) == session['user']['id']:
                flash('មិនអាចលុបគណនីផ្ទាល់ខ្លួនដែលកំពុងប្រើប្រាស់បានទេ!', 'danger')
            else:
                cursor.execute('DELETE FROM users WHERE id = ?', (user_id,))
                conn.commit()
                flash('បានលុបគណនីដោយជោគជ័យ!', 'info')

        return redirect(url_for('users_management'))

    all_users = [dict(r) for r in cursor.execute('SELECT id, username, full_name, role, phone, is_active, created_at FROM users ORDER BY id ASC').fetchall()]
    conn.close()
    return render_template('users.html', users=all_users)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Starting Social Assistance Management System on http://0.0.0.0:{port}")
    app.run(host='0.0.0.0', port=port, debug=False)
