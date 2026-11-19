import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

EXCEL_LOCAL_PATH = os.path.join(os.path.dirname(__file__), 'data', 'ទិន្នន័យជំនួយសង្គម_Local.xlsx')

HEADERS = [
    'ល.រ',
    'អត្តលេខ',
    'សិស្សឈ្មោះ',
    'ភេទ',
    'ថ្ងៃខែឆ្នាំកំណើត',
    'ទីកន្លែងកំណើត',
    'ឪពុកឈ្មោះ',
    'មុខរបរឪពុក',
    'ម្តាយឈ្មោះ',
    'មុខរបរម្តាយ',
    'អាសយដ្ឋានបច្ចុប្បន្ន',
    'អាព្យាបាលឈ្មោះ',
    'មុខរបរអាព្យាបាល',
    'ត្រូវជា',
    'ថ្នាក់ទី',
    'ឆ្នាំសិក្សា',
    'ប្រវត្តិចុះឈ្មោះ',
    'ប្រភេទបណ្ណសមធម៌ក្រីក្រ',
    'លេខប័ណ្ណសមធម៌',
    'លេខទូរស័ព្ទ',
    'កំណត់សម្គាល់'
]

def export_students_to_excel(students, output_path=None, title="បញ្ជីរាយនាមសិស្សទទួលកម្មវិធីជំនួយសង្គម"):
    """
    Exports students list to an elegantly styled Excel workbook.
    """
    if output_path is None:
        output_path = EXCEL_LOCAL_PATH

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ទិន្នន័យសិស្ស"
    ws.views.sheetView[0].showGridLines = True

    # Styling elements
    font_title = Font(name="Khmer OS Muol Light", size=13, bold=True, color="1F4E79")
    font_sub = Font(name="Khmer OS Siemreap", size=11, bold=True, color="333333")
    font_header = Font(name="Khmer OS Siemreap", size=10, bold=True, color="FFFFFF")
    font_data = Font(name="Khmer OS Siemreap", size=10)
    font_bold = Font(name="Khmer OS Siemreap", size=10, bold=True)

    fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    fill_even = PatternFill(start_color="F2F7FA", end_color="F2F7FA", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )

    # 1. School Header Rows
    ws.merge_cells("A1:U1")
    ws["A1"] = "ព្រះរាជាណាចក្រកម្ពុជា ជាតិ សាសនា ព្រះមហាក្សត្រ"
    ws["A1"].font = font_sub
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A2:U2")
    ws["A2"] = title
    ws["A2"].font = font_title
    ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 28

    # 2. Table Headers (Row 4)
    row_idx = 4
    ws.row_dimensions[row_idx].height = 26
    for col_idx, header in enumerate(HEADERS, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=header)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # 3. Data Rows
    for i, s in enumerate(students, start=1):
        row_idx += 1
        ws.row_dimensions[row_idx].height = 22
        row_data = [
            i,
            s.get('student_code', ''),
            s.get('full_name', ''),
            s.get('gender', ''),
            s.get('dob', ''),
            s.get('pob', ''),
            s.get('father_name', ''),
            s.get('father_job', ''),
            s.get('mother_name', ''),
            s.get('mother_job', ''),
            s.get('current_address', ''),
            s.get('guardian_name', ''),
            s.get('guardian_job', ''),
            s.get('guardian_relation', ''),
            s.get('grade', ''),
            s.get('academic_year', ''),
            s.get('registration_status', 'ចុះឈ្មោះថ្មី'),
            s.get('equity_card_type', ''),
            s.get('equity_card_no', ''),
            s.get('phone', ''),
            s.get('note', '')
        ]

        is_even = (i % 2 == 0)
        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = font_bold if col_idx in (1, 3, 4, 15, 17, 18) else font_data
            if is_even:
                cell.fill = fill_even
            cell.border = thin_border

            # Alignment
            if col_idx in (1, 4, 15, 16, 17):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # 4. Auto adjust column widths
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row >= 4 and cell.value:
                # Approximate width for Khmer text
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(output_path)
    return output_path

def import_students_from_excel(file_path, academic_year_default='២០២៥-២០២៦', created_by='admin'):
    """
    Reads an Excel file and extracts student records.
    """
    wb = openpyxl.load_workbook(file_path, data_only=True)
    ws = wb.active

    # Find the header row
    header_row = None
    for r in range(1, 10):
        row_vals = [str(ws.cell(row=r, column=c).value or '').strip() for c in range(1, 25)]
        if any('សិស្សឈ្មោះ' in v or 'ឈ្មោះសិស្ស' in v or 'ឈ្មោះ' in v for v in row_vals):
            header_row = r
            break

    if not header_row:
        header_row = 4

    col_map = {}
    for c in range(1, 25):
        val = str(ws.cell(row=header_row, column=c).value or '').strip()
        if 'អត្តលេខ' in val: col_map['student_code'] = c
        elif 'សិស្សឈ្មោះ' in val or 'ឈ្មោះ' in val: col_map['full_name'] = c
        elif 'ភេទ' in val: col_map['gender'] = c
        elif 'ថ្ងៃខែឆ្នាំកំណើត' in val or 'កំណើត' in val: col_map['dob'] = c
        elif 'ទីកន្លែងកំណើត' in val: col_map['pob'] = c
        elif 'ឪពុក' in val and 'មុខរបរ' not in val: col_map['father_name'] = c
        elif 'មុខរបរឪពុក' in val: col_map['father_job'] = c
        elif 'ម្តាយ' in val and 'មុខរបរ' not in val: col_map['mother_name'] = c
        elif 'មុខរបរម្តាយ' in val: col_map['mother_job'] = c
        elif 'អាសយដ្ឋាន' in val or 'ទីលំនៅ' in val: col_map['current_address'] = c
        elif 'អាព្យាបាល' in val and 'មុខរបរ' not in val and 'ត្រូវជា' not in val: col_map['guardian_name'] = c
        elif 'មុខរបរអាព្យាបាល' in val: col_map['guardian_job'] = c
        elif 'ត្រូវជា' in val: col_map['guardian_relation'] = c
        elif 'ថ្នាក់' in val: col_map['grade'] = c
        elif 'ឆ្នាំសិក្សា' in val: col_map['academic_year'] = c
        elif 'ប្រវត្តិចុះឈ្មោះ' in val or 'ស្ថានភាពចុះឈ្មោះ' in val: col_map['registration_status'] = c
        elif 'បណ្ណ' in val or 'សមធម៌' in val: col_map['equity_card_type'] = c
        elif 'លេខប័ណ្ណ' in val: col_map['equity_card_no'] = c
        elif 'ទូរស័ព្ទ' in val: col_map['phone'] = c
        elif 'ចំណាំ' in val: col_map['note'] = c

    students = []
    max_row = ws.max_row
    for r in range(header_row + 1, max_row + 1):
        name_col = col_map.get('full_name', 3)
        name = str(ws.cell(row=r, column=name_col).value or '').strip()
        if not name or name == 'None':
            continue

        def get_val(key, default=''):
            c = col_map.get(key)
            if c:
                v = ws.cell(row=r, column=c).value
                if v is not None and str(v).strip() != 'None':
                    return str(v).strip()
            return default

        student = {
            'student_code': get_val('student_code', f"STU-{len(students)+1:03d}"),
            'full_name': name,
            'gender': get_val('gender', 'ស្រី'),
            'dob': get_val('dob', '២០០០-០១-០១'),
            'pob': get_val('pob', ''),
            'father_name': get_val('father_name', ''),
            'father_job': get_val('father_job', 'កសិករ'),
            'mother_name': get_val('mother_name', ''),
            'mother_job': get_val('mother_job', 'កសិករ'),
            'current_address': get_val('current_address', ''),
            'guardian_name': get_val('guardian_name', ''),
            'guardian_job': get_val('guardian_job', 'កសិករ'),
            'guardian_relation': get_val('guardian_relation', 'ឪពុក'),
            'grade': get_val('grade', '១០'),
            'academic_year': get_val('academic_year', academic_year_default),
            'registration_status': get_val('registration_status', 'ចុះឈ្មោះថ្មី'),
            'equity_card_type': get_val('equity_card_type', 'បណ្ណសមធម៌ក្រីក្រ កម្រិត២'),
            'equity_card_no': get_val('equity_card_no', ''),
            'phone': get_val('phone', ''),
            'note': get_val('note', ''),
            'created_by': created_by
        }
        students.append(student)

    return students
