import os
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_cell_border(cell, **kwargs):
    """
    kwargs: top, bottom, left, right
    values: dict(sz=12, val='single', color='000000')
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in("w:tcBorders")
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        edge_data = kwargs.get(edge)
        if edge_data:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key, val in edge_data.items():
                element.set(qn('w:{}'.format(key)), str(val))

def set_khmer_font(run, font_name="Khmer OS Siemreap", size_pt=11, bold=False, italic=False, color=None):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color:
        run.font.color.rgb = color
    # Set complex script font for Word to properly render Khmer
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)

def generate_certificate_docx(student, school_settings, output_path, custom_date=None, custom_lunar=None):
    """
    Generates a Word document matching the exact official template of លិខិតបញ្ជាក់.pdf
    Guaranteed strictly 1 single page on A4.
    """
    doc = Document()
    
    # Page setup: A4, tight margins for strict 1-page fit
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(0.45)
    section.bottom_margin = Inches(0.4)
    section.left_margin = Inches(0.65)
    section.right_margin = Inches(0.65)

    # 1. Header: Kingdom of Cambodia
    p_king = doc.add_paragraph()
    p_king.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_king.paragraph_format.space_before = Pt(0)
    p_king.paragraph_format.space_after = Pt(1)
    p_king.paragraph_format.line_spacing = 1.15
    run1 = p_king.add_run("ព្រះរាជាណាចក្រកម្ពុជា\n")
    set_khmer_font(run1, "Khmer OS Muol Light", 12, bold=False)
    run2 = p_king.add_run("ជាតិ សាសនា ព្រះមហាក្សត្រ\n")
    set_khmer_font(run2, "Khmer OS Muol Light", 11, bold=False)
    run_sym = p_king.add_run("3")
    set_khmer_font(run_sym, "Tacteing", 20)

    # 2. Sub-header Table: Left = Dept & School, Right = 4x6 Photo box (Portrait 1.1" x 1.4", exact match to web view)
    table_head = doc.add_table(rows=1, cols=3)
    table_head.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_head.autofit = False

    cell_left = table_head.rows[0].cells[0]
    cell_mid = table_head.rows[0].cells[1]
    cell_photo = table_head.rows[0].cells[2]

    cell_left.width = Inches(5.45)
    cell_mid.width = Inches(0.67)
    cell_photo.width = Inches(0.85)

    # Left cell content: Font Khmer OS Muol Light Size:10
    p_dept = cell_left.paragraphs[0]
    p_dept.paragraph_format.space_before = Pt(0)
    p_dept.paragraph_format.space_after = Pt(1)
    p_dept.paragraph_format.line_spacing = 1.15
    r_dept = p_dept.add_run(school_settings.get('province_dept', 'មន្ទីរអប់រំ យុវជន និងកីឡា ខេត្តតាកែវ') + "\n")
    set_khmer_font(r_dept, "Khmer OS Muol Light", 10, bold=False)
    r_sch = p_dept.add_run(school_settings.get('school_name', 'វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍'))
    set_khmer_font(r_sch, "Khmer OS Muol Light", 10, bold=False)

    # Middle spacer cell (empty)
    cell_mid.paragraphs[0].text = ""

    # Right photo box: Portrait vertical 4x6 (narrower width, taller height)
    p_photo = cell_photo.paragraphs[0]
    p_photo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_photo.paragraph_format.space_before = Pt(22)
    p_photo.paragraph_format.space_after = Pt(22)
    p_photo.paragraph_format.line_spacing = 1.15
    r_ph1 = p_photo.add_run("រូបថត\n")
    set_khmer_font(r_ph1, "Khmer OS Siemreap", 9)
    r_ph2 = p_photo.add_run("៤x៦")
    set_khmer_font(r_ph2, "Khmer OS Siemreap", 9)

    set_cell_border(cell_photo, 
                    top=dict(sz=6, val='single', color='666666'),
                    bottom=dict(sz=6, val='single', color='666666'),
                    left=dict(sz=6, val='single', color='666666'),
                    right=dict(sz=6, val='single', color='666666'))

    # 3. Main Title
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    r_title = p_title.add_run("លិខិតបញ្ជាក់ការចុះឈ្មោះចូលរៀន")
    set_khmer_font(r_title, "Khmer OS Muol Light", 13.5, bold=False)

    # 4. Director subtitle
    p_director = doc.add_paragraph()
    p_director.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_director.paragraph_format.space_before = Pt(1)
    p_director.paragraph_format.space_after = Pt(2)
    r_dir = p_director.add_run(f"នាយក{school_settings.get('school_name', 'វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍')}\n")
    set_khmer_font(r_dir, "Khmer OS Muol Light", 11, bold=False)
    r_cert = p_director.add_run("សូមបញ្ជាក់ថា៖")
    set_khmer_font(r_cert, "Khmer OS Muol Light", 11, bold=False)

    # 5. Body lines: Font Khmer OS Siemreap size:12
    def add_body_line(label, value="", suffix="", line_space=1.15, space_after=2, val_bold=False):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_space
        r_lbl = p.add_run(label)
        set_khmer_font(r_lbl, "Khmer OS Siemreap", 12, bold=False)
        if value:
            r_val = p.add_run(str(value))
            set_khmer_font(r_val, "Khmer OS Siemreap", 12, bold=val_bold)
        if suffix:
            r_suf = p.add_run(str(suffix))
            set_khmer_font(r_suf, "Khmer OS Siemreap", 12, bold=False)
        return p

    # Gender checkboxes
    gender = student.get('gender', 'ស្រី')
    female_box = "☑" if gender == 'ស្រី' else "☐"
    male_box = "☑" if gender == 'ប្រុស' else "☐"

    p_stu = doc.add_paragraph()
    p_stu.paragraph_format.space_before = Pt(3)
    p_stu.paragraph_format.space_after = Pt(2)
    p_stu.paragraph_format.line_spacing = 1.15
    r_sl = p_stu.add_run("សិស្សឈ្មោះ៖ ")
    set_khmer_font(r_sl, "Khmer OS Siemreap", 12, bold=False)
    r_sn = p_stu.add_run(f"{student.get('full_name', '')}     ")
    set_khmer_font(r_sn, "Khmer OS Siemreap", 12, bold=True)
    r_gl = p_stu.add_run("ភេទ៖  ")
    set_khmer_font(r_gl, "Khmer OS Siemreap", 12, bold=False)
    r_gv = p_stu.add_run(f"{female_box} ស្រី   {male_box} ប្រុស")
    set_khmer_font(r_gv, "Khmer OS Siemreap", 12, bold=False)

    add_body_line("ថ្ងៃខែឆ្នាំកំណើត៖ ", student.get('dob', ''), val_bold=False)
    add_body_line("ទីកន្លែងកំណើត៖ ", student.get('pob', ''), " ។", val_bold=False)

    # Father & Mother
    p_parents1 = doc.add_paragraph()
    p_parents1.paragraph_format.space_before = Pt(0)
    p_parents1.paragraph_format.space_after = Pt(2)
    p_parents1.paragraph_format.line_spacing = 1.15
    r_fl = p_parents1.add_run("ឪពុកឈ្មោះ៖ ")
    set_khmer_font(r_fl, "Khmer OS Siemreap", 12, bold=False)
    r_fn = p_parents1.add_run(f"{student.get('father_name', '')} ")
    set_khmer_font(r_fn, "Khmer OS Siemreap", 12, bold=True)
    r_fjl = p_parents1.add_run("មុខរបរ ")
    set_khmer_font(r_fjl, "Khmer OS Siemreap", 12, bold=False)
    r_fjv = p_parents1.add_run(f"{student.get('father_job', 'កសិករ')}")
    set_khmer_font(r_fjv, "Khmer OS Siemreap", 12, bold=False)

    p_parents2 = doc.add_paragraph()
    p_parents2.paragraph_format.space_before = Pt(0)
    p_parents2.paragraph_format.space_after = Pt(2)
    p_parents2.paragraph_format.line_spacing = 1.15
    r_ml = p_parents2.add_run("ម្តាយឈ្មោះ៖ ")
    set_khmer_font(r_ml, "Khmer OS Siemreap", 12, bold=False)
    r_mn = p_parents2.add_run(f"{student.get('mother_name', '')} ")
    set_khmer_font(r_mn, "Khmer OS Siemreap", 12, bold=True)
    r_mjl = p_parents2.add_run("មុខរបរ ")
    set_khmer_font(r_mjl, "Khmer OS Siemreap", 12, bold=False)
    r_mjv = p_parents2.add_run(f"{student.get('mother_job', 'កសិករ')}")
    set_khmer_font(r_mjv, "Khmer OS Siemreap", 12, bold=False)

    add_body_line("អស័យដ្ឋានបច្ចុប្បន្ន៖ ", student.get('current_address', ''), " ។", val_bold=False)

    # Guardian
    p_guard = doc.add_paragraph()
    p_guard.paragraph_format.space_before = Pt(0)
    p_guard.paragraph_format.space_after = Pt(2)
    p_guard.paragraph_format.line_spacing = 1.15
    r_gl = p_guard.add_run("អ្នកអាណាព្យាបាលឈ្មោះ៖ ")
    set_khmer_font(r_gl, "Khmer OS Siemreap", 12, bold=False)
    r_gn = p_guard.add_run(f"{student.get('guardian_name', '')} ")
    set_khmer_font(r_gn, "Khmer OS Siemreap", 12, bold=True)
    r_gjl = p_guard.add_run("មុខរបរ ")
    set_khmer_font(r_gjl, "Khmer OS Siemreap", 12, bold=False)
    r_gjv = p_guard.add_run(f"{student.get('guardian_job', 'កសិករ')} ")
    set_khmer_font(r_gjv, "Khmer OS Siemreap", 12, bold=False)
    r_grl = p_guard.add_run("ត្រូវជា ")
    set_khmer_font(r_grl, "Khmer OS Siemreap", 12, bold=False)
    r_grv = p_guard.add_run(f"{student.get('guardian_relation', 'ឪពុក')}")
    set_khmer_font(r_grv, "Khmer OS Siemreap", 12, bold=False)

    # Certification clause (Line 1 strictly unbroken, Line 2 on next line)
    p_cert_clause = doc.add_paragraph()
    p_cert_clause.paragraph_format.space_before = Pt(2)
    p_cert_clause.paragraph_format.space_after = Pt(2)
    p_cert_clause.paragraph_format.line_spacing = 1.15
    r_cc1 = p_cert_clause.add_run("ពិតជាសិស្សរៀនថ្នាក់ទី ")
    set_khmer_font(r_cc1, "Khmer OS Siemreap", 11.5, bold=False)
    r_grade = p_cert_clause.add_run(f"{student.get('grade', '')}")
    set_khmer_font(r_grade, "Khmer OS Siemreap", 11.5, bold=True)
    r_cc2 = p_cert_clause.add_run(f" សាលា{school_settings.get('school_name', 'វិទ្យាល័យ ហ៊ុន សែន អង្គប្រីយ៍')} ក្នុងឆ្នាំសិក្សា ")
    set_khmer_font(r_cc2, "Khmer OS Siemreap", 11.5, bold=False)
    r_yr = p_cert_clause.add_run(f"{student.get('academic_year', '')}")
    set_khmer_font(r_yr, "Khmer OS Siemreap", 11.5, bold=True)
    r_cc3 = p_cert_clause.add_run(" បានចុះឈ្មោះ\n")
    set_khmer_font(r_cc3, "Khmer OS Siemreap", 11.5, bold=False)
    r_cc4 = p_cert_clause.add_run("ចូលរៀនពិតប្រាកដមែន ដែលមាន៖")
    set_khmer_font(r_cc4, "Khmer OS Siemreap", 11.5, bold=False)

    # Equity card checkboxes
    card_type = student.get('equity_card_type', 'បណ្ណសមធម៌ក្រីក្រ កម្រិត២')
    lvl1_box = "☑" if "កម្រិត១" in card_type else "☐"
    lvl2_box = "☑" if "កម្រិត២" in card_type else "☐"

    p_card = doc.add_paragraph()
    p_card.paragraph_format.space_before = Pt(2)
    p_card.paragraph_format.space_after = Pt(2)
    p_card.paragraph_format.line_spacing = 1.15
    r_card = p_card.add_run(f"{lvl1_box} ប័ណ្ណសមធម៌ក្រីក្រ កម្រិត១       {lvl2_box} ប័ណ្ណសមធម៌ក្រីក្រ កម្រិត២")
    set_khmer_font(r_card, "Khmer OS Siemreap", 12, bold=False)

    # Closing statement
    p_close = doc.add_paragraph()
    p_close.paragraph_format.space_before = Pt(2)
    p_close.paragraph_format.space_after = Pt(4)
    p_close.paragraph_format.line_spacing = 1.15
    r_cl = p_close.add_run("លិខិតបញ្ជាក់នេះ ផ្តល់ជូនសាមីខ្លួនយកទៅប្រើប្រាស់ផ្លូវការ តាមដែលអាចប្រើបាន។")
    set_khmer_font(r_cl, "Khmer OS Siemreap", 12)

    # 6. Bottom Table: Left = Notes, Right = Date & Signature
    table_foot = doc.add_table(rows=1, cols=2)
    table_foot.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_foot.autofit = False

    cell_f_left = table_foot.rows[0].cells[0]
    cell_f_right = table_foot.rows[0].cells[1]
    cell_f_left.width = Inches(3.7)
    cell_f_right.width = Inches(3.27)

    # Left: Notes (Size: 8pt, aligned ~2 lines below នាយកសាលា)
    p_note = cell_f_left.paragraphs[0]
    p_note.paragraph_format.space_before = Pt(16)
    p_note.paragraph_format.space_after = Pt(1)
    p_note.paragraph_format.line_spacing = 1.15
    r_nl = p_note.add_run("កំណត់សម្គាល់\n")
    set_khmer_font(r_nl, "Khmer OS Siemreap", 8, bold=True)
    r_n1 = p_note.add_run("- លិខិតនេះប្រើបាន លុះត្រាតែ គ្មានការកោសលុប\n")
    set_khmer_font(r_n1, "Khmer OS Siemreap", 8)
    r_n2 = p_note.add_run("- លិខិតនេះ ពុំចេញជូនជាលើកទី២ទេ\n")
    set_khmer_font(r_n2, "Khmer OS Siemreap", 8)
    r_n3 = p_note.add_run("- សាមីជនអាចថតចម្លងតាមច្បាប់ដើម សម្រាប់ប្រើប្រាស់ផ្ទាល់ខ្លួន។")
    set_khmer_font(r_n3, "Khmer OS Siemreap", 8)

    # Right: Date and Signature (Single line dates)
    p_sign = cell_f_right.paragraphs[0]
    p_sign.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sign.paragraph_format.space_before = Pt(0)
    p_sign.paragraph_format.space_after = Pt(1)
    p_sign.paragraph_format.line_spacing = 1.15

    lunar_text = custom_lunar or student.get('cert_lunar_date') or school_settings.get('default_lunar_day', 'ថ្ងៃសុក្រ ៥កើត ខែបឋមាសាឍ ឆ្នាំមមី អដ្ឋស័ក ព.ស.២៥៧០')
    solar_text = custom_date or student.get('cert_solar_date') or f"ធ្វើនៅ {school_settings.get('school_short_name', 'វិ.ហ.ស អង្គប្រីយ៍')} ថ្ងៃទី១៩ ខែមិថុនា ឆ្នាំ២០២៦"

    r_s1 = p_sign.add_run(f"{lunar_text}\n")
    set_khmer_font(r_s1, "Khmer OS Siemreap", 11)
    r_s2 = p_sign.add_run(f"{solar_text}\n")
    set_khmer_font(r_s2, "Khmer OS Siemreap", 11)
    
    # នាយកសាលា: Font Khmer OS Muol Light Size:12, NOT BOLD, right below solar date
    r_s3 = p_sign.add_run(f"\n{school_settings.get('principal_title', 'នាយកសាលា')}\n\n\n")
    set_khmer_font(r_s3, "Khmer OS Muol Light", 12, bold=False)
    if school_settings.get('principal_name'):
        r_s4 = p_sign.add_run(f"{school_settings.get('principal_name')}")
        set_khmer_font(r_s4, "Khmer OS Muol Light", 12, bold=False)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path
