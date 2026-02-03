import pythoncom
from flask import Flask, render_template, request, jsonify, send_file, Response
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import json
import os
from datetime import datetime
import subprocess
import platform
import sys
import webbrowser
import threading
import base64
from io import BytesIO

app = Flask(__name__)

# კონფიგურაცია
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
    TEMPLATE_FOLDER = os.path.join(sys._MEIPASS, 'templates')
    STATIC_FOLDER = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, template_folder=TEMPLATE_FOLDER, static_folder=STATIC_FOLDER)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOCUMENTS_FOLDER = os.path.join(BASE_DIR, 'documents')
TEMPLATES_FOLDER = os.path.join(BASE_DIR, 'saved_templates')
SIGNATURES_FOLDER = os.path.join(BASE_DIR, 'signatures')

for folder in [DOCUMENTS_FOLDER, TEMPLATES_FOLDER, SIGNATURES_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)


# ==================== Helper Functions ====================

def set_cell_shading(cell, color):
    """უჯრის ფონის ფერი"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    cell._tc.get_or_add_tcPr().append(shading)


def decode_base64_image(base64_string):
    """Base64 სურათის დეკოდირება და BytesIO დაბრუნება"""
    if not base64_string:
        return None

    if not isinstance(base64_string, str):
        return None

    if not base64_string.startswith('data:image'):
        return None

    try:
        # Base64 დეკოდირება
        header, data = base64_string.split(',', 1)
        image_data = base64.b64decode(data)
        image_stream = BytesIO(image_data)
        return image_stream
    except Exception as e:
        print(f"Image decode error: {e}")
        return None


# ==================== Document Creation ====================

def create_form_100_document(data):
    """ფორმა №100/ა - ცნობა ჯანმრთელობის მდგომარეობის შესახებ"""
    doc = Document()

    # გვერდის პარამეტრები
    for section in doc.sections:
        section.top_margin = Cm(1)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1)

    style = doc.styles['Normal']
    style.font.name = 'Sylfaen'
    style.font.size = Pt(10)

    # ===== დამტკიცების ინფო =====
    approval_para = doc.add_paragraph()
    approval_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    approval_run = approval_para.add_run(
        'დანართი №2 დამტკიცებულია საქართველოს შრომის\n'
        'ჯანმრთელობისა და სოციალური დაცვის მინისტრის\n'
        '2013 წ 03.12 №01-42/ნ ბრძანებით')
    approval_run.font.size = Pt(8)
    approval_run.font.italic = True

    # ფორმის ტიპი
    form_para = doc.add_paragraph()
    form_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    form_run = form_para.add_run(data.get('form_type', 'სამედიცინო დოკუმენტაცია ფორმა № IV-100/ა'))
    form_run.font.bold = True
    form_run.font.size = Pt(11)

    # სათაური
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run('ცნობა ჯანმრთელობის მდგომარეობის შესახებ')
    title_run.font.bold = True
    title_run.font.size = Pt(12)

    # თარიღი და რეგისტრაცია
    date_para = doc.add_paragraph()
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    date_para.add_run(
        f"გაცემის თარიღი: {data.get('document_date', '')}     ბარათის №: {data.get('registration_number', '')}")

    doc.add_paragraph()

    # ===== 1. გამცემი ორგანიზაცია =====
    t1 = doc.add_table(rows=4, cols=2)
    t1.style = 'Table Grid'
    t1.rows[0].cells[0].merge(t1.rows[0].cells[1])
    t1.rows[0].cells[0].text = "1. გამცემი ორგანიზაცია"
    set_cell_shading(t1.rows[0].cells[0], "D9E2F3")

    t1.rows[1].cells[0].text = "დასახელება:"
    t1.rows[1].cells[1].text = data.get('facility_name', '')
    t1.rows[2].cells[0].text = "საიდენტიფიკაციო კოდი:"
    t1.rows[2].cells[1].text = data.get('identification_code', '')
    t1.rows[3].cells[0].text = "მისამართი:"
    t1.rows[3].cells[1].text = data.get('facility_address', '')

    doc.add_paragraph()

    # ===== 2. მიმღები ორგანიზაცია =====
    t2 = doc.add_table(rows=2, cols=2)
    t2.style = 'Table Grid'
    t2.rows[0].cells[0].merge(t2.rows[0].cells[1])
    t2.rows[0].cells[0].text = "2. მიმღები ორგანიზაცია"
    set_cell_shading(t2.rows[0].cells[0], "D9E2F3")
    t2.rows[1].cells[0].text = "დასახელება:"
    t2.rows[1].cells[1].text = data.get('recipient_name', '')

    doc.add_paragraph()

    # ===== პაციენტის მონაცემები =====
    t3 = doc.add_table(rows=6, cols=2)
    t3.style = 'Table Grid'
    t3.rows[0].cells[0].merge(t3.rows[0].cells[1])
    t3.rows[0].cells[0].text = "პაციენტის მონაცემები"
    set_cell_shading(t3.rows[0].cells[0], "E2EFDA")

    patient_fields = [
        ("3. სახელი, გვარი:", data.get('patient_name', '')),
        ("4. დაბადების თარიღი:", data.get('birth_date', '')),
        ("5. პირადი ნომერი:", data.get('personal_id', '')),
        ("6. მისამართი:", data.get('patient_address', '')),
        ("7. სამუშაო ადგილი:", data.get('occupation', '')),
    ]
    for i, (label, value) in enumerate(patient_fields):
        t3.rows[i + 1].cells[0].text = label
        t3.rows[i + 1].cells[1].text = str(value) if value else ''

    doc.add_paragraph()

    # ===== 8. ჰოსპიტალიზაციის ვადები =====
    t4 = doc.add_table(rows=2, cols=2)
    t4.style = 'Table Grid'
    t4.rows[0].cells[0].merge(t4.rows[0].cells[1])
    t4.rows[0].cells[0].text = "8. ჰოსპიტალიზაციის ვადები"
    set_cell_shading(t4.rows[0].cells[0], "D9E2F3")
    t4.rows[1].cells[0].text = f"მიღება: {data.get('hospitalization_date', '')}"
    t4.rows[1].cells[1].text = f"გაწერა: {data.get('discharge_date', '')}"

    doc.add_paragraph()

    # ===== 9. დიაგნოზი =====
    t5 = doc.add_table(rows=3, cols=2)
    t5.style = 'Table Grid'
    t5.rows[0].cells[0].merge(t5.rows[0].cells[1])
    t5.rows[0].cells[0].text = "9. დიაგნოზი"
    set_cell_shading(t5.rows[0].cells[0], "FCE4D6")
    t5.rows[1].cells[0].text = "ძირითადი:"
    t5.rows[1].cells[1].text = data.get('main_diagnosis', '')
    t5.rows[2].cells[0].text = "ექიმის მიერ დაზუსტება:"
    t5.rows[2].cells[1].text = data.get('case_code', '')

    doc.add_paragraph()

    # ===== 10. გადატანილი დაავადებები =====
    t6 = doc.add_table(rows=2, cols=1)
    t6.style = 'Table Grid'
    t6.rows[0].cells[0].text = "10. გადატანილი დაავადებები"
    set_cell_shading(t6.rows[0].cells[0], "D9E2F3")
    t6.rows[1].cells[0].text = data.get('past_diseases', '')

    doc.add_paragraph()

    # ===== 11. მოკლე ანამნეზი =====
    t7 = doc.add_table(rows=2, cols=1)
    t7.style = 'Table Grid'
    t7.rows[0].cells[0].text = "11. მოკლე ანამნეზი"
    set_cell_shading(t7.rows[0].cells[0], "D9E2F3")
    t7.rows[1].cells[0].text = data.get('anamnesis', '')

    doc.add_paragraph()

    # ===== 12. ჩატარებული გამოკვლევები =====
    t8 = doc.add_table(rows=4, cols=1)
    t8.style = 'Table Grid'
    t8.rows[0].cells[0].text = "12. ჩატარებული გამოკვლევები"
    set_cell_shading(t8.rows[0].cells[0], "D9E2F3")
    t8.rows[1].cells[0].text = f"სისხლის საერთო ანალიზი BL.6: {data.get('blood_analysis', '')}"
    t8.rows[2].cells[0].text = f"გლუკოზის განსაზღვრა სისხლის შრატში BL.12.1: {data.get('biochemistry', '')}"
    t8.rows[3].cells[0].text = f"ინსტრუმენტული კვლევები: {data.get('instrumental', '')}"

    doc.add_paragraph()

    # ===== 13. დაავადების მიმდინარეობა =====
    t9 = doc.add_table(rows=2, cols=1)
    t9.style = 'Table Grid'
    t9.rows[0].cells[0].text = "13. დაავადების მიმდინარეობა"
    set_cell_shading(t9.rows[0].cells[0], "D9E2F3")

    course_text = f"""ტიპი: {data.get('course_type', '')}

მიღებისას: {data.get('admission_status', '')}
ვიტალური მაჩვენებლები: T-{data.get('admission_temp', '')}°C | HR-{data.get('admission_hr', '')} | BP-{data.get('admission_bp', '')} | RR-{data.get('admission_rr', '')} | SpO2-{data.get('admission_spo2', '')}

გაწერისას: {data.get('discharge_status', '')}
ვიტალური მაჩვენებლები: T-{data.get('discharge_temp', '')}°C | HR-{data.get('discharge_hr', '')} | BP-{data.get('discharge_bp', '')} | RR-{data.get('discharge_rr', '')} | SpO2-{data.get('discharge_spo2', '')}"""

    t9.rows[1].cells[0].text = course_text

    doc.add_paragraph()

    # ===== 14. ჩატარებული მკურნალობა =====
    t10 = doc.add_table(rows=2, cols=1)
    t10.style = 'Table Grid'
    t10.rows[0].cells[0].text = "14. ჩატარებული მკურნალობა"
    set_cell_shading(t10.rows[0].cells[0], "D9E2F3")
    t10.rows[1].cells[
        0].text = f"მედიკამენტები:\n{data.get('medications', '')}\n\nკოდი: {data.get('treatment_code', '')}"

    doc.add_paragraph()

    # ===== 15-17. გამოსავალი =====
    t11 = doc.add_table(rows=4, cols=2)
    t11.style = 'Table Grid'
    t11.rows[0].cells[0].merge(t11.rows[0].cells[1])
    t11.rows[0].cells[0].text = "გამოსავალი"
    set_cell_shading(t11.rows[0].cells[0], "E2EFDA")
    t11.rows[1].cells[0].text = "15. სტაციონარში გადაყვანა:"
    t11.rows[1].cells[1].text = data.get('transfer_to_hospital', '-') or '-'
    t11.rows[2].cells[0].text = "16. გაწერის მდგომარეობა:"
    t11.rows[2].cells[1].text = data.get('discharge_condition', '')
    t11.rows[3].cells[0].text = "17. რეკომენდაციები:"
    t11.rows[3].cells[1].text = data.get('recommendations', '')

    doc.add_paragraph()

    # ===== 18-20. ხელმოწერები =====
    t12 = doc.add_table(rows=4, cols=2)
    t12.style = 'Table Grid'
    t12.rows[0].cells[0].merge(t12.rows[0].cells[1])
    t12.rows[0].cells[0].text = "ხელმოწერები"
    set_cell_shading(t12.rows[0].cells[0], "D9E2F3")

    t12.rows[1].cells[0].text = "18. მკურნალი ექიმი:"
    t12.rows[1].cells[1].text = data.get('attending_doctor', '')

    t12.rows[2].cells[0].text = "19. დაწესებულების ხელმძღვანელი:"
    t12.rows[2].cells[1].text = data.get('facility_head', '')

    t12.rows[3].cells[0].text = "20. ცნობის გაცემის თარიღი:"
    t12.rows[3].cells[1].text = data.get('issue_date', '')

    doc.add_paragraph()

    # ===== ელექტრონული ხელმოწერები და ბეჭედი =====
    sig_table = doc.add_table(rows=2, cols=3)
    sig_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # ლეიბლები
    sig_table.rows[0].cells[0].text = "ექიმის ხელმოწერა"
    sig_table.rows[0].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    sig_table.rows[0].cells[1].text = "ბეჭედი"
    sig_table.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    sig_table.rows[0].cells[2].text = "ხელმძღვანელის ხელმოწერა"
    sig_table.rows[0].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ექიმის ხელმოწერა
    doctor_sig_data = data.get('doctor_signature_image', '')
    doctor_img = decode_base64_image(doctor_sig_data)
    if doctor_img:
        try:
            run = sig_table.rows[1].cells[0].paragraphs[0].add_run()
            run.add_picture(doctor_img, width=Inches(1.2))
            sig_table.rows[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            print(f"Doctor signature error: {e}")
            sig_table.rows[1].cells[0].text = "________________"
            sig_table.rows[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        sig_table.rows[1].cells[0].text = "________________"
        sig_table.rows[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ბეჭედი
    stamp_data = data.get('stamp_image', '')
    stamp_img = decode_base64_image(stamp_data)
    if stamp_img:
        try:
            run = sig_table.rows[1].cells[1].paragraphs[0].add_run()
            run.add_picture(stamp_img, width=Inches(1.2))
            sig_table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            print(f"Stamp error: {e}")
            sig_table.rows[1].cells[1].text = "ბ.ა."
            sig_table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        sig_table.rows[1].cells[1].text = "ბ.ა."
        sig_table.rows[1].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # ხელმძღვანელის ხელმოწერა
    head_sig_data = data.get('head_signature_image', '')
    head_img = decode_base64_image(head_sig_data)
    if head_img:
        try:
            run = sig_table.rows[1].cells[2].paragraphs[0].add_run()
            run.add_picture(head_img, width=Inches(1.2))
            sig_table.rows[1].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except Exception as e:
            print(f"Head signature error: {e}")
            sig_table.rows[1].cells[2].text = "________________"
            sig_table.rows[1].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        sig_table.rows[1].cells[2].text = "________________"
        sig_table.rows[1].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    return doc


def create_medical_record_document(data):
    """სამედიცინო ჩანაწერი - კურსუსი"""
    doc = Document()

    for section in doc.sections:
        section.top_margin = Cm(1)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1)

    style = doc.styles['Normal']
    style.font.name = 'Sylfaen'
    style.font.size = Pt(10)

    # სათაური
    h1 = doc.add_paragraph()
    h1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r1 = h1.add_run(data.get('facility_name', 'პრემიუმ მედ გრუპი'))
    r1.font.bold = True
    r1.font.size = Pt(14)

    h2 = doc.add_paragraph()
    h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h2.add_run(data.get('department', 'გადაუდებელი მედიცინა'))

    doc.add_paragraph()

    # პაციენტის მონაცემები
    t1 = doc.add_table(rows=4, cols=2)
    t1.style = 'Table Grid'
    t1.rows[0].cells[0].merge(t1.rows[0].cells[1])
    t1.rows[0].cells[0].text = "პაციენტის მონაცემები"
    set_cell_shading(t1.rows[0].cells[0], "E2EFDA")
    t1.rows[1].cells[0].text = "ბარათის №:"
    t1.rows[1].cells[1].text = data.get('card_number', '-') or '-'
    t1.rows[2].cells[0].text = "სახელი, გვარი:"
    t1.rows[2].cells[1].text = data.get('patient_name', '-') or '-'
    t1.rows[3].cells[0].text = "მიღების სტატუსი:"
    t1.rows[3].cells[1].text = data.get('admission_status', 'თვითდინებით')

    doc.add_paragraph()

    # დიაგნოზი
    t2 = doc.add_table(rows=2, cols=2)
    t2.style = 'Table Grid'
    t2.rows[0].cells[0].merge(t2.rows[0].cells[1])
    t2.rows[0].cells[0].text = "დიაგნოზი (ICD-10)"
    set_cell_shading(t2.rows[0].cells[0], "FCE4D6")
    t2.rows[1].cells[0].text = f"ZYZA10 ამბულატორია (გადაუდებელი): {data.get('icd_code', '')}"
    t2.rows[1].cells[1].text = data.get('diagnosis_description', '')

    doc.add_paragraph()

    # ჩივილები
    t3 = doc.add_table(rows=2, cols=1)
    t3.style = 'Table Grid'
    t3.rows[0].cells[0].text = "ჩივილები"
    set_cell_shading(t3.rows[0].cells[0], "D9E2F3")
    t3.rows[1].cells[0].text = data.get('complaints', '')

    doc.add_paragraph()

    # ანამნეზი
    t4 = doc.add_table(rows=2, cols=1)
    t4.style = 'Table Grid'
    t4.rows[0].cells[0].text = "ანამნეზი"
    set_cell_shading(t4.rows[0].cells[0], "D9E2F3")
    t4.rows[1].cells[0].text = data.get('anamnesis', '')

    doc.add_paragraph()

    # ალერგიები
    al = doc.add_paragraph()
    alr = al.add_run("ალერგიები: ")
    alr.font.bold = True
    al.add_run(data.get('allergies', 'არა'))

    doc.add_paragraph()

    # ობიექტური სტატუსი
    obj_h = doc.add_paragraph()
    obj_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    obj_r = obj_h.add_run("ობიექტური სტატუსი")
    obj_r.font.bold = True
    obj_r.font.size = Pt(12)

    # ვიტალები
    vt = doc.add_table(rows=2, cols=5)
    vt.style = 'Table Grid'
    headers = ["T°C", "BP", "HR", "RR", "SpO₂"]
    values = [
        data.get('temperature', ''),
        data.get('blood_pressure', ''),
        data.get('heart_rate', ''),
        data.get('respiratory_rate', ''),
        data.get('spo2', '')
    ]
    for i, h in enumerate(headers):
        vt.rows[0].cells[i].text = h
        set_cell_shading(vt.rows[0].cells[i], "D0D0D0")
        vt.rows[0].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        vt.rows[1].cells[i].text = str(values[i]) if values[i] else ''
        vt.rows[1].cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph()

    # ორგანო-სისტემები
    sys_t = doc.add_table(rows=9, cols=2)
    sys_t.style = 'Table Grid'
    systems = [
        ("ზოგადი მდგომარეობა:", data.get('general_condition', '')),
        ("კანი:", data.get('skin', '')),
        ("პერიფერიული შეშუპება:", data.get('edema', '')),
        ("გულ-სისხლძარღვთა:", data.get('cardiovascular', '')),
        ("სასუნთქი:", data.get('respiratory', '')),
        ("საჭმლის მომნელებელი:", data.get('digestive', '')),
        ("შარდგამომყოფი:", data.get('urinary', '')),
        ("ნერვული სისტემა:", data.get('neurological', '')),
        ("საყრდენ-მამოძრავებელი:", data.get('musculoskeletal', '')),
    ]
    for i, (label, value) in enumerate(systems):
        sys_t.rows[i].cells[0].text = label
        sys_t.rows[i].cells[1].text = str(value) if value else ''
        set_cell_shading(sys_t.rows[i].cells[0], "F2F2F2")

    doc.add_paragraph()

    # წინასწარი დიაგნოზი
    pd = doc.add_paragraph()
    pdr = pd.add_run("წინასწარი დიაგნოზი: ")
    pdr.font.bold = True
    pd.add_run(data.get('preliminary_diagnosis', ''))

    dr = doc.add_paragraph()
    drr = dr.add_run("მკურნალი ექიმი: ")
    drr.font.bold = True
    dr.add_run(data.get('doctor', ''))

    # გვერდი 2
    doc.add_page_break()

    pg_h = doc.add_paragraph()
    pg_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pg_r = pg_h.add_run("პაციენტის მიმდინარეობის ფურცელი (დღიური)")
    pg_r.font.bold = True
    pg_r.font.size = Pt(14)

    doc.add_paragraph()

    # პირველადი შეფასება
    init_t = doc.add_table(rows=2, cols=1)
    init_t.style = 'Table Grid'
    initial_date = data.get('initial_date', '')
    header_text = "პირველადი შეფასება / მიღება"
    if initial_date:
        header_text += f"  (თარიღი: {initial_date})"
    init_t.rows[0].cells[0].text = header_text
    set_cell_shading(init_t.rows[0].cells[0], "D9E2F3")
    init_t.rows[1].cells[0].text = data.get('initial_narrative', '')

    doc.add_paragraph()

    id_p = doc.add_paragraph()
    id_r = id_p.add_run("წინასწარი დიაგნოზი: ")
    id_r.font.bold = True
    id_p.add_run(data.get('initial_diagnosis', ''))

    doc.add_paragraph()

    # დანიშნულებები
    ord_h = doc.add_paragraph()
    ord_r = ord_h.add_run("დანიშნულებები:")
    ord_r.font.bold = True
    ord_r.font.size = Pt(11)

    inv_h = doc.add_paragraph()
    inv_r = inv_h.add_run("გამოკვლევები:")
    inv_r.font.italic = True

    investigations = data.get('investigations', '')
    if investigations:
        for inv in investigations.split('\n'):
            if inv.strip():
                b = doc.add_paragraph(style='List Bullet')
                b.add_run(inv.strip())

    med_h = doc.add_paragraph()
    med_r = med_h.add_run("მედიკამენტები:")
    med_r.font.italic = True

    medications = data.get('medications', '')
    if medications:
        for med in medications.split('\n'):
            if med.strip():
                b = doc.add_paragraph(style='List Bullet')
                b.add_run(med.strip())

    doc.add_paragraph()

    # ექიმის ხელმოწერა
    sig1 = doc.add_paragraph()
    sig1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig1.add_run(f"ექიმი: {data.get('doctor_signature', '')}")

    # ექიმის ხელმოწერის სურათი
    doctor_sig_data = data.get('doctor_signature_image', '')
    doctor_img = decode_base64_image(doctor_sig_data)
    if doctor_img:
        try:
            sig_para = doc.add_paragraph()
            sig_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = sig_para.add_run()
            run.add_picture(doctor_img, width=Inches(1.0))
        except Exception as e:
            print(f"Doctor signature error: {e}")

    doc.add_paragraph()
    doc.add_paragraph()

    # გადაფასება / გაწერა
    dis_t = doc.add_table(rows=2, cols=1)
    dis_t.style = 'Table Grid'
    discharge_note_date = data.get('discharge_note_date', '')
    dis_header = "გადაფასება / გაწერა"
    if discharge_note_date:
        dis_header += f"  (თარიღი: {discharge_note_date})"
    dis_t.rows[0].cells[0].text = dis_header
    set_cell_shading(dis_t.rows[0].cells[0], "E2EFDA")
    dis_t.rows[1].cells[0].text = data.get('discharge_narrative', '')

    doc.add_paragraph()

    sig2 = doc.add_paragraph()
    sig2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig2.add_run(f"ექიმი: {data.get('discharge_doctor', '')}")

    return doc


# ==================== PDF Conversion ====================

def find_libreoffice():
    """LibreOffice-ის პოვნა"""
    if platform.system() == 'Windows':
        paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            os.path.expandvars(r"%PROGRAMFILES%\LibreOffice\program\soffice.exe"),
            os.path.expandvars(r"%PROGRAMFILES(X86)%\LibreOffice\program\soffice.exe"),
        ]
        for path in paths:
            if os.path.exists(path):
                return path
    else:
        for cmd in ['libreoffice', 'soffice']:
            try:
                subprocess.run([cmd, '--version'], capture_output=True, check=True)
                return cmd
            except:
                pass
    return None


def convert_to_pdf(docx_path, output_folder):
    """DOCX -> PDF (ჯერ docx2pdf + Word, შემდეგ LibreOffice როგორც fallback)"""
    system = platform.system()
    pdf_path = docx_path.replace('.docx', '.pdf')

    # 1) Windows + Word (docx2pdf)
    if system == 'Windows':
        try:
            # !!! მთავარი ხაზი: COM ინიციალიზაცია მიმდინარე თრედში !!!
            pythoncom.CoInitialize()

            from docx2pdf import convert
            convert(docx_path, pdf_path)

            # თუ PDF შეიქმნა, ვაბრუნებთ მის გზას
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception as e:
            print("docx2pdf error:", e)
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    # 2) LibreOffice (თუ დაყენებულია) – fallback
    lo = find_libreoffice()
    if lo:
        try:
            subprocess.run([
                lo, '--headless', '--convert-to', 'pdf',
                '--outdir', output_folder, docx_path
            ], capture_output=True, timeout=60)
            if os.path.exists(pdf_path):
                return pdf_path
        except Exception as e:
            print("LibreOffice error:", e)

    # თუ ვერცერთი იმუშავა, ვაბრუნებთ None-ს → frontend გადადის DOCX ჩამოტვირთვაზე
    return None


# ==================== Flask Routes ====================

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/save-document', methods=['POST'])
def save_document():
    try:
        data = request.json
        doc_type = data.get('document_type', 'form_100')
        filename = data.get('filename', f'document_{datetime.now().strftime("%Y%m%d_%H%M%S")}')

        safe_chars = set('აბგდევზთიკლმნოპჟრსტუფქღყშჩცძწჭხჯჰ_- ')
        filename = ''.join(c for c in filename if c.isalnum() or c in safe_chars)

        if doc_type == 'form_100':
            doc = create_form_100_document(data)
        else:
            doc = create_medical_record_document(data)

        filepath = os.path.join(DOCUMENTS_FOLDER, f'{filename}.docx')
        doc.save(filepath)

        return jsonify({
            'success': True,
            'message': 'დოკუმენტი წარმატებით შეინახა',
            'filename': f'{filename}.docx'
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/print-document', methods=['POST'])
def print_document():
    try:
        data = request.json
        doc_type = data.get('document_type', 'form_100')
        filename = f'print_{datetime.now().strftime("%Y%m%d_%H%M%S")}'

        if doc_type == 'form_100':
            doc = create_form_100_document(data)
        else:
            doc = create_medical_record_document(data)

        docx_path = os.path.join(DOCUMENTS_FOLDER, f'{filename}.docx')
        doc.save(docx_path)

        pdf_path = convert_to_pdf(docx_path, DOCUMENTS_FOLDER)

        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(docx_path)
            except:
                pass

            return jsonify({
                'success': True,
                'message': 'PDF წარმატებით შეიქმნა',
                'filename': os.path.basename(pdf_path),
                'is_pdf': True
            })
        else:
            return jsonify({
                'success': True,
                'message': 'PDF კონვერტაცია ვერ მოხერხდა',
                'filename': f'{filename}.docx',
                'is_pdf': False
            })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/print-page/<filename>')
def print_page(filename):
    filepath = os.path.join(DOCUMENTS_FOLDER, filename)
    if not os.path.exists(filepath):
        return "ფაილი ვერ მოიძებნა", 404
    return render_template('print.html', filename=filename)


@app.route('/api/view-pdf/<filename>')
def view_pdf(filename):
    try:
        filepath = os.path.join(DOCUMENTS_FOLDER, filename)
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': 'ფაილი ვერ მოიძებნა'}), 404

        if filename.endswith('.pdf'):
            mimetype = 'application/pdf'
        else:
            mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

        with open(filepath, 'rb') as f:
            file_data = f.read()

        response = Response(file_data, mimetype=mimetype)
        response.headers['Content-Disposition'] = f'inline; filename="{filename}"'
        return response
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/download/<filename>')
def download_file(filename):
    try:
        filepath = os.path.join(DOCUMENTS_FOLDER, filename)
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': 'ფაილი ვერ მოიძებნა'}), 404
        return send_file(filepath, as_attachment=True)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 404


# ==================== Signature Upload ====================

@app.route('/api/upload-signature', methods=['POST'])
def upload_signature():
    """ხელმოწერის/ბეჭდის ატვირთვა"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': 'ფაილი არ არის'}), 400

        file = request.files['file']
        sig_type = request.form.get('type', 'doctor')

        if file.filename == '':
            return jsonify({'success': False, 'error': 'ფაილი არჩეული არ არის'}), 400

        ext = file.filename.rsplit('.', 1)[-1].lower()
        if ext not in ['png', 'jpg', 'jpeg', 'gif']:
            return jsonify({'success': False, 'error': 'მხოლოდ სურათები (PNG, JPG)'}), 400

        filename = f'{sig_type}_signature.{ext}'
        filepath = os.path.join(SIGNATURES_FOLDER, filename)
        file.save(filepath)

        with open(filepath, 'rb') as f:
            base64_data = base64.b64encode(f.read()).decode('utf-8')

        return jsonify({
            'success': True,
            'filename': filename,
            'base64': f'data:image/{ext};base64,{base64_data}'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/get-signatures')
def get_signatures():
    """შენახული ხელმოწერების მიღება"""
    signatures = {}

    for sig_type in ['doctor', 'head', 'stamp']:
        for ext in ['png', 'jpg', 'jpeg', 'gif']:
            filepath = os.path.join(SIGNATURES_FOLDER, f'{sig_type}_signature.{ext}')
            if os.path.exists(filepath):
                with open(filepath, 'rb') as f:
                    base64_data = base64.b64encode(f.read()).decode('utf-8')
                signatures[sig_type] = f'data:image/{ext};base64,{base64_data}'
                break

    return jsonify({'success': True, 'signatures': signatures})


@app.route('/api/search-patients', methods=['GET'])
def search_patients():
    query = request.args.get('q', '').lower()
    if not query:
        return jsonify({'success': True, 'results': []})

    results = []

    # 1. შენახული დოკუმენტების ძებნა (documents/)
    if os.path.exists(DOCUMENTS_FOLDER):
        for filename in os.listdir(DOCUMENTS_FOLDER):
            if filename.endswith('.docx') and query in filename.lower():
                results.append({
                    'type': 'document',
                    'name': filename,
                    'path': f'/api/download/{filename}',
                    'date': datetime.fromtimestamp(os.path.getmtime(os.path.join(DOCUMENTS_FOLDER, filename))).strftime(
                        '%Y-%m-%d %H:%M')
                })

    # 2. შაბლონების ძებნა (saved_templates/) - JSON შიგთავსში
    if os.path.exists(TEMPLATES_FOLDER):
        for filename in os.listdir(TEMPLATES_FOLDER):
            if filename.endswith('.json'):
                try:
                    path = os.path.join(TEMPLATES_FOLDER, filename)
                    with open(path, 'r', encoding='utf-8') as f:
                        data = json.load(f)

                    # ვეძებთ პაციენტის სახელს ან პირად ნომერს
                    patient_name = data.get('patient_name', '').lower()
                    personal_id = data.get('personal_id', '')
                    template_name = data.get('template_name', '').lower()

                    if query in patient_name or query in personal_id or query in template_name:
                        results.append({
                            'type': 'template',
                            'name': data.get('template_name', filename),
                            'id': filename.replace('.json', ''),
                            'patient': data.get('patient_name', '-'),
                            'date': data.get('created', '')[:16].replace('T', ' ')
                        })
                except:
                    continue

    return jsonify({'success': True, 'results': results})

# ==================== Templates API ====================

@app.route('/api/templates', methods=['GET'])
def get_templates():
    try:
        templates = []
        if os.path.exists(TEMPLATES_FOLDER):
            for filename in os.listdir(TEMPLATES_FOLDER):
                if filename.endswith('.json'):
                    filepath = os.path.join(TEMPLATES_FOLDER, filename)
                    with open(filepath, 'r', encoding='utf-8') as f:
                        template_data = json.load(f)
                        templates.append({
                            'id': filename.replace('.json', ''),
                            'name': template_data.get('template_name', filename),
                            'type': template_data.get('document_type', 'unknown'),
                            'created': template_data.get('created', ''),
                            'data': template_data
                        })
        return jsonify({'success': True, 'templates': templates})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/templates', methods=['POST'])
def save_template():
    try:
        data = request.json
        template_name = data.get('template_name', f'template_{datetime.now().strftime("%Y%m%d_%H%M%S")}')
        template_id = template_name.replace(' ', '_').lower()
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f'{template_id}_{timestamp}.json'

        data['created'] = datetime.now().isoformat()
        data['template_name'] = template_name

        filepath = os.path.join(TEMPLATES_FOLDER, filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        return jsonify({'success': True, 'message': 'შაბლონი წარმატებით შეინახა'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/templates/<template_id>', methods=['DELETE'])
def delete_template(template_id):
    try:
        for filename in os.listdir(TEMPLATES_FOLDER):
            if filename.startswith(template_id) and filename.endswith('.json'):
                os.remove(os.path.join(TEMPLATES_FOLDER, filename))
                return jsonify({'success': True, 'message': 'შაბლონი წაშლილია'})
        return jsonify({'success': False, 'error': 'შაბლონი ვერ მოიძებნა'}), 404
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ==================== Startup ====================

def open_browser():
    webbrowser.open('http://127.0.0.1:5000')


if __name__ == '__main__':
    print("=" * 50)
    print("🏥 სამედიცინო დოკუმენტაცია")
    print("=" * 50)

    lo = find_libreoffice()
    if lo:
        print(f"✅ LibreOffice: {lo}")
    else:
        print("⚠️  LibreOffice ვერ მოიძებნა")

    print(f"\n📁 დოკუმენტები: {DOCUMENTS_FOLDER}")
    print(f"📁 შაბლონები: {TEMPLATES_FOLDER}")
    print(f"📁 ხელმოწერები: {SIGNATURES_FOLDER}")
    print("\n🌐 მისამართი: http://127.0.0.1:5000")
    print("=" * 50)

    threading.Timer(1.5, open_browser).start()
    app.run(debug=False, host='127.0.0.1', port=5000, threaded=True)