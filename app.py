import os
import shutil
import zipfile
import random
import uuid
import pandas as pd
from flask import Flask, render_template, request, send_file, url_for, flash, redirect
from jinja2 import Environment, FileSystemLoader
from playwright.sync_api import sync_playwright

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash_messages'

# Directories setup
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')

for d in [INPUT_DIR, OUTPUT_DIR, TEMPLATES_DIR]:
    os.makedirs(d, exist_ok=True)

# Helper function: Excel parser
def process_excel(file_path):
    try:
        # Read without headers first to find the actual header row
        df_student_raw = pd.read_excel(file_path, sheet_name=0, header=None)
        df_class_raw = pd.read_excel(file_path, sheet_name=1, header=None)
    except Exception as e:
        return None, f"Excel parsing error: {e}. Make sure the excel file has at least 2 sheets."

    def find_header_and_build_df(df_raw, keyword='분반'):
        header_idx = 0
        for idx, row in df_raw.iterrows():
            # Check if any cell in this row contains the keyword (as a string)
            if any(keyword in str(cell).replace(' ', '') for cell in row.values if pd.notna(cell)):
                header_idx = idx
                break
        
        # Set the found row as header
        new_header = df_raw.iloc[header_idx]
        df_clean = df_raw[header_idx+1:].copy()
        df_clean.columns = new_header
        # Reset index
        df_clean.reset_index(drop=True, inplace=True)
        return df_clean

    df_student = find_header_and_build_df(df_student_raw, '분반')
    df_class = find_header_and_build_df(df_class_raw, '분반')

    # Clean columns mapping just in case, removing newlines and spaces
    df_class.columns = df_class.columns.astype(str).str.replace('\n', '').str.replace(' ', '').str.strip()
    df_student.columns = df_student.columns.astype(str).str.replace('\n', '').str.replace(' ', '').str.strip()
    
    print("Student Columns:", df_student.columns.tolist())
    print("Class Columns:", df_class.columns.tolist())
    
    # Safely find the class name column in both sheets by searching for '분반'
    class_col_student = next((col for col in df_student.columns if '분반' in str(col)), None)
    class_col_class = next((col for col in df_class.columns if '분반' in str(col)), None)
    
    if not class_col_student or not class_col_class:
        return None, f"Merge Error: '분반' 이 포함된 열(Column)을 찾을 수 없습니다. (학생시트: {class_col_student}, 분반시트: {class_col_class})"
        
    # Rename them to a standard key to guarantee merge
    df_student.rename(columns={class_col_student: 'merge_class_key'}, inplace=True)
    df_class.rename(columns={class_col_class: 'merge_class_key'}, inplace=True)
    
    # Merge Student and Class data on the standard key
    try:
        df_merged = pd.merge(df_student, df_class, on='merge_class_key', how='left')
    except Exception as e:
        return None, f"Merge Error: {e}"
    
    reports_data = []
    
    for _, row in df_merged.iterrows():
        # Exception Logic Handling
        test_held_raw = row.get('테스트 실시 여부', False)
        if pd.isna(test_held_raw) or str(test_held_raw).strip() == '':
            test_held = False
        else:
            test_held = str(test_held_raw).strip().upper() not in ['FALSE', 'N', 'X', '미실시', '0']

        test_score = row.get('테스트 점수')
        
        # Test Status
        if not test_held:
            display_test_score = "미실시"
            display_test_percent = 0
            test_status = "미실시"
        elif pd.isna(test_score) or str(test_score).strip() == '':
            display_test_score = "미응시"
            display_test_percent = 0
            test_status = "미응시"
        else:
            try:
                display_test_score = int(float(str(test_score).replace('점', '').strip()))
                display_test_percent = display_test_score
            except ValueError:
                display_test_score = str(test_score)
                display_test_percent = 0
            test_status = "응시"

        # Absence & Notes
        absence_reason = row.get('결석 사유')
        if pd.isna(absence_reason) or str(absence_reason).strip() == '':
            absence_reason = "-"

        special_note = row.get('특이사항 (밀린 과제, 테스트 미응시 사유 등)')
        if pd.isna(special_note) or str(special_note).strip() == '':
            special_note = "당일 특이사항 없습니다."

        notice = row.get('공지사항')
        if pd.isna(notice) or str(notice).strip() == '':
            notice = "별도 공지사항 없습니다."

        # Date formatting safely
        date_raw = row.get('날짜', 'Unknown Date')
        if hasattr(date_raw, 'strftime'):
            date_str = date_raw.strftime('%Y-%m-%d')
        else:
            date_str = str(date_raw)

        # Homework safely convert to int if possible, else keep as text
        hw_raw = row.get('과제 이행도')
        if pd.isna(hw_raw) or str(hw_raw).strip() == '':
            hw_completion = 0
            hw_text = "0"
            hw_is_percent = True
            hw_status = "미완료"
        else:
            hw_str = str(hw_raw).strip()
            try:
                hw_completion = int(float(hw_str.replace('%', '')))
                hw_text = str(hw_completion)
                hw_is_percent = True
                hw_status = "완료" if hw_completion >= 100 else "부분 완료" if hw_completion > 0 else "미완료"
            except ValueError:
                # It's a text string like "첫 시간", "미확인"
                hw_completion = 0
                hw_text = hw_str
                hw_is_percent = False
                hw_status = hw_str

        class_avg_raw = row.get('반평균')
        if pd.isna(class_avg_raw) or str(class_avg_raw).strip() == '':
            class_avg = "-"
        else:
            try:
                class_avg = int(float(str(class_avg_raw).replace('점', '').strip()))
            except ValueError:
                class_avg = str(class_avg_raw)

        # Data map
        # Find mapping for dynamic columns
        def get_val(row, keyword, default=''):
            col = next((c for c in row.index if keyword in str(c)), None)
            return row[col] if col else default

        # Attendance parsing
        raw_attendance = get_val(row, '출석', '')
        if pd.isna(raw_attendance) or str(raw_attendance).strip() == '' or str(raw_attendance).lower() == 'nan':
            attendance_val = "출석"
        else:
            attendance_val = str(raw_attendance).strip()

        data = {
            "date": date_str,
            "student_name": str(get_val(row, '학생', 'Unknown')),
            "school": str(get_val(row, '학교', '')),
            "grade": str(get_val(row, '학년', '')).replace('학년', '').strip(),
            "class_name": str(row.get('merge_class_key', 'Unknown')),
            "attendance": attendance_val,
            "absence_reason": str(absence_reason),
            "test_status": test_status,
            "homework_status": hw_status,
            "display_test_score": display_test_score,
            "test_score_percent": display_test_percent, # for SVG gauge offset
            "class_average": class_avg,
            "homework_completion": hw_completion, # for SVG gauge offset
            "homework_text": hw_text,
            "homework_is_percent": hw_is_percent,
            "lesson_content": str(row.get('학습 내용', '')),
            "next_homework": str(row.get('다음 시간 과제', '')),
            "special_notes": str(special_note),
            "announcements": str(notice),
        }
        reports_data.append(data)
        
    return reports_data, None

# HTML rendering to PNG
def generate_images(reports_data, job_id):
    env = Environment(loader=FileSystemLoader(TEMPLATES_DIR))
    template = env.get_template('report.html')
    
    job_output_dir = os.path.join(OUTPUT_DIR, job_id)
    os.makedirs(job_output_dir, exist_ok=True)
    
    generated_files = []
    
    # Read the CSS file so we can inject it inline and avoid path resolution issues
    css_path = os.path.join(BASE_DIR, 'static', 'style.css')
    with open(css_path, 'r', encoding='utf-8') as f:
        inline_css = f.read()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        # Using device_scale_factor=2 for crisp high-res UI
        page = browser.new_page(
            viewport={"width": 1080, "height": 1560},
            device_scale_factor=2
        )
        
        for data in reports_data:
            data['inline_css'] = inline_css
            html_content = template.render(**data)
            temp_html_path = os.path.join(job_output_dir, f"temp_{uuid.uuid4().hex[:6]}.html")
            
            with open(temp_html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
            file_url = f"file:///{temp_html_path.replace(os.sep, '/')}"
            page.goto(file_url, wait_until='networkidle')
            
            # Directory structure: output/JobId/Date/ClassName/StudentName.png
            date_dir = os.path.join(job_output_dir, data['date'])
            class_dir = os.path.join(date_dir, data['class_name'])
            os.makedirs(class_dir, exist_ok=True)
            
            png_filename = f"{data['student_name']}.png"
            png_path = os.path.join(class_dir, png_filename)
            
            # Full_page=True inside snapshot takes dynamic height into account
            page.screenshot(path=png_path, full_page=True)
            generated_files.append((png_path, data))
            
            # clean up temp html
            try:
                os.remove(temp_html_path)
            except:
                pass
            
        browser.close()
        
    return job_output_dir, generated_files

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash("파일을 첨부해주세요.", "error")
            return redirect(request.url)
        
        file = request.files['file']
        if file.filename == '':
            flash("선택된 파일이 없습니다.", "error")
            return redirect(request.url)
            
        if not file.filename.endswith(('.xls', '.xlsx')):
            flash("엑셀 파일(.xlsx)만 업로드 가능합니다.", "error")
            return redirect(request.url)
            
        # Secure saving
        job_id = uuid.uuid4().hex[:8]
        file_path = os.path.join(INPUT_DIR, f"{job_id}_{file.filename}")
        file.save(file_path)
        
        reports_data, err = process_excel(file_path)
        if err:
            flash(err, "error")
            return redirect(request.url)
            
        # Create Images
        job_dir, files = generate_images(reports_data, job_id)
        
        # Create ZIP
        zip_filename = f"reports_{job_id}.zip"
        zip_path = os.path.join(OUTPUT_DIR, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # Walk directory to maintain Date/Class folder tree in ZIP
            for root, _, arc_files in os.walk(job_dir):
                for f in arc_files:
                    abs_file = os.path.join(root, f)
                    arc_name = os.path.relpath(abs_file, start=job_dir)
                    zipf.write(abs_file, arcname=arc_name)
        
        # Select random preview
        preview_image = None
        if files:
            # Copy random image to static folder so browser can serve it easily mapping to static/preview/
            preview_dir = os.path.join(BASE_DIR, 'static', 'previews')
            os.makedirs(preview_dir, exist_ok=True)
            
            random_file, rand_data = random.choice(files)
            preview_filename = f"preview_{job_id}.png"
            preview_dest = os.path.join(preview_dir, preview_filename)
            shutil.copy(random_file, preview_dest)
            preview_image = url_for('static', filename=f'previews/{preview_filename}')
            
        return render_template('result.html', zip_url=f"/download/{zip_filename}", preview_url=preview_image, total=len(files))
        
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return "File not found", 404

if __name__ == '__main__':
    app.run(debug=True, port=5000)
