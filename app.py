import os
import shutil
import zipfile
import random
import time
import threading
import uuid
import datetime
import gc
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from flask import Flask, render_template, request, send_file, url_for, flash, redirect
from playwright.sync_api import sync_playwright
from playwright.sync_api import sync_playwright

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash_messages'

# Global store for job progress
progress_store = {}

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
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        if len(sheet_names) < 2:
            return None, "Excel parsing error: Make sure the excel file has at least 2 sheets."
            
        # Determine actual sheets by looking for keywords, else default to index 0 and 1
        student_sheet = next((s for s in sheet_names if '학생' in s), sheet_names[0])
        class_sheet = next((s for s in sheet_names if '분반' in s), sheet_names[1])
        
        print(f"Loading Student Sheet: {student_sheet}, Class Sheet: {class_sheet}")
        
        # Read without headers first to find the actual header row
        df_student_raw = pd.read_excel(xls, sheet_name=student_sheet, header=None)
        df_class_raw = pd.read_excel(xls, sheet_name=class_sheet, header=None)
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
        
        # Remove any columns that are NaN or unnamed to prevent merge issues
        df_clean = df_clean.loc[:, df_clean.columns.notna()]
        df_clean = df_clean.loc[:, ~df_clean.columns.astype(str).str.contains('^Unnamed:')]
        
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
    
    # Robust value extractor that tries multiple keywords and cleans whitespace
    def get_val(row, keywords, default=''):
        if isinstance(keywords, str):
            keywords = [keywords]
        
        # 1. Try to find the first column that contains ANY of the keywords
        for keyword in keywords:
            col = next((c for c in row.index if keyword in str(c)), None)
            if col:
                val = row[col]
                if pd.isna(val):
                    return default
                return val
        return default

    def format_bullets(text):
        if not text:
            return ""
        lines = []
        for line in str(text).split('\n'):
            stripped = line.strip()
            if stripped.startswith('-'):
                # remove the first dash and wrap in styled HTML
                content = stripped[1:].strip()
                lines.append(f'<div class="bullet-line"><span class="bullet-dot"></span><span class="bullet-text">{content}</span></div>')
            else:
                if stripped:
                    lines.append(f'<div>{stripped}</div>')
                else:
                    lines.append('<div style="height: 10px;"></div>')
        return ''.join(lines)

    def format_class_date(date_val):
        if not date_val or str(date_val).lower() == 'nan':
            return "날짜 미표기"
        
        # datetime 객체인 경우
        if hasattr(date_val, 'strftime'):
            return f"{date_val.month}월 {date_val.day}일 수업"
            
        date_str = str(date_val).strip()
        
        # "3/5, 3/6" 형태 처리
        import re
        matches = re.findall(r'(\d+)/(\d+)', date_str)
        if matches:
            month_map = {}
            for m, d in matches:
                if m not in month_map:
                    month_map[m] = []
                month_map[m].append(d)
            
            result_parts = []
            for m in month_map:
                days = ', '.join(month_map[m])
                result_parts.append(f"{m}월 {days}일")
            
            return f"{', '.join(result_parts)} 수업"
        return f"{date_str} 수업"

    # 명시적 자원 해제를 위해 openpyxl은 엔진으로만 사용하거나 필요시 close
    def get_val(row, keywords, default=''):
        if isinstance(keywords, str):
            keywords = [keywords]
        
        for keyword in keywords:
            col = next((c for c in row.index if keyword in str(c)), None)
            if col:
                val = row[col]
                if pd.isna(val) or str(val).lower() == 'nan':
                    return default
                return str(val).strip()
        return default

    for i, (_, row) in enumerate(df_merged.iterrows()):
        # Exception Logic Handling
        test_held_raw = get_val(row, ['테스트실시', '테스트진행', '테스트여부'], False)
        if test_held_raw is False: # fallback if column not found at all
            test_held = False
        else:
            test_held = str(test_held_raw).strip().upper() not in ['FALSE', 'N', 'X', '미실시', '0']

        test_score = get_val(row, ['테스트점수', '시험점수', '결과점수'])
        test_max = get_val(row, ['만점', '기준점수', '최대점수'])
        
        # Test Status mapping
        if not test_held:
            display_test_score = "미실시"
            display_test_percent = 0
            test_status = "미실시"
        elif str(test_score).strip() == '':
            display_test_score = "미응시"
            display_test_percent = 0
            test_status = "미응시"
        else:
            try:
                # Extract numeric score
                s_val = str(test_score).replace('점', '').strip()
                display_test_score = int(float(s_val))
                
                # Max Score for Gauge Math (Relative scale)
                if str(test_max).strip() != '':
                    m_val = str(test_max).replace('점', '').strip()
                    max_val = float(m_val)
                    if max_val > 0:
                        # Scaling: (Score / Max) * 100
                        display_test_percent = int((display_test_score / max_val) * 100)
                    else:
                        display_test_percent = 0
                else: 
                    # Default to 100-base if no max found
                    display_test_percent = display_test_score 
            except ValueError:
                display_test_score = str(test_score)
                display_test_percent = 0
            test_status = "응시"

        # 추가 테스트 지표 추출
        total_q = get_val(row, ['전체문항수', '총문항수'])
        obj_q = get_val(row, ['객관식문항수', '객관식'])
        subj_q = get_val(row, ['주관식문항수', '주관식'])
        difficulty = get_val(row, ['난이도', '테스트난이도'])

        try:
            if difficulty and str(difficulty).strip() != '' and str(difficulty).lower() != 'nan':
                difficulty_val = f"{float(difficulty):.2f}"
            else:
                difficulty_val = "-"
        except ValueError:
            difficulty_val = str(difficulty)

        # Absence & Notes
        absence_reason = get_val(row, ['결석사유', '불참사유'])
        if str(absence_reason).strip() == '' or str(absence_reason).lower() == 'nan':
            absence_reason = "-"

        special_note = get_val(row, ['특이사항', '기타사항', '비고'])
        if str(special_note).strip() == '' or str(special_note).lower() == 'nan':
            special_note = "당일 특이사항 없습니다."

        notice = get_val(row, ['공지사항', '전달사항'])
        if str(notice).strip() == '' or str(notice).lower() == 'nan':
            notice = "별도 공지사항 없습니다."

        # Date formatting safely
        date_raw = get_val(row, ['날짜', '일시'])
        date_display = format_class_date(date_raw)
        
        # 파일 저장용 date_str (여전히 YYYY-MM-DD 형식이 필요할 수 있으므로)
        if hasattr(date_raw, 'strftime'):
            date_str = date_raw.strftime('%Y-%m-%d')
        else:
            # "3/5, 3/6" 같은 경우 파일 경로에 부적절한 문자가 있을 수 있으므로 정제
            date_str = str(date_raw).replace('/', '-').replace(',', '').replace(' ', '_')[:20]

        # Homework safely convert to int if possible, else keep as text
        hw_raw = get_val(row, ['과제이행도', '과제수행', '숙제'])
        if str(hw_raw).strip() == '' or str(hw_raw).lower() == 'nan':
            hw_completion = 0
            hw_text = "0"
            hw_is_percent = True
            hw_status = "미완료"
        else:
            hw_str = str(hw_raw).strip()
            try:
                # Remove % and convert
                val_clean = hw_str.replace('%', '')
                raw_float = float(val_clean)
                
                # Handle ratio floats (e.g., 0.78 -> 78%) vs whole numbers (e.g., 78 -> 78%)
                # If there's no % sign and it's a decimal <= 1.0, treat as ratio
                if 0 < raw_float <= 1.0 and '%' not in hw_str:
                     hw_completion = int(round(raw_float * 100))
                else:
                     hw_completion = int(raw_float)
                
                hw_text = str(hw_completion)
                hw_is_percent = True
                hw_status = "완료" if hw_completion >= 100 else "부분 완료" if hw_completion > 0 else "미완료"
            except ValueError:
                # It's a text string like "첫 시간", "미확인"
                hw_completion = 0
                hw_text = hw_str
                hw_is_percent = False
                hw_status = hw_str

        class_avg_raw = get_val(row, '반평균')
        if pd.isna(class_avg_raw) or str(class_avg_raw).strip() == '' or str(class_avg_raw).lower() == 'nan':
            class_avg = "-"
        else:
            try:
                class_avg = int(float(str(class_avg_raw).replace('점', '').strip()))
            except ValueError:
                class_avg = str(class_avg_raw)



        # Attendance parsing
        raw_attendance = get_val(row, ['출석', '출결', '출석여부'], '출석')
        if str(raw_attendance).strip() == '' or str(raw_attendance).lower() == 'nan':
            attendance_val = "출석"
        else:
            attendance_val = str(raw_attendance).strip()
        
        # 1. 오늘의 학습 내용
        raw_lesson = get_val(row, '학습내용')
        
        # 2. 다음 시간 과제
        raw_next_hw = get_val(row, '다음시간과제')
        
        # 3. 이전 시간 과제 (강력한 필터링 적용)
        raw_prev_hw = ""
        # 키워드 우선순위
        potential_vals = [
            get_val(row, k) for k in ['이전시간과제', '이전시간 과제', '이전 과제', '이전과제', '지난과제', '전시간과제', '지난 시간 과제']
        ]
        for val in potential_vals:
            if val and not val.isdigit() and val not in ["첫시간", "미응시", "A", "B", "C", "D"]:
                raw_prev_hw = val
                break
        
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
            "lesson_content": format_bullets(raw_lesson),
            "next_homework": format_bullets(raw_next_hw),
            "special_notes": format_bullets(str(special_note).replace('nan', '')),
            "announcements": format_bullets(str(notice).replace('nan', '')),
            "date_display": date_display,
            "total_q": total_q if str(total_q).strip() != '' else "-",
            "obj_q": obj_q if str(obj_q).strip() != '' else "-",
            "subj_q": subj_q if str(subj_q).strip() != '' else "-",
            "difficulty": difficulty_val,
            "prev_homework_content": raw_prev_hw.replace('\n', '<br>') if raw_prev_hw else "ㅡ",
        }
        
        reports_data.append(data)
        
    xls.close()
    return reports_data, None

# HTML rendering to PNG
def generate_images(reports_data, job_id, inline_css, template):
    job_output_dir = os.path.join(OUTPUT_DIR, job_id)
    os.makedirs(job_output_dir, exist_ok=True)
    
    generated_files = []
    progress_store[job_id] = {"percent": 0, "status": "준비 중..."}
    
    try:
        with sync_playwright() as p:
            def launch_browser():
                b = p.chromium.launch(
                    headless=True,
                    args=[
                        '--disable-dev-shm-usage',
                        '--no-sandbox',
                        '--disable-setuid-sandbox',
                        '--disable-gpu',
                        '--single-process',
                        '--js-flags="--max-old-space-size=256"',
                        '--disable-extensions',
                        '--disable-component-update',
                        '--no-zygote'
                    ]
                )
                c = b.new_context(viewport={"width": 1080, "height": 1560}, device_scale_factor=2.0)
                pg = c.new_page()
                pg.set_default_timeout(300000)
                return b, c, pg

            browser, context, page = launch_browser()
            
            total_reports = len(reports_data)
            for i, data in enumerate(reports_data):
                progress_store[job_id] = {
                    "percent": int((i / total_reports) * 95),
                    "status": f"[{i+1}/{total_reports}] 리포트 생성 중: {data['student_name']}..."
                }
                
                # Context restart logic - Re-launch ENTIRE browser every 50 reports to clear memory
                if i > 0 and i % 50 == 0:
                    try:
                        page.close()
                        context.close()
                        browser.close()
                    except: pass
                    gc.collect()
                    time.sleep(0.5) # Reduced cooling time for faster batching
                    browser, context, page = launch_browser()

                try:
                    html_content = template.render(**data)
                    html_content = html_content.replace('</head>', f'<style>{inline_css}</style></head>')
                    page.set_content(html_content, wait_until='domcontentloaded')
                    
                    date_dir = os.path.join(job_output_dir, data['date'])
                    class_dir = os.path.join(date_dir, data['class_name'])
                    os.makedirs(class_dir, exist_ok=True)
                    
                    clean_school = str(data['school']).replace('등학교', '').strip()
                    png_filename = f"{data['student_name']}({clean_school}{data['grade']}).png"
                    png_path = os.path.join(class_dir, png_filename)
                    
                    page.screenshot(path=png_path, full_page=True)
                    generated_files.append(png_path)
                except Exception as e_inner:
                    # If page fails, try one re-launch
                    print(f"Error during page render ({data['student_name']}): {e_inner}. Retrying...")
                    try: browser.close()
                    except: pass
                    browser, context, page = launch_browser()
                    page.set_content(html_content, wait_until='load')
                    page.screenshot(path=png_path, full_page=True)
                    generated_files.append(png_path)

                gc.collect()
                
            try:
                page.close()
                context.close()
                browser.close()
            except: pass

        # ZIP creation
        progress_store[job_id] = {"percent": 95, "status": "압축 파일 생성 중..."}
        today_str = datetime.datetime.now().strftime('%Y%m%d')
        zip_filename = f"report_{today_str}_{job_id}.zip"
        zip_path = os.path.join(OUTPUT_DIR, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, arc_files in os.walk(job_output_dir):
                for f in arc_files:
                    abs_file = os.path.join(root, f)
                    arc_name = os.path.relpath(abs_file, start=job_output_dir)
                    zipf.write(abs_file, arcname=arc_name)
        
        # Immediate cleanup of PNGs
        if os.path.exists(job_output_dir):
            shutil.rmtree(job_output_dir)
            
        progress_store[job_id] = {"percent": 100, "status": "완료!", "zip": zip_filename}
        return True
    except Exception as e:
        progress_store[job_id] = {"percent": 0, "status": f"오류 발생: {str(e)}", "error": True}
        return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return {"error": "파일을 첨부해주세요."}, 400
        
        file = request.files['file']
        if file.filename == '':
            return {"error": "선택된 파일이 없습니다."}, 400
            
        job_id = uuid.uuid4().hex[:8]
        file_path = os.path.join(INPUT_DIR, f"{job_id}_{file.filename}")
        file.save(file_path)
        
        reports_data, err = process_excel(file_path)
        if err:
            if os.path.exists(file_path): os.remove(file_path)
            return {"error": err}, 400
            
        # Resources for worker
        css_path = os.path.join(BASE_DIR, 'static', 'style.css')
        with open(css_path, 'r', encoding='utf-8') as f:
            inline_css = f.read()
        env = Environment(loader=FileSystemLoader(TEMPLATES_DIR))
        template = env.get_template('report.html')

        def worker():
            try:
                generate_images(reports_data, job_id, inline_css, template)
            except Exception as e:
                progress_store[job_id] = {"percent": 0, "status": f"오류: {str(e)}", "error": True}
                
            # Cleanup after 10 minutes
            time.sleep(600)
            if job_id in progress_store:
                job_info = progress_store.pop(job_id)
                zip_name = job_info.get('zip')
                if zip_name:
                    zpath = os.path.join(OUTPUT_DIR, zip_name)
                    if os.path.exists(zpath): os.remove(zpath)
            if os.path.exists(file_path): os.remove(file_path)

        threading.Thread(target=worker, daemon=True).start()
        return {"job_id": job_id}, 200
        
    return render_template('index.html')

@app.route('/status/<job_id>')
def get_status(job_id):
    info = progress_store.get(job_id)
    if not info:
        return {"error": "Job not found", "percent": 0}, 404
    return info


@app.route('/download_job/<job_id>')
def download_job(job_id):
    info = progress_store.get(job_id)
    if info and info.get('percent') == 100:
        zip_name = info['zip']
        zpath = os.path.join(OUTPUT_DIR, zip_name)
        if os.path.exists(zpath):
            return send_file(zpath, as_attachment=True)
    return "파일이 아직 준비되지 않았거나 만료되었습니다.", 404

if __name__ == '__main__':
    app.run(debug=True, port=int(os.environ.get('PORT', 5000)))
