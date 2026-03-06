import os
import shutil
import zipfile
import random
import time
import threading
import uuid
import datetime
import gc
import json
import pandas as pd
from jinja2 import Environment, FileSystemLoader
from flask import Flask, render_template, request, send_file, url_for, flash, redirect
from concurrent.futures import ThreadPoolExecutor

app = Flask(__name__)
app.secret_key = 'super_secret_key_for_flash_messages'

# Directories setup
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, 'input')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')
STATUS_DIR = os.path.join(BASE_DIR, 'status')
PREVIEW_DIR = os.path.join(BASE_DIR, 'static', 'previews')

for d in [INPUT_DIR, OUTPUT_DIR, TEMPLATES_DIR, STATUS_DIR, PREVIEW_DIR]:
    os.makedirs(d, exist_ok=True)

def save_status(job_id, data):
    path = os.path.join(STATUS_DIR, f"{job_id}.json")
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

def load_status(job_id):
    path = os.path.join(STATUS_DIR, f"{job_id}.json")
    if os.path.exists(path):
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except: pass
    return None

# Helper function: Excel parser
def process_excel(file_path):
    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        if len(sheet_names) < 2:
            return None, "Excel parsing error: Make sure the excel file has at least 2 sheets."
            
        student_sheet = next((s for s in sheet_names if '학생' in s), sheet_names[0])
        class_sheet = next((s for s in sheet_names if '분반' in s), sheet_names[1])
        
        df_student_raw = pd.read_excel(xls, sheet_name=student_sheet, header=None)
        df_class_raw = pd.read_excel(xls, sheet_name=class_sheet, header=None)
    except Exception as e:
        return None, f"Excel parsing error: {e}. Make sure the excel file has at least 2 sheets."

    def find_header_and_build_df(df_raw, keyword='분반'):
        header_idx = 0
        for idx, row in df_raw.iterrows():
            if any(keyword in str(cell).replace(' ', '') for cell in row.values if pd.notna(cell)):
                header_idx = idx
                break
        new_header = df_raw.iloc[header_idx]
        df_clean = df_raw[header_idx+1:].copy()
        df_clean.columns = new_header
        df_clean = df_clean.loc[:, df_clean.columns.notna()]
        df_clean = df_clean.loc[:, ~df_clean.columns.astype(str).str.contains('^Unnamed:')]
        df_clean.reset_index(drop=True, inplace=True)
        return df_clean

    df_student = find_header_and_build_df(df_student_raw, '분반')
    df_class = find_header_and_build_df(df_class_raw, '분반')

    df_class.columns = df_class.columns.astype(str).str.replace('\n', '').str.replace(' ', '').str.strip()
    df_student.columns = df_student.columns.astype(str).str.replace('\n', '').str.replace(' ', '').str.strip()
    
    class_col_student = next((col for col in df_student.columns if '분반' in str(col)), None)
    class_col_class = next((col for col in df_class.columns if '분반' in str(col)), None)
    
    if not class_col_student or not class_col_class:
        return None, f"Merge Error: '분반' 이 포함된 열을 찾을 수 없습니다."
        
    df_student.rename(columns={class_col_student: 'merge_class_key'}, inplace=True)
    df_class.rename(columns={class_col_class: 'merge_class_key'}, inplace=True)
    
    try:
        df_merged = pd.merge(df_student, df_class, on='merge_class_key', how='left')
    except Exception as e:
        return None, f"Merge Error: {e}"
    
    reports_data = []

    def format_bullets(text):
        if not text: return ""
        lines = []
        for line in str(text).split('\n'):
            stripped = line.strip()
            if stripped.startswith('-'):
                content = stripped[1:].strip()
                lines.append(f'<div class="bullet-line"><span class="bullet-dot"></span><span class="bullet-text">{content}</span></div>')
            else:
                if stripped: lines.append(f'<div>{stripped}</div>')
                else: lines.append('<div style="height: 10px;"></div>')
        return ''.join(lines)

    def format_class_date(date_val):
        if not date_val or str(date_val).lower() == 'nan': return "날짜 미표기"
        if hasattr(date_val, 'strftime'): return f"{date_val.month}월 {date_val.day}일 수업"
        date_str = str(date_val).strip()
        import re
        matches = re.findall(r'(\d+)/(\d+)', date_str)
        if matches:
            month_map = {}
            for m, d in matches:
                if m not in month_map: month_map[m] = []
                month_map[m].append(d)
            result_parts = [f"{m}월 {', '.join(month_map[m])}일" for m in month_map]
            return f"{', '.join(result_parts)} 수업"
        return f"{date_str} 수업"

    def get_val(row, keywords, default=''):
        if isinstance(keywords, str): keywords = [keywords]
        for keyword in keywords:
            col = next((c for c in row.index if keyword in str(c)), None)
            if col:
                val = row[col]
                if pd.isna(val) or str(val).lower() == 'nan': return default
                return str(val).strip()
        return default

    for i, (_, row) in enumerate(df_merged.iterrows()):
        test_held_raw = get_val(row, ['테스트실시', '테스트진행', '테스트여부'], False)
        if test_held_raw is False: test_held = False
        else: test_held = str(test_held_raw).strip().upper() not in ['FALSE', 'N', 'X', '미실시', '0']

        test_score = get_val(row, ['테스트점수', '시험점수', '결과점수'])
        test_max = get_val(row, ['만점', '기준점수', '최대점수'])
        
        if not test_held:
            display_test_score, display_test_percent, test_status = "미실시", 0, "미실시"
        elif str(test_score).strip() == '':
            display_test_score, display_test_percent, test_status = "미응시", 0, "미응시"
        else:
            try:
                s_val = str(test_score).replace('점', '').strip()
                display_test_score = int(float(s_val))
                if str(test_max).strip() != '':
                    m_val = str(test_max).replace('점', '').strip()
                    max_val = float(m_val)
                    display_test_percent = int((display_test_score / max_val) * 100) if max_val > 0 else 0
                else: display_test_percent = display_test_score 
            except ValueError:
                display_test_score, display_test_percent = str(test_score), 0
            test_status = "응시"

        total_q = get_val(row, ['전체문항수', '총문항수'])
        obj_q = get_val(row, ['객관식문항수', '객관식'])
        subj_q = get_val(row, ['주관식문항수', '주관식'])
        difficulty = get_val(row, ['난이도', '테스트난이도'])
        try: difficulty_val = f"{float(difficulty):.2f}" if difficulty and str(difficulty).strip() != '' and str(difficulty).lower() != 'nan' else "-"
        except ValueError: difficulty_val = str(difficulty)

        absence_reason = get_val(row, ['결석사유', '불참사유'], "-")
        special_note = get_val(row, ['특이사항', '기타사항', '비고'], "당일 특이사항 없습니다.")
        notice = get_val(row, ['공지사항', '전달사항'], "별도 공지사항 없습니다.")

        date_raw = get_val(row, ['날짜', '일시'])
        date_display = format_class_date(date_raw)
        date_str = date_raw.strftime('%Y-%m-%d') if hasattr(date_raw, 'strftime') else str(date_raw).replace('/', '-').replace(',', '').replace(' ', '_')[:20]

        hw_raw = get_val(row, ['과제이행도', '과제수행', '숙제'])
        if str(hw_raw).strip() == '' or str(hw_raw).lower() == 'nan':
            hw_completion, hw_text, hw_is_percent, hw_status = 0, "0", True, "미완료"
        else:
            hw_str = str(hw_raw).strip()
            try:
                val_clean = hw_str.replace('%', '')
                raw_float = float(val_clean)
                if 0 < raw_float <= 1.0 and '%' not in hw_str: hw_completion = int(round(raw_float * 100))
                else: hw_completion = int(raw_float)
                hw_text, hw_is_percent = str(hw_completion), True
                hw_status = "완료" if hw_completion >= 100 else "부분 완료" if hw_completion > 0 else "미완료"
            except ValueError:
                hw_completion, hw_text, hw_is_percent, hw_status = 0, hw_str, False, hw_str

        class_avg_raw = get_val(row, '반평균')
        try: class_avg = int(float(str(class_avg_raw).replace('점', '').strip())) if not pd.isna(class_avg_raw) and str(class_avg_raw).strip() != '' else "-"
        except ValueError: class_avg = str(class_avg_raw)

        raw_attendance = get_val(row, ['출석', '출결', '출석여부'], '출석')
        raw_lesson = get_val(row, '학습내용')
        raw_next_hw = get_val(row, '다음시간과제')
        
        raw_prev_hw = ""
        for k in ['이전시간과제', '이전시간 과제', '이전 과제', '이전과제', '지난과제', '전시간과제', '지난 시간 과제']:
            val = get_val(row, k)
            if val and not val.isdigit() and val not in ["첫시간", "미응시", "A", "B", "C", "D"]:
                raw_prev_hw = val
                break
        
        reports_data.append({
            "date": date_str,
            "student_name": str(get_val(row, '학생', 'Unknown')),
            "school": str(get_val(row, '학교', '')),
            "grade": str(get_val(row, '학년', '')).replace('학년', '').strip(),
            "class_name": str(row.get('merge_class_key', 'Unknown')),
            "attendance": raw_attendance,
            "absence_reason": str(absence_reason),
            "test_status": test_status,
            "homework_status": hw_status,
            "display_test_score": display_test_score,
            "test_score_percent": display_test_percent,
            "class_average": class_avg,
            "homework_completion": hw_completion,
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
        })
    xls.close()
    return reports_data, None

# HTML rendering to PNG - Parallel version for high performance
def generate_images(reports_data, job_id, inline_css, template):
    job_output_dir = os.path.join(OUTPUT_DIR, job_id)
    os.makedirs(job_output_dir, exist_ok=True)
    save_status(job_id, {"percent": 0, "status": "준비 중..."})
    
    total_reports = len(reports_data)
    completed_count = [0]
    counter_lock = threading.Lock()
    first_image_saved = [False]

    def render_single_report(report_data, index):
        from playwright.sync_api import sync_playwright
        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True, args=['--disable-dev-shm-usage', '--no-sandbox', '--disable-gpu'])
                context = browser.new_context(viewport={"width": 1080, "height": 1560}, device_scale_factor=2.0)
                page = context.new_page()
                page.set_default_timeout(60000)

                html_content = template.render(**report_data)
                html_content = html_content.replace('</head>', f'<style>{inline_css}</style></head>')
                page.set_content(html_content, wait_until='domcontentloaded')
                
                date_dir = os.path.join(job_output_dir, report_data['date'])
                class_dir = os.path.join(date_dir, report_data['class_name'])
                os.makedirs(class_dir, exist_ok=True)
                
                clean_school = str(report_data['school']).replace('등학교', '').strip()
                png_filename = f"{report_data['student_name']}({clean_school}{report_data['grade']}).png"
                png_path = os.path.join(class_dir, png_filename)
                page.screenshot(path=png_path, full_page=True)
                
                with counter_lock:
                    if not first_image_saved[0]:
                        preview_path = os.path.join(PREVIEW_DIR, f"{job_id}.png")
                        shutil.copy(png_path, preview_path)
                        first_image_saved[0] = True
                    completed_count[0] += 1
                    save_status(job_id, {
                        "percent": int((completed_count[0] / total_reports) * 95),
                        "status": f"[{completed_count[0]}/{total_reports}] 리포트 생성 완료: {report_data['student_name']}"
                    })
                browser.close()
        except Exception as e:
            print(f"Error rendering report {index}: {e}")

    try:
        with ThreadPoolExecutor(max_workers=3) as executor:
            futures = [executor.submit(render_single_report, data, i) for i, data in enumerate(reports_data)]
            for future in futures: future.result() # Wait for all

        save_status(job_id, {"percent": 95, "status": "압축 파일 생성 중..."})
        today_str = datetime.datetime.now().strftime('%Y%m%d')
        zip_filename = f"report_{today_str}_{job_id}.zip"
        zip_path = os.path.join(OUTPUT_DIR, zip_filename)
        
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, arc_files in os.walk(job_output_dir):
                for f in arc_files:
                    abs_file = os.path.join(root, f)
                    arc_name = os.path.relpath(abs_file, start=job_output_dir)
                    zipf.write(abs_file, arcname=arc_name)
        
        if os.path.exists(job_output_dir): shutil.rmtree(job_output_dir)
        save_status(job_id, {"percent": 100, "status": "완료!", "zip": zip_filename})
        return True
    except Exception as e:
        save_status(job_id, {"percent": 0, "status": f"오류 발생: {str(e)}", "error": True})
        return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files: return {"error": "파일을 첨부해주세요."}, 400
        file = request.files['file']
        if file.filename == '': return {"error": "선택된 파일이 없습니다."}, 400
            
        job_id = uuid.uuid4().hex[:8]
        file_path = os.path.join(INPUT_DIR, f"{job_id}_{file.filename}")
        file.save(file_path)
        
        reports_data, err = process_excel(file_path)
        if err:
            if os.path.exists(file_path): os.remove(file_path)
            return {"error": err}, 400
            
        css_path = os.path.join(BASE_DIR, 'static', 'style.css')
        with open(css_path, 'r', encoding='utf-8') as f: inline_css = f.read()
        env = Environment(loader=FileSystemLoader(TEMPLATES_DIR))
        template = env.get_template('report.html')

        def worker():
            try: generate_images(reports_data, job_id, inline_css, template)
            except Exception as e: save_status(job_id, {"percent": 0, "status": f"오류: {str(e)}", "error": True})
            time.sleep(600)
            job_info = load_status(job_id)
            if job_info and job_info.get('zip'):
                zpath = os.path.join(OUTPUT_DIR, job_info['zip'])
                if os.path.exists(zpath): os.remove(zpath)
            for path in [os.path.join(STATUS_DIR, f"{job_id}.json"), os.path.join(PREVIEW_DIR, f"{job_id}.png"), file_path]:
                if os.path.exists(path):
                    try: os.remove(path)
                    except: pass
        threading.Thread(target=worker, daemon=True).start()
        return {"job_id": job_id}, 200
    return render_template('index.html')

@app.route('/status/<job_id>')
def get_status(job_id):
    info = load_status(job_id)
    return info if info else ({"error": "Job not found", "percent": 0}, 404)

@app.route('/success/<job_id>')
def success(job_id):
    info = load_status(job_id)
    if not info or info.get('percent') != 100: return redirect(url_for('index'))
    return render_template('success.html', job_id=job_id)

@app.route('/download_job/<job_id>')
def download_job(job_id):
    info = load_status(job_id)
    if info and info.get('percent') == 100:
        zpath = os.path.join(OUTPUT_DIR, info['zip'])
        if os.path.exists(zpath): return send_file(zpath, as_attachment=True)
    return "파일이 아직 준비되지 않았거나 만료되었습니다.", 404

if __name__ == '__main__':
    app.run(debug=True, port=int(os.environ.get('PORT', 5000)))
