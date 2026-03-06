import os
from jinja2 import Environment, FileSystemLoader

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates')

env = Environment(loader=FileSystemLoader(TEMPLATES_DIR))
template = env.get_template('report.html')

css_path = os.path.join(BASE_DIR, 'static', 'style.css')
with open(css_path, 'r', encoding='utf-8') as f:
    inline_css = f.read()

data = {
    "date": "2026-03-06",
    "student_name": "Test",
    "school": "Test",
    "grade": "1",
    "class_name": "Class",
    "attendance": "출석",
    "absence_reason": "-",
    "test_status": "응시",
    "homework_status": "완료",
    "display_test_score": "100",
    "test_score_percent": 100,
    "class_average": "90",
    "homework_completion": 100,
    "homework_text": "100",
    "homework_is_percent": True,
    "lesson_content": "Test",
    "next_homework": "Test",
    "special_notes": "Test",
    "announcements": "Test",
    "inline_css": inline_css
}

html = template.render(**data)
with open('test_output.html', 'w', encoding='utf-8') as f:
    f.write(html)
print("Done")
