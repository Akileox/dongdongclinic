# Student Learning Report Generator

A Python script to generate a premium, glassmorphism-style student learning report as a screenshot using HTML, CSS, and Playwright.

## Requirements

* Python 3.8+
* Node.js (for Playwright, optional if using python playwright package solely)

## Installation

```bash
# Install dependencies
pip install -r requirements.txt

# Install playwright browsers
playwright install chromium
```

## Usage

```bash
python app.py
```

This will run the script, render the template `templates/report.html` with sample data, and output a high-resolution PNG to `output/report_sample_dark.png`.
