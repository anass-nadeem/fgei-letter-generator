# FGEI Official Letter Generator

> Automated document generation system for military-grade official correspondence — eliminating manual formatting errors and reducing letter preparation time from 15+ minutes to under 60 seconds.

![Python](https://img.shields.io/badge/Python-3.12-blue?logo=python)
![Flask](https://img.shields.io/badge/Flask-3.0-black?logo=flask)
![License](https://img.shields.io/badge/License-MIT-green)

## Problem

Military and government institutions spend significant time manually formatting official letters — ensuring correct classification headers, monogram placement, address blocks, body spacing, and signatory positioning according to strict FGEI correspondence standards. A single formatting mistake requires the entire letter to be retyped.

## Solution

A web-based letter generation system where staff enter data into a structured form and receive a perfectly formatted, print-ready `.docx` file in seconds — every time, with zero formatting errors.

## Screenshots

<img width="916" height="481" alt="123" src="https://github.com/user-attachments/assets/3db3a4ee-2940-4088-b788-a070246fb11a" />
<img width="271" height="340" alt="234" src="https://github.com/user-attachments/assets/b7cd89ec-d3f0-4d95-bd6a-9595f357f173" />

## Live Demo

https://letter-generator-4gr7.onrender.com/

## Key Features

- ✅ Pixel-perfect replication of official FGEI letter format
- ✅ Real-time live preview as you type
- ✅ Bold text support using `**text**` markdown syntax
- ✅ Supports RESTD, CONFIDENTIAL, SECRET, UNCLASSIFIED classifications
- ✅ Correct Arial 12pt font enforced at XML level — overrides Word defaults
- ✅ One-click `.docx` download
- ✅ Deployable as `.exe` for offline use or hosted for team access

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python, Flask |
| Document Engine | python-docx, lxml |
| Frontend | HTML5, CSS3, Vanilla JS |
| Deployment | PyInstaller (desktop), Render (web) |

## Technical Highlights

- Built a **custom Word 2016 template from scratch** in Python to bypass MS Word compatibility mode — resolving a deep font-override issue caused by `w:asciiTheme` in `docDefaults`
- Implemented **raw XML manipulation** of `.docx` internals (theme fonts, paragraph spacing, tab stops) using `lxml` — going beyond what `python-docx` exposes at the API level
- Packaged as a **standalone `.exe`** using PyInstaller for offline deployment on client machines with no Python required

## Installation
```bash
git clone https://github.com/anass-nadeem/fgei-letter-generator
cd fgei-letter-generator
pip install -r requirements.txt
python app.py
```

Open `http://localhost:5000` in your browser.

## Project Structure
```
letter_app/
├── app.py              # Flask backend + DOCX builder
├── requirements.txt    # Dependencies
├── launcher.py         # PyInstaller entry point
└── templates/
    └── index.html      # Live preview form
```

## License

MIT
