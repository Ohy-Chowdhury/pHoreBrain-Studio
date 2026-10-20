# pHoreBrain Studio — MVP 0.1

A starter educational content builder inspired by question-to-document and slide workflows.

## Features
- Add MCQs with four options, answer, and optional explanation
- Preview/delete questions
- Export question bank as JSON
- Generate editable DOCX question sheets
- Generate editable PPTX slide decks
- Toggle answers in exported files

## Run locally (Windows)
1. Install Python 3.11 or 3.12 from https://www.python.org/downloads/
2. Extract this ZIP.
3. Open the extracted folder.
4. Click the folder address bar in File Explorer, type `cmd`, and press Enter.
5. Run:

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
streamlit run app.py
```

Your browser should open the app. If not, open the local URL shown in the terminal, usually `http://localhost:8501`.

## Run on Mac/Linux
```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## Free cloud deployment
You can deploy from a GitHub repository using Streamlit Community Cloud:
https://docs.streamlit.io/deploy/streamlit-community-cloud/deploy-your-app

Upload `app.py`, `requirements.txt`, and `README.md` to a GitHub repository, then select `app.py` as the app entrypoint.

## Important limitations
- MVP questions are held in Streamlit session state, not a permanent database.
- Do not use it as the only copy of important question content.
- Review answer keys and exported formatting before sharing with students.
- PDF export, bulk import, permanent storage, login, and online quiz mode are planned future features.

## Suggested next milestones
1. Add import from JSON/CSV and persistent question bank.
2. Add polished Biology/Chemistry/Physics templates.
3. Add PDF export.
4. Add image/diagram support and equations.
5. Add student quiz mode and analytics.
