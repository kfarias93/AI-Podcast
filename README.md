# 🎙️ AI Podcast Generator

Upload a PDF, Word (.docx) or text (.txt) document and get back a downloadable podcast: a natural
conversation between two AI hosts (**Maya** and **Sam**) who discuss the document's main ideas, with a
different voice for each host.

Built with **Python + Streamlit** and the official **OpenAI Python SDK**.

## What you need
- **Python 3.10 or newer** – check with `python --version` (download from python.org if needed)
- An **OpenAI API key** with a little billing credit – create one at <https://platform.openai.com/api-keys>
  (each podcast costs a small amount of credit; see OpenAI's pricing page)
- *(Optional)* **ffmpeg**, only if you also want an MP3 download

## 1. Install

Open a terminal in the project folder (the one containing `app.py`).

**Windows (PowerShell)**
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
```

**macOS / Linux**
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
```

## 2. Add your API key
Open the new **`.env`** file in any text editor and replace `your_api_key_here` with your real key:
```
OPENAI_API_KEY=sk-...your key...
```
Save the file. The key is only ever read from this file / the environment – it is never written in the code.
**Never share `.env` or upload it to GitHub** (it is already listed in `.gitignore`).

## 3. Run
```bash
streamlit run app.py
```
Your browser opens at <http://localhost:8501>. (To stop the app press `Ctrl + C` in the terminal.)

## 4. Use it
1. **Upload** a document (try `sample_documents/sleep_and_learning.pdf`). You'll see the file name, word count and a preview.
2. **Choose** a length (Short 2–3 min / Standard 4–5 min) and a tone (Conversational / Educational).
3. Click **Generate Podcast** and watch the progress bar (usually 30–90 seconds).
4. **Listen** with the built-in player, read the script, then use **Download Podcast** (WAV, plus MP3 if ffmpeg is installed) or **Download Script**.

### Optional: MP3 downloads
Install ffmpeg and restart the app; an MP3 button appears automatically.
Windows: `winget install ffmpeg` · macOS: `brew install ffmpeg` · Ubuntu/Debian: `sudo apt install ffmpeg`

### Optional: change models
Add to `.env` (defaults shown):
```
OPENAI_TEXT_MODEL=gpt-4o-mini
OPENAI_TTS_MODEL=gpt-4o-mini-tts
```

## Put it online
See **DEPLOY.md** for free, no-zip hosting on Streamlit Community Cloud (a shareable URL).

## How it works
```
document ──► extract text ──► (long? summarise each ~12,000-char section) ──► write 2-host script (JSON)
         ──► text-to-speech per line (Maya = "nova", Sam = "onyx") ──► join with short pauses ──► WAV (+MP3)
```
| File | Purpose |
|---|---|
| `app.py` | The Streamlit interface |
| `podcast_core.py` | Text extraction, summarising, script writing, text-to-speech, audio joining, error handling |
| `tests/` | Automated tests (run with the command below) |
| `sample_documents/` | A sample PDF, DOCX and TXT to try |
| `docs/ASSIGNMENT_NOTES.md` | Write-up material for the university assignment |

Libraries: `pypdf` (PDF), `python-docx` (Word), `python-dotenv` (.env), `openai` (script + voices), Python's built-in `wave` (joining audio).

## Run the tests
```bash
python -m unittest discover -s tests -t . -v
```
The tests use a fake OpenAI client, so they are free and need no internet or API key.

## Troubleshooting
| Message | Fix |
|---|---|
| "OpenAI API key not found" | Make sure `.env` is in the same folder as `app.py`, contains `OPENAI_API_KEY=...`, and restart the app |
| "rejected your API key" | The key is wrong or revoked – create a new one |
| "run out of credit" | Add billing credit in your OpenAI account |
| "No readable text in this PDF" | The PDF is a scan (images). Use a text-based PDF – OCR is not included |
| "model was not found / no access" | Set `OPENAI_TEXT_MODEL` / `OPENAI_TTS_MODEL` in `.env` to models your account can use |
| `streamlit: command not found` | Activate the virtual environment first (see step 1) |

## Privacy note
The document's text is sent to OpenAI to create the script and audio. Don't upload confidential material you aren't allowed to share.
AI-generated audio should be described as AI-generated when you share it.
