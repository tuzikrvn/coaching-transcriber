"""
Coaching Transcriber — FastAPI backend.

Flow:
  POST /upload   → save file, start background transcription, return {job_id}
  GET  /status/{job_id} → poll for progress
  GET  /download/{job_id} → download finished .xlsx
"""
import os
import uuid
import secrets
from fastapi import FastAPI, UploadFile, File, BackgroundTasks, HTTPException, Depends
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.security import HTTPBasic, HTTPBasicCredentials
import assemblyai as aai

from excel_utils import create_excel, detect_pauses

# ── Config ──────────────────────────────────────────────────────────────────
ASSEMBLYAI_API_KEY = os.environ.get("ASSEMBLYAI_API_KEY", "")
PAUSE_THRESHOLD_SEC = float(os.environ.get("PAUSE_THRESHOLD_SEC", "2.0"))
APP_USER     = os.environ.get("APP_USER", "coach")
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")

ALLOWED_EXTENSIONS = {".mp3", ".mp4", ".m4a", ".wav", ".webm", ".ogg", ".flac", ".aac"}
MAX_FILE_SIZE_MB = 500

# ── App ──────────────────────────────────────────────────────────────────────
app = FastAPI(title="Coaching Transcriber")
security = HTTPBasic()

# In-memory job store — sufficient for single-instance free-tier deployment
jobs: dict = {}


# ── Auth ─────────────────────────────────────────────────────────────────────

def require_auth(credentials: HTTPBasicCredentials = Depends(security)):
    """HTTP Basic Auth. Пропускается если APP_PASSWORD не задан (локальная разработка)."""
    if not APP_PASSWORD:
        return  # локально — без пароля
    ok_user = secrets.compare_digest(credentials.username.encode(), APP_USER.encode())
    ok_pass = secrets.compare_digest(credentials.password.encode(), APP_PASSWORD.encode())
    if not (ok_user and ok_pass):
        raise HTTPException(
            status_code=401,
            detail="Неверный логин или пароль",
            headers={"WWW-Authenticate": "Basic"},
        )


# ── Routes ───────────────────────────────────────────────────────────────────

@app.get("/", response_class=HTMLResponse, dependencies=[Depends(require_auth)])
async def index():
    html_path = os.path.join(os.path.dirname(__file__), "templates", "index.html")
    with open(html_path, "r", encoding="utf-8") as f:
        return HTMLResponse(f.read())


@app.post("/upload", dependencies=[Depends(require_auth)])
async def upload_audio(
    background_tasks: BackgroundTasks,
    file: UploadFile = File(...),
):
    ext = os.path.splitext(file.filename or "")[1].lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            400,
            f"Неподдерживаемый формат. Можно загружать: {', '.join(sorted(ALLOWED_EXTENSIONS))}"
        )

    content = await file.read()
    size_mb = len(content) / (1024 * 1024)
    if size_mb > MAX_FILE_SIZE_MB:
        raise HTTPException(
            400,
            f"Файл слишком большой ({size_mb:.0f} МБ). Максимум — {MAX_FILE_SIZE_MB} МБ."
        )

    job_id = str(uuid.uuid4())
    audio_path = f"/tmp/{job_id}{ext}"

    with open(audio_path, "wb") as f:
        f.write(content)

    jobs[job_id] = {
        "status": "processing",
        "step": 1,
        "message": "Файл получен, отправляю на транскрибацию…",
        "filename": file.filename,
        "output_path": None,
        "error": None,
    }

    background_tasks.add_task(_process_audio, job_id, audio_path, file.filename or "session")
    return {"job_id": job_id}


@app.get("/status/{job_id}", dependencies=[Depends(require_auth)])
async def get_status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(404, "Задача не найдена")
    job = jobs[job_id]
    return {
        "status": job["status"],   # processing | done | error
        "step": job["step"],       # 1..3
        "message": job["message"],
        "error": job.get("error"),
    }


@app.get("/download/{job_id}", dependencies=[Depends(require_auth)])
async def download_result(job_id: str):
    if job_id not in jobs:
        raise HTTPException(404, "Задача не найдена")
    job = jobs[job_id]
    if job["status"] != "done":
        raise HTTPException(400, "Файл ещё не готов")

    base = os.path.splitext(job["filename"])[0]
    download_name = f"{base}_транскрипция.xlsx"

    return FileResponse(
        job["output_path"],
        filename=download_name,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Background task ──────────────────────────────────────────────────────────

def _process_audio(job_id: str, audio_path: str, original_filename: str) -> None:
    try:
        if not ASSEMBLYAI_API_KEY:
            raise ValueError("Переменная окружения ASSEMBLYAI_API_KEY не задана.")

        # Step 1 — transcribe
        _update(job_id, step=1, message="Транскрибирую аудио… (обычно 1–3 минуты)")

        aai.settings.api_key = ASSEMBLYAI_API_KEY
        config = aai.TranscriptionConfig(
            speech_models=["universal-3-pro", "universal-2"],  # поддержка 99 языков вкл. русский
            speaker_labels=True,
            punctuate=True,
            format_text=True,
        )
        transcript = aai.Transcriber().transcribe(audio_path, config=config)

        if transcript.status == aai.TranscriptStatus.error:
            raise RuntimeError(f"AssemblyAI: {transcript.error}")

        # Step 2 — build rows
        _update(job_id, step=2, message="Анализирую паузы и формирую таблицу…")
        rows = detect_pauses(transcript.utterances, PAUSE_THRESHOLD_SEC)

        # Step 3 — create Excel
        _update(job_id, step=3, message="Создаю Excel-файл…")
        output_path = f"/tmp/{job_id}_result.xlsx"
        create_excel(rows, output_path, original_filename)

        jobs[job_id].update({
            "status": "done",
            "step": 3,
            "message": "Готово! Файл можно скачать.",
            "output_path": output_path,
        })

    except Exception as exc:
        jobs[job_id].update({
            "status": "error",
            "message": "Произошла ошибка при обработке.",
            "error": str(exc),
        })
    finally:
        if os.path.exists(audio_path):
            os.remove(audio_path)


def _update(job_id: str, step: int, message: str) -> None:
    jobs[job_id]["step"] = step
    jobs[job_id]["message"] = message
