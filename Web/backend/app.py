import logging
import os
import shutil
import sys
import tempfile
import threading
import time
from collections import defaultdict, deque
from pathlib import Path
from uuid import UUID, uuid4

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from Web.backend.generator import create_outputs


logging.basicConfig(level=logging.INFO)
log = logging.getLogger("fillerbunny.web")
PROJECT_ROOT = Path(__file__).resolve().parents[2]
WEB_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "Product" / "Blackbox"))
from dispatcher import MAX_ACTIONS, MAX_PROMPT_LENGTH, parse_prompt  # noqa: E402

JOB_ROOT = Path(tempfile.gettempdir()) / "fillerbunny-web-jobs"
JOB_TTL_SECONDS = 60 * 60
REQUESTS_PER_MINUTE = 5
_request_times: dict[str, deque[float]] = defaultdict(deque)
_rate_lock = threading.Lock()
_generation_lock = threading.Lock()

app = FastAPI(title="Fillerbunny Generator", docs_url=None, redoc_url=None)
origins = [origin.strip() for origin in os.getenv("CORS_ORIGINS", "*").split(",") if origin.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins or ["*"],
    allow_methods=["GET", "POST"],
    allow_headers=["Content-Type"],
)


class GenerateRequest(BaseModel):
    prompt: str = Field(min_length=5, max_length=MAX_PROMPT_LENGTH)


def _check_rate_limit(client_id: str) -> None:
    now = time.monotonic()
    with _rate_lock:
        times = _request_times[client_id]
        while times and now - times[0] > 60:
            times.popleft()
        if len(times) >= REQUESTS_PER_MINUTE:
            raise HTTPException(status_code=429, detail="Please wait a minute before generating again.")
        times.append(now)


def _remove_expired_jobs() -> None:
    JOB_ROOT.mkdir(parents=True, exist_ok=True)
    expiry = time.time() - JOB_TTL_SECONDS
    for job_dir in JOB_ROOT.iterdir():
        if job_dir.is_dir() and job_dir.stat().st_mtime < expiry:
            shutil.rmtree(job_dir)


@app.get("/health")
def health() -> dict:
    return {"status": "ok"}


@app.post("/api/generate")
def generate(request: GenerateRequest, http_request: Request) -> dict:
    client_id = http_request.client.host if http_request.client else "unknown"
    _check_rate_limit(client_id)
    prompt = request.prompt.strip()
    if len(prompt) < 5:
        raise HTTPException(status_code=422, detail="Enter a prompt with at least five characters.")

    actions = parse_prompt(prompt)
    if not actions:
        raise HTTPException(
            status_code=422,
            detail="No supported actions found. Try: “A red fox appears, then spins 180 degrees.”",
        )
    if len(actions) > MAX_ACTIONS:
        raise HTTPException(status_code=422, detail=f"Please limit your prompt to {MAX_ACTIONS} actions.")

    if not _generation_lock.acquire(blocking=False):
        raise HTTPException(status_code=503, detail="A presentation is being generated. Please try again shortly.")

    job_id = str(uuid4())
    job_dir = JOB_ROOT / job_id
    try:
        _remove_expired_jobs()
        job_dir.mkdir(parents=True)
        pptx_path = job_dir / "presentation.pptx"
        mp4_path = job_dir / "animation.mp4"
        create_outputs(actions, pptx_path, mp4_path)
    except (OSError, RuntimeError, ValueError) as exc:
        if job_dir.exists():
            shutil.rmtree(job_dir)
        log.exception("Failed to generate outputs for job %s", job_id)
        raise HTTPException(status_code=500, detail="Generation failed. Please simplify the prompt and try again.") from exc
    finally:
        _generation_lock.release()

    return {
        "job_id": job_id,
        "actions": len(actions),
        "files": [
            {"name": "presentation.pptx", "url": f"/api/jobs/{job_id}/presentation.pptx"},
            {"name": "animation.mp4", "url": f"/api/jobs/{job_id}/animation.mp4"},
        ],
    }


@app.get("/api/jobs/{job_id}/{filename}")
def download(job_id: str, filename: str) -> FileResponse:
    try:
        canonical_id = str(UUID(job_id))
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="Download not found or expired.") from exc
    if filename not in {"presentation.pptx", "animation.mp4"}:
        raise HTTPException(status_code=404, detail="Download not found or expired.")

    path = JOB_ROOT / canonical_id / filename
    if not path.is_file() or time.time() - path.stat().st_mtime > JOB_TTL_SECONDS:
        raise HTTPException(status_code=404, detail="Download not found or expired.")
    media_type = (
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        if filename.endswith(".pptx") else "video/mp4"
    )
    return FileResponse(path, media_type=media_type, filename=filename)


app.mount("/", StaticFiles(directory=WEB_DIR, html=True), name="website")
