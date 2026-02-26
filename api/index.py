"""
Vercel serverless entry point.
Copies the bundled SQLite DB to /tmp on cold start, then exposes the Flask app.
"""
import shutil
import os
import sys
from pathlib import Path

# Mark as Vercel environment BEFORE importing app modules
os.environ["VERCEL"] = "1"

# Ensure project root is on sys.path so `from euromilhoes import ...` works
PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

# Copy bundled DB to /tmp on cold start (Vercel filesystem is read-only except /tmp)
_src_db = PROJECT_ROOT / "euromilhoes.db"
_tmp_db = Path("/tmp/euromilhoes.db")
if _src_db.exists() and not _tmp_db.exists():
    shutil.copy2(_src_db, _tmp_db)

# Create temp dirs needed by the app
Path("/tmp/chaves_geradas").mkdir(exist_ok=True)

# Import Flask app — Vercel Python runtime picks up the `app` variable
from app import app  # noqa: E402
