"""Generates profile/master_cv.docx directly from canonical profile/master_cv.md."""
import sys
from pathlib import Path

PROFILE_DIR = Path(__file__).resolve().parent
BASE_DIR = PROFILE_DIR.parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

from core.engine import JobAgentEngine


def build_master_cv_docx():
    md_path = PROFILE_DIR / "master_cv.md"
    docx_path = PROFILE_DIR / "master_cv.docx"
    markdown_text = md_path.read_text(encoding="utf-8")
    engine = JobAgentEngine()
    engine.generate_docx(markdown_text, docx_path)
    print(f"Canonical Master CV saved to: {docx_path}")


if __name__ == "__main__":
    build_master_cv_docx()
