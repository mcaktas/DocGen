import json
import subprocess
from pathlib import Path
from jinja2 import Environment, FileSystemLoader

# --- Paths ---
DATA_PATH = Path("data/data.json")
TEMPLATE_PATH = "JinjaTemplates/template.md.j2"
OUT_MD_DIR = Path("Output_MarkDown")
OUT_DOCX_DIR = Path("Output_Docx")

OUT_MD_PATH = OUT_MD_DIR / "report.md"
OUT_DOCX_PATH = OUT_DOCX_DIR / "report.docx"

# --- Ensure output folders exist ---
OUT_MD_DIR.mkdir(parents=True, exist_ok=True)
OUT_DOCX_DIR.mkdir(parents=True, exist_ok=True)

# --- Load data ---
data = json.loads(DATA_PATH.read_text(encoding="utf-8"))

# --- Render Markdown via Jinja ---
env = Environment(
    loader=FileSystemLoader("."),
    autoescape=False,  # markdown, not html
    trim_blocks=True,
    lstrip_blocks=True,
)

tpl = env.get_template(TEMPLATE_PATH)
out = tpl.render(**data)

OUT_MD_PATH.write_text(out, encoding="utf-8")
print(f"Wrote Markdown: {OUT_MD_PATH}")

# --- Convert Markdown -> DOCX using Pandoc (no template) ---
# Requires: pandoc installed and on PATH
try:
    REFERENCE_DOC = Path("WordTemplates/reference.docx")

    subprocess.run(
        ["pandoc", "report.md",
         "--reference-doc", str(REFERENCE_DOC.resolve()),
         "-o", str(OUT_DOCX_PATH.resolve())],
        check=True,
        cwd=str(OUT_MD_DIR),
    )

    print(f"Wrote DOCX: {OUT_DOCX_PATH}")
except FileNotFoundError:
    raise SystemExit("Pandoc not found. Install it and ensure `pandoc` is on your PATH.")
except subprocess.CalledProcessError as e:
    raise SystemExit(f"Pandoc conversion failed with exit code {e.returncode}.")
