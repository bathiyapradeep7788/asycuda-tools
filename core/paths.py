"""Working folders (cloud container eke temp folder). Real data Supabase storage eke thiyenawa."""
import os

WORK_DIR = os.environ.get("ASYCUDA_WORK_DIR", "/tmp/asycuda_work")
DATA_XLSX = os.path.join(WORK_DIR, "data.xlsx")
PDF_DIR = os.path.join(WORK_DIR, "pdf")
MERGED_DIR = os.path.join(WORK_DIR, "merged")
LOGS_DIR = os.path.join(WORK_DIR, "logs")
EXPORTERS_XLSX = os.path.join(WORK_DIR, "exporters.xlsx")
RECALC_DIR = os.path.join(WORK_DIR, "recalc")


def ensure_dirs():
    for d in (WORK_DIR, PDF_DIR, MERGED_DIR, LOGS_DIR, RECALC_DIR):
        os.makedirs(d, exist_ok=True)


ensure_dirs()
