"""LibreOffice helpers (Excel COM wenuwata). Cloud eke Excel naha, ethakota LibreOffice headless use karanawa."""
import os
import shutil
import subprocess
import tempfile


class OfficeError(RuntimeError):
    pass


def _soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    raise OfficeError("LibreOffice hambune naha (packages.txt eke libreoffice-calc thiyenna ona).")


def convert(src, fmt, out_dir, timeout=240):
    """src file eka LibreOffice walin fmt ('pdf' | 'xlsx') ekata convert karala path eka return karanawa."""
    os.makedirs(out_dir, exist_ok=True)
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        cmd = [_soffice(), f"-env:UserInstallation=file://{profile}", "--headless", "--norestore",
               "--convert-to", fmt, "--outdir", out_dir, src]
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout)
        out = os.path.join(out_dir, os.path.splitext(os.path.basename(src))[0] + "." + fmt.split(":")[0])
        if not os.path.exists(out) or os.path.getsize(out) == 0:
            raise OfficeError(f"LibreOffice convert fail: {r.stderr.strip() or r.stdout.strip()}")
        return out
    except subprocess.TimeoutExpired:
        raise OfficeError("LibreOffice timeout una.")
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def recalculated_copy(xlsx_path, out_dir):
    """Formulas recalculate karapu copy ekak (cached values ekka). mail.py walata AA3/AA4 gannanna."""
    tmp_in = os.path.join(out_dir, "_recalc_in.xlsx")
    os.makedirs(out_dir, exist_ok=True)
    shutil.copy2(xlsx_path, tmp_in)
    sub_out = os.path.join(out_dir, "_out")
    shutil.rmtree(sub_out, ignore_errors=True)
    return convert(tmp_in, "xlsx", sub_out)
