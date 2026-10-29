import os
import shutil
import subprocess


def recalcular(xlsx, pasta):
    """Recalcula no LibreOffice (se houver) e devolve o caminho do arquivo recalculado."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    out = os.path.join(pasta, "recalc")
    os.makedirs(out, exist_ok=True)
    subprocess.run([soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, xlsx],
                   capture_output=True, timeout=180)
    p = os.path.join(out, os.path.basename(xlsx))
    return p if os.path.exists(p) else None
