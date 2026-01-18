# file_manager.py
from pathlib import Path
from collections import defaultdict
import re
import shutil
from typing import Optional

OUTPUT_SUFFIXES = [
    "_Green_highlighted.xlsx",
    "_SigTestTable_Green_highlighted.xlsx",
    "_SigTestTable_highlighted.xlsx",
    "_highlighted.xlsx",
    "_Green_highlighted_cells.csv",
    "_highlighted_cells.csv",
    "_highlighted.csv",
    "_SigTestTable.xlsx",
]

def save_uploaded_file(uploaded_file, target_dir: Path, filename: Optional[str] = None) -> Path:
    """
    Save a Streamlit uploaded file to target_dir.

    - uploaded_file: object returned by st.file_uploader
    - target_dir: Path to destination folder (created if missing)
    - filename: desired filename (if None, uses the uploaded file's original name)

    Returns absolute Path to saved file.
    """
    target_dir = Path(target_dir)
    target_dir.mkdir(parents=True, exist_ok=True)

    # Determine target name: explicit filename wins, else preserve original extension
    if filename:
        out_name = filename
    else:
        # uploaded_file usually has `.name` attr (e.g., "myfile.xlsx")
        orig = getattr(uploaded_file, "name", None) or "uploaded_file.xlsx"
        out_name = orig

    out_path = target_dir / out_name

    # Ensure we start from the beginning of the file-like object
    try:
        uploaded_file.seek(0)
    except Exception:
        # Some objects may not support seek; ignore if so
        pass

    # Try the efficient path first (getbuffer), fallback to read() or copyfileobj
    try:
        if hasattr(uploaded_file, "getbuffer"):
            data = uploaded_file.getbuffer()
            with open(out_path, "wb") as fh:
                fh.write(data)
        else:
            # try read() — many Streamlit UploadedFile objects support this
            content = uploaded_file.read()
            # If read() returns a string accidentally, encode
            if isinstance(content, str):
                content = content.encode()
            with open(out_path, "wb") as fh:
                fh.write(content)
    except Exception:
        # Last-resort safe copy using file-like streaming
        try:
            uploaded_file.seek(0)
        except Exception:
            pass
        with open(out_path, "wb") as fh:
            shutil.copyfileobj(uploaded_file, fh)

    return out_path.resolve()

def list_output_files(output_dir: Path):
    """
    Return list(Path) of files in output_dir sorted by name.
    """
    p = Path(output_dir)
    if not p.exists():
        return []
    return sorted([f for f in p.iterdir() if f.is_file()])

def group_outputs_by_table(output_dir: Path):
    """
    Group files by base table title extracted from filename.
    E.g. "B1_Stance_Green_highlighted.xlsx" -> "B1_Stance"
    Returns dict: {table_title: [Path, ...], ...}
    """
    files = list_output_files(Path(output_dir))
    groups = defaultdict(list)
    for f in files:
        name = f.name
        base = name
        # remove any known suffix exactly
        for suf in OUTPUT_SUFFIXES:
            if base.endswith(suf):
                base = base[: -len(suf)]
                break
        # remove repeated trailing suffix chunks like _Green or _highlighted etc
        base = re.sub(r'(_(Green|highlighted|SigTestTable|cells|SigTestTable_Green))+$', '', base, flags=re.IGNORECASE)
        base = base.rstrip("_")
        groups[base].append(f)
    for k in groups:
        groups[k] = sorted(groups[k], key=lambda p: p.name)
    return dict(groups)
