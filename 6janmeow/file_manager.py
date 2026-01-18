# file_manager.py
from pathlib import Path
from collections import defaultdict
import re

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

def save_uploaded_file(uploaded_file, target_dir: Path, filename: str = None) -> Path:
    """
    Save a streamlit uploaded file to target_dir with filename (or original name).
    Returns the Path to the saved file.
    """
    target_dir.mkdir(parents=True, exist_ok=True)
    name = filename or uploaded_file.name
    out_path = target_dir / name
    with open(out_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return out_path

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
