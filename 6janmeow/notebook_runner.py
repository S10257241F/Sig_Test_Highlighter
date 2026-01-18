# notebook_runner.py
import importlib
import subprocess
import sys
import os
import re
from pathlib import Path


def run_with_papermill(nb_path: Path, input_file: Path, output_dir: Path, on_progress=None):
    """
    Runs notebook with papermill.
    If on_progress is provided, it will receive (progress_float_0_to_1, message_str).
    """
    import papermill as pm

    executed_nb = output_dir / f"executed_{nb_path.name}"
    params = {
        "input_file": str(input_file),
        "output_dir": str(output_dir),
    }

    # papermill logs progress like: "Executing:  42%|..."
    # We'll attach a logging handler to papermill's logger and parse %.
    if on_progress is not None:
        import logging

        class _PMProgressHandler(logging.Handler):
            def emit(self, record):
                try:
                    msg = record.getMessage()
                    m = re.search(r"Executing:\s*(\d+)%", msg)
                    if m:
                        pct = int(m.group(1))
                        on_progress(pct / 100.0, msg)
                except Exception:
                    # never let logging break execution
                    pass

        pm_logger = logging.getLogger("papermill")
        handler = _PMProgressHandler()
        handler.setLevel(logging.INFO)

        # attach temporarily
        pm_logger.addHandler(handler)
        try:
            pm.execute_notebook(
                str(nb_path),
                str(executed_nb),
                parameters=params,
                kernel_name="python3",
            )
        finally:
            pm_logger.removeHandler(handler)
    else:
        pm.execute_notebook(
            str(nb_path),
            str(executed_nb),
            parameters=params,
            kernel_name="python3",
        )

    # make sure we end at 100% if we were tracking
    if on_progress is not None:
        on_progress(1.0, "Papermill finished.")

    return f"Papermill wrote executed notebook: {executed_nb}"


def run_with_nbconvert(nb_path: Path, input_file: Path, output_dir: Path, on_progress=None):
    """
    Fallback: run nbconvert --execute. We expose the paths via environment variables
    so the notebook can read them with os.environ if it doesn't use papermill parameters.

    nbconvert doesn't expose clean per-cell progress without changing execution engine,
    so we only send coarse progress updates (start/end).
    """
    env = os.environ.copy()
    env["UPLOADED_FILE"] = str(input_file)
    env["HIGHLIGHTED_OUTPUTS"] = str(output_dir)

    if on_progress is not None:
        on_progress(0.05, "Starting nbconvert execution...")

    cmd = [
        sys.executable, "-m", "nbconvert",
        "--to", "notebook",
        "--execute",
        "--inplace",
        str(nb_path),
    ]
    subprocess.check_call(cmd, env=env)

    if on_progress is not None:
        on_progress(1.0, "nbconvert finished.")

    return f"nbconvert executed notebook in place: {nb_path}"


def run_notebook(nb_path, input_file, output_dir, on_progress=None):
    """
    Attempt to run notebook with papermill first. If papermill is not installed,
    fall back to nbconvert and rely on environment variables.

    Added: on_progress callback (optional)
      on_progress(progress_float_0_to_1, message_str)

    Returns (success: bool, message: str)
    """
    nb_path = Path(nb_path)
    input_file = Path(input_file)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    try:
        # prefer papermill if present
        if importlib.util.find_spec("papermill") is not None:
            if on_progress is not None:
                on_progress(0.01, "Papermill detected. Starting execution...")
            msg = run_with_papermill(nb_path, input_file, output_dir, on_progress=on_progress)
        else:
            if on_progress is not None:
                on_progress(0.01, "Papermill not installed. Using nbconvert...")
            msg = run_with_nbconvert(nb_path, input_file, output_dir, on_progress=on_progress)

        return True, msg

    except subprocess.CalledProcessError as e:
        return False, f"Subprocess error: {e}\nOutput: {getattr(e, 'output', '')}"
    except Exception as ex:
        return False, f"Unexpected error: {ex}"
