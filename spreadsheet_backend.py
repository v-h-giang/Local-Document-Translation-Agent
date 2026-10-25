"""Auto-select Microsoft Excel COM or LibreOffice UNO for spreadsheet files.

UNO runs in a separate LibreOffice-compatible Python interpreter. This module
does not copy or vendor LibreOffice DLLs; it reuses the native runtime already
installed with LibreOffice, avoiding incompatible CPython/pyuno builds.
"""
from __future__ import annotations

import json
import hashlib
import re
import os
import shutil
import socket
import subprocess
import sys
import tempfile
import time
from contextlib import contextmanager
from pathlib import Path

from excel_translate import EXCEL_EXTENSIONS
from excel_optimized import translate_excel_optimized
from translation_batch import translate_context_batch, literal_terms_from_glossary, BATCH_ALGORITHM_VERSION
from shape_review import write_shape_review
from rich_text_review import write_rich_text_review


BATCH_MAX_CHARS = 1100
CHECKPOINT_INTERVAL = 75


def _translate_context_batch(texts, translate_func):
    """Translate multiple strings in a contextual numbered block.

    Explicit IDs enable 1:1 validation. A model that loses IDs is handled
    through individual translations, never by guessing the mapping.
    """
    if not texts:
        return []
    results = []
    current = []
    current_size = 0

    def translate_group(group):
        if len(group) == 1:
            return [translate_func(group[0])]
        source = "\n".join(
            f"[[SEG_{index:03d}]] {value}"
            for index, value in enumerate(group)
        )
        prompt = (
            "Translate each numbered Vietnamese segment to Japanese. "
            "Return one line per ID. Keep IDs identical, do not merge or "
            "drop segments. No explanations.\n" + source
        )
        try:
            output = translate_func(prompt)
            hits = re.findall(
                r"(?m)^\s*\[\[SEG_(\d{3})\]\]\s*(.*)$", output
            )
            values = {}
            for index, translated in hits:
                number = int(index)
                if number in values or number >= len(group):
                    raise ValueError("Duplicate or invalid translation ID")
                values[number] = translated
            if len(values) != len(group):
                raise ValueError("Incomplete contextual translation")
            return [values[i] for i in range(len(group))]
        except Exception as exc:
            print(
                "WARNING: contextual batch returned invalid mapping; "
                f"falling back to individual strings: {exc}"
            )
            return [translate_func(value) for value in group]

    for value in texts:
        # Preserve the context budget and avoid one model request per cell.
        estimated = len(value) + 24
        if current and (current_size + estimated > BATCH_MAX_CHARS
                        or len(current) >= 16):
            results.extend(translate_group(current))
            current = []
            current_size = 0
        current.append(value)
        current_size += estimated
    if current:
        results.extend(translate_group(current))
    return results


def _excel_available():
    """Require a real EXCEL.EXE installation, not a stale COM registration."""
    if os.name != "nt":
        return False
    try:
        import win32com.client
        application = win32com.client.DispatchEx("Excel.Application")
        try:
            _ = application.Version
            program_path = Path(str(application.Path))
            executable = program_path / "EXCEL.EXE"
            if not executable.is_file():
                print(
                    "Excel COM is registered but Excel desktop executable "
                    f"is absent: {executable}. Using LibreOffice instead."
                )
                return False
        finally:
            application.Quit()
        return True
    except Exception as exc:
        print(f"Excel desktop unavailable ({type(exc).__name__}: {exc})")
        return False


def _locate_libreoffice():
    configured = os.environ.get("LIBREOFFICE_HOME", "").strip().strip('"')
    candidates = []

    if configured:
        home = Path(configured)
        candidates += [home / "program" / "soffice.exe", home / "soffice.exe",
                       home / "program" / "soffice", home / "soffice"]

    for variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        root = os.environ.get(variable)
        if root:
            candidates += [
                Path(root) / "LibreOffice" / "program" / "soffice.exe",
                Path(root) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
            ]

    for executable in ("soffice.exe", "soffice", "libreoffice"):
        found = shutil.which(executable)
        if found:
            candidates.append(Path(found))

    for candidate in candidates:
        if candidate.is_file():
            return candidate.resolve()

    raise RuntimeError(
        "LibreOffice was not found. Install LibreOffice or set "
        "LIBREOFFICE_HOME=C:\\Program Files\\LibreOffice."
    )


def _select_uno_python(program_dir):
    env = os.environ.copy()
    env["PATH"] = str(program_dir) + os.pathsep + env.get("PATH", "")
    env["PYTHONPATH"] = str(program_dir) + os.pathsep + env.get("PYTHONPATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    candidates = [
        program_dir / "python.exe",
        program_dir / "python.bin",
        program_dir / "python",
        Path(sys.executable),
    ]
    # On Linux a system Python can have python3-uno while a virtualenv cannot.
    if os.name != "nt":
        candidates.append(Path("/usr/bin/python3"))

    for interpreter in candidates:
        if not interpreter.is_file():
            continue
        try:
            probe = subprocess.run(
                [str(interpreter), "-c", "import uno"],
                env=env, capture_output=True, text=True, timeout=10,
            )
            if probe.returncode == 0:
                return str(interpreter), env
        except (OSError, subprocess.TimeoutExpired):
            continue

    raise RuntimeError(
        "LibreOffice exists but no compatible Python-UNO interpreter was "
        "found. The agent loads pyuno and its DLL dependencies dynamically "
        "from LibreOffice's program directory. Ensure the LibreOffice Python "
        "component is installed, or use a Python build compatible with pyuno. "
        f"LibreOffice directory: {program_dir}"
    )


def _free_local_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


@contextmanager
def _uno_workdir():
    """Clean scratch data after the UNO processes exit.

    Windows may briefly lock LibreOffice's profile/extension files even after
    the worker has exited. Cleanup failure is not translation failure, since
    the workbook and review reports have already been published.
    """
    temporary = Path(tempfile.mkdtemp(prefix="jpvn_uno_"))
    try:
        yield str(temporary)
    finally:
        last_error = None
        for attempt in range(8):
            try:
                shutil.rmtree(temporary)
                last_error = None
                break
            except FileNotFoundError:
                last_error = None
                break
            except OSError as exc:
                last_error = exc
                if attempt < 7:
                    time.sleep(0.5)
        if last_error is not None:
            print(
                "WARNING: LibreOffice output was saved, but a temporary "
                "profile file is still in use. Cleanup was deferred: "
                f"{temporary} ({type(last_error).__name__}: {last_error})"
            )
            print(
                "  Close any remaining LibreOffice processes associated "
                "with this job, then remove only that temporary directory."
            )


def _libreoffice_translate(input_path, output_path, translate_func, max_chars, signature, glossary):
    input_path = Path(input_path).resolve()
    output_path = Path(output_path).resolve()

    if input_path.suffix.lower() in {".xlsm", ".xlsb"}:
        raise RuntimeError(
            "LibreOffice backend cannot safely round-trip .xlsm macros or "
            ".xlsb files. Convert a COPY to .xlsx with LibreOffice first; "
            "the original file will not be modified."
        )

    soffice = _locate_libreoffice()
    python_exe, environment = _select_uno_python(soffice.parent)
    worker_file = Path(__file__).resolve().with_name("libreoffice_uno_worker.py")
    if not worker_file.is_file():
        raise RuntimeError(f"UNO worker missing: {worker_file}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    port = _free_local_port()

    print(f"LibreOffice runtime : {soffice}")
    print(f"UNO Python          : {python_exe}")
    print(f"UNO input           : {input_path}")
    print(f"UNO output          : {output_path}")
    print("WARNING: LibreOffice may reserialize XLSX drawings, charts, macros "
          "and other advanced formatting. Inspect the output before use.")

    office = None
    worker = None
    with _uno_workdir() as temporary:
        profile_url = (Path(temporary) / "profile").resolve().as_uri()
        staging_path = Path(temporary) / ("source" + input_path.suffix.lower())
        shutil.copy2(input_path, staging_path)
        print(f"UNO editable copy   : {staging_path}")
        office_args = [
            str(soffice), "-env:UserInstallation=" + profile_url,
            "--headless", "--nologo", "--nodefault", "--norestore", "--nofirststartwizard",
            "--accept=socket,host=127.0.0.1,port="
            + str(port) + ";urp;StarOffice.ServiceManager",
        ]

        try:
            office = subprocess.Popen(
                office_args, env=environment, stdin=subprocess.DEVNULL,
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            worker = subprocess.Popen(
                [python_exe, "-u", str(worker_file), str(port),
                 str(staging_path), str(output_path), str(input_path),
                 str(CHECKPOINT_INTERVAL), signature],
                env=environment, text=True, encoding="utf-8", errors="replace",
                stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, bufsize=1,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )

            finished = False
            for line in worker.stdout:
                line = line.strip()
                if not line:
                    continue
                try:
                    message = json.loads(line)
                except json.JSONDecodeError:
                    print("  UNO worker: " + line[:500])
                    continue

                kind = message.get("type")
                if kind == "translate_batch":
                    try:
                        translations = translate_context_batch(
                            message["texts"], translate_func,
                            literal_terms_from_glossary(glossary)
                        )
                        reply = {"ok": True, "translations": translations}
                    except Exception as exc:
                        reply = {"ok": False, "error": f"{type(exc).__name__}: {exc}"}
                    worker.stdin.write(json.dumps(reply, ensure_ascii=True) + "\n")
                    worker.stdin.flush()
                elif kind == "log":
                    print("  UNO: " + str(message.get("text", "")))
                elif kind == "error":
                    office_code = office.poll() if office is not None else None
                    raise RuntimeError(
                        "LibreOffice UNO failed (soffice exit code="
                        f"{office_code}):\n" + str(message["text"])
                    )
                elif kind == "done":
                    report = write_shape_review(output_path, message.get("review", []))
                    print("Shape review report : " + str(report))
                    skipped = message.get("rich_text_warnings", [])
                    rich_report = write_rich_text_review(output_path, skipped)
                    print("Rich-text review    : " + str(rich_report))
                    if skipped:
                        print(f"WARNING: {len(skipped)} cells/shapes kept untranslated "
                              "to protect formatting; see rich-text review CSV.")
                    print("LibreOffice translation completed: " + str(message.get("stats", {})))
                    finished = True
                else:
                    print("  UNO: " + line[:500])

            worker.wait(timeout=20)
            if worker.returncode != 0 or not finished:
                raise RuntimeError(
                    f"LibreOffice worker exited with code {worker.returncode}; "
                    f"finished={finished}. Check UNO messages above."
                )
        finally:
            if worker is not None and worker.poll() is None:
                worker.kill()
                worker.wait(timeout=10)
            if office is not None and office.poll() is None:
                office.terminate()
                try:
                    office.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    office.kill()
                    office.wait(timeout=10)


def translate_spreadsheet(input_path, output_path, translate_func, max_chars=1200,
                          model="", glossary=None):
    """Select backend automatically; allow a manual override for diagnostics."""
    if Path(input_path).suffix.lower() not in EXCEL_EXTENSIONS:
        raise ValueError(f"Unsupported spreadsheet: {input_path}")

    glossary = glossary or {}
    signature = hashlib.sha256(json.dumps({
        "model": model, "glossary": glossary,
        "batch_algorithm": BATCH_ALGORITHM_VERSION, "format_revision": 2,
    }, sort_keys=True, ensure_ascii=False).encode("utf-8")).hexdigest()
    choice = os.environ.get("SPREADSHEET_BACKEND", "auto").strip().lower()
    if choice not in {"auto", "excel", "libreoffice"}:
        raise ValueError(
            "SPREADSHEET_BACKEND must be auto, excel, or libreoffice"
        )

    if choice == "excel" or (choice == "auto" and _excel_available()):
        try:
            print("Spreadsheet backend : Microsoft Excel COM")
            return translate_excel_optimized(
                input_path, output_path, translate_func, max_chars=max_chars,
                signature=signature, glossary=glossary
            )
        except Exception as exc:
            if choice == "excel":
                raise
            print(
                "WARNING: Excel COM failed, attempting LibreOffice UNO: "
                f"{type(exc).__name__}: {exc}"
            )

    print("Spreadsheet backend : LibreOffice UNO")
    return _libreoffice_translate(input_path, output_path, translate_func,
                                  max_chars, signature, glossary)
