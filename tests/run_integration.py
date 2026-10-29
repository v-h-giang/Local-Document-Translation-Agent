"""Real XLSX end-to-end regression suite, without Ollama/customer data."""
import argparse
import contextlib
import csv
import hashlib
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tests.integration_cases import CASES, SMOKE_IDS
from spreadsheet_backend import (_free_local_port, _locate_libreoffice,
                                 _select_uno_python, translate_spreadsheet)


class FakeTranslator:
    def __init__(self, fail_on=None):
        self.calls = 0
        self.batch_calls = 0
        self.fail_on = fail_on

    def _one(self, text):
        if self.fail_on and self.fail_on in text:
            raise RuntimeError("Injected integration-test translator failure")
        return "訳:" + text

    def __call__(self, prompt):
        self.calls += 1
        matches = re.findall(r"(?m)^\s*\[\[SEG_(\d{3})\]\]\s*(.+?)\s*$", prompt)
        if matches:
            self.batch_calls += 1
            return "\n".join(
                "[[SEG_{}]] {}".format(idx, json.dumps(self._one(json.loads(payload)),
                                                    ensure_ascii=False))
                for idx, payload in matches)
        return self._one(prompt)


def sha(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def worker(py, env, port, action, manifest, folder):
    result = subprocess.run(
        [py, "-u", str(Path(__file__).with_name("uno_fixture_worker.py")),
         str(port), action, str(manifest), str(folder)],
        env=env, capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=180)
    if result.returncode:
        raise RuntimeError(f"UNO fixture {action}: {result.stderr[-4000:]}")
    return json.loads(result.stdout)


@contextlib.contextmanager
def fixture_office(folder):
    soffice = _locate_libreoffice()
    python_exe, env = _select_uno_python(soffice.parent)
    port = _free_local_port()
    profile = (folder / "fixture_profile").resolve().as_uri()
    office = subprocess.Popen(
        [str(soffice), "-env:UserInstallation=" + profile,
         "--headless", "--nologo", "--nodefault", "--norestore",
         "--nofirststartwizard",
         "--accept=socket,host=127.0.0.1,port=" + str(port)
         + ";urp;StarOffice.ServiceManager"],
        env=env, stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    try:
        yield python_exe, env, port
    finally:
        if office.poll() is None:
            office.terminate()
            try:
                office.wait(timeout=15)
            except subprocess.TimeoutExpired:
                office.kill()
                office.wait(timeout=10)


def review_csv(output, suffix):
    file = output.with_name(output.stem + suffix)
    if not file.is_file():
        raise AssertionError(f"Missing review CSV: {file}")
    with file.open("r", newline="", encoding="utf-8-sig") as handle:
        return list(csv.DictReader(handle))


def check_results(case, workbook, original_hash, input_file, output_file, translator):
    if "error" in workbook:
        raise AssertionError(workbook["error"])
    cells = workbook["cells"]
    shapes = workbook["shapes"]
    issues = []

    def require(ok, message):
        if not ok:
            issues.append(message)

    for key in case.get("translated", []):
        require("訳:" in cells.get(key, {}).get("text", ""), f"not translated: {key}")
    for key in case.get("unchanged", []):
        name, addr = key.rsplit("!", 1)
        source = next(x for x in case["sheets"] if x["name"] == name)["cells"][addr]
        target = "".join(r["text"] for r in source["runs"]) if "runs" in source else source["text"]
        require(cells.get(key, {}).get("text") == target, f"unchanged cell modified: {key}")
    for key, literals in case.get("literal", {}).items():
        value = cells.get(key, {}).get("text", "")
        for literal in literals:
            require(literal in value, f"literal corrupted: {key}: {literal}")
        require("訳:" in value, f"surrounding prose not translated: {key}")
    for key in case.get("contains_newline", []):
        require("\n" in cells.get(key, {}).get("text", ""), f"newline lost: {key}")
    for key, expected in case.get("formulas", {}).items():
        require(cells.get(key, {}).get("formula") == expected, f"formula modified: {key}")
    for key, expected in case.get("numbers", {}).items():
        require(cells.get(key, {}).get("number") == expected, f"number modified: {key}")
    for name, ranges in case.get("merged", {}).items():
        for r in ranges:
            require(workbook.get("merged", {}).get(name, {}).get(r), f"merge lost: {name}!{r}")
    for key, styles in case.get("check_rich", {}).items():
        runs = cells.get(key, {}).get("runs", [])
        for style in styles:
            require(any(x.get("style", {}).get(style) for x in runs),
                    f"lost rich style {style}: {key}")
    for key in case.get("shape_translated", []):
        require("訳:" in shapes.get(key, {}).get("text", ""), f"shape not translated: {key}")
    for key, styles in case.get("shape_rich", {}).items():
        runs = shapes.get(key, {}).get("runs", [])
        for style in styles:
            require(any(x.get("style", {}).get(style) for x in runs),
                    f"lost shape style {style}: {key}")
    if case.get("require_batch"):
        require(translator.batch_calls > 0, "contextual batching not used")
    if case.get("no_model_calls"):
        require(translator.calls == 0, "Japanese-only text sent to model")
    shape_review = review_csv(output_file, ".shape_review.csv")
    rich_review = review_csv(output_file, ".rich_text_review.csv")
    if case.get("review_shape"):
        require(bool(shape_review), "shape shrink/resize was not logged")
    if case.get("translated") or case.get("shape_translated"):
        require(not rich_review, "source left untranslated: " +
                str([x.get("object_id") for x in rich_review]))
    require(sha(input_file) == original_hash, "input source workbook changed")
    if issues:
        raise AssertionError("; ".join(issues))
    return {"shape_review": len(shape_review), "rich_review": len(rich_review)}


def translate(case, source, output, backend, failure=None):
    fake = FakeTranslator(failure)
    logs = io.StringIO()
    previous = os.environ.get("SPREADSHEET_BACKEND")
    os.environ["SPREADSHEET_BACKEND"] = backend
    try:
        with contextlib.redirect_stdout(logs):
            translate_spreadsheet(
                source, output, fake, model="integration-fake-v1",
                glossary={}, source_lang=case.get("source", "vi"),
                target_lang=case.get("target", "ja"))
    finally:
        if previous is None:
            os.environ.pop("SPREADSHEET_BACKEND", None)
        else:
            os.environ["SPREADSHEET_BACKEND"] = previous
    return fake, logs.getvalue()


def run():
    ap = argparse.ArgumentParser(description=__doc__)
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--smoke", action="store_true", help="Five critical cases")
    mode.add_argument("--full", action="store_true", help="All twenty cases (default)")
    ap.add_argument("--backend", choices=("libreoffice", "excel"),
                    default="libreoffice")
    ap.add_argument("--keep", action="store_true", help="Keep all fixture XLSX")
    args = ap.parse_args()
    cases = [x for x in CASES if not args.smoke or x["id"] in SMOKE_IDS]
    root = Path(__file__).resolve().parent
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    artifacts = root / "artifacts"
    artifacts.mkdir(exist_ok=True)
    work = Path(tempfile.mkdtemp(prefix=stamp + "_", dir=artifacts))
    inp, out = work / "input", work / "output"
    inp.mkdir(); out.mkdir()
    manifest = work / "manifest.json"
    manifest.write_text(json.dumps(cases, ensure_ascii=False), encoding="utf-8")
    reportdir = root / "reports"
    reportdir.mkdir(exist_ok=True)
    records, translated = [], {}
    start = time.monotonic()
    code = 1
    try:
        with fixture_office(work) as (python_exe, env, port):
            print(f"Creating {len(cases)} real XLSX workbooks...", flush=True)
            created = worker(python_exe, env, port, "create", manifest, inp)
            if created.get("created") != len(cases):
                raise AssertionError("Fixture creation count mismatch")
            for idx, case in enumerate(cases, 1):
                name = case["id"]
                source, output = inp / (name + ".xlsx"), out / (name + ".xlsx")
                digest = sha(source)
                rec = {"id": name, "status": "FAIL"}
                print(f"[{idx}/{len(cases)}] {name}", flush=True)
                try:
                    if case.get("checkpoint"):
                        try:
                            translate(case, source, output, args.backend, failure="xác nhận")
                            raise AssertionError("Expected injected failure")
                        except RuntimeError as exc:
                            if "Injected integration-test" not in str(exc):
                                raise
                        checkpoint = output.with_name(output.stem + ".checkpoint.json")
                        if not checkpoint.exists() or output.exists():
                            raise AssertionError("Checkpoint not retained on failure")
                    fake, log = translate(case, source, output, args.backend)
                    (work / (name + ".log")).write_text(log, encoding="utf-8")
                    if case.get("checkpoint") and "Resuming checkpoint:" not in log:
                        raise AssertionError("Did not resume from checkpoint")
                    translated[name] = (fake, digest)
                    rec.update({"model_calls": fake.calls, "batch_calls": fake.batch_calls})
                except Exception as exc:
                    rec["error"] = f"{type(exc).__name__}: {exc}"
                records.append(rec)
            inspected = worker(python_exe, env, port, "inspect", manifest, out)
            for case, rec in zip(cases, records):
                if rec.get("error"):
                    continue
                name = case["id"]
                try:
                    fake, digest = translated[name]
                    rec.update(check_results(case, inspected[name], digest,
                                  inp / (name + ".xlsx"),
                                  out / (name + ".xlsx"), fake))
                    rec["status"] = "PASS"
                except Exception as exc:
                    rec["error"] = f"{type(exc).__name__}: {exc}"
        code = int(any(x["status"] != "PASS" for x in records))
    except Exception as exc:
        records.append({"id": "suite", "status": "FAIL",
                        "error": f"{type(exc).__name__}: {exc}"})
        code = 2
    finally:
        report = {
            "timestamp_utc": datetime.now(timezone.utc).isoformat(),
            "backend": args.backend,
            "duration_seconds": round(time.monotonic() - start, 2),
            "passed": sum(x["status"] == "PASS" for x in records),
            "failed": sum(x["status"] != "PASS" for x in records),
            "cases": records,
        }
        file = reportdir / (stamp + "_integration.json")
        file.write_text(json.dumps(report, ensure_ascii=False, indent=2),
                        encoding="utf-8")
        for row in records:
            print(row["status"], row["id"], row.get("error", ""))
        print(f"PASS={report['passed']} FAIL={report['failed']}")
        print(f"Report: {file}")
        if code or args.keep:
            print(f"Retained failed-test artifacts: {work}")
        else:
            try:
                shutil.rmtree(work)
            except OSError as exc:
                print(f"WARNING: fixture cleanup deferred: {exc}")
    return code


if __name__ == "__main__":
    sys.exit(run())
