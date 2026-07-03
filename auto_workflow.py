from __future__ import annotations

import argparse
import csv
import hashlib
import shutil
from datetime import datetime
from pathlib import Path

from db import connect
from importers import import_excel
from orderboek_converter import convert_orderboek


PROJECT_ROOT = Path(__file__).resolve().parent
AUTO_IMPORT_DIR = PROJECT_ROOT / "data" / "auto_import"
INBOX_DIR = AUTO_IMPORT_DIR / "inbox"
PROCESSED_DIR = AUTO_IMPORT_DIR / "processed"
ERROR_DIR = AUTO_IMPORT_DIR / "error"
WORK_DIR = AUTO_IMPORT_DIR / "work"
LOG_PATH = PROJECT_ROOT / "logs" / "auto_workflow.log"
IMPORT_JOBS_LOG_PATH = PROJECT_ROOT / "logs" / "import_jobs.log"


def _now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _safe_name(path: Path) -> str:
    name = path.name.replace("\\", "_").replace("/", "_")
    return "".join(ch if ch.isalnum() or ch in " ._-()" else "_" for ch in name).strip() or "bestand.xlsx"


def _ensure_dirs() -> None:
    for directory in [INBOX_DIR, PROCESSED_DIR, ERROR_DIR, WORK_DIR, LOG_PATH.parent]:
        directory.mkdir(parents=True, exist_ok=True)


def _log(path: Path, *parts) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("a", encoding="utf-8", newline="") as f:
            writer = csv.writer(f)
            writer.writerow([_now(), *parts])
    except Exception as e:
        print("Auto workflow log mislukt:", repr(e))


def log_auto_workflow(*parts) -> None:
    _log(LOG_PATH, *parts)


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def ensure_auto_workflow_schema(conn) -> None:
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS imported_files (
            id SERIAL PRIMARY KEY,
            filename TEXT,
            file_hash TEXT UNIQUE,
            batch_id TEXT,
            imported_at TEXT
        )
        """
    )
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS auto_import_jobs (
            id SERIAL PRIMARY KEY,
            filename TEXT,
            batch_id TEXT,
            status TEXT,
            imported_count INTEGER,
            error_message TEXT,
            created_at TEXT,
            completed_at TEXT
        )
        """
    )
    conn.commit()


def _create_job(conn, filename: str) -> int:
    row = conn.execute(
        """
        INSERT INTO auto_import_jobs (filename, status, imported_count, created_at)
        VALUES (%s, 'gestart', 0, %s)
        RETURNING id
        """,
        (filename, _now()),
    ).fetchone()
    conn.commit()
    return int(row["id"])


def _complete_job(conn, job_id: int, status: str, imported_count: int, batch_id: str = "", error_message: str = "") -> None:
    conn.execute(
        """
        UPDATE auto_import_jobs
        SET status = %s,
            imported_count = %s,
            batch_id = %s,
            error_message = %s,
            completed_at = %s
        WHERE id = %s
        """,
        (status, imported_count, batch_id, error_message, _now(), job_id),
    )
    conn.commit()
    _log(IMPORT_JOBS_LOG_PATH, job_id, status, batch_id, imported_count, error_message)


def _hash_already_imported(conn, file_hash: str) -> bool:
    row = conn.execute(
        "SELECT id FROM imported_files WHERE file_hash = %s LIMIT 1",
        (file_hash,),
    ).fetchone()
    return row is not None


def _record_imported_file(conn, filename: str, file_hash: str, batch_id: str) -> None:
    conn.execute(
        """
        INSERT INTO imported_files (filename, file_hash, batch_id, imported_at)
        VALUES (%s, %s, %s, %s)
        ON CONFLICT (file_hash) DO NOTHING
        """,
        (filename, file_hash, batch_id, _now()),
    )
    conn.commit()


def _move_unique(src: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    base_name = _safe_name(src)
    target = target_dir / base_name
    if not target.exists():
        shutil.move(str(src), str(target))
        return target

    stem = target.stem
    suffix = target.suffix
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    target = target_dir / f"{stem}_{timestamp}{suffix}"
    shutil.move(str(src), str(target))
    return target


def _import_with_converter_fallback(path: Path) -> tuple[int, str]:
    try:
        return import_excel(str(path), return_batch=True)
    except ValueError as e:
        if "Onbekend Excel-formaat" not in str(e) and "wordt niet herkend" not in str(e):
            raise

    converted = WORK_DIR / f"{path.stem}_geconverteerd.xlsx"
    convert_orderboek(path, converted, remove_duplicates=True)
    return import_excel(str(converted), return_batch=True)


def process_auto_import_folder() -> dict:
    _ensure_dirs()
    summary = {
        "found": 0,
        "imported_files": 0,
        "skipped_files": 0,
        "error_files": 0,
        "imported_rows": 0,
        "batches": [],
        "errors": [],
    }

    files = sorted(
        [p for p in INBOX_DIR.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"}],
        key=lambda p: p.stat().st_mtime,
    )
    summary["found"] = len(files)

    with connect() as conn:
        ensure_auto_workflow_schema(conn)

    for path in files:
        job_id = None
        filename = path.name
        try:
            file_hash = file_sha256(path)
            with connect() as conn:
                ensure_auto_workflow_schema(conn)
                job_id = _create_job(conn, filename)
                if _hash_already_imported(conn, file_hash):
                    _complete_job(conn, job_id, "overgeslagen", 0, "", "Bestandshash is al geimporteerd.")
                    summary["skipped_files"] += 1
                    moved = _move_unique(path, PROCESSED_DIR)
                    log_auto_workflow("", filename, "duplicate", "overgeslagen", f"verplaatst naar {moved}")
                    continue

            imported_count, batch_id = _import_with_converter_fallback(path)

            with connect() as conn:
                ensure_auto_workflow_schema(conn)
                _record_imported_file(conn, filename, file_hash, batch_id)
                if job_id is not None:
                    _complete_job(conn, job_id, "gelukt", imported_count, batch_id, "")

            summary["imported_files"] += 1
            summary["imported_rows"] += int(imported_count or 0)
            summary["batches"].append(batch_id)
            moved = _move_unique(path, PROCESSED_DIR)
            log_auto_workflow(batch_id, filename, "import", "gelukt", f"{imported_count} leads; verplaatst naar {moved}")
        except Exception as e:
            error_message = f"{type(e).__name__}: {e}"
            summary["error_files"] += 1
            summary["errors"].append({"filename": filename, "error": error_message})
            try:
                with connect() as conn:
                    ensure_auto_workflow_schema(conn)
                    if job_id is None:
                        job_id = _create_job(conn, filename)
                    _complete_job(conn, job_id, "fout", 0, "", error_message)
            except Exception as log_error:
                print("Auto import job fout loggen mislukt:", repr(log_error))
            try:
                moved = _move_unique(path, ERROR_DIR)
                log_auto_workflow("", filename, "import", "fout", f"{error_message}; verplaatst naar {moved}")
            except Exception:
                log_auto_workflow("", filename, "import", "fout", error_message)

    return summary


def get_auto_workflow_summary() -> dict:
    _ensure_dirs()
    pending_files = len([p for p in INBOX_DIR.iterdir() if p.is_file() and p.suffix.lower() in {".xls", ".xlsx"}])
    result = {
        "inbox_path": str(INBOX_DIR.relative_to(PROJECT_ROOT)).replace("\\", "/"),
        "pending_files": pending_files,
        "last_job": None,
        "last_imported_at": "",
        "last_imported_count": 0,
    }

    try:
        with connect() as conn:
            ensure_auto_workflow_schema(conn)
            result["last_job"] = conn.execute(
                """
                SELECT filename, batch_id, status, imported_count, error_message, created_at, completed_at
                FROM auto_import_jobs
                ORDER BY id DESC
                LIMIT 1
                """
            ).fetchone()
            imported = conn.execute(
                """
                SELECT imported_at, batch_id
                FROM imported_files
                ORDER BY id DESC
                LIMIT 1
                """
            ).fetchone()
            if imported:
                result["last_imported_at"] = imported["imported_at"] or ""
    except Exception as e:
        result["error"] = f"{type(e).__name__}: {e}"

    return result


def main() -> None:
    parser = argparse.ArgumentParser(description="Automatische lead-import workflow")
    parser.add_argument("--import-only", action="store_true", help="Verwerk Excelbestanden uit data/auto_import/inbox")
    args = parser.parse_args()

    if not args.import_only:
        parser.print_help()
        return

    summary = process_auto_import_folder()
    print(summary)


if __name__ == "__main__":
    main()
