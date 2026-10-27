#!/usr/bin/env python3

"""
STEP Garmin v3 token verification and seeding — full account sweep.

Every run checks each unique Garmin email in the STEP workbook:
  * Existing, structurally valid v3 tokens are authenticated with Garmin.
  * Missing, malformed, or authentication-rejected tokens are freshly seeded
    when the workbook contains a password.
  * New tokens are tested again by logging in with only the saved token file.
  * Fresh logins use a PRIVATE STAGING PATH, never the live token filename.
    The existing token is backed up and replaced ONLY after full verification.
  * Rate limits/network failures NEVER count as verified and NEVER cause the
    existing token to be deleted. They are marked unresolved in the reports.
  * Existing tokens with missing workbook passwords can still be verified;
    seeding those accounts requires adding a password to the workbook.

No Garmin wellness/activity measurements or REDCap APIs are accessed.
The original Excel workbook is not changed. No password/email is put in audit.

Token location: <workbook folder>/.garmin_tokens_v3/<email sha256[:24]>.json
Original compatibility: garminconnect >= 0.3.10, including v3 DI OAuth JSON.

Usage:
  python step_v3_verify_and_seed.py --file "/path/to/STEP.xlsx" \
    --sheet "Master Spreadsheet" --verify-existing

  # Read-only inspection; does not authenticate with Garmin:
  python step_v3_verify_and_seed.py --file "/path/workbook.xlsx" --dry-run

  # Optional slower rate-limit-conscious pacing between accounts:
  python step_v3_verify_and_seed.py --file "/path/workbook.xlsx" \
    --inter-account-delay 1.0

All accounts are attempted; online verification depends on Garmin responding.
Exit 0 = all listed Garmin accounts verified; 2 = any account unresolved.
"""

from __future__ import annotations

import argparse

import csv

import hashlib

import importlib.metadata

import json

import logging

import os
import shutil
import tempfile
import uuid

import re

import sys

import time

from datetime import datetime, timezone

from pathlib import Path

from garminconnect import Garmin

from openpyxl import load_workbook

try:

    from garminconnect.exceptions import GarminConnectTooManyRequestsError

except Exception:

    class GarminConnectTooManyRequestsError(Exception):
        """Fallback marker used only when garminconnect exception import fails."""
        pass

try:

    import tkinter as tk

    from tkinter import filedialog

    _HAS_TK = True

except Exception:

    _HAS_TK = False

# =============================================================================

# CONFIGURATION

# =============================================================================

# This matches the modern auth generation used for the Rove no-wait seeder.

MIN_GARMINCONNECT_VERSION = (0, 3, 10)

TOKEN_DIR_NAME = ".garmin_tokens_v3"

LATEST_SUMMARY_FILENAME = "STEP_Garmin_Token_Seed_Summary.json"

EMAIL_HEADERS = (

    "Watch Email",

    "watch email",

    "account email",

    "Garmin Email",

    "email",

    "username",

    "user name",

)

PASSWORD_HEADERS = (

    "Watch Password",

    "watch password",

    "account password",

    "Garmin Password",

    "password",

)

RECORD_ID_HEADERS = (

    "REDCap ID",

    "redcap id",

    "record id",

    "record_id",

)

# =============================================================================

# BASIC HELPERS

# =============================================================================

def utc_now() -> datetime:

    return datetime.now(timezone.utc)

def iso_utc(dt: datetime) -> str:

    return (

        dt.astimezone(timezone.utc)

        .replace(microsecond=0)

        .isoformat()

        .replace("+00:00", "Z")

    )

def version_tuple(version_text: str):

    nums = re.findall(r"\d+", str(version_text or ""))

    vals = [int(x) for x in nums[:3]]

    while len(vals) < 3:

        vals.append(0)

    return tuple(vals)

def get_garminconnect_version():

    for package_name in ("garminconnect", "python-garminconnect"):

        try:

            return importlib.metadata.version(package_name)

        except Exception:

            pass

    return None

def ensure_supported_garminconnect() -> str:

    version = get_garminconnect_version()

    if version is None or version_tuple(version) < MIN_GARMINCONNECT_VERSION:

        wanted = ".".join(str(x) for x in MIN_GARMINCONNECT_VERSION)

        raise RuntimeError(

            f"This utility requires garminconnect >= {wanted}. "

            f"Installed version: {version or 'not found'}.\n"

            "Upgrade with:\n"

            "  python -m pip install --upgrade garminconnect"

        )

    return version

def canon(value) -> str:

    if value is None:

        return ""

    text = str(value).replace("\xa0", " ").lower().strip()

    text = re.sub(r"[^a-z0-9]+", " ", text)

    return re.sub(r"\s+", " ", text).strip()

def cell_text(cell, preserve_spaces: bool = False) -> str:

    """Read a cell while preserving zero-padded numeric display when possible."""

    value = cell.value

    if value is None or cell.data_type == "e":

        return ""

    if isinstance(value, float) and value.is_integer():

        value = int(value)

    if (

        isinstance(value, int)

        and not isinstance(value, bool)

        and re.fullmatch(r"0+", cell.number_format or "")

    ):

        text = str(value).zfill(len(cell.number_format))

    else:

        text = str(value)

    return text if preserve_spaces else text.strip()

def account_hash(email: str) -> str:

    return hashlib.sha256(

        str(email or "").strip().lower().encode("utf-8")

    ).hexdigest()[:24]

def mask_email(email: str) -> str:

    text = str(email or "").strip()

    if "@" not in text:

        if len(text) <= 2:

            return "***"

        return f"{text[0]}***{text[-1]}"

    local, domain = text.split("@", 1)

    return f"{(local[:1] or '*')}***@{domain}"

def sanitize_error(exc, email: str = "", password: str = "") -> str:

    text = str(exc)

    if email:

        text = text.replace(str(email), "<email>")

    if password:

        text = text.replace(str(password), "<password>")

    return text[:1500]

def is_rate_limit_error(exc) -> bool:

    if isinstance(exc, GarminConnectTooManyRequestsError):

        return True

    text = str(exc).lower()

    return any(

        marker in text

        for marker in (

            "429",

            "too many requests",

            "rate limit",

            "too many login attempts",

            "all login strategies rate limited",

        )

    )

# =============================================================================

# WORKBOOK / SHEET SELECTION

# =============================================================================

def choose_source_path(cli_path: str | None) -> Path:

    if cli_path:

        path = Path(cli_path).expanduser().resolve()

        if path.is_file():

            return path

        print(f"[warn] --file not found: {path}")

    if _HAS_TK:

        try:

            root = tk.Tk()

            root.withdraw()

            path = filedialog.askopenfilename(

                title="Select STEP account workbook",

                filetypes=[

                    ("Excel workbooks", "*.xlsx *.xlsm *.xltx *.xltm"),

                    ("All files", "*.*"),

                ],

            )

            root.destroy()

            if path:

                return Path(path).expanduser().resolve()

        except Exception:

            pass

    while True:

        raw = input("STEP workbook path (Enter to cancel): ").strip().strip('"')

        if not raw:

            raise ValueError("No STEP workbook selected.")

        path = Path(raw).expanduser().resolve()

        if path.is_file():

            return path

        print("Path not found. Try again.\n")

def build_header_map(ws, header_row: int) -> dict[str, int]:

    result: dict[str, int] = {}

    for col in range(1, ws.max_column + 1):

        key = canon(ws.cell(header_row, col).value)

        if key:

            result.setdefault(key, col)

    return result

def resolve_header(header_map: dict[str, int], candidates, required: bool = True):

    for candidate in candidates:

        key = canon(candidate)

        if key in header_map:

            return header_map[key]

    for candidate in candidates:

        key = canon(candidate)

        for existing, col in header_map.items():

            if key and (key == existing or key in existing):

                return col

    if required:

        raise KeyError(f"Could not find any of these workbook columns: {tuple(candidates)}")

    return None

def find_header_row(ws, max_scan_rows: int = 20):

    email_keys = {canon(x) for x in EMAIL_HEADERS}

    password_keys = {canon(x) for x in PASSWORD_HEADERS}

    best_row = None

    best_hits = -1

    for row in range(1, min(ws.max_row, max_scan_rows) + 1):

        values = {

            canon(ws.cell(row, col).value)

            for col in range(1, ws.max_column + 1)

            if canon(ws.cell(row, col).value)

        }

        has_email = bool(values & email_keys)

        has_password = bool(values & password_keys)

        has_record = bool(values & {canon(x) for x in RECORD_ID_HEADERS})

        hits = int(has_email) + int(has_password) + int(has_record)

        if hits > best_hits:

            best_hits = hits

            best_row = row

        if has_email and has_password:

            return row

    return best_row if best_hits >= 2 else None

def sheet_has_credentials(workbook: Path, sheet_name: str) -> bool:

    wb = load_workbook(workbook, read_only=True, data_only=False)

    try:

        ws = wb[sheet_name]

        header_row = find_header_row(ws)

        if header_row is None:

            return False

        header_map = build_header_map(ws, header_row)

        return (

            resolve_header(header_map, EMAIL_HEADERS, required=False) is not None

            and resolve_header(header_map, PASSWORD_HEADERS, required=False) is not None

        )

    finally:

        wb.close()

def choose_sheet_name(workbook: Path, cli_sheet: str | None) -> str:

    wb = load_workbook(workbook, read_only=True, data_only=False)

    try:

        names = list(wb.sheetnames)

    finally:

        wb.close()

    if cli_sheet:

        if cli_sheet not in names:

            raise KeyError(

                f"Worksheet '{cli_sheet}' was not found. Available sheets: {names}"

            )

        return cli_sheet

    # STEP's normal location wins when it actually contains credentials.

    if "Master Spreadsheet" in names:

        try:

            if sheet_has_credentials(workbook, "Master Spreadsheet"):

                return "Master Spreadsheet"

        except Exception:

            pass

    matching = []

    for name in names:

        try:

            if sheet_has_credentials(workbook, name):

                matching.append(name)

        except Exception:

            pass

    if len(matching) == 1:

        return matching[0]

    if len(matching) > 1:

        print("Multiple worksheets contain STEP Garmin credentials:")

        for i, name in enumerate(matching, 1):

            print(f"  {i}. {name}")

        while True:

            raw = input("Select worksheet number: ").strip()

            if raw.isdigit() and 1 <= int(raw) <= len(matching):

                return matching[int(raw) - 1]

            print("Invalid selection.\n")

    raise RuntimeError(

        "Could not find a worksheet containing both Watch Email and Watch Password."

    )

def load_unique_accounts(workbook: Path, sheet_name: str):

    """Return one in-memory record per unique normalized Garmin email."""

    wb = load_workbook(workbook, read_only=True, data_only=False)

    try:

        if sheet_name not in wb.sheetnames:

            raise KeyError(

                f"Worksheet '{sheet_name}' was not found. Available sheets: {wb.sheetnames}"

            )

        ws = wb[sheet_name]

        header_row = find_header_row(ws)

        if header_row is None:

            raise RuntimeError(

                f"Could not find Watch Email/Watch Password headers in '{sheet_name}'."

            )

        header_map = build_header_map(ws, header_row)

        email_col = resolve_header(header_map, EMAIL_HEADERS)

        password_col = resolve_header(header_map, PASSWORD_HEADERS)

        record_col = resolve_header(header_map, RECORD_ID_HEADERS, required=False)

        by_email: dict[str, dict] = {}

        rows_missing_password = 0

        rows_missing_email = 0

        for row in range(header_row + 1, ws.max_row + 1):

            email = cell_text(ws.cell(row, email_col))

            # Password is intentionally preserved exactly as stored in Excel.

            password = cell_text(ws.cell(row, password_col), preserve_spaces=True)

            record_id = (

                cell_text(ws.cell(row, record_col))

                if record_col is not None

                else ""

            )

            if not email:

                if password.strip() or record_id:

                    rows_missing_email += 1

                continue

            key = email.lower()

            if not password.strip():

                rows_missing_password += 1

                by_email.setdefault(

                    key,

                    {

                        "email": email,

                        "password": "",

                        "record_id": record_id,

                        "source_row": row,

                    },

                )

                continue

            # Last nonblank credential row wins for duplicate Garmin accounts.

            by_email[key] = {

                "email": email,

                "password": password,

                "record_id": record_id,

                "source_row": row,

            }

        usable = [

            account

            for account in by_email.values()

            if account["email"] and account["password"].strip()

        ]

        unusable = [

            account

            for account in by_email.values()

            if account["email"] and not account["password"].strip()

        ]

        return {

            "header_row": header_row,

            "accounts": sorted(by_email.values(), key=lambda a: a["source_row"]),
            "unique_with_credentials": len(usable),

            "unique_with_missing_password": len(unusable),

            "rows_with_missing_password": rows_missing_password,

            "rows_with_missing_email": rows_missing_email,

        }

    finally:

        wb.close()

# =============================================================================

# TOKEN STORAGE

# =============================================================================

def token_dir_for(workbook: Path, explicit_token_dir: str | None = None) -> Path:

    if explicit_token_dir:

        root = Path(explicit_token_dir).expanduser().resolve()

        # Accept either the actual .garmin_tokens_v3 directory or its parent.

        if root.name != TOKEN_DIR_NAME:

            root = root / TOKEN_DIR_NAME

    else:

        root = workbook.resolve().parent / TOKEN_DIR_NAME

    root.mkdir(parents=True, exist_ok=True)

    try:

        root.chmod(0o700)

    except Exception:

        pass

    return root

def token_path_for(token_root: Path, email: str) -> Path:

    return token_root / f"{account_hash(email)}.json"

def token_file_looks_usable(path: Path) -> bool:

    try:

        return path.is_file() and path.stat().st_size > 20

    except Exception:

        return False

def token_json_has_expected_fields(path: Path) -> bool:

    """Structural check only; does not contact Garmin."""

    try:

        data = json.loads(path.read_text(encoding="utf-8"))

        return bool(

            isinstance(data, dict)

            and data.get("di_token")

            and data.get("di_refresh_token")

            and data.get("di_client_id")

        )

    except Exception:

        return False

def is_connection_or_transient_error(exc) -> bool:
    """A connectivity/server error cannot establish that a saved token is bad."""
    message = f"{type(exc).__name__} {exc}".lower()
    markers = (
        "connection", "timeout", "timed out", "dns", "network", "ssl",
        "temporary", "unavailable", "server error", "503", "502", "504", "500",
        "cloudflare", "waf", "challenge", "failed to establish", "proxy",
    )
    return isinstance(exc, (OSError, TimeoutError, ConnectionError)) or any(
        marker in message for marker in markers
    )

def verify_existing_token(path: Path) -> None:
    """Authenticate a session against Garmin; structural checks alone don't count."""
    if not token_json_has_expected_fields(path):
        raise ValueError("token JSON is malformed or missing required DI OAuth fields")
    client = Garmin(retry_attempts=0, verify_login=True)
    client.login(str(path))
    # If the library refreshed/wrote a token, ensure it is still usable on disk.
    if not token_json_has_expected_fields(path):
        raise RuntimeError("Garmin authentication returned but persisted token is malformed")

def seed_one_account(email: str, password: str, stage_path: Path) -> None:
    """Use only a NEW file path to force a credential login, not cached login."""
    if stage_path.exists():
        raise RuntimeError("fresh-login staging path unexpectedly exists")
    client = Garmin(
        email=email,
        password=password,
        prompt_mfa=lambda: input("  Garmin MFA code: ").strip(),
        retry_attempts=0,
        verify_login=True,
    )
    # Don't remove/override the library's normal login strategies.
    inner_client = getattr(client, "client", None)
    skip_strategies = getattr(inner_client, "skip_strategies", None)
    if skip_strategies is not None and hasattr(skip_strategies, "clear"):
        skip_strategies.clear()

    client.login(str(stage_path))
    if not stage_path.exists():
        inner_client = getattr(client, "client", None)
        if inner_client is None or not hasattr(inner_client, "dump"):
            raise RuntimeError("Garmin login succeeded but no token file was saved")
        inner_client.dump(str(stage_path))
    if not token_json_has_expected_fields(stage_path):
        raise RuntimeError("Fresh login did not create a valid v3 DI OAuth JSON token")
    # Verify a second time FROM DISK, with no username/password supplied.
    verify_existing_token(stage_path)

def backup_and_promote(stage_path: Path, live_path: Path, token_root: Path) -> bool:
    """Back up any old token before atomically promoting a fully verified token.

    Backup copy is made only after the new token has passed token-only login.
    The live token is NEVER moved away before success is confirmed.
    """
    if not token_json_has_expected_fields(stage_path):
        raise ValueError("refusing to promote an unverified/malformed staged token")
    replaced_old = live_path.exists()
    if replaced_old:
        backup_dir = token_root / ".replaced_token_backups"
        backup_dir.mkdir(mode=0o700, parents=True, exist_ok=True)
        try:
            backup_dir.chmod(0o700)
        except OSError:
            pass
        stamp = utc_now().strftime("%Y%m%dT%H%M%S%fZ")
        backup = backup_dir / f"{live_path.stem}.{stamp}.{uuid.uuid4().hex[:8]}.bak"
        shutil.copy2(live_path, backup)
        try:
            backup.chmod(0o600)
        except OSError:
            pass
    # Stage lives inside token_root (same file system); os.replace is atomic.
    os.replace(stage_path, live_path)
    try:
        live_path.chmod(0o600)
    except OSError:
        pass
    return replaced_old

def token_state(path: Path) -> str:
    if path.is_symlink():
        return "symlink_rejected"
    if not path.exists():
        return "missing"
    if token_json_has_expected_fields(path):
        return "structurally_valid"
    return "malformed"

def write_outputs(workbook: Path, sheet: str, token_root: Path,
                  audit_rows: list[dict], package_version: str,
                  counts: dict, loaded: dict) -> tuple[Path, str]:
    stamp = utc_now().strftime("%Y%m%dT%H%M%SZ")
    audit_path = workbook.parent / f"STEP_Garmin_Token_Verify_Audit_{stamp}.csv"
    fields = (
        "account_hash", "workbook_row", "password_present", "initial_token_state",
        "action", "outcome", "verified_online", "token_present_after",
        "seconds", "error",
    )
    with audit_path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=fields)
        writer.writeheader()
        for row in audit_rows:
            writer.writerow({k: row.get(k, "") for k in fields})

    verified_count = sum(row["verified_online"] is True for row in audit_rows)
    unresolved = len(audit_rows) - verified_count
    status = "all_verified" if audit_rows and not unresolved else "partial_unverified"
    summary = {
        "timestamp_utc": iso_utc(utc_now()),
        "project": "SyncStep",
        "status": status,
        "workbook": str(workbook),
        "sheet": sheet,
        "token_directory": str(token_root),
        "token_format": "v3 / garminconnect 0.3.x DI OAuth JSON",
        "garminconnect_version": package_version,
        "unique_workbook_accounts_with_email": len(audit_rows),
        "unique_workbook_accounts_with_credentials": loaded["unique_with_credentials"],
        "unique_accounts_missing_password": loaded["unique_with_missing_password"],
        "rows_with_missing_email": loaded["rows_with_missing_email"],
        "rows_with_missing_password": loaded["rows_with_missing_password"],
        "verified_online_count": verified_count,
        "unresolved_count": unresolved,
        "verified_existing_count": counts["verified_existing"],
        "verified_new_count": counts["verified_new"],
        "verified_reseeded_count": counts["verified_reseeded"],
        "rate_limited_count": counts["rate_limited"],
        "transient_error_count": counts["transient"],
        "seed_failed_count": counts["seed_failed"],
        "missing_password_no_token_count": counts["missing_password_no_token"],
        "token_files_structurally_valid_after_run": sum(
            row["token_present_after"] is True for row in audit_rows
        ),
        "outcome_counts": dict(sorted({
            outcome: sum(row["outcome"] == outcome for row in audit_rows)
            for outcome in {row["outcome"] for row in audit_rows}
        }.items())),
        "audit_csv": str(audit_path),
        "zero_script_added_delay_by_default": True,
        "downloads_health_data": False,
        "contacts_redcap": False,
        "modifies_workbook": False,
        "contains_plaintext_email_or_password": False,
        "note": (
            "Only verified_online=true means authenticated successfully with Garmin. "
            "A structurally valid token alone is not proof that it can authenticate. "
            "429/network failures remain unresolved, not valid and not invalid."
        ),
    }
    path = workbook.parent / "STEP_Garmin_Token_Seed_Summary.json"
    path.write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return path, status

def main() -> int:
    parser = argparse.ArgumentParser(description=(
        "Verify every STEP v3 Garmin session, reseed failures and missing accounts, "
        "and report unresolved/rate-limited accounts without destroying saved tokens."
    ))
    parser.add_argument("--file", "-f", help="STEP workbook; picker if omitted")
    parser.add_argument("--sheet", "-s", help="Worksheet name; autodetected if omitted")
    parser.add_argument("--token-dir", help="Override .garmin_tokens_v3 directory or its parent")
    parser.add_argument("--verify-existing", action="store_true", default=True,
                        help="Verify existing sessions (always enabled in this edition)")
    parser.add_argument("--dry-run", action="store_true", help="Local token inventory only")
    parser.add_argument("--debug-login", action="store_true", help="Verbose library logging")
    parser.add_argument("--inter-account-delay", type=float, default=0.0,
                        help="Optional seconds between accounts (default: 0)")
    parser.add_argument("--stop-on-rate-limit", action="store_true",
                        help="Stop after first 429 rather than continuing (recommended if 429s recur)")
    args = parser.parse_args()
    if args.inter_account_delay < 0:
        parser.error("--inter-account-delay must be >= 0")
    if args.debug_login:
        logging.basicConfig(level=logging.DEBUG,
                            format="%(asctime)s %(levelname)s %(name)s: %(message)s")

    workbook = choose_source_path(args.file)
    sheet = choose_sheet_name(workbook, args.sheet)
    version = ensure_supported_garminconnect()
    token_root = token_dir_for(workbook, args.token_dir)
    loaded = load_unique_accounts(workbook, sheet)
    accounts = loaded["accounts"]
    counts = dict(verified_existing=0, verified_new=0, verified_reseeded=0,
                  rate_limited=0, transient=0, seed_failed=0,
                  missing_password_no_token=0)
    audit_rows: list[dict] = []
    token_states = [token_state(token_path_for(token_root, a["email"])) for a in accounts]
    print("=== STEP V3 GARMIN: VERIFY + SEED ALL UNIQUE ACCOUNTS ===")
    print(f"Workbook: {workbook}\nSheet: {sheet}\ngarminconnect: {version}")
    print(f"Token directory: {token_root}")
    print(f"Unique accounts with email: {len(accounts)}")
    print(f"Accounts with missing passwords: {loaded['unique_with_missing_password']}")
    print(f"Existing structural tokens: {token_states.count('structurally_valid')}")
    print(f"Missing: {token_states.count('missing')}; malformed: {token_states.count('malformed')}")
    print("Every existing token will be tested online; missing/rejected accounts will be seeded.")
    print("Only verified sessions count as success. Source workbook will not change.\n")
    if args.dry_run:
        print("DRY RUN: no Garmin authentication performed and no tokens created.")
        print("Account count in each token state:")
        for state in sorted(set(token_states)):
            print(f"  {state}: {token_states.count(state)}")
        return 0

    stop_after_429 = False
    for idx, account in enumerate(accounts, 1):
        email = account["email"]
        password = account["password"]
        live_path = token_path_for(token_root, email)
        initial = token_state(live_path)
        start = time.perf_counter()
        row = {
            "account_hash": account_hash(email),
            "workbook_row": account["source_row"],
            "password_present": bool(password.strip()),
            "initial_token_state": initial,
            "action": "", "outcome": "", "verified_online": False,
            "token_present_after": False, "seconds": 0.0, "error": "",
        }
        print(f"[{idx}/{len(accounts)}] {mask_email(email)} ({account_hash(email)[:12]}...): ", end="", flush=True)

        try:
            if stop_after_429:
                row["action"] = "not_attempted"
                row["outcome"] = "deferred_after_rate_limit"
                print("DEFERRED after earlier 429")
                continue

            if initial == "symlink_rejected":
                row["action"] = "none"
                row["outcome"] = "symlink_rejected"
                print("UNRESOLVED: token file is a symlink; no overwrite attempted")
                continue

            if initial == "structurally_valid":
                row["action"] = "verify_existing"
                # Retain original bytes in case a failed verification attempt has
                # modified the cache as a side effect of refresh.
                original_bytes = live_path.read_bytes()
                try:
                    verify_existing_token(live_path)
                    row["outcome"] = "verified_existing"
                    row["verified_online"] = True
                    counts["verified_existing"] += 1
                    print("VERIFIED existing token")
                    continue
                except Exception as exc:
                    # Rate-limit and network failure are NOT evidence of bad cache.
                    # Restore the original saved token only if verification changed it.
                    if not live_path.is_file() or live_path.read_bytes() != original_bytes:
                        tmp = live_path.with_name(f".{live_path.stem}.{uuid.uuid4().hex}.rollback")
                        tmp.write_bytes(original_bytes)
                        tmp.chmod(0o600)
                        os.replace(tmp, live_path)
                    row["error"] = sanitize_error(exc, email, password)
                    if is_rate_limit_error(exc):
                        counts["rate_limited"] += 1
                        row["outcome"] = "unverified_rate_limited_existing_kept"
                        print("UNVERIFIED (429); retained original token")
                        if args.stop_on_rate_limit:
                            stop_after_429 = True
                        continue
                    if is_connection_or_transient_error(exc):
                        counts["transient"] += 1
                        row["outcome"] = "unverified_connection_existing_kept"
                        print("UNVERIFIED (network/service); retained original token")
                        continue
                    print("REJECTED/UNUSABLE; attempting isolated credential re-login")

            if not password.strip():
                row["outcome"] = "unverified_missing_password"
                row["action"] = "verify_only_or_missing"
                if initial == "missing":
                    counts["missing_password_no_token"] += 1
                print("  UNRESOLVED: no workbook password to seed/reseed")
                continue

            row["action"] = "fresh_login_isolated" if initial == "missing" else "reseed_isolated"
            # tempfile guarantees a nonexistent path, and cleans up on every
            # failure; the live token remains in place until success.
            with tempfile.TemporaryDirectory(dir=token_root, prefix=".step_seed_stage_") as d:
                stage_path = Path(d) / live_path.name
                seed_one_account(email, password, stage_path)
                backup_and_promote(stage_path, live_path, token_root)
            row["verified_online"] = True
            row["outcome"] = "verified_new" if initial == "missing" else "verified_reseeded"
            if initial == "missing":
                counts["verified_new"] += 1
            else:
                counts["verified_reseeded"] += 1
            row["error"] = ""  # old verification error no longer relevant
            print("  VERIFIED and saved fresh token")

        except Exception as exc:
            err = sanitize_error(exc, email, password)
            row["error"] = err
            if is_rate_limit_error(exc):
                row["outcome"] = "unverified_seed_rate_limited"
                counts["rate_limited"] += 1
                print("  UNVERIFIED: fresh login/recheck rate limited; existing token preserved")
                if args.stop_on_rate_limit:
                    stop_after_429 = True
            elif is_connection_or_transient_error(exc):
                row["outcome"] = "unverified_seed_connection_error"
                counts["transient"] += 1
                print("  UNVERIFIED: network/service failure; existing token preserved")
            else:
                row["outcome"] = "unverified_seed_failed"
                counts["seed_failed"] += 1
                print(f"  UNVERIFIED: credential login/verification failed: {type(exc).__name__}: {err}")
        finally:
            row["token_present_after"] = token_state(live_path) == "structurally_valid"
            row["seconds"] = round(time.perf_counter() - start, 6)
            audit_rows.append(row)
            if args.inter_account_delay > 0 and idx < len(accounts) and not stop_after_429:
                time.sleep(args.inter_account_delay)

    summary_path, status = write_outputs(workbook, sheet, token_root,
                                         audit_rows, version, counts, loaded)
    total_verified = sum(row["verified_online"] for row in audit_rows)
    print("\n=== STEP GARMIN V3 TOKEN VERIFICATION SUMMARY ===")
    print(f"Status: {status}")
    print(f"Online verified: {total_verified}/{len(accounts)}")
    print(f"  Existing: {counts['verified_existing']}")
    print(f"  New: {counts['verified_new']}")
    print(f"  Reseeded: {counts['verified_reseeded']}")
    print(f"Unresolved (NOT counted valid): {len(accounts) - total_verified}")
    print(f"  Rate limited: {counts['rate_limited']}; network issues: {counts['transient']}")
    print(f"  Seed failures: {counts['seed_failed']}")
    print(f"Summary JSON: {summary_path}")
    if status != "all_verified":
        print("Check the audit CSV in the STEP workbook folder; rerun after resolving issues.")
        return 2
    return 0

if __name__ == "__main__":
    raise SystemExit(main())