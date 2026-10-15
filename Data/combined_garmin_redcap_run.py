#!/usr/bin/env python3
"""
Combined Garmin -> REDCap runner for SyncStep + SyncRove
=========================================================

Purpose
-------
Run BOTH already-seeded Garmin projects in one invocation while keeping each
project's data model separate:

SyncStep
  - Full Garmin dataset: steps, sleep, HR, HRV, respiration, SpO2, intensity.
  - Missingness is preserved independently for every metric; missing values are
    never converted to zero based on another metric.
  - 10 most recent pre-op days with any wearable data.
  - Post-op fields map to the ACTUAL postoperative day 1..90.
  - STEP wearable fields are pipeline-managed: re-upload uses REDCap overwrite
    semantics for those fields so stale placeholder/shifted values are cleared.
  - Existing STEP REDCap payload structure:
      record_id, redcap_data_access_group,
      watch_type, watch_number,
      preop{metric}_dayN, postop{metric}_dayN
  - Overnight HRV (ms): sleep dailySleepDTO.avgSleepHRV, then the dedicated
    HRV endpoint's hrvSummary.lastNightAvg. Garmin hrvSummary.status is stored
    separately. --hrv-mode auto inspects REDCap metadata to upload the correct
    concept. Ambiguous HRV fields are omitted, including blank overwrites.
  - Total intensity = moderate minutes + 2 * vigorous minutes. Both components
    must be present; an observed zero stays zero. Failed HRV retrieval blocks
    that account's upload instead of overwriting existing values with blanks.

SyncRove
  - Walking/steps-only project using the recovered Click4Steps download behavior.
  - Source workbook mapping is fixed to the established project headers:
      Study ID, Vivofit4 Serial Number, username, DOS, Start Date,
      Expected End Date, Walking data captured.
  - Query stops at the earliest planned end, recorded return, next assignment
    minus one day, or today. Missing expected end falls back to DOS + 90.
  - Both projects' complete assignment logs are checked before completed rows
    are skipped. Unresolved ownership conflicts block the affected records.
  - Missing return dates are flagged. A later Start Date for the same serial OR
    account stops the older assignment on the previous calendar day, including
    when the later assignment has different credentials or is already completed.
    Undated reuse and conflicting/same-day assignments require log correction.
    Log every handover: unrecorded reuse cannot be inferred from step patterns.
  - Exact return dates are inclusive, except that a same-day reassignment owns
    the handover day. Month-only return dates exclude the entire return month
    conservatively; no exact return date is invented. Missing return plus missing
    serial blocks the record because cross-account device reuse cannot be checked.
  - Existing out-of-window REDCap repeats are listed in REDCap_Out_Of_Window.csv.
    They are preserved for review; this update does not delete past imports.
  - Daily REDCap rows are retained for relative days -90 through +120.
  - Current REDCap import structure is a flat CSV with REDCap routing columns
    followed by the project's live metadata fields. Daily step rows populate:
      record_id, redcap_event_name, redcap_repeat_instrument,
      redcap_repeat_instance, step_data_relative_day, and steps;
    unrelated project fields and step_data_complete remain blank.
  - Existing step_data repeats are read before upload. Existing relative days
    reuse their REDCap repeat instance; new relative days receive new instances.
    Duplicate existing relative days fail safely instead of creating ambiguity.

Authentication
--------------
This script is TOKEN-FIRST / TOKEN-ONLY for Garmin. It does NOT perform fresh
Garmin username/password logins. That is intentional because both projects
have already been seeded. If a token is missing or invalid, the account is
logged as failed and the run continues.

You can select a seed folder in either form:
    /path/to/Seed
or:
    /path/to/Seed/.garmin_tokens_v3

The script resolves the actual .garmin_tokens_v3 folder automatically.

REDCap tokens are kept separate by project:
  SyncStep search order:
      SYNCSTEP_REDCAP_TOKEN -> api_token -> REDCAP_API_TOKEN -> secure prompt
  SyncRove search order:
      SYNCROVE_REDCAP_TOKEN -> ROVE_REDCAP_TOKEN -> secure prompt

SyncRove intentionally does NOT silently fall back to the STEP token, which
prevents accidental upload to the wrong REDCap project.

Performance outputs
-------------------
The manuscript/pipeline wall clock starts only after preflight is complete and
after the user types RUN (or, with --yes, immediately before automated account
processing begins). File selection, token entry, REDCap metadata validation, and
local staging of the Rove workbook/token folder are therefore excluded from
total_wall_seconds.

A timestamped combined-results folder is created. Major files include:
  Combined_Performance_Attempts.csv
      One row per account attempt, including login, date-loop download,
      data sort/transform, REDCap upload, and total timing.
  Combined_Download_Iterations.csv
      One row per requested date iteration, with project/watch type and timing.
  Combined_Performance_By_Project_Watch.csv
      Mean/SD/median/min/max timing by PROJECT + WATCH TYPE.
  Combined_Performance_Manuscript_Long.csv
      Long-form descriptive statistics for manuscript/table use.
  Combined_Run_Summary.csv
      One row per project plus an overall row.
  Combined_Rate_Limit_Events.csv
      429/rate-limit events without passwords or REDCap tokens.
  Assignment_Window_Audit.csv
      All logged rows with chosen boundaries, return precision, reuse warnings,
      and reasons for blocked ownership. Account identifiers are hashed.
  STEP_HRV_Preflight.json
      HRV concept chosen from live REDCap metadata and status choice-code mapping.

Project-specific payload/audit snapshots are also written under:
  <run folder>/SyncStep/
  <run folder>/SyncRove/

Workbook write-back
-------------------
SyncStep write-back is formula-safe by default: xlwings/Excel is used when
available. If xlwings is unavailable, the STEP workbook is NOT saved with
openpyxl unless --allow-openpyxl-step-writeback is explicitly supplied.
No rows/columns are inserted by this combined script.

SyncRove writes only to already-existing status columns. Missing columns are
skipped rather than inserted. For timing runs, the Rove workbook and token folder
are copied from the mounted drive to local staging before RUN; timed execution uses
only the local copies. After timing stops, the script can optionally copy the
updated workbook and refreshed token files back to the mounted drive.

Examples
--------
Interactive selection:
    python combined_garmin_redcap_run.py

Fully specified:
    python combined_garmin_redcap_run.py \
      --step-workbook "/path/to/STEP.xlsx" \
      --step-seed-folder "/path/to/STEP/Seed" \
      --rove-workbook "/path/to/ROVE.xlsx" \
      --rove-seed-folder "/path/to/ROVE/Seed" \
      --yes

Download + transform only, no REDCap writes and no workbook write-back:
    python combined_garmin_redcap_run.py --no-upload --no-writeback

Disable intentional Garmin inter-call pacing (retry backoff remains enabled):
    python combined_garmin_redcap_run.py --pacing false

HRV metadata normally resolves automatically. To upload overnight HRV to numeric
text fields after confirming their definition is milliseconds:
    python combined_garmin_redcap_run.py --hrv-mode overnight-ms
To require categorical status fields/choices:
    python combined_garmin_redcap_run.py --hrv-mode status
An incompatible explicit mode stops in preflight. No REDCap metadata is changed.

Dependencies
------------
    python -m pip install pandas requests openpyxl "garminconnect>=0.3.10"
Optional for formula-safe Excel write-back:
    python -m pip install xlwings
"""

from __future__ import annotations

import argparse
import csv
import io
import getpass
import hashlib
import importlib.metadata
import inspect
import json
import math
import os
import random
import re
import shutil
import sys
import time
import traceback
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Callable, Iterable, Optional

import pandas as pd
import requests
from openpyxl import load_workbook
from garminconnect import Garmin

try:
    from garminconnect.exceptions import GarminConnectTooManyRequestsError
except Exception:
    GarminConnectTooManyRequestsError = Exception

try:
    import tkinter as tk
    from tkinter import filedialog
    _HAS_TK = True
except Exception:
    _HAS_TK = False

try:
    import xlwings as xw
    _HAS_XLW = True
except Exception:
    _HAS_XLW = False


# =============================================================================
# CONSTANTS
# =============================================================================

# Set REDCAP_API_URL in the environment to point at your institution's REDCap API.
REDCAP_API_URL = os.getenv("REDCAP_API_URL", "https://redcap.ucdenver.edu/api/")
MIN_GARMINCONNECT_VERSION = (0, 3, 10)
TOKEN_FOLDER_NAME = ".garmin_tokens_v3"

# The user already tested modern Garmin data retrieval around this ceiling.
DEFAULT_DATA_QPS = 2.0
DEFAULT_RETRY_QPS = 0.8
DEFAULT_NETWORK_RETRIES = 3
DEFAULT_RETRY_WAIT_SECONDS = 5 * 60
DEFAULT_JITTER_SECONDS = 0.10

STEP_PREOP_DAYS = 10
STEP_POSTOP_DAYS = 90
ROVE_PREOP_DAYS = 90  # retained only for descriptive compatibility
ROVE_POSTOP_DAYS = 120
ROVE_DEFAULT_POSTOP_DAYS = 90  # fallback only when Expected End Date is blank
ROVE_MIN_RELATIVE_DAY = -90

STEP_FIELDS = [
    "steps", "totalsleep", "lightsleep", "deepsleep", "awake", "remsleep",
    "restinghr", "maxhr", "hrv", "rr", "pulseox",
    "totalintensity", "modintensity", "vigintensity",
]

STEP_SLEEP_FIELDS = (
    "totalsleep", "lightsleep", "deepsleep", "awake", "remsleep",
)

# STEP workbook team labels -> live REDCap DAG unique names.
# Keys are canonicalized with canon(), so case and surrounding whitespace do not
# matter (for example, " Gastro " and "gastro" both resolve to "gastro").
#
# IMPORTANT: "UNK" is intentionally NOT mapped. An unknown team must be corrected
# in the workbook (or otherwise explicitly resolved) rather than silently routed to
# a REDCap data access group.
STEP_DAG_MAP = {
    "thoracic": "cardiothoracic",
    "ortho": "orthopedics",
    "esoph": "esophageal",
    "esophageal": "esophageal",
    "gastro": "gastro",
    "kidney": "kidney",
    "pancreas": "pancreas",
    "plastics": "plastics",
    "transplant": "transplant",
    "urology": "urology",
    "vascular": "vascular",
}


def normalize_step_team(value: Any) -> str:
    """Canonical STEP team label used only for DAG routing/validation."""
    return canon(value)


def resolve_step_dag(value: Any) -> Optional[str]:
    """Return the REDCap DAG unique name for a STEP team, or None if unresolved."""
    return STEP_DAG_MAP.get(normalize_step_team(value))

TRANSIENT_ERROR_TOKENS = (
    "timed out", "timeout", "read timed out", "connection reset",
    "connection aborted", "connection error", "temporarily unavailable",
    "max retries exceeded", "502", "503", "504",
)

RATE_LIMIT_TOKENS = (
    "429", "rate limit", "too many requests", "too many login attempts",
)


# =============================================================================
# BASIC HELPERS
# =============================================================================


def now_local_iso() -> str:
    return datetime.now(timezone.utc).astimezone().replace(microsecond=0).isoformat()


def utc_stamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def canon(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").lower().strip()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def canon_compact(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def clean_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def parse_bool(value: Any) -> bool:
    """argparse helper accepting T/F, true/false, yes/no, and 1/0."""
    if isinstance(value, bool):
        return value
    normalized = canon(value)
    if normalized in {"true", "t", "yes", "y", "1", "on"}:
        return True
    if normalized in {"false", "f", "no", "n", "0", "off"}:
        return False
    raise argparse.ArgumentTypeError(f"Expected T/F (or true/false), got {value!r}")


def safe_date(value: Any) -> Optional[date]:
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def is_truthy(value: Any) -> bool:
    return canon(value) in {"true", "1", "yes", "y", "done", "complete", "completed"}


def first_not_none(*values: Any) -> Any:
    for value in values:
        if value is not None:
            return value
    return None


def version_tuple(version_text: str) -> tuple[int, int, int]:
    nums = re.findall(r"\d+", str(version_text or ""))
    vals = [int(x) for x in nums[:3]]
    while len(vals) < 3:
        vals.append(0)
    return tuple(vals)  # type: ignore[return-value]


def ensure_supported_garminconnect() -> str:
    version = None
    for package_name in ("garminconnect", "python-garminconnect"):
        try:
            version = importlib.metadata.version(package_name)
            break
        except Exception:
            pass
    if version is None or version_tuple(version) < MIN_GARMINCONNECT_VERSION:
        wanted = ".".join(str(x) for x in MIN_GARMINCONNECT_VERSION)
        raise RuntimeError(
            f"This runner requires garminconnect >= {wanted}. "
            f"Installed version: {version or 'not found'}.\n"
            "Upgrade with: python -m pip install --upgrade garminconnect"
        )
    return version


def is_rate_limit_error(message: str) -> bool:
    low = str(message or "").lower()
    return any(token in low for token in RATE_LIMIT_TOKENS)


def is_transient_error(message: str) -> bool:
    low = str(message or "").lower()
    return any(token in low for token in TRANSIENT_ERROR_TOKENS)


def account_hash(username: str) -> str:
    return hashlib.sha256(clean_text(username).lower().encode("utf-8")).hexdigest()[:24]


def descriptive_stats(values: Iterable[Any]) -> dict[str, Any]:
    s = pd.to_numeric(pd.Series(list(values)), errors="coerce").dropna()
    n = int(len(s))
    if n == 0:
        return {
            "n": 0, "mean": None, "sd": None, "median": None,
            "q1": None, "q3": None, "iqr": None, "min": None, "max": None,
        }
    q1 = float(s.quantile(0.25))
    q3 = float(s.quantile(0.75))
    return {
        "n": n,
        "mean": float(s.mean()),
        "sd": float(s.std(ddof=1)) if n > 1 else None,
        "median": float(s.median()),
        "q1": q1,
        "q3": q3,
        "iqr": q3 - q1,
        "min": float(s.min()),
        "max": float(s.max()),
    }


def json_safe(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, (pd.Timestamp,)):
        return value.isoformat()
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def sanitize_record(record: dict[str, Any]) -> dict[str, Any]:
    return {str(k): json_safe(v) for k, v in record.items()}


def ensure_dir(path: Path) -> Path:
    path.mkdir(parents=True, exist_ok=True)
    return path


def stage_rove_inputs(
    source_workbook: Path,
    source_token_dir: Path,
    staging_parent: Path,
) -> tuple[Path, Path, Path]:
    """Copy all mounted-drive Rove inputs to local storage before RUN.

    The returned workbook/token paths are the ONLY Rove input paths used during
    timed execution. This deliberately removes SMB/VPN-mounted-drive latency
    from the manuscript timing interval.
    """
    stage_root = ensure_dir(staging_parent / f"SyncRove_Stage_{utc_stamp()}")
    staged_workbook = stage_root / source_workbook.name
    staged_token_dir = stage_root / ".garmin_tokens_v3"

    print("\n[SyncRove] Staging mounted-drive inputs locally BEFORE timing...")
    shutil.copy2(source_workbook, staged_workbook)
    shutil.copytree(
        source_token_dir,
        staged_token_dir,
        dirs_exist_ok=True,
        copy_function=shutil.copy2,
    )

    if not staged_workbook.is_file():
        raise RuntimeError(f"SyncRove workbook staging failed: {staged_workbook}")
    token_files = [p for p in staged_token_dir.rglob("*") if p.is_file()]
    if not token_files:
        raise RuntimeError(
            f"SyncRove token staging produced no files in {staged_token_dir}"
        )

    print(f"  Source workbook: {source_workbook}")
    print(f"  Local workbook:  {staged_workbook}")
    print(f"  Source tokens:   {source_token_dir}")
    print(f"  Local tokens:    {staged_token_dir}")
    print(f"  Token files staged: {len(token_files)}")
    return staged_workbook, staged_token_dir, stage_root


def sync_rove_stage_back(
    staged_workbook: Path,
    staged_token_dir: Path,
    source_workbook: Path,
    source_token_dir: Path,
    *,
    copy_workbook: bool,
) -> None:
    """Copy staged Rove state back to the mounted drive AFTER timing stops."""
    if not source_workbook.parent.exists():
        raise RuntimeError(
            f"Mounted-drive workbook folder is not available: {source_workbook.parent}"
        )
    if not source_token_dir.parent.exists():
        raise RuntimeError(
            f"Mounted-drive token parent is not available: {source_token_dir.parent}"
        )

    if copy_workbook:
        shutil.copy2(staged_workbook, source_workbook)
        print(f"[SyncRove] Updated staged workbook copied back to: {source_workbook}")

    source_token_dir.mkdir(parents=True, exist_ok=True)
    shutil.copytree(
        staged_token_dir,
        source_token_dir,
        dirs_exist_ok=True,
        copy_function=shutil.copy2,
    )
    print(f"[SyncRove] Staged/refreshed Garmin tokens copied back to: {source_token_dir}")


def write_csv(rows: Iterable[dict[str, Any]], path: Path, columns: Optional[list[str]] = None) -> None:
    rows_list = list(rows)
    if rows_list:
        df = pd.DataFrame(rows_list)
    else:
        df = pd.DataFrame(columns=columns or [])
    df.to_csv(path, index=False)


# =============================================================================
# FILE / FOLDER SELECTION
# =============================================================================


def _tk_root():
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
    except Exception:
        pass
    return root


def choose_file(cli_value: Optional[str], title: str) -> Path:
    if cli_value:
        p = Path(cli_value).expanduser().resolve()
        if p.is_file():
            return p
        raise FileNotFoundError(f"File not found: {p}")

    if _HAS_TK:
        try:
            root = _tk_root()
            try:
                selected = filedialog.askopenfilename(
                    title=title,
                    filetypes=[
                        ("Excel workbooks", "*.xlsx *.xlsm *.xltx *.xltm"),
                        ("All files", "*.*"),
                    ],
                )
            finally:
                root.destroy()
            if selected:
                return Path(selected).expanduser().resolve()
        except Exception:
            pass

    while True:
        raw = input(f"{title} path: ").strip().strip('"')
        p = Path(raw).expanduser().resolve()
        if p.is_file():
            return p
        print("File not found. Try again.")


def choose_folder(cli_value: Optional[str], title: str) -> Path:
    if cli_value:
        p = Path(cli_value).expanduser().resolve()
        if p.is_dir():
            return p
        raise FileNotFoundError(f"Folder not found: {p}")

    if _HAS_TK:
        try:
            root = _tk_root()
            try:
                selected = filedialog.askdirectory(title=title)
            finally:
                root.destroy()
            if selected:
                return Path(selected).expanduser().resolve()
        except Exception:
            pass

    while True:
        raw = input(f"{title} path: ").strip().strip('"')
        p = Path(raw).expanduser().resolve()
        if p.is_dir():
            return p
        print("Folder not found. Try again.")


def resolve_token_dir(selected: Path) -> Path:
    selected = selected.expanduser().resolve()
    if selected.name == TOKEN_FOLDER_NAME and selected.is_dir():
        return selected
    child = selected / TOKEN_FOLDER_NAME
    if child.is_dir():
        return child
    # Also accept selecting the directory containing flat token JSONs directly.
    flat_tokens = list(selected.glob("*.json"))
    if flat_tokens:
        return selected
    raise FileNotFoundError(
        f"Could not find {TOKEN_FOLDER_NAME} under {selected}, and the selected "
        "folder does not contain token JSON files."
    )


def token_path(token_dir: Path, username: str) -> Path:
    return token_dir / f"{account_hash(username)}.json"


def token_file_structurally_valid(path: Path) -> bool:
    try:
        if not path.is_file() or path.stat().st_size <= 20:
            return False
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return False
        # Modern 0.3.x DI OAuth token files normally contain these fields. Keep
        # the check permissive enough for compatible library revisions.
        expected_any = (
            "di_token", "di_refresh_token", "di_client_id",
            "access_token", "refresh_token",
        )
        return any(data.get(key) for key in expected_any)
    except Exception:
        return False


def choose_sheet(workbook: Path, cli_sheet: Optional[str], project: str) -> str:
    wb = load_workbook(workbook, read_only=True, data_only=False)
    try:
        names = list(wb.sheetnames)
    finally:
        wb.close()

    if cli_sheet:
        if cli_sheet not in names:
            raise KeyError(f"Sheet {cli_sheet!r} not found in {workbook}. Available: {names}")
        return cli_sheet

    if project == "SyncStep" and "Master Spreadsheet" in names:
        return "Master Spreadsheet"

    # Let each project header detector choose among sheets.
    candidates = []
    for name in names:
        try:
            if project == "SyncStep":
                header_row, raw = autodetect_step_header_row(workbook, name)
                if header_row is not None:
                    candidates.append(name)
            else:
                header_row, raw = autodetect_rove_header_row(workbook, name)
                if header_row is not None:
                    candidates.append(name)
        except Exception:
            pass

    if len(candidates) == 1:
        return candidates[0]
    if candidates:
        print(f"Multiple {project} sheets look usable:")
        for i, name in enumerate(candidates, 1):
            print(f"  {i}. {name}")
        while True:
            raw = input("Select sheet number: ").strip()
            if raw.isdigit() and 1 <= int(raw) <= len(candidates):
                return candidates[int(raw) - 1]

    if len(names) == 1:
        return names[0]

    print(f"Available sheets in {workbook.name}:")
    for i, name in enumerate(names, 1):
        print(f"  {i}. {name}")
    while True:
        raw = input("Select sheet number: ").strip()
        if raw.isdigit() and 1 <= int(raw) <= len(names):
            return names[int(raw) - 1]


# =============================================================================
# REDCAP CLIENT + METADATA ADAPTER
# =============================================================================


class RedcapError(RuntimeError):
    pass


@dataclass
class RedcapSchema:
    project: str
    record_field: str
    field_names: set[str]
    metadata: list[dict[str, Any]]
    serial_field: Optional[str] = None
    watch_type_field: Optional[str] = None
    preop_steps_fields: dict[int, str] = field(default_factory=dict)
    postop_steps_fields: dict[int, str] = field(default_factory=dict)
    # SyncRove may store daily steps in a repeating/long instrument instead of
    # one participant-wide preop/postop record. STEP continues to use wide mode.
    mode: str = "wide"
    relative_day_field: Optional[str] = None
    steps_field: Optional[str] = None
    completion_field: Optional[str] = None
    repeat_instrument: Optional[str] = None
    repeat_event: Optional[str] = None
    repeat_is_event: bool = False
    # Exact flat REDCap import-column order for SyncRove. The first four are
    # REDCap system columns; the remainder follow live metadata order.
    import_headers: list[str] = field(default_factory=list)
    hrv_mode: str = "overnight-ms"
    hrv_status_choices: dict[str, dict[str, str]] = field(default_factory=dict)
    hrv_warning: str = ""


class RedcapClient:
    def __init__(self, project: str, token: str, api_url: str = REDCAP_API_URL):
        self.project = project
        self.token = token.strip()
        self.api_url = api_url
        self.session = requests.Session()
        self._metadata: Optional[list[dict[str, Any]]] = None

    def metadata(self) -> list[dict[str, Any]]:
        if self._metadata is not None:
            return self._metadata
        payload = {
            "token": self.token,
            "content": "metadata",
            "format": "json",
            "returnFormat": "json",
        }
        try:
            r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap metadata request failed: {exc}") from exc
        if r.status_code != 200:
            raise RedcapError(
                f"{self.project} REDCap metadata HTTP {r.status_code}: {r.text[:500]}"
            )
        try:
            parsed = r.json()
        except Exception as exc:
            raise RedcapError(
                f"{self.project} REDCap metadata was not JSON: {r.text[:500]}"
            ) from exc
        if isinstance(parsed, dict) and parsed.get("error"):
            raise RedcapError(f"{self.project} REDCap metadata error: {parsed['error']}")
        if not isinstance(parsed, list) or not parsed:
            raise RedcapError(f"{self.project} REDCap metadata response was empty/unexpected")
        self._metadata = parsed
        return parsed

    def field_names(self) -> set[str]:
        return {
            clean_text(row.get("field_name"))
            for row in self.metadata()
            if clean_text(row.get("field_name"))
        }

    def repeating_forms_events(self) -> list[dict[str, Any]]:
        payload = {
            "token": self.token,
            "content": "repeatingFormsEvents",
            "format": "json",
            "returnFormat": "json",
        }
        try:
            r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
        except Exception as exc:
            raise RedcapError(
                f"{self.project} REDCap repeating-form request failed: {exc}"
            ) from exc
        if r.status_code != 200:
            raise RedcapError(
                f"{self.project} REDCap repeating-form HTTP {r.status_code}: {r.text[:500]}"
            )
        try:
            parsed = r.json()
        except Exception as exc:
            raise RedcapError(
                f"{self.project} REDCap repeating-form response was not JSON: {r.text[:500]}"
            ) from exc
        if isinstance(parsed, dict) and parsed.get("error"):
            raise RedcapError(
                f"{self.project} REDCap repeating-form error: {parsed['error']}"
            )
        if parsed in (None, ""):
            return []
        if not isinstance(parsed, list):
            raise RedcapError(
                f"{self.project} REDCap repeating-form response was unexpected: {str(parsed)[:500]}"
            )
        return parsed

    def form_event_mapping(self) -> list[dict[str, Any]]:
        payload = {
            "token": self.token,
            "content": "formEventMapping",
            "format": "json",
            "returnFormat": "json",
        }
        try:
            r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
        except Exception as exc:
            raise RedcapError(
                f"{self.project} REDCap form-event mapping request failed: {exc}"
            ) from exc
        if r.status_code != 200:
            raise RedcapError(
                f"{self.project} REDCap form-event mapping HTTP {r.status_code}: {r.text[:500]}"
            )
        try:
            parsed = r.json()
        except Exception as exc:
            raise RedcapError(
                f"{self.project} REDCap form-event mapping was not JSON: {r.text[:500]}"
            ) from exc
        if isinstance(parsed, dict) and parsed.get("error"):
            # Classic/non-longitudinal projects may not need this endpoint.
            return []
        return parsed if isinstance(parsed, list) else []

    def probe_rove_repeat_rows(
        self,
        record_id: str,
        record_field: str,
        relative_day_field: str,
        steps_field: str,
    ) -> list[dict[str, Any]]:
        """Export only the fields needed to infer Rove repeat layout.

        This uses ordinary API Export privileges and intentionally avoids the
        repeatingFormsEvents / Project Setup endpoint. REDCap still includes
        redcap_event_name, redcap_repeat_instrument, and redcap_repeat_instance
        in flat exports when applicable.
        """
        payload: dict[str, Any] = {
            "token": self.token,
            "content": "record",
            "format": "json",
            "type": "flat",
            "records[0]": record_id,
            "fields[0]": record_field,
            "fields[1]": relative_day_field,
            "fields[2]": steps_field,
            "returnFormat": "json",
            "exportDataAccessGroups": "false",
        }
        try:
            r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap repeat-layout probe failed: {exc}") from exc
        if r.status_code != 200:
            raise RedcapError(
                f"{self.project} REDCap repeat-layout probe HTTP {r.status_code}: {r.text[:500]}"
            )
        try:
            parsed = r.json()
        except Exception as exc:
            raise RedcapError(
                f"{self.project} REDCap repeat-layout probe was not JSON: {r.text[:500]}"
            ) from exc
        if isinstance(parsed, dict) and parsed.get("error"):
            raise RedcapError(f"{self.project} REDCap repeat-layout probe error: {parsed['error']}")
        if not isinstance(parsed, list):
            raise RedcapError(f"{self.project} REDCap repeat-layout probe response was unexpected")
        return parsed

    def export_record_exists(self, record_id: str, record_field: str) -> bool:
        payload = {
            "token": self.token,
            "content": "record",
            "format": "json",
            "type": "flat",
            "records[0]": record_id,
            "fields[0]": record_field,
            "returnFormat": "json",
            "exportDataAccessGroups": "false",
        }
        try:
            r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap record-existence export failed: {exc}") from exc
        if r.status_code != 200:
            raise RedcapError(f"{self.project} REDCap record-existence HTTP {r.status_code}: {r.text[:500]}")
        try:
            parsed = r.json()
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap record-existence response was not JSON: {r.text[:500]}") from exc
        if isinstance(parsed, dict) and parsed.get("error"):
            raise RedcapError(f"{self.project} REDCap record-existence error: {parsed['error']}")
        return isinstance(parsed, list) and len(parsed) > 0

    def export_rove_step_rows(self, record_id: str, schema: RedcapSchema) -> list[dict[str, Any]]:
        fields = [schema.record_field, schema.relative_day_field, schema.steps_field, schema.completion_field]
        payload: dict[str, Any] = {
            "token": self.token,
            "content": "record",
            "format": "json",
            "type": "flat",
            "records[0]": record_id,
            "returnFormat": "json",
            "exportDataAccessGroups": "false",
        }
        for i, field_name in enumerate([f for f in fields if f]):
            payload[f"fields[{i}]"] = field_name
        try:
            r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap existing-step export failed: {exc}") from exc
        if r.status_code != 200:
            raise RedcapError(f"{self.project} REDCap existing-step HTTP {r.status_code}: {r.text[:500]}")
        try:
            parsed = r.json()
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap existing-step response was not JSON: {r.text[:500]}") from exc
        if isinstance(parsed, dict) and parsed.get("error"):
            raise RedcapError(f"{self.project} REDCap existing-step error: {parsed['error']}")
        if not isinstance(parsed, list):
            raise RedcapError(f"{self.project} REDCap existing-step response was unexpected")
        rows = []
        for row in parsed:
            if schema.repeat_event and clean_text(row.get("redcap_event_name")) != schema.repeat_event:
                continue
            repeat_instrument = clean_text(row.get("redcap_repeat_instrument"))
            if schema.repeat_is_event:
                if repeat_instrument:
                    continue
            elif canon_compact(repeat_instrument) != canon_compact(schema.repeat_instrument):
                continue
            rows.append(row)
        return rows

    def data_access_groups(self) -> list[dict[str, Any]]:
        """Return REDCap DAG metadata when the API token is permitted to view it."""
        payload = {
            "token": self.token,
            "content": "dag",
            "format": "json",
            "returnFormat": "json",
        }
        try:
            r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap DAG request failed: {exc}") from exc
        if r.status_code != 200:
            raise RedcapError(f"{self.project} REDCap DAG HTTP {r.status_code}: {r.text[:500]}")
        try:
            parsed = r.json()
        except Exception as exc:
            raise RedcapError(f"{self.project} REDCap DAG response was not JSON: {r.text[:500]}") from exc
        if isinstance(parsed, dict) and parsed.get("error"):
            raise RedcapError(f"{self.project} REDCap DAG error: {parsed['error']}")
        if not isinstance(parsed, list):
            raise RedcapError(f"{self.project} REDCap DAG response was unexpected")
        return parsed

    def export_record_fields(
        self,
        record_id: str,
        record_field: str,
        fields: Iterable[str],
        batch_size: int = 800,
    ) -> dict[str, Any]:
        """Export selected fields for one record in bounded batches for readback verification."""
        requested: list[str] = []
        seen = {clean_text(record_field)}
        for field_name in list(fields):
            field_name = clean_text(field_name)
            if field_name and field_name not in seen:
                seen.add(field_name)
                requested.append(field_name)

        merged: dict[str, Any] = {}
        chunk_size = max(1, int(batch_size))
        chunks = [requested[i:i + chunk_size] for i in range(0, len(requested), chunk_size)] or [[]]
        for chunk_values in chunks:
            chunk = [record_field, *chunk_values]
            payload: dict[str, Any] = {
                "token": self.token,
                "content": "record",
                "format": "json",
                "type": "flat",
                "records[0]": record_id,
                "returnFormat": "json",
                "exportDataAccessGroups": "false",
            }
            for i, field_name in enumerate(chunk):
                payload[f"fields[{i}]"] = field_name
            try:
                r = self.session.post(self.api_url, data=payload, timeout=(10, 120))
            except Exception as exc:
                raise RedcapError(f"{self.project} REDCap readback failed: {exc}") from exc
            if r.status_code != 200:
                raise RedcapError(f"{self.project} REDCap readback HTTP {r.status_code}: {r.text[:500]}")
            try:
                parsed = r.json()
            except Exception as exc:
                raise RedcapError(f"{self.project} REDCap readback was not JSON: {r.text[:500]}") from exc
            if isinstance(parsed, dict) and parsed.get("error"):
                raise RedcapError(f"{self.project} REDCap readback error: {parsed['error']}")
            if not isinstance(parsed, list) or not parsed:
                raise RedcapError(f"{self.project} REDCap readback returned no rows for record {record_id}")
            matching = [row for row in parsed if clean_text(row.get(record_field)) == clean_text(record_id)]
            if len(matching) != 1:
                raise RedcapError(
                    f"{self.project} REDCap readback expected exactly one flat row for {record_id}; "
                    f"found {len(matching)}"
                )
            merged.update(matching[0])
        return merged

    def upload_records(
        self,
        records: list[dict[str, Any]],
        overwrite_behavior: str = "normal",
        retries: int = 2,
    ) -> tuple[bool, int, str, float]:
        if not records:
            return True, 200, '{"count":0}', 0.0
        payload_json = json.dumps(
            [sanitize_record(record) for record in records], ensure_ascii=False
        )
        payload = {
            "token": self.token,
            "content": "record",
            "format": "json",
            "type": "flat",
            "overwriteBehavior": overwrite_behavior,
            "data": payload_json,
            "returnContent": "count",
            "returnFormat": "json",
        }

        last_exc: Optional[Exception] = None
        for attempt in range(1, retries + 2):
            t0 = time.perf_counter()
            try:
                r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
                elapsed = time.perf_counter() - t0
                response_text = r.text or ""
                ok = r.status_code == 200 and "error" not in response_text.lower()
                if ok:
                    return True, r.status_code, response_text, elapsed
                if attempt <= retries and (r.status_code >= 500 or r.status_code == 429):
                    time.sleep(min(2 ** (attempt - 1), 8))
                    continue
                return False, r.status_code, response_text, elapsed
            except Exception as exc:
                elapsed = time.perf_counter() - t0
                last_exc = exc
                if attempt <= retries and is_transient_error(str(exc)):
                    time.sleep(min(2 ** (attempt - 1), 8))
                    continue
                raise RedcapError(
                    f"{self.project} REDCap upload failed: {exc}"
                ) from exc

        raise RedcapError(f"{self.project} REDCap upload failed: {last_exc}")

    def upload_records_csv(
        self,
        records: list[dict[str, Any]],
        headers: list[str],
        overwrite_behavior: str = "normal",
        retries: int = 2,
    ) -> tuple[bool, int, str, float]:
        """Import flat REDCap rows using an explicit CSV header/order."""
        if not records:
            return True, 200, '{"count":0}', 0.0
        required = {"redcap_event_name", "redcap_repeat_instrument", "redcap_repeat_instance",
                    "step_data_relative_day", "steps"}
        if not headers or not headers[0]:
            raise RedcapError("SyncRove CSV import header has no record identifier column.")
        required.add(headers[0])
        missing = sorted(required - set(headers))
        if missing:
            raise RedcapError(f"SyncRove CSV import header is missing required columns: {missing}")
        csv_text = redcap_rows_to_csv_text(records, headers)
        payload = {
            "token": self.token,
            "content": "record",
            "format": "csv",
            "type": "flat",
            "overwriteBehavior": overwrite_behavior,
            "data": csv_text,
            "returnContent": "count",
            "returnFormat": "json",
        }
        last_exc: Optional[Exception] = None
        for attempt in range(1, retries + 2):
            t0 = time.perf_counter()
            try:
                r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
                elapsed = time.perf_counter() - t0
                response_text = r.text or ""
                ok = r.status_code == 200 and "error" not in response_text.lower()
                if ok:
                    return True, r.status_code, response_text, elapsed
                if attempt <= retries and (r.status_code >= 500 or r.status_code == 429):
                    time.sleep(min(2 ** (attempt - 1), 8))
                    continue
                return False, r.status_code, response_text, elapsed
            except Exception as exc:
                elapsed = time.perf_counter() - t0
                last_exc = exc
                if attempt <= retries and is_transient_error(str(exc)):
                    time.sleep(min(2 ** (attempt - 1), 8))
                    continue
                raise RedcapError(f"{self.project} REDCap CSV upload failed: {exc}") from exc
        raise RedcapError(f"{self.project} REDCap CSV upload failed: {last_exc}")


    def upload_record(
        self,
        record: dict[str, Any],
        overwrite_behavior: str = "normal",
        retries: int = 2,
    ) -> tuple[bool, int, str, float]:
        payload_json = json.dumps([sanitize_record(record)], ensure_ascii=False)
        payload = {
            "token": self.token,
            "content": "record",
            "format": "json",
            "type": "flat",
            "overwriteBehavior": overwrite_behavior,
            "data": payload_json,
            "returnContent": "count",
            "returnFormat": "json",
        }

        last_exc: Optional[Exception] = None
        for attempt in range(1, retries + 2):
            t0 = time.perf_counter()
            try:
                r = self.session.post(self.api_url, data=payload, timeout=(10, 60))
                elapsed = time.perf_counter() - t0
                text = r.text or ""
                ok = r.status_code == 200 and "error" not in text.lower()
                if ok:
                    return True, r.status_code, text, elapsed
                if attempt <= retries and (r.status_code >= 500 or r.status_code == 429):
                    time.sleep(min(2 ** (attempt - 1), 8))
                    continue
                return False, r.status_code, text, elapsed
            except Exception as exc:
                elapsed = time.perf_counter() - t0
                last_exc = exc
                if attempt <= retries and is_transient_error(str(exc)):
                    time.sleep(min(2 ** (attempt - 1), 8))
                    continue
                raise RedcapError(f"{self.project} REDCap upload failed: {exc}") from exc

        raise RedcapError(f"{self.project} REDCap upload failed: {last_exc}")


def get_redcap_token(project: str, no_upload: bool) -> Optional[str]:
    if no_upload:
        return None

    if project == "SyncStep":
        env_names = ["SYNCSTEP_REDCAP_TOKEN", "api_token", "REDCAP_API_TOKEN"]
    else:
        env_names = ["SYNCROVE_REDCAP_TOKEN", "ROVE_REDCAP_TOKEN"]

    for env_name in env_names:
        value = os.getenv(env_name)
        if value and value.strip():
            print(f"[{project}] REDCap token loaded from environment variable {env_name}.")
            return value.strip()

    print(f"[{project}] REDCap token not found in the expected environment variables.")
    value = getpass.getpass(f"Enter {project} REDCap API token (input hidden): ").strip()
    if not value:
        raise ValueError(f"A REDCap token is required for {project} unless --no-upload is used.")
    return value


def find_field_by_canonical(field_names: set[str], *candidates: str) -> Optional[str]:
    by_compact = {canon_compact(name): name for name in field_names}
    for candidate in candidates:
        exact = clean_text(candidate)
        if exact in field_names:
            return exact
        compact = canon_compact(candidate)
        if compact in by_compact:
            return by_compact[compact]
    return None


def build_step_redcap_schema(client: RedcapClient, hrv_mode: str = "auto") -> RedcapSchema:
    metadata = client.metadata()
    fields = client.field_names()
    record_field = find_field_by_canonical(fields, "record_id")
    if not record_field:
        raise RedcapError(
            "SyncStep REDCap metadata does not contain the expected primary field 'record_id'. "
            "Upload stopped before any records were written."
        )
    for required in ("watch_type", "watch_number", "preopsteps_day1", "postopsteps_day1"):
        if find_field_by_canonical(fields, required) is None:
            raise RedcapError(
                f"SyncStep REDCap metadata is missing expected field {required!r}. "
                "Upload stopped before any records were written."
            )

    expected_wearable_fields = [
        *(f"preop{metric}_day{day}" for day in range(1, STEP_PREOP_DAYS + 1) for metric in STEP_FIELDS),
        *(f"postop{metric}_day{day}" for day in range(1, STEP_POSTOP_DAYS + 1) for metric in STEP_FIELDS),
    ]
    missing_wearable = [name for name in expected_wearable_fields if name not in fields]
    if missing_wearable:
        preview = ", ".join(missing_wearable[:20])
        suffix = " ..." if len(missing_wearable) > 20 else ""
        raise RedcapError(
            f"SyncStep REDCap metadata is missing {len(missing_wearable)} pipeline-managed wearable "
            f"field(s): {preview}{suffix}. Upload stopped before Garmin retrieval so values cannot be "
            "silently dropped or misaligned."
        )

    resolved_hrv, status_choices, hrv_warning = resolve_hrv_schema(metadata, hrv_mode)
    return RedcapSchema(
        project="SyncStep",
        record_field=record_field,
        field_names=fields,
        metadata=metadata,
        serial_field=find_field_by_canonical(fields, "watch_number"),
        watch_type_field=find_field_by_canonical(fields, "watch_type"),
        hrv_mode=resolved_hrv, hrv_status_choices=status_choices, hrv_warning=hrv_warning,
    )


def validate_step_dag_map(
    client: RedcapClient,
    rows: pd.DataFrame,
    cols: dict[str, Any],
) -> dict[str, Any]:
    """Validate only DAG values actually used by current STEP rows before timing starts."""
    # Preserve a readable representative workbook label in the audit while using
    # canonicalized labels for routing. This makes case/whitespace variants harmless.
    used_mappings: dict[str, str] = {}
    for _, row in rows.iterrows():
        raw_team = clean_text(row.get(cols["team"]))
        dag = resolve_step_dag(raw_team)
        if raw_team and dag:
            used_mappings[raw_team] = dag
    try:
        dags = client.data_access_groups()
    except RedcapError as exc:
        print(f"  [SyncStep] WARNING: DAG preflight unavailable: {exc}")
        return {
            "status": "unavailable",
            "error": str(exc),
            "used_mappings": used_mappings,
        }

    accepted = set()
    for row in dags:
        for key in ("unique_group_name", "data_access_group_name", "group_name"):
            value = clean_text(row.get(key))
            if value:
                accepted.add(value)
    invalid = {team: dag for team, dag in used_mappings.items() if clean_text(dag) not in accepted}
    result = {
        "status": "ok" if not invalid else "invalid",
        "used_mappings": used_mappings,
        "invalid_mappings": invalid,
        "redcap_dag_values": sorted(accepted),
    }
    return result


def build_rove_import_headers(metadata: list[dict[str, Any]], record_field: str) -> list[str]:
    """Build the minimal flat CSV header used for SyncRove REDCap import.

    Rove only writes objective Garmin step observations.  Do not rebuild or
    populate the project's clinical/demographic data dictionary.  The routing
    columns identify the repeating REDCap row; the only project data fields
    written are relative day and daily steps.
    """
    return [
        record_field,
        "redcap_event_name",
        "redcap_repeat_instrument",
        "redcap_repeat_instance",
        "step_data_relative_day",
        "steps",
    ]


def expand_rove_import_row(sparse: dict[str, Any], headers: list[str]) -> dict[str, Any]:
    row = {name: "" for name in headers}
    for key, value in sparse.items():
        if key in row:
            row[key] = value
    return row


def redcap_rows_to_csv_text(records: list[dict[str, Any]], headers: list[str]) -> str:
    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=headers, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()
    for record in records:
        writer.writerow({h: sanitize_record(record).get(h, "") for h in headers})
    return buf.getvalue()


def build_rove_redcap_schema(
    client: RedcapClient,
    candidate_record_ids: Optional[list[str]] = None,
) -> RedcapSchema:
    """Validate Rove daily-step fields and determine the repeat layout safely.

    Preferred route: REDCap repeatingFormsEvents metadata when the token has
    Project Setup/Design privilege.

    Fallback route: infer the layout from ordinary flat record exports for
    existing workbook records. This requires API Export only and does not expose
    any additional clinical fields: the probe requests only record_id,
    step_data_relative_day, and steps. REDCap supplies the repeat/event metadata
    columns automatically.
    """
    metadata = client.metadata()
    fields = client.field_names()

    # The Rove import template supplied for this project is authoritative for
    # these field names.  Some REDCap roles can import/export records while the
    # metadata endpoint omits fields they do not have Project Setup/Design-level
    # visibility for.  Therefore metadata visibility is NOT used as proof that
    # the daily step fields exist.  Actual record export is probed below before
    # any timed Garmin work or REDCap write occurs.
    record_field = find_field_by_canonical(fields, "record_id") or "record_id"
    relative_day_field = find_field_by_canonical(fields, "step_data_relative_day") or "step_data_relative_day"
    steps_field = find_field_by_canonical(fields, "steps") or "steps"
    completion_field = find_field_by_canonical(fields, "step_data_complete")

    metadata_missing = [
        name for name in ("record_id", "step_data_relative_day", "steps")
        if find_field_by_canonical(fields, name) is None
    ]
    if metadata_missing:
        print(
            "  [SyncRove] Metadata endpoint does not expose: "
            + ", ".join(metadata_missing)
            + ". Using the supplied REDCap import-template field names and validating by record export instead."
        )

    # Only objective Rove fields are imported. Clinical/demographic fields are
    # deliberately absent from the CSV payload.
    import_headers = build_rove_import_headers(metadata, record_field)

    metadata_by_name = {clean_text(m.get("field_name")): m for m in metadata}
    rel_meta = metadata_by_name.get(relative_day_field, {})
    steps_meta = metadata_by_name.get(steps_field, {})
    rel_form_visible = clean_text(rel_meta.get("form_name"))
    steps_form_visible = clean_text(steps_meta.get("form_name"))
    if rel_form_visible and steps_form_visible and rel_form_visible != steps_form_visible:
        raise RedcapError(
            "SyncRove metadata places step_data_relative_day and steps on different forms: "
            f"{rel_form_visible!r} vs {steps_form_visible!r}. No records were uploaded."
        )

    # The project import/export headers establish these as the step_data repeat
    # fields. If metadata exposes the form name, use it; otherwise use the
    # instrument name demonstrated by the project structure. The fallback record
    # export below must still confirm the repeat routing before uploads begin.
    rel_form = rel_form_visible or steps_form_visible or "step_data"

    # First try the formal project-design endpoint. Some REDCap roles deny this
    # despite allowing record export/import, so a 400/privilege error is not fatal.
    privileged_error: Optional[str] = None
    try:
        repeat_cfg = client.repeating_forms_events()
        matches: list[dict[str, Any]] = []
        for item in repeat_cfg:
            form_name = clean_text(item.get("form_name") or item.get("form"))
            event_name = clean_text(item.get("event_name") or item.get("unique_event_name"))
            if canon_compact(form_name) == canon_compact(rel_form):
                matches.append({"event_name": event_name, "form_name": form_name, "repeat_is_event": False})

        if not matches:
            form_map = client.form_event_mapping()
            assigned_events = {
                clean_text(item.get("unique_event_name") or item.get("event_name"))
                for item in form_map
                if canon_compact(clean_text(item.get("form") or item.get("form_name"))) == canon_compact(rel_form)
            }
            for item in repeat_cfg:
                form_name = clean_text(item.get("form_name") or item.get("form"))
                event_name = clean_text(item.get("event_name") or item.get("unique_event_name"))
                if not form_name and event_name and event_name in assigned_events:
                    matches.append({"event_name": event_name, "form_name": "", "repeat_is_event": True})

        unique_matches: list[dict[str, Any]] = []
        seen = set()
        for item in matches:
            key = (item["event_name"], item["form_name"], item["repeat_is_event"])
            if key not in seen:
                seen.add(key)
                unique_matches.append(item)
        if len(unique_matches) == 1:
            match = unique_matches[0]
            repeat_is_event = bool(match["repeat_is_event"])
            return RedcapSchema(
                project="SyncRove",
                record_field=record_field,
                field_names=fields,
                metadata=metadata,
                mode="long_repeating",
                relative_day_field=relative_day_field,
                steps_field=steps_field,
                completion_field=completion_field,
                repeat_instrument=None if repeat_is_event else rel_form,
                repeat_event=match["event_name"] or None,
                repeat_is_event=repeat_is_event,
                import_headers=import_headers,
            )
        privileged_error = (
            f"repeatingFormsEvents did not uniquely identify {rel_form!r}: {unique_matches!r}"
        )
    except RedcapError as exc:
        privileged_error = str(exc)

    print(
        "  [SyncRove] Project-design repeat metadata is unavailable; "
        "inferring repeat layout from existing record exports instead."
    )

    # Fallback: inspect only workbook record IDs, stopping once the observed
    # layout is unambiguous. No full-project export is performed.
    ids: list[str] = []
    seen_ids = set()
    for raw in candidate_record_ids or []:
        rid = clean_text(raw)
        if rid and rid not in seen_ids:
            seen_ids.add(rid)
            ids.append(rid)

    observed: list[dict[str, Any]] = []
    probe_errors: list[str] = []
    for rid in ids:
        try:
            rows = client.probe_rove_repeat_rows(
                rid, record_field, relative_day_field, steps_field
            )
        except RedcapError as exc:
            probe_errors.append(f"{rid}: {exc}")
            continue
        for row in rows:
            rel_value = clean_text(row.get(relative_day_field))
            steps_value = clean_text(row.get(steps_field))
            repeat_instance = clean_text(row.get("redcap_repeat_instance"))
            if not (rel_value or steps_value or repeat_instance):
                continue
            observed.append(row)
        if observed:
            # One populated participant is normally sufficient; keep this
            # bounded and avoid unnecessary REDCap reads during preflight.
            break

    if not observed:
        detail = f" Prior endpoint result: {privileged_error}" if privileged_error else ""
        if probe_errors:
            detail += " Probe errors: " + " | ".join(probe_errors[:3])
        raise RedcapError(
            "SyncRove repeat layout could not be inferred because the API token lacks "
            "Project Setup/Design privilege and no existing populated step_data repeat "
            "row was found among the workbook records. No records were uploaded." + detail
        )

    instrument_rows = [
        row for row in observed
        if canon_compact(clean_text(row.get("redcap_repeat_instrument"))) == canon_compact(rel_form)
    ]
    event_rows = [
        row for row in observed
        if not clean_text(row.get("redcap_repeat_instrument"))
        and clean_text(row.get("redcap_repeat_instance"))
        and (clean_text(row.get(relative_day_field)) or clean_text(row.get(steps_field)))
    ]

    if instrument_rows:
        repeat_is_event = False
        repeat_instrument: Optional[str] = rel_form
        source_rows = instrument_rows
    elif event_rows:
        repeat_is_event = True
        repeat_instrument = None
        source_rows = event_rows
    else:
        raise RedcapError(
            "Existing SyncRove step_data rows were exported, but their repeat metadata "
            "did not identify either a repeating step_data instrument or a repeating event. "
            "No records were uploaded."
        )

    event_names = sorted({
        clean_text(row.get("redcap_event_name"))
        for row in source_rows
        if clean_text(row.get("redcap_event_name"))
    })
    if len(event_names) > 1:
        raise RedcapError(
            "SyncRove existing step_data rows span multiple REDCap events, so the upload "
            f"target is ambiguous: {event_names!r}. No records were uploaded."
        )
    repeat_event = event_names[0] if event_names else None

    print(
        "  [SyncRove] Repeat layout inferred from existing data: "
        f"instrument={repeat_instrument!r}; event={repeat_event!r}; "
        f"repeating_event={repeat_is_event}."
    )
    return RedcapSchema(
        project="SyncRove",
        record_field=record_field,
        field_names=fields,
        metadata=metadata,
        mode="long_repeating",
        relative_day_field=relative_day_field,
        steps_field=steps_field,
        completion_field=completion_field,
        repeat_instrument=repeat_instrument,
        repeat_event=repeat_event,
        repeat_is_event=repeat_is_event,
        import_headers=import_headers,
    )


def _rove_int(value: Any) -> Optional[int]:
    value = clean_text(value)
    if not value:
        return None
    try:
        return int(float(value))
    except Exception:
        return None


def build_rove_repeat_instance_map(existing_rows: list[dict[str, Any]], schema: RedcapSchema) -> tuple[dict[int, int], int]:
    by_day: dict[int, int] = {}
    used_instances: list[int] = []
    duplicates: dict[int, list[int]] = {}
    for row in existing_rows:
        instance = _rove_int(row.get("redcap_repeat_instance"))
        relative_day = _rove_int(row.get(schema.relative_day_field or "step_data_relative_day"))
        if instance is not None:
            used_instances.append(instance)
        if instance is None or relative_day is None:
            continue
        if relative_day in by_day and by_day[relative_day] != instance:
            duplicates.setdefault(relative_day, [by_day[relative_day]]).append(instance)
        else:
            by_day[relative_day] = instance
    if duplicates:
        raise RedcapError(
            "Duplicate existing SyncRove step_data relative days were found; refusing an ambiguous upload: "
            + repr(duplicates)
        )
    return by_day, max(used_instances, default=0)


def prepare_rove_redcap_rows(
    record_id: str,
    daily_rows: list[dict[str, Any]],
    existing_rows: list[dict[str, Any]],
    schema: RedcapSchema,
) -> tuple[list[dict[str, Any]], int, int]:
    if not schema.relative_day_field or not schema.steps_field:
        raise RedcapError("SyncRove schema is missing daily-step fields.")
    by_day, max_instance = build_rove_repeat_instance_map(existing_rows, schema)
    output: list[dict[str, Any]] = []
    reused = 0
    created = 0
    for daily in sorted(daily_rows, key=lambda x: int(x["relative_day"])):
        relative_day = int(daily["relative_day"])
        if relative_day in by_day:
            instance = by_day[relative_day]; reused += 1
        else:
            max_instance += 1; instance = max_instance; by_day[relative_day] = instance; created += 1
        sparse = {
            schema.record_field: record_id,
            "redcap_event_name": schema.repeat_event or "",
            "redcap_repeat_instrument": "" if schema.repeat_is_event else (schema.repeat_instrument or ""),
            "redcap_repeat_instance": instance,
            schema.relative_day_field: relative_day,
            schema.steps_field: daily.get("steps"),
        }
        output.append(expand_rove_import_row(sparse, schema.import_headers))
    return output, reused, created


# =============================================================================
# GARMIN RATE CONTROL / TOKEN SESSION MANAGER
# =============================================================================


class GarminRateLimit(RuntimeError):
    pass


@dataclass
class GarminCounters:
    api_attempts: int = 0
    api_errors: int = 0
    retry_events: int = 0
    api_active_seconds: float = 0.0
    throttle_sleep_seconds: float = 0.0
    retry_sleep_seconds: float = 0.0
    method_counts: dict[str, int] = field(default_factory=dict)

    def snapshot(self) -> dict[str, Any]:
        return {
            "api_attempts": self.api_attempts,
            "api_errors": self.api_errors,
            "retry_events": self.retry_events,
            "api_active_seconds": self.api_active_seconds,
            "throttle_sleep_seconds": self.throttle_sleep_seconds,
            "retry_sleep_seconds": self.retry_sleep_seconds,
            "method_counts": dict(self.method_counts),
        }


class GarminCaller:
    def __init__(
        self,
        qps: float,
        retry_qps: float,
        max_network_retries: int,
        jitter_seconds: float,
        rate_limit_events: list[dict[str, Any]],
        pacing_enabled: bool = True,
    ):
        self.qps = max(float(qps), 0.01)
        self.retry_qps = max(float(retry_qps), 0.01)
        self.max_network_retries = max(0, int(max_network_retries))
        self.jitter_seconds = max(0.0, float(jitter_seconds))
        self.rate_limit_events = rate_limit_events
        self.pacing_enabled = bool(pacing_enabled)
        self.last_call_monotonic = 0.0
        self.counters = GarminCounters()

    def _throttle(self, retry_mode: bool = False) -> None:
        # Pacing controls only intentional inter-call QPS spacing. Network retry
        # backoff remains active even when pacing is disabled.
        if not self.pacing_enabled:
            self.last_call_monotonic = time.monotonic()
            return
        qps = self.retry_qps if retry_mode else self.qps
        min_interval = 1.0 / qps
        now = time.monotonic()
        wait = (self.last_call_monotonic + min_interval) - now
        if wait > 0:
            wait += random.uniform(0.0, self.jitter_seconds)
            t0 = time.perf_counter()
            time.sleep(wait)
            self.counters.throttle_sleep_seconds += time.perf_counter() - t0
        self.last_call_monotonic = time.monotonic()

    def _retry_sleep(self, seconds: float) -> None:
        delay = max(0.0, seconds) + random.uniform(0.0, self.jitter_seconds)
        t0 = time.perf_counter()
        time.sleep(delay)
        self.counters.retry_sleep_seconds += time.perf_counter() - t0

    def call(
        self,
        fn: Callable,
        *args,
        project: str,
        record_id: str,
        watch_category: str,
        date_value: Optional[str] = None,
        **kwargs,
    ) -> Any:
        method_name = getattr(fn, "__name__", fn.__class__.__name__)
        attempt = 0
        while True:
            attempt += 1
            self._throttle(retry_mode=attempt > 1)
            self.counters.api_attempts += 1
            self.counters.method_counts[method_name] = self.counters.method_counts.get(method_name, 0) + 1
            t0 = time.perf_counter()
            try:
                result = fn(*args, **kwargs)
                self.counters.api_active_seconds += time.perf_counter() - t0
                return result
            except Exception as exc:
                # Only time spent inside the actual Garmin method counts as API-active
                # time. Throttle and retry sleeps are tracked separately.
                self.counters.api_active_seconds += time.perf_counter() - t0
                self.counters.api_errors += 1
                message = str(exc)
                if is_rate_limit_error(message):
                    self.rate_limit_events.append({
                        "timestamp": now_local_iso(),
                        "project": project,
                        "record_id": record_id,
                        "watch_category": watch_category,
                        "date": date_value,
                        "method": method_name,
                        "kind": "DATA_429",
                        "message": message,
                    })
                    raise GarminRateLimit(message) from exc
                if is_transient_error(message) and attempt <= self.max_network_retries:
                    self.counters.retry_events += 1
                    self._retry_sleep(min(2 ** (attempt - 1), 30))
                    continue
                raise


class GarminSessionManager:
    def __init__(self, rate_limit_events: list[dict[str, Any]]):
        self.clients: dict[tuple[str, str], Garmin] = {}
        self.rate_limit_events = rate_limit_events

    def get_client(
        self,
        project: str,
        username: str,
        token_dir: Path,
        record_id: str,
        watch_category: str,
    ) -> tuple[Garmin, str]:
        username = clean_text(username)
        key = (project, username.lower())
        if key in self.clients:
            return self.clients[key], "memory-cache"

        path = token_path(token_dir, username)
        if not token_file_structurally_valid(path):
            raise FileNotFoundError(
                f"No structurally usable seeded token for account hash {account_hash(username)} "
                f"in {token_dir}"
            )

        try:
            client = Garmin(retry_attempts=2, retry_min_wait=1.0, retry_max_wait=5.0)
            client.login(str(path))
        except GarminConnectTooManyRequestsError as exc:
            self.rate_limit_events.append({
                "timestamp": now_local_iso(),
                "project": project,
                "record_id": record_id,
                "watch_category": watch_category,
                "date": None,
                "method": "login(token)",
                "kind": "LOGIN_429",
                "message": str(exc),
            })
            raise GarminRateLimit(f"Token restore/refresh rate limited: {exc}") from exc
        except Exception as exc:
            if is_rate_limit_error(str(exc)):
                self.rate_limit_events.append({
                    "timestamp": now_local_iso(),
                    "project": project,
                    "record_id": record_id,
                    "watch_category": watch_category,
                    "date": None,
                    "method": "login(token)",
                    "kind": "LOGIN_429",
                    "message": str(exc),
                })
                raise GarminRateLimit(f"Token restore/refresh rate limited: {exc}") from exc
            raise RuntimeError(f"Saved token authentication failed: {exc}") from exc

        self.clients[key] = client
        return client, "seeded-token"

    def clear(self) -> None:
        self.clients.clear()


# =============================================================================
# STEP WORKBOOK PARSING
# =============================================================================


def build_lookup(columns: Iterable[Any]) -> dict[str, Any]:
    return {canon(col): col for col in columns if canon(col)}


def resolve_header(
    lookup: dict[str, Any],
    *candidates: str,
    required: bool = True,
) -> Optional[Any]:
    for candidate in candidates:
        c = canon(candidate)
        if c in lookup:
            return lookup[c]
    for candidate in candidates:
        c = canon(candidate)
        for key, original in lookup.items():
            if c and (c == key or c in key):
                return original
    if required:
        raise KeyError(f"Could not find any of these columns: {candidates}")
    return None


def autodetect_step_header_row(
    workbook: Path,
    sheet_name: str,
    max_scan_rows: int = 15,
) -> tuple[Optional[int], pd.DataFrame]:
    raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None, dtype=object)
    expected = {
        "redcap id", "watch email", "watch password", "watch model",
        "watch serial number", "redcap team", "patient enrollment date",
        "patient surgery date",
    }
    best_row = None
    best_hits = -1
    for i in range(min(max_scan_rows, len(raw))):
        values = {canon(v) for v in raw.iloc[i].tolist() if canon(v)}
        hits = sum(1 for e in expected if any(e == v or e in v for v in values))
        if hits > best_hits:
            best_hits = hits
            best_row = i
        if hits >= 5:
            return i, raw
    if best_row is not None and best_hits >= 3:
        return best_row, raw
    return None, raw


def load_step_rows(workbook: Path, sheet_name: str, process_all: bool) -> tuple[int, pd.DataFrame, dict[str, Any]]:
    header_row, raw = autodetect_step_header_row(workbook, sheet_name)
    if header_row is None:
        raise RuntimeError(f"Could not identify SyncStep header row in sheet {sheet_name!r}")

    df = raw.copy()
    df.columns = df.iloc[header_row].tolist()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df["__excel_row__"] = [header_row + 2 + i for i in range(len(df))]

    lk = build_lookup(df.columns)
    cols = {
        "record_id": resolve_header(lk, "REDCap ID", "redcap id", "record id"),
        "username": resolve_header(lk, "Watch Email", "account email", "email"),
        "password": resolve_header(lk, "Watch Password", "account password", "password", required=False),
        "model": resolve_header(lk, "Watch Model", "garmin model"),
        "serial": resolve_header(lk, "Watch Serial Number", "garmin serial no", "serial"),
        "team": resolve_header(lk, "REDCap Team", "step team using tracker", "team"),
        "start": resolve_header(lk, "Patient Enrollment Date", "enrollment date", "start date"),
        "dos": resolve_header(lk, "Patient Surgery Date", "day of surgery", "dos"),
        "end": resolve_header(lk, "Patient Completion Date", "tracking completion date", "end date"),
        "returned": exact_optional_header(lk, "Device back", "Device back date", "Watch returned", "Watch Return Date", "Return Date", "Date Returned"),
        "verified": resolve_header(
            lk,
            "REDCap Upload Completed and Verified",
            "upload completed and verified",
            "verified",
            required=False,
        ),
    }

    for key in ("record_id", "username", "model", "serial", "team"):
        df[cols[key]] = df[cols[key]].map(clean_text)
    # Keep raw boundaries: an invalid/partial date must not silently become a
    # blank and enable a longer fallback window.
    for key in ("start", "dos", "end"):
        df[f"__raw_{key}__"] = df[cols[key]]

    df = df[
        df[cols["record_id"]].ne("")
        | df[cols["username"]].ne("")
        | df[cols["serial"]].ne("")
    ].copy()

    preserve_return_precision(workbook, sheet_name, header_row, df, cols)
    cols["_assignment_rows"] = df.copy(deep=True)
    if not process_all and cols["verified"] is not None:
        df = df[~df[cols["verified"]].map(is_truthy)].copy()

    df["__dedupe_key__"] = (
        df[cols["username"]].str.lower().str.strip()
        + "||"
        + df[cols["record_id"]].str.lower().str.strip()
    )
    df = df.drop_duplicates(subset="__dedupe_key__", keep="last").drop(columns="__dedupe_key__")
    return header_row, df, cols


def classify_step_watch(model: Any) -> tuple[bool, int, str]:
    raw = clean_text(model)
    low = raw.lower()
    is_vivoactive = "vivoactive" in low or "vívoactive" in raw
    if is_vivoactive:
        label = "vívoactive 5" if "5" in raw else (raw or "vívoactive")
        return True, 2, label
    if "vivofit" in low or "vívofit" in raw:
        label = "vívofit 4" if "4" in raw else (raw or "vívofit")
        return False, 1, label
    return False, 1, raw or "Type 1 / model not specified"


# =============================================================================
# SYNCROVE WORKBOOK PARSING
# =============================================================================


def autodetect_rove_header_row(
    workbook: Path,
    sheet_name: str,
    max_scan_rows: int = 20,
) -> tuple[Optional[int], pd.DataFrame]:
    raw = pd.read_excel(workbook, sheet_name=sheet_name, header=None, dtype=object)
    expected = {
        "username", "password", "study id", "vivofit4 serial number",
        "dos", "start date", "expected end date",
    }
    best_row = None
    best_hits = -1
    for i in range(min(max_scan_rows, len(raw))):
        values = {canon(v) for v in raw.iloc[i].tolist() if canon(v)}
        hits = len(expected & values)
        if hits > best_hits:
            best_row, best_hits = i, hits
        if hits >= 5:
            return i, raw
    if best_row is not None and best_hits >= 3:
        return best_row, raw
    return None, raw


def load_rove_rows(workbook: Path, sheet_name: str, process_all: bool) -> tuple[int, pd.DataFrame, dict[str, Any]]:
    header_row, raw = autodetect_rove_header_row(workbook, sheet_name)
    if header_row is None:
        raise RuntimeError(f"Could not identify SyncRove header row in sheet {sheet_name!r}")

    df = raw.copy()
    df.columns = df.iloc[header_row].tolist()
    df = df.iloc[header_row + 1:].reset_index(drop=True)
    df["__excel_row__"] = [header_row + 2 + i for i in range(len(df))]

    lk = build_lookup(df.columns)
    cols = {
        "username": resolve_header(lk, "username", "Watch Email", "account email", "email"),
        "password": resolve_header(lk, "password", "Watch Password", required=False),
        "study_id": resolve_header(lk, "Study ID", "record id", "REDCap ID"),
        "serial": resolve_header(
            lk, "Vivofit4 Serial Number", "Watch Serial Number", "serial number", "serial"
        ),
        "dos": resolve_header(lk, "DOS", "Patient Surgery Date", "day of surgery"),
        "start": resolve_header(lk, "Start Date", "Patient Enrollment Date", "enrollment date"),
        "end": resolve_header(
            lk, "Expected End Date", "Patient Completion Date", "completion date", "end date", required=False
        ),
        "returned": exact_optional_header(lk, "Device back", "Device back date", "Watch returned", "Watch Return Date", "Return Date", "Date Returned"),
        "captured": resolve_header(lk, "Walking data captured", "walking captured", required=False),
    }

    for key in ("username", "study_id", "serial"):
        df[cols[key]] = df[cols[key]].map(clean_text)
    for key in ("dos", "start", "end"):
        df[f"__raw_{key}__"] = df[cols[key]] if cols.get(key) else None

    df = df[
        df[cols["username"]].ne("")
        | df[cols["study_id"]].ne("")
        | df[cols["serial"]].ne("")
    ].copy()

    preserve_return_precision(workbook, sheet_name, header_row, df, cols)
    cols["_assignment_rows"] = df.copy(deep=True)
    if not process_all and cols["captured"] is not None:
        df = df[~df[cols["captured"]].map(is_truthy)].copy()

    df["__dedupe_key__"] = (
        df[cols["username"]].str.lower().str.strip()
        + "||"
        + df[cols["study_id"]].str.lower().str.strip()
    )
    df = df.drop_duplicates(subset="__dedupe_key__", keep="last").drop(columns="__dedupe_key__")
    return header_row, df, cols


# =============================================================================
# STEP DAILY DATA EXTRACTION
# =============================================================================


def exact_optional_header(lookup: dict[str, Any], *candidates: str) -> Optional[Any]:
    """Actual-return columns use exact aliases, never an expected-return match."""
    return next((lookup[canon(c)] for c in candidates if canon(c) in lookup), None)


def preserve_return_precision(workbook: Path, sheet: str, header_row: int,
                              rows: pd.DataFrame, cols: dict[str, Any]) -> None:
    """A month-only Excel display is kept month-only instead of inventing a day."""
    if not cols.get("returned"):
        return
    wb = load_workbook(workbook, read_only=True, data_only=True)
    try:
        ws = wb[sheet]
        col = next((c.column for c in ws[header_row + 1]
                    if canon(c.value) == canon(cols["returned"])), None)
        if col is None:
            return
        for index, row in rows.iterrows():
            cell = ws.cell(int(row["__excel_row__"]), col)
            value = cell.value
            # Read-only Excel blank cells can have number_format=None. Leave
            # blanks and text returns untouched; only native dates need a
            # display-format check to preserve month-only precision.
            if not isinstance(value, (date, datetime)):
                continue
            fmt = re.sub(r'"[^"]*"|\[[^\]]*\]', '', (cell.number_format or '').lower())
            if 'm' in fmt and 'y' in fmt and 'd' not in fmt:
                rows.at[index, cols["returned"]] = value.strftime('%b-%Y')
    finally:
        wb.close()


def strict_day(value: Any) -> Optional[date]:
    if not clean_text(value):
        return None
    if isinstance(value, (int, float, bool)):
        raise ValueError('Date is an unformatted number, not an exact calendar date')
    text = clean_text(value)
    if not isinstance(value, (date, datetime)):
        parts = re.findall(r'\d+|[A-Za-z]+', text)
        if len(parts) < 3 or re.fullmatch(r'(?:\d{4}-\d{1,2}|\d{1,2}/\d{4}|[A-Za-z]{3,9}[- /]\d{2,4})', text):
            raise ValueError(f'Exact day missing: {text}')
    day = safe_date(value)
    if day is None:
        raise ValueError(f'Invalid calendar date: {text}')
    return day


def return_boundary(value: Any) -> tuple[Optional[date], Optional[date], str]:
    """Return (last safe date, exact return if known, precision)."""
    if not clean_text(value):
        return None, None, 'not recorded'
    text = clean_text(value)
    if not isinstance(value, (date, datetime)) and re.fullmatch(
        r'(?:\d{4}-\d{1,2}|\d{1,2}/\d{4}|[A-Za-z]{3,9}[- /]\d{2,4})', text
    ):
        month = None
        for fmt in ('%Y-%m', '%m/%Y', '%b-%y', '%b-%Y', '%B-%y', '%B-%Y', '%b %y', '%b %Y', '%B %y', '%B %Y', '%b/%y', '%b/%Y', '%B/%y', '%B/%Y'):
            try:
                month = datetime.strptime(text, fmt)
                break
            except ValueError:
                continue
        if month is None:
            raise ValueError(f'Invalid return month: {text}')
        first = date(month.year, month.month, 1)
        return first - timedelta(days=1), None, 'month only; stop before earliest possible return'
    exact = strict_day(value)
    return exact, exact, 'exact date'


def normalized_serial(value: Any) -> str:
    text = canon_compact(clean_text(value))
    return '' if text in {'', '0', 'unknown', 'tbd', 'na', 'none', 'missing'} else text


def annotate_assignment_windows(project_frames: list[tuple[str, pd.DataFrame, dict[str, Any]]],
                                today: Optional[date] = None) -> list[dict[str, Any]]:
    """Join full logs by serial OR account across projects before skip/dedupe.

    Start Date means the logged assignment start. These limits are protection
    against documented reuse, not proof that unlogged reuse never happened.
    """
    today = today or date.today()
    assignments = []
    for project, processed, cols in project_frames:
        all_rows = cols.get('_assignment_rows', processed)
        id_key = 'record_id' if project == 'SyncStep' else 'study_id'
        for _, row in all_rows.iterrows():
            item = {'project': project, 'excel_row': int(row['__excel_row__']),
                    'record_id': clean_text(row.get(cols[id_key])),
                    'serial': normalized_serial(row.get(cols.get('serial'))),
                    'account': clean_text(row.get(cols['username'])).lower().strip(),
                    'start': None, 'planned': None, 'return_limit': None,
                    'returned': None, 'return_precision': 'not recorded',
                    'next_start': None, 'errors': [], 'warnings': []}
            item['inventory_only'] = not item['record_id'] and not any(
                clean_text(row.get(f'__raw_{k}__', row.get(cols.get(k)))) for k in ('start', 'dos', 'end')) and not clean_text(row.get(cols.get('returned')))
            for key in ('start', 'dos', 'end'):
                raw = row.get(f'__raw_{key}__', row.get(cols.get(key)))
                try:
                    item[key] = strict_day(raw)
                except ValueError as exc:
                    item[key] = None
                    item['errors'].append(f'{key}: {exc}')
            try:
                item['return_limit'], item['returned'], item['return_precision'] = return_boundary(
                    row.get(cols.get('returned')))
            except ValueError as exc:
                item['errors'].append(f'Device back: {exc}')
            item['return_latest'] = item['returned']
            if item['return_limit'] and item['return_precision'].startswith('month only'):
                first = item['return_limit'] + timedelta(days=1)
                following = date(first.year + 1, 1, 1) if first.month == 12 else date(first.year, first.month + 1, 1)
                item['return_latest'] = following - timedelta(days=1)
            if not item['start'] and not item['inventory_only']:
                item['errors'].append('Missing assignment Start Date')
            item['planned'] = item['end']
            if not item['planned'] and project == 'SyncRove' and item['dos']:
                item['planned'] = item['dos'] + timedelta(days=ROVE_DEFAULT_POSTOP_DAYS)
                item['warnings'].append('Expected End Date missing; DOS + 90 fallback used')
            if item['start'] and item['return_limit'] and item['return_limit'] < item['start']:
                item['errors'].append('Return boundary precedes Start Date')
            if item['start'] and item['planned'] and item['planned'] < item['start']:
                item['errors'].append('Expected end precedes Start Date')
            if not item['return_limit']:
                item['warnings'].append('Return date missing; possession end is unconfirmed')
                if not item['serial']:
                    item['errors'].append('Missing return and serial: physical-device reuse cannot be checked')
            elif item['return_precision'] != 'exact date':
                item['warnings'].append('Month-only return: exclude entire return month conservatively')
            assignments.append(item)

    for i, item in enumerate(assignments):
        for other in assignments[i + 1:]:
            if item['inventory_only'] or other['inventory_only']:
                continue
            same_serial = bool(item['serial'] and item['serial'] == other['serial'])
            same_account = bool(item['account'] and item['account'] == other['account'])
            same_record = bool(item['project'] == other['project'] and item['record_id']
                               and item['record_id'] == other['record_id'])
            if not (same_serial or same_account or same_record):
                continue
            exact_duplicate = same_record and all(item[k] == other[k] for k in
                ('serial', 'account', 'start', 'dos', 'planned', 'return_limit'))
            if exact_duplicate:
                continue
            if same_record:
                for x in (item, other):
                    x['errors'].append('Conflicting duplicate participant assignments; consolidate the log')
            if not (same_serial or same_account):
                continue
            if not item['start'] or not other['start']:
                missing, known = (item, other) if not item['start'] else (other, item)
                if known['start'] and missing['return_latest'] and missing['return_latest'] <= known['start']:
                    known['warnings'].append('Undated earlier assignment was returned before this Start Date')
                    continue
                for x in (item, other):
                    x['errors'].append('Reused serial/account has an undated assignment; ownership order unknown')
                continue
            if item['start'] == other['start']:
                for x in (item, other):
                    x['errors'].append('Same serial/account assigned to different records on the same day')
                continue
            old, new = sorted((item, other), key=lambda x: x['start'])
            if old['next_start'] is None or new['start'] < old['next_start']:
                old['next_start'] = new['start']
            # A same-day handover is resolved by excluding that day for the old
            # assignment. A return strictly later contradicts the new assignment.
            if old['returned'] and old['returned'] > new['start']:
                for x in (old, new):
                    x['errors'].append('Recorded return overlaps a later serial/account assignment')
            elif old['return_precision'].startswith('month only') and old['return_latest'] > new['start']:
                for x in (old, new):
                    x['errors'].append('Month-only return overlaps possible reassignment; exact date needed')

    lookup, audit = {}, []
    for item in assignments:
        candidates = [(item['planned'], 'expected end' if item['end'] else 'DOS + 90 fallback'),
                      (item['return_limit'], 'recorded return')]
        if item['dos']:
            maximum = STEP_POSTOP_DAYS if item['project'] == 'SyncStep' else ROVE_POSTOP_DAYS
            candidates.append((item['dos'] + timedelta(days=maximum), 'REDCap maximum relative day'))
            if item['project'] == 'SyncRove' and item['start'] and (item['start'] - item['dos']).days < ROVE_MIN_RELATIVE_DAY:
                item['errors'].append('Start Date is outside the supported Rove preop relative-day range')
        if item['next_start']:
            candidates.append((item['next_start'] - timedelta(days=1), 'next assignment minus one day'))
            item['warnings'].append('Serial/account reused; prior assignment ends before next Start Date')
        candidates = [(d, reason) for d, reason in candidates if d is not None]
        final_end = min((d for d, _ in candidates), default=None)
        end = min(final_end, today) if final_end else None
        if not final_end:
            item['errors'].append('No valid monitoring end boundary')
        if item['start'] and end and end < item['start']:
            item['errors'].append('Assignment has no elapsed eligible days')
        record = {'project': item['project'], 'record_id': item['record_id'],
                  'excel_row': item['excel_row'], 'serial': item['serial'],
                  'account_hash': account_hash(item['account']) if item['account'] else '',
                  'query_start': item['start'], 'planned_end': item['planned'],
                  'recorded_return': item['returned'], 'return_precision': item['return_precision'],
                  'next_assignment_start': item['next_start'], 'final_end': final_end,
                  'query_end': end, 'limit_reason': ', '.join(reason for d, reason in candidates if d == final_end),
                  'warnings': '; '.join(dict.fromkeys(item['warnings'])),
                  'error': '; '.join(dict.fromkeys(item['errors']))}
        audit.append(record)
        lookup[(item['project'], item['excel_row'])] = record
    for project, rows, _ in project_frames:
        for index, row in rows.iterrows():
            record = lookup[(project, int(row['__excel_row__']))]
            for key in ('query_end', 'final_end', 'limit_reason', 'warnings', 'error'):
                rows.at[index, f'__assignment_{key}__'] = record[key]
    return audit


def assignment_window(row: pd.Series) -> tuple[date, date]:
    """Fail closed if a caller bypasses the full-log preflight."""
    error = clean_text(row.get('__assignment_error__'))
    if error:
        raise ValueError(error)
    end = safe_date(row.get('__assignment_query_end__'))
    final = safe_date(row.get('__assignment_final_end__'))
    if end is None or final is None:
        raise ValueError('Ownership preflight is required before account processing')
    return end, final


def finite_nonnegative(value: Any, metric: str) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool) or isinstance(value, (dict, list)):
        raise ValueError(f'{metric} is not numeric')
    try:
        numeric = float(value)
    except (ValueError, TypeError) as exc:
        raise ValueError(f'{metric} is not numeric') from exc
    if not math.isfinite(numeric) or numeric < 0:
        raise ValueError(f'{metric} is not finite/nonnegative')
    # Keep integral measurements as integers for REDCap integer validation;
    # preserve genuine fractional precision instead of rounding it away.
    return int(numeric) if numeric.is_integer() else numeric


def extract_hrv(sleep: dict[str, Any], hrv: Any, cdate: str) -> dict[str, Any]:
    if hrv is not None and not isinstance(hrv, dict):
        raise ValueError('Malformed HRV response; do not treat retrieval failure as missing')
    dto = sleep.get('dailySleepDTO') or {}
    summary = (hrv or {}).get('hrvSummary') or {}
    if not isinstance(dto, dict) or not isinstance(summary, dict):
        raise ValueError('Malformed sleep/HRV summary')
    for payload in (dto, summary):
        if payload.get('calendarDate') and clean_text(payload['calendarDate']) != cdate:
            raise ValueError('Sleep/HRV response date differs from requested date')
    sleep_value = finite_nonnegative(dto.get('avgSleepHRV'), 'overnight HRV')
    endpoint_value = finite_nonnegative(summary.get('lastNightAvg'), 'overnight HRV')
    value = sleep_value if sleep_value is not None else endpoint_value
    source = 'sleep.dailySleepDTO.avgSleepHRV' if sleep_value is not None else (
        'hrv.hrvSummary.lastNightAvg' if endpoint_value is not None else 'successful response; no HRV value')
    return {'hrv_overnight_ms': value, 'hrv_status': clean_text(summary.get('status')) or None,
            'hrv_source': source, 'hrv_sources_disagree': bool(sleep_value is not None and endpoint_value is not None and sleep_value != endpoint_value)}


def extract_intensity(stats: dict[str, Any], summary: dict[str, Any]) -> dict[str, Any]:
    moderate = finite_nonnegative(actual_from(stats, summary, 'moderateIntensityMinutes'), 'moderate intensity')
    vigorous = finite_nonnegative(actual_from(stats, summary, 'vigorousIntensityMinutes'), 'vigorous intensity')
    total = moderate + 2 * vigorous if moderate is not None and vigorous is not None else None
    return {'modintensity': moderate, 'vigintensity': vigorous, 'totalintensity': total,
            'totalintensity_source': 'moderate + 2 * vigorous' if total is not None else 'missing component; no total derived'}


def hrv_status_key(value: Any) -> str:
    return re.sub(r'[^a-z0-9]', '', clean_text(value).lower())


def resolve_hrv_schema(metadata: list[dict[str, Any]], requested: str) -> tuple[str, dict[str, dict[str, str]], str]:
    """Determine the actual HRV concept without putting milliseconds in status fields."""
    fields = [m for m in metadata if re.fullmatch(r'(preop|postop)hrv_day\d+', str(m.get('field_name', '')))]
    if requested == 'off':
        return 'off', {}, 'HRV upload disabled; raw numeric/status data are still exported'
    modes, mappings = set(), {}
    for m in fields:
        kind = clean_text(m.get('field_type')).lower()
        label = clean_text(m.get('field_label')).lower()
        validation = clean_text(m.get('text_validation_type_or_show_slider_number')).lower()
        choices = {}
        for choice in clean_text(m.get('select_choices_or_calculations')).split('|'):
            if ',' in choice:
                code, text = choice.split(',', 1)
                choices[hrv_status_key(text)] = code.strip()
        name = m['field_name']
        if kind in {'radio', 'dropdown'}:
            concept = 'status'
            if not choices or not any(k in choices for k in ('balanced', 'unbalanced', 'low', 'poor')):
                concept = 'ambiguous'
            mappings[name] = choices
        elif kind == 'text' and 'status' in label and not validation:
            concept = 'status'
            mappings[name] = {}
        elif kind == 'text' and 'status' not in label and (validation.startswith(('number', 'integer')) or 'overnight' in label or 'millisecond' in label or '(ms)' in label):
            concept = 'overnight-ms'
        else:
            concept = 'ambiguous'
        # An explicit numeric override still requires a numeric text field.
        if requested == 'overnight-ms' and kind == 'text' and validation.startswith(('number', 'integer')):
            concept = 'overnight-ms'
        modes.add(concept)
    if requested != 'auto' and modes != {requested}:
        raise RedcapError(f'HRV mode {requested!r} conflicts with REDCap field types/choices; no upload attempted')
    if len(modes) != 1 or 'ambiguous' in modes or not fields:
        return 'off', {}, 'HRV metadata is ambiguous/mixed. HRV fields are omitted (existing values preserved); configure numeric overnight HRV in ms or explicit status choices.'
    mode = modes.pop()
    return mode, mappings, ''


def step_metric_value(metric: str, values: dict[str, Any], target_field: str,
                      schema: Optional[RedcapSchema]) -> Any:
    if metric != 'hrv':
        return values.get(metric)
    if schema is None or schema.hrv_mode == 'overnight-ms':
        return values.get('hrv_overnight_ms')
    if schema.hrv_mode == 'off':
        return None
    status = values.get('hrv_status')
    if status is None:
        return None
    choices = schema.hrv_status_choices.get(target_field, {})
    if not choices:
        return status
    key = hrv_status_key(status)
    if key not in choices:
        raise RedcapError(f'HRV status {status!r} has no choice mapping for {target_field}; no upload attempted')
    return choices[key]


def extract_actual(payload: dict[str, Any], key: str) -> Any:
    if not isinstance(payload, dict):
        return None
    value = payload.get(key)
    if isinstance(value, dict):
        if value.get("value") is not None:
            return value.get("value")
        if value.get("total") is not None:
            return value.get("total")
        return value.get("summary")
    return value


def actual_from(stats: dict[str, Any], summary: dict[str, Any], *keys: str) -> Any:
    for key in keys:
        value = extract_actual(stats, key)
        if value is not None:
            return value
        value = extract_actual(summary, key)
        if value is not None:
            return value
    return None


def as_minutes(seconds: Any) -> Optional[int]:
    if seconds is None:
        return None
    try:
        return round(float(seconds) / 60.0)
    except Exception:
        return None


def has_any_step_metric(row: dict[str, Any]) -> bool:
    """True when Garmin returned at least one STEP wearable metric for the date."""
    return any(row.get(key) is not None for key in (*STEP_FIELDS, 'hrv_overnight_ms', 'hrv_status'))


def step_availability_pattern(row: dict[str, Any]) -> dict[str, Any]:
    """Classify independent step/sleep availability without altering source values."""
    steps_present = row.get("steps") is not None
    sleep_present = any(row.get(key) is not None for key in STEP_SLEEP_FIELDS)
    if steps_present and sleep_present:
        pattern = "steps_and_sleep"
    elif steps_present:
        pattern = "steps_only"
    elif sleep_present:
        pattern = "sleep_only"
    else:
        pattern = "no_steps_or_sleep"
    return {
        "steps_present": steps_present,
        "sleep_present": sleep_present,
        "any_metric_present": has_any_step_metric(row),
        "data_pattern": pattern,
    }


def redcap_values_equivalent(expected: Any, observed: Any) -> bool:
    """Compare a submitted value with REDCap's string-oriented readback value."""
    exp_blank = expected is None or clean_text(expected) == ""
    obs_blank = observed is None or clean_text(observed) == ""
    if exp_blank or obs_blank:
        return exp_blank and obs_blank
    try:
        exp_num = float(expected)
        obs_num = float(observed)
        if math.isfinite(exp_num) and math.isfinite(obs_num):
            return math.isclose(exp_num, obs_num, rel_tol=1e-9, abs_tol=1e-9)
    except Exception:
        pass
    return clean_text(expected) == clean_text(observed)


def verify_step_readback(
    payload: dict[str, Any],
    readback: dict[str, Any],
    schema: RedcapSchema,
) -> list[str]:
    mismatches: list[str] = []
    for field_name, expected in payload.items():
        if field_name == "redcap_data_access_group":
            continue
        if field_name != schema.record_field and field_name not in schema.field_names:
            continue
        observed = readback.get(field_name)
        if not redcap_values_equivalent(expected, observed):
            mismatches.append(
                f"{field_name}: sent={expected!r}, readback={observed!r}"
            )
    return mismatches


def verify_rove_readback(
    submitted_rows: list[dict[str, Any]],
    readback_rows: list[dict[str, Any]],
    schema: RedcapSchema,
) -> list[str]:
    by_key: dict[tuple[int, int], dict[str, Any]] = {}
    duplicates: list[tuple[int, int]] = []
    for row in readback_rows:
        rel = _rove_int(row.get(schema.relative_day_field or "step_data_relative_day"))
        inst = _rove_int(row.get("redcap_repeat_instance"))
        if rel is None or inst is None:
            continue
        key = (rel, inst)
        if key in by_key:
            duplicates.append(key)
        by_key[key] = row
    mismatches: list[str] = []
    if duplicates:
        mismatches.append(f"duplicate readback relative-day/instance keys: {duplicates[:10]!r}")
    for row in submitted_rows:
        rel = _rove_int(row.get(schema.relative_day_field or "step_data_relative_day"))
        inst = _rove_int(row.get("redcap_repeat_instance"))
        if rel is None or inst is None:
            mismatches.append(f"submitted row missing relative day/instance: {row!r}")
            continue
        observed = by_key.get((rel, inst))
        if observed is None:
            mismatches.append(f"relative_day={rel}, instance={inst}: row missing on readback")
            continue
        expected_steps = row.get(schema.steps_field or "steps")
        observed_steps = observed.get(schema.steps_field or "steps")
        if not redcap_values_equivalent(expected_steps, observed_steps):
            mismatches.append(
                f"relative_day={rel}, instance={inst}: steps sent={expected_steps!r}, "
                f"readback={observed_steps!r}"
            )
    return mismatches


def resolve_step_stats_mode(requested: str) -> tuple[str, dict[str, Any]]:
    requested = str(requested or "auto").lower()
    if requested not in {"auto", "optimized", "legacy"}:
        raise ValueError("stats mode must be auto, optimized, or legacy")

    version = None
    try:
        version = importlib.metadata.version("garminconnect")
    except Exception:
        try:
            version = importlib.metadata.version("python-garminconnect")
        except Exception:
            pass

    alias = None
    try:
        src = inspect.getsource(Garmin.get_stats)
        alias = "self.get_user_summary" in re.sub(r"\s+", " ", src)
    except Exception:
        alias = None

    if requested == "auto":
        effective = "optimized" if alias is True else "legacy"
    else:
        effective = requested
    return effective, {
        "garminconnect_version": version,
        "get_stats_aliases_get_user_summary": alias,
    }


def get_step_daily_summary_sources(
    client: Garmin,
    cdate: str,
    is_vivoactive: bool,
    stats_mode: str,
    caller: GarminCaller,
    project: str,
    record_id: str,
    watch_category: str,
) -> tuple[dict[str, Any], dict[str, Any]]:
    if stats_mode == "optimized":
        if is_vivoactive:
            stats = caller.call(
                client.get_stats,
                cdate,
                project=project,
                record_id=record_id,
                watch_category=watch_category,
                date_value=cdate,
            ) or {}
            return stats, stats
        summary = caller.call(
            client.get_user_summary,
            cdate,
            project=project,
            record_id=record_id,
            watch_category=watch_category,
            date_value=cdate,
        ) or {}
        return summary, summary

    stats = {}
    if is_vivoactive:
        stats = caller.call(
            client.get_stats,
            cdate,
            project=project,
            record_id=record_id,
            watch_category=watch_category,
            date_value=cdate,
        ) or {}
    summary = caller.call(
        client.get_user_summary,
        cdate,
        project=project,
        record_id=record_id,
        watch_category=watch_category,
        date_value=cdate,
    ) or {}
    return stats, summary


# =============================================================================
# PERFORMANCE RECORDER
# =============================================================================


@dataclass
class CombinedRecorder:
    run_id: str
    attempts: list[dict[str, Any]] = field(default_factory=list)
    download_iterations: list[dict[str, Any]] = field(default_factory=list)
    rate_limit_events: list[dict[str, Any]] = field(default_factory=list)

    def add_attempt(self, row: dict[str, Any]) -> None:
        self.attempts.append(row)

    def add_iteration(self, row: dict[str, Any]) -> None:
        self.download_iterations.append(row)


# =============================================================================
# PROJECT RESULT CONTAINERS
# =============================================================================


@dataclass
class ProjectResult:
    project: str
    success: bool
    retryable: bool
    record_id: str
    excel_row: int
    watch_category: str
    payload: Optional[dict[str, Any]] = None
    # For long/repeating REDCap projects such as SyncRove, keep the actual
    # imported REDCap rows separately from the participant-level summary payload.
    redcap_rows: list[dict[str, Any]] = field(default_factory=list)
    daily_rows: list[dict[str, Any]] = field(default_factory=list)
    sync_row: Optional[dict[str, Any]] = None
    writeback: Optional[dict[str, Any]] = None
    error: Optional[str] = None
    metrics: Optional[dict[str, Any]] = None
    boundary_rows: list[dict[str, Any]] = field(default_factory=list)


# =============================================================================
# STEP ACCOUNT PROCESSOR
# =============================================================================


def step_validation_reasons(row: pd.Series, cols: dict[str, Any]) -> list[str]:
    reasons = []
    try:
        assignment_window(row)
    except ValueError as exc:
        reasons.append(str(exc))
    if not clean_text(row.get(cols["record_id"])):
        reasons.append("Missing REDCap ID")
    if not clean_text(row.get(cols["username"])):
        reasons.append("Missing Watch Email")
    if not clean_text(row.get(cols["serial"])):
        reasons.append("Missing Watch Serial Number")
    if safe_date(row.get(cols["start"])) is None:
        reasons.append("Missing/invalid Patient Enrollment Date")
    if safe_date(row.get(cols["dos"])) is None:
        reasons.append("Missing/invalid Patient Surgery Date")
    if safe_date(row.get(cols["end"])) is None:
        reasons.append("Missing/invalid Patient Completion Date")
    team = clean_text(row.get(cols["team"]))
    if not team:
        reasons.append("Missing REDCap Team")
    elif resolve_step_dag(team) is None:
        if normalize_step_team(team) == "unk":
            reasons.append("Unresolved REDCap Team 'UNK' (must be corrected before upload)")
        else:
            reasons.append(f"Unknown REDCap Team {team!r}")
    return reasons


def filter_step_payload_to_schema(
    payload: dict[str, Any],
    schema: Optional[RedcapSchema],
) -> dict[str, Any]:
    if schema is None:
        return payload
    special = {"redcap_data_access_group"}
    filtered = {}
    for key, value in payload.items():
        if key in schema.field_names or key in special or key == schema.record_field:
            filtered[key] = value
    return filtered


def process_step_account(
    row: pd.Series,
    cols: dict[str, Any],
    token_dir: Path,
    sessions: GarminSessionManager,
    caller: GarminCaller,
    recorder: CombinedRecorder,
    redcap: Optional[RedcapClient],
    redcap_schema: Optional[RedcapSchema],
    stats_mode: str,
    run_pass: str,
    attempt_sequence: int,
    project_output_dir: Path,
) -> ProjectResult:
    project = "SyncStep"
    account_t0 = time.perf_counter()
    counter_start = caller.counters.snapshot()

    record_id = clean_text(row.get(cols["record_id"]))
    username = clean_text(row.get(cols["username"]))
    model = clean_text(row.get(cols["model"]))
    serial = clean_text(row.get(cols["serial"]))
    team = clean_text(row.get(cols["team"]))
    excel_row = int(row["__excel_row__"])
    is_vivoactive, watch_type, watch_category = classify_step_watch(model)

    metrics: dict[str, Any] = {
        "run_id": recorder.run_id,
        "attempt_sequence": attempt_sequence,
        "project": project,
        "pacing_enabled": caller.pacing_enabled,
        "configured_qps": caller.qps if caller.pacing_enabled else None,
        "pass_name": run_pass,
        "record_id": record_id,
        "watch_category": watch_category,
        "watch_type_code": watch_type,
        "outcome": "started",
        "upload_success": False,
        "redcap_import_accepted": None,
        "token_source": None,
        "login_seconds": None,
        "download_iteration_seconds": None,
        "download_api_active_seconds": None,
        "data_sort_seconds": None,
        "redcap_existing_read_seconds": 0.0,
        "redcap_upload_seconds": None,
        "redcap_verify_seconds": 0.0,
        "redcap_verified": None,
        "redcap_verify_mismatches": None,
        "total_account_seconds": None,
        "days_requested": None,
        "valid_days_with_steps": None,
        "days_with_sleep_data": None,
        "days_with_any_data": None,
        "days_with_steps_and_sleep": None,
        "days_with_steps_only": None,
        "days_with_sleep_only": None,
        "days_with_no_steps_or_sleep": None,
        "percent_days_with_steps": None,
        "percent_days_with_sleep": None,
        "percent_days_worn": None,  # backward-compatible alias for percent_days_with_steps
        "download_api_attempts": None,
        "download_api_errors": None,
        "download_retry_events": None,
        "preop_days_found": None,
        "postop_days_found": None,
        "preop_days_uploaded": None,
        "postop_days_uploaded": None,
        "payload_json_bytes": None,
        "fields_nonblank_uploaded": None,
        "error_text": None,
    }

    try:
        reasons = step_validation_reasons(row, cols)
        if reasons:
            message = "; ".join(reasons)
            metrics["outcome"] = "validation-fail"
            metrics["error_text"] = message
            return ProjectResult(
                project=project, success=False, retryable=False, record_id=record_id,
                excel_row=excel_row, watch_category=watch_category,
                error=message, metrics=metrics,
            )

        start_date = safe_date(row.get(cols["start"]))
        dos_date = safe_date(row.get(cols["dos"]))
        end_planned = safe_date(row.get(cols["end"]))
        assert start_date is not None and dos_date is not None and end_planned is not None
        end_date, final_end = assignment_window(row)
        metrics.update({'query_start': start_date.isoformat(), 'query_end': end_date.isoformat(),
                        'assignment_final_end': final_end.isoformat(),
                        'window_limit_reason': row.get('__assignment_limit_reason__'),
                        'ownership_warning': row.get('__assignment_warnings__'),
                        'hrv_mode': redcap_schema.hrv_mode if redcap_schema else 'overnight-ms (export only)',
                        'hrv_configuration_warning': redcap_schema.hrv_warning if redcap_schema else ''})
        if start_date > end_date:
            message = f"Invalid/future window: {start_date} -> {end_date}"
            metrics["outcome"] = "validation-fail"
            metrics["error_text"] = message
            return ProjectResult(
                project=project, success=False, retryable=False, record_id=record_id,
                excel_row=excel_row, watch_category=watch_category,
                error=message, metrics=metrics,
            )

        metrics["days_requested"] = (end_date - start_date).days + 1

        login_t0 = time.perf_counter()
        try:
            client, token_source = sessions.get_client(
                project, username, token_dir, record_id, watch_category
            )
            metrics["token_source"] = token_source
            metrics["login_seconds"] = time.perf_counter() - login_t0
        except GarminRateLimit as exc:
            metrics["login_seconds"] = time.perf_counter() - login_t0
            metrics["outcome"] = "login-429"
            metrics["error_text"] = str(exc)
            return ProjectResult(
                project=project, success=False, retryable=True, record_id=record_id,
                excel_row=excel_row, watch_category=watch_category,
                error=str(exc), metrics=metrics,
            )
        except Exception as exc:
            metrics["login_seconds"] = time.perf_counter() - login_t0
            metrics["outcome"] = "login-fail"
            metrics["error_text"] = str(exc)
            return ProjectResult(
                project=project, success=False, retryable=False, record_id=record_id,
                excel_row=excel_row, watch_category=watch_category,
                error=str(exc), metrics=metrics,
            )

        preop_days: dict[int, dict[str, Any]] = {}
        postop_days: dict[int, dict[str, Any]] = {}
        daily_rows: list[dict[str, Any]] = []
        latest_data_date: Optional[date] = None
        latest_step_date: Optional[date] = None
        valid_days = 0
        days_with_sleep_data = 0
        days_with_any_data = 0
        days_with_steps_and_sleep = 0
        days_with_steps_only = 0
        days_with_sleep_only = 0
        days_with_no_steps_or_sleep = 0

        download_counter_start = caller.counters.snapshot()
        download_t0 = time.perf_counter()
        current = start_date
        iteration_index = 0
        try:
            while current <= end_date:
                iteration_index += 1
                cur_str = current.isoformat()
                relative_day = (current - dos_date).days
                iter_t0 = time.perf_counter()
                iter_counter_start = caller.counters.snapshot()
                iter_status = "success"
                iter_error = None

                try:
                    summary_t0 = time.perf_counter()
                    stats, summary = get_step_daily_summary_sources(
                        client,
                        cur_str,
                        is_vivoactive,
                        stats_mode,
                        caller,
                        project,
                        record_id,
                        watch_category,
                    )
                    summary_wall = time.perf_counter() - summary_t0

                    steps = first_not_none(
                        extract_actual(stats, "totalSteps"),
                        extract_actual(summary, "totalSteps"),
                    )
                    sleep_t0 = time.perf_counter()
                    sleep = caller.call(
                        client.get_sleep_data,
                        cur_str,
                        project=project,
                        record_id=record_id,
                        watch_category=watch_category,
                        date_value=cur_str,
                    ) or {}
                    sleep_wall = time.perf_counter() - sleep_t0
                    daily_sleep = sleep.get("dailySleepDTO", {}) if isinstance(sleep, dict) else {}
                    daily_sleep = daily_sleep or {}
                    # Failed retrievals never become blanks in an overwrite payload.
                    hrv_payload = None
                    if is_vivoactive:
                        hrv_payload = caller.call(
                            client.get_hrv_data, cur_str, project=project,
                            record_id=record_id, watch_category=watch_category, date_value=cur_str,
                        )
                    hrv_values = extract_hrv(sleep, hrv_payload, cur_str)
                    if not is_vivoactive:
                        hrv_values['hrv_source'] = 'HRV not collected for vivofit model'
                    if redcap_schema and redcap_schema.hrv_mode == 'status' and hrv_values['hrv_status'] is not None:
                        status_key = hrv_status_key(hrv_values['hrv_status'])
                        if any(mapping and status_key not in mapping for mapping in redcap_schema.hrv_status_choices.values()):
                            raise RedcapError('Observed Garmin HRV status has no matching REDCap choice; account upload held')
                    intensity_values = extract_intensity(stats, summary)
                    light = daily_sleep.get("lightSleepSeconds")
                    deep = daily_sleep.get("deepSleepSeconds")
                    awake = daily_sleep.get("awakeSleepSeconds")
                    rem = daily_sleep.get("remSleepSeconds") if is_vivoactive else None
                    duration = daily_sleep.get("durationInSeconds")
                    if duration is None:
                        # Only derive total sleep when every stage expected for that
                        # watch type is actually available. Partial stage data remains
                        # partial; it is never silently summed into a misleading total.
                        expected_parts = [light, deep]
                        if is_vivoactive:
                            expected_parts.append(rem)
                        if expected_parts and all(isinstance(v, (int, float)) for v in expected_parts):
                            duration = sum(expected_parts)

                    # IMPORTANT: each metric preserves Garmin's own missingness. A
                    # missing sleep/HR/HRV/etc value stays None even if steps exist,
                    # and a true numeric zero stays zero even if steps are absent.
                    data = {
                        "steps": steps,
                        "totalsleep": as_minutes(duration),
                        "lightsleep": as_minutes(light),
                        "deepsleep": as_minutes(deep),
                        "awake": as_minutes(awake),
                        "remsleep": as_minutes(rem),
                        "restinghr": actual_from(stats, summary, "restingHeartRate"),
                        "maxhr": actual_from(stats, summary, "maxHeartRate"),
                        "hrv": hrv_values['hrv_status'] if redcap_schema and redcap_schema.hrv_mode == 'status' else hrv_values['hrv_overnight_ms'],
                        "rr": actual_from(stats, summary, "avgWakingRespirationValue", "averageRespirationValue"),
                        "pulseox": actual_from(stats, summary, "averageSpo2", "avgSpO2", "avgOvernightSpo2"),
                        **hrv_values,
                        **intensity_values,
                    }
                    availability = step_availability_pattern(data)
                    if availability["steps_present"]:
                        valid_days += 1
                        latest_step_date = current
                    if availability["sleep_present"]:
                        days_with_sleep_data += 1
                    if availability["any_metric_present"]:
                        days_with_any_data += 1
                        latest_data_date = current
                    if availability["data_pattern"] == "steps_and_sleep":
                        days_with_steps_and_sleep += 1
                    elif availability["data_pattern"] == "steps_only":
                        days_with_steps_only += 1
                    elif availability["data_pattern"] == "sleep_only":
                        days_with_sleep_only += 1
                    else:
                        days_with_no_steps_or_sleep += 1

                    if relative_day < 0:
                        preop_days[relative_day] = data
                    elif relative_day > 0:
                        postop_days[relative_day] = data

                    daily_rows.append({
                        "project": project,
                        "record_id": record_id,
                        "watch_category": watch_category,
                        "date": cur_str,
                        "relative_day": relative_day,
                        **availability,
                        **data,
                    })

                except GarminRateLimit as exc:
                    iter_status = "rate-limited"
                    iter_error = str(exc)
                    raise
                except Exception as exc:
                    iter_status = "error"
                    iter_error = str(exc)
                    raise
                finally:
                    iter_counter_end = caller.counters.snapshot()
                    recorder.add_iteration({
                        "run_id": recorder.run_id,
                        "project": project,
                        "pacing_enabled": caller.pacing_enabled,
                        "record_id": record_id,
                        "watch_category": watch_category,
                        "pass_name": run_pass,
                        "iteration_index": iteration_index,
                        "date": cur_str,
                        "relative_day": relative_day,
                        "iteration_seconds": round(time.perf_counter() - iter_t0, 6),
                        "garmin_api_attempts": (
                            iter_counter_end["api_attempts"] - iter_counter_start["api_attempts"]
                        ),
                        "garmin_api_errors": (
                            iter_counter_end["api_errors"] - iter_counter_start["api_errors"]
                        ),
                        "garmin_api_active_seconds": round(
                            iter_counter_end["api_active_seconds"]
                            - iter_counter_start["api_active_seconds"], 6
                        ),
                        "status": iter_status,
                        "error": iter_error,
                    })

                current += timedelta(days=1)

        except GarminRateLimit as exc:
            metrics["download_iteration_seconds"] = time.perf_counter() - download_t0
            metrics["outcome"] = "download-429"
            metrics["error_text"] = str(exc)
            return ProjectResult(
                project=project, success=False, retryable=True, record_id=record_id,
                excel_row=excel_row, watch_category=watch_category,
                daily_rows=daily_rows, error=str(exc), metrics=metrics,
            )
        except Exception as exc:
            metrics["download_iteration_seconds"] = time.perf_counter() - download_t0
            metrics["outcome"] = "download-error"
            metrics["error_text"] = str(exc)
            return ProjectResult(
                project=project,
                success=False,
                retryable=is_transient_error(str(exc)),
                record_id=record_id,
                excel_row=excel_row,
                watch_category=watch_category,
                daily_rows=daily_rows,
                error=str(exc),
                metrics=metrics,
            )
        finally:
            if metrics["download_iteration_seconds"] is None:
                metrics["download_iteration_seconds"] = time.perf_counter() - download_t0

        download_counter_end = caller.counters.snapshot()
        metrics["download_api_active_seconds"] = (
            download_counter_end["api_active_seconds"] - download_counter_start["api_active_seconds"]
        )
        metrics["download_api_attempts"] = (
            download_counter_end["api_attempts"] - download_counter_start["api_attempts"]
        )
        metrics["download_api_errors"] = (
            download_counter_end["api_errors"] - download_counter_start["api_errors"]
        )
        metrics["download_retry_events"] = (
            download_counter_end["retry_events"] - download_counter_start["retry_events"]
        )
        metrics["valid_days_with_steps"] = valid_days
        metrics["days_with_sleep_data"] = days_with_sleep_data
        metrics["days_with_any_data"] = days_with_any_data
        metrics["days_with_steps_and_sleep"] = days_with_steps_and_sleep
        metrics["days_with_steps_only"] = days_with_steps_only
        metrics["days_with_sleep_only"] = days_with_sleep_only
        metrics["days_with_no_steps_or_sleep"] = days_with_no_steps_or_sleep
        metrics['days_with_numeric_hrv'] = sum(x.get('hrv_overnight_ms') is not None for x in daily_rows)
        metrics['days_with_hrv_status'] = sum(x.get('hrv_status') is not None for x in daily_rows)
        metrics['days_with_total_intensity'] = sum(x.get('totalintensity') is not None for x in daily_rows)
        metrics["percent_days_with_steps"] = (
            round(100 * valid_days / metrics["days_requested"], 1)
            if metrics["days_requested"] else None
        )
        metrics["percent_days_with_sleep"] = (
            round(100 * days_with_sleep_data / metrics["days_requested"], 1)
            if metrics["days_requested"] else None
        )
        # Retain the old output name only for workbook/backward compatibility.
        # It is step-data availability, not validated device wear.
        metrics["percent_days_worn"] = metrics["percent_days_with_steps"]
        metrics["preop_days_found"] = sum(1 for values in preop_days.values() if has_any_step_metric(values))
        metrics["postop_days_found"] = sum(1 for values in postop_days.values() if has_any_step_metric(values))

        sort_t0 = time.perf_counter()
        payload = {
            "record_id": record_id,
            "redcap_data_access_group": resolve_step_dag(team) or "",
            "watch_type": watch_type,
            "watch_number": serial,
        }

        # STEP wearable fields are pipeline-managed. Initialize all supported
        # pre-op slots and POD1..POD90 to blank, then fill only values Garmin
        # actually returned. With REDCap overwriteBehavior='overwrite', this
        # removes stale placeholder zeros and prior shifted values safely without
        # touching unrelated REDCap fields.
        for slot in range(1, STEP_PREOP_DAYS + 1):
            for field_name in STEP_FIELDS:
                if field_name == 'hrv' and redcap_schema and redcap_schema.hrv_mode == 'off':
                    continue
                payload[f"preop{field_name}_day{slot}"] = ""
        for pod in range(1, STEP_POSTOP_DAYS + 1):
            for field_name in STEP_FIELDS:
                if field_name == 'hrv' and redcap_schema and redcap_schema.hrv_mode == 'off':
                    continue
                payload[f"postop{field_name}_day{pod}"] = ""

        preop_candidates = {
            delta: values for delta, values in preop_days.items()
            if has_any_step_metric(values)
        }
        selected_preop = sorted(preop_candidates.items(), reverse=True)[:STEP_PREOP_DAYS]
        selected_preop.sort()
        for slot, (_delta, values) in enumerate(selected_preop, 1):
            for field_name in STEP_FIELDS:
                if field_name == 'hrv' and redcap_schema and redcap_schema.hrv_mode == 'off':
                    continue
                value = step_metric_value(field_name, values, f'preop{field_name}_day{slot}', redcap_schema)
                if value is not None:
                    payload[f"preop{field_name}_day{slot}"] = value

        # Postoperative data are mapped by the ACTUAL date minus surgery date.
        # A window beginning on POD3 therefore writes POD3, never POD1.
        selected_postop = [
            (relative_day, values)
            for relative_day, values in sorted(postop_days.items())
            if 1 <= relative_day <= STEP_POSTOP_DAYS
        ]
        for relative_day, values in selected_postop:
            for field_name in STEP_FIELDS:
                if field_name == 'hrv' and redcap_schema and redcap_schema.hrv_mode == 'off':
                    continue
                value = step_metric_value(field_name, values, f'postop{field_name}_day{relative_day}', redcap_schema)
                if value is not None:
                    payload[f"postop{field_name}_day{relative_day}"] = value

        payload = filter_step_payload_to_schema(payload, redcap_schema)
        metrics["preop_days_uploaded"] = len(selected_preop)
        metrics["postop_days_uploaded"] = sum(
            1 for _relative_day, values in selected_postop if has_any_step_metric(values)
        )
        metrics["fields_nonblank_uploaded"] = sum(
            1 for v in payload.values() if v not in (None, "")
        )
        metrics["payload_json_bytes"] = len(
            json.dumps([sanitize_record(payload)], ensure_ascii=False).encode("utf-8")
        )
        metrics["data_sort_seconds"] = time.perf_counter() - sort_t0

        upload_ok = True
        upload_error = None
        if redcap is not None:
            upload_t0 = time.perf_counter()
            try:
                ok, status, response_text, elapsed = redcap.upload_record(payload, overwrite_behavior="overwrite")
                metrics["redcap_upload_seconds"] = elapsed
                if not ok:
                    metrics["redcap_import_accepted"] = False
                    upload_ok = False
                    upload_error = f"REDCap HTTP {status}: {response_text[:1000]}"
                else:
                    metrics["redcap_import_accepted"] = True
                if ok and redcap_schema is None:
                    upload_ok = False
                    upload_error = "SyncStep REDCap schema unavailable for post-import verification"
                elif ok:
                    verify_fields = [
                        key for key in payload
                        if key != "redcap_data_access_group"
                        and (key == redcap_schema.record_field or key in redcap_schema.field_names)
                    ]
                    verify_t0 = time.perf_counter()
                    try:
                        readback = redcap.export_record_fields(
                            record_id, redcap_schema.record_field, verify_fields
                        )
                    finally:
                        metrics["redcap_verify_seconds"] = time.perf_counter() - verify_t0
                    mismatches = verify_step_readback(payload, readback, redcap_schema)
                    metrics["redcap_verify_mismatches"] = len(mismatches)
                    metrics["redcap_verified"] = not mismatches
                    if mismatches:
                        upload_ok = False
                        upload_error = (
                            f"REDCap readback mismatch in {len(mismatches)} field(s): "
                            + " | ".join(mismatches[:10])
                        )
            except Exception as exc:
                if metrics.get("redcap_upload_seconds") is None:
                    metrics["redcap_upload_seconds"] = time.perf_counter() - upload_t0
                upload_ok = False
                upload_error = str(exc)
        else:
            metrics["redcap_upload_seconds"] = 0.0
            metrics["redcap_verify_seconds"] = 0.0
            metrics["redcap_import_accepted"] = None
            metrics["redcap_verified"] = None
            metrics["redcap_verify_mismatches"] = None

        if not upload_ok:
            metrics["outcome"] = "redcap-error"
            metrics["error_text"] = upload_error
            return ProjectResult(
                project=project, success=False, retryable=False, record_id=record_id,
                excel_row=excel_row, watch_category=watch_category,
                payload=payload, daily_rows=daily_rows, error=upload_error, metrics=metrics,
            )

        metrics["upload_success"] = redcap is not None
        metrics["outcome"] = "success" if redcap is not None else "download-transform-success-no-upload"

        sync_row = {
            "record_id": record_id,
            "team": team,
            "end_date_planned": final_end,
            "query_end": end_date, "window_limit_reason": row.get('__assignment_limit_reason__'),
            "ownership_warning": row.get('__assignment_warnings__'),
            "last_synced": latest_data_date,
            "last_step_date": latest_step_date,
            "days_since_sync": (date.today() - latest_data_date).days if latest_data_date else None,
            "percent_days_with_steps": metrics["percent_days_with_steps"],
            "percent_days_with_sleep": metrics["percent_days_with_sleep"],
            "percent_days_worn": metrics["percent_days_worn"],
        }
        writeback = {
            "record_id": record_id,
            "excel_row": excel_row,
            "verified": bool(redcap is not None and final_end < date.today() and not (redcap_schema and redcap_schema.hrv_warning)),
            "last_successful_upload": now_local_iso() if redcap is not None else None,
            "reason_failure": redcap_schema.hrv_warning if redcap_schema else "",
            **sync_row,
        }
        return ProjectResult(
            project=project, success=True, retryable=False, record_id=record_id,
            excel_row=excel_row, watch_category=watch_category,
            payload=payload, daily_rows=daily_rows,
            sync_row=sync_row, writeback=writeback, metrics=metrics,
        )

    finally:
        metrics["total_account_seconds"] = time.perf_counter() - account_t0
        counter_end = caller.counters.snapshot()
        # Fill counters if the account exited before the normal download-finalization path.
        if metrics.get("download_api_active_seconds") is None:
            metrics["download_api_active_seconds"] = (
                counter_end["api_active_seconds"] - counter_start["api_active_seconds"]
            )
        if metrics.get("download_api_attempts") is None:
            metrics["download_api_attempts"] = (
                counter_end["api_attempts"] - counter_start["api_attempts"]
            )
        if metrics.get("download_api_errors") is None:
            metrics["download_api_errors"] = (
                counter_end["api_errors"] - counter_start["api_errors"]
            )
        if metrics.get("download_retry_events") is None:
            metrics["download_retry_events"] = (
                counter_end["retry_events"] - counter_start["retry_events"]
            )
        for key in (
            "login_seconds", "download_iteration_seconds", "download_api_active_seconds",
            "data_sort_seconds", "redcap_upload_seconds", "redcap_verify_seconds", "total_account_seconds",
        ):
            if isinstance(metrics.get(key), (int, float)):
                metrics[key] = round(float(metrics[key]), 6)
        recorder.add_attempt(metrics.copy())


# =============================================================================
# SYNCROVE ACCOUNT PROCESSOR
# =============================================================================


def rove_validation_reasons(row: pd.Series, cols: dict[str, Any]) -> list[str]:
    reasons = []
    try:
        assignment_window(row)
    except ValueError as exc:
        reasons.append(str(exc))
    if not clean_text(row.get(cols["study_id"])):
        reasons.append("Missing Study ID")
    if not clean_text(row.get(cols["username"])):
        reasons.append("Missing username")
    # Vivofit4 Serial Number is retained for audit/reference only. It is not
    # required for Garmin token restoration or for the REDCap daily-step upload.
    if safe_date(row.get(cols["start"])) is None:
        reasons.append("Missing/invalid Start Date")
    if safe_date(row.get(cols["dos"])) is None:
        reasons.append("Missing/invalid DOS")
    return reasons


def process_rove_account(
    row: pd.Series,
    cols: dict[str, Any],
    token_dir: Path,
    sessions: GarminSessionManager,
    caller: GarminCaller,
    recorder: CombinedRecorder,
    redcap: Optional[RedcapClient],
    redcap_schema: Optional[RedcapSchema],
    run_pass: str,
    attempt_sequence: int,
) -> ProjectResult:
    project = "SyncRove"
    account_t0 = time.perf_counter()
    counter_start = caller.counters.snapshot()
    record_id = clean_text(row.get(cols["study_id"]))
    username = clean_text(row.get(cols["username"]))
    serial = clean_text(row.get(cols["serial"]))
    excel_row = int(row["__excel_row__"])
    watch_category = "vívofit 4"

    metrics: dict[str, Any] = {
        "run_id": recorder.run_id, "attempt_sequence": attempt_sequence, "project": project,
        "pacing_enabled": caller.pacing_enabled, "configured_qps": caller.qps if caller.pacing_enabled else None,
        "pass_name": run_pass, "record_id": record_id, "watch_category": watch_category,
        "watch_type_code": 1, "outcome": "started", "upload_success": False, "redcap_import_accepted": None,
        "token_source": None, "login_seconds": None, "download_iteration_seconds": None,
        "download_api_active_seconds": None, "data_sort_seconds": None,
        "redcap_existing_read_seconds": 0.0, "redcap_upload_seconds": None,
        "redcap_verify_seconds": 0.0, "redcap_verified": None, "redcap_verify_mismatches": None,
        "total_account_seconds": None, "days_requested": None, "days_retained_for_redcap": None,
        "valid_days_with_steps": None, "percent_days_with_steps": None,
        "percent_days_worn": None,  # backward-compatible alias for percent_days_with_steps
        "download_api_attempts": None, "download_api_errors": None, "download_retry_events": None,
        "preop_days_found": None, "postop_days_found": None, "preop_days_uploaded": None,
        "postop_days_uploaded": None, "existing_instances_reused": 0, "new_instances_created": 0,
        "payload_json_bytes": None, "payload_csv_bytes": None, "fields_nonblank_uploaded": None, "error_text": None,
    }

    try:
        reasons = rove_validation_reasons(row, cols)
        if reasons:
            message = "; ".join(reasons)
            metrics["outcome"] = "validation-fail"; metrics["error_text"] = message
            return ProjectResult(project=project, success=False, retryable=False, record_id=record_id,
                                 excel_row=excel_row, watch_category=watch_category, error=message, metrics=metrics)

        start_date = safe_date(row.get(cols["start"]))
        dos_date = safe_date(row.get(cols["dos"]))
        assert start_date is not None and dos_date is not None
        end_date, planned_end = assignment_window(row)
        metrics.update({'query_start': start_date.isoformat(), 'query_end': end_date.isoformat(),
                        'assignment_final_end': planned_end.isoformat(),
                        'window_limit_reason': row.get('__assignment_limit_reason__'),
                        'ownership_warning': row.get('__assignment_warnings__')})
        if start_date > end_date:
            message = f"Invalid/future window: {start_date} -> {end_date}"
            metrics["outcome"] = "validation-fail"; metrics["error_text"] = message
            return ProjectResult(project=project, success=False, retryable=False, record_id=record_id,
                                 excel_row=excel_row, watch_category=watch_category, error=message, metrics=metrics)
        metrics["days_requested"] = (end_date - start_date).days + 1

        login_t0 = time.perf_counter()
        try:
            client, token_source = sessions.get_client(project, username, token_dir, record_id, watch_category)
            metrics["token_source"] = token_source
            metrics["login_seconds"] = time.perf_counter() - login_t0
        except GarminRateLimit as exc:
            metrics["login_seconds"] = time.perf_counter() - login_t0; metrics["outcome"] = "login-429"; metrics["error_text"] = str(exc)
            return ProjectResult(project=project, success=False, retryable=True, record_id=record_id, excel_row=excel_row, watch_category=watch_category, error=str(exc), metrics=metrics)
        except Exception as exc:
            metrics["login_seconds"] = time.perf_counter() - login_t0; metrics["outcome"] = "login-fail"; metrics["error_text"] = str(exc)
            return ProjectResult(project=project, success=False, retryable=False, record_id=record_id, excel_row=excel_row, watch_category=watch_category, error=str(exc), metrics=metrics)

        daily_rows: list[dict[str, Any]] = []
        valid_days = 0
        latest_data_date: Optional[date] = None
        preop_found = postop_found = 0
        download_counter_start = caller.counters.snapshot()
        download_t0 = time.perf_counter()
        current = start_date
        iteration_index = 0
        try:
            while current <= end_date:
                iteration_index += 1
                cur_str = current.isoformat()
                relative_day = (current - dos_date).days
                iter_t0 = time.perf_counter()
                iter_counter_start = caller.counters.snapshot()
                iter_status = "success"; iter_error = None
                try:
                    stats = caller.call(client.get_stats, cur_str, project=project, record_id=record_id, watch_category=watch_category, date_value=cur_str) or {}
                    steps = extract_actual(stats, "totalSteps")
                    if steps is not None:
                        try: steps = int(float(steps))
                        except Exception: pass
                        valid_days += 1; latest_data_date = current
                    if ROVE_MIN_RELATIVE_DAY <= relative_day <= ROVE_POSTOP_DAYS:
                        if relative_day < 0: preop_found += 1
                        elif relative_day > 0: postop_found += 1
                        daily_rows.append({
                            "project": project, "record_id": record_id, "watch_category": watch_category,
                            "vivofit4_serial_number": serial, "date": cur_str, "relative_day": relative_day,
                            "phase": "preop" if relative_day < 0 else ("DOS" if relative_day == 0 else "postop"),
                            "steps_present": steps is not None,
                            "steps": steps,
                        })
                except GarminRateLimit as exc:
                    iter_status = "rate-limited"; iter_error = str(exc); raise
                except Exception as exc:
                    iter_status = "error"; iter_error = str(exc); raise
                finally:
                    iter_counter_end = caller.counters.snapshot()
                    recorder.add_iteration({
                        "run_id": recorder.run_id, "project": project, "pacing_enabled": caller.pacing_enabled, "record_id": record_id,
                        "watch_category": watch_category, "pass_name": run_pass, "iteration_index": iteration_index,
                        "date": cur_str, "relative_day": relative_day,
                        "iteration_seconds": round(time.perf_counter() - iter_t0, 6),
                        "garmin_api_attempts": iter_counter_end["api_attempts"] - iter_counter_start["api_attempts"],
                        "garmin_api_errors": iter_counter_end["api_errors"] - iter_counter_start["api_errors"],
                        "garmin_api_active_seconds": round(iter_counter_end["api_active_seconds"] - iter_counter_start["api_active_seconds"], 6),
                        "status": iter_status, "error": iter_error,
                    })
                current += timedelta(days=1)
        except GarminRateLimit as exc:
            metrics["download_iteration_seconds"] = time.perf_counter() - download_t0; metrics["outcome"] = "download-429"; metrics["error_text"] = str(exc)
            return ProjectResult(project=project, success=False, retryable=True, record_id=record_id, excel_row=excel_row, watch_category=watch_category, daily_rows=daily_rows, error=str(exc), metrics=metrics)
        except Exception as exc:
            metrics["download_iteration_seconds"] = time.perf_counter() - download_t0; metrics["outcome"] = "download-error"; metrics["error_text"] = str(exc)
            return ProjectResult(project=project, success=False, retryable=is_transient_error(str(exc)), record_id=record_id, excel_row=excel_row, watch_category=watch_category, daily_rows=daily_rows, error=str(exc), metrics=metrics)
        finally:
            if metrics["download_iteration_seconds"] is None:
                metrics["download_iteration_seconds"] = time.perf_counter() - download_t0

        download_counter_end = caller.counters.snapshot()
        metrics["download_api_active_seconds"] = download_counter_end["api_active_seconds"] - download_counter_start["api_active_seconds"]
        metrics["download_api_attempts"] = download_counter_end["api_attempts"] - download_counter_start["api_attempts"]
        metrics["download_api_errors"] = download_counter_end["api_errors"] - download_counter_start["api_errors"]
        metrics["download_retry_events"] = download_counter_end["retry_events"] - download_counter_start["retry_events"]
        metrics["valid_days_with_steps"] = valid_days
        metrics["percent_days_with_steps"] = round(100 * valid_days / metrics["days_requested"], 1) if metrics["days_requested"] else None
        metrics["percent_days_worn"] = metrics["percent_days_with_steps"]
        metrics["days_retained_for_redcap"] = len(daily_rows)
        metrics["preop_days_found"] = preop_found; metrics["postop_days_found"] = postop_found

        sort_t0 = time.perf_counter()
        payload: dict[str, Any] = {
            "record_id": record_id, "redcap_mode": "long_repeating",
            "daily_rows_prepared": len(daily_rows),
            "repeat_instrument": redcap_schema.repeat_instrument if redcap_schema else None,
            "repeat_event": redcap_schema.repeat_event if redcap_schema else None,
        }
        redcap_rows: list[dict[str, Any]] = []
        boundary_rows: list[dict[str, Any]] = []
        try:
            if redcap is not None:
                if redcap_schema is None or redcap_schema.mode != "long_repeating":
                    raise RedcapError("SyncRove REDCap schema is not the required long/repeating step_data structure.")
                if not redcap.export_record_exists(record_id, redcap_schema.record_field):
                    raise RedcapError(f"SyncRove record {record_id} does not exist; refusing to create it implicitly.")
                read_t0 = time.perf_counter()
                existing_rows = redcap.export_rove_step_rows(record_id, redcap_schema)
                metrics["redcap_existing_read_seconds"] = time.perf_counter() - read_t0
                for existing in existing_rows:
                    relative = _rove_int(existing.get(redcap_schema.relative_day_field))
                    if relative is None:
                        continue
                    measured_date = dos_date + timedelta(days=relative)
                    if measured_date < start_date or measured_date > end_date:
                        boundary_rows.append({
                            'record_id': record_id, 'date': measured_date.isoformat(),
                            'relative_day': relative, 'redcap_repeat_instance': existing.get('redcap_repeat_instance'),
                            'steps': existing.get(redcap_schema.steps_field),
                            'query_start': start_date.isoformat(), 'query_end': end_date.isoformat(),
                            'action': 'Review existing out-of-window entry; no automatic deletion',
                        })
                metrics['existing_rows_outside_assignment'] = len(boundary_rows)
                redcap_rows, reused, created = prepare_rove_redcap_rows(record_id, daily_rows, existing_rows, redcap_schema)
                metrics["existing_instances_reused"] = reused; metrics["new_instances_created"] = created
                payload["existing_instances_reused"] = reused; payload["new_instances_created"] = created
                metrics["preop_days_uploaded"] = sum(1 for x in daily_rows if int(x["relative_day"]) < 0)
                metrics["postop_days_uploaded"] = sum(1 for x in daily_rows if int(x["relative_day"]) > 0)
                metrics["fields_nonblank_uploaded"] = sum(1 for rec in redcap_rows for v in rec.values() if v not in (None, ""))
                csv_payload = redcap_rows_to_csv_text(redcap_rows, redcap_schema.import_headers)
                metrics["payload_csv_bytes"] = len(csv_payload.encode("utf-8"))
                metrics["payload_json_bytes"] = None
            else:
                metrics["preop_days_uploaded"] = 0; metrics["postop_days_uploaded"] = 0
                metrics["fields_nonblank_uploaded"] = 0; metrics["payload_json_bytes"] = None; metrics["payload_csv_bytes"] = 0
        except Exception as exc:
            metrics["data_sort_seconds"] = time.perf_counter() - sort_t0; metrics["outcome"] = "redcap-map-error"; metrics["error_text"] = str(exc)
            return ProjectResult(project=project, success=False, retryable=False, record_id=record_id, excel_row=excel_row, watch_category=watch_category, payload=payload, redcap_rows=redcap_rows, daily_rows=daily_rows, error=str(exc), metrics=metrics, boundary_rows=boundary_rows)
        metrics["data_sort_seconds"] = time.perf_counter() - sort_t0

        if redcap is not None:
            upload_t0 = time.perf_counter()
            try:
                ok, status, response_text, elapsed = redcap.upload_records_csv(
                    redcap_rows, redcap_schema.import_headers, overwrite_behavior="overwrite"
                )
                metrics["redcap_upload_seconds"] = elapsed
                if not ok:
                    metrics["redcap_import_accepted"] = False
                    raise RedcapError(f"REDCap HTTP {status}: {response_text[:1000]}")
                metrics["redcap_import_accepted"] = True
                verify_t0 = time.perf_counter()
                try:
                    readback_rows = redcap.export_rove_step_rows(record_id, redcap_schema)
                finally:
                    metrics["redcap_verify_seconds"] = time.perf_counter() - verify_t0
                mismatches = verify_rove_readback(redcap_rows, readback_rows, redcap_schema)
                metrics["redcap_verify_mismatches"] = len(mismatches)
                metrics["redcap_verified"] = not mismatches
                if mismatches:
                    raise RedcapError(
                        f"REDCap readback mismatch in {len(mismatches)} Rove row(s): "
                        + " | ".join(mismatches[:10])
                    )
            except Exception as exc:
                if metrics.get("redcap_upload_seconds") is None:
                    metrics["redcap_upload_seconds"] = time.perf_counter() - upload_t0
                metrics["outcome"] = "redcap-error"; metrics["error_text"] = str(exc)
                return ProjectResult(project=project, success=False, retryable=False, record_id=record_id, excel_row=excel_row, watch_category=watch_category, payload=payload, redcap_rows=redcap_rows, daily_rows=daily_rows, error=str(exc), metrics=metrics, boundary_rows=boundary_rows)
        else:
            metrics["redcap_upload_seconds"] = 0.0
            metrics["redcap_verify_seconds"] = 0.0
            metrics["redcap_import_accepted"] = None
            metrics["redcap_verified"] = None
            metrics["redcap_verify_mismatches"] = None

        metrics["upload_success"] = redcap is not None
        metrics["outcome"] = "success" if redcap is not None else "download-transform-success-no-upload"
        sync_row = {
            "record_id": record_id, "planned_final_end": planned_end, "last_synced": latest_data_date,
            "days_since_sync": (date.today() - latest_data_date).days if latest_data_date else None,
            "percent_days_worn": metrics["percent_days_worn"], "final_window_reached": date.today() > planned_end,
            "query_end": end_date, "window_limit_reason": row.get('__assignment_limit_reason__'),
            "ownership_warning": row.get('__assignment_warnings__'),
            "existing_rows_outside_assignment": len(boundary_rows),
        }
        writeback = {
            "record_id": record_id, "excel_row": excel_row,
            "capture_complete": bool(redcap is not None and date.today() > planned_end and not boundary_rows),
            "last_successful_capture": now_local_iso() if redcap is not None else None,
            "last_synced": latest_data_date, "days_since_sync": sync_row["days_since_sync"],
            "percent_days_worn": metrics["percent_days_worn"], "reason_failure": ('Existing REDCap rows outside assignment require review' if boundary_rows else ''),
        }
        return ProjectResult(project=project, success=True, retryable=False, record_id=record_id, excel_row=excel_row, watch_category=watch_category, payload=payload, redcap_rows=redcap_rows, daily_rows=daily_rows, sync_row=sync_row, writeback=writeback, metrics=metrics, boundary_rows=boundary_rows)
    finally:
        metrics["total_account_seconds"] = time.perf_counter() - account_t0
        counter_end = caller.counters.snapshot()
        if metrics.get("download_api_active_seconds") is None:
            metrics["download_api_active_seconds"] = counter_end["api_active_seconds"] - counter_start["api_active_seconds"]
        if metrics.get("download_api_attempts") is None:
            metrics["download_api_attempts"] = counter_end["api_attempts"] - counter_start["api_attempts"]
        if metrics.get("download_api_errors") is None:
            metrics["download_api_errors"] = counter_end["api_errors"] - counter_start["api_errors"]
        if metrics.get("download_retry_events") is None:
            metrics["download_retry_events"] = counter_end["retry_events"] - counter_start["retry_events"]
        for key in ("login_seconds", "download_iteration_seconds", "download_api_active_seconds", "data_sort_seconds", "redcap_existing_read_seconds", "redcap_upload_seconds", "redcap_verify_seconds", "total_account_seconds"):
            if isinstance(metrics.get(key), (int, float)):
                metrics[key] = round(float(metrics[key]), 6)
        recorder.add_attempt(metrics.copy())


# =============================================================================
# WORKBOOK WRITE-BACK HELPERS (NO STRUCTURE EDITS)
# =============================================================================


def _find_header_col_openpyxl(ws, header_row: int, candidates: list[str]) -> Optional[int]:
    wanted = {canon(c) for c in candidates}
    for col in range(1, ws.max_column + 1):
        if canon(ws.cell(header_row, col).value) in wanted:
            return col
    return None


def _find_id_row_map_openpyxl(ws, header_row: int, id_col: int) -> dict[str, int]:
    mapping = {}
    for row in range(header_row + 1, ws.max_row + 1):
        value = clean_text(ws.cell(row, id_col).value)
        if value:
            mapping[value] = row
    return mapping


def write_step_workbook_status(
    workbook: Path,
    sheet_name: str,
    header_row_0based: int,
    results: list[ProjectResult],
    allow_openpyxl_fallback: bool,
    output_dir: Path,
) -> None:
    successful = [r for r in results if r.success and r.writeback]
    failures = [r for r in results if not r.success]
    if not successful and not failures:
        return

    pending_rows = []
    for r in successful:
        pending_rows.append({"record_id": r.record_id, "status": "success", "error": ""})
    for r in failures:
        pending_rows.append({"record_id": r.record_id, "status": "failure", "error": r.error})

    header_row = header_row_0based + 1

    if _HAS_XLW:
        app = None
        try:
            app = xw.App(visible=False, add_book=False)
            wb_xw = app.books.open(str(workbook))
            wb_op = load_workbook(workbook, read_only=True, data_only=False)
            ws_op = wb_op[sheet_name]

            id_col = _find_header_col_openpyxl(ws_op, header_row, ["REDCap ID", "record id"])
            ver_col = _find_header_col_openpyxl(
                ws_op, header_row,
                ["REDCap Upload Completed and Verified", "upload completed and verified", "verified"],
            )
            last_col = _find_header_col_openpyxl(
                ws_op, header_row, ["Last Successful Upload"]
            )
            err_col = _find_header_col_openpyxl(
                ws_op, header_row, ["RedCap Upload Error", "REDCap Upload Error", "Upload Error"]
            )
            reason_col = _find_header_col_openpyxl(ws_op, header_row, ["Reason Failure"])

            if not id_col:
                raise RuntimeError("Could not locate REDCap ID for STEP write-back")
            id_rows = _find_id_row_map_openpyxl(ws_op, header_row, id_col)
            sht = wb_xw.sheets[sheet_name]

            for result in successful:
                rr = id_rows.get(result.record_id)
                if not rr:
                    continue
                wb_row = result.writeback or {}
                if ver_col and wb_row.get("verified"):
                    sht.cells(rr, ver_col).value = True
                if last_col and wb_row.get("last_successful_upload"):
                    sht.cells(rr, last_col).value = wb_row["last_successful_upload"]
                if err_col:
                    sht.cells(rr, err_col).value = None
                if reason_col:
                    sht.cells(rr, reason_col).value = None

            for result in failures:
                rr = id_rows.get(result.record_id)
                if not rr:
                    continue
                if err_col:
                    sht.cells(rr, err_col).value = result.error
                if reason_col:
                    sht.cells(rr, reason_col).value = result.error

            # Team-sheet sync metrics, row 2 headers, existing columns only.
            by_team: dict[str, list[ProjectResult]] = {}
            for result in successful:
                if result.sync_row:
                    team = clean_text(result.sync_row.get("team"))
                    if team:
                        by_team.setdefault(team, []).append(result)

            for team, team_results in by_team.items():
                if team not in wb_op.sheetnames:
                    continue
                ws_team = wb_op[team]
                idc = _find_header_col_openpyxl(ws_team, 2, ["REDCap ID", "record id"])
                lastc = _find_header_col_openpyxl(ws_team, 2, ["Last Sync Auto", "Last Synced", "Last Sync"])
                daysc = _find_header_col_openpyxl(ws_team, 2, ["Days Since Sync"])
                pctc = _find_header_col_openpyxl(ws_team, 2, ["Percent Days Worn"])
                checkc = _find_header_col_openpyxl(ws_team, 2, ["Synced ≤7d (Active)", "Synced <=7d (Active)"])
                if not idc:
                    continue
                id_rows_team = _find_id_row_map_openpyxl(ws_team, 2, idc)
                sht_team = wb_xw.sheets[team]
                for result in team_results:
                    rr = id_rows_team.get(result.record_id)
                    if not rr:
                        continue
                    sync = result.sync_row or {}
                    last_synced = safe_date(sync.get("last_synced"))
                    end_planned = safe_date(sync.get("end_date_planned"))
                    if lastc:
                        sht_team.cells(rr, lastc).value = last_synced
                    if daysc:
                        sht_team.cells(rr, daysc).value = sync.get("days_since_sync")
                    if pctc:
                        sht_team.cells(rr, pctc).value = sync.get("percent_days_worn")
                    if checkc:
                        active = bool(end_planned and end_planned > date.today())
                        recent = bool(last_synced and (date.today() - last_synced).days <= 7)
                        sht_team.cells(rr, checkc).value = "✓" if (active and recent) else None

            wb_op.close()
            wb_xw.app.calculate()
            wb_xw.save()
            wb_xw.close()
            print("[SyncStep] Workbook write-back saved through Excel/xlwings.")
            return
        except Exception as exc:
            print(f"[SyncStep] xlwings write-back failed: {exc}")
        finally:
            if app is not None:
                try:
                    app.quit()
                except Exception:
                    pass

    if not allow_openpyxl_fallback:
        write_csv(pending_rows, output_dir / "STEP_Workbook_Writeback_Pending.csv")
        print(
            "[SyncStep] Workbook write-back skipped because xlwings/Excel was unavailable or failed. "
            "No openpyxl save was attempted, protecting dynamic-array formulas. "
            "See STEP_Workbook_Writeback_Pending.csv."
        )
        return

    # Explicitly requested fallback: existing cells only, no structural edits.
    keep_vba = workbook.suffix.lower() == ".xlsm"
    wb = load_workbook(workbook, keep_vba=keep_vba)
    try:
        ws = wb[sheet_name]
        id_col = _find_header_col_openpyxl(ws, header_row, ["REDCap ID", "record id"])
        if not id_col:
            raise RuntimeError("Could not locate REDCap ID for STEP write-back")
        id_rows = _find_id_row_map_openpyxl(ws, header_row, id_col)
        ver_col = _find_header_col_openpyxl(ws, header_row, ["REDCap Upload Completed and Verified", "verified"])
        last_col = _find_header_col_openpyxl(ws, header_row, ["Last Successful Upload"])
        err_col = _find_header_col_openpyxl(ws, header_row, ["RedCap Upload Error", "REDCap Upload Error"])
        reason_col = _find_header_col_openpyxl(ws, header_row, ["Reason Failure"])
        for result in successful:
            rr = id_rows.get(result.record_id)
            if not rr:
                continue
            wb_row = result.writeback or {}
            if ver_col and wb_row.get("verified"):
                ws.cell(rr, ver_col, True)
            if last_col and wb_row.get("last_successful_upload"):
                ws.cell(rr, last_col, wb_row["last_successful_upload"])
            if err_col:
                ws.cell(rr, err_col, None)
            if reason_col:
                ws.cell(rr, reason_col, None)
        for result in failures:
            rr = id_rows.get(result.record_id)
            if not rr:
                continue
            if err_col:
                ws.cell(rr, err_col, result.error)
            if reason_col:
                ws.cell(rr, reason_col, result.error)
        wb.save(workbook)
        print("[SyncStep] Workbook write-back saved with openpyxl fallback.")
    finally:
        wb.close()


def write_rove_workbook_status(
    workbook: Path,
    sheet_name: str,
    header_row_0based: int,
    results: list[ProjectResult],
) -> None:
    successful = [r for r in results if r.success and r.writeback]
    failures = [r for r in results if not r.success]
    if not successful and not failures:
        return

    header_row = header_row_0based + 1
    keep_vba = workbook.suffix.lower() == ".xlsm"
    wb = load_workbook(workbook, keep_vba=keep_vba)
    try:
        ws = wb[sheet_name]
        id_col = _find_header_col_openpyxl(ws, header_row, ["Study ID", "REDCap ID", "record id"])
        if not id_col:
            print("[SyncRove] Could not locate Study ID column; workbook write-back skipped.")
            return
        id_rows = _find_id_row_map_openpyxl(ws, header_row, id_col)

        captured_col = _find_header_col_openpyxl(ws, header_row, ["Walking data captured"])
        last_capture_col = _find_header_col_openpyxl(ws, header_row, ["Last Successful Capture"])
        last_sync_col = _find_header_col_openpyxl(ws, header_row, ["Last Synced", "Last Sync"])
        days_col = _find_header_col_openpyxl(ws, header_row, ["Days Since Sync"])
        pct_col = _find_header_col_openpyxl(ws, header_row, ["Percent Days Worn"])
        reason_col = _find_header_col_openpyxl(ws, header_row, ["Reason Failure"])

        for result in successful:
            rr = id_rows.get(result.record_id)
            if not rr:
                continue
            item = result.writeback or {}
            if captured_col:
                ws.cell(rr, captured_col, bool(item.get("capture_complete")))
            if last_capture_col and item.get("last_successful_capture"):
                ws.cell(rr, last_capture_col, item["last_successful_capture"])
            if last_sync_col:
                ws.cell(rr, last_sync_col, item.get("last_synced"))
            if days_col:
                ws.cell(rr, days_col, item.get("days_since_sync"))
            if pct_col:
                ws.cell(rr, pct_col, item.get("percent_days_worn"))
            if reason_col:
                ws.cell(rr, reason_col, item.get('reason_failure', ''))

        for result in failures:
            rr = id_rows.get(result.record_id)
            if rr and captured_col and (result.metrics or {}).get('outcome') == 'validation-fail':
                ws.cell(rr, captured_col, False)
            if rr and reason_col:
                ws.cell(rr, reason_col, result.error)

        wb.save(workbook)
        print("[SyncRove] Workbook status updated in existing columns only.")
    finally:
        wb.close()


# =============================================================================
# OUTPUT / SUMMARY HELPERS
# =============================================================================


def final_success_attempts(attempt_df: pd.DataFrame) -> pd.DataFrame:
    if attempt_df.empty:
        return attempt_df.copy()
    success = attempt_df[attempt_df["outcome"].isin([
        "success", "download-transform-success-no-upload"
    ])].copy()
    if success.empty:
        return success
    success = success.sort_values(["project", "record_id", "attempt_sequence"])
    return success.drop_duplicates(subset=["project", "record_id"], keep="last")


def write_combined_performance_outputs(
    recorder: CombinedRecorder,
    out_dir: Path,
    project_summaries: list[dict[str, Any]],
    run_started: datetime,
    run_ended: datetime,
    caller: GarminCaller,
) -> None:
    attempt_path = out_dir / "Combined_Performance_Attempts.csv"
    iteration_path = out_dir / "Combined_Download_Iterations.csv"
    group_path = out_dir / "Combined_Performance_By_Project_Watch.csv"
    long_path = out_dir / "Combined_Performance_Manuscript_Long.csv"
    summary_path = out_dir / "Combined_Run_Summary.csv"
    rate_path = out_dir / "Combined_Rate_Limit_Events.csv"
    dict_path = out_dir / "Combined_Performance_Data_Dictionary.csv"

    attempt_df = pd.DataFrame(recorder.attempts)
    attempt_df.to_csv(attempt_path, index=False)
    pd.DataFrame(recorder.download_iterations).to_csv(iteration_path, index=False)
    pd.DataFrame(recorder.rate_limit_events).to_csv(rate_path, index=False)

    successful_final = final_success_attempts(attempt_df)
    metrics = [
        "login_seconds",
        "download_iteration_seconds",
        "download_api_active_seconds",
        "data_sort_seconds",
        "redcap_existing_read_seconds",
        "redcap_upload_seconds",
        "total_account_seconds",
        "days_requested",
        "valid_days_with_steps",
        "days_with_sleep_data",
        "days_with_any_data",
        "days_with_steps_and_sleep",
        "days_with_steps_only",
        "days_with_sleep_only",
        "days_with_no_steps_or_sleep",
        "percent_days_with_steps",
        "percent_days_with_sleep",
        "redcap_verify_seconds",
        "redcap_verify_mismatches",
        "download_api_attempts",
        "payload_json_bytes",
    ]

    grouped_rows = []
    long_rows = []
    if not attempt_df.empty:
        group_keys = (
            attempt_df[["project", "watch_category"]]
            .drop_duplicates()
            .itertuples(index=False, name=None)
        )
        for project, watch_category in group_keys:
            g_all = attempt_df[
                (attempt_df["project"] == project)
                & (attempt_df["watch_category"] == watch_category)
            ].copy()
            g_success = successful_final[
                (successful_final["project"] == project)
                & (successful_final["watch_category"] == watch_category)
            ].copy() if not successful_final.empty else pd.DataFrame()

            row = {
                "run_id": recorder.run_id,
                "project": project,
                "watch_category": watch_category,
                "unique_records_attempted": int(g_all["record_id"].nunique()),
                "total_attempts": int(len(g_all)),
                "successful_records": int(g_success["record_id"].nunique()) if not g_success.empty else 0,
                "failed_attempts": int((~g_all["outcome"].isin(["success", "download-transform-success-no-upload"])).sum()),
                "download_seconds_all_attempts": float(pd.to_numeric(g_all["download_iteration_seconds"], errors="coerce").fillna(0).sum()),
                "sort_seconds_all_attempts": float(pd.to_numeric(g_all["data_sort_seconds"], errors="coerce").fillna(0).sum()),
                "upload_seconds_all_attempts": float(pd.to_numeric(g_all["redcap_upload_seconds"], errors="coerce").fillna(0).sum()),
            }
            for metric in metrics:
                st = descriptive_stats(g_success[metric] if (not g_success.empty and metric in g_success.columns) else [])
                for stat in ("n", "mean", "sd", "median", "q1", "q3", "iqr", "min", "max"):
                    row[f"{metric}_{stat}"] = st[stat]
                    if st["n"] and stat == "mean":
                        pass
                if st["n"]:
                    long_rows.append({
                        "run_id": recorder.run_id,
                        "project": project,
                        "watch_category": watch_category,
                        "metric": metric,
                        **st,
                    })
            grouped_rows.append(row)

    pd.DataFrame(grouped_rows).to_csv(group_path, index=False)
    pd.DataFrame(long_rows).to_csv(long_path, index=False)

    total_wall = (run_ended - run_started).total_seconds()
    overall = {
        "project": "Overall",
        "run_id": recorder.run_id,
        "rows_considered": sum(int(x.get("rows_considered", 0)) for x in project_summaries),
        "records_successful": sum(int(x.get("records_successful", 0)) for x in project_summaries),
        "records_failed": sum(int(x.get("records_failed", 0)) for x in project_summaries),
        "retry_attempts": sum(int(x.get("retry_attempts", 0)) for x in project_summaries),
        "total_wall_seconds": round(total_wall, 6),
        "garmin_api_attempts": caller.counters.api_attempts,
        "garmin_api_errors": caller.counters.api_errors,
        "garmin_retry_events": caller.counters.retry_events,
        "garmin_api_active_seconds": round(caller.counters.api_active_seconds, 6),
        "garmin_throttle_sleep_seconds": round(caller.counters.throttle_sleep_seconds, 6),
        "garmin_retry_sleep_seconds": round(caller.counters.retry_sleep_seconds, 6),
        "garmin_pacing_enabled": caller.pacing_enabled,
        "redcap_upload_seconds_total": round(float(pd.to_numeric(attempt_df.get("redcap_upload_seconds", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()), 6),
        "redcap_verify_seconds_total": round(float(pd.to_numeric(attempt_df.get("redcap_verify_seconds", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()), 6),
        "redcap_verify_mismatches_total": int(pd.to_numeric(attempt_df.get("redcap_verify_mismatches", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()),
    }
    summary_rows = project_summaries + [overall]
    pd.DataFrame(summary_rows).to_csv(summary_path, index=False)

    dictionary = [
        ("project", "SyncStep or SyncRove; project-specific payload logic is never mixed."),
        ("watch_category", "Readable watch grouping used for timing summaries."),
        ("attempt_sequence", "Global sequence of account-processing attempts across both projects."),
        ("pass_name", "P1 full sweep or P2-retry final sweep."),
        ("pacing_enabled", "Whether intentional Garmin inter-call QPS pacing was enabled for this run/attempt."),
        ("configured_qps", "Configured Garmin QPS when pacing is enabled; blank when unpaced."),
        ("login_seconds", "Wall time to restore a seeded Garmin token or reuse the in-memory client."),
        ("download_iteration_seconds", "Wall time for the complete date-by-date Garmin download loop for one account."),
        ("download_api_active_seconds", "Time spent inside Garmin API calls during the account download loop; excludes throttle sleeps."),
        ("data_sort_seconds", "Time to select/sort pre-op/post-op days and construct the project-specific REDCap payload."),
        ("redcap_existing_read_seconds", "SyncRove only: time to export existing step_data repeats so relative-day rows can safely reuse repeat instances."),
        ("redcap_upload_seconds", "Wall time for the REDCap record import HTTP request/response, including REDCap retry time only when retried."),
        ("total_account_seconds", "Total wall time for one account-processing attempt."),
        ("iteration_seconds", "In Combined_Download_Iterations.csv: wall time for one requested calendar-date iteration."),
        ("days_requested", "Inclusive number of dates requested from Garmin for the account."),
        ("valid_days_with_steps", "Number of requested dates where Garmin returned a non-null totalSteps value."),
        ("days_with_sleep_data", "SyncStep: requested dates with at least one nonmissing sleep metric."),
        ("days_with_any_data", "SyncStep: requested dates with at least one nonmissing wearable metric."),
        ("days_with_steps_and_sleep", "SyncStep: dates with both steps and at least one sleep metric."),
        ("days_with_steps_only", "SyncStep: dates with steps but no sleep metric."),
        ("days_with_sleep_only", "SyncStep: dates with sleep data but no steps."),
        ("days_with_no_steps_or_sleep", "SyncStep: dates with neither steps nor a sleep metric."),
        ("percent_days_with_steps", "valid_days_with_steps / days_requested * 100; this is data availability, not verified wear adherence."),
        ("percent_days_with_sleep", "SyncStep: days_with_sleep_data / days_requested * 100."),
        ("percent_days_worn", "Backward-compatible alias for percent_days_with_steps; do not interpret as validated wear duration."),
        ("redcap_verify_seconds", "Wall time for post-import REDCap readback verification."),
        ("redcap_import_accepted", "Whether REDCap accepted the import request before post-import readback verification."),
        ("redcap_verified", "True only when post-import REDCap readback matched the submitted pipeline-managed values; null when upload is disabled."),
        ("redcap_verify_mismatches", "Count of submitted STEP fields or Rove rows that failed post-import readback comparison."),
        ("download_api_attempts", "Garmin method invocations during the account attempt, including any network retries."),
        ("payload_json_bytes", "UTF-8 JSON byte size of the project-specific REDCap record payload (SyncStep)."),
        ("payload_csv_bytes", "UTF-8 CSV byte size of the flat REDCap import payload (SyncRove)."),
    ]
    pd.DataFrame(dictionary, columns=["field", "definition"]).to_csv(dict_path, index=False)


def save_project_outputs(
    project: str,
    results: list[ProjectResult],
    project_dir: Path,
) -> None:
    """Persist download/audit artifacts even when a later REDCap step fails."""
    ensure_dir(project_dir)
    boundary_rows = [row for result in results for row in result.boundary_rows]
    if project == 'SyncRove':
        write_csv(boundary_rows, project_dir / 'REDCap_Out_Of_Window.csv', columns=[
            'record_id', 'date', 'relative_day', 'redcap_repeat_instance', 'steps', 'query_start', 'query_end', 'action'])
    successes = [r for r in results if r.success]
    failures = [r for r in results if not r.success]

    successful_payload_rows = [sanitize_record(r.payload) for r in successes if r.payload]
    attempted_payload_rows = [sanitize_record(r.payload) for r in results if r.payload]
    # Daily Garmin data are valuable source/audit data even if REDCap import or
    # readback later fails, so retain them for all completed/partial downloads.
    daily_rows = [row for r in results for row in r.daily_rows]
    sync_rows = [sanitize_record(r.sync_row) for r in successes if r.sync_row]
    failure_rows = [
        {
            "record_id": r.record_id,
            "excel_row": r.excel_row,
            "watch_category": r.watch_category,
            "retryable": r.retryable,
            "reason": r.error,
        }
        for r in failures
    ]

    if project == "SyncStep":
        write_csv(successful_payload_rows, project_dir / "REDCap_Upload.csv")
        write_csv(attempted_payload_rows, project_dir / "REDCap_Attempted_Payloads.csv")
        write_csv(daily_rows, project_dir / "STEP_Daily_Data.csv")
        write_csv(sync_rows, project_dir / "STEP_Sync_Metrics.csv")
        write_csv(failure_rows, project_dir / "Watch_Skipped.csv")
    else:
        successful_redcap_rows = [
            sanitize_record(row)
            for r in successes
            for row in (r.redcap_rows or [])
        ]
        attempted_redcap_rows = [
            sanitize_record(row)
            for r in results
            for row in (r.redcap_rows or [])
        ]
        write_csv(successful_redcap_rows or successful_payload_rows, project_dir / "REDCap_Upload.csv")
        write_csv(attempted_redcap_rows or attempted_payload_rows, project_dir / "REDCap_Attempted_Payloads.csv")
        write_csv(daily_rows, project_dir / "Garmin_Walking_Daily.csv")
        write_csv(sync_rows, project_dir / "Garmin_Walking_Sync.csv")
        write_csv(failure_rows, project_dir / "Garmin_Walking_Skipped.csv")


# =============================================================================
# PROJECT RUNNERS
# =============================================================================


def load_latest_seed_audit_status(search_dirs: Iterable[Path]) -> dict[str, dict[str, Any]]:
    """Load hashed seed outcomes when available; never reads/stores credential columns."""
    candidates: list[Path] = []
    seen_paths = set()
    for directory in search_dirs:
        try:
            directory = Path(directory).expanduser().resolve()
        except Exception:
            continue
        if not directory.is_dir():
            continue
        for pattern in ("*Garmin_Token_Seed_Audit_*.csv", "*Token_Seed_Audit_*.csv"):
            for path in directory.glob(pattern):
                key = str(path)
                if key not in seen_paths and path.is_file():
                    seen_paths.add(key)
                    candidates.append(path)

    # Oldest -> newest so the latest observed outcome for each hash wins.
    candidates.sort(key=lambda p: (p.stat().st_mtime, p.name))
    latest: dict[str, dict[str, Any]] = {}
    for path in candidates:
        try:
            df = pd.read_csv(path, dtype=object)
        except Exception:
            continue
        if "account_hash" not in df.columns:
            continue
        for _, row in df.iterrows():
            aid = clean_text(row.get("account_hash"))
            if not aid:
                continue
            latest[aid] = {
                "seed_audit_file": path.name,
                "last_seed_action": clean_text(row.get("action")),
                "last_seed_outcome": clean_text(row.get("outcome")),
                # Intentionally do not copy the audit error text into the combined
                # output; outcome is sufficient and avoids propagating any chance
                # of sensitive text from an older/custom seeder.
            }
    return latest


def preflight_token_coverage(
    project: str,
    rows: pd.DataFrame,
    username_col: Any,
    token_dir: Path,
    password_col: Optional[Any] = None,
    seed_audit: Optional[dict[str, dict[str, Any]]] = None,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Audit token coverage without persisting plaintext credentials."""
    grouped: dict[str, list[pd.Series]] = {}
    for _, row in rows.iterrows():
        username = clean_text(row.get(username_col)).lower()
        if username:
            grouped.setdefault(username, []).append(row)

    present = 0
    missing: list[str] = []
    malformed: list[str] = []
    details: list[dict[str, Any]] = []
    missing_with_password = 0
    missing_without_password = 0

    for username in sorted(grouped):
        account_rows = grouped[username]
        aid = account_hash(username)
        path = token_path(token_dir, username)
        if not path.exists():
            status = "missing"
            missing.append(aid)
        elif not token_file_structurally_valid(path):
            status = "malformed"
            malformed.append(aid)
        else:
            status = "usable"
            present += 1

        password_present = None
        if password_col is not None:
            password_present = any(clean_text(row.get(password_col)) for row in account_rows)
        if status == "missing" and password_present is True:
            missing_with_password += 1
        elif status == "missing" and password_present is False:
            missing_without_password += 1

        historical = (seed_audit or {}).get(aid, {})
        details.append({
            "project": project,
            "account_hash": aid,
            "rows_using_account": len(account_rows),
            "token_status": status,
            "password_column_available": password_col is not None,
            "password_present_in_any_row": password_present,
            "last_seed_action": historical.get("last_seed_action"),
            "last_seed_outcome": historical.get("last_seed_outcome"),
            "seed_audit_file": historical.get("seed_audit_file"),
        })

    missing_outcomes: dict[str, int] = {}
    for item in details:
        if item["token_status"] == "missing":
            outcome = clean_text(item.get("last_seed_outcome")) or "no_seed_audit_match"
            missing_outcomes[outcome] = missing_outcomes.get(outcome, 0) + 1

    result = {
        "project": project,
        "unique_accounts": len(grouped),
        "valid_seed_tokens": present,
        "missing_token_hashes": missing,
        "malformed_token_hashes": malformed,
        "missing_tokens_with_password_present": missing_with_password,
        "missing_tokens_with_password_absent": missing_without_password,
        "missing_token_seed_outcomes": missing_outcomes,
        "token_dir": str(token_dir),
    }
    print(
        f"[{project}] Token preflight: {present}/{len(grouped)} unique workbook accounts "
        f"have structurally usable seeded tokens."
    )
    if missing:
        print(
            f"[{project}] Missing token count: {len(missing)}; "
            f"password currently present for {missing_with_password}, absent for {missing_without_password}."
        )
        if missing_outcomes:
            print(f"[{project}] Historical seed outcomes for missing tokens: {missing_outcomes}")
    if malformed:
        print(f"[{project}] Malformed token count: {len(malformed)}.")
    return result, details


def run_project_rows(
    project: str,
    rows: pd.DataFrame,
    cols: dict[str, Any],
    token_dir: Path,
    sessions: GarminSessionManager,
    caller: GarminCaller,
    recorder: CombinedRecorder,
    redcap: Optional[RedcapClient],
    schema: Optional[RedcapSchema],
    step_stats_mode: str,
    retry_wait_seconds: int,
    project_output_dir: Path,
) -> tuple[list[ProjectResult], dict[str, Any]]:
    print(f"\n{'=' * 80}\n{project}: PASS 1 FULL SWEEP ({len(rows)} row(s))\n{'=' * 80}")

    final_by_excel_row: dict[int, ProjectResult] = {}
    retry_rows: list[pd.Series] = []
    global_attempt_sequence = len(recorder.attempts)

    for index, (_, row) in enumerate(rows.iterrows(), 1):
        global_attempt_sequence += 1
        record_col = cols["record_id"] if project == "SyncStep" else cols["study_id"]
        rid = clean_text(row.get(record_col)) or f"excel-row-{int(row['__excel_row__'])}"
        print(f"[{project} P1 {index}/{len(rows)}] record={rid}", flush=True)
        if project == "SyncStep":
            result = process_step_account(
                row, cols, token_dir, sessions, caller, recorder,
                redcap, schema, step_stats_mode, "P1", global_attempt_sequence,
                project_output_dir,
            )
        else:
            result = process_rove_account(
                row, cols, token_dir, sessions, caller, recorder,
                redcap, schema, "P1", global_attempt_sequence,
            )

        final_by_excel_row[int(row["__excel_row__"])] = result
        if result.retryable:
            retry_rows.append(row)
        elif not result.success:
            print(f"  -> FAILED: {result.error}")
        else:
            print("  -> success")

    retry_attempts = 0
    if retry_rows:
        wait_seconds = max(0, int(retry_wait_seconds))
        print(
            f"\n[{project}] {len(retry_rows)} retryable account(s) remain after P1. "
            f"Waiting {wait_seconds / 60:.1f} minute(s), then doing ONE final retry sweep."
        )
        if wait_seconds:
            time.sleep(wait_seconds)
        sessions.clear()

        for index, row in enumerate(retry_rows, 1):
            retry_attempts += 1
            global_attempt_sequence += 1
            record_col = cols["record_id"] if project == "SyncStep" else cols["study_id"]
            rid = clean_text(row.get(record_col)) or f"excel-row-{int(row['__excel_row__'])}"
            print(f"[{project} P2 {index}/{len(retry_rows)}] record={rid}", flush=True)
            if project == "SyncStep":
                result = process_step_account(
                    row, cols, token_dir, sessions, caller, recorder,
                    redcap, schema, step_stats_mode, "P2-retry", global_attempt_sequence,
                    project_output_dir,
                )
            else:
                result = process_rove_account(
                    row, cols, token_dir, sessions, caller, recorder,
                    redcap, schema, "P2-retry", global_attempt_sequence,
                )
            final_by_excel_row[int(row["__excel_row__"])] = result
            print("  -> success" if result.success else f"  -> FAILED: {result.error}")

    final_results = list(final_by_excel_row.values())
    save_project_outputs(project, final_results, project_output_dir)

    summary = {
        "project": project,
        "run_id": recorder.run_id,
        "rows_considered": len(rows),
        "records_successful": sum(1 for r in final_results if r.success),
        "records_failed": sum(1 for r in final_results if not r.success),
        "retry_attempts": retry_attempts,
        "token_dir": str(token_dir),
        "output_dir": str(project_output_dir),
    }
    return final_results, summary


# =============================================================================
# MAIN
# =============================================================================


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Run SyncStep + SyncRove Garmin downloads and project-specific REDCap uploads "
            "in one token-only seeded run with combined timing metrics."
        )
    )
    parser.add_argument("--step-workbook", help="SyncStep source workbook")
    parser.add_argument("--step-sheet", help="SyncStep worksheet (default Master Spreadsheet/autodetect)")
    parser.add_argument(
        "--step-seed-folder",
        help="SyncStep Seed folder or .garmin_tokens_v3 folder",
    )
    parser.add_argument("--rove-workbook", help="SyncRove source workbook")
    parser.add_argument("--rove-sheet", help="SyncRove worksheet (autodetect if omitted)")
    parser.add_argument(
        "--rove-seed-folder",
        help="SyncRove Seed folder or .garmin_tokens_v3 folder",
    )
    parser.add_argument(
        "--output-dir",
        help="Parent directory for timestamped combined results (default script directory)",
    )
    parser.add_argument(
        "--pacing",
        type=parse_bool,
        default=True,
        metavar="T/F",
        help=(
            "Enable intentional Garmin inter-call QPS pacing (default True). "
            "Set --pacing false to remove QPS/jitter sleeps. Network/account retry "
            "backoff remains enabled for safety."
        ),
    )
    parser.add_argument(
        "--no-pacing",
        dest="pacing",
        action="store_false",
        help="Alias for --pacing false.",
    )
    parser.add_argument(
        "--qps",
        type=float,
        default=DEFAULT_DATA_QPS,
        help=f"Global Garmin data-call QPS across BOTH projects (default {DEFAULT_DATA_QPS})",
    )
    parser.add_argument(
        "--retry-qps",
        type=float,
        default=DEFAULT_RETRY_QPS,
        help=f"Garmin QPS while retrying transient calls (default {DEFAULT_RETRY_QPS})",
    )
    parser.add_argument(
        "--network-retries",
        type=int,
        default=DEFAULT_NETWORK_RETRIES,
        help=f"In-place retries for non-429 transient network failures (default {DEFAULT_NETWORK_RETRIES})",
    )
    parser.add_argument(
        "--retry-wait-seconds",
        type=int,
        default=DEFAULT_RETRY_WAIT_SECONDS,
        help=(
            "Wait before one final account-level retry sweep for 429/transient failures "
            f"(default {DEFAULT_RETRY_WAIT_SECONDS})"
        ),
    )
    parser.add_argument(
        "--stats-mode",
        choices=["auto", "optimized", "legacy"],
        default="auto",
        help="SyncStep vivoactive daily-summary compatibility mode",
    )
    parser.add_argument('--hrv-mode', choices=['auto', 'overnight-ms', 'status', 'off'], default='auto',
                        help='STEP HRV upload concept. Auto inspects REDCap metadata; ambiguous HRV fields are omitted, preserving existing values. Raw overnight ms and categorical status remain separate.')
    parser.add_argument(
        "--skip-completed",
        action="store_true",
        help=(
            "Incremental mode: skip already-verified STEP rows and already-captured Rove rows. "
            "Default is FULL SWEEP of all eligible rows, as requested for the timing run."
        ),
    )
    parser.add_argument(
        "--no-upload",
        action="store_true",
        help="Download/transform only; do not write to REDCap.",
    )
    parser.add_argument(
        "--no-writeback",
        action="store_true",
        help="Do not modify either source workbook after the run.",
    )
    parser.add_argument(
        "--allow-openpyxl-step-writeback",
        action="store_true",
        help=(
            "If Excel/xlwings is unavailable, allow openpyxl to save the STEP workbook. "
            "Off by default to protect dynamic-array formulas."
        ),
    )
    parser.add_argument(
        "--sync-rove-back",
        action="store_true",
        help=(
            "After timed execution has ended, automatically copy the staged Rove "
            "workbook/status and refreshed token files back to the mounted-drive source. "
            "Use only if the mount will still be available after the run."
        ),
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Skip the final RUN confirmation prompt.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    garmin_version = ensure_supported_garminconnect()

    # -------------------------
    # Select workbooks + tokens
    # -------------------------
    print("\n[1/4] Select the SyncStep Excel workbook.")
    step_workbook = choose_file(args.step_workbook, "1 of 4 - Select SyncStep workbook")
    step_sheet = choose_sheet(step_workbook, args.step_sheet, "SyncStep")
    print(f"  SyncStep workbook: {step_workbook}")

    print("\n[2/4] Select the SyncStep Seed folder OR its .garmin_tokens_v3 folder.")
    step_seed_selected = choose_folder(args.step_seed_folder, "2 of 4 - Select SyncStep Seed/token folder")
    step_token_dir = resolve_token_dir(step_seed_selected)
    print(f"  SyncStep token folder resolved to: {step_token_dir}")

    print("\n[3/4] Select the SyncRove Excel workbook (the sheet with Study ID / username / DOS).")
    rove_workbook_source = choose_file(args.rove_workbook, "3 of 4 - Select SyncRove workbook")
    rove_sheet = choose_sheet(rove_workbook_source, args.rove_sheet, "SyncRove")
    print(f"  SyncRove workbook: {rove_workbook_source}")

    print("\n[4/4] Select the SyncRove Seed folder OR its .garmin_tokens_v3 folder.")
    rove_seed_selected = choose_folder(args.rove_seed_folder, "4 of 4 - Select SyncRove Seed/token folder")
    rove_token_dir_source = resolve_token_dir(rove_seed_selected)
    print(f"  SyncRove token folder resolved to: {rove_token_dir_source}")

    # ---------------------------------------------------------------------
    # LOCAL ROVE STAGING -- deliberately BEFORE manuscript timing begins.
    # The mounted drive is no longer used by Rove during the timed interval.
    # ---------------------------------------------------------------------
    staging_parent = ensure_dir(Path(__file__).resolve().parent / "_local_rove_staging")
    rove_workbook, rove_token_dir, rove_stage_dir = stage_rove_inputs(
        rove_workbook_source,
        rove_token_dir_source,
        staging_parent,
    )

    process_all = not args.skip_completed
    step_header_row, step_rows, step_cols = load_step_rows(step_workbook, step_sheet, process_all)
    rove_header_row, rove_rows, rove_cols = load_rove_rows(rove_workbook, rove_sheet, process_all)

    # -------------------------
    # Output folder
    # -------------------------
    output_parent = Path(args.output_dir).expanduser().resolve() if args.output_dir else Path(__file__).resolve().parent
    out_dir = ensure_dir(output_parent / f"Combined_Garmin_REDCap_Run_{utc_stamp()}")
    step_out = ensure_dir(out_dir / "SyncStep")
    rove_out = ensure_dir(out_dir / "SyncRove")
    ownership_audit = annotate_assignment_windows([
        ('SyncStep', step_rows, step_cols), ('SyncRove', rove_rows, rove_cols)])
    write_csv(ownership_audit, out_dir / 'Assignment_Window_Audit.csv')

    # -------------------------
    # Token coverage preflight
    # -------------------------
    step_seed_audit = load_latest_seed_audit_status([step_workbook.parent, step_token_dir.parent])
    rove_seed_audit = load_latest_seed_audit_status([rove_workbook_source.parent, rove_token_dir_source.parent])
    step_token_status, step_token_details = preflight_token_coverage(
        "SyncStep", step_rows, step_cols["username"], step_token_dir,
        step_cols.get("password"), step_seed_audit,
    )
    rove_token_status, rove_token_details = preflight_token_coverage(
        "SyncRove", rove_rows, rove_cols["username"], rove_token_dir,
        rove_cols.get("password"), rove_seed_audit,
    )
    pd.DataFrame([step_token_status, rove_token_status]).to_json(
        out_dir / "Token_Preflight.json", orient="records", indent=2
    )
    pd.DataFrame(step_token_details + rove_token_details).to_csv(
        out_dir / "Token_Preflight_Accounts.csv", index=False
    )

    # A completely wrong token-folder selection should never progress to REDCap
    # prompts or the manuscript run. Partial coverage is allowed because some
    # accounts may legitimately have failed seeding.
    zero_token_projects = [
        item["project"]
        for item in (step_token_status, rove_token_status)
        if int(item.get("valid_seed_tokens", 0)) == 0
        and int(item.get("unique_accounts", 0)) > 0
    ]
    if zero_token_projects:
        raise RuntimeError(
            "Token preflight found ZERO usable seeded tokens for: "
            + ", ".join(zero_token_projects)
            + ". This almost certainly means the wrong Seed/.garmin_tokens_v3 folder "
              "was selected. No REDCap upload was attempted. Re-run and select the "
              "token folder created by that project's seeding run."
        )

    # -------------------------
    # REDCap preflight
    # -------------------------
    step_redcap: Optional[RedcapClient] = None
    rove_redcap: Optional[RedcapClient] = None
    step_schema: Optional[RedcapSchema] = None
    rove_schema: Optional[RedcapSchema] = None

    if not args.no_upload:
        step_token = get_redcap_token("SyncStep", no_upload=False)
        rove_token = get_redcap_token("SyncRove", no_upload=False)
        assert step_token and rove_token
        step_redcap = RedcapClient("SyncStep", step_token)
        rove_redcap = RedcapClient("SyncRove", rove_token)
        print("\nValidating SyncStep REDCap metadata...")
        step_schema = build_step_redcap_schema(step_redcap, args.hrv_mode)
        print(f'  STEP HRV upload mode: {step_schema.hrv_mode}')
        if step_schema.hrv_warning:
            print(f'  STEP HRV: {step_schema.hrv_warning}')
        (out_dir / 'STEP_HRV_Preflight.json').write_text(json.dumps({
            'requested_mode': args.hrv_mode, 'effective_mode': step_schema.hrv_mode,
            'warning': step_schema.hrv_warning, 'status_choices': step_schema.hrv_status_choices}, indent=2), encoding='utf-8')
        print(f"  SyncStep REDCap fields: {len(step_schema.field_names)}; primary key={step_schema.record_field}")
        print("  Validating SyncStep REDCap DAG mappings used by this workbook...")
        step_dag_preflight = validate_step_dag_map(step_redcap, step_rows, step_cols)
        (out_dir / "STEP_DAG_Preflight.json").write_text(
            json.dumps(step_dag_preflight, indent=2), encoding="utf-8"
        )
        if step_dag_preflight.get("status") == "invalid":
            invalid = step_dag_preflight.get("invalid_mappings", {})
            detail = ", ".join(f"{team}->{dag}" for team, dag in invalid.items())
            raise RedcapError(
                "SyncStep DAG preflight found mappings that do not exist in REDCap: " + detail
            )
        print("Validating SyncRove REDCap metadata and storage model...")
        rove_candidate_record_ids = [
            clean_text(row.get(rove_cols["study_id"]))
            for _, row in rove_rows.iterrows()
            if clean_text(row.get(rove_cols["study_id"]))
        ]
        rove_schema = build_rove_redcap_schema(rove_redcap, rove_candidate_record_ids)
        print(
            f"  SyncRove REDCap mode: LONG / REPEATING DAILY ROWS; "
            f"primary key={rove_schema.record_field}; "
            f"relative-day field={rove_schema.relative_day_field}; "
            f"steps field={rove_schema.steps_field}; "
            f"completion field left blank={rove_schema.completion_field}; "
            f"repeat instrument={rove_schema.repeat_instrument}; "
            f"event={rove_schema.repeat_event}; "
            f"repeating event={rove_schema.repeat_is_event}"
        )

    effective_stats_mode, step_lib_info = resolve_step_stats_mode(args.stats_mode)

    # -------------------------
    # User-visible preflight summary
    # -------------------------
    print("\n" + "=" * 80)
    print("COMBINED RUN PREFLIGHT")
    print("=" * 80)
    print(f"garminconnect: {garmin_version}")
    print(f"Garmin pacing: {'ON' if args.pacing else 'OFF'}")
    print(f"Global Garmin data QPS: {args.qps if args.pacing else 'UNPACED'}")
    print(f"Retry Garmin QPS: {args.retry_qps if args.pacing else 'UNPACED (retry backoff still active)'}")
    print(f"Token-only Garmin auth: YES (no fresh credential logins)")
    print(f"Process mode: {'FULL SWEEP' if process_all else 'INCREMENTAL / SKIP COMPLETED'}")
    print(f"REDCap upload: {'OFF' if args.no_upload else 'ON'}")
    print(f"Workbook write-back: {'OFF' if args.no_writeback else 'ON'}")
    print(f"SyncStep workbook: {step_workbook}")
    print(f"SyncStep sheet: {step_sheet}")
    print(f"SyncStep token folder: {step_token_dir}")
    print(f"SyncStep rows: {len(step_rows)}")
    print(f"SyncStep stats mode: {effective_stats_mode}")
    print(f"SyncRove source workbook (mounted): {rove_workbook_source}")
    print(f"SyncRove STAGED workbook used for run: {rove_workbook}")
    print(f"SyncRove sheet: {rove_sheet}")
    print(f"SyncRove source token folder (mounted): {rove_token_dir_source}")
    print(f"SyncRove STAGED token folder used for run: {rove_token_dir}")
    print(f"SyncRove local staging folder: {rove_stage_dir}")
    print(f"SyncRove rows: {len(rove_rows)}")
    print(f"Ownership audit: {len(ownership_audit)} full-log assignments; {sum(bool(x['error']) for x in ownership_audit)} blocked; {sum(bool(x['warnings']) for x in ownership_audit)} flagged")
    print('Windows: expected end / recorded return / next assignment minus one day / today, whichever is earliest.')
    if rove_schema is not None:
        print(f"SyncRove REDCap mode: {rove_schema.mode}")
        if rove_schema.mode == "long_repeating":
            print(
                f"SyncRove REDCap daily form: instrument={rove_schema.repeat_instrument}, "
                f"event={rove_schema.repeat_event}, relative_day={rove_schema.relative_day_field}, "
                f"steps={rove_schema.steps_field}, completion field LEFT BLANK"
            )
    print(f"Combined output folder: {out_dir}")
    print("=" * 80)
    print(
        "SyncRove mounted-drive inputs are fully staged locally. "
        "The timed run will not read from the mounted network drive.\n"
        "If REDCap is reachable without the VPN, you may disconnect the VPN NOW, "
        "before typing RUN."
    )

    if not args.yes:
        prompt = "Type RUN to begin both projects: "
        if input(prompt).strip().upper() != "RUN":
            print("Cancelled before Garmin data calls or REDCap uploads.")
            return 0

    # -------------------------------------------------------------------------
    # START OF MANUSCRIPT / PIPELINE WALL-CLOCK TIMING
    #
    # Deliberately starts AFTER all human interaction and preflight work:
    #   - workbook / seed-folder selection
    #   - REDCap token entry
    #   - REDCap metadata validation
    #   - RUN confirmation
    # This keeps total_wall_seconds limited to the automated execution itself.
    # -------------------------------------------------------------------------
    run_started = datetime.now(timezone.utc).astimezone()
    run_id = run_started.strftime("%Y%m%dT%H%M%S%z") + f"_pid{os.getpid()}"
    print(f"\nTIMED EXECUTION STARTED: {run_started.replace(microsecond=0).isoformat()}")

    # -------------------------
    # Shared run-scoped objects
    # -------------------------
    recorder = CombinedRecorder(run_id=run_id)
    caller = GarminCaller(
        qps=args.qps,
        retry_qps=args.retry_qps,
        max_network_retries=args.network_retries,
        jitter_seconds=DEFAULT_JITTER_SECONDS,
        rate_limit_events=recorder.rate_limit_events,
        pacing_enabled=args.pacing,
    )
    sessions = GarminSessionManager(recorder.rate_limit_events)

    project_summaries: list[dict[str, Any]] = []

    # -------------------------
    # SyncStep
    # -------------------------
    step_results, step_summary = run_project_rows(
        project="SyncStep",
        rows=step_rows,
        cols=step_cols,
        token_dir=step_token_dir,
        sessions=sessions,
        caller=caller,
        recorder=recorder,
        redcap=step_redcap,
        schema=step_schema,
        step_stats_mode=effective_stats_mode,
        retry_wait_seconds=args.retry_wait_seconds,
        project_output_dir=step_out,
    )
    project_summaries.append(step_summary)

    # Clear live Garmin clients between projects while preserving global QPS/counters.
    sessions.clear()

    # -------------------------
    # SyncRove
    # -------------------------
    rove_results, rove_summary = run_project_rows(
        project="SyncRove",
        rows=rove_rows,
        cols=rove_cols,
        token_dir=rove_token_dir,
        sessions=sessions,
        caller=caller,
        recorder=recorder,
        redcap=rove_redcap,
        schema=rove_schema,
        step_stats_mode=effective_stats_mode,
        retry_wait_seconds=args.retry_wait_seconds,
        project_output_dir=rove_out,
    )
    project_summaries.append(rove_summary)

    # -------------------------
    # Workbook write-backs
    # -------------------------
    if not args.no_writeback and not args.no_upload:
        print("\n=== WORKBOOK WRITE-BACK ===")
        try:
            write_step_workbook_status(
                step_workbook,
                step_sheet,
                step_header_row,
                step_results,
                allow_openpyxl_fallback=args.allow_openpyxl_step_writeback,
                output_dir=step_out,
            )
        except Exception as exc:
            print(f"[SyncStep] Workbook write-back failed but data run will remain complete: {exc}")
            traceback.print_exc()
        try:
            write_rove_workbook_status(
                rove_workbook,
                rove_sheet,
                rove_header_row,
                rove_results,
            )
        except Exception as exc:
            print(f"[SyncRove] Workbook write-back failed but data run will remain complete: {exc}")
            traceback.print_exc()

    # -------------------------
    # Combined outputs
    # -------------------------
    run_ended = datetime.now(timezone.utc).astimezone()
    write_combined_performance_outputs(
        recorder,
        out_dir,
        project_summaries,
        run_started,
        run_ended,
        caller,
    )

    run_manifest = {
        "run_id": run_id,
        "started_at": run_started.replace(microsecond=0).isoformat(),
        "ended_at": run_ended.replace(microsecond=0).isoformat(),
        "garminconnect_version": garmin_version,
        "token_only_auth": True,
        "fresh_credential_logins_attempted": 0,
        "pacing_enabled": args.pacing,
        "data_qps": args.qps if args.pacing else None,
        "retry_qps": args.retry_qps if args.pacing else None,
        "retry_backoff_enabled": True,
        "missing_data_policy": "metric-independent; missing remains blank; true zeros retained",
        "ownership_policy": "full unfiltered logs across both projects; serial OR account reuse; next Start Date minus one day; ambiguous assignments blocked",
        "return_date_policy": "exact date inclusive unless reassigned that day; month-only excludes entire return month",
        "total_intensity_definition": "moderateIntensityMinutes + 2 * vigorousIntensityMinutes; both components required",
        "hrv_numeric_source_priority": "dailySleepDTO.avgSleepHRV then hrvSummary.lastNightAvg (milliseconds)",
        "hrv_status_source": "hrvSummary.status; categorical; metadata choice codes only",
        "hrv_upload_mode": step_schema.hrv_mode if step_schema else 'overnight-ms export only',
        "step_redcap_overwrite_policy": "pipeline-managed wearable fields blank-initialized and imported with overwrite; post-import readback required",
        "network_retries": args.network_retries,
        "retry_wait_seconds": args.retry_wait_seconds,
        "process_all_rows": process_all,
        "redcap_upload_enabled": not args.no_upload,
        "workbook_writeback_enabled": not args.no_writeback,
        "step": {
            "workbook": str(step_workbook),
            "sheet": step_sheet,
            "token_dir": str(step_token_dir),
            "rows": len(step_rows),
            "stats_mode_requested": args.stats_mode,
            "stats_mode_effective": effective_stats_mode,
            "postop_day_max": STEP_POSTOP_DAYS,
            "preop_selection": "10 most recent pre-op dates with any wearable metric",
            **step_lib_info,
        },
        "rove": {
            "source_workbook": str(rove_workbook_source),
            "staged_workbook_used_for_timed_run": str(rove_workbook),
            "sheet": rove_sheet,
            "source_token_dir": str(rove_token_dir_source),
            "staged_token_dir_used_for_timed_run": str(rove_token_dir),
            "local_staging_dir": str(rove_stage_dir),
            "rows": len(rove_rows),
            "redcap_record_field": rove_schema.record_field if rove_schema else None,
            "redcap_mode": rove_schema.mode if rove_schema else None,
            "redcap_relative_day_field": rove_schema.relative_day_field if rove_schema else None,
            "redcap_steps_field": rove_schema.steps_field if rove_schema else None,
            "redcap_completion_field": rove_schema.completion_field if rove_schema else None,
            "garmin_query_policy": "Start Date through earliest expected end, recorded return, next assignment minus one day, and today",
            "redcap_max_relative_day": ROVE_POSTOP_DAYS,
            "missing_expected_end_fallback_postop_days": ROVE_DEFAULT_POSTOP_DAYS,
            "redcap_min_relative_day": ROVE_MIN_RELATIVE_DAY,
            "redcap_repeat_instrument": rove_schema.repeat_instrument if rove_schema else None,
            "redcap_repeat_event": rove_schema.repeat_event if rove_schema else None,
            "redcap_repeat_is_event": rove_schema.repeat_is_event if rove_schema else None,
            "redcap_preop_step_fields_found": len(rove_schema.preop_steps_fields) if rove_schema else None,
            "redcap_postop_step_fields_found": len(rove_schema.postop_steps_fields) if rove_schema else None,
        },
        "combined_output_dir": str(out_dir),
    }
    (out_dir / "Run_Manifest.json").write_text(
        json.dumps(run_manifest, indent=2), encoding="utf-8"
    )

    print("\n" + "=" * 80)
    print("COMBINED RUN COMPLETE")
    print("=" * 80)
    for summary in project_summaries:
        print(
            f"{summary['project']}: {summary['records_successful']} successful, "
            f"{summary['records_failed']} failed, {summary['retry_attempts']} retry attempt(s)"
        )
    total_wall_seconds = (run_ended - run_started).total_seconds()
    print(f"Timed automated wall clock: {total_wall_seconds:.2f} seconds ({total_wall_seconds / 60.0:.2f} minutes)")
    print(f"Garmin API attempts: {caller.counters.api_attempts}")
    print(f"Garmin API errors: {caller.counters.api_errors}")
    print(f"Combined outputs: {out_dir}")
    print("Key timing files:")
    print(f"  {out_dir / 'Combined_Performance_Attempts.csv'}")
    print(f"  {out_dir / 'Combined_Download_Iterations.csv'}")
    print(f"  {out_dir / 'Combined_Performance_By_Project_Watch.csv'}")
    print(f"  {out_dir / 'Combined_Performance_Manuscript_Long.csv'}")
    print(f"  {out_dir / 'Combined_Run_Summary.csv'}")
    print("=" * 80)

    # ---------------------------------------------------------------------
    # OPTIONAL POST-TIMING COPY-BACK
    # This happens only AFTER run_ended has been recorded, so mounted-drive
    # transfer time is never included in manuscript wall-clock timing.
    # ---------------------------------------------------------------------
    should_sync_back = bool(args.sync_rove_back)
    if not args.yes and not args.sync_rove_back:
        print(
            "\nSyncRove ran entirely from local staging. To preserve any workbook status "
            "updates and refreshed Garmin token files on the mounted drive, reconnect "
            "the VPN/mount now if needed."
        )
        answer = input(
            "Type SYNC to copy the staged Rove workbook/tokens back to the mounted drive, "
            "or press Enter to leave them staged locally: "
        ).strip().upper()
        should_sync_back = answer == "SYNC"

    if should_sync_back:
        try:
            sync_rove_stage_back(
                staged_workbook=rove_workbook,
                staged_token_dir=rove_token_dir,
                source_workbook=rove_workbook_source,
                source_token_dir=rove_token_dir_source,
                copy_workbook=(not args.no_writeback and not args.no_upload),
            )
            print("[SyncRove] Post-timing copy-back complete.")
        except Exception as exc:
            print(
                "[SyncRove] Post-timing copy-back FAILED. The timed run is still valid and "
                f"the local staged files are preserved at {rove_stage_dir}. Error: {exc}"
            )
    else:
        print(
            f"[SyncRove] Local staged workbook/tokens retained at: {rove_stage_dir}"
        )
        if args.yes:
            print(
                "[SyncRove] Batch mode does not prompt for copy-back. Re-run with "
                "--sync-rove-back if the mounted drive will be available after the run."
            )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
