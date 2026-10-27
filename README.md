# Combined Garmin-to-REDCap Program: Leveraging AI for Perioperative Activity Tracking

This repository contains the combined program used in the operational evaluation reported in:

> Daly J, Skalicky M, Arai E, et al. *Leveraging Artificial Intelligence to Reduce Manual Burden in Perioperative Activity Data Collection.* Manuscript under review.

The program retrieves daily activity data from study-managed Garmin Connect accounts, aligns each day to the date of surgery, prepares project-specific REDCap imports, uploads them, and reads the stored values back to verify agreement. It ran two projects in one timed execution while keeping each project's data model separate:

- **Project 1 (STEP; Activity and Movement Monitoring in Peri-Operative Setting):** 14 daily variables (steps, sleep stages, heart rate, HRV, respiration, pulse oximetry, intensity minutes) stored in fixed REDCap fields for the 10 most recent preoperative days with data and postoperative days 1-90.
- **Project 2 (ACCS; Ambulation and Cognition after Cardiac Surgery):** daily step counts stored as REDCap repeating instances, one per requested day.

The program was developed with large language model assistance (ChatGPT, OpenAI; Claude, Anthropic) under investigator direction, as described in the manuscript. Routine execution does not call any language model.

## Important notes before use

- **Unofficial interface.** Garmin data are retrieved with the community-developed [`python-garminconnect`](https://github.com/cyberjunky/python-garminconnect) package, which is not affiliated with Garmin. Review Garmin's [Terms of Use](https://www.garmin.com/en-US/legal/terms-of-use/) and your institution's requirements before use. The interface may change or stop working.
- **Approvals.** Use only with accounts and data covered by your ethics approval and participant consent.
- **No credentials in this repository.** Garmin passwords, session tokens, REDCap API tokens, study workbooks, and run outputs must never be committed. The `.gitignore` excludes them.

## Requirements

- Python 3.11 or 3.12
- `pip install -r requirements.txt`
- Optional: `xlwings` (with Microsoft Excel) for formula-safe write-back to the STEP workbook

## Inputs

1. **Tracking workbooks**, one per project, maintained by the study team. Templates with made-up rows are in `examples/`:
   - `STEP_tracking_template.xlsx`: REDCap ID, Watch Email, Watch Password, Watch Model, Watch Serial Number, REDCap Team, Patient Enrollment Date, Patient Surgery Date, Patient Completion Date, Device back, REDCap Upload Completed and Verified
   - `ACCS_tracking_template.xlsx`: Study ID, username, password, Vivofit4 Serial Number, DOS, Start Date, Expected End Date, Return Date, Walking data captured
2. **Seeded Garmin session tokens.** The combined program is token-only: it never performs a fresh username and password login. Each account is first authenticated once ("seeded") with the project's seeding script, which reads credentials from the tracking workbook, logs in, verifies the new session from disk, and saves a reusable token in a `.garmin_tokens_v3` folder next to the workbook. Token files are named by a SHA-256 hash of the account, so no email address appears in file names, audit files, or summaries. Existing tokens are never deleted when a login is rate limited or fails.
   - Project 1: `python step_v3_verify_and_seed.py --file "/path/to/STEP.xlsx"`. Verifies every existing token online and reseeds missing or rejected ones. Add `--dry-run` for a local inventory with no Garmin contact, or `--stop-on-rate-limit` to pause after the first 429 response.
   - Project 2: `python seed_syncrove_tokens.py --file "/path/to/ACCS.xlsx" --auth modern`. Add `--verify-existing` for token-only checks, `--prompt-mfa` for accounts with multifactor authentication, or `--dry-run` for no network calls.
   - Seeding is rate limited by Garmin; spread large batches over time.
3. **REDCap API tokens**, supplied through environment variables (never hard-coded):
   - Project 1: `SYNCSTEP_REDCAP_TOKEN` (or `REDCAP_API_TOKEN`)
   - Project 2: `SYNCROVE_REDCAP_TOKEN` (or `ROVE_REDCAP_TOKEN`)
   - If unset, the program prompts with hidden input.
   - Set `REDCAP_API_URL` to your institution's REDCap API endpoint.

The REDCap field names expected for each project are described in the docstring at the top of `combined_garmin_redcap_run.py`. Adapt the field mapping and the `STEP_DAG_MAP` data access group table to your own REDCap projects.

## Running

Interactive (file pickers):

```bash
python combined_garmin_redcap_run.py
```

Fully specified:

```bash
python combined_garmin_redcap_run.py \
  --step-workbook "/path/to/STEP.xlsx" --step-seed-folder "/path/to/STEP/Seed" \
  --rove-workbook "/path/to/ACCS.xlsx" --rove-seed-folder "/path/to/ACCS/Seed" --yes
```

Download and transform only, with no REDCap writes and no workbook write-back:

```bash
python combined_garmin_redcap_run.py --no-upload --no-writeback
```

Run `python combined_garmin_redcap_run.py --help` for all options (pacing, retries, HRV mode, full sweep versus incremental).

## Repository contents

| File | Purpose |
|---|---|
| `combined_garmin_redcap_run.py` | Combined timed run for both projects (retrieval, alignment, REDCap import, readback verification, timing outputs) |
| `step_v3_verify_and_seed.py` | Project 1 token verification and seeding |
| `seed_syncrove_tokens.py` | Project 2 token seeding |
| `requirements.txt` | Python package requirements |
| `examples/` | Tracking workbook templates with made-up rows |
| `data/` | De-identified processing-time data from the evaluated run, with a data dictionary and a script that reproduces the reported timing results |

## Outputs

Each run creates a timestamped results folder containing timing files (per attempt, per date request, and summaries by project and device), REDCap payload snapshots, daily data exports, skipped-record logs, and an assignment-window audit. Account identifiers in audit files are hashed. **Run outputs contain study record codes and calendar dates and must stay on secure institutional storage.**

## Processing-time data

The `data/` folder contains de-identified processing-time data from the evaluated run (September 28, 2026):

| File | Contents |
|---|---|
| `attempts_deidentified.csv` | One row per tracking-workbook row (203): outcome, stage timings, days requested, and Garmin call counts |
| `download_iterations_deidentified.csv` | One row per requested date for the 143 retrieved records (18,319 rows): per-date timing and Garmin call counts |
| `run_summary.csv` | Run-level totals (wall time, Garmin calls and errors, REDCap import and readback time) |
| `group_timing_summary.csv` | Timing summary statistics among verified records, by project and device |
| `data_dictionary.csv` | Field definitions |
| `reproduce_timing_results.py` | Recomputes the reported processing-time results (`cd data && python reproduce_timing_results.py`) |

Study record codes were replaced with random keys, and calendar dates, surgery-relative days, per-record wearable data availability, error messages, account identifiers, and file paths were removed. See `data/README.md` for details.

## Citation

If you use this code, please cite the article above and the Zenodo archive of this repository: https://doi.org/10.5281/zenodo.23332019 (all versions). The version evaluated in the article is v2.0.1: https://doi.org/10.5281/zenodo.23403624.

## License

MIT License. See `LICENSE`.
