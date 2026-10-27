# Combined Garmin-to-REDCap run: processing-time data (September 28, 2026)

De-identified record-level processing-time data for the evaluated execution described in
"Leveraging Artificial Intelligence to Reduce Manual Burden in Perioperative Activity Data Collection" (Daly J, Skalicky M, Arai E, et al.; manuscript under review).

## Files

| File | Contents |
|---|---|
| `attempts_deidentified.csv` | One row per tracking-workbook row (203): outcome, stage timings, days requested, Garmin call counts |
| `download_iterations_deidentified.csv` | One row per requested date for the 143 retrieved records (18,319 rows): per-date timing and Garmin call counts |
| `run_summary.csv` | Run-level totals (wall time, Garmin calls and errors, REDCap import and readback time) |
| `group_timing_summary.csv` | Summary statistics of timing metrics among verified records, by project and device |
| `data_dictionary.csv` | Field definitions |
| `reproduce_timing_results.py` | Recomputes Table 2 outcome counts, Table 3, the processing-time results, the Figure 2B slopes, and the eTable 3 inputs |

## What was removed

- Study record IDs (replaced by random keys; no linkage file was kept)
- Calendar dates and surgery-relative days
- Wearable data availability per record (days with steps, sleep, or any data; preoperative and postoperative day counts; nonblank fields uploaded)
- Error messages, account hashes, local file paths, token locations, and process IDs

No rate-limit events were recorded during the run, so that log is not included.

## Reproducing the results

```
python reproduce_timing_results.py
```

Requires Python 3 with pandas and SciPy. Data availability results (Figure 2C and the Availability paragraph) use participant-level wearable data and cannot be reproduced from these files.

## Citation

Please cite the article above and the Zenodo archive of this repository: https://doi.org/10.5281/zenodo.23332019 (all versions). The version evaluated in the article is v2.0.1: https://doi.org/10.5281/zenodo.23403624.
