"""Recompute the manuscript's processing-time results from the de-identified files.
Usage: python reproduce_timing_results.py   (run from this folder; needs pandas and scipy)"""
import math
import pandas as pd
from scipy import stats

a = pd.read_csv('attempts_deidentified.csv')
it = pd.read_csv('download_iterations_deidentified.csv')
rs = pd.read_csv('run_summary.csv')
total = rs.loc[rs.project == 'Overall', 'total_wall_seconds'].item()
v = a[a.redcap_verified == True]

def ci_t(x):
    h = stats.t.ppf(0.975, len(x) - 1) * x.std() / math.sqrt(len(x))
    return x.mean() - h, x.mean() + h

print('Table 2: outcomes by project'); print(a.groupby(['project', 'outcome']).size().to_string(), '\n')

x = v.total_account_seconds; lo, hi = ci_t(x)
print(f'Verified records: {len(v)}')
print(f'Full execution: {total:.2f} s ({total/60:.1f} min); {total/len(v):.2f} s per verified record; '
      f'{3600/(total/len(v)):.0f} records/h')
print(f'Processing time: mean {x.mean():.1f} (SD {x.std():.1f}; 95% CI {lo:.1f}-{hi:.1f}); '
      f'median {x.median():.1f} (IQR {x.quantile(.25):.1f}-{x.quantile(.75):.1f})\n')

s = rs[rs.project == 'Overall'].iloc[0]
print(f'Share of run time: Garmin API {100*s.garmin_api_active_seconds/total:.1f}%, '
      f'REDCap import {100*s.redcap_upload_seconds_total/total:.1f}%, '
      f'readback {100*s.redcap_verify_seconds_total/total:.1f}%')
print(f'Garmin calls {int(s.garmin_api_attempts)} for {len(it)} dates; errors {int(s.garmin_api_errors)}\n')

print('Table 3: mean (SD) by group')
cols = ['days_requested', 'login_seconds', 'download_iteration_seconds', 'data_sort_seconds',
        'redcap_existing_read_seconds', 'redcap_upload_seconds', 'redcap_verify_seconds', 'total_account_seconds']
for (p, w), g in v.groupby(['project', 'watch_category']):
    print(f'  {p} {w} (n={len(g)})')
    for c in cols:
        print(f'    {c:30s} {g[c].mean():9.3f} ({g[c].std():.3f})')
    if p == 'Project 2':
        prep = g.data_sort_seconds - g.redcap_existing_read_seconds
        print(f'    {"prep excl. existing lookup":30s} {prep.mean():9.3f} ({prep.std():.3f})')
    t = g.total_account_seconds
    print(f'    total median (IQR)             {t.median():.1f} ({t.quantile(.25):.1f}-{t.quantile(.75):.1f})')
    print(f'    retrieval s per requested date {g.download_iteration_seconds.sum()/g.days_requested.sum():.2f}')

print('\nFigure 2B: retrieval time per requested day (Project 1)')
for w, g in v[v.project == 'Project 1'].groupby('watch_category'):
    r = stats.linregress(g.days_requested, g.download_iteration_seconds)
    t = stats.t.ppf(0.975, len(g) - 2)
    print(f'  {w}: {r.slope:.2f} s/day (95% CI {r.slope - t*r.stderr:.2f}-{r.slope + t*r.stderr:.2f})')

print(f'\neTable 3: automated elapsed time {total/3600:.3f} h; break-even {60*total/3600/len(v):.2f} min per device-week')
