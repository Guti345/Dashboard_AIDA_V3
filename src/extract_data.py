"""
extract_data.py — AIDA Ventures Integrated Dashboard
Reads all Excel sources and returns a unified data dictionary.
"""
import os
import json
import pandas as pd
from utils import safe_str, clean_val

RAW_DIR = os.path.join(os.path.dirname(__file__), '..', 'data', 'raw')
PROCESSED_DIR = os.path.join(os.path.dirname(__file__), '..', 'data', 'processed')


def _path(filename):
    return os.path.join(RAW_DIR, filename)


# ─── FINTECH SECTORS ────────────────────────────────────────────────────────

def read_fintech():
    path = _path('Fintech_Sectors.xlsx')

    # Investment by Country
    df_country = pd.read_excel(path, sheet_name='Investment by Country 2024', header=0)
    countries = []
    for _, r in df_country.iterrows():
        countries.append({
            'country': safe_str(r.iloc[0]),
            'investment': safe_str(r.iloc[1]),
            'share': safe_str(r.iloc[2]),
            'fintechs': safe_str(r.iloc[3]),
            'deals': safe_str(r.iloc[4]),
            'unicorns': safe_str(r.iloc[5]),
            'growth': safe_str(r.iloc[6]),
        })

    # LATAM Subsectors
    df_latam = pd.read_excel(path, sheet_name='LATAM Subsectors 2024', header=0)
    latam_sub = []
    for _, r in df_latam.iterrows():
        latam_sub.append({
            'subsector': safe_str(r.iloc[0]),
            'investment': safe_str(r.iloc[1]),
            'pct': safe_str(r.iloc[2]),
            'startups': safe_str(r.iloc[3]),
            'leaders': safe_str(r.iloc[4]),
        })

    # USA Subsectors
    df_usa = pd.read_excel(path, sheet_name='USA Subsectors 2024', header=0)
    usa_sub = []
    for _, r in df_usa.iterrows():
        usa_sub.append({
            'subsector': safe_str(r.iloc[0]),
            'investment': safe_str(r.iloc[1]),
            'pct': safe_str(r.iloc[2]),
            'startups': safe_str(r.iloc[3]),
            'trend': safe_str(r.iloc[4]),
        })

    # USA vs LATAM Comparison
    df_cmp = pd.read_excel(path, sheet_name='USA vs LATAM Comparison', header=0)
    comparison = []
    for _, r in df_cmp.iterrows():
        comparison.append({
            'metric': safe_str(r.iloc[0]),
            'usa': safe_str(r.iloc[1]),
            'latam': safe_str(r.iloc[2]),
            'ratio': safe_str(r.iloc[3]),
        })

    return {
        'countries': countries,
        'latam_subsectors': latam_sub,
        'usa_subsectors': usa_sub,
        'comparison': comparison,
    }


# ─── VC FUNDS METRICS ───────────────────────────────────────────────────────

def read_vcfunds():
    path = _path('VCFunds_Metrics.xlsx')

    def read_sheet(name):
        df = pd.read_excel(path, sheet_name=name, header=None)
        # Find header row (has 'Métrica')
        header_row = None
        for i, row in df.iterrows():
            if any('trica' in str(v) for v in row.values):
                header_row = i
                break
        if header_row is None:
            return []
        df.columns = [safe_str(df.iloc[header_row][c]) for c in range(len(df.columns))]
        df = df.iloc[header_row + 1:].reset_index(drop=True)
        records = []
        for _, r in df.iterrows():
            vals = [safe_str(v) for v in r.values]
            if all(v == '' for v in vals):
                continue
            records.append(vals)
        col_names = [safe_str(df.columns[i]) for i in range(len(df.columns))]
        return {'headers': col_names, 'rows': records}

    early = read_sheet('EarlyStageFunds')
    benchmark = read_sheet('VCFunds Benchmark')
    latam_vs_us = read_sheet('LATAM vs US')
    return {
        'early_stage': early,
        'benchmark': benchmark,
        'latam_vs_us': latam_vs_us,
    }


# ─── VENTURE STUDIO ─────────────────────────────────────────────────────────

def read_venture_studio():
    path = _path('Venture_Studio_Metrics_Reference.xlsx')
    df = pd.read_excel(path, sheet_name='Venture Studio Metrics Ref', header=None)

    sections = {'A': [], 'B': [], 'C': []}
    current = None
    for _, row in df.iterrows():
        vals = [safe_str(v) for v in row.values]
        first = vals[0]
        if 'A. STUDIO' in first:
            current = 'A'
            continue
        if 'B. PORTFOLIO' in first:
            current = 'B'
            continue
        if 'C. FUND' in first:
            current = 'C'
            continue
        if current and vals[0] not in ('', 'Metric') and any(v != '' for v in vals):
            sections[current].append({
                'metric': vals[0],
                'early': vals[1] if len(vals) > 1 else '',
                'mature': vals[2] if len(vals) > 2 else '',
                'industry': vals[3] if len(vals) > 3 else '',
                'notes': vals[4] if len(vals) > 4 else '',
            })
    return sections


# ─── AIDA BENCHMARKS ────────────────────────────────────────────────────────

def read_aida_benchmarks():
    path = _path('_AIDA_Ventures_-_Startups_Benchmarks.xlsx')

    # Revenue Multiples - SaaS section (rows 3-11)
    df_rev = pd.read_excel(path, sheet_name='Revenue Multiples', header=None)
    rev_saas, rev_fintech, rev_logtech = [], [], []
    section = None
    for _, row in df_rev.iterrows():
        vals = [safe_str(v) for v in row.values]
        first = vals[0]
        if 'SaaS' in first:
            section = 'saas'
            continue
        if 'Fintech' in first:
            section = 'fintech'
            continue
        if 'LogTech' in first:
            section = 'logtech'
            continue
        if section and vals[0] not in ('', 'Stage') and any(v != '' for v in vals):
            rec = {
                'stage': vals[0],
                'type': vals[1],
                'global': vals[2],
                'latam': vals[3],
                'use_case': vals[4],
                'source': vals[5],
            }
            if section == 'saas':
                rev_saas.append(rec)
            elif section == 'fintech':
                rev_fintech.append(rec)
            elif section == 'logtech':
                rev_logtech.append(rec)

    # US Valuations
    df_us = pd.read_excel(path, sheet_name='US Valuations', header=None)
    us_val_saas, us_val_fintech, us_val_logtech = [], [], []
    section = None
    for _, row in df_us.iterrows():
        vals = [safe_str(v) for v in row.values]
        first = vals[0]
        if 'Enterprise SaaS' in first:
            section = 'saas'
            continue
        if 'Fintech' in first:
            section = 'fintech'
            continue
        if 'LogTech' in first:
            section = 'logtech'
            continue
        if section and vals[0] not in ('', 'Stage') and any(v != '' for v in vals):
            rec = {'stage': vals[0], 'val': vals[1], 'multiple': vals[2], 'notes': vals[3], 'round': vals[4]}
            if section == 'saas':
                us_val_saas.append(rec)
            elif section == 'fintech':
                us_val_fintech.append(rec)
            elif section == 'logtech':
                us_val_logtech.append(rec)

    # LATAM Valuations - discount table
    df_la = pd.read_excel(path, sheet_name='LatAm Valuations', header=None)
    latam_discount, latam_by_sector, latam_notable = [], [], []
    section = None
    for _, row in df_la.iterrows():
        vals = [safe_str(v) for v in row.values]
        first = vals[0]
        if 'LATAM Valuation Discount' in first:
            section = 'discount'
            continue
        if 'Estimated LATAM Valuations by Stage' in first:
            section = 'by_sector'
            continue
        if 'Notable LATAM Funding' in first:
            section = 'notable'
            continue
        if 'Why the LATAM Discount' in first:
            section = None
            continue
        if section == 'discount' and vals[0] not in ('', 'Factor') and any(v != '' for v in vals):
            latam_discount.append({'factor': vals[0], 'discount': vals[1], 'source': vals[2], 'notes': vals[3]})
        elif section == 'by_sector' and vals[0] not in ('', 'Stage') and any(v != '' for v in vals):
            latam_by_sector.append({'stage': vals[0], 'saas': vals[1], 'fintech': vals[2], 'logtech': vals[3],
                                     'us_ref': vals[4] if len(vals) > 4 else '', 'discount': vals[5] if len(vals) > 5 else ''})
        elif section == 'notable' and vals[0] not in ('', 'Company') and any(v != '' for v in vals):
            latam_notable.append({'company': vals[0], 'country': vals[1], 'stage': vals[2],
                                   'amount': vals[3], 'valuation': vals[4], 'date': vals[5], 'sector': vals[6] if len(vals) > 6 else ''})

    # Graduation Rates
    df_grad = pd.read_excel(path, sheet_name='Graduation Rates', header=None)
    grad_rates, grad_sector, grad_cohort = [], [], []
    section = None
    for _, row in df_grad.iterrows():
        vals = [safe_str(v) for v in row.values]
        first = vals[0]
        if 'Stage Graduation Rates by Geography' in first:
            section = 'geography'
            continue
        if 'Sector-Specific Graduation Rates' in first:
            section = 'sector'
            continue
        if 'Cumulative Survival' in first:
            section = 'cohort'
            continue
        if section == 'geography' and vals[0] not in ('', 'Stage Transition') and any(v != '' for v in vals[:5]):
            grad_rates.append({'transition': vals[0], 'us_2024': vals[1], 'us_2021': vals[2], 'latam': vals[3], 'latam_top': vals[4]})
        elif section == 'sector' and vals[0] not in ('', 'Sector') and any(v != '' for v in vals[:4]):
            grad_sector.append({'sector': vals[0], 'us_rate': vals[1], 'latam_rate': vals[2], 'key': vals[3]})
        elif section == 'cohort' and vals[0] not in ('', 'Stage') and any(v != '' for v in vals[:3]):
            grad_cohort.append({'stage': vals[0], 'us': vals[1], 'latam': vals[2], 'pct_us': vals[3] if len(vals) > 3 else ''})

    # Time Between Rounds - general section
    df_time = pd.read_excel(path, sheet_name='Time Between Rounds', header=None)
    time_all = []
    section = None
    for _, row in df_time.iterrows():
        vals = [safe_str(v) for v in row.values]
        first = vals[0]
        if 'Section 1: All Sectors' in first:
            section = 'all'
            continue
        if 'Fintech' in first and 'Section' not in first and section == 'all':
            section = None
        if section == 'all' and vals[0] not in ('', 'Transition') and '→' in vals[0]:
            time_all.append({
                'transition': vals[0],
                'us_2021': vals[1],
                'us_2024': vals[2],
                'change': vals[3],
                'latam_median': vals[4],
                'latam_top': vals[5],
            })

    return {
        'rev_multiples': {'saas': rev_saas, 'fintech': rev_fintech, 'logtech': rev_logtech},
        'us_valuations': {'saas': us_val_saas, 'fintech': us_val_fintech, 'logtech': us_val_logtech},
        'latam_discount': latam_discount,
        'latam_by_sector': latam_by_sector,
        'latam_notable': latam_notable,
        'graduation': {'geography': grad_rates, 'sector': grad_sector, 'cohort': grad_cohort},
        'time_between_rounds': time_all,
    }


# ─── STARTUP METRICS ────────────────────────────────────────────────────────

def read_metricas():
    path = _path('_Metricas_Startups.xlsx')

    def read_benchmark_sheet(sheet_name):
        df = pd.read_excel(path, sheet_name=sheet_name, header=None)
        # Find header row (row with 'Métrica')
        header_row = None
        for i, row in df.iterrows():
            if any('trica' in str(v) for v in row.values):
                header_row = i
                break
        if header_row is None:
            return []
        rows = []
        for i in range(header_row + 1, len(df)):
            r = df.iloc[i]
            vals = [safe_str(v) for v in r.values]
            non_empty = [v for v in vals if v]
            if len(non_empty) < 2:
                continue
            rows.append({
                'metric': vals[1] if len(vals) > 1 else '',
                'preseed': vals[2] if len(vals) > 2 else '',
                'seed': vals[3] if len(vals) > 3 else '',
                'series_a': vals[4] if len(vals) > 4 else '',
            })
        return rows

    return {
        'income_growth': read_benchmark_sheet('Income&Growth US'),
        'capital_efficiency': read_benchmark_sheet('Capital Efficiency US'),
        'retention': read_benchmark_sheet('Retention&Customers'),
        'unit_economics': read_benchmark_sheet('Unit Economics'),
        'market_valuation': read_benchmark_sheet('Market&Valuation'),
        'operating': read_benchmark_sheet('Operating Metrics'),
    }


# ─── MAIN EXTRACTOR ─────────────────────────────────────────────────────────

def extract_all():
    data = {
        'fintech': read_fintech(),
        'vcfunds': read_vcfunds(),
        'venture_studio': read_venture_studio(),
        'benchmarks': read_aida_benchmarks(),
        'metricas': read_metricas(),
    }
    return data


def save_json(data, filename='dashboard_data.json'):
    os.makedirs(PROCESSED_DIR, exist_ok=True)
    out_path = os.path.join(PROCESSED_DIR, filename)
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"  → Saved: {out_path}")
    return out_path


if __name__ == '__main__':
    print("Extracting data from Excel sources...")
    data = extract_all()
    save_json(data)
    print("Done.")
