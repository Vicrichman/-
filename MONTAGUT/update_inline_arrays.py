#!/usr/bin/env python3
"""Regenerate inline data arrays in index.html"""
import json

DATA_DIR = '/home/Vic/dewu-reports/MONTAGUT/2026/data'
HTML_PATH = '/home/Vic/dewu-reports/MONTAGUT/2026/index.html'

# ── output-dir support ──
import argparse as _ap, os as _os
_args = _ap.ArgumentParser()
_args.add_argument('--output-dir', default=None)
_args, _ = _args.parse_known_args()
if _args.output_dir:
    _od = _args.output_dir
    assert _os.path.isdir(_od), f"output-dir not found: {_od}"
    # Template still read from production (managed by git)
    HTML_TEMPLATE = HTML_PATH
    HTML_PATH = _os.path.join(_od, 'index.html')
    DATA_DIR = _os.path.join(_od, 'data')


with open(f'{DATA_DIR}/ALL_DATES.json') as f:
    all_dates = json.load(f)
with open(f'{DATA_DIR}/ALL_CATS.json') as f:
    all_cats = json.load(f)
with open(f'{DATA_DIR}/ALL_MONTHS.json') as f:
    all_months = json.load(f)
with open(f'{DATA_DIR}/ALL_GOODS.json') as f:
    all_goods = json.load(f)

with open(HTML_TEMPLATE if '_od' in dir() else HTML_PATH) as f:
    html = f.read()

# Generate new JS array strings
dates_js = 'const ALL_DATES = ' + json.dumps(all_dates, ensure_ascii=False) + ';'
cats_js = 'const ALL_CATS = ' + json.dumps(all_cats, ensure_ascii=False) + ';'
goods_js = 'const ALL_GOODS = ' + json.dumps(all_goods, ensure_ascii=False) + ';'
months_js = 'const ALL_MONTHS = ' + json.dumps(all_months, ensure_ascii=False) + ';'

# Replace in HTML
import re

# ALL_DATES
html = re.sub(r'const ALL_DATES = \[.*?\];', dates_js, html, count=1, flags=re.DOTALL)
# ALL_CATS
html = re.sub(r'const ALL_CATS = \[.*?\];', cats_js, html, count=1, flags=re.DOTALL)
# ALL_GOODS
html = re.sub(r'const ALL_GOODS = \[.*?\];', goods_js, html, count=1, flags=re.DOTALL)
# ALL_MONTHS
html = re.sub(r'const ALL_MONTHS = \[.*?\];', months_js, html, count=1, flags=re.DOTALL)

with open(HTML_PATH, 'w') as f:
    f.write(html)

print(f'Updated inline arrays:')
print(f'  ALL_DATES: {len(all_dates)} items')
print(f'  ALL_CATS: {len(all_cats)} items')
print(f'  ALL_GOODS: {len(all_goods)} items')
print(f'  ALL_MONTHS: {len(all_months)} items')

# Verify replacements
with open(HTML_PATH) as f:
    new_html = f.read()
for name in ['ALL_DATES', 'ALL_CATS', 'ALL_GOODS', 'ALL_MONTHS']:
    idx = new_html.index(f'const {name} = ')
    end = new_html.index(';', idx)
    snippet = new_html[idx:end+1]
    print(f'  {name}: {snippet[:80]}...{snippet[-30:]}')
