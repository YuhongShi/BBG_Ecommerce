#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Check Sheet1: for each row, Betrag must equal sum of cols 5+6+7+8+9+10."""
import sys, argparse
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

parser = argparse.ArgumentParser()
parser.add_argument("--umsatz", required=True)
parser.add_argument("--platform", default=None, help="Filter by Zahlungsverkehr (e.g. Kaufland)")
parser.add_argument("--threshold", type=float, default=0.02)
args = parser.parse_args()

wb = openpyxl.load_workbook(args.umsatz, data_only=True)
ws = wb["Sheet1"]

# Col indices (1-based): Nr=1, Datum=2, Zahlungsverkehr=3, Betrag=4,
# Brrutto_Umsatz=5, Andere=6, KTO_Guthaben=7, Behalten=8, Einzahlung=9,
# Geb_Ohne_Steuer=10, Steuersatz=11(skip), Marktplatz=12(skip)

COL_NAMES = {
    4:  "Betrag",
    5:  "Brrutto_Umsatz",
    6:  "Andere Gebühren",
    7:  "KTO_Guthaben",
    8:  "Behalten_in_KTO",
    9:  "Einzahlung",
    10: "Geb_Ohne_Steuer",
}

def v(cell):
    val = cell.value
    if val is None: return 0.0
    try: return float(val)
    except: return 0.0

print(f"File: {args.umsatz}")
print(f"Platform filter: {args.platform or 'ALL'}")
print(f"Threshold: ±{args.threshold}")
print()

problems = []
all_rows = []

for row in ws.iter_rows(min_row=2):
    nr      = row[0].value
    zahlv   = str(row[2].value or "")
    betrag  = v(row[3])

    if not nr:
        continue
    if args.platform and args.platform.lower() not in zahlv.lower():
        continue
    if betrag == 0:
        continue

    brutto  = v(row[4])
    andere  = v(row[5])
    kto_gut = v(row[6])
    behalten= v(row[7])
    einzahl = v(row[8])
    geb     = v(row[9])

    # Behalten_in_KTO is stored as negative (deduction), so add it
    col_sum = round(brutto + andere + kto_gut + behalten + einzahl + geb, 2)
    diff    = round(betrag - col_sum, 2)

    all_rows.append((int(nr), zahlv, betrag, brutto, andere, kto_gut, behalten, einzahl, geb, col_sum, diff))

    if abs(diff) > args.threshold:
        problems.append((int(nr), zahlv, betrag, brutto, andere, kto_gut, behalten, einzahl, geb, col_sum, diff))

# Print header
print(f"{'Nr':>4}  {'Zahlungsverkehr':<22}  {'Betrag':>10}  {'Brutto':>10}  {'Andere':>10}  {'KTO_Gut':>9}  {'Behalten':>9}  {'Einzahl':>8}  {'GebOhne':>8}  {'Sum':>10}  {'Diff':>8}")
print("─" * 130)

for r in all_rows:
    nr, zv, betrag, brutto, andere, kto, behalten, einzahl, geb, csum, diff = r
    flag = " ◄ DIFF" if abs(diff) > args.threshold else ""
    print(f"{nr:>4}  {zv:<22}  {betrag:>10.2f}  {brutto:>10.2f}  {andere:>10.2f}  {kto:>9.2f}  {behalten:>9.2f}  {einzahl:>8.2f}  {geb:>8.2f}  {csum:>10.2f}  {diff:>8.2f}{flag}")

print()
if problems:
    print(f"⚠ {len(problems)} row(s) with |diff| > {args.threshold}:")
    for r in problems:
        nr, zv, betrag, _, _, _, _, _, _, csum, diff = r
        print(f"  Nr={nr:>3}  {zv:<22}  Betrag={betrag:.2f}  Sum={csum:.2f}  Diff={diff:+.2f}")
else:
    print(f"✓ All rows balance (threshold ±{args.threshold})")
