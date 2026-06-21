import csv, pandas as pd, os, re, sys

raw_all = pd.read_csv('58807-54909-BRAINS-wiki/data/FREEZE/raw_exports/redcap/all/58807SchatzbergRapid_DATA_LABELS_2026-06-11_1539.csv')

print((set(list(raw_all.columns[0]))))
raw_all['Event Name'].unique()