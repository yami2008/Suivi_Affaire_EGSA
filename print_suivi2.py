# -*- coding: utf-8 -*-
import sys, io
import pandas as pd
from datetime import datetime

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

path = r'C:\Users\hp\Desktop\Suivi_Affaire_EGSA\Suivi_Affaires_EGSA.xlsx'
df = pd.read_excel(path, sheet_name='Suivi Affaires')

today = datetime.strptime('2026-04-15', '%Y-%m-%d')

# Pick required columns
wanted = ['ID', 'Titre', 'Statut', 'Priorité', 'Prochaine Action',
          'Date Prochaine Action', 'Date Limite', 'Date Alerte',
          'Bloquant', 'Responsable', 'Imprévu']

# Match columns case-insensitively
def find_col(name):
    for c in df.columns:
        if c.strip().lower() == name.strip().lower():
            return c
    for c in df.columns:
        if name.lower() in c.lower():
            return c
    return None

mapped = {w: find_col(w) for w in wanted}
print("Column mapping:", mapped)
print()

SEP = "=" * 110

print(SEP)
print(f"  SUIVI AFFAIRES EGSA  |  Date du jour : 15/04/2026  |  {len(df)} affaires")
print(SEP)

for idx, row in df.iterrows():
    def g(col_label):
        col = mapped.get(col_label)
        if col is None:
            return "—"
        val = row[col]
        if pd.isna(val):
            return "—"
        if hasattr(val, 'strftime'):
            return val.strftime('%d/%m/%Y')
        return str(val).strip()

    print()
    print(f"  ┌─ [{g('ID')}]  {g('Titre')}")
    print(f"  │  Statut              : {g('Statut')}")
    print(f"  │  Priorité            : {g('Priorité')}")
    print(f"  │  Responsable         : {g('Responsable')}")
    print(f"  │  Imprévu             : {g('Imprévu')}")
    print(f"  │  Bloquant            : {g('Bloquant')}")
    print(f"  │  Prochaine Action    : {g('Prochaine Action')}")
    print(f"  │  Date Prochaine Act. : {g('Date Prochaine Action')}")
    print(f"  │  Date Limite         : {g('Date Limite')}")
    print(f"  │  Date Alerte         : {g('Date Alerte')}")

    # Urgency
    dl_raw = mapped.get('Date Limite')
    if dl_raw:
        dl_val = row[dl_raw]
        if pd.notna(dl_val):
            try:
                dl_dt = pd.to_datetime(dl_val)
                diff = (dl_dt.to_pydatetime().replace(tzinfo=None) - today).days
                if diff < 0:
                    flag = f"⛔  DÉPASSÉE de {abs(diff)} jour(s) !"
                elif diff == 0:
                    flag = "🔴  ÉCHÉANCE AUJOURD'HUI !"
                elif diff <= 3:
                    flag = f"🟠  Dans {diff} jour(s) — URGENT"
                elif diff <= 7:
                    flag = f"🟡  Dans {diff} jours"
                else:
                    flag = f"🟢  Dans {diff} jours"
                print(f"  │  Délai restant       : {flag}")
            except Exception as e:
                pass

    print(f"  └{'─'*100}")

print()
print("Fin du rapport.")
