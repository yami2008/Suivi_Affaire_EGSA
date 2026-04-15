import pandas as pd
from datetime import datetime

path = r'C:\Users\hp\Desktop\Suivi_Affaire_EGSA\Suivi_Affaires_EGSA.xlsx'
df = pd.read_excel(path, sheet_name='Suivi Affaires')

today = datetime.strptime('2026-04-15', '%Y-%m-%d')

cols = ['ID', 'Titre', 'Statut', 'Priorite', 'Prochaine Action',
        'Date Prochaine Action', 'Date Limite', 'Date Alerte',
        'Bloquant', 'Responsable', 'Imprevu']

# Map accented column names to actual ones
col_map = {}
for c in df.columns:
    col_map[c] = c

df_view = df[['ID', 'Titre', 'Statut',
              'Priorite' if 'Priorite' in df.columns else [x for x in df.columns if 'Priorit' in x][0],
              'Prochaine Action', 'Date Prochaine Action', 'Date Limite',
              'Date Alerte', 'Bloquant', 'Responsable',
              'Imprevu' if 'Imprevu' in df.columns else [x for x in df.columns if 'Impr' in x][0]
             ]].copy()
print(df_view.to_string())
