
"""
SHARED UTILITIES
================
Fonctions utilitaires partagées par tous les modules de la console.
Centralisé pour éviter la duplication. Aucun customtkinter ici, juste
des helpers pandas/IO.

BNP Paribas Cash Management - Direction Monétique
Mai 2026
"""

import pandas as pd
import os
import csv

# ─── COULEURS BNP (centralisées) ───────────────────────────────────────
VERT_BNP = "#00915A"
VERT_FONCE = "#003D2E"
VERT_CLAIR = "#E8F5E9"
GRIS = "#666666"
GRIS_CLAIR = "#E0E0E0"
GRIS_TRES_CLAIR = "#F5F5F5"
BLANC = "#FFFFFF"
NOIR = "#333333"
ORANGE = "#FF9800"
ROUGE = "#E53935"

def load_csv_smart(path, nrows=None):
    """Charge un CSV en testant plusieurs séparateurs et encodages."""
    for sep in [';', ',', '\t']:
        for enc in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
            try:
                df_test = pd.read_csv(
                    path, sep=sep, encoding=enc, dtype=str,
                    keep_default_na=False, na_values=[],
                    on_bad_lines='skip', nrows=5
                )
                if df_test.shape[1] > 1:
                    return pd.read_csv(
                        path, sep=sep, encoding=enc, dtype=str,
                        keep_default_na=False, na_values=[],
                        on_bad_lines='skip', nrows=nrows
                    )
            except Exception:
                continue
    return pd.read_csv(
        path, sep=None, engine='python', dtype=str,
        keep_default_na=False, na_values=[],
        on_bad_lines='skip', nrows=nrows
    )

def load_xlsx(path, sheet_name=0):
    """Charge un fichier Excel en str (préserve les zéros initiaux)."""
    return pd.read_excel(
        path, sheet_name=sheet_name, dtype=str,
        keep_default_na=False, na_values=[]
    )

def load_file_auto(path, nrows=None):
    """Charge un fichier CSV ou XLSX selon son extension."""
    ext = os.path.splitext(path)[1].lower()
    if ext in ['.xlsx', '.xls']:
        return load_xlsx(path)
    return load_csv_smart(path, nrows)

def protect_long_ids(series, threshold=15):
    """Protège les identifiants longs en format ="VALEUR" pour Excel."""
    def protect(val):
        s = str(val).strip()
        if not s or s in ['', 'nan', 'NaN', 'None', 'N/A']:
            return s
        clean = s.replace(' ', '').replace('.0', '')
        if clean.isdigit() and len(clean) >= threshold:
            return f'="{clean}"'
        return s
    return series.apply(protect)

def save_csv_protected(df, path, protect_ids=True, id_threshold=15):
    """Sauvegarde un CSV avec protection des IDs longs et BOM utf-8-sig."""
    df_out = df.copy()
    if protect_ids:
        for col in df_out.columns:
            sample = df_out[col].dropna().head(100).astype(str)
            has_long = sample.str.replace(' ', '').str.replace('.0', '').str.match(
                r'^\d{' + str(id_threshold) + r',}$'
            ).any()
            if has_long:
                df_out[col] = protect_long_ids(df_out[col], id_threshold)
    df_out.to_csv(
        path, index=False, sep=';',
        encoding='utf-8-sig', quoting=csv.QUOTE_ALL
    )

def get_column_preview(df, col, n=3):
    """Renvoie une string d'aperçu des n premières valeurs d'une colonne."""
    values = df[col].dropna().head(n).astype(str).tolist()
    return " | ".join([v[:40] for v in values])

