"""
TITRE [FRPCH]
=============
REMPLACEMENT DE CARACTERES [FRPCH]

DESCRIPTION
-----------
Outil CLI autonome de recherche et remplacement d'une chaine de caracteres
dans un fichier CSV ou XLSX. Le remplacement s'effectue soit sur UNE colonne
precise (par son nom), soit sur TOUTES les colonnes (valeur speciale ALL).
Le remplacement est litteral (regex=False), applique sur les valeurs
converties en chaine (astype(str)). La logique metier est preservee a
l'identique depuis la version GUI customtkinter d'origine.

SOURCES REQUISES
----------------
- --input : fichier source CSV ou XLSX a traiter.

OUTPUTS PRODUITS
----------------
- --output-file : fichier CSV resultat, sauvegarde en utf-8-sig avec
  protection des identifiants longs (format ="VALEUR" pour Excel),
  separateur ';' et quoting QUOTE_ALL.

ARGUMENTS CLI
-------------
- --input        (requis) : chemin du fichier source CSV/XLSX.
- --output-file  (requis) : chemin du fichier CSV resultat.
- --column       (optionnel, defaut: premiere colonne du fichier) :
                  nom de la colonne cible, ou la valeur speciale 'ALL'
                  pour appliquer le remplacement sur TOUTES les colonnes.
- --search       (requis) : chaine de caracteres a rechercher.
- --replace      (optionnel, defaut: '') : chaine de remplacement.

DECOMPOSITION
-------------
1. Chargement du fichier source via load_file_auto (CSV/XLSX, tout en str).
2. Determination de la colonne cible (nom precis ou ALL).
3. Application du remplacement litteral (str.replace, regex=False).
4. Sauvegarde CSV protegee via save_csv_protected (utf-8-sig, IDs longs).

BNP Paribas Cash Management - Direction Monetique
Mai 2026
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from pathlib import Path

import pandas as pd

VERSION_ID = "FRPCH"

# Valeur speciale reproduisant le bouton "Appliquer a TOUTES" de l'UI.
ALL_COLUMNS = "ALL"


# ─── Helpers inlines depuis _shared.py (logique inchangee) ──────────────

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


# ─── Logique de remplacement (inchangee depuis apply_replace / apply_all) ──

def apply_replace(df: pd.DataFrame, col: str, s: str, r: str) -> pd.DataFrame:
    """Remplacement sur UNE colonne (reproduit ReplaceCharsApp.apply_replace)."""
    if col in df.columns and s:
        df[col] = df[col].astype(str).str.replace(s, r, regex=False)
        print(f"'{s}' -> '{r}' applique sur {col}")
    return df


def apply_all(df: pd.DataFrame, s: str, r: str) -> pd.DataFrame:
    """Remplacement sur TOUTES les colonnes (reproduit ReplaceCharsApp.apply_all)."""
    if s:
        for col in df.columns:
            df[col] = df[col].astype(str).str.replace(s, r, regex=False)
        print(f"'{s}' -> '{r}' applique sur TOUTES les colonnes")
    return df


def main() -> int:
    parser = argparse.ArgumentParser(
        prog="05.FRPCH",
        description=f"REMPLACEMENT DE CARACTERES [{VERSION_ID}] - "
                    "recherche/remplacement dans un CSV/XLSX (CLI).",
    )
    parser.add_argument("--input", required=True,
                        help="Fichier source CSV ou XLSX a traiter.")
    parser.add_argument("--output-file", required=True,
                        help="Fichier CSV resultat (utf-8-sig).")
    parser.add_argument("--column", default=None,
                        help=f"Nom de colonne cible, ou '{ALL_COLUMNS}' pour "
                             "TOUTES les colonnes. Defaut: premiere colonne.")
    parser.add_argument("--search", required=True,
                        help="Chaine a rechercher.")
    parser.add_argument("--replace", default="",
                        help="Chaine de remplacement (defaut: vide).")

    args = parser.parse_args()

    try:
        input_path = Path(args.input)
        output_path = Path(args.output_file)

        if not input_path.is_file():
            print(f"ERREUR : fichier introuvable : {input_path}")
            return 2

        print(f"Chargement : {input_path}")
        df = load_file_auto(str(input_path))
        print(f"{len(df):,} lignes, {len(df.columns)} colonnes")

        if len(df.columns) == 0:
            print("ERREUR : aucune colonne dans le fichier source.")
            return 1

        # Colonne cible : defaut = premiere colonne (comme l'UI au chargement).
        column = args.column if args.column is not None else df.columns[0]
        s = args.search
        r = args.replace

        if column == ALL_COLUMNS:
            df = apply_all(df, s, r)
        else:
            if column not in df.columns:
                print(f"ERREUR : colonne introuvable : {column}")
                print(f"Colonnes disponibles : {list(df.columns)}")
                return 1
            df = apply_replace(df, column, s, r)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        save_csv_protected(df, str(output_path))
        print(f"Sauvegarde : {output_path}")
        return 0

    except Exception as exc:
        print(f"ERREUR : {exc}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
