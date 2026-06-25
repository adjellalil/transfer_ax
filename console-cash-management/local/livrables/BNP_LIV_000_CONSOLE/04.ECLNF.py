"""
NETTOYAGE DE FICHIERS [ECLNF]
==============================

DESCRIPTION
-----------
Outil CLI autonome de nettoyage et de restructuration d'un fichier CSV / XLSX.
Reproduit a l'identique la logique metier de l'application GUI d'origine :
  - Supprimer des lignes par indices
  - Renommer une colonne existante
  - Ajouter une nouvelle colonne avec valeur par defaut
  - Supprimer une colonne
  - Sauvegarde CSV utf-8-sig avec protection des IDs longs

Les operations sont appliquees dans l'ordre suivant (identique a l'UI) :
drop-rows -> rename-col -> add-col -> drop-col, puis sauvegarde.
Chaque option peut etre repetee plusieurs fois.

SOURCES REQUISES
----------------
- --input : un fichier source CSV (.csv) ou Excel (.xlsx / .xls).

OUTPUTS PRODUITS
----------------
- --output-file : un fichier CSV (separateur ';', encodage utf-8-sig,
  toutes les valeurs entre guillemets, IDs longs proteges au format ="VALEUR").

ARGUMENTS CLI
-------------
  --input PATH                 Fichier source CSV/XLSX (requis).
  --output-file PATH           Fichier CSV de sortie (requis).
  --drop-rows I,J,K            Indices (post reset_index) des lignes a
                               supprimer, separes par des virgules.
                               Repetable ; cumulatif.
  --rename-col ANCIEN:NOUVEAU  Renomme la colonne ANCIEN en NOUVEAU.
                               Repetable.
  --add-col NOM:VALEUR         Ajoute une colonne NOM remplie avec VALEUR.
                               Repetable. (NOM:  => valeur vide).
  --drop-col NOM               Supprime la colonne NOM. Repetable.

DECOMPOSITION
-------------
1. Parsing des arguments CLI (argparse) et validation de l'entree.
2. Chargement du fichier source (load_file_auto, detection auto CSV/XLSX).
3. Application sequentielle des operations de nettoyage (logique UI inchangee).
4. Sauvegarde protegee (save_csv_protected) vers le fichier de sortie.
5. Sorties : 0 (succes), 2 (arguments invalides), 1 (erreur d'execution).

BNP Paribas Cash Management - Direction Monetique
Mai 2026
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from pathlib import Path
from typing import List, Tuple

import pandas as pd

VERSION_ID = "ECLNF"


# ─── HELPERS INLINES DEPUIS _shared.py (logique VERBATIM) ──────────────

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


# ─── PARSERS D'ARGUMENTS "CLE:VALEUR" ──────────────────────────────────

def _parse_pair(value: str) -> Tuple[str, str]:
    """Decoupe 'GAUCHE:DROITE' en (GAUCHE, DROITE). DROITE peut etre vide."""
    if ':' not in value:
        raise argparse.ArgumentTypeError(
            f"Format invalide '{value}' : attendu GAUCHE:DROITE"
        )
    left, right = value.split(':', 1)
    if not left.strip():
        raise argparse.ArgumentTypeError(
            f"Format invalide '{value}' : partie gauche vide"
        )
    return left, right


# ─── OPERATIONS DE NETTOYAGE (logique UI inchangee) ────────────────────

def op_delete_rows(df: pd.DataFrame, val: str) -> pd.DataFrame:
    """Supprime des lignes par indices (logique identique a delete_rows)."""
    rows = [int(x.strip()) for x in val.split(',')]
    df.drop(rows, inplace=True)
    df.reset_index(drop=True, inplace=True)
    return df


def op_rename_col(df: pd.DataFrame, old: str, new: str) -> pd.DataFrame:
    """Renomme une colonne (logique identique a rename_col)."""
    if old and old in df.columns:
        if new:
            df.rename(columns={old: new.strip()}, inplace=True)
    return df


def op_add_col(df: pd.DataFrame, name: str, val: str) -> pd.DataFrame:
    """Ajoute une colonne (logique identique a add_col)."""
    if name:
        val = val or ""
        df[name.strip()] = val
    return df


def op_delete_col(df: pd.DataFrame, col: str) -> pd.DataFrame:
    """Supprime une colonne (logique identique a delete_col)."""
    if col and col in df.columns:
        df.drop(columns=[col], inplace=True)
    return df


# ─── POINT D'ENTREE CLI ────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="04.ECLNF.py",
        description=f"NETTOYAGE DE FICHIERS [{VERSION_ID}] - "
                    "nettoyage/restructuration CSV/XLSX (CLI).",
    )
    parser.add_argument(
        "--input", required=True, type=Path,
        help="Fichier source CSV ou XLSX.",
    )
    parser.add_argument(
        "--output-file", required=True, type=Path,
        help="Fichier CSV de sortie.",
    )
    parser.add_argument(
        "--drop-rows", action="append", default=[], metavar="I,J,K",
        help="Indices des lignes a supprimer (ex: 0,1,2). Repetable.",
    )
    parser.add_argument(
        "--rename-col", action="append", default=[], type=_parse_pair,
        metavar="ANCIEN:NOUVEAU",
        help="Renomme une colonne. Repetable.",
    )
    parser.add_argument(
        "--add-col", action="append", default=[], type=_parse_pair,
        metavar="NOM:VALEUR",
        help="Ajoute une colonne avec valeur par defaut. Repetable.",
    )
    parser.add_argument(
        "--drop-col", action="append", default=[], metavar="NOM",
        help="Supprime une colonne. Repetable.",
    )
    return parser


def main(argv: List[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    input_path: Path = args.input
    output_path: Path = args.output_file

    if not input_path.is_file():
        print(f"[ERREUR] Fichier introuvable : {input_path}")
        return 2

    try:
        print(f"[INFO] Chargement : {input_path}")
        df = load_file_auto(str(input_path))
        print(f"[INFO] Charge : {len(df):,} lignes x {len(df.columns)} colonnes")

        # --- drop-rows ---
        for val in args.drop_rows:
            print(f"[INFO] Suppression lignes : {val}")
            df = op_delete_rows(df, val)

        # --- rename-col ---
        for old, new in args.rename_col:
            if old not in df.columns:
                print(f"[ALERTE] Colonne a renommer absente : {old}")
            print(f"[INFO] Renommage colonne : {old} -> {new.strip()}")
            df = op_rename_col(df, old, new)

        # --- add-col ---
        for name, val in args.add_col:
            print(f"[INFO] Ajout colonne : {name.strip()}")
            df = op_add_col(df, name, val)

        # --- drop-col ---
        for col in args.drop_col:
            if col not in df.columns:
                print(f"[ALERTE] Colonne a supprimer absente : {col}")
            print(f"[INFO] Suppression colonne : {col}")
            df = op_delete_col(df, col)

        print(f"[INFO] Resultat : {len(df):,} lignes x {len(df.columns)} colonnes")

        output_path.parent.mkdir(parents=True, exist_ok=True)
        save_csv_protected(df, str(output_path))
        print(f"[SUCCES] Sauvegarde : {output_path}")
        return 0

    except Exception as e:
        print(f"[ERREUR] {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
