"""
MONEXT FUSION PECC52 [N7M4P]
==============================

DESCRIPTION :
-------------
Script CLI autonome de fusion des fichiers MONEXT PECC-5212 mensuels en un
unique CSV consolide multi-annee. Deux modes d'execution :
  1. CONCATENATION HISTORIQUE (mode par defaut) : fusionne tous les fichiers
     PECC mensuels presents dans --input-folder, tries par periode.
  2. AJOUT INCREMENTAL : lorsque --existing-file est fourni, ajoute les
     fichiers du dossier a un consolide deja existant.

Reconnaissance intelligente du mois via le nom du fichier (pattern _MMYYYY).
Sacralisation des colonnes critiques au format ="valeur" pour preserver les
zeros initiaux. Export CSV separateur ; encodage UTF-8 BOM (QUOTE_NONE).

SOURCES REQUISES :
------------------
- --input-folder : dossier contenant les fichiers MONEXT PECC-5212 mensuels
  (1 fichier par mois, nom contenant une periode au format MMYYYY).
- --existing-file (mode incremental uniquement) : CSV consolide existant
  auquel ajouter les fichiers du dossier.

OUTPUTS PRODUITS :
------------------
- --output-file : CSV consolide multi-annee, separateur ;, encodage UTF-8 BOM.
  Colonnes sacralisees au format ="valeur".

ARGUMENTS CLI :
---------------
- --input-folder PATH   (obligatoire) : dossier des fichiers PECC mensuels.
- --output-file  PATH   (obligatoire) : chemin du CSV consolide a produire.
- --existing-file PATH  (optionnel)   : active le mode incremental ; ajoute les
                                        fichiers du dossier a ce consolide.
- --rc-col       INT    (optionnel)   : n. de colonne RC a sacraliser (base 1,
                                        defaut 10 -> index 9 base 0).

COLONNES SACRALISEES (fixes) :
------------------------------
- Colonne 5  : Identifiant bancaire
- Colonne 9  : Identifiant RP (Relation Personne)
- Colonne 10 : Identifiant RC (Relation Commerciale) - configurable via --rc-col
- Colonne 11 : IBAN

DECOMPOSITION :
---------------
03.N7M4P.py
|-- CONFIGURATION
|   |-- SACRAL_COL_FIXED, RC_COL_DEFAULT, MOIS_NAMES
|-- HELPERS (logique metier preservee verbatim)
|   |-- extract_period_from_filename()
|   |-- period_to_label()
|   |-- period_sort_key()
|   |-- sacralise_id()
|   |-- load_csv_smart()
|   |-- sacralise_colonnes()
|-- CONFIG SACRALISATION
|   |-- get_rc_col_index()
|   |-- get_all_sacral_indices()
|-- COLLECTE
|   |-- collect_monext_files()
|-- WORKERS
|   |-- worker_historique()
|   |-- worker_incremental()
|   |-- export_result()
|-- main() argparse -> try/except -> sorties 0/1/2

BNP Paribas Cash Management - Direction Monetique
Mars 2026
"""

import argparse
import os
import re
import csv
import sys
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import numpy as np

# --- CONFIGURATION ---
VERSION_ID = "N7M4P"

# Index des colonnes a sacraliser (base 0)
SACRAL_COL_FIXED = [4, 8, 10]  # Col 5, 9, 11 (fixes, pas configurables)
RC_COL_DEFAULT = 9              # Colonne 10 en base 1 -> index 9 en base 0 (configurable)

MOIS_NAMES = {
    '01': 'JANVIER', '02': 'FEVRIER', '03': 'MARS', '04': 'AVRIL',
    '05': 'MAI', '06': 'JUIN', '07': 'JUILLET', '08': 'AOUT',
    '09': 'SEPTEMBRE', '10': 'OCTOBRE', '11': 'NOVEMBRE', '12': 'DECEMBRE'
}


def extract_period_from_filename(filepath):
    """
    Extrait la periode YYYYMM depuis le nom du fichier MONEXT PECC52.

    Les fichiers suivent le pattern : extract_pecc-5212_V4_MMYYYY
    Exemple : extract_pecc-5212_V4_012025 -> periode = 012025 -> YYYYMM = 202501
    """
    basename = os.path.basename(filepath)

    # Pattern principal : _MMYYYY en fin de nom (avant extension eventuelle)
    match = re.search(r'_(\d{2})(\d{4})(?:\.|$)', basename)
    if match:
        month = match.group(1)
        year = match.group(2)
        # Validation basique
        if 1 <= int(month) <= 12 and 2000 <= int(year) <= 2099:
            return f"{year}{month}"

    # Fallback : cherche toute sequence MMYYYY plausible
    match = re.search(r'(\d{2})(20\d{2})', basename)
    if match:
        month = match.group(1)
        year = match.group(2)
        if 1 <= int(month) <= 12:
            return f"{year}{month}"

    return None


def period_to_label(period_yyyymm):
    """Convertit YYYYMM en label lisible : ex 202501 -> JANVIER 2025"""
    if not period_yyyymm or len(period_yyyymm) != 6:
        return period_yyyymm or "INCONNU"
    year = period_yyyymm[:4]
    month = period_yyyymm[4:]
    month_name = MOIS_NAMES.get(month, f"MOIS{month}")
    return f"{month_name} {year}"


def period_sort_key(period_yyyymm):
    """Cle de tri pour les periodes YYYYMM"""
    try:
        return int(period_yyyymm)
    except (ValueError, TypeError):
        return 0


def sacralise_id(val):
    """
    Sacralise un identifiant pour empecher toute conversion Excel.
    Retourne le format ="valeur" qui preserve les zeros initiaux.
    """
    if pd.isna(val):
        return ''
    s = str(val).strip()
    if s == '' or s.lower() in ('nan', 'none', 'null', 'na', 'n/a'):
        return ''
    # Si deja au format ="...", ne pas re-sacraliser
    if s.startswith('="') and s.endswith('"'):
        return s
    # Nettoyer le .0 eventuel (arrondi pandas/Excel)
    if s.endswith('.0') and s[:-2].replace('-', '').isdigit():
        s = s[:-2]
    # Retirer quotes simples eventuelles
    s = s.lstrip("'")
    return f'="{s}"'


def load_csv_smart(path, nrows=None):
    """Chargement CSV intelligent avec detection auto separateur/encodage"""
    for sep in [';', ',', '\t']:
        for enc in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
            try:
                df_test = pd.read_csv(path, sep=sep, encoding=enc, dtype=str,
                                      keep_default_na=False, na_values=[],
                                      on_bad_lines='skip', nrows=5)
                if df_test.shape[1] > 1:
                    return pd.read_csv(path, sep=sep, encoding=enc, dtype=str,
                                       keep_default_na=False, na_values=[],
                                       on_bad_lines='skip', nrows=nrows)
            except Exception:
                continue
    return pd.read_csv(path, sep=None, engine='python', dtype=str,
                       keep_default_na=False, na_values=[], on_bad_lines='skip', nrows=nrows)


def sacralise_colonnes(df, col_indices, col_names):
    """
    Sacralise un ensemble de colonnes par leurs indices (base 0).
    col_indices : liste d'indices (base 0)
    col_names   : liste des noms de colonnes du dataframe (reference)
    """
    nb_cols = len(col_names)
    for idx in col_indices:
        if idx < nb_cols:
            col_name = col_names[idx]
            df[col_name] = df[col_name].apply(sacralise_id)
    return df


# =============================================================================
# CONFIG SACRALISATION (numero de colonne RC configurable via --rc-col)
# =============================================================================
def get_rc_col_index(rc_col_base1: Optional[int]) -> int:
    """Recupere l'index de la colonne RC configurable (convertit base 1 -> base 0)"""
    try:
        val = int(rc_col_base1)
        return val - 1
    except (ValueError, TypeError):
        return RC_COL_DEFAULT


def get_all_sacral_indices(rc_col_base1: Optional[int]) -> List[int]:
    """Retourne la liste complete des indices a sacraliser (fixes + RC configurable)"""
    rc_idx = get_rc_col_index(rc_col_base1)
    all_indices = list(SACRAL_COL_FIXED)
    if rc_idx not in all_indices:
        all_indices.append(rc_idx)
    return sorted(set(all_indices))


# =============================================================================
# COLLECTE DES FICHIERS MONEXT
# =============================================================================
def collect_monext_files(input_folder: Path) -> List[Tuple[str, str]]:
    """
    Liste les fichiers MONEXT du dossier, detecte leur periode et les trie.
    Retourne une liste de (path, period_yyyymm) triee par periode.
    """
    fichiers_monext: List[Tuple[str, str]] = []
    erreurs: List[str] = []

    paths = sorted(
        str(p) for p in input_folder.iterdir()
        if p.is_file() and p.suffix.lower() == '.csv'
    )

    for path in paths:
        period = extract_period_from_filename(path)
        if period:
            fichiers_monext.append((path, period))
        else:
            erreurs.append(f"  x PERIODE NON DETECTEE  <-  {os.path.basename(path)}")

    # Trier par periode
    fichiers_monext.sort(key=lambda x: period_sort_key(x[1]))

    print(f"{'='*60}")
    print(f"  {len(fichiers_monext)} fichier(s) detecte(s) - tries par periode")
    print(f"{'='*60}")
    for path, period in fichiers_monext:
        label = period_to_label(period)
        print(f"  + {label}  ({period})  <-  {os.path.basename(path)}")

    if erreurs:
        print("\n  Fichiers non reconnus :")
        for line in erreurs:
            print(line)

    return fichiers_monext


# =============================================================================
# WORKERS (logique de fusion / sacralisation preservee a l'identique)
# =============================================================================
def worker_historique(fichiers_monext: List[Tuple[str, str]],
                      output_file: Path,
                      rc_col_base1: Optional[int]) -> pd.DataFrame:
    """Mode concatenation historique : fusionne tous les fichiers MONEXT du dossier"""
    if not fichiers_monext:
        raise ValueError("Aucun fichier MONEXT detecte dans le dossier d'entree.")

    total_files = len(fichiers_monext)
    print(f"[0/{total_files}] Demarrage - {total_files} fichier(s) a traiter...")

    # Charger le premier fichier pour recuperer les en-tetes de reference
    first_path, first_period = fichiers_monext[0]
    print(f"[0/{total_files}] Lecture des en-tetes depuis {os.path.basename(first_path)}...")
    df_first = load_csv_smart(first_path)
    original_columns = list(df_first.columns)
    nb_cols_source = len(original_columns)

    # Verification des indices a sacraliser
    sacral_indices = get_all_sacral_indices(rc_col_base1)
    indices_invalides = [i for i in sacral_indices if i >= nb_cols_source]
    if indices_invalides:
        cols_invalides_str = ", ".join([str(i + 1) for i in indices_invalides])
        raise ValueError(
            f"Les colonnes suivantes depassent le nombre de colonnes du fichier ({nb_cols_source}) : "
            f"Colonnes n. : {cols_invalides_str}. Verifiez le numero de colonne RC.")

    sacral_names = [original_columns[i] for i in sacral_indices if i < nb_cols_source]
    print(f"[0/{total_files}] Colonnes sacralisees : {', '.join(sacral_names)}")

    # Traiter tous les fichiers
    all_chunks = []
    total_rows = 0

    for file_idx, (path, period) in enumerate(fichiers_monext):
        label = period_to_label(period)
        print(f"[{file_idx + 1}/{total_files}] Traitement : {label}...")

        df = load_csv_smart(path)

        # Aligner le nombre de colonnes sur le premier fichier
        current_nb_cols = len(df.columns)
        if current_nb_cols < nb_cols_source:
            for i in range(nb_cols_source - current_nb_cols):
                df[f'_extra_{i}'] = ''
        elif current_nb_cols > nb_cols_source:
            df = df.iloc[:, :nb_cols_source]

        # Renommer les colonnes pour matcher le premier fichier
        df.columns = original_columns

        # Sacraliser les colonnes critiques
        df = sacralise_colonnes(df, sacral_indices, original_columns)

        all_chunks.append(df)
        total_rows += len(df)

    # Concatener
    print(f"[{total_files}/{total_files}] Concatenation de {total_rows:,} lignes...")
    df_final = pd.concat(all_chunks, ignore_index=True)

    # Export
    print(f"[{total_files}/{total_files}] Export CSV...")
    export_result(df_final, output_file, rc_col_base1)
    return df_final


def worker_incremental(fichiers_monext: List[Tuple[str, str]],
                       existing_file: Path,
                       output_file: Path,
                       rc_col_base1: Optional[int]) -> pd.DataFrame:
    """Mode ajout incremental : ajoute les fichiers MONEXT du dossier a un consolide existant"""
    if not fichiers_monext:
        raise ValueError("Aucun fichier MONEXT a ajouter detecte dans le dossier d'entree.")

    total_files = len(fichiers_monext)

    # Charger le fichier base
    print(f"[0/{total_files}] Chargement du fichier consolide existant...")
    df_base = load_csv_smart(str(existing_file))
    base_columns = list(df_base.columns)
    nb_rows_base = len(df_base)
    nb_cols_base = len(base_columns)

    print(f"[0/{total_files}] Fichier base : {nb_rows_base:,} lignes, {nb_cols_base} colonnes")

    # Verification des indices a sacraliser par rapport au fichier base
    sacral_indices = get_all_sacral_indices(rc_col_base1)
    indices_invalides = [i for i in sacral_indices if i >= nb_cols_base]
    if indices_invalides:
        cols_invalides_str = ", ".join([str(i + 1) for i in indices_invalides])
        raise ValueError(
            f"Les colonnes suivantes depassent le nombre de colonnes du fichier ({nb_cols_base}) : "
            f"Colonnes n. : {cols_invalides_str}")

    all_chunks = [df_base]
    total_added = 0

    for file_idx, (path, period) in enumerate(fichiers_monext):
        # Charger le fichier a ajouter
        df_ajout = load_csv_smart(path)

        period_detected = extract_period_from_filename(path)
        if not period_detected:
            raise ValueError(
                "Impossible de detecter la periode depuis le nom du fichier. "
                "Le fichier doit contenir une date au format MMYYYY dans son nom. "
                "Exemple attendu : extract_pecc-5212_V4_022026.csv")

        label = period_to_label(period_detected)
        print(f"[{file_idx + 1}/{total_files}] Nouveau fichier : {len(df_ajout):,} lignes - Periode : {label}")

        # Aligner les colonnes du fichier ajoute sur le fichier base
        current_nb_cols = len(df_ajout.columns)
        if current_nb_cols < nb_cols_base:
            for i in range(nb_cols_base - current_nb_cols):
                df_ajout[f'_extra_{i}'] = ''
        elif current_nb_cols > nb_cols_base:
            df_ajout = df_ajout.iloc[:, :nb_cols_base]

        df_ajout.columns = base_columns

        # Sacraliser les colonnes critiques
        print(f"[{file_idx + 1}/{total_files}] Sacralisation des colonnes critiques...")
        df_ajout = sacralise_colonnes(df_ajout, sacral_indices, base_columns)

        all_chunks.append(df_ajout)
        total_added += len(df_ajout)

    # Concatener
    print(f"[{total_files}/{total_files}] Concatenation...")
    df_final = pd.concat(all_chunks, ignore_index=True)

    nb_rows_final = len(df_final)
    print(f"[{total_files}/{total_files}] Resultat : {nb_rows_final:,} lignes ({nb_rows_base:,} + {total_added:,})")

    # Export
    print(f"[{total_files}/{total_files}] Export CSV...")
    export_result(df_final, output_file, rc_col_base1)
    return df_final


def export_result(df_final: pd.DataFrame, output_file: Path, rc_col_base1: Optional[int]) -> None:
    """Export du fichier fusionne"""
    save_path = str(output_file)

    print("Ecriture du fichier...")

    # Export CSV avec separateur ; et encodage UTF-8 BOM
    # QUOTE_NONE pour que les ="valeur" ne soient pas re-quotes
    df_final.to_csv(
        save_path,
        sep=';',
        index=False,
        encoding='utf-8-sig',
        quoting=csv.QUOTE_NONE,
        escapechar='\\'
    )

    nb_rows = len(df_final)
    nb_cols = len(df_final.columns)

    print(f"Termine - {nb_rows:,} lignes exportees")

    stats = (
        f"{'='*50}\n"
        f"  FUSION TERMINEE\n"
        f"{'='*50}\n\n"
        f"Fichier : {os.path.basename(save_path)}\n"
        f"Lignes : {nb_rows:,}\n"
        f"Colonnes : {nb_cols}\n\n"
        f"Colonnes sacralisees :\n"
        f"  - Col 5  : Identifiant bancaire\n"
        f"  - Col 9  : Identifiant RP\n"
        f"  - Col {get_rc_col_index(rc_col_base1) + 1}  : Identifiant RC\n"
        f"  - Col 11 : IBAN\n\n"
        f"Emplacement :\n{save_path}"
    )

    print(stats)


# =============================================================================
# MAIN
# =============================================================================
def main() -> int:
    parser = argparse.ArgumentParser(
        prog="03.N7M4P.py",
        description=(
            "MONEXT FUSION PECC52 [N7M4P] - Fusion des fichiers MONEXT PECC-5212 "
            "mensuels en un CSV consolide multi-annee. Mode par defaut = "
            "concatenation historique de tous les fichiers du dossier. "
            "Mode incremental active par --existing-file."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--input-folder", required=True, type=Path,
        help="Dossier contenant les fichiers PECC mensuels (obligatoire).")
    parser.add_argument(
        "--output-file", required=True, type=Path,
        help="Chemin du CSV consolide a produire (obligatoire).")
    parser.add_argument(
        "--existing-file", required=False, type=Path, default=None,
        help="CSV consolide existant : active le mode incremental "
             "(ajoute les fichiers du dossier a ce consolide).")
    parser.add_argument(
        "--rc-col", required=False, type=int, default=10,
        help="N. de colonne RC a sacraliser (base 1, defaut 10).")

    args = parser.parse_args()

    try:
        input_folder: Path = args.input_folder
        output_file: Path = args.output_file
        existing_file: Optional[Path] = args.existing_file
        rc_col_base1: int = args.rc_col

        # Validation des entrees
        if not input_folder.is_dir():
            print(f"[ERREUR] Dossier d'entree introuvable : {input_folder}", file=sys.stderr)
            return 2
        if existing_file is not None and not existing_file.is_file():
            print(f"[ERREUR] Fichier consolide existant introuvable : {existing_file}", file=sys.stderr)
            return 2

        # Preparer le dossier de sortie
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # Collecte des fichiers MONEXT
        fichiers_monext = collect_monext_files(input_folder)

        if existing_file is not None:
            print("\n>>> MODE : AJOUT INCREMENTAL\n")
            worker_incremental(fichiers_monext, existing_file, output_file, rc_col_base1)
        else:
            print("\n>>> MODE : CONCATENATION HISTORIQUE\n")
            worker_historique(fichiers_monext, output_file, rc_col_base1)

        print(f"[OK] Fusion terminee - fichier consolide ecrit : {output_file}")
        return 0

    except ValueError as exc:
        print(f"[ERREUR] {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # noqa: BLE001
        import traceback
        traceback.print_exc()
        print(f"[ERREUR] Echec inattendu : {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
