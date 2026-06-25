"""
CCO STOCK AGREGATEUR v4 [66Z3Q]
=================================

DESCRIPTION
-----------
Agrege les fichiers CCO Stock d'un DOSSIER (CSV ou XLSX onglet REPORT_)
en 1 SEUL CSV normalise au schema fixe. Detection auto de la colonne
mois (ex: JANUARY-2025) servant de "Stock", sacralisation des identifiants
Portfolio_ID / Client_Mgr_ID / ID_RP (format ="VALEUR" anti notation
scientifique Excel), extraction vectorisee de l'ID RP depuis
Subsidiary_Department.
Deux modes : complet (from scratch) ou incremental (CSV existant + nouveaux
fichiers). La logique metier est strictement identique a l'application GUI
d'origine ; seule l'interface a ete remplacee par une CLI argparse autonome.

MEME LOGIQUE QUE CCO FLUX AGREGATEUR v4 [N3096]
Sauf que :
- 1 SEUL CSV en sortie (pas de separation par type Data)
- Schema adapte au fichier Stock

SOURCES REQUISES
----------------
- --input-folder : dossier contenant les fichiers CCO_Stock
    * CSV  : separateur auto-detecte (; , tab), encodage auto
    * XLSX : un onglet dont le nom commence par "REPORT_"
    * La DATE est extraite du NOM DU FICHIER (motif AAAAMM, ex: 202501)
- --existing-file (mode incremental uniquement) : CSV Stock deja agrege
    a completer (separateur ; , encodage utf-8-sig)

OUTPUTS PRODUITS
----------------
- 1 CSV normalise (--output-file), separateur ";", encodage utf-8-sig.
  Schema de sortie fixe (12 colonnes) :
    DATE | Organization_Unit_Code | Portfolio_ID | Client_Mgr_ID |
    Company | Subsidiary_Department | ID_RP | Product | Data | Stock |
    Bank_Type | Bin_Type

  - DATE         : extrait du NOM DU FICHIER (format MOIS_ANNEE, ex: JANVIER_2025)
  - Portfolio_ID : sacralise ="VALEUR" (peut commencer par 0)
  - Client_Mgr_ID: sacralise ="VALEUR" (peut commencer par 0)
  - Subsidiary_Department : colonne encodee originale (2_VCE1_NOM_IDRP...)
  - ID_RP        : extrait de Subsidiary_Department, sacralise ="VALEUR"
  - Product      : nom du programme carte
  - Data         : Active / New Subscriber / etc. (pas de separation en datasets)
  - Stock        : contenu de la colonne mois (entier relatif = nombre de cartes)

ARGUMENTS CLI
-------------
  --input-folder PATH   (OBLIGATOIRE) Dossier source des CSV/XLSX a agreger.
  --output-file  PATH   (OBLIGATOIRE) Chemin du CSV de sortie normalise.
  --existing-file PATH  (OPTIONNEL)   CSV Stock deja agrege a completer.
                                      Sa presence active le MODE INCREMENTAL ;
                                      en son absence, MODE COMPLET par defaut.

  Codes de sortie : 0 = succes | 1 = erreur d'execution | 2 = erreur d'arguments

DECOMPOSITION
-------------
main(argv)
├── parse_args()                      lecture/validation des flags argparse
├── run_aggregation(input_folder, output_file, existing_file)
│   ├── glob des fichiers CSV/XLSX du dossier
│   ├── [incremental] chargement du CSV existant
│   ├── boucle fichiers -> process_file(...)
│   │   ├── load_raw_df(...)          chargement brut CSV/XLSX
│   │   ├── find_header_row_idx(...)  detection ligne d'entete
│   │   ├── identify_columns(...)     mapping des roles de colonnes
│   │   ├── sacralise_vec(...)        sacralisation ="VALEUR"
│   │   └── extract_id_rp_vec(...)    extraction ID RP vectorisee
│   ├── alignement + concat des chunks
│   └── export CSV final (sep=';', utf-8-sig)
└── print("[OK] ...")                 rapport final

REPERAGE DES COLONNES (robuste aux coquilles et espaces) :
  - Organization_Unit_Code : header contient "organization" ou "unit"
  - Portfolio_ID           : header contient "portfolio"
  - Client_Mgr_ID          : header contient "client" et "mgr"
  - Company                : header contient "company"
  - Subsidiary_Department  : detecte par CONTENU (valeurs matchant N_XXX_NOM_IDRP)
  - Product                : header contient "product"
  - Data                   : header == "data" (strip lowercase)
  - Stock                  : header matchant pattern mois (JANUARY-2022, SEPTEMBER-2025...)
  - Bank_Type              : header contient "bank"
  - Bin_Type               : header contient "bin"

OPTIMISATIONS :
  - Sacralisation : vectorisee via np.where
  - Extraction ID_RP : vectorisee via str.extract (regex pandas)
  - Zero iterrows()

BNP Paribas Cash Management - Direction Monetique
Mars 2026
"""

import argparse
import glob
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd

VERSION_ID = "66Z3Q"

# ─── Schema de sortie fixe ───────────────────────────────────────────────────
OUTPUT_COLS = [
    "DATE",
    "Organization_Unit_Code",
    "Portfolio_ID",
    "Client_Mgr_ID",
    "Company",
    "Subsidiary_Department",
    "ID_RP",
    "Product",
    "Data",
    "Stock",
    "Bank_Type",
    "Bin_Type",
]

# Pattern mois dans les headers
MONTH_HDR_RE = re.compile(
    r"(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|"
    r"SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER|"
    r"JAN|FEB|MAR|APR|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
    r"[-_]?\d{4}",
    re.IGNORECASE,
)

MOIS_FR = {
    "01": "JANVIER",  "02": "FEVRIER",  "03": "MARS",
    "04": "AVRIL",    "05": "MAI",       "06": "JUIN",
    "07": "JUILLET",  "08": "AOUT",      "09": "SEPTEMBRE",
    "10": "OCTOBRE",  "11": "NOVEMBRE",  "12": "DECEMBRE",
}

MIN_NONEMPTY_HEADER = 4


# =============================================================================
# UTILITAIRES
# =============================================================================

def extract_date_from_filename(filepath):
    bn = os.path.basename(filepath)
    m = re.search(r"(\d{4})(\d{2})", bn)
    if m:
        year, month = m.group(1), m.group(2)
        if 2000 <= int(year) <= 2099 and 1 <= int(month) <= 12:
            return f"{MOIS_FR.get(month, month)}_{year}"
    return "INCONNU"


def count_nonempty(row):
    return sum(1 for v in row
               if str(v).strip() not in ("", "nan", "None", "NaN"))


def detect_report_sheet(sheet_names):
    for s in sheet_names:
        if s.upper().startswith("REPORT_"):
            return s
    return None


def normalize_header(h):
    return str(h).strip().lower()


# =============================================================================
# SACRALISATION VECTORISEE
# =============================================================================

def sacralise_vec(series):
    """
    Sacralise les valeurs numeriques (y compris celles commencant par 0) :
    format ="VALEUR" pour empecher la notation scientifique Excel.
    Vectorise via np.where.
    """
    s = series.astype(str).str.strip().str.lstrip("'")
    already = s.str.startswith('="') & s.str.endswith('"')
    is_num  = s.str.match(r"^\d+$") | s.str.match(r"^0\d+$")
    result  = np.where(already, s,
              np.where(is_num, '="' + s + '"', s))
    return pd.Series(result, index=series.index)


# =============================================================================
# EXTRACTION ID_RP VECTORISEE
# =============================================================================

def extract_id_rp_vec(series):
    """
    Extrait l'ID RP depuis la colonne Subsidiary_Department.
    Format : 2_VCE1_NOM_032700044008000000
    Regle  : dernier segment numerique >= 10 chiffres apres "_"
    Vectorise via str.extract.
    """
    s = series.astype(str).str.strip()
    extracted = s.str.extract(r"_(\d{10,})$", expand=False)
    # Fallback pour les cas sans match en fin de chaine
    mask_empty = extracted.isna() | (extracted == "")
    if mask_empty.any():
        fallback = s[mask_empty].str.findall(r"\d{10,}").apply(
            lambda lst: lst[-1] if lst else ""
        )
        extracted[mask_empty] = fallback
    return extracted.fillna("")


# =============================================================================
# CHARGEMENT BRUT
# =============================================================================

def load_raw_df(filepath):
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".xlsx", ".xls"):
        try:
            xl = pd.ExcelFile(filepath)
            sheet = detect_report_sheet(xl.sheet_names)
            if not sheet:
                return None, None, f"Aucun onglet REPORT_ : {os.path.basename(filepath)}"
            df = pd.read_excel(filepath, sheet_name=sheet, header=None,
                               dtype=str, keep_default_na=False, na_values=[])
            return df, sheet, None
        except Exception as e:
            return None, None, str(e)

    elif ext == ".csv":
        for sep in [";", ",", "\t"]:
            for enc in ["utf-8-sig", "utf-8", "latin1", "cp1252"]:
                try:
                    df = pd.read_csv(filepath, sep=sep, encoding=enc,
                                     header=None, dtype=str,
                                     keep_default_na=False, na_values=[],
                                     on_bad_lines="skip")
                    if df.shape[1] > 3:
                        return df, None, None
                except Exception:
                    continue
        return None, None, f"Impossible de lire le CSV : {os.path.basename(filepath)}"

    return None, None, f"Format non supporte : {os.path.basename(filepath)}"


def find_header_row_idx(df_raw):
    for i in range(min(6, len(df_raw))):
        if count_nonempty(df_raw.iloc[i].tolist()) >= MIN_NONEMPTY_HEADER:
            return i
    return 0


# =============================================================================
# IDENTIFICATION DES COLONNES
# =============================================================================

def identify_columns(df):
    """
    Identifie le role de chaque colonne par son header.
    Robuste aux coquilles, espaces superflus, variantes de noms.
    """
    mapping = {}
    cols = list(df.columns)

    for col in cols:
        nh = normalize_header(col)

        if ("organization" in nh or "unit_code" in nh or
                ("unit" in nh and "org" in nh)) and "org_unit" not in mapping:
            mapping["org_unit"] = col

        elif "portfolio" in nh and "portfolio" not in mapping:
            mapping["portfolio"] = col

        elif "client" in nh and "mgr" in nh and "client_mgr" not in mapping:
            mapping["client_mgr"] = col

        elif "company" in nh and "company" not in mapping:
            mapping["company"] = col

        elif "subsidiary" in nh and "subsidiary" not in mapping:
            mapping["subsidiary"] = col

        elif nh.strip() == "product" and "product" not in mapping:
            mapping["product"] = col

        elif nh.strip() == "data" and "data" not in mapping:
            mapping["data"] = col

        elif "bank" in nh and "bank_type" not in mapping:
            mapping["bank_type"] = col

        elif "bin" in nh and "bin_type" not in mapping:
            mapping["bin_type"] = col

        # Colonne mois : header ressemble a JANUARY-2022, SEPTEMBER-2025...
        elif MONTH_HDR_RE.search(col) and "stock_col" not in mapping:
            mapping["stock_col"] = col

    # Fallback Subsidiary : detecte par contenu si header non reconnu
    if "subsidiary" not in mapping:
        mapped_so_far = set(mapping.values())
        for col in cols:
            if col in mapped_so_far:
                continue
            sample = df[col].dropna().astype(str).head(30)
            if sample.str.contains(r"_\d{10,}", regex=True).any():
                mapping["subsidiary"] = col
                break

    return mapping


# =============================================================================
# TRAITEMENT D'UN FICHIER -> DataFrame normalise
# =============================================================================

def process_file(filepath, date_str):
    warnings = []
    fname = os.path.basename(filepath)

    df_raw, sheet, err = load_raw_df(filepath)
    if df_raw is None:
        return None, [err or f"Erreur inconnue : {fname}"]

    header_idx = find_header_row_idx(df_raw)

    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(filepath, sheet_name=sheet, header=header_idx,
                               dtype=str, keep_default_na=False, na_values=[])
        else:
            df_raw2, _, _ = load_raw_df(filepath)
            col_names = df_raw2.iloc[header_idx].astype(str).tolist()
            df = df_raw2.iloc[header_idx + 1:].copy()
            df.columns = col_names
            df = df.reset_index(drop=True)
    except Exception as e:
        return None, [f"{fname} : erreur rechargement -> {e}"]

    if df.empty:
        return None, [f"{fname} : DataFrame vide apres header"]

    df.columns = [str(c).strip() for c in df.columns]

    m = identify_columns(df)

    n = len(df)
    out = {col: pd.Series([""] * n, dtype=str) for col in OUTPUT_COLS}

    out["DATE"] = pd.Series([date_str] * n)

    if "org_unit" in m:
        out["Organization_Unit_Code"] = df[m["org_unit"]].astype(str).str.strip()

    if "portfolio" in m:
        # Sacraliser Portfolio_ID (peut commencer par 0)
        out["Portfolio_ID"] = sacralise_vec(df[m["portfolio"]].astype(str).str.strip())

    if "client_mgr" in m:
        # Sacraliser Client_Mgr_ID aussi
        out["Client_Mgr_ID"] = sacralise_vec(df[m["client_mgr"]].astype(str).str.strip())

    if "company" in m:
        out["Company"] = df[m["company"]].astype(str).str.strip()

    if "subsidiary" in m:
        out["Subsidiary_Department"] = df[m["subsidiary"]].astype(str).str.strip()
        out["ID_RP"] = sacralise_vec(extract_id_rp_vec(df[m["subsidiary"]]))
    else:
        warnings.append(f"{fname} : colonne Subsidiary_Department non trouvee — ID_RP vide")

    if "product" in m:
        out["Product"] = df[m["product"]].astype(str).str.strip()

    if "data" in m:
        out["Data"] = df[m["data"]].astype(str).str.strip()
    else:
        warnings.append(f"{fname} : colonne Data non trouvee")

    if "stock_col" in m:
        out["Stock"] = df[m["stock_col"]].astype(str).str.strip()
    else:
        warnings.append(f"{fname} : colonne Stock (mois) non trouvee")

    if "bank_type" in m:
        out["Bank_Type"] = df[m["bank_type"]].astype(str).str.strip()

    if "bin_type" in m:
        out["Bin_Type"] = df[m["bin_type"]].astype(str).str.strip()

    df_out = pd.DataFrame(out)[OUTPUT_COLS]

    # Supprimer lignes completement vides (hors DATE)
    mask = df_out[OUTPUT_COLS[1:]].apply(
        lambda row: row.str.strip().ne("").any(), axis=1
    )
    df_out = df_out[mask].reset_index(drop=True)

    return df_out, warnings


# =============================================================================
# AGREGATION (worker, logique inchangee)
# =============================================================================

def run_aggregation(input_folder: Path, output_file: Path,
                    existing_file: Path | None) -> int:
    """
    Reproduit a l'identique le worker GUI : glob des fichiers du dossier,
    chargement optionnel d'un CSV existant (mode incremental), traitement
    fichier par fichier, concatenation alignee et export du CSV final.
    Progression remplacee par des print("[i/N] ...").
    Retourne le nombre de fichiers traites.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    mode = "increment" if existing_file is not None else "complet"

    # ── Detection des fichiers du dossier ──
    fichiers = []
    for ext in ("*.csv", "*.xlsx", "*.xls"):
        fichiers += glob.glob(os.path.join(str(input_folder), ext))
    fichiers = sorted(fichiers)

    if not fichiers:
        raise RuntimeError(
            f"Aucun fichier CSV/XLSX trouve dans le dossier : {input_folder}")

    n = len(fichiers)
    n_csv  = sum(1 for f in fichiers if f.lower().endswith(".csv"))
    n_xlsx = n - n_csv
    print(f"{n} fichier(s) detecte(s) — {n_csv} CSV | {n_xlsx} XLSX")
    print(f"Mode : {'Incremental' if mode == 'increment' else 'Complet'}")

    all_chunks   = []
    all_warnings = []

    # ── Charger CSV existant (incremental) ──
    if mode == "increment" and existing_file:
        print("Chargement CSV existant...")
        try:
            df_ex = pd.read_csv(str(existing_file), sep=";", dtype=str,
                                keep_default_na=False, na_values=[],
                                encoding="utf-8-sig")
            for col in OUTPUT_COLS:
                if col not in df_ex.columns:
                    df_ex[col] = ""
            all_chunks.append(df_ex[OUTPUT_COLS])
        except Exception as e:
            all_warnings.append(f"CSV existant non charge : {e}")

    # ── Traiter les fichiers du dossier ──
    for idx, filepath in enumerate(fichiers):
        date_str = extract_date_from_filename(filepath)
        print(f"[{idx+1}/{n}] Traitement : {os.path.basename(filepath)}...")

        df_norm, warns = process_file(filepath, date_str)
        all_warnings.extend(warns)

        if df_norm is not None and not df_norm.empty:
            all_chunks.append(df_norm)

    # ── Concatener ──
    print("Concatenation finale...")
    if not all_chunks:
        raise RuntimeError("Aucune donnee a exporter.")

    # Aligner toutes les colonnes avant concat
    aligned = []
    for chunk in all_chunks:
        for col in OUTPUT_COLS:
            if col not in chunk.columns:
                chunk = chunk.copy()
                chunk[col] = ""
        aligned.append(chunk[OUTPUT_COLS])

    df_final = pd.concat(aligned, ignore_index=True)

    # ── Exporter ──
    print("Export CSV...")
    output_file.parent.mkdir(parents=True, exist_ok=True)
    df_final.to_csv(str(output_file), sep=";", index=False, encoding="utf-8-sig")

    lines = [
        f"Agregation terminee — {n} fichier(s) traite(s)",
        f"Mode : {'Incremental' if mode == 'increment' else 'Complet'}",
        f"Schema : {' | '.join(OUTPUT_COLS)}",
        f"Fichier : {os.path.basename(str(output_file))}",
        f"Lignes  : {len(df_final):,}",
    ]
    if all_warnings:
        lines.append(f"Avertissements ({len(all_warnings)}) :")
        for w in all_warnings[:20]:
            lines.append(f"  - {w}")
        if len(all_warnings) > 20:
            lines.append(f"  ... et {len(all_warnings)-20} autres")
    for ln in lines:
        print(ln)

    return n


# =============================================================================
# CLI
# =============================================================================

def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="02.66Z3Q",
        description="CCO STOCK AGREGATEUR v4 [66Z3Q] - agrege un dossier de "
                    "fichiers CCO_Stock (CSV/XLSX onglet REPORT_) en 1 CSV "
                    "normalise.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--input-folder", required=True, type=Path,
        help="(OBLIGATOIRE) Dossier source contenant les CSV/XLSX a agreger.",
    )
    parser.add_argument(
        "--output-file", required=True, type=Path,
        help="(OBLIGATOIRE) Chemin du CSV de sortie normalise.",
    )
    parser.add_argument(
        "--existing-file", required=False, type=Path, default=None,
        help="(OPTIONNEL) CSV Stock deja agrege a completer. Sa presence "
             "active le MODE INCREMENTAL ; en son absence, MODE COMPLET.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        input_folder: Path = args.input_folder
        output_file: Path = args.output_file
        existing_file: Path | None = args.existing_file

        if not input_folder.is_dir():
            print(f"[ERREUR] Dossier source introuvable : {input_folder}",
                  file=sys.stderr)
            return 1

        if existing_file is not None and not existing_file.is_file():
            print(f"[ERREUR] CSV existant introuvable : {existing_file}",
                  file=sys.stderr)
            return 1

        n = run_aggregation(input_folder, output_file, existing_file)

        print(f"[OK] Agregation terminee : {n} fichier(s) -> {output_file}")
        return 0

    except Exception as e:
        print(f"[ERREUR] {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    sys.exit(main())
