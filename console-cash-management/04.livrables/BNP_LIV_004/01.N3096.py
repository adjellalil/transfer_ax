"""
CCO FLUX AGREGATEUR v4 [N3096]
================================

DESCRIPTION :
  Agrege les fichiers CCO Flux (CSV ou XLSX onglet REPORT_) contenus dans un
  DOSSIER source et produit 3 CSV normalises : NUMBER / AMOUNT / INTERCHANGE.
  Detection automatique du separateur, des colonnes et du mois ; extraction et
  sacralisation de l'ID_RP ; modes complet (par defaut) ou incremental
  (completion de CSV consolides existants). Aucune regle metier n'est modifiee
  par rapport a la version GUI d'origine. Version CLI autonome, sans interface
  graphique.

  SCHEMA DE SORTIE FIXE :
    DATE | Company | Code | ID_RP | Product | Trx_Type | Currency | Data |
    Valeur | Bank_Type | Bin_Type
    - DATE   : extrait du NOM DU FICHIER (format MOIS_ANNEE, ex: JANVIER_2025)
    - Code   : colonne concatenee originale (2_VCE1_NOM_IDRP)
    - ID_RP  : extrait de Code (dernier segment numerique >= 10 chiffres),
               sacralise ="VALEUR"
    - Valeur : contenu de la colonne mois (unique par fichier)

SOURCES REQUISES :
  - Dossier source contenant des fichiers CCO Flux (*.csv, *.xlsx, *.xls).
    Pour les XLSX, l'onglet doit commencer par REPORT_.
  - (Mode incremental, optionnel) Dossier contenant les CSV consolides
    existants nommes <base>_NUMBER.csv, <base>_AMOUNT.csv,
    <base>_INTERCHANGE.csv a completer.

OUTPUTS PRODUITS :
  - <base sans extension>_NUMBER.csv      (dataset NUMBER, sep ";", utf-8-sig)
  - <base sans extension>_AMOUNT.csv      (dataset AMOUNT, sep ";", utf-8-sig)
  - <base sans extension>_INTERCHANGE.csv (dataset INTERCHANGE, sep ";", utf-8-sig)

ARGUMENTS CLI :
  --input-folder PATH (obligatoire) dossier source contenant les CSV/XLSX a agreger
  --output-file PATH (obligatoire) chemin de BASE de sortie ; les 3 fichiers en
    sont derives par ajout des suffixes _NUMBER, _AMOUNT, _INTERCHANGE (.csv)
  --existing-folder PATH (optionnel) dossier des CSV consolides existants a
    completer (active le mode incremental ; absent = mode complet)

DECOMPOSITION :
  main()
  |- argparse (parse des arguments CLI)
  |- run_aggregation()
  |  |- scan du dossier source (glob *.csv / *.xlsx / *.xls)
  |  |- [increment] chargement des CSV consolides existants
  |  |- pour chaque fichier :
  |  |  |- extract_date_from_filename()
  |  |  |- process_file()
  |  |  |  |- load_raw_df()        (XLSX onglet REPORT_ / CSV detection sep+enc)
  |  |  |  |- find_header_row_idx()
  |  |  |  |- identify_columns()
  |  |  |  |- extract_id_rp_vec() + sacralise_vec()
  |  |  |- split_datasets()        (NUMBER / AMOUNT / INTERCHANGE)
  |  |- concatenation finale (pd.concat alignee sur OUTPUT_COLS)
  |  |- export des 3 CSV derives
  |- sorties 0 (succes) / 1 (erreur traitement) / 2 (arguments invalides)

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
from typing import Optional

import numpy as np
import pandas as pd

VERSION_ID = "N3096"

# ─── Schema de sortie fixe ───────────────────────────────────────────────────
OUTPUT_COLS = [
    "DATE", "Company", "Code", "ID_RP",
    "Product", "Trx_Type", "Currency",
    "Data", "Valeur", "Bank_Type", "Bin_Type"
]

# ─── Reperage colonnes ────────────────────────────────────────────────────────
DATA_NUMBER      = "number"
DATA_AMOUNT      = "amount"
DATA_INTERCHANGE = "interchange"

# Pattern mois dans les headers de colonnes
MONTH_HDR_RE = re.compile(
    r"(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|"
    r"SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER|"
    r"JAN|FEB|MAR|APR|JUN|JUL|AUG|SEP|OCT|NOV|DEC)"
    r"[-_]?\d{4}",
    re.IGNORECASE
)

# Pattern ID_RP encode dans la colonne Code
CODE_IDRP_RE = re.compile(r"_(\d{10,})$")

# Mapping numero de mois -> nom FR
MOIS_FR = {
    "01": "JANVIER",  "02": "FEVRIER",  "03": "MARS",
    "04": "AVRIL",    "05": "MAI",       "06": "JUIN",
    "07": "JUILLET",  "08": "AOUT",      "09": "SEPTEMBRE",
    "10": "OCTOBRE",  "11": "NOVEMBRE",  "12": "DECEMBRE",
}

MIN_NONEMPTY_HEADER = 4  # nb min de cellules non vides pour valider une ligne header


# =============================================================================
# UTILITAIRES
# =============================================================================

def extract_date_from_filename(filepath):
    """
    Extrait MOIS_ANNEE depuis le nom de fichier.
    Ex: CCO_Flux_202507.xlsx  -> JUILLET_2025
        CCO_Flux_030004-20250806.xlsx -> AOUT_2025
    """
    bn = os.path.basename(filepath)
    m = re.search(r"(\d{4})(\d{2})", bn)
    if m:
        year, month = m.group(1), m.group(2)
        if 2000 <= int(year) <= 2099 and 1 <= int(month) <= 12:
            return f"{MOIS_FR.get(month, month)}_{year}"
    return "INCONNU"


def count_nonempty(row):
    return sum(1 for v in row if str(v).strip() not in ("", "nan", "None", "NaN"))


def detect_report_sheet(sheet_names):
    for s in sheet_names:
        if s.upper().startswith("REPORT_"):
            return s
    return None


def normalize_header(h):
    """Nettoie un nom de colonne : strip, lowercase, sans espaces multiples."""
    return str(h).strip().lower()


def identify_columns(df):
    """
    Identifie le role de chaque colonne du DataFrame par son header.
    Retourne un dict role -> col_name.
    Robuste aux coquilles et espaces superflus.
    """
    mapping = {}
    cols = list(df.columns)

    for col in cols:
        nh = normalize_header(col)

        if "company" in nh and "company" not in mapping:
            mapping["company"] = col

        elif "product" in nh and "product" not in mapping:
            mapping["product"] = col

        elif ("trx" in nh or "transaction" in nh) and "trx_type" not in mapping:
            mapping["trx_type"] = col

        elif "currenc" in nh and "currency" not in mapping:
            mapping["currency"] = col

        elif nh.strip() == "data" and "data" not in mapping:
            mapping["data"] = col

        elif "bank" in nh and "type" in nh and "bank_type" not in mapping:
            mapping["bank_type"] = col

        elif "bin" in nh and "bin_type" not in mapping:
            mapping["bin_type"] = col

        # Colonne mois : header ressemble a JULY-2025
        elif MONTH_HDR_RE.search(col) and "valeur_col" not in mapping:
            mapping["valeur_col"] = col

        # Colonne ID_RP isolee
        elif re.sub(r"[\s_\-]", "", nh) in ("idrp", "id_rp", "identifiantrp") \
                and "id_rp_isolated" not in mapping:
            mapping["id_rp_isolated"] = col

    # Colonne Code (concatenee) : detectee par contenu (pattern _IDRP)
    # On cherche parmi les colonnes pas encore mappees
    mapped_cols = set(mapping.values())
    for col in cols:
        if col in mapped_cols:
            continue
        sample = df[col].dropna().astype(str).head(30)
        if sample.str.contains(r"_\d{10,}", regex=True).any():
            mapping["code_col"] = col
            break

    return mapping


# =============================================================================
# EXTRACTION ID_RP (vectorisee)
# =============================================================================

def extract_id_rp_vec(series):
    """
    Extrait l'ID RP depuis la colonne Code encodee.
    Vectorise via str.extract sur regex.
    """
    s = series.astype(str).str.strip()
    extracted = s.str.extract(r"_(\d{10,})$", expand=False)
    # Fallback : si pas de match en fin, chercher n'importe quel segment numerique long
    mask_empty = extracted.isna() | (extracted == "")
    if mask_empty.any():
        fallback = s[mask_empty].str.findall(r"\d{10,}").apply(
            lambda lst: lst[-1] if lst else ""
        )
        extracted[mask_empty] = fallback
    return extracted.fillna("")


def sacralise_vec(series):
    """
    Sacralise les IDs numeriques longs : ="VALEUR"
    Vectorise via np.where.
    """
    s = series.astype(str).str.strip().str.lstrip("'")
    # Deja sacralise
    already = s.str.startswith('="') & s.str.endswith('"')
    # Numerique pur ou commence par 0
    is_num = s.str.match(r"^\d+$") | (s.str.match(r"^0\d+$"))
    result = np.where(already, s,
             np.where(is_num, '="' + s + '"', s))
    return pd.Series(result, index=series.index)


# =============================================================================
# CHARGEMENT D'UN FICHIER (XLSX ou CSV)
# =============================================================================

def load_raw_df(filepath):
    """
    Charge le contenu brut du fichier (XLSX -> onglet REPORT_ | CSV -> direct).
    Retourne (df_raw sans header, sheet_name ou None, erreur ou None).
    """
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

    else:
        return None, None, f"Format non supporte : {os.path.basename(filepath)}"


def find_header_row_idx(df_raw):
    """
    Trouve la premiere ligne avec >= MIN_NONEMPTY_HEADER cellules non vides.
    """
    for i in range(min(6, len(df_raw))):
        if count_nonempty(df_raw.iloc[i].tolist()) >= MIN_NONEMPTY_HEADER:
            return i
    return 0


# =============================================================================
# TRAITEMENT D'UN FICHIER -> DataFrame normalise
# =============================================================================

def process_file(filepath, date_str):
    """
    Charge, detecte le header, identifie les colonnes, normalise.
    Retourne (df_normalise, warnings).
    """
    warnings = []
    fname = os.path.basename(filepath)

    df_raw, sheet, err = load_raw_df(filepath)
    if df_raw is None:
        return None, [err or f"Erreur inconnue : {fname}"]

    # Trouver le header
    header_idx = find_header_row_idx(df_raw)

    # Recharger avec le bon header
    ext = os.path.splitext(filepath)[1].lower()
    try:
        if ext in (".xlsx", ".xls"):
            df = pd.read_excel(filepath, sheet_name=sheet, header=header_idx,
                               dtype=str, keep_default_na=False, na_values=[])
        else:
            # Pour CSV : re-lire avec header=header_idx
            df_raw2, _, _ = load_raw_df(filepath)
            # Extraire les noms de colonnes depuis la ligne header_idx
            col_names = df_raw2.iloc[header_idx].astype(str).tolist()
            df = df_raw2.iloc[header_idx + 1:].copy()
            df.columns = col_names
            df = df.reset_index(drop=True)
    except Exception as e:
        return None, [f"{fname} : erreur rechargement -> {e}"]

    if df.empty:
        return None, [f"{fname} : DataFrame vide apres header"]

    # Nettoyer les noms de colonnes
    df.columns = [str(c).strip() for c in df.columns]

    # Identifier les colonnes
    m = identify_columns(df)

    n = len(df)

    # Construire le DataFrame normalise
    out = {col: pd.Series([""] * n, dtype=str) for col in OUTPUT_COLS}

    out["DATE"] = pd.Series([date_str] * n)

    if "company" in m:
        out["Company"] = df[m["company"]].astype(str).str.strip()

    if "code_col" in m:
        out["Code"] = df[m["code_col"]].astype(str).str.strip()
        # Extraire ID_RP depuis Code
        out["ID_RP"] = sacralise_vec(extract_id_rp_vec(df[m["code_col"]]))

    # Si colonne ID_RP isolee presente, elle prend le dessus
    if "id_rp_isolated" in m:
        isolated = df[m["id_rp_isolated"]].astype(str).str.strip()
        # Sacraliser uniquement si pas deja fait
        out["ID_RP"] = sacralise_vec(isolated)

    if "product" in m:
        out["Product"] = df[m["product"]].astype(str).str.strip()

    if "trx_type" in m:
        out["Trx_Type"] = df[m["trx_type"]].astype(str).str.strip()

    if "currency" in m:
        out["Currency"] = df[m["currency"]].astype(str).str.strip()

    if "data" in m:
        out["Data"] = df[m["data"]].astype(str).str.strip()
    else:
        warnings.append(f"{fname} : colonne Data non trouvee")

    if "valeur_col" in m:
        out["Valeur"] = df[m["valeur_col"]].astype(str).str.strip()
    else:
        warnings.append(f"{fname} : colonne Valeur (mois) non trouvee")

    if "bank_type" in m:
        out["Bank_Type"] = df[m["bank_type"]].astype(str).str.strip()

    if "bin_type" in m:
        out["Bin_Type"] = df[m["bin_type"]].astype(str).str.strip()

    df_out = pd.DataFrame(out)[OUTPUT_COLS]

    # Supprimer les lignes completement vides (hors DATE)
    mask_nonempty = df_out[OUTPUT_COLS[1:]].apply(
        lambda row: row.str.strip().ne("").any(), axis=1
    )
    df_out = df_out[mask_nonempty].reset_index(drop=True)

    return df_out, warnings


# =============================================================================
# SEPARATION EN 3 DATASETS (vectorisee)
# =============================================================================

def split_datasets(df):
    """Separe le DataFrame en 3 selon la colonne Data. Vectorise."""
    if "Data" not in df.columns or df.empty:
        return df.copy(), pd.DataFrame(columns=OUTPUT_COLS), pd.DataFrame(columns=OUTPUT_COLS)

    data_norm = df["Data"].astype(str).str.lower().str.strip()

    is_interchange = data_norm.str.contains(DATA_INTERCHANGE, na=False)
    is_number      = data_norm.str.contains(DATA_NUMBER, na=False) & ~is_interchange
    is_amount      = data_norm.str.contains(DATA_AMOUNT, na=False) & ~is_interchange & ~is_number

    return (
        df[is_number].copy().reset_index(drop=True),
        df[is_amount].copy().reset_index(drop=True),
        df[is_interchange].copy().reset_index(drop=True),
    )


# =============================================================================
# DERIVATION DES CHEMINS DE SORTIE
# =============================================================================

def derive_output_paths(output_file: Path) -> dict:
    """
    Derive les 3 chemins de sortie a partir du chemin de BASE fourni.
    <base sans extension>_NUMBER.csv / _AMOUNT.csv / _INTERCHANGE.csv
    """
    base = output_file.with_suffix("")
    parent = base.parent
    stem = base.name
    return {
        "NUMBER":      parent / f"{stem}_NUMBER.csv",
        "AMOUNT":      parent / f"{stem}_AMOUNT.csv",
        "INTERCHANGE": parent / f"{stem}_INTERCHANGE.csv",
    }


# =============================================================================
# WORKER (logique d'agregation, hors GUI) — comportement preserve a l'identique
# =============================================================================

def run_aggregation(input_folder: Path,
                    output_file: Path,
                    existing_folder: Optional[Path]) -> int:
    """
    Agrege les fichiers du dossier source et exporte les 3 CSV derives.
    Retourne un code de sortie (0 succes, 1 erreur).
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")  # noqa: F841 (conserve)
    mode = "increment" if existing_folder is not None else "complet"

    # Scanner les fichiers valides
    fichiers = []
    for ext in ("*.csv", "*.xlsx", "*.xls"):
        fichiers += glob.glob(os.path.join(str(input_folder), ext))
    fichiers = sorted(fichiers)

    if not fichiers:
        print(f"[ERREUR] Aucun fichier CSV/XLSX trouve dans : {input_folder}")
        return 1

    n = len(fichiers)
    all_warnings = []

    # CSV existants derives du meme schema de nommage que les sorties
    existing_paths = {}
    if mode == "increment":
        base_stem = output_file.with_suffix("").name
        for label in ("NUMBER", "AMOUNT", "INTERCHANGE"):
            existing_paths[label] = existing_folder / f"{base_stem}_{label}.csv"

    print(f"[i] Mode : {'Incremental' if mode == 'increment' else 'Complet'}")
    print(f"[i] {n} fichier(s) detecte(s) dans : {input_folder}")

    # ── Etape 1 : charger CSV existants (mode incremental) ──
    all_num, all_amt, all_int = [], [], []
    if mode == "increment":
        for label, chunks in [("NUMBER", all_num), ("AMOUNT", all_amt),
                               ("INTERCHANGE", all_int)]:
            path = str(existing_paths[label]) if label in existing_paths else ""
            if path and os.path.exists(path):
                print(f"[i] Chargement CSV existant {label} : {os.path.basename(path)}")
                try:
                    df_ex = pd.read_csv(path, sep=";", dtype=str,
                                        keep_default_na=False, na_values=[],
                                        encoding="utf-8-sig")
                    # Aligner sur OUTPUT_COLS
                    for col in OUTPUT_COLS:
                        if col not in df_ex.columns:
                            df_ex[col] = ""
                    chunks.append(df_ex[OUTPUT_COLS])
                except Exception as e:
                    all_warnings.append(f"CSV {label} non charge : {e}")

    # ── Etape 2 : traiter chaque fichier ──
    for idx, filepath in enumerate(fichiers):
        date_str = extract_date_from_filename(filepath)
        print(f"[{idx+1}/{n}] Traitement : {os.path.basename(filepath)} ({date_str})")

        df_norm, warns = process_file(filepath, date_str)
        all_warnings.extend(warns)

        if df_norm is None or df_norm.empty:
            continue

        df_num, df_amt, df_int = split_datasets(df_norm)
        if not df_num.empty:  all_num.append(df_num)
        if not df_amt.empty:  all_amt.append(df_amt)
        if not df_int.empty:  all_int.append(df_int)

    # ── Etape 3 : concatener ──
    print("[i] Concatenation finale...")

    results = {}
    labels_map = {
        "NUMBER":      all_num,
        "AMOUNT":      all_amt,
        "INTERCHANGE": all_int,
    }
    for label, chunks in labels_map.items():
        if chunks:
            # Aligner avant concat
            aligned = []
            for chunk in chunks:
                for col in OUTPUT_COLS:
                    if col not in chunk.columns:
                        chunk = chunk.copy()
                        chunk[col] = ""
                aligned.append(chunk[OUTPUT_COLS])
            results[label] = pd.concat(aligned, ignore_index=True)
        else:
            results[label] = pd.DataFrame(columns=OUTPUT_COLS)

    # ── Etape 4 : export ──
    print("[i] Export CSV...")
    out_paths = derive_output_paths(output_file)
    out_paths["NUMBER"].parent.mkdir(parents=True, exist_ok=True)
    saved = {}
    for label, df_out in results.items():
        if df_out.empty:
            all_warnings.append(f"Dataset {label} vide — pas de fichier cree")
            continue
        save_path = out_paths[label]
        df_out.to_csv(str(save_path), sep=";", index=False, encoding="utf-8-sig")
        saved[label] = (str(save_path), len(df_out))

    # ── Resume ──
    print(f"[i] Agregation terminee — {n} fichier(s) traite(s)")
    print(f"[i] Schema : {' | '.join(OUTPUT_COLS)}")
    for label, (path, nb) in saved.items():
        print(f"    {label:<12} : {nb:>9,} lignes  ->  {os.path.basename(path)}")
    if all_warnings:
        print(f"[!] Avertissements ({len(all_warnings)}) :")
        for w in all_warnings[:20]:
            print(f"    - {w}")
        if len(all_warnings) > 20:
            print(f"    ... et {len(all_warnings)-20} autres")

    if not saved:
        print("[ERREUR] Aucun dataset non vide — aucun fichier produit.")
        return 1

    return 0


# =============================================================================
# MAIN (CLI argparse)
# =============================================================================

def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="01.N3096",
        description=f"CCO FLUX AGREGATEUR v4 [{VERSION_ID}] — "
                    "agregation CCO_Flux (CSV/XLSX) vers 3 CSV normalises "
                    "NUMBER / AMOUNT / INTERCHANGE.",
    )
    parser.add_argument(
        "--input-folder", required=True, type=Path,
        help="Dossier source contenant les CSV/XLSX a agreger (obligatoire).",
    )
    parser.add_argument(
        "--output-file", required=True, type=Path,
        help="Chemin de BASE de sortie ; les 3 fichiers _NUMBER/_AMOUNT/"
             "_INTERCHANGE.csv en sont derives (obligatoire).",
    )
    parser.add_argument(
        "--existing-folder", required=False, type=Path, default=None,
        help="Dossier des CSV consolides existants a completer "
             "(active le mode incremental ; absent = mode complet).",
    )

    try:
        args = parser.parse_args(argv)
    except SystemExit:
        # argparse a deja signale l'erreur d'arguments
        return 2

    # Validation des chemins
    if not args.input_folder.is_dir():
        print(f"[ERREUR] Dossier source introuvable : {args.input_folder}")
        return 2
    if args.existing_folder is not None and not args.existing_folder.is_dir():
        print(f"[ERREUR] Dossier des CSV existants introuvable : {args.existing_folder}")
        return 2

    try:
        rc = run_aggregation(args.input_folder, args.output_file, args.existing_folder)
    except Exception as e:
        print(f"[ERREUR] {e}")
        import traceback
        traceback.print_exc()
        return 1

    if rc == 0:
        print("[OK] Agregation CCO Flux terminee avec succes.")
    return rc


if __name__ == "__main__":
    sys.exit(main())
