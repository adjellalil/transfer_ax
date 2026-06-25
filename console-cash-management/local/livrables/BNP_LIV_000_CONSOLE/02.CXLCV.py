"""
XLSX -> CSV CONVERTER [CXLCV]
============================================================================

DESCRIPTION :
    Conversion en masse de fichiers Excel (.xlsx, .xls) vers CSV. Preserve
    les zeros initiaux (lecture en dtype=str) et protege optionnellement les
    identifiants longs au format ="..." pour que Excel ne les transforme pas
    en notation scientifique. Sortie CSV utf-8-sig, separateur ; (compatible
    Excel francais).

SOURCES REQUISES :
    - Un ou plusieurs fichiers Excel (.xlsx ou .xls) passes via --inputs.

OUTPUTS PRODUITS :
    - Un fichier CSV par fichier Excel converti (meme nom de base, extension
      .csv), ecrit dans --output-dir si fourni, sinon a cote du fichier source.

ARGUMENTS CLI :
    --inputs PATH [PATH ...] (obligatoire) liste des fichiers Excel a convertir
    --output-dir DIR (optionnel) dossier de destination des CSV (defaut: a cote du source)
    --no-protect-ids (optionnel) desactive la protection des identifiants longs (=\"...\")

DECOMPOSITION :
    1. Lecture des arguments CLI (argparse)
    2. Conversion des fichiers
        2.1 Chargement du fichier Excel en str (load_xlsx)
        2.2 Determination du dossier de sortie
        2.3 Sauvegarde CSV avec protection optionnelle des IDs (save_csv_protected)
    3. Rapport final et code de sortie (0 OK / 1 erreur(s) / 2 usage)

BNP Paribas Cash Management - Direction Monetique
Mai 2026
"""

import argparse
import csv
import sys
from pathlib import Path

import pandas as pd

VERSION_ID = "CXLCV"


# ============================================================================
# FONCTIONS UTILITAIRES (inlinees depuis _shared.py) - LOGIQUE A L'IDENTIQUE
# ============================================================================

def load_xlsx(path, sheet_name=0):
    """Charge un fichier Excel en str (préserve les zéros initiaux)."""
    return pd.read_excel(
        path, sheet_name=sheet_name, dtype=str,
        keep_default_na=False, na_values=[]
    )


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


# ============================================================================
# LOGIQUE DE CONVERSION (anciennement classe GUI, attributs lus depuis argparse)
# ============================================================================

class XlsxToCsvConverter:
    def __init__(self, args: argparse.Namespace) -> None:
        self.files_to_convert: list[str] = [str(p) for p in args.inputs]
        self.dest_folder: str | None = (
            str(args.output_dir) if args.output_dir is not None else None
        )
        self.protect_ids: bool = args.protect_ids

    def convert_files(self) -> int:
        total = len(self.files_to_convert)
        converted = 0
        errors: list[tuple[str, str]] = []
        for i, path in enumerate(self.files_to_convert):
            try:
                fname = Path(path).name
                print(f"[{i + 1}/{total}] Conversion : {fname}")
                df = load_xlsx(path)
                out_dir = self.dest_folder or str(Path(path).parent)
                out_path = str(Path(out_dir) / (Path(fname).stem + ".csv"))
                save_csv_protected(df, out_path, protect_ids=self.protect_ids)
                converted += 1
                print(f"[OK] {fname} -> {out_path}")
            except Exception as e:
                errors.append((Path(path).name, str(e)))
                print(f"[ERREUR] {Path(path).name}: {str(e)[:200]}")

        print(f"[OK] Termine : {converted}/{total} fichier(s) converti(s)")
        if errors:
            print(f"[ERREUR] {len(errors)} erreur(s) :")
            for fn, err in errors[:5]:
                print(f"  - {fn}: {err[:80]}")
            return 1
        return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="02.CXLCV.py",
        description=f"Conversion en masse XLSX/XLS -> CSV [{VERSION_ID}]",
    )
    parser.add_argument(
        "--inputs", nargs="+", required=True, type=Path,
        help="Liste des fichiers Excel (.xlsx, .xls) a convertir",
    )
    parser.add_argument(
        "--output-dir", type=Path, default=None,
        help="Dossier de destination des CSV (defaut: a cote du fichier source)",
    )
    parser.add_argument(
        "--no-protect-ids", dest="protect_ids", action="store_false",
        help='Desactive la protection des identifiants longs (format ="...")',
    )
    parser.set_defaults(protect_ids=True)
    return parser


def main() -> int:
    parser = build_parser()
    try:
        args = parser.parse_args()
    except SystemExit:
        return 2

    try:
        if not args.inputs:
            print("[ERREUR] Aucun fichier selectionne")
            return 2
        if args.output_dir is not None and not Path(args.output_dir).is_dir():
            print(f"[ERREUR] Dossier de destination introuvable : {args.output_dir}")
            return 2

        converter = XlsxToCsvConverter(args)
        return converter.convert_files()
    except Exception as e:
        print(f"[ERREUR] {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
