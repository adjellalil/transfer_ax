"""
TXT -> CSV CONVERTER [BTXCV]
============================
DESCRIPTION
-----------
Conversion de fichiers TXT (notamment exports Teradata) vers CSV.
- Detection automatique du separateur (virgule, point-virgule, tab, pipe)
- Detection automatique des guillemets (doubles, simples, aucun)
- Strip des guillemets entourant les valeurs
- Protection des identifiants longs au format ="..." pour Excel

SOURCES REQUISES
----------------
- Un ou plusieurs fichiers TXT (txt) -- exports a convertir (ex: exports Teradata).

OUTPUTS PRODUITS
----------------
- Un fichier CSV par fichier TXT en entree (separateur ';', encodage utf-8-sig,
  toutes valeurs entre guillemets), nomme <base>.csv dans le dossier de sortie.

ARGUMENTS CLI
-------------
--inputs PATH...           (obligatoire) Un ou plusieurs fichiers TXT a convertir
--output-dir PATH          (obligatoire) Dossier de sortie
--sep VALUE                (optionnel) Separateur : auto | , | ; | tab | | (defaut: auto)
--quote VALUE              (optionnel) Guillemets : auto | " | ' | none (defaut: auto)
--no-protect-ids           (optionnel) Desactive la protection des identifiants longs (="...")
--no-strip                 (optionnel) Conserve les guillemets entourant les valeurs

DECOMPOSITION
-------------
1. Lecture des arguments CLI
   1.1 Resolution des fichiers d'entree et du dossier de sortie
   1.2 Normalisation des options separateur / guillemets / protection / strip
2. Conversion de chaque fichier
   2.1 Lecture du fichier TXT (detection encodage/separateur/guillemets)
   2.2 Strip optionnel des guillemets
   2.3 Sauvegarde CSV protege (separateur ';', utf-8-sig, QUOTE_ALL)
3. Rapport final (compteur converti / total, erreurs eventuelles)

BNP Paribas Cash Management - Direction Monetique
Mai 2026
"""

import argparse
import csv
import os
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd

VERSION_ID = "BTXCV"


# ─── FONCTIONS UTILITAIRES (inlinees depuis _shared.py) ─────────────────

def protect_long_ids(series, threshold=15):
    """Protege les identifiants longs en format ="VALEUR" pour Excel."""
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


# ─── LOGIQUE DE CONVERSION ──────────────────────────────────────────────

class TxtToCsvConverter:
    def __init__(self, args: argparse.Namespace):
        self.files_to_convert: List[str] = [str(p) for p in args.inputs]
        self.output_dir: Optional[str] = str(args.output_dir) if args.output_dir else None
        self.sep_value: str = args.sep
        self.quote_value: str = args.quote
        self.protect_ids: bool = not args.no_protect_ids
        self.strip_quotes: bool = not args.no_strip

    @staticmethod
    def detect_separator(lines):
        candidates = [',', ';', '\t', '|']
        best_sep, best_score = ',', -1
        for sep in candidates:
            counts = [l.count(sep) for l in lines if l.strip()]
            if not counts:
                continue
            if min(counts) > 0 and max(counts) == min(counts):
                if min(counts) > best_score:
                    best_score = min(counts); best_sep = sep
            elif min(counts) > 0 and min(counts) > best_score:
                best_score = min(counts); best_sep = sep
        return best_sep

    @staticmethod
    def detect_quotechar(lines):
        sample = ''.join(lines[:5])
        if sample.count('"') > 10: return '"'
        if sample.count("'") > 10: return "'"
        return None

    def read_txt_file(self, path, nrows=None):
        raw_lines = []
        for enc in ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']:
            try:
                with open(path, 'r', encoding=enc) as f:
                    raw_lines = f.readlines()
                break
            except Exception:
                continue
        if not raw_lines:
            with open(path, 'r', errors='replace') as f:
                raw_lines = f.readlines()
        if not raw_lines:
            return pd.DataFrame(), ',', None, raw_lines

        sep_choice = self.sep_value
        sep = self.detect_separator(raw_lines[:20]) if sep_choice == "auto" else sep_choice

        quote_choice = self.quote_value
        if quote_choice == "auto":
            quotechar = self.detect_quotechar(raw_lines[:10])
        elif quote_choice == "none":
            quotechar = None
        else:
            quotechar = quote_choice

        for enc in ['utf-8', 'latin1', 'cp1252']:
            try:
                kwargs = {
                    'sep': sep, 'encoding': enc, 'dtype': str,
                    'keep_default_na': False, 'na_values': [],
                    'on_bad_lines': 'skip',
                }
                if quotechar:
                    kwargs['quotechar'] = quotechar
                if nrows:
                    kwargs['nrows'] = nrows
                df = pd.read_csv(path, **kwargs)
                if df.shape[1] > 1:
                    if self.strip_quotes:
                        for col in df.columns:
                            df[col] = df[col].astype(str).str.strip('"').str.strip("'")
                        df.columns = [c.strip('"').strip("'").strip() for c in df.columns]
                    return df, sep, quotechar, raw_lines
            except Exception:
                continue
        return pd.DataFrame(), sep, quotechar, raw_lines

    def convert_files(self, dest_folder: Optional[str]) -> Tuple[int, int, List[Tuple[str, str]]]:
        total = len(self.files_to_convert)
        converted = 0
        errors: List[Tuple[str, str]] = []
        for i, path in enumerate(self.files_to_convert):
            try:
                fname = os.path.basename(path)
                print(f"[{i + 1}/{total}] Conversion : {fname}")
                df, _, _, _ = self.read_txt_file(path)
                base = os.path.splitext(fname)[0]
                out_path = os.path.join(dest_folder or os.path.dirname(path), f"{base}.csv")
                save_csv_protected(df, out_path, protect_ids=self.protect_ids)
                converted += 1
            except Exception as e:
                errors.append((os.path.basename(path), str(e)))
        return converted, total, errors


# ─── INTERFACE CLI ──────────────────────────────────────────────────────

def _normalize_sep(value: str) -> str:
    """Convertit la valeur CLI du separateur vers le format attendu par la logique."""
    if value == "tab":
        return "\t"
    return value


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=f"Conversion TXT -> CSV [{VERSION_ID}]"
    )
    parser.add_argument(
        "--inputs", nargs='+', required=True, type=Path,
        help="Un ou plusieurs fichiers TXT a convertir"
    )
    parser.add_argument(
        "--output-dir", required=True, type=Path,
        help="Dossier de sortie"
    )
    parser.add_argument(
        "--sep", default="auto", choices=["auto", ",", ";", "tab", "|"],
        help="Separateur (defaut: auto)"
    )
    parser.add_argument(
        "--quote", default="auto", choices=["auto", '"', "'", "none"],
        help="Guillemets (defaut: auto)"
    )
    parser.add_argument(
        "--no-protect-ids", action="store_true",
        help='Desactive la protection des identifiants longs (="...")'
    )
    parser.add_argument(
        "--no-strip", action="store_true",
        help="Conserve les guillemets entourant les valeurs"
    )
    args = parser.parse_args(argv)
    args.sep = _normalize_sep(args.sep)
    return args


def main(argv: Optional[List[str]] = None) -> int:
    try:
        args = parse_args(argv)

        print(f"[1/3] Lecture des arguments [{VERSION_ID}]")
        missing = [str(p) for p in args.inputs if not Path(p).is_file()]
        if missing:
            raise FileNotFoundError(
                "Fichier(s) introuvable(s) : " + ", ".join(missing)
            )

        output_dir = Path(args.output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)

        converter = TxtToCsvConverter(args)

        print(f"[2/3] Conversion de {len(converter.files_to_convert)} fichier(s)")
        converted, total, errors = converter.convert_files(str(output_dir))

        print(f"[3/3] Rapport final")
        if errors:
            print(f"{len(errors)} erreur(s) de conversion :")
            for fn, err in errors[:5]:
                print(f"  - {fn}: {err[:80]}")

        print(f"[OK] Conversion terminee : {converted}/{total} fichier(s) converti(s)")
        return 0

    except (FileNotFoundError, PermissionError, ValueError) as e:
        print(f"[ERREUR] Erreur fonctionnelle : {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"[ERREUR] Erreur technique : {e}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
