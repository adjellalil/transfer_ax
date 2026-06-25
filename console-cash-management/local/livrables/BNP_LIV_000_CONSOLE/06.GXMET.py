"""
EXTRACTION DE METADONNEES [GXMET]
==================================

TITRE [GXMET]
    EXTRACTION DE METADONNEES [GXMET]

DESCRIPTION
    Outil CLI autonome d'extraction de metadonnees (noms de colonnes +
    exemples de valeurs) depuis plusieurs fichiers CSV / XLSX / TXT.
    Pour chaque fichier analyse, produit une ligne par colonne avec sa
    position, son nom et un nombre configurable d'exemples de valeurs.
    La logique metier est strictement identique a la version GUI d'origine
    (chargement auto du fichier, extraction des exemples, format de sortie).

SOURCES REQUISES
    - Un dossier complet (--input-folder) contenant des fichiers
      .csv / .xlsx / .xls / .txt
      OU
    - Une liste explicite de fichiers (--inputs), un ou plusieurs chemins.
    L'un OU l'autre doit etre fourni (exclusif).

OUTPUTS PRODUITS
    - 1 fichier CSV consolide (--output-file), separateur ';',
      encodage utf-8-sig, 1 ligne par colonne :
      Fichier / Position / Nom_Colonne / Exemple_1 .. Exemple_N

ARGUMENTS CLI
    --input-folder   Dossier source (tous les CSV/XLSX/XLS/TXT du dossier).
    --inputs         Liste de fichiers source (nargs='+').
                     --input-folder ET --inputs sont mutuellement exclusifs ;
                     l'un des deux est requis.
    --output-file    Chemin du CSV consolide a produire (requis).
    --examples       Nombre d'exemples par colonne, 1 a 5 (defaut 3,
                     valeur par defaut de l'UI d'origine).

DECOMPOSITION
    1. Resolution de la source (dossier ou liste de fichiers).
    2. Pour chaque fichier : chargement auto (load_file_auto, nrows=100).
    3. Pour chaque colonne : extraction de N exemples (logique identique).
    4. Consolidation dans un DataFrame et ecriture du CSV utf-8-sig.

BNP Paribas Cash Management - Direction Monetique
Mai 2026
"""

import argparse
import os
import sys
from pathlib import Path
from typing import List, Optional

import pandas as pd

VERSION_ID = "GXMET"


# ─── HELPERS INLINE (depuis _shared.py, VERBATIM) ──────────────────────

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


# ─── RESOLUTION DES SOURCES ────────────────────────────────────────────

def resolve_files(input_folder: Optional[str], inputs: Optional[List[str]]) -> List[str]:
    """Resout la liste des fichiers source (dossier ou liste explicite).

    Reproduit la logique de l'UI :
    - dossier : tous les fichiers .csv/.xlsx/.xls/.txt du dossier
    - liste   : les chemins fournis tels quels
    """
    if input_folder:
        folder = input_folder
        if not os.path.isdir(folder):
            raise FileNotFoundError(f"Dossier introuvable : {folder}")
        files = [
            os.path.join(folder, f) for f in os.listdir(folder)
            if f.lower().endswith(('.csv', '.xlsx', '.xls', '.txt'))
        ]
        return files

    files = list(inputs or [])
    manquants = [f for f in files if not os.path.isfile(f)]
    if manquants:
        raise FileNotFoundError(
            "Fichier(s) introuvable(s) : " + ", ".join(manquants)
        )
    return files


# ─── EXTRACTION (logique metier identique) ─────────────────────────────

def extract_metadata(files_in_folder: List[str], n_ex: int) -> pd.DataFrame:
    """Extrait les metadonnees (1 ligne par colonne) pour chaque fichier.

    Logique strictement identique a la methode extract() de l'UI.
    """
    metadata = []

    for i, fpath in enumerate(files_in_folder):
        try:
            print(f"Analyse {i + 1}/{len(files_in_folder)} : {os.path.basename(fpath)}")
            df = load_file_auto(fpath, nrows=100)
            for ci, col in enumerate(df.columns):
                examples = df[col].dropna().head(n_ex).astype(str).tolist()
                while len(examples) < n_ex:
                    examples.append("")
                row = {
                    'Fichier': os.path.basename(fpath),
                    'Position': ci + 1,
                    'Nom_Colonne': col
                }
                for j, ex in enumerate(examples, 1):
                    row[f'Exemple_{j}'] = str(ex)[:100]
                metadata.append(row)
        except Exception as e:
            metadata.append({
                'Fichier': os.path.basename(fpath),
                'Position': 0,
                'Nom_Colonne': f"ERREUR: {str(e)[:50]}"
            })

    return pd.DataFrame(metadata)


# ─── MAIN / CLI ────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="06.GXMET.py",
        description=f"EXTRACTION DE METADONNEES [{VERSION_ID}] - "
                    "extraction des noms de colonnes et exemples de valeurs "
                    "depuis plusieurs fichiers CSV/XLSX/TXT.",
    )
    src = parser.add_mutually_exclusive_group(required=True)
    src.add_argument(
        "--input-folder",
        help="Dossier source (tous les .csv/.xlsx/.xls/.txt du dossier).",
    )
    src.add_argument(
        "--inputs",
        nargs='+',
        help="Liste explicite de fichiers source.",
    )
    parser.add_argument(
        "--output-file",
        required=True,
        help="Chemin du CSV consolide a produire.",
    )
    parser.add_argument(
        "--examples",
        type=int,
        default=3,
        choices=range(1, 6),
        metavar="{1..5}",
        help="Nombre d'exemples par colonne (1 a 5, defaut 3).",
    )
    return parser


def main(argv: Optional[List[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    try:
        files_in_folder = resolve_files(args.input_folder, args.inputs)

        if not files_in_folder:
            print("Aucun fichier", file=sys.stderr)
            return 2

        print(f"{len(files_in_folder)} fichier(s) a analyser")
        for i, f in enumerate(files_in_folder, 1):
            print(f"{i:03d}. {os.path.basename(f)}")

        n_ex = int(args.examples)
        mdf = extract_metadata(files_in_folder, n_ex)

        save_path = args.output_file
        out_parent = Path(save_path).parent
        if str(out_parent) and not out_parent.exists():
            out_parent.mkdir(parents=True, exist_ok=True)

        mdf.to_csv(save_path, index=False, sep=';', encoding='utf-8-sig')

        print("Termine")
        print(
            f"Extraction terminee : "
            f"{len(files_in_folder)} fichiers -> {len(mdf)} colonnes"
        )
        print(f"Sortie : {save_path}")
        return 0

    except (FileNotFoundError, OSError, ValueError) as e:
        print(f"ERREUR : {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"ERREUR INATTENDUE : {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
