"""
TITRE [DMRGE]
=============
FUSION DE FICHIERS [DMRGE] - version CLI autonome (sans GUI).

DESCRIPTION
-----------
Outil de fusion de plusieurs fichiers CSV en un seul, avec alignement de
colonnes sur le premier fichier (reference) et transformations prealables
optionnelles (renommer / ajouter / supprimer / remplacer des colonnes,
par fichier ou sur TOUS les fichiers).

La logique metier (chargement multi-separateurs, transformations,
verification de compatibilite, fusion alignee sur les colonnes du premier
fichier, sauvegarde CSV utf-8-sig avec protection des IDs longs) est
PRESERVEE A L'IDENTIQUE par rapport a la version GUI customtkinter
d'origine. Les fonctions du module `_shared` utilisees (`load_csv_smart`,
`save_csv_protected` et l'aide `protect_long_ids`) ont ete INLINEES pour
rendre ce script AUTONOME (zero dependance a tkinter / _shared).

GESTION DES TRANSFORMATIONS INTERACTIVES
----------------------------------------
La version GUI permettait de saisir interactivement des transformations
(renommer / ajouter / supprimer / remplacer des colonnes, par fichier ou
sur TOUS). En CLI ces saisies interactives sont remplacees par un fichier
JSON optionnel `--transforms-json`. Si ce fichier n'est PAS fourni, le
script effectue une FUSION SIMPLE ALIGNEE sur les colonnes du premier
fichier (comportement par defaut de la fusion d'origine), sans aucune
transformation. Format du JSON (toutes les cles sont optionnelles) :

    {
      "per_file": {
        "<nom_de_fichier.csv>": [
          {"action": "rename",  "column": "ANCIEN", "new_name": "NOUVEAU"},
          {"action": "add",     "name": "SOURCE",   "value": "BNP"},
          {"action": "delete",  "column": "INUTILE"},
          {"action": "replace", "column": "MNT", "search": ",", "replace": "."}
        ]
      },
      "all_files": [
        {"action": "add_all",    "name": "LOT",  "value": "2026"},
        {"action": "rename_all", "old": "ID",     "new_name": "IDENTIFIANT"}
      ]
    }

La cle `per_file` est indexee par le nom de base du fichier (basename), tel
qu'affiche par la version GUI. Les transformations `all_files` sont
appliquees apres les transformations `per_file`, dans l'ordre du tableau.

SOURCES REQUISES
----------------
- --inputs        : 2 fichiers CSV ou plus a fusionner (au moins 2).
- --transforms-json (optionnel) : fichier JSON de transformations.

OUTPUTS PRODUITS
----------------
- --output-file   : fichier CSV fusionne (sep ';', utf-8-sig, IDs proteges).

ARGUMENTS CLI
-------------
- --inputs          (nargs='+', requis) : chemins des CSV a fusionner.
- --output-file     (requis)            : chemin du CSV de sortie.
- --transforms-json (optionnel)         : chemin du JSON de transformations.

DECOMPOSITION
-------------
03.DMRGE.py
|
+-- Helpers metier inlines depuis _shared (logique INCHANGEE)
|   +-- load_csv_smart(path, nrows)        : chargement CSV multi-sep/enc
|   +-- protect_long_ids(series, threshold): protection IDs longs ="..."
|   +-- save_csv_protected(df, path, ...)  : ecriture CSV utf-8-sig protege
|
+-- Logique de transformation (issue des actions GUI, INCHANGEE)
|   +-- apply_transforms(dataframes, transforms) : applique per_file + all_files
|       +-- rename / add / delete / replace      (par fichier)
|       +-- add_all / rename_all                 (sur TOUS)
|
+-- Logique de verification / fusion (INCHANGEE)
|   +-- check_compatibility(dataframes)    : colonnes manquantes / en trop
|   +-- merge_dataframes(dataframes)       : fusion alignee sur 1er fichier
|
+-- Orchestration CLI
    +-- load_inputs(paths)                 : chargement + noms uniques
    +-- main()                             : argparse -> try/except -> 0/1/2
"""

import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List

import pandas as pd

VERSION_ID = "DMRGE"


# ─── HELPERS METIER INLINES DEPUIS _shared (logique INCHANGEE) ──────────

def load_csv_smart(path: str, nrows=None) -> pd.DataFrame:
    """Charge un CSV en testant plusieurs separateurs et encodages."""
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


def protect_long_ids(series: pd.Series, threshold: int = 15) -> pd.Series:
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


def save_csv_protected(df: pd.DataFrame, path: str, protect_ids: bool = True,
                       id_threshold: int = 15) -> None:
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


# ─── CHARGEMENT DES ENTREES (noms uniques, logique du GUI add_files) ────

def load_inputs(paths: List[str]) -> Dict[str, pd.DataFrame]:
    """Charge les CSV et leur attribue un nom de base unique (comme le GUI)."""
    dataframes: Dict[str, pd.DataFrame] = {}
    for path in paths:
        p = Path(path)
        df = load_csv_smart(str(p))
        name = p.name
        counter = 1
        base = name
        while name in dataframes:
            name = f"{base} ({counter})"
            counter += 1
        dataframes[name] = df
        print(f"  Charge : {name} ({len(df):,} lignes x {len(df.columns)} colonnes)")
    return dataframes


# ─── TRANSFORMATIONS (issues des actions GUI, logique INCHANGEE) ────────

def apply_transforms(dataframes: Dict[str, pd.DataFrame], transforms: dict) -> None:
    """Applique les transformations per_file puis all_files (in place).

    Reproduit a l'identique les actes des boutons GUI :
      - rename / add / delete / replace : par fichier (combo_transform)
      - add_all / rename_all            : sur TOUS les fichiers
    """
    per_file = transforms.get("per_file", {}) or {}
    for fname, actions in per_file.items():
        if fname not in dataframes:
            print(f"  AVERTISSEMENT : fichier '{fname}' introuvable pour transformation, ignore")
            continue
        for act in (actions or []):
            action = act.get("action")
            if action == "rename":
                c = act["column"]
                new = act["new_name"]
                # GUI: act_rename -> rename(columns={c: new.strip()})
                dataframes[fname].rename(columns={c: new.strip()}, inplace=True)
                print(f"  [{fname}] renomme '{c}' -> '{new.strip()}'")
            elif action == "add":
                name = act["name"]
                val = act.get("value", "") or ""
                # GUI: act_add -> df[name.strip()] = val
                dataframes[fname][name.strip()] = val
                print(f"  [{fname}] ajoute colonne '{name.strip()}' = '{val}'")
            elif action == "delete":
                c = act["column"]
                # GUI: act_delete -> drop(columns=[c])
                dataframes[fname].drop(columns=[c], inplace=True)
                print(f"  [{fname}] supprime colonne '{c}'")
            elif action == "replace":
                c = act["column"]
                search = act["search"]
                repl = act.get("replace", "") or ""
                # GUI: act_replace -> astype(str).str.replace(search, repl, regex=False)
                dataframes[fname][c] = dataframes[fname][c].astype(str).str.replace(
                    search, repl, regex=False
                )
                print(f"  [{fname}] remplace '{search}' -> '{repl}' dans '{c}'")
            else:
                print(f"  AVERTISSEMENT : action per_file inconnue '{action}', ignoree")

    all_files = transforms.get("all_files", []) or []
    for act in all_files:
        action = act.get("action")
        if action == "add_all":
            name = act["name"]
            val = act.get("value", "") or ""
            # GUI: act_add_all -> for df: df[name.strip()] = val
            for df in dataframes.values():
                df[name.strip()] = val
            print(f"  [TOUS] ajoute colonne '{name.strip()}' = '{val}'")
        elif action == "rename_all":
            old = act["old"]
            new = act["new_name"]
            # GUI: act_rename_all -> for df: if old in cols: rename({old: new.strip()})
            for df in dataframes.values():
                if old in df.columns:
                    df.rename(columns={old: new.strip()}, inplace=True)
            print(f"  [TOUS] renomme '{old}' -> '{new.strip()}'")
        else:
            print(f"  AVERTISSEMENT : action all_files inconnue '{action}', ignoree")


# ─── VERIFICATION DE COMPATIBILITE (logique GUI check_compat INCHANGEE) ─

def check_compatibility(dataframes: Dict[str, pd.DataFrame]) -> None:
    """Affiche le rapport de compatibilite (manquantes / en trop)."""
    files = list(dataframes.keys())
    ref_cols = list(dataframes[files[0]].columns)
    print(f"Reference : {files[0]} ({len(ref_cols)} colonnes)\n")
    for fname in files[1:]:
        cols = list(dataframes[fname].columns)
        missing = [c for c in ref_cols if c not in cols]
        extra = [c for c in cols if c not in ref_cols]
        if not missing and not extra:
            print(f"OK : {fname}")
        else:
            print(f"DIFF : {fname}")
            if missing:
                print(f"  Manquantes : {', '.join(missing)}")
            if extra:
                print(f"  En trop : {', '.join(extra)}")


# ─── FUSION (logique GUI execute_merge INCHANGEE) ───────────────────────

def merge_dataframes(dataframes: Dict[str, pd.DataFrame]) -> pd.DataFrame:
    """Fusionne les DataFrames alignes sur les colonnes du premier fichier."""
    files = list(dataframes.keys())
    ref_cols = list(dataframes[files[0]].columns)
    merged = pd.DataFrame()
    for fname in files:
        df = dataframes[fname].copy()
        for col in ref_cols:
            if col not in df.columns:
                df[col] = ""
        df = df[ref_cols]
        merged = df if merged.empty else pd.concat([merged, df], ignore_index=True)
    return merged


# ─── ORCHESTRATION CLI ──────────────────────────────────────────────────

def main() -> int:
    parser = argparse.ArgumentParser(
        prog="03.DMRGE.py",
        description=f"FUSION DE FICHIERS [{VERSION_ID}] - CLI autonome (sans GUI)",
    )
    parser.add_argument(
        "--inputs", nargs="+", required=True,
        help="Chemins des fichiers CSV a fusionner (au moins 2).",
    )
    parser.add_argument(
        "--output-file", required=True,
        help="Chemin du fichier CSV fusionne en sortie.",
    )
    parser.add_argument(
        "--transforms-json", default=None,
        help="(Optionnel) Fichier JSON de transformations. Si absent : "
             "fusion simple alignee sur le 1er fichier.",
    )
    args = parser.parse_args()

    try:
        print(f"=== FUSION DE FICHIERS [{VERSION_ID}] ===")
        print(f"Demarrage : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

        # Validation des entrees
        if len(args.inputs) < 2:
            print("ERREUR : fournir au moins 2 fichiers via --inputs.", file=sys.stderr)
            return 2

        missing_inputs = [p for p in args.inputs if not Path(p).is_file()]
        if missing_inputs:
            for p in missing_inputs:
                print(f"ERREUR : fichier d'entree introuvable : {p}", file=sys.stderr)
            return 2

        transforms = None
        if args.transforms_json is not None:
            tj = Path(args.transforms_json)
            if not tj.is_file():
                print(f"ERREUR : fichier --transforms-json introuvable : {tj}", file=sys.stderr)
                return 2
            transforms = json.loads(tj.read_text(encoding="utf-8"))

        # Chargement
        print("Chargement des fichiers :")
        dataframes = load_inputs(args.inputs)
        print()

        if len(dataframes) < 2:
            print("ERREUR : moins de 2 fichiers charges, fusion impossible.", file=sys.stderr)
            return 1

        # Transformations (optionnelles)
        if transforms is not None:
            print("Application des transformations :")
            apply_transforms(dataframes, transforms)
            print()
        else:
            print("Aucune transformation fournie : fusion simple alignee sur le 1er fichier.\n")

        # Verification de compatibilite
        print("Verification de compatibilite :")
        check_compatibility(dataframes)
        print()

        # Fusion
        print("Fusion en cours...")
        merged = merge_dataframes(dataframes)

        # Sauvegarde
        out_path = Path(args.output_file)
        if out_path.parent and not out_path.parent.exists():
            out_path.parent.mkdir(parents=True, exist_ok=True)
        save_csv_protected(merged, str(out_path))

        print(f"Fusion reussie : {len(dataframes)} fichiers -> {len(merged):,} lignes")
        print(f"Sortie : {out_path}")
        print(f"\nTermine : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        return 0

    except KeyError as e:
        print(f"ERREUR : cle de transformation/colonne manquante : {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"ERREUR : {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
