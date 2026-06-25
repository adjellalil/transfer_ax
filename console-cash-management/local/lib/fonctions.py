"""
Bibliothèque de fonctions de référence — Cash Management Console.

Ce fichier est une RÉFÉRENCE pour la génération/refonte de scripts. Les fonctions
ne sont JAMAIS importées par les livrables : elles sont INLINEÉES dans chaque script
pour garantir l'autonomie sur PC BNP.

Conventions :
- snake_case, type hints, docstring avec exemple d'usage.
- Dépendances autorisées : pandas, openpyxl (présents sur PC BNP).

Source des implémentations : module partagé `_shared.py` (BNP_LIV_000_CONSOLE) pour les
helpers d'I/O, et patterns récurrents extraits des analyseurs (clean_*, norm_rs, to_float,
parse_mois, pays_to_geo).
"""

from __future__ import annotations

import csv
import os
import re
import unicodedata

import pandas as pd

# ─── Palette BNP (sorties Excel) ───────────────────────────────────────────
VERT_BNP = "#00915A"
VERT_FONCE = "#003D2E"
VERT_CLAIR = "#E8F5E9"
GRIS = "#666666"
BLANC = "#FFFFFF"
NOIR = "#333333"
ORANGE = "#FF9800"
ROUGE = "#E53935"


# ═══════════════════════════════════════════════════════════════════════════
# I/O — lecture / écriture (depuis _shared.py)
# ═══════════════════════════════════════════════════════════════════════════
def load_csv_smart(path: str, nrows: int | None = None) -> pd.DataFrame:
    """Charge un CSV en testant plusieurs séparateurs et encodages.

    Tout est lu en str (zéros initiaux préservés). Essaie ; , tab puis
    auto-détection. Exemple : ``df = load_csv_smart('work/MONEXT.csv')``.
    """
    for sep in [";", ",", "\t"]:
        for enc in ["utf-8", "latin1", "cp1252", "iso-8859-1"]:
            try:
                df_test = pd.read_csv(
                    path, sep=sep, encoding=enc, dtype=str,
                    keep_default_na=False, na_values=[],
                    on_bad_lines="skip", nrows=5,
                )
                if df_test.shape[1] > 1:
                    return pd.read_csv(
                        path, sep=sep, encoding=enc, dtype=str,
                        keep_default_na=False, na_values=[],
                        on_bad_lines="skip", nrows=nrows,
                    )
            except Exception:
                continue
    return pd.read_csv(
        path, sep=None, engine="python", dtype=str,
        keep_default_na=False, na_values=[],
        on_bad_lines="skip", nrows=nrows,
    )


def load_xlsx(path: str, sheet_name: int | str = 0) -> pd.DataFrame:
    """Charge un fichier Excel en str (préserve les zéros initiaux)."""
    return pd.read_excel(
        path, sheet_name=sheet_name, dtype=str,
        keep_default_na=False, na_values=[],
    )


def load_file_auto(path: str, nrows: int | None = None) -> pd.DataFrame:
    """Charge un fichier CSV ou XLSX selon son extension."""
    ext = os.path.splitext(path)[1].lower()
    if ext in [".xlsx", ".xls"]:
        return load_xlsx(path)
    return load_csv_smart(path, nrows)


def protect_long_ids(series: pd.Series, threshold: int = 15) -> pd.Series:
    """Protège les identifiants numériques longs au format ``="VALEUR"`` pour Excel."""
    def protect(val: object) -> str:
        s = str(val).strip()
        if not s or s in ["", "nan", "NaN", "None", "N/A"]:
            return s
        clean = s.replace(" ", "").replace(".0", "")
        if clean.isdigit() and len(clean) >= threshold:
            return f'="{clean}"'
        return s
    return series.apply(protect)


def save_csv_protected(df: pd.DataFrame, path: str, protect_ids: bool = True,
                       id_threshold: int = 15) -> None:
    """Sauvegarde un CSV (utf-8-sig, séparateur ;, QUOTE_ALL) avec protection des IDs longs."""
    df_out = df.copy()
    if protect_ids:
        for col in df_out.columns:
            sample = df_out[col].dropna().head(100).astype(str)
            has_long = sample.str.replace(" ", "").str.replace(".0", "").str.match(
                r"^\d{" + str(id_threshold) + r",}$"
            ).any()
            if has_long:
                df_out[col] = protect_long_ids(df_out[col], id_threshold)
    df_out.to_csv(path, index=False, sep=";", encoding="utf-8-sig", quoting=csv.QUOTE_ALL)


def get_column_preview(df: pd.DataFrame, col: str, n: int = 3) -> str:
    """Renvoie une string d'aperçu des n premières valeurs d'une colonne."""
    values = df[col].dropna().head(n).astype(str).tolist()
    return " | ".join([v[:40] for v in values])


# ═══════════════════════════════════════════════════════════════════════════
# Nettoyage / normalisation (patterns récurrents des analyseurs)
# ═══════════════════════════════════════════════════════════════════════════
def clean_id(series: pd.Series) -> pd.Series:
    """Normalise un identifiant numérique : retire espaces, suffixe .0 et zéros initiaux.

    Exemple : ``df['_RC'] = clean_id(df['ID_RC'])`` -> '0012345' devient '12345'.
    """
    def _clean(val: object) -> str:
        s = str(val).strip().replace(" ", "")
        if s.endswith(".0"):
            s = s[:-2]
        s = s.lstrip("0")
        return s
    return series.apply(_clean)


def clean_ga(series: pd.Series) -> pd.Series:
    """Normalise un code GA/agence (retire espaces et préfixe 00 superflu)."""
    def _clean(val: object) -> str:
        s = str(val).strip().replace(" ", "")
        if s.endswith(".0"):
            s = s[:-2]
        return s.lstrip("0")
    return series.apply(_clean)


def norm_rs(series: pd.Series) -> pd.Series:
    """Normalise une Raison Sociale pour le matching (majuscules, sans accents/ponctuation).

    Exemple : ``df['_RS'] = norm_rs(df['RS'])`` -> 'Acmé, S.A.' devient 'ACME SA'.
    """
    def _norm(val: object) -> str:
        s = str(val).strip().upper()
        s = "".join(
            c for c in unicodedata.normalize("NFKD", s)
            if not unicodedata.combining(c)
        )
        s = re.sub(r"[^A-Z0-9 ]", " ", s)
        s = re.sub(r"\s+", " ", s).strip()
        return s
    return series.apply(_norm)


def to_float(series: pd.Series) -> pd.Series:
    """Conversion numérique robuste : gère vides, virgule décimale et valeurs invalides (-> 0.0)."""
    def _f(val: object) -> float:
        s = str(val).strip().replace(" ", "").replace(" ", "")
        if not s or s in ["nan", "NaN", "None", "N/A", "-"]:
            return 0.0
        s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return 0.0
    return series.apply(_f)


def parse_mois(value: str) -> str:
    """Parse un mois ``YYYYMM`` en label ``MM/YYYY`` (renvoie la valeur brute si non parsable).

    Exemple : ``parse_mois('202501')`` -> '01/2025'.
    """
    s = str(value).strip()
    m = re.match(r"^(\d{4})(\d{2})$", s)
    if m:
        return f"{m.group(2)}/{m.group(1)}"
    return s


def pays_to_geo(pays: str) -> str:
    """Mappe un pays vers une région géo : 'France' / 'Hors_France' / 'Pays_non_trouve'."""
    s = str(pays).strip().upper()
    if not s or s in ["", "NAN", "NONE", "N/A"]:
        return "Pays_non_trouve"
    if s in ["FRANCE", "FR", "FRA"]:
        return "France"
    return "Hors_France"
