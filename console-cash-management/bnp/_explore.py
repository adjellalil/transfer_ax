"""
_explore.py — Utilitaire d'exploration de fichiers (CSV/XLSX) pour la Console.

Appelé par server.js via subprocess. Renvoie du JSON sur stdout.
Actions : preview (aperçu N lignes), filter (filtres + aperçu), stats (stats d'une colonne).
Autonome : aucune dépendance hors pandas/openpyxl. Messages d'erreur en français.

Usage :
  python _explore.py --action preview --path FICHIER --limit 100
  python _explore.py --action filter  --path FICHIER --limit 100 --filters '[{"col":"RS","op":"contains","value":"ACME"}]'
  python _explore.py --action stats   --path FICHIER --column MONTANT
"""
from __future__ import annotations

import argparse
import json
import os
import sys

import pandas as pd


def load(path: str) -> pd.DataFrame:
    """Charge un CSV (détection séparateur/encodage) ou un XLSX, tout en str."""
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path, dtype=str, keep_default_na=False, na_values=[])
    for sep in [";", ",", "\t"]:
        for enc in ["utf-8", "latin1", "cp1252"]:
            try:
                test = pd.read_csv(path, sep=sep, encoding=enc, dtype=str,
                                   keep_default_na=False, na_values=[], on_bad_lines="skip", nrows=5)
                if test.shape[1] > 1:
                    return pd.read_csv(path, sep=sep, encoding=enc, dtype=str,
                                       keep_default_na=False, na_values=[], on_bad_lines="skip")
            except Exception:
                continue
    return pd.read_csv(path, sep=None, engine="python", dtype=str,
                       keep_default_na=False, na_values=[], on_bad_lines="skip")


def to_num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.str.replace(" ", "", regex=False).str.replace(",", ".", regex=False),
                         errors="coerce")


def col_types(df: pd.DataFrame) -> dict:
    types = {}
    for c in df.columns:
        nums = to_num(df[c].head(200))
        ratio = nums.notna().mean() if len(nums) else 0
        types[c] = "number" if ratio >= 0.8 else "string"
    return types


def apply_filters(df: pd.DataFrame, filters: list) -> pd.DataFrame:
    for f in filters:
        col, op, val = f.get("col"), f.get("op"), f.get("value", "")
        if col not in df.columns:
            continue
        s = df[col]
        if op == "=":
            df = df[s == str(val)]
        elif op == "!=":
            df = df[s != str(val)]
        elif op == "contains":
            df = df[s.str.contains(str(val), case=False, na=False)]
        elif op in (">", "<", ">=", "<="):
            n, ref = to_num(s), float(str(val).replace(",", "."))
            df = df[{">": n > ref, "<": n < ref, ">=": n >= ref, "<=": n <= ref}[op]]
    return df


def preview(df: pd.DataFrame, limit: int) -> dict:
    head = df.head(limit)
    return {
        "columns": list(df.columns),
        "rows": head.astype(str).values.tolist(),
        "total_rows": int(len(df)),
        "column_types": col_types(df),
    }


def stats(df: pd.DataFrame, column: str) -> dict:
    if column not in df.columns:
        raise ValueError(f"Colonne introuvable : {column}")
    s = df[column]
    nums = to_num(s)
    is_num = nums.notna().mean() >= 0.8
    out = {"count": int(len(s)), "nulls": int((s.str.strip() == "").sum()),
           "unique": int(s.nunique())}
    if is_num:
        v = nums.dropna()
        out.update({"type": "numeric", "sum": float(v.sum()), "mean": float(v.mean()) if len(v) else 0.0,
                    "min": float(v.min()) if len(v) else 0.0, "max": float(v.max()) if len(v) else 0.0})
        edges = list(pd.cut(v, bins=10).value_counts().sort_index().items()) if len(v) else []
        out["histogram"] = [{"label": f"{iv.left:.0f}–{iv.right:.0f}", "count": int(n)} for iv, n in edges]
    else:
        top = s[s.str.strip() != ""].value_counts().head(10)
        out.update({"type": "categorical", "top_values": [[str(k), int(n)] for k, n in top.items()]})
    return out


def main() -> None:
    p = argparse.ArgumentParser(description="Exploration de fichiers CSV/XLSX")
    p.add_argument("--action", required=True, choices=["preview", "filter", "stats"])
    p.add_argument("--path", required=True)
    p.add_argument("--limit", type=int, default=100)
    p.add_argument("--column", default=None)
    p.add_argument("--filters", default="[]")
    args = p.parse_args()

    try:
        if not os.path.exists(args.path):
            raise FileNotFoundError(f"Fichier introuvable : {args.path}")
        df = load(args.path)
        if args.action == "preview":
            result = preview(df, args.limit)
        elif args.action == "filter":
            result = preview(apply_filters(df, json.loads(args.filters)), args.limit)
        else:
            result = stats(df, args.column)
        print(json.dumps(result, ensure_ascii=False))
        sys.exit(0)
    except FileNotFoundError as e:
        print(json.dumps({"error": str(e)}, ensure_ascii=False)); sys.exit(1)
    except Exception as e:  # noqa: BLE001
        print(json.dumps({"error": f"{type(e).__name__}: {e}"}, ensure_ascii=False)); sys.exit(2)


if __name__ == "__main__":
    main()
