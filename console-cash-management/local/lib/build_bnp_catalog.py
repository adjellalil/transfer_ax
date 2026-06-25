"""
build_bnp_catalog.py — Génère le miroir UI bnp/catalog.json depuis local/catalog/.

Lit local/catalog/columns.json, local/catalog/functions.json, local/catalog/aggregators.json
et le catalogue racine column.json. Produit un fichier consolidé et aplati, optimisé pour
l'affichage dans la Console web. Relançable à tout moment après enrichissement du catalogue.

Usage : python build_bnp_catalog.py
"""

from __future__ import annotations

import datetime
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent          # local/
CATALOG_DIR = ROOT / "catalog"
BNP_CATALOG = ROOT.parent / "bnp" / "catalog.json"
ROOT_COLUMN = ROOT.parent / "column.json"
LOCAL_ROOT_COLUMN = ROOT / "catalog" / "column.json"


def _read(name: str) -> dict:
    path = CATALOG_DIR / name
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _read_root() -> dict:
    for path in (LOCAL_ROOT_COLUMN, ROOT_COLUMN):
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    return {}


def build() -> dict:
    columns = _read("columns.json")
    functions = _read("functions.json").get("functions", {})
    aggregators = _read("aggregators.json").get("aggregators", {})
    root_catalog = _read_root()
    root_sources = root_catalog.get("sources", {})
    root_concepts = root_catalog.get("concepts_canoniques_referentiels", {})
    root_livrables = root_catalog.get("livrables_referentiel", {})

    flat_sources = []
    source_columns = []

    def add_source_columns(source_ref: str, source_name: str, columns_list, livrables, description, source_type):
        seen = set()
        if not columns_list:
            return
        for item in columns_list:
            if isinstance(item, dict):
                col_name = item.get("nom_dans_fichier") or item.get("concept_canonique")
                col_desc = item.get("description", "")
                col_format = item.get("format", "")
                concept = item.get("concept_canonique", "")
            else:
                col_name = str(item)
                col_desc = ""
                col_format = ""
                concept = ""
            if not col_name:
                continue
            key = (source_ref, col_name)
            if key in seen:
                continue
            seen.add(key)
            source_columns.append({
                "source_ref": source_ref,
                "source_name": source_name,
                "source_type": source_type,
                "description": description,
                "column_name": col_name,
                "concept_canonique": concept,
                "column_description": col_desc,
                "column_format": col_format,
                "livrables": livrables or [],
            })

    def merge_lists(first, second):
        out = list(dict.fromkeys((first or []) + (second or [])))
        return out

    for ref, s in root_sources.items():
        if not isinstance(s, dict):
            continue
        if not (s.get("colonnes") or s.get("colonnes_principales")):
            continue
        columns_list = s.get("colonnes", s.get("colonnes_principales", []))
        def extract_column_names(items):
            if not items:
                return []
            names = []
            for x in items:
                if isinstance(x, dict):
                    name = x.get("nom_dans_fichier") or x.get("concept_canonique")
                else:
                    name = str(x)
                if name:
                    names.append(name)
            return names

        flat_sources.append({
            "ref": ref,
            "description": s.get("description", ""),
            "type": s.get("type", ""),
            "format_original": s.get("format_recu", s.get("format_original", "")),
            "format_travail": s.get("format_travail", s.get("format_recu", "")),
            "chemin_travail": s.get("chemin_relatif_travail", ""),
            "chemin_original": s.get("chemin_relatif_original", ""),
            "agregateur": s.get("agregateur"),
            "colonnes_principales": s.get("colonnes_principales", []),
            "columns": extract_column_names(columns_list),
            "utilisee_dans": s.get("utilisee_dans", s.get("livrables_utilisateurs", [])),
            "status": s.get("status", ""),
            "nom_canonique": s.get("nom_canonique", ""),
            "nom_historique": s.get("nom_historique", ""),
            "owner": s.get("owner", ""),
            "notes": s.get("notes", ""),
            "plateforme_origine": s.get("plateforme_origine", ""),
            "frequence": s.get("frequence", ""),
            "separateur": s.get("separateur", ""),
        })
        add_source_columns(ref, s.get("nom_canonique", ref), columns_list, s.get("utilisee_dans", s.get("livrables_utilisateurs", [])), s.get("description", ""), s.get("type", ""))

    flat_columns = []
    for name, c in columns.get("columns_source", {}).items():
        flat_columns.append({
            "name": name, "kind": "source",
            "definition": c.get("definition", ""), "format": c.get("format", ""),
            "exemple": c.get("exemple", ""), "sources": c.get("sources_utilisees", []),
        })
    for name, c in columns.get("columns_calculees", {}).items():
        flat_columns.append({
            "name": name, "kind": "calculee",
            "definition": c.get("definition", ""), "format": c.get("format", ""),
            "exemple": c.get("exemple", ""), "depends_on": c.get("depends_on", []),
            "produced_by": c.get("produced_by", []),
            "formula_description": c.get("formula_description", ""),
        })

    flat_aggregators = [
        {"ref": ref, **{k: a.get(k) for k in
                        ("description", "input_folder", "output_file", "script", "status")}}
        for ref, a in aggregators.items()
    ]

    flat_functions = [
        {"name": name, "description": fn.get("description", ""),
         "signature": fn.get("signature", ""), "utilisee_dans": fn.get("utilisee_dans", [])}
        for name, fn in functions.items()
    ]

    output = {
        "version": "1.0",
        "generated_at": datetime.date.today().isoformat(),
        "sources": flat_sources,
        "columns": flat_columns,
        "aggregators": flat_aggregators,
        "functions": flat_functions,
        "source_columns": source_columns,
    }
    if root_concepts:
        output["concepts_canoniques"] = root_concepts
    if root_livrables:
        output["livrables_referentiel"] = root_livrables
    return output


def main() -> None:
    catalog = build()
    BNP_CATALOG.parent.mkdir(parents=True, exist_ok=True)
    with open(BNP_CATALOG, "w", encoding="utf-8") as f:
        json.dump(catalog, f, ensure_ascii=False, indent=2)
    print(f"[OK] {BNP_CATALOG} généré : "
          f"{len(catalog['sources'])} sources, {len(catalog['columns'])} colonnes, "
          f"{len(catalog['aggregators'])} agrégateurs, {len(catalog['functions'])} fonctions.")


if __name__ == "__main__":
    main()
