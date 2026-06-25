# catalog/ — Référentiel de données

Le catalogue est la **carte vivante** de l'écosystème de données : il décrit les fichiers
sources, leurs colonnes, les colonnes calculées, les fonctions communes et les agrégateurs.
Claude Code le lit **avant de coder** un nouveau livrable ; Ali l'enrichit à la main.

## Fichiers

| Fichier | Contenu |
|---|---|
| `sources.json` | Fichiers sources : description, type (agrégé/référentiel), format, chemins, agrégateur, colonnes principales, livrables utilisateurs. |
| `columns.json` | `columns_source` (colonnes brutes) + `columns_calculees` (avec `depends_on`, `produced_by`, `formula_description`). |
| `functions.json` | Fonctions communes (signature, fichier de référence, livrables). Version canonique dans `../lib/fonctions.py`. |
| `aggregators.json` | Programmes d'agrégation XLSX/CSV mensuels -> CSV consolidé (statut `actif` / `a_creer`). |

## Champs `a_remplir`
Les valeurs `"a_remplir": true` ou `"a_remplir"` marquent ce que l'analyse automatique n'a pas pu
déduire (définitions métier exactes, formats, exemples réels). **À compléter par Ali.**

## Workflow d'enrichissement (Ali)
1. **Nouvelle source** -> ajouter une entrée dans `sources.json` (+ déposer le fichier dans
   `../sources/originals/<SOURCE>/`).
2. **Nouvelles colonnes** -> les décrire dans `columns.json` (source ou calculée).
3. **Nouvel agrégateur** -> entrée dans `aggregators.json`.
4. Régénérer le miroir UI : `python ../lib/build_bnp_catalog.py` -> met à jour `../../bnp/catalog.json`.
5. (Optionnel) envoyer le `bnp/catalog.json` mis à jour par mail au PC BNP.

## Workflow de lecture (Claude Code)
Avant d'écrire/refactorer un script : lire `sources.json` (entrées attendues), `columns.json`
(colonnes manipulées) et `functions.json` (helpers à inliner depuis `../lib/fonctions.py`).
