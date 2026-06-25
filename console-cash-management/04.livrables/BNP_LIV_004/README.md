# BNP_LIV_004 — AGREGATION_FICHIERS_MENSUELS_MONEXT

Agrégateurs de fichiers mensuels, **refactorés en CLI** (Session 2). Exposés dans l'onglet
Agrégateurs de la Console (flags `--input-folder` / `--output-file`).

| Fichier | Code | Rôle | Statut |
|---|---|---|---|
| `01.N3096.py` | N3096 | CCO Flux → 3 CSV (NUMBER/AMOUNT/INTERCHANGE) | Refactoré CLI |
| `02.66Z3Q.py` | 66Z3Q | CCO Stock → 1 CSV consolidé | Refactoré CLI |
| `03.N7M4P.py` | N7M4P | Fusion MONEXT PECC-5212 → CSV multi-année | Refactoré CLI |

Paramètres et étapes : voir le docstring en tête de chaque script.
