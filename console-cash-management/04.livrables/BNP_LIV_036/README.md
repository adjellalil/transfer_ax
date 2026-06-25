# BNP_LIV_036 — BNP_LIV_36_ANALYSE_KPI_CORPO

| | |
|---|---|
| Script | `01.M5VTQ.py` |
| Code | M5VTQ |
| Statut | Refactoré CLI (argparse, sans GUI) — Session 2 |

Analyse des revenus MONEXT par client (ID RP) sur une période sélectionnable (`--mois-debut` →
`--mois-fin`). Entrée unique MONEXT consolidé. Sortie XLSX 2 feuilles (DATA + SYNTHESE) :
FLUX, NB_CARTES_MOYEN, DIFFERES, PNB + décomposition mensuelle.

Paramètres et étapes : docstring en tête du script (ARGUMENTS CLI / DECOMPOSITION), aussi dans la Console.
