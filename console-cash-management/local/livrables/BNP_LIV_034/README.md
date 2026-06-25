# BNP_LIV_034 — ANALYSE_REVENUS_CLIENTS_CIB

| | |
|---|---|
| Script | `01.W3RKN.py` |
| Code | W3RKN |
| Statut | Refactoré CLI (argparse, sans GUI) — Session 2 |

Analyse des revenus clients CIB (flux CCO/MONEXT + CPC/Worldline) : cascade de matching,
récupération étendue (PARC/SINGLETON), conversion devises WL→EUR, snapshots, XLSX nombres natifs
(TCD-ready). 8 sources obligatoires + optionnelles.

Paramètres et étapes : docstring en tête du script (ARGUMENTS CLI / DECOMPOSITION), aussi dans la Console.
