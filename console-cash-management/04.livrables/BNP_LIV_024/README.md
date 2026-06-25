# BNP_LIV_024 — ANALYSE_REVENUS_ET_COMMISSIONS_CLIENTS_CIB_HISTORIQUE

| | |
|---|---|
| Script | `01.Q8YY0.py` |
| Code | Q8YY0 |
| Statut | Refactoré CLI (argparse, sans GUI) — Session 2 |

Analyse commissionnement CIB : 8 sources obligatoires + 7 optionnelles, recalcul complet par année
(plafond, périodicité, différé), RWA, sortie XLSX avec formules vivantes.

Paramètres et étapes : docstring en tête du script (ARGUMENTS CLI / DECOMPOSITION), aussi dans la Console.

> Note : certains fichiers optionnels (BPE_RETAIL, SEG_AGENCE, USAGE, MC1, MC2) sont exposés en flags
> mais non consommés par le worker (parité UI) — voir HANDOFF.md.
