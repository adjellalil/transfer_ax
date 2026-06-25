# BNP_LIV_035 — ANALYSE_COMMISSIONS_VERSEES_CLIENTS_CIB

| | |
|---|---|
| Script | `01.K6BZP.py` |
| Code | K6BZP |
| Statut | Refactoré CLI (argparse, sans GUI) — Session 2 |

Analyse des commissions versées aux clients CIB (flux CCO/Monext + CPC/Worldline) : 8 sources
obligatoires + 3 optionnelles, recalcul par année, REBATE cloisonné, XLSX 7 onglets, formules vivantes.

Paramètres et étapes : docstring en tête du script (ARGUMENTS CLI / DECOMPOSITION), aussi dans la Console.

> Note : pays exclus et plage de mois reprennent les défauts de l'ex-UI (non paramétrables en CLI) — voir HANDOFF.md.
