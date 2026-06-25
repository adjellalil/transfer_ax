# BNP_LIV_033 — TABLE_MAPPING_ID_RC_ID_RMPM_POUR_WORLDLINE

| | |
|---|---|
| Script | `01.Q3JKL.py` |
| Code | Q3JKL |
| Statut | Refactoré CLI (argparse, sans GUI) — Session 2 |

Résolution des identifiants RC Worldline : pipeline 5 blocs + SYNTHÈSE (validation format,
matching RC→RMPM via PARC et IBAN→RMPM via FORTIS, zero-padding, matching raison sociale, absorbés).
Sortie XLSX 2 feuilles (DATA + ANALYSE waterfall).

Paramètres et étapes : docstring en tête du script (ARGUMENTS CLI / DECOMPOSITION), aussi dans la Console.

> Note : 2 modes (`--mode MODE_1|MODE_2`) ; en MODE_2, `--lookup-input` devient obligatoire.
