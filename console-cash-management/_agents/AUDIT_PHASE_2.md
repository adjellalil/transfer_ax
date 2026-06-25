# AUDIT PHASE A (Session 2) — Codes alphanumériques & renommage

Source de vérité : le code `[XXXXX]` du **docstring d'en-tête** (ligne titre) de chaque script.
« AVANT » = code déduit best-effort en Session 1 (nom de fichier / description Supabase).
Renommage appliqué : `0N.{CODE}.py`.

| Livrable | Ancien fichier | Code AVANT (S1) | Code APRÈS (docstring) | Nouveau fichier | Statut |
|---|---|---|---|---|---|
| BNP_LIV_000_CONSOLE | _shared.py | (aucun) | (aucun — lib) | _shared.py | inchangé (bibliothèque) |
| BNP_LIV_000_CONSOLE | a_console.py | ACNSL | ACNSL | a_console.py | inchangé (*legacy* — lanceur remplacé par le GUI web) |
| BNP_LIV_000_CONSOLE | b_txt_to_csv.py | BTXCV | BTXCV | 01.BTXCV.py | ok |
| BNP_LIV_000_CONSOLE | c_xlsx_to_csv.py | CXLCV | CXLCV | 02.CXLCV.py | ok |
| BNP_LIV_000_CONSOLE | d_merge_files.py | DMRGE | DMRGE | 03.DMRGE.py | ok |
| BNP_LIV_000_CONSOLE | e_clean_files.py | ECLNF | ECLNF | 04.ECLNF.py | ok |
| BNP_LIV_000_CONSOLE | f_replace_chars.py | FRPCH | FRPCH | 05.FRPCH.py | ok |
| BNP_LIV_000_CONSOLE | g_extract_metadata.py | GXMET | GXMET | 06.GXMET.py | ok |
| BNP_LIV_000_CONSOLE | h_renamer.py | HRNAM | HRNAM | 07.HRNAM.py | ok |
| BNP_LIV_004 | cco_flux_agregateur.py | N3096 | N3096 | 01.N3096.py | ok |
| BNP_LIV_004 | cco_stock_agregateur.py | (aucun) | 66Z3Q | 02.66Z3Q.py | **corrigé** |
| BNP_LIV_004 | monext_fusion_pecc52.py | (PECC-5212) | N7M4P | 03.N7M4P.py | **corrigé** |
| BNP_LIV_017 | sales_pnb_analyzer.py | (aucun) | M3X9R | 01.M3X9R.py | **corrigé** |
| BNP_LIV_018 | worldline_analyzer.py | GA14B | V2MRG | 01.V2MRG.py | **corrigé** (GA14B = prédécesseur) |
| BNP_LIV_019 | monext_analyzer.py | (aucun, X5DET) | S8XPL | 01.S8XPL.py | **corrigé** |
| BNP_LIV_024 | cib_commissionnement_analyzer.py | Q8YY0 | Q8YY0 | 01.Q8YY0.py | ok |
| BNP_LIV_026 | cco_pme_dashboard.py | C2SME (pattern) | B2PME | 01.B2PME.py | **corrigé** |
| BNP_LIV_027 | monext_comparaison.py | (aucun, K2P8N) | FSUB3 | 01.FSUB3.py | **corrigé** |
| BNP_LIV_029 | sme_dataset_analyzer.py | (aucun, B6SME) | C2SME | 01.C2SME.py | **corrigé** |
| BNP_LIV_030 | sme_dataset_analyzer.py | C7MWZ | C7MWZ | 01.C7MWZ.py | ok |
| BNP_LIV_032 | kelly_animation_analyzer_TXVLV.py | TXVLV | TXVLV | 01.TXVLV.py | ok |
| BNP_LIV_033 | rc_identifier_analyzer_Q3JKL.py | Q3JKL | Q3JKL | 01.Q3JKL.py | ok |
| BNP_LIV_034 | cib_revenus_analyzer_W3RKN.py | W3RKN / [H4WQZ] | W3RKN | 01.W3RKN.py | ok (discordance résolue) |
| BNP_LIV_035 | cib_commission_analyzer_K6BZP.py | K6BZP / [NO7WK] | K6BZP | 01.K6BZP.py | ok (discordance résolue) |
| BNP_LIV_036 | monext_revenus_periode_M5VTQ.py | M5VTQ | M5VTQ | 01.M5VTQ.py | ok |

## Synthèse
- **23 fichiers renommés** en `0N.{CODE}.py` ; `_shared.py` et `a_console.py` conservés tels quels.
- **9 corrections de code** vs Session 1 (best-effort erroné ou absent) : LIV_004 (×2), 017, 018,
  019, 026, 027, 029.
- Aucune anomalie : tous les scripts portent un code `[XXXXX]` en ligne 3, sauf `_shared.py`
  (bibliothèque sans version).
- Les READMEs de livrables seront régénérés en fin de session (code exact + nouveau nom + statut
  refactoré), afin de refléter l'état final post-Phase D en une seule passe.
