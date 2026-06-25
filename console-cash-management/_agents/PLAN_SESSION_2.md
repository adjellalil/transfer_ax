# PLAN D'EXÉCUTION — SESSION 2 (refonte)

Plan détaillé suivi pendant la session autonome du 2026-06-23. Décisions argumentées dans
[JOURNAL.md](JOURNAL.md) ; transmission à Ali dans [HANDOFF.md](HANDOFF.md).

## Décisions structurantes
1. Renommage `0N.{CODE}.py` (code = `[XXXXX]` du docstring, ligne titre). Renverse la décision
   Session 1 (« noms réels ») car le prompt Phase A est explicite et le code est fiable.
2. Charte : thème **clair** « shadcn sans shadcn » — blanc dominant, vert BNP `#009A44`.
3. `BNP_LIV_000_CONSOLE` : `_shared.py` conservé (lib) ; `a_console.py` retiré du CLI (lanceur
   remplacé par le GUI web) ; `b..h` refactorés. → **23 scripts refactorés**.
4. Refactoring par sous-agents chirurgicaux (logique préservée), vérif `py_compile`.
5. Zéro écriture Supabase, zéro exécution de script (pas de données).

## Table de renommage (Phase A) — codes vérifiés ligne 3

| Livrable | Ancien nom | Nouveau nom |
|---|---|---|
| BNP_LIV_000_CONSOLE | _shared.py | _shared.py (inchangé, lib) |
| BNP_LIV_000_CONSOLE | a_console.py | a_console.py (inchangé, *legacy*) |
| BNP_LIV_000_CONSOLE | b_txt_to_csv.py | 01.BTXCV.py |
| BNP_LIV_000_CONSOLE | c_xlsx_to_csv.py | 02.CXLCV.py |
| BNP_LIV_000_CONSOLE | d_merge_files.py | 03.DMRGE.py |
| BNP_LIV_000_CONSOLE | e_clean_files.py | 04.ECLNF.py |
| BNP_LIV_000_CONSOLE | f_replace_chars.py | 05.FRPCH.py |
| BNP_LIV_000_CONSOLE | g_extract_metadata.py | 06.GXMET.py |
| BNP_LIV_000_CONSOLE | h_renamer.py | 07.HRNAM.py |
| BNP_LIV_004 | cco_flux_agregateur.py | 01.N3096.py |
| BNP_LIV_004 | cco_stock_agregateur.py | 02.66Z3Q.py |
| BNP_LIV_004 | monext_fusion_pecc52.py | 03.N7M4P.py |
| BNP_LIV_017 | sales_pnb_analyzer.py | 01.M3X9R.py |
| BNP_LIV_018 | worldline_analyzer.py | 01.V2MRG.py |
| BNP_LIV_019 | monext_analyzer.py | 01.S8XPL.py |
| BNP_LIV_024 | cib_commissionnement_analyzer.py | 01.Q8YY0.py |
| BNP_LIV_026 | cco_pme_dashboard.py | 01.B2PME.py |
| BNP_LIV_027 | monext_comparaison.py | 01.FSUB3.py |
| BNP_LIV_029 | sme_dataset_analyzer.py | 01.C2SME.py |
| BNP_LIV_030 | sme_dataset_analyzer.py | 01.C7MWZ.py |
| BNP_LIV_032 | kelly_animation_analyzer_TXVLV.py | 01.TXVLV.py |
| BNP_LIV_033 | rc_identifier_analyzer_Q3JKL.py | 01.Q3JKL.py |
| BNP_LIV_034 | cib_revenus_analyzer_W3RKN.py | 01.W3RKN.py |
| BNP_LIV_035 | cib_commission_analyzer_K6BZP.py | 01.K6BZP.py |
| BNP_LIV_036 | monext_revenus_periode_M5VTQ.py | 01.M5VTQ.py |

## Ordre d'exécution
1. **PLAN_SESSION_2.md** (ce fichier).
2. **Phase A** : renommer (23 fichiers), MAJ des 15 READMEs (code exact, retrait « best-effort »),
   `AUDIT_PHASE_2.md`.
3. **Phase B** : `local/catalog/` (4 JSON + README), `local/sources/{originals,work}/` (miroir vide
   + README), enrichir `lib/fonctions.py` (depuis `_shared.py` + helpers récurrents),
   `lib/build_bnp_catalog.py`.
4. **Phase C** : `bnp/_explore.py`, `config.json` (étendu), `build_bnp_catalog.py` → `catalog.json`,
   `server.js` (routes étendues), `index.html` (thème clair, 5 onglets, SVG natif), `README.md`.
5. **Phase D** : refactoring des 23 scripts (sous-agents), format docstring + argparse, `py_compile`.
6. **Phase E** : `ARCHITECTURE.md`, `CONVENTIONS.md`, `JOURNAL.md`, `HANDOFF.md`.
7. **Vérif** : `json.tool` (JSON), `py_compile` (scripts), smoke-test serveur (port libre).

## Sources/colonnes/fonctions amorçables (depuis exploration)
- Sources (~23) : PRGM, MONEXT, CCO_FLUX, CCO_STOCK, PARC_CLIENT, REFERENTIEL_CLIENT, IDSEG,
  OPTIFLUX, IBAN_ACCOUNT, DEVISES, REBATE, CM360, MONITORING, MATCHING_USAGE, MC1, MC2,
  BPE_RETAIL, SEG_AGENCE, OVERRIDE_PAYS, BEJO_CARTES, BEJO_FLUX, SINGLETON, COUNTRY, YANNICK.
- Fonctions communes (~26) : load_csv_smart, load_xlsx, load_file_auto, save_csv_protected,
  protect_long_ids, clean_id, clean_ga, clean_iban*, norm_rs, to_float, parse_mois, pays_to_geo,
  get_column_preview, write_xlsx, sacralise_vec, etc.
