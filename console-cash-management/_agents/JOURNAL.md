# JOURNAL D'AVANCEMENT

## Session 1 — Initialisation (Phases 1 & 2)
- Date : 2026-06-22
- Agent : Claude Code (Opus 4.8)

### Phases complétées
- **Phase 1** : arborescence créée + fichiers `_agents/` (STARTER, ARCHITECTURE, CONVENTIONS,
  JOURNAL) + README racine.
- **Phase 2.A** : pull Supabase — 15 livrables / 25 scripts Python écrits **verbatim** dans
  `local/livrables/`, fidélité vérifiée par md5 (0 écart) + 1 `README.md` par livrable (15).
- **Phase 2.B** : `local/lib/fonctions.py` créé (vide, à enrichir en Phase 3).
- **Phase 2.C** : application `bnp/` (server.js, index.html, config.json, README.md), zéro npm.

### Décisions prises
- Schéma Supabase réel ≠ hypothèses du prompt (cf. addendum d'ARCHITECTURE.md). Pas de
  `is_current`, pas de colonne `code`, pas de `name`. Filtre retenu :
  `fichier_type IN ('python','python_gui') AND content_text IS NOT NULL`.
- **Noms de fichiers réels** conservés (`fichier_shortname`), pas de renommage `01.{CODE}.py`
  (décision Ali). Code alphanumérique consigné dans les README (best-effort).
- **17 livrables conceptuels** (dossiers/documents sans code) **exclus** (décision Ali).
- Pull via script Node temporaire + API REST Supabase (clé anon) pour une copie byte-perfect ;
  script supprimé après usage.
- `bnp/server.js` durci : streaming NDJSON, anti-traversal, gestion EADDRINUSE, surcharge PORT,
  API WHATWG URL (remplacement de `url.parse()` déprécié).

### Vérifications
- 25 `.py` + 15 `README.md` + 15 dossiers `BNP_LIV_*` (comptes exacts).
- Smoke-test serveur OK (port 3999) : /api/livrables → 15 livrables / 25 scripts, / → UI
  (titre OK), /api/browse → 200, route inconnue → 404. Fallback `../local/livrables` OK.
- Environnement PC ARTEONYS : Node v24.16.0, Python 3.14.6, Git 2.54.0.

### En attente / prochaine session
- ⛔ **Validation d'Ali** : tester `cd bnp ; node server.js` → `localhost:3000`, vérifier l'UI
  et la cohérence du pull, puis feu vert explicite pour la **Phase 3**.
- Phase 3 (1 prompt par script) : retirer Tkinter/customtkinter, argparse, print(), inliner les
  fonctions de `lib/`. Cas spécial `BNP_LIV_000_CONSOLE` (imports `_shared`).
- À clarifier avec Ali : codes fichier vs description divergents — LIV_034 (W3RKN / H4WQZ) et
  LIV_035 (K6BZP / NO7WK).
- Port 3000 occupé sur PC ARTEONYS par une autre app (SlideForge) — sans impact sur PC BNP.

## Session 2 — Refonte (autonome)
- Date : 2026-06-23
- Agent : Claude Code (Opus 4.8)
- Plan d'exécution : `PLAN_SESSION_2.md`. Transmission : `HANDOFF.md`.

### Phases complétées (A→E)
- **A** — Renommage des 23 scripts en `0N.{CODE}.py` (code `[XXXXX]` du docstring, source de
  vérité). Audit dans `AUDIT_PHASE_2.md` (9 corrections de code vs S1). `_shared.py` (lib) et
  `a_console.py` (lanceur legacy) conservés.
- **B** — Catalogue `local/catalog/` (sources 24, colonnes 39, fonctions 11, agrégateurs 4) +
  `sources/{originals,work}/` (miroir vide) + `lib/fonctions.py` enrichi + `build_bnp_catalog.py`.
- **C** — Refonte `bnp/` : `server.js` (routes catalog/file/script-tree/run/aggregator),
  `index.html` thème **clair** 5 onglets + dataviz SVG, `_explore.py` (pandas), `catalog.json`
  (généré), `config.json` étendu, `README.md`. Smoke-test complet OK (pandas 3.0.3 présent).
- **D** — **23/23 scripts refactorés** en CLI argparse (sous-agents chirurgicaux). Tous compilent
  (`py_compile`), zéro GUI résiduel (hors `a_console.py`). Logique métier préservée à l'identique.
- **E** — `ARCHITECTURE.md` + `CONVENTIONS.md` réécrits, 15 READMEs régénérés, `HANDOFF.md` créé.

### Décisions (argumentées)
- Renommage `0N.{CODE}.py` : **renverse** la décision S1 (« noms réels ») car le prompt S2 l'impose
  et le code est fiable (docstring). Discordances S1 (LIV_034/035) résolues : W3RKN/K6BZP font foi.
- Charte : passage **sombre → clair** (blanc + vert BNP `#009A44`).
- `BNP_LIV_000_CONSOLE` : 7 outils refactorés ; `_shared.py` gardé (lib) ; `a_console.py` retiré du
  CLI (remplacé par le GUI web).
- Refactoring par sous-agents parallèles (1/script), consignes chirurgicales + vérif `py_compile`.

### Incident & remédiation
- **Limite de session** atteinte pendant le batch 3 (9 gros analyseurs) : 5 ont fini (mort au
  reporting), 2 non faits (Q8YY0, K6BZP). Diagnostic par `py_compile` + grep marqueurs GUI →
  Q8YY0/K6BZP restaurés byte-perfect depuis Supabase (md5) puis **re-refactorés** après reset.
  Résultat final : 23/23.
- **Formats docstring hétérogènes** (selon les sous-agents) : parser `server.js` rendu robuste
  (titre tolérant aux séparateurs, marqueurs `(oblig)`/`(obligatoire)`, puces, `:`), + repli
  collecte de `--flags` et **fallback texte** pour les DECOMPOSITION en arbre ASCII. Vérifié sur
  les 23 : code OK ×23, args OK ×23, décomposition affichable ×23.

### Vérifications finales
- 25 `.py` compilent ; 23 sans GUI ; 6 JSON valides ; smoke-test serveur OK (toutes routes) ;
  `__pycache__` et scripts temporaires supprimés.

### En attente / points d'attention (voir HANDOFF.md)
- Scripts refactorés **non exécutés** (pas de données) → Ali teste sur PC BNP.
- Mapping de colonnes : reproduit les **défauts** des ex-GUI ; certains réglages interactifs ne sont
  plus paramétrables en CLI (signalés par livrable). Cas notable : LIV_017 (colonnes PNB Worldline
  vides par défaut).
- Catalogue : champs `a_remplir` à compléter par Ali.

## Session 3 — Refonte du front `bnp/` et intégration du catalogue racine
- Date : 2026-06-24
- Agent : Claude Code (Opus 4.8)
- Objectif : intégrer le nouveau `column.json` racine dans la structure catalogue existante et
  refondre le front `bnp/index.html` autour de 5 pages : Sources / Données / Exécution /
  Exploration / Configuration.

### Plan d'action
1. Déplacer `column.json` dans `local/catalog/column.json` pour en faire un catalogue interne
   lisible par `local/lib/build_bnp_catalog.py`.
2. Étendre `build_bnp_catalog.py` pour exposer les colonnes par source (`source_columns`) et
   fusionner intelligemment les `livrables_utilisateurs` issus des deux catalogues.
3. Ajouter des endpoints Node natifs dans `bnp/server.js` pour lire/sauver/créer des scripts et
   gérer la configuration persistante de chemins dans `bnp/config.json`.
4. Refaire `bnp/index.html` en un seul monolithe clair, dense et professionnel :
   - Page Sources : référence source × colonne.
   - Page Données : dictionnaire de colonnes dédupliqué.
   - Page Exécution : sélection de livrable, version de script, visualisation + édition + création.
   - Page Exploration : preview CSV/XLSX, filtres et dataviz.
   - Page Configuration : racines livrables/sources, et mapping source → chemin réel.
5. Moderniser `_explore.py` pour préférer DuckDB lors de la lecture CSV/XLSX, avec fallback
   pandas si DuckDB absent.

### Décision catalogue
- Le `column.json` racine est un format de documentation élargi, pas un simple doublon du
  catalogue `local/catalog/columns.json`.
- Je l’intègre comme source de vérité documentaire dans `local/catalog/column.json` et
  l’expose dans `bnp/catalog.json` via le générateur UI.
- Les définitions `nom_canonique`, `livrables_utilisateurs` et `colonnes` y sont désormais
  accessibles au front, sans perte des sources existantes de `local/catalog/sources.json`.

## Session 4 — Catalogue à 3 fichiers liés + câblage direct (autonome, validé par Ali)
- Date : 2026-06-25
- Agent : Claude Code (Opus 4.8)
- Contexte : QA du travail de la session 3 + refonte du modèle catalogue selon la vision d'Ali.

### Diagnostic session 3 (régression trouvée)
- Le build lisait les `sources` de `column.json` (où les 24 sources historiques ne sont que des
  noms sans colonnes) → `bnp/catalog.json` tombé de **24 à 9 sources**. La mention « sans perte »
  du JOURNAL S3 était fausse (vérifié). Nommage incohérent sur 3 niveaux (scripts / sources.json /
  column.json), sans clé de jointure.

### Nouveau modèle (décisions validées par Ali)
- Dossier racine **`configuration/`** = 3 JSON maîtres liés, **source de vérité unique** :
  `data.json` (109 données), `sources.json` (36 sources, `nom_canonique`+`alias`+`chemin_local`+
  `colonnes[].data_ref`), `livrables.json` (15 livrables, `inputs[].source_ref`). Jointure :
  livrables → source_ref → sources → data_ref → data. Intégrité vérifiée à 100 %.
- `bnp/` **lit les 3 JSON directement** (pas d'étape de build). `server.js` : `/api/catalog` renvoie
  les 3 objets ; nouveaux endpoints `/api/source/path` (enregistre `chemin_local`) et
  `/api/source/validate` (compare colonnes réelles vs attendues). Agrégateurs lus depuis
  `local/catalog/aggregators.json`.
- `index.html` refondu : page **Sources** (chemin réel par source + indicateur ✓/⚠ colonnes,
  non bloquant) ; page **Données** (dictionnaire data.json). Onglets : Sources / Données /
  Exécution / Exploration / Configuration.
- Réconciliation de nommage validée par Ali : IBAN_ACCOUNT = IBAN_GROUP (fusionnés) ;
  IBAN_SINGLETON distinct ; GA_GESTION_DIRECTE (yannick = source-sales) ; MC1/MC2 identiques ;
  K2P8N = LIV_036 ; OVERRIDE_PAYEE_PAYER/DEBIT_REVENUE_MAPPING/OVERRIDE_COUNTRY = data clean ;
  FEF/SCOPE abandonnés. Restent `a_confirmer` : CM360, FICHIER_ACHETEUR, CLEAN_SALES,
  TYPE_ACTIONNAIRE, COMMISSIONED_PROCESSING.

### Retraits
- Supprimés : `column.json` (racine + copie catalog/), `local/catalog/sources.json`,
  `local/catalog/columns.json`, `local/lib/build_bnp_catalog.py`, `bnp/catalog.json`.
- Conservés : `local/catalog/functions.json`, `aggregators.json`, `lib/fonctions.py`.
- **ARCHITECTURE.md non mis à jour** (règle : seul JOURNAL.md est éditable) → la doc faisant foi du
  nouveau modèle est `configuration/README.md`.

### Vérifs
- `node --check server.js` OK ; serveur démarré (port 3999) ; `/api/catalog` (109/36/15),
  `/api/aggregators` (4), `/api/source/path` + `/api/source/validate` testés sur CSV → `ok:true`.
- Repo poussé sur GitHub `adjellalil/transfer_ax` (dossier `console-cash-management`).
