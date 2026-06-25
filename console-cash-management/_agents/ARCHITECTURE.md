# ARCHITECTURE — Cash Management Console

Plateforme = **environnement de travail complet** pour la mission BNP (Direction Monétique) :
stocker les scripts livrables, décrire les données (catalogue), lancer/explorer depuis un GUI
web unifié qui remplace les anciennes GUI Tkinter.

## Les deux espaces

| Espace | Vit où | Rôle | Transféré ? |
|---|---|---|---|
| `local/` | PC ARTEONYS | Atelier : scripts, lib, catalogue, miroir des sources | Non |
| `bnp/` | ARTEONYS → PC BNP | Application web minimaliste (zéro npm) | Oui (par mail) |

## Arborescence

```
local/
├── livrables/<BNP_LIV_xxx>/
│   ├── 0N.{CODE}.py        ← script CLI argparse autonome (refactoré, sans Tkinter)
│   └── README.md           ← métadonnées du livrable
│   (BNP_LIV_000_CONSOLE : _shared.py = lib, a_console.py = lanceur legacy conservé)
├── lib/
│   ├── fonctions.py        ← bibliothèque de RÉFÉRENCE (inlinée dans les scripts, jamais importée)
│   └── build_bnp_catalog.py ← génère bnp/catalog.json depuis catalog/
├── catalog/
│   ├── sources.json        ← fichiers sources (description, format, chemins, colonnes)
│   ├── columns.json        ← colonnes source + calculées (depends_on, produced_by, formule)
│   ├── functions.json      ← fonctions communes
│   ├── aggregators.json    ← programmes d'agrégation
│   └── README.md
└── sources/                ← miroir (structure vide, AUCUNE donnée)
    ├── originals/          ← XLSX bruts (déposés par Ali sur PC BNP)
    └── work/               ← CSV consolidés (produits par les agrégateurs)

bnp/                        ← 6 fichiers, zéro dépendance npm
├── server.js               ← serveur Node natif (API + service UI)
├── index.html              ← UI monolithique (HTML+CSS+JS inline), thème clair, 5 onglets
├── _explore.py             ← utilitaire pandas (preview/filter/stats) appelé par le serveur
├── catalog.json            ← miroir du catalogue (généré)
├── config.json
└── README.md
```

## Le catalogue (cœur de l'écosystème)
Carte vivante des données. Claude Code le LIT avant de coder un livrable ; Ali l'enrichit (champs
`a_remplir`). Le miroir `bnp/catalog.json` (généré par `build_bnp_catalog.py`) alimente les onglets
Sources / Catalogue / Agrégateurs de l'UI. Voir `catalog/README.md`.

## Format des scripts refactorés
Chaque `0N.{CODE}.py` est un script CLI **autonome** (argparse, aucune GUI, fonctions communes
inlinées). Son docstring d'en-tête suit un format imposé (voir `CONVENTIONS.md`) avec, notamment :
- ligne titre `NOM [CODE]` (le code `[XXXXX]` fait foi, il vient du docstring d'origine) ;
- section `ARGUMENTS CLI` → **lue par la Console** pour générer le formulaire de paramètres ;
- section `DECOMPOSITION` (arborescence) → **lue par la Console** pour l'arbre de traitement.
Codes de sortie : `0` succès, `1` erreur fonctionnelle, `2` erreur technique. Logs `print()` FR.

## Le GUI web (bnp/)
`server.js` expose : `/api/livrables`, `/api/catalog`, `/api/aggregators`, `/api/script/tree`
(parse le docstring), `/api/script/run` (args structurés → flags + streaming NDJSON),
`/api/file/{preview,filter,stats}` (subprocess `_explore.py`), `/api/aggregator/run`, `/api/browse`.
5 onglets : Livrables, Sources, Catalogue, Explorer (dataviz SVG), Agrégateurs.

## Workflow de mise à jour
1. Nouvelle source → `catalog/sources.json` (+ colonnes dans `columns.json`).
2. `python local/lib/build_bnp_catalog.py` → régénère `bnp/catalog.json`.
3. Envoi par mail au PC BNP : le plus souvent **1 seul fichier** (`index.html` ou `catalog.json`).

## Règles d'or
1. `bnp/` : **aucune dépendance npm**, modules Node natifs uniquement, nombre de fichiers minimal.
2. `local/lib/fonctions.py` est une **référence** : on **inline** les fonctions dans chaque script
   (autonomie sur PC BNP), on ne l'importe jamais.
3. **Aucun fichier de données réel** sur ARTEONYS (ni `.xlsx`/`.csv` rempli). Que du code/structure.
4. `_agents/` : tu ne modifies que `JOURNAL.md` (le reste est référence).
5. **Jamais de push GitHub.** Commits locaux optionnels.
