# Cash Management Console

Environnement unifié des livrables de la mission **BNP Paribas Cash Management — Direction
Monétique** (Ali Djellalil / Arteonys). Une console web locale (Node.js natif, **zéro npm**) pilote
l'exécution des scripts ; les sources de données sont résolues automatiquement par l'arborescence.

## Arborescence (dossiers numérotés)

```
00.console/        App web (server.js + index.html). Lancer : cd 00.console ; node server.js -> http://localhost:3000
01.agents/         Documentation interne / journal de bord (référence ; seul JOURNAL.md évolue)
03.sources/        Dépôt des fichiers sources, 1 sous-dossier par source (voir 03.sources/README.md)
   01.interne/        sources maintenues à la main (overrides, data-clean, matching, devises…)
   02.plateforme/     données plateformes monétiques (Worldline PRGM, Monext, CCO flux/stock, acheteur)
   03.client/         référentiels clients BNP (parc, référentiel client, IBAN, segment, agences, GA)
04.livrables/      Scripts de livrables (BNP_LIV_xxx / NN.CODE.py), CLI argparse autonomes
05.configuration/  Catalogue source de vérité : data.json · sources.json · livrables.json (+ aggregators.json)
06.fonctions/      fonctions.py — bibliothèque de référence (résolution sources + lecture DuckDB)
07.outputs/        Sorties produites par les scripts (non versionnées)
.venv/             Environnement Python local (non versionné ; recréé sur le PC BNP)
```

## Le principe « zéro chemin en dur »
Un script de `04.livrables/<LIV>/` **remonte de 2 dossiers**, va dans `03.sources/`, trouve le
sous-dossier au **nom canonique** de la source (peu importe la catégorie) et prend le fichier au
**numéro de préfixe le plus élevé**. Donc : où que tu poses le repo sur le PC BNP, il suffit de
déposer les fichiers dans les bons sous-dossiers de `03.sources/` (préfixés `01.`, `02.`…) et tout
fonctionne. On peut aussi forcer un fichier précis via l'argument CLI correspondant.

## Pré-requis (PC BNP)
- **Node.js** (sert la console, aucun `npm install`).
- **Python 3** : `pip install -r requirements.txt` (pandas, openpyxl, duckdb).
  *DuckDB est optionnel : sans lui, les scripts retombent automatiquement sur pandas.*

## Démarrage
1. Déposer les fichiers sources dans `03.sources/<catégorie>/<SOURCE>/` (préfixe numérique croissant).
2. `cd 00.console ; node server.js` → http://localhost:3000.
3. Onglet **Sources** : pour chaque source, « Vérifier colonnes » contrôle le fichier détecté.
4. Onglet **Exécution** : lancer un livrable (les sources sont résolues automatiquement).

## Catalogue
3 fichiers liés dans `05.configuration/` (voir son README) : `data.json` (dictionnaire des données),
`sources.json` (sources : nom canonique, alias, catégorie, colonnes→données, chemin), `livrables.json`
(livrables : arguments CLI → source canonique). Lus directement par la console (pas d'étape de build).
