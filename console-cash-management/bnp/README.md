# Cash Management Console — Application BNP

GUI web unifié de la mission BNP — Direction Monétique. Remplace les GUI Tkinter par une seule
console locale (Node.js natif, **zéro dépendance npm**) : lancer les livrables, explorer les
fichiers, consulter le catalogue de données.

## Pré-requis sur PC BNP
- **Node.js** (sert l'app, aucun `npm install`).
- **Python 3** avec **pandas** et **openpyxl** (pour lancer les scripts et l'explorateur de fichiers).

## Fichiers de l'app
| Fichier | Rôle |
|---|---|
| `server.js` | Serveur Node natif (API + service de l'UI). |
| `index.html` | UI complète (HTML+CSS+JS inline, monolithique — **par design**, pour l'envoi par mail). |
| `_explore.py` | Utilitaire Python (aperçu / filtre / stats d'un CSV-XLSX). Appelé par le serveur. |
| `catalog.json` | Miroir du catalogue de données (généré par `local/lib/build_bnp_catalog.py`). |
| `config.json` | Chemins et options. |
| `README.md` | Ce fichier. |

## Lancement
```
cd <dossier de l'app>
node server.js
```
Puis ouvrir `http://localhost:3000`. *(Port modifiable dans `config.json` ; surcharge ponctuelle
possible via la variable d'environnement `PORT`.)*

## Onglets
- **Livrables** — documentation, arbre de décomposition, formulaire de paramètres, exécution (logs live).
- **Sources** — fichiers sources catalogués, accès rapide à l'explorateur et aux agrégateurs.
- **Catalogue** — colonnes (source + calculées) consultables.
- **Explorer** — ouvrir un CSV/XLSX, filtrer, statistiques par colonne (dataviz SVG).
- **Agrégateurs** — consolidation des fichiers mensuels.

## Configuration (`config.json`)
Sur le PC ARTEONYS, les chemins pointent vers `../local/...`. Sur le PC BNP, adapte selon ton
arborescence (par ex. `livrables_path: "./livrables"`, `sources_*` vers tes dossiers réels,
`default_output_path` vers tes sorties).

## Mises à jour par mail
- Changement d'UI -> Ali reçoit un nouveau `index.html` et l'écrase. (1 seul fichier, le plus souvent.)
- Changement de catalogue -> nouveau `catalog.json` (regénéré via `build_bnp_catalog.py`).
- Changement de logique serveur -> nouveau `server.js`.

## Note
Les scripts de `livrables/` sont refactorés en CLI (argparse, sans Tkinter). Chaque script expose
ses paramètres via la section `ARGUMENTS CLI` de son docstring, lue par la Console pour générer le
formulaire. Codes de sortie : 0 succès, 1 erreur fonctionnelle, 2 erreur technique.
