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
| `config.json` | Chemins et options (dont `configuration_path` -> les 3 JSON maîtres). |
| `README.md` | Ce fichier. |

> Le catalogue n'est plus un `catalog.json` généré : le serveur lit **directement** les 3 fichiers
> de `../configuration/` (`data.json`, `sources.json`, `livrables.json`). Voir `configuration/README.md`.

## Lancement
```
cd <dossier de l'app>
node server.js
```
Puis ouvrir `http://localhost:3000`. *(Port modifiable dans `config.json` ; surcharge ponctuelle
possible via la variable d'environnement `PORT`.)*

## Onglets
- **Sources** — pour chaque source : renseigner le **fichier réel sur ce PC**, puis « Vérifier
  colonnes » contrôle les noms/nombre de colonnes (indicateur ✓/⚠, non bloquant). Le chemin est
  enregistré dans `configuration/sources.json` (`chemin_local`).
- **Données** — dictionnaire des données (`data.json`) : définition, format, alias, sources où la
  donnée apparaît, et pour les données calculées : dépendances / livrable producteur / formule.
- **Exécution** — livrables (doc, décomposition, paramètres, logs live) et agrégateurs.
- **Exploration** — ouvrir un CSV/XLSX, filtrer, statistiques par colonne (dataviz SVG).

## Configuration (`config.json`)
Sur le PC ARTEONYS, les chemins pointent vers `../local/...`. Sur le PC BNP, adapte selon ton
arborescence (par ex. `livrables_path: "./livrables"`, `sources_*` vers tes dossiers réels,
`default_output_path` vers tes sorties).

## Mises à jour par mail
- Changement d'UI -> Ali reçoit un nouveau `index.html` et l'écrase. (1 seul fichier, le plus souvent.)
- Changement de catalogue -> nouveau(x) fichier(s) de `configuration/` (`data/sources/livrables.json`).
- Changement de logique serveur -> nouveau `server.js`.

## Note
Les scripts de `livrables/` sont refactorés en CLI (argparse, sans Tkinter). Chaque script expose
ses paramètres via la section `ARGUMENTS CLI` de son docstring, lue par la Console pour générer le
formulaire. Codes de sortie : 0 succès, 1 erreur fonctionnelle, 2 erreur technique.
