# catalog/ — Fonctions & agrégateurs

> **Le catalogue de données a déménagé.** Les sources et les colonnes ne sont plus ici : elles
> vivent désormais dans **`../../configuration/`** (`data.json`, `sources.json`, `livrables.json`),
> lues directement par l'app `bnp/`. Voir `configuration/README.md`.
>
> Fichiers retirés de ce dossier : `sources.json`, `columns.json`, `column.json` (consolidés dans
> `configuration/`). `build_bnp_catalog.py` est supprimé (plus d'étape de build).

Ne restent ici que les deux référentiels non migrés :

| Fichier | Contenu |
|---|---|
| `functions.json` | Fonctions communes (signature, fichier de référence, livrables). Version canonique dans `../lib/fonctions.py`. |
| `aggregators.json` | Programmes d'agrégation XLSX/CSV mensuels -> CSV consolidé (statut `actif` / `a_creer`). Lu directement par `bnp/server.js` (onglet Agrégateurs). |

## Enrichir le catalogue de données
Voir `configuration/README.md` : éditer `data.json` / `sources.json` / `livrables.json`, puis
envoyer le(s) fichier(s) modifié(s) au PC BNP (l'app les lit directement, aucune régénération).
