# sources/ — Structure miroir des fichiers de travail

**Aucune donnée réelle ici.** Ce dossier reproduit, côté ARTEONYS, l'organisation que tu auras
sur le PC BNP, afin que Claude Code comprenne où vivent les fichiers. Sur le PC BNP, c'est ici
que tu déposes et produis tes fichiers.

## Organisation

```
sources/
├── originals/   ← fichiers BRUTS reçus (XLSX mensuels Worldline/Monext/CCO, référentiels...)
│                  Convention : un sous-dossier par source, ex. originals/PRGM/, originals/MONEXT/
└── work/        ← fichiers de TRAVAIL (CSV consolidés produits par les agrégateurs, intermédiaires)
                   Convention : un CSV consolidé par source, ex. work/PRGM_consolide.csv
```

Les sous-dossiers par source ne sont **pas pré-créés** (pas de prolifération) : ils se
matérialisent quand tu déposes des fichiers. Les chemins-types sont décrits dans
[../catalog/sources.json](../catalog/sources.json).

## Workflow type
1. Déposer les XLSX bruts dans `originals/<SOURCE>/`.
2. Lancer l'agrégateur correspondant (onglet « Agrégateurs » de la Console) → produit
   `work/<SOURCE>_consolide.csv`.
3. Les scripts livrables consomment les CSV de `work/` et écrivent dans `outputs/`.
