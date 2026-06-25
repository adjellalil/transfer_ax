# 03.sources/ — Dépôt des fichiers sources

Trois catégories, **un sous-dossier par source** (nommé exactement comme le nom canonique du
catalogue `05.configuration/sources.json`) :

```
03.sources/
├── 01.interne/      sources maintenues à la main (overrides, data-clean, matching, devises, monitoring…)
├── 02.plateforme/   données reçues des plateformes monétiques (Worldline PRGM, Monext, CCO flux/stock, acheteur…)
└── 03.client/       référentiels clients BNP (parc, référentiel client, IBAN, segment, agences, GA…)
```

## Règle de dépôt (Ali, sur le PC BNP)
Dans le sous-dossier d'une source, dépose tes fichiers `.csv` / `.xlsx` **préfixés d'un numéro
croissant**, le reste du nom est libre (repère visuel) :

```
03.sources/03.client/IBAN_ACCOUNT/
    01.IBAN_ACCOUNT_JANV_2026.csv
    02.IBAN_ACCOUNT_MAI_2026.csv     ← c'est CELUI-CI qui sera utilisé (numéro le plus élevé)
```

## Règle de lecture (les programmes)
Un script de `04.livrables/<LIV>/` **remonte de 2 dossiers** (→ racine du repo), entre dans
`03.sources/`, trouve le sous-dossier portant le **nom canonique** de la source (quelle que soit la
catégorie), et prend le fichier au **numéro le plus élevé**. Aucun chemin en dur : où que tu poses
le repo sur le PC BNP, la résolution fonctionne (helper dans `06.fonctions/fonctions.py`).

> La catégorie (interne/plateforme/client) est purement organisationnelle : la résolution se fait
> par **nom de source**, qui est unique. Tu peux ranger une source dans la catégorie qui te parle.
