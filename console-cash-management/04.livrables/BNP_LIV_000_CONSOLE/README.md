# BNP_LIV_000_CONSOLE — CONSOLE_OUTILS_REUTILISABLES

Boîte à outils de 7 utilitaires fichiers, **refactorés en CLI** (Session 2). Plus un module de
bibliothèque et un lanceur de bureau legacy (conservés tels quels).

| Fichier | Code | Rôle | Statut |
|---|---|---|---|
| `01.BTXCV.py` | BTXCV | TXT → CSV (exports Teradata) | Refactoré CLI |
| `02.CXLCV.py` | CXLCV | XLSX/XLS → CSV en masse | Refactoré CLI |
| `03.DMRGE.py` | DMRGE | Fusion de CSV (+ transformations) | Refactoré CLI |
| `04.ECLNF.py` | ECLNF | Nettoyage/restructuration CSV-XLSX | Refactoré CLI |
| `05.FRPCH.py` | FRPCH | Recherche/remplacement de caractères | Refactoré CLI |
| `06.GXMET.py` | GXMET | Extraction de métadonnées colonnes | Refactoré CLI |
| `07.HRNAM.py` | HRNAM | Renommage en masse via Excel | Refactoré CLI |
| `_shared.py` | — | Bibliothèque partagée (helpers pandas/IO) | Conservé (lib) |
| `a_console.py` | ACNSL | Lanceur de bureau Tkinter | **Legacy** (remplacé par le GUI web) |

Paramètres et étapes de chaque outil : voir le docstring en tête du script (sections ARGUMENTS CLI
/ DECOMPOSITION), aussi affichés dans l'onglet Livrables de la Console.
