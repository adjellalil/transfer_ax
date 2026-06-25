# configuration/ — Catalogue source de vérité (3 fichiers liés)

Ce dossier contient les **3 JSON maîtres** que tu édites à la main. L'application `bnp/` les lit
**directement** (pas d'étape de build). Tu modifies un fichier → tu l'envoies → tout est à jour.

## Les 3 fichiers et leurs liens

```
livrables.json ──(inputs[].source_ref)──▶ sources.json ──(colonnes[].data_ref)──▶ data.json
   (qui utilise quoi)                        (où vit la donnée, + chemin_local)      (le dictionnaire)
```

| Fichier | Rôle | Clé |
|---|---|---|
| `data.json` | Dictionnaire des **données** (colonnes/concepts). Une donnée (ex. `IBAN`, `ID_RP`) peut vivre dans plusieurs sources. | `datas.<ID_CANONIQUE>` |
| `sources.json` | Catalogue des **sources** (fichiers). `nom_canonique` unique + `alias` (tous les autres noms), `chemin_local` éditable, et `colonnes[].data_ref` vers `data.json`. | `sources.<NOM_CANONIQUE>` |
| `livrables.json` | Catalogue des **livrables**. Chaque `inputs[].arg` (argument CLI) est relié à `source_ref` (source canonique) — c'est ce qui corrige les mauvais noms dans le code. | `livrables.<BNP_LIV_xxx>` |

## Pourquoi `alias` ?
Le même fichier portait jusqu'à 3 noms différents (script / ancien `sources.json` / ancien `column.json`).
Exemple : le script LIV_018 demande `--account "BG-LE-RMPM ACCOUNT"` → source canonique **`IBAN_ACCOUNT`**.
Tous ces noms sont listés dans `alias`, donc on retrouve toujours la source canonique à partir d'un nom de code.

## `chemin_local` — enregistrement des sources au lancement (PC BNP)
Chaque source a un champ `chemin_local` (vide par défaut). Sur le PC BNP, tu le renseignes (depuis la page
**Configuration** de l'app, ou directement dans le JSON) avec le chemin réel du fichier sur ton bureau.
C'est l'étape « identifier / enregistrer les sources ». Aucune synchro : c'est juste une valeur dans le JSON.

## Workflow d'évolution (ajouter une source)
1. Ajouter l'entrée dans `sources.json` (+ ses `colonnes[].data_ref`).
2. Ajouter les nouvelles données dans `data.json` si besoin.
3. Relier la source aux livrables qui l'utilisent (`livrables.json` → `inputs[].source_ref`).
4. Renseigner `chemin_local` sur le PC BNP. Envoyer le(s) fichier(s) modifié(s). Fini.

---

## ⚠️ Cas à valider (`"a_confirmer": true`)
Reconstitués par analyse du code mais à confirmer par toi (ton workflow « ce fichier a l'IBAN en col 4,
donc ça doit être tel fichier »). Liste à trancher :

| # | Sujet | Question |
|---|---|---|
| 1 | **`IBAN_ACCOUNT` vs `IBAN_GROUP`** | Le « BG-LE Account » (LIV_018/019/024/032/034) et le « FORTIS groupe » (LIV_033) sont-ils **le même fichier** ou deux distincts ? Structurellement quasi identiques (IBAN→RMPM/entité/pays). |
| 2 | **`IBAN_SINGLETON` vs `--singleton` (LIV_034)** | Le `--singleton` « Nom entité → IBAN » du LIV_034 est-il bien le FORTIS singleton, ou un autre fichier ? |
| 3 | **`GA_GESTION_DIRECTE`** | `--yannick` (LIV_018/019) et `--source-sales` (LIV_032) pointent vers le même fichier « GA → Sales → Gestion Indirecte » ? (supposé oui) |
| 4 | **`CLEAN_SALES` vs `GA_GESTION_DIRECTE`** | Collision de nom : `--sales` (LIV_034 = fichier de nettoyage) ≠ `--source-sales` (LIV_032 = gestion directe). Confirmé distincts. |
| 5 | **`MONITORING` vs `MONITORING_ONBOARDING_CIB`** | Le `--monitoring` du LIV_024 est-il bien un fichier différent du pivot CIB du LIV_034 ? |
| 6 | **`MONEXT_DETAIL_K2P8N` (LIV_027)** | Le code `K2P8N` (MONEXT ANALYZER) correspond à quel `BNP_LIV_xxx` ? (probablement LIV_036 / M5VTQ). LIV_027 consomme sa sortie. |
| 7 | **`ABSORBES`, `FEF`, `SCOPE_HOLDING`** | Nouvelles sources absentes de l'ancien catalogue — colonnes à documenter. |
| 8 | **`MC1` vs `MC2`** | Différence métier exacte des deux tables Matching Client (Ludovic) ? |
| 9 | **`FICHIER_ACHETEUR`** | 72 colonnes ; décalage des exemples col 16-35 à vérifier ; positions pivots mappées provisoirement. |
| 10 | **Override LIV_034** | `TYPE_ACTIONNAIRE`, `OVERRIDE_PAYEE_PAYER`, `DEBIT_REVENUE_MAPPING`, `COMMISSIONED_PROCESSING`, `OVERRIDE_COUNTRY` : sémantique exacte + à quel argument CLI chacun se rattache. |

## Remplace
Ce dossier **remplace** : `column.json` (racine), `local/catalog/sources.json`, `local/catalog/columns.json`.
Tout leur contenu utile a été consolidé ici. (`functions.json` / `aggregators.json` restent dans `local/catalog/`.)
