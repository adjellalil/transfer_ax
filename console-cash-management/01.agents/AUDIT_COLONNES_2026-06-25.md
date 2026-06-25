# Audit colonnes & matching — 2026-06-25

## Problème signalé
La page « Vérifier colonnes » indiquait des colonnes manquantes (ex. PRGM : `ID_PROG`, `RC`,
`RS` ; MONEXT idem) **alors qu'elles sont bien présentes** dans les fichiers.

## Cause racine
1. **Les livrables lisent les colonnes par POSITION** (`DEFAULT_POSITIONS`), jamais par nom
   d'en-tête. Le commentaire de LIV_018 le dit : « Vérification colonnes : déléguée à l'UI web ».
2. **L'ancienne vérification** (`server.js` / `handleSourceValidate`) exigeait `noms identiques`
   **ET** `même nombre de colonnes` que `sources.json`.
3. Or le catalogue documentait des **positions fictives** et **incomplètes** : PRGM avait 12
   colonnes documentées (positions inventées) alors que le vrai fichier en a **~41** ; MONEXT 15
   documentées pour **≥55** réelles. Résultat : `count_match` toujours faux + noms jamais trouvés
   → faux « colonnes manquantes ». Le bug ne venait PAS des fichiers.

## Positions réelles (consensus des 15 livrables)

### PRGM_AGREGE (Worldline, ~41 colonnes)
| pos | champ | | pos | champ |
|----|----|----|----|----|
| 2 | MOIS | | 14 | PERIODICITE |
| 3 | NOM_PROGRAMME | | 17 | NB_CARTES *(voir ambiguïté)* |
| 4 | **ID_PROGRAMME** (jointure) | | 24-29 | DEPENSES (flux) |
| 5 | PRODUIT | | 30-36 | PNB |
| 6 | CODE_AGENCE | | 39 | PAYS_APPORTEUR |
| 8 | RS | | **40** | **RC** |
| 9 | IBAN | | 41 | DIFFERE |
| 12 | DEVISE / 13 PLAFOND | | | |

### MONEXT_AGREGE (CCO, ≥55 colonnes)
1 MOIS · 4 RS · 7 PAYS · 9 ID_RP · 10 ID_RC · 11 IBAN · 12 NB_CARTES · 13 DIFFERE ·
15 DEPENSES · 16 NB_TRANSACTIONS · 17 RETRAITS · 19→55 PNB (excl. 33) · 21 INTERCHANGE.
Colonnes 5/9/10/11 « sacralisées » par l'agrégateur LIV_004/N7M4P (confirme les positions).

## Corrections appliquées
1. **`05.configuration/sources.json`** (v2.1) :
   - `colonnes` reconstruites sur les positions réelles pour 13 sources : PRGM_AGREGE,
     MONEXT_AGREGE, CCO_FLUX, CCO_STOCK, PARC_CLIENT, IBAN_ACCOUNT, OPTIFLUX, CM360,
     MATCHING_USAGE, SEG_AGENCE, GA_GESTION_DIRECTE, FICHIER_ACHETEUR, MONITORING.
   - Nouveau champ **`colonnes_min`** sur les 36 sources = nombre minimum de colonnes attendu
     (= position max consommée). Base de la vérification.
   - Nouveau champ **`mode_matching`** (`position` partout, `nom` pour CCO_FLUX/CCO_STOCK qui
     sont lus par mot-clé d'en-tête dans LIV_026).
   - CCO_FLUX/CCO_STOCK : en-têtes réels repris de leur producteur (LIV_004/N3096 et /66Z3Q).
2. **`00.console/server.js`** (`handleSourceValidate`) : vérification désormais **par position +
   nombre de colonnes**. `ok = nb_colonnes_fichier ≥ colonnes_min`. Pour chaque colonne
   documentée on renvoie l'en-tête réel trouvé à sa position (contrôle visuel). Plus de faux
   « manquant » sur écart de nom.
3. **`00.console/index.html`** : rendu mis à jour (tableau position → attendu → trouvé,
   positions absentes si fichier trop court).
4. **Alignement de 3 livrables** sur le consensus (incohérences de lecture confirmées) :
   - LIV_017 (`01.M3X9R.py`) : `prgm_date` 1 → **2** (MOIS).
   - LIV_018 (`01.V2MRG.py`) : `prgm_rc` 41 → **40** (la 41 = DIFFERE).
   - LIV_026 (`01.B2PME.py`) : `WL_POS["id_rc"]` 41 → **40** (idem).

## ⚠️ Ambiguïtés restantes — à confirmer sur un fichier réel
Ces conflits entre livrables ne peuvent pas être tranchés sans voir un en-tête réel. Non modifiés.
- **PRGM NB_CARTES** : col **17** (LIV_024/034/035) vs col **15** (LIV_032, marqué « POSITIONS
  CONFIRMEES »). Catalogue = 17 + note. Métrique, pas une clé de jointure.
- **PARC RC** : col **13** (LIV_029/030/032/033) vs col **14** (LIV_017/019/024/026). Catalogue =
  13 + note ; col 14 = FDC/RC.
- **IBAN_SINGLETON** : 4 colonnes (LIV_033, IBAN col 4) vs ≥7 (LIV_034, IBAN col 7). `colonnes_min`=4.
- **COUNTRY** : original/new en col 1/2 (LIV_035) vs col 2/3 (LIV_034).
- **REFERENTIEL_CLIENT SEGMENT** : col **9** (LIV_019/024/035, = catalogue) vs col **8** (LIV_034,
  qui lit peut-être CATEGORIE). À vérifier côté LIV_034.

➡️ Pour lever ces ambiguïtés : fournir la 1ʳᵉ ligne (en-têtes) d'un vrai PRGM, MONEXT, PARC,
SINGLETON et COUNTRY ; on fige alors les positions exactes et on peut activer aussi le contrôle
par nom.
