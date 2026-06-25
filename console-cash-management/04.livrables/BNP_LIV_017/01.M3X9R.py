# -*- coding: utf-8 -*-
"""
SALES & PNB ANALYZER [M3X9R]
============================

DESCRIPTION :
-------------
Analyse croisee du fichier PARC_CLIENT avec les bases CM360, PRGM (Worldline) et
OPTIFLUX. Identification des Sales, calcul du PNB par mois (Janvier a Decembre),
classification des clients (PARC/BPE, SALES/RESEAU) et export CSV enrichi.
Version CLI autonome (sans interface graphique).

SOURCES REQUISES :
------------------
- PARC_CLIENT (csv) - source principale a enrichir, obligatoire
- CM360 (csv) - Sales et RMPM extraits du Legal Entity, obligatoire
- PRGM_CONSOLIDE (csv) - donnees Worldline Jan-Dec, PNB, produits, obligatoire
- OPTIFLUX (csv) - base de reference BPE, obligatoire
- MATCHING_USAGE (csv) - mapping produits -> usages, optionnel

OUTPUTS PRODUITS :
------------------
- <output-dir>/<output-filename>.csv : copie du PARC_CLIENT enrichie (ID_PROGRAMME,
  PRODUIT, USAGE, matching PRGM/CM360/OPTIFLUX, Sales, PNB total + PNB par mois,
  classification client, PRODUIT_OU_USAGE).

ARGUMENTS CLI :
---------------
--source-parc PATH (obligatoire) fichier PARC_CLIENT (CSV)
--source-cm360 PATH (obligatoire) fichier CM360 (CSV)
--source-prgm PATH (obligatoire) fichier PRGM_CONSOLIDE Worldline (CSV)
--opti PATH (obligatoire) fichier OPTIFLUX (CSV)
--matching-usage PATH (optionnel) fichier MATCHING_USAGE produits -> usages (CSV)
--output-dir PATH (obligatoire) repertoire de sortie
--output-filename NAME (obligatoire) nom du fichier de sortie (sans extension)

DECOMPOSITION :
---------------
1. Chargement des fichiers
   1.1 Lecture intelligente CSV (auto-detection encodage/separateur)
2. Preparation des donnees de reference
   2.1 OPTIFLUX : index RP et RS
   2.2 CM360 : extraction RMPM/RS du Legal Entity, index RMPM et RS
   2.3 PRGM : conversion PNB en float, cles RC/RS/produit/id_prog/mois
   2.4 Pre-agregation PNB total et PNB par mois (par RC et par RS)
   2.5 Construction du mapping produit -> usage (fichier MATCHING_USAGE)
   2.6 Pre-calcul ID_PROGRAMME / PRODUIT / USAGE par client
3. Preparation PARC_CLIENT (cles RP/RC/RMPM/RS/Sales)
4. Traitement principal (matching ligne a ligne)
   4.1 Matching OPTIFLUX
   4.2 Matching CM360 (+ recuperation Sales)
   4.3 Matching PRGM (+ PNB total et PNB par mois)
   4.4 ID_PROGRAMME / PRODUIT / USAGE / PRODUIT_OU_USAGE
   4.5 Comparaison Sales PARC vs CM360
   4.6 Classification client (PARC/BPE, SALES/RESEAU)
5. Construction du DataFrame resultat
6. Export CSV

CORRECTIF M3X9R :
- Comparaison Sales amelioree pour gerer les noms colles vs avec espace
  (ex: "FabienLemeunier" vs "Fabien Lemeunier")

BNP Paribas Cash Management - Direction Monetique
Janvier 2025
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
import numpy as np
import re
import unicodedata


# ─── Résolution auto des sources + lecture DuckDB (inliné depuis 06.fonctions) ──
def _find_sources_root():
    here = Path(__file__).resolve()
    for d in [here] + list(here.parents):
        c = d / "03.sources"
        if c.is_dir():
            return c
    return None


def resolve_source(name, required=False):
    root = _find_sources_root()
    if not root:
        if required:
            raise FileNotFoundError(f"Dossier 03.sources introuvable depuis {__file__}")
        return None
    folder = next((m for m in root.glob("*/" + name) if m.is_dir()), None)
    if not folder:
        if required:
            raise FileNotFoundError(f"Source introuvable : 03.sources/*/{name}")
        return None
    exts = (".csv", ".xlsx", ".xls", ".txt")
    files = [f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in exts]
    if not files:
        if required:
            raise FileNotFoundError(f"Aucun fichier source dans {folder}")
        return None

    def _np(f):
        m = re.match(r"\s*(\d+)", f.name)
        return (int(m.group(1)) if m else -1, f.name)
    return max(files, key=_np)


def _read_duck(path, nrows=None):
    p = Path(path)
    if nrows is not None or p.suffix.lower() not in (".csv", ".txt", ""):
        return None
    try:
        import duckdb
        return duckdb.read_csv(str(p), all_varchar=True).df()
    except Exception:
        return None


# --- CONSTANTES ---
VERSION_ID = "M3X9R"

# Mapping code mois -> nom francais
MOIS_MAPPING = {
    '202501': 'JANVIER', '202502': 'FEVRIER', '202503': 'MARS',
    '202504': 'AVRIL', '202505': 'MAI', '202506': 'JUIN',
    '202507': 'JUILLET', '202508': 'AOUT', '202509': 'SEPTEMBRE',
    '202510': 'OCTOBRE', '202511': 'NOVEMBRE', '202512': 'DECEMBRE'
}

# Positions par défaut des colonnes (1-indexed)
DEFAULT_POSITIONS = {
    'parc_rp': 1, 'parc_rc': 14, 'parc_rmpm': 6, 'parc_rs': 8, 'parc_sales': 12,
    'prgm_date': 1, 'prgm_rc': 40, 'prgm_rs': 8, 'prgm_id_prog': 4, 'prgm_produit': 5,
    'cm360_legal_entity': 7, 'cm360_sales': 6,
    'opti_rp': 8, 'opti_rs': 4,
    'ludo_id_prog': 2, 'ludo_produit': 3, 'ludo_usage': 4
}


class SalesPNBAnalyzer_M3X9R:
    def __init__(self, args):
        # Fichiers (mêmes attributs que ceux lus par le worker)
        self.files = {
            "PARC": str(args.source_parc),
            "CM360": str(args.source_cm360),
            "PRGM": str(args.source_prgm),
            "OPTI": str(args.opti),
            "LUDO": str(args.matching_usage) if args.matching_usage else ""
        }

        # Option fichier optionnel (remplace la checkbox GUI)
        self._use_ludo = bool(args.matching_usage)

        # Sortie (remplace asksaveasfilename)
        self.output_dir = Path(args.output_dir)
        self.output_filename = str(args.output_filename)

    # =========================================================================
    # CHARGEMENT FICHIERS
    # =========================================================================
    def load_csv_smart(self, path, nrows=None):
        """Charge un CSV en testant automatiquement encodages et séparateurs."""
        _d = _read_duck(path, nrows)
        if _d is not None:
            return _d
        separators = [';', ',', '\t']
        encodings = ['utf-8', 'latin1', 'cp1252', 'iso-8859-1']

        for sep in separators:
            for enc in encodings:
                try:
                    df_test = pd.read_csv(
                        path, sep=sep, encoding=enc, dtype=str,
                        keep_default_na=False, na_values=[],
                        on_bad_lines='skip', nrows=5
                    )
                    if df_test.shape[1] > 1:
                        return pd.read_csv(
                            path, sep=sep, encoding=enc, dtype=str,
                            keep_default_na=False, na_values=[],
                            on_bad_lines='skip', nrows=nrows
                        )
                except Exception:
                    continue

        return pd.read_csv(
            path, sep=None, engine='python', dtype=str,
            keep_default_na=False, na_values=[],
            on_bad_lines='skip', nrows=nrows
        )

    # =========================================================================
    # FONCTIONS UTILITAIRES
    # =========================================================================
    def clean_id_safe(self, val):
        """Nettoie un identifiant en préservant sa valeur."""
        if pd.isna(val):
            return ""

        s = str(val).strip()

        if s == "" or s.upper() in ["NAN", "NONE", "NULL", "NA", "N/A"]:
            return ""

        if s.startswith('="') and s.endswith('"'):
            s = s[2:-1]

        if s.startswith("'"):
            s = s[1:]

        if s.endswith('.0'):
            base = s[:-2]
            if base.isdigit():
                s = base

        return s.strip()

    def normalize_rs(self, val):
        """Normalise une raison sociale pour comparaison."""
        if pd.isna(val) or str(val).strip() == "":
            return ""

        s = str(val).strip().upper()
        s = unicodedata.normalize('NFD', s)
        s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')

        return s

    def extract_rmpm_from_legal_entity(self, val):
        """Extrait le RMPM de la fin du Legal Entity (format: RS (RMPM))."""
        if pd.isna(val) or str(val).strip() == "":
            return ""

        s = str(val).strip()
        match = re.search(r'\(([^)]+)\)$', s)
        if match:
            return self.clean_id_safe(match.group(1))

        return ""

    def extract_rs_from_legal_entity(self, val):
        """Extrait la RS du Legal Entity (avant les parenthèses)."""
        if pd.isna(val) or str(val).strip() == "":
            return ""

        s = str(val).strip()
        s = re.sub(r'\([^)]*\)$', '', s).strip()

        return s

    def normalize_sales_for_comparison(self, val):
        """
        Normalise un nom de Sales pour comparaison.
        Gère les cas où le nom est collé (FabienLemeunier) ou avec espace (Fabien Lemeunier).
        Retourne une version sans espace, en majuscules, sans accents.
        """
        if pd.isna(val) or str(val).strip() == "":
            return ""

        s = str(val).strip().upper()

        # Supprimer les accents
        s = unicodedata.normalize('NFD', s)
        s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')

        # Supprimer tous les espaces, tirets et caractères spéciaux
        # Pour que "Fabien Lemeunier" et "FabienLemeunier" deviennent identiques
        s = re.sub(r'[\s\-_\.\,\']+', '', s)

        return s

    def sales_are_identical(self, sales1, sales2):
        """
        Compare deux noms de sales.
        Gère les cas :
        - Avec/sans espace : "Fabien Lemeunier" vs "FabienLemeunier"
        - Avec/sans tiret : "Le-Meunier" vs "Lemeunier"
        - Inversion prénom/nom : "Fabien Lemeunier" vs "Lemeunier Fabien"
        """
        if not sales1 or not sales2:
            return False

        # Normalisation : tout en majuscules, sans accents, sans espaces
        norm1 = self.normalize_sales_for_comparison(sales1)
        norm2 = self.normalize_sales_for_comparison(sales2)

        # Comparaison directe (gère collé vs avec espace)
        if norm1 == norm2:
            return True

        # Si pas identique, essayer l'inversion prénom/nom
        # On repart des valeurs originales pour identifier les "mots"
        def get_words(val):
            s = str(val).strip().upper()
            s = unicodedata.normalize('NFD', s)
            s = ''.join(c for c in s if unicodedata.category(c) != 'Mn')
            # Séparer sur espaces et tirets
            words = re.split(r'[\s\-]+', s)
            return [w for w in words if w]

        words1 = get_words(sales1)
        words2 = get_words(sales2)

        # Si les deux ont au moins 2 mots, essayer l'inversion
        if len(words1) >= 2 and len(words2) >= 2:
            # Inverser words1 et comparer
            words1_inverted = words1[1:] + [words1[0]]
            if ''.join(words1_inverted) == ''.join(words2):
                return True
            # Inverser words2 et comparer
            words2_inverted = words2[1:] + [words2[0]]
            if ''.join(words1) == ''.join(words2_inverted):
                return True

        # Comparaison des ensembles de mots (ordre indifférent)
        if len(words1) >= 2 and len(words2) >= 2:
            if set(words1) == set(words2):
                return True

        return False

    def has_valid_sales(self, sales_parc, sales_cm360):
        """Vérifie si au moins un sales valide existe."""
        invalid_values = {'', 'N/A', 'NA', 'NAN', 'NONE', 'NULL'}

        parc_valid = sales_parc and str(sales_parc).strip().upper() not in invalid_values
        cm360_valid = sales_cm360 and str(sales_cm360).strip().upper() not in invalid_values

        return parc_valid or cm360_valid

    def to_float(self, x):
        """Convertit en float de manière robuste."""
        try:
            s = str(x).replace(',', '.').replace(' ', '').replace('\xa0', '')
            return float(s)
        except:
            return 0.0

    def get_mois_name(self, code_mois):
        """Convertit 202501 → JANVIER."""
        return MOIS_MAPPING.get(str(code_mois).strip(), None)

    # =========================================================================
    # WORKER
    # =========================================================================
    def update_prog(self, val, txt):
        print(txt)

    def worker(self, m, pnb_prgm_cols):
        self.update_prog(0.05, "Chargement des données complètes...")

        # Charger tous les fichiers
        df_parc = self.load_csv_smart(self.files["PARC"])
        df_cm360 = self.load_csv_smart(self.files["CM360"])
        df_prgm = self.load_csv_smart(self.files["PRGM"])
        df_opti = self.load_csv_smart(self.files["OPTI"])

        df_ludo = None
        if self._use_ludo and self.files["LUDO"]:
            df_ludo = self.load_csv_smart(self.files["LUDO"])

        total_rows = len(df_parc)

        self.update_prog(0.10, "Préparation des index de recherche...")

        # =====================================================================
        # PREPARATION DES DONNEES DE REFERENCE
        # =====================================================================

        # --- OPTIFLUX ---
        df_opti['_K_RP'] = df_opti[m['opti_rp']].apply(self.clean_id_safe)
        df_opti['_K_RS'] = df_opti[m['opti_rs']].apply(self.normalize_rs)
        opti_set_rp = set(df_opti['_K_RP']) - {''}
        opti_dict_rs = {rs: idx for idx, rs in zip(df_opti.index, df_opti['_K_RS']) if rs}

        # --- CM360 ---
        df_cm360['_RMPM'] = df_cm360[m['cm360_legal_entity']].apply(self.extract_rmpm_from_legal_entity)
        df_cm360['_RS'] = df_cm360[m['cm360_legal_entity']].apply(self.extract_rs_from_legal_entity)
        df_cm360['_K_RS'] = df_cm360['_RS'].apply(self.normalize_rs)
        df_cm360['_SALES'] = df_cm360[m['cm360_sales']]

        cm360_dict_rmpm = {rmpm: idx for idx, rmpm in zip(df_cm360.index, df_cm360['_RMPM']) if rmpm}
        cm360_dict_rs = {rs: idx for idx, rs in zip(df_cm360.index, df_cm360['_K_RS']) if rs}

        self.update_prog(0.15, "Préparation PRGM Worldline...")

        # --- PRGM (Worldline) ---
        # Conversion PNB en float
        for c in pnb_prgm_cols:
            df_prgm[c] = df_prgm[c].apply(self.to_float)
        if pnb_prgm_cols:
            df_prgm['_PNB_TOTAL'] = df_prgm[pnb_prgm_cols].sum(axis=1)
        else:
            df_prgm['_PNB_TOTAL'] = 0.0

        # Colonnes clés
        df_prgm['_K_RC'] = df_prgm[m['prgm_rc']].apply(self.clean_id_safe)
        df_prgm['_K_RS'] = df_prgm[m['prgm_rs']].apply(self.normalize_rs)
        df_prgm['_PRODUIT'] = df_prgm[m['prgm_produit']].astype(str).str.strip()
        df_prgm['_ID_PROG'] = df_prgm[m['prgm_id_prog']].astype(str).str.strip()
        df_prgm['_MOIS'] = df_prgm[m['prgm_date']].astype(str).str.strip()

        # Créer la clé de concaténation pour le matching usage
        df_prgm['_CLE_USAGE'] = df_prgm['_ID_PROG'] + '-' + df_prgm['_PRODUIT']

        # Dictionnaires PRGM pour lookup existence
        prgm_set_rc = set(df_prgm['_K_RC']) - {''}
        prgm_dict_rs = {rs: idx for idx, rs in zip(df_prgm.index, df_prgm['_K_RS']) if rs}

        self.update_prog(0.20, "Pré-agrégation PNB total par clé...")

        # Agrégation PNB TOTAL par RC et RS
        prgm_pnb_total_by_rc = df_prgm.groupby('_K_RC')['_PNB_TOTAL'].sum().to_dict()
        prgm_pnb_total_by_rs = df_prgm.groupby('_K_RS')['_PNB_TOTAL'].sum().to_dict()

        self.update_prog(0.25, "Pré-agrégation PNB par mois...")

        # Agrégation PNB par (clé, mois)
        prgm_pnb_by_rc_mois = df_prgm.groupby(['_K_RC', '_MOIS'])['_PNB_TOTAL'].sum().to_dict()
        prgm_pnb_by_rs_mois = df_prgm.groupby(['_K_RS', '_MOIS'])['_PNB_TOTAL'].sum().to_dict()

        self.update_prog(0.28, "Construction du dictionnaire de mapping usage...")

        # --- FICHIER LUDO (Mapping Produit → Usage) ---
        ludo_mapping = {}
        nb_mappings_ludo = 0

        if df_ludo is not None and m.get('ludo_id_prog') and m.get('ludo_produit') and m.get('ludo_usage'):
            for idx, row in df_ludo.iterrows():
                id_prog = str(row[m['ludo_id_prog']]).strip()
                produit = str(row[m['ludo_produit']]).strip()
                usage = str(row[m['ludo_usage']]).strip()

                if id_prog and produit and usage:
                    if id_prog.upper() not in ['NAN', 'NONE', 'NA', 'N/A', '']:
                        if produit.upper() not in ['NAN', 'NONE', 'NA', 'N/A', '']:
                            if usage.upper() not in ['NAN', 'NONE', 'NA', 'N/A', '']:
                                cle = f"{id_prog}-{produit}"
                                ludo_mapping[cle] = usage
                                nb_mappings_ludo += 1

        self.update_prog(0.30, "Pré-calcul ID_PROGRAMME, PRODUIT et USAGE par client...")

        # Pour chaque client (RC ou RS), stocker le premier ID_PROGRAMME, PRODUIT trouvé
        id_prog_by_rc = {}
        id_prog_by_rs = {}
        produit_by_rc = {}
        produit_by_rs = {}
        usage_by_rc = {}
        usage_by_rs = {}

        for _, row in df_prgm.iterrows():
            rc = row['_K_RC']
            rs = row['_K_RS']
            id_prog = row['_ID_PROG']
            produit = row['_PRODUIT']
            cle_usage = row['_CLE_USAGE']

            id_prog_valid = id_prog and id_prog.upper() not in ['NAN', 'NONE', 'NA', 'N/A', '']
            produit_valid = produit and produit.upper() not in ['NAN', 'NONE', 'NA', 'N/A', '']

            if id_prog_valid and produit_valid:
                if rc and rc not in id_prog_by_rc:
                    id_prog_by_rc[rc] = id_prog
                if rs and rs not in id_prog_by_rs:
                    id_prog_by_rs[rs] = id_prog

                if rc and rc not in produit_by_rc:
                    produit_by_rc[rc] = produit
                if rs and rs not in produit_by_rs:
                    produit_by_rs[rs] = produit

                if ludo_mapping and cle_usage in ludo_mapping:
                    usage = ludo_mapping[cle_usage]
                    if rc and rc not in usage_by_rc:
                        usage_by_rc[rc] = usage
                    if rs and rs not in usage_by_rs:
                        usage_by_rs[rs] = usage

        self.update_prog(0.35, "Préparation PARC_CLIENT...")

        # PREPARATION PARC_CLIENT
        df_parc['_K_RP'] = df_parc[m['parc_rp']].apply(self.clean_id_safe)
        df_parc['_K_RC'] = df_parc[m['parc_rc']].apply(self.clean_id_safe)
        df_parc['_K_RMPM'] = df_parc[m['parc_rmpm']].apply(self.clean_id_safe)
        df_parc['_K_RS'] = df_parc[m['parc_rs']].apply(self.normalize_rs)
        df_parc['_SALES_PARC'] = df_parc[m['parc_sales']].astype(str).str.strip()

        self.update_prog(0.40, "Matching en cours...")

        # =====================================================================
        # TRAITEMENT PRINCIPAL
        # =====================================================================

        # Arrays numpy pour accès rapide
        arr_rp = df_parc['_K_RP'].values
        arr_rc = df_parc['_K_RC'].values
        arr_rmpm = df_parc['_K_RMPM'].values
        arr_rs_norm = df_parc['_K_RS'].values
        arr_rs_source = df_parc[m['parc_rs']].values
        arr_sales_parc = df_parc['_SALES_PARC'].values

        # Pré-allocation des résultats
        results = {
            'ID_PROGRAMME': np.empty(total_rows, dtype=object),
            'PRODUIT': np.empty(total_rows, dtype=object),
            'USAGE': np.empty(total_rows, dtype=object),
            'FOUND_PRGM': np.empty(total_rows, dtype=object),
            'METHOD_FOUND_PRGM': np.empty(total_rows, dtype=object),
            'VALUE_FOUND_PRGM': np.empty(total_rows, dtype=object),
            'FOUND_CM360': np.empty(total_rows, dtype=object),
            'METHOD_FOUND_CM360': np.empty(total_rows, dtype=object),
            'VALUE_FOUND_CM360': np.empty(total_rows, dtype=object),
            'FOUND_OPTIFLUX': np.empty(total_rows, dtype=object),
            'METHOD_FOUND_OPTIFLUX': np.empty(total_rows, dtype=object),
            'VALUE_FOUND_OPTIFLUX': np.empty(total_rows, dtype=object),
            'SALES_PARC': np.empty(total_rows, dtype=object),
            'SALES_CM360': np.empty(total_rows, dtype=object),
            'SALES_IDENTIQUE': np.empty(total_rows, dtype=object),
            'SALES_DIFFERENT': np.empty(total_rows, dtype=object),
            'PNB_PRGM_TOTAL': np.zeros(total_rows, dtype=float),
            'CLIENT_PARC': np.empty(total_rows, dtype=object),
            'CLIENT_PARC_SALES': np.empty(total_rows, dtype=object),
            'CLIENT_PARC_RESEAU': np.empty(total_rows, dtype=object),
            'CLIENT_BPE': np.empty(total_rows, dtype=object),
            'CLIENT_BPE_SALES': np.empty(total_rows, dtype=object),
            'CLIENT_BPE_RESEAU': np.empty(total_rows, dtype=object),
            'PRODUIT_FINAL': np.empty(total_rows, dtype=object),
            'PRODUIT_OU_USAGE': np.empty(total_rows, dtype=object)
        }

        # PNB par mois
        pnb_mois_arrays = {mois: np.zeros(total_rows, dtype=float) for mois in MOIS_MAPPING.keys()}

        # Compteurs pour stats
        nb_sales_identiques = 0

        for i in range(total_rows):
            if i % 2000 == 0:
                pct = 0.40 + (0.45 * i / total_rows)
                self.update_prog(pct, f"Matching ligne {i}/{total_rows}...")

            rp_source = arr_rp[i]
            rc_source = arr_rc[i]
            rmpm_source = arr_rmpm[i]
            rs_norm = arr_rs_norm[i]
            rs_source = arr_rs_source[i]
            sales_parc = arr_sales_parc[i]

            # --- OPTIFLUX ---
            found_opti, method_opti, value_opti = "NO", "N/A", "N/A"
            if rp_source and rp_source in opti_set_rp:
                found_opti, method_opti, value_opti = "YES", "ID_RP_EXACT", rp_source
            elif rs_norm and rs_norm in opti_dict_rs:
                found_opti, method_opti, value_opti = "YES", "RS_EXACT", rs_source

            # --- CM360 ---
            found_cm360, method_cm360, value_cm360, sales_cm360 = "NO", "N/A", "N/A", ""
            if rmpm_source and rmpm_source in cm360_dict_rmpm:
                idx = cm360_dict_rmpm[rmpm_source]
                found_cm360, method_cm360, value_cm360 = "YES", "ID_RMPM_EXACT", rmpm_source
                sales_cm360 = str(df_cm360.at[idx, '_SALES']).strip()
            elif rs_norm and rs_norm in cm360_dict_rs:
                idx = cm360_dict_rs[rs_norm]
                found_cm360, method_cm360, value_cm360 = "YES", "RS_EXACT", rs_source
                sales_cm360 = str(df_cm360.at[idx, '_SALES']).strip()

            # --- PRGM (Worldline) ---
            found_prgm, method_prgm, value_prgm = "NO", "N/A", "N/A"
            pnb_total = 0.0
            prgm_key_type = None
            prgm_key_value = None

            if rc_source and rc_source in prgm_set_rc:
                found_prgm, method_prgm, value_prgm = "YES", "ID_RC_EXACT", rc_source
                pnb_total = prgm_pnb_total_by_rc.get(rc_source, 0.0)
                prgm_key_type = 'rc'
                prgm_key_value = rc_source
            elif rs_norm and rs_norm in prgm_dict_rs:
                found_prgm, method_prgm, value_prgm = "YES", "RS_EXACT", rs_source
                pnb_total = prgm_pnb_total_by_rs.get(rs_norm, 0.0)
                prgm_key_type = 'rs'
                prgm_key_value = rs_norm

            # PNB par mois
            for code_mois in MOIS_MAPPING.keys():
                pnb_mois = 0.0
                if prgm_key_type == 'rc':
                    pnb_mois = prgm_pnb_by_rc_mois.get((prgm_key_value, code_mois), 0.0)
                elif prgm_key_type == 'rs':
                    pnb_mois = prgm_pnb_by_rs_mois.get((prgm_key_value, code_mois), 0.0)
                pnb_mois_arrays[code_mois][i] = pnb_mois

            # --- ID_PROGRAMME, PRODUIT, USAGE ---
            id_programme = "N/A"
            produit = "N/A"
            usage = "N/A"

            if prgm_key_type == 'rc':
                id_programme = id_prog_by_rc.get(prgm_key_value, "N/A")
                produit = produit_by_rc.get(prgm_key_value, "N/A")
                usage = usage_by_rc.get(prgm_key_value, "N/A")
            elif prgm_key_type == 'rs':
                id_programme = id_prog_by_rs.get(prgm_key_value, "N/A")
                produit = produit_by_rs.get(prgm_key_value, "N/A")
                usage = usage_by_rs.get(prgm_key_value, "N/A")

            # PRODUIT_OU_USAGE
            if usage and usage != "N/A":
                produit_ou_usage = usage
            else:
                produit_ou_usage = produit

            # --- Comparaison Sales (CORRIGEE) ---
            sales_identique = "N/A"
            sales_different = "N/A"

            sales_parc_clean = sales_parc if sales_parc and sales_parc.upper() not in ['', 'N/A', 'NAN', 'NONE'] else ""
            sales_cm360_clean = sales_cm360 if sales_cm360 and sales_cm360.upper() not in ['', 'N/A', 'NAN', 'NONE'] else ""

            if sales_parc_clean and sales_cm360_clean:
                if self.sales_are_identical(sales_parc_clean, sales_cm360_clean):
                    sales_identique = "OUI"
                    sales_different = "NON"
                    nb_sales_identiques += 1
                else:
                    sales_identique = "NON"
                    sales_different = "OUI"
            elif sales_parc_clean or sales_cm360_clean:
                sales_identique = "NON"
                sales_different = "OUI"

            # --- Classification Client ---
            has_sales = self.has_valid_sales(sales_parc, sales_cm360)

            if found_opti == "NO":
                client_parc = "YES"
                client_bpe = "NO"
                client_parc_sales = "YES" if has_sales else "NO"
                client_parc_reseau = "NO" if has_sales else "YES"
                client_bpe_sales = "NO"
                client_bpe_reseau = "NO"
            else:
                client_parc = "NO"
                client_bpe = "YES"
                client_parc_sales = "NO"
                client_parc_reseau = "NO"
                client_bpe_sales = "YES" if has_sales else "NO"
                client_bpe_reseau = "NO" if has_sales else "YES"

            # Stocker résultats
            results['ID_PROGRAMME'][i] = id_programme
            results['PRODUIT'][i] = produit
            results['USAGE'][i] = usage
            results['FOUND_PRGM'][i] = found_prgm
            results['METHOD_FOUND_PRGM'][i] = method_prgm
            results['VALUE_FOUND_PRGM'][i] = value_prgm
            results['FOUND_CM360'][i] = found_cm360
            results['METHOD_FOUND_CM360'][i] = method_cm360
            results['VALUE_FOUND_CM360'][i] = value_cm360
            results['FOUND_OPTIFLUX'][i] = found_opti
            results['METHOD_FOUND_OPTIFLUX'][i] = method_opti
            results['VALUE_FOUND_OPTIFLUX'][i] = value_opti
            results['SALES_PARC'][i] = sales_parc if sales_parc else "N/A"
            results['SALES_CM360'][i] = sales_cm360 if sales_cm360 else "N/A"
            results['SALES_IDENTIQUE'][i] = sales_identique
            results['SALES_DIFFERENT'][i] = sales_different
            results['PNB_PRGM_TOTAL'][i] = pnb_total
            results['CLIENT_PARC'][i] = client_parc
            results['CLIENT_PARC_SALES'][i] = client_parc_sales
            results['CLIENT_PARC_RESEAU'][i] = client_parc_reseau
            results['CLIENT_BPE'][i] = client_bpe
            results['CLIENT_BPE_SALES'][i] = client_bpe_sales
            results['CLIENT_BPE_RESEAU'][i] = client_bpe_reseau
            results['PRODUIT_FINAL'][i] = produit
            results['PRODUIT_OU_USAGE'][i] = produit_ou_usage

        self.update_prog(0.88, "Construction du DataFrame résultat...")

        # =====================================================================
        # CONSTRUCTION DU DATAFRAME RESULTAT
        # =====================================================================

        df_parc['ID_PROGRAMME'] = results['ID_PROGRAMME']
        df_parc['PRODUIT'] = results['PRODUIT']
        df_parc['USAGE'] = results['USAGE']

        for col in ['FOUND_PRGM', 'METHOD_FOUND_PRGM', 'VALUE_FOUND_PRGM',
                    'FOUND_CM360', 'METHOD_FOUND_CM360', 'VALUE_FOUND_CM360',
                    'FOUND_OPTIFLUX', 'METHOD_FOUND_OPTIFLUX', 'VALUE_FOUND_OPTIFLUX']:
            df_parc[col] = results[col]

        for col in ['SALES_PARC', 'SALES_CM360', 'SALES_IDENTIQUE', 'SALES_DIFFERENT']:
            df_parc[col] = results[col]

        df_parc['PNB_PRGM_TOTAL'] = results['PNB_PRGM_TOTAL']

        for code_mois, nom_mois in MOIS_MAPPING.items():
            col_name = f"PNB_PRGM_{nom_mois}"
            df_parc[col_name] = pnb_mois_arrays[code_mois]

        for col in ['CLIENT_PARC', 'CLIENT_PARC_SALES', 'CLIENT_PARC_RESEAU',
                    'CLIENT_BPE', 'CLIENT_BPE_SALES', 'CLIENT_BPE_RESEAU']:
            df_parc[col] = results[col]

        df_parc['PRODUIT_FINAL'] = results['PRODUIT_FINAL']
        df_parc['PRODUIT_OU_USAGE'] = results['PRODUIT_OU_USAGE']

        # Supprimer colonnes temporaires
        cols_to_drop = ['_K_RP', '_K_RC', '_K_RMPM', '_K_RS', '_SALES_PARC']
        df_parc = df_parc.drop(columns=[c for c in cols_to_drop if c in df_parc.columns])

        # =====================================================================
        # EXPORT CSV
        # =====================================================================

        self.update_prog(0.95, "Sauvegarde CSV...")

        self.output_dir.mkdir(parents=True, exist_ok=True)
        save_path = str(self.output_dir / f"{self.output_filename}.csv")

        df_parc.to_csv(save_path, sep=';', index=False, encoding='utf-8-sig')

        self.update_prog(1.0, "Terminé !")

        # Statistiques
        found_prgm_count = np.sum(results['FOUND_PRGM'] == 'YES')
        found_cm360_count = np.sum(results['FOUND_CM360'] == 'YES')
        found_opti_count = np.sum(results['FOUND_OPTIFLUX'] == 'YES')
        client_parc_count = np.sum(results['CLIENT_PARC'] == 'YES')
        client_bpe_count = np.sum(results['CLIENT_BPE'] == 'YES')
        client_parc_sales_count = np.sum(results['CLIENT_PARC_SALES'] == 'YES')
        client_bpe_sales_count = np.sum(results['CLIENT_BPE_SALES'] == 'YES')
        usage_found_count = np.sum(results['USAGE'] != 'N/A')

        # Compter Sales
        sales_oui = np.sum(results['SALES_IDENTIQUE'] == 'OUI')
        sales_non = np.sum(results['SALES_IDENTIQUE'] == 'NON')

        stats_msg = f"""Fichier CSV généré : {save_path}

STATISTIQUES MATCHING :
- Total lignes : {total_rows:,}
- Trouvés PRGM : {found_prgm_count:,} ({100*found_prgm_count/total_rows:.1f}%)
- Trouvés CM360 : {found_cm360_count:,} ({100*found_cm360_count/total_rows:.1f}%)
- Trouvés OPTIFLUX : {found_opti_count:,} ({100*found_opti_count/total_rows:.1f}%)

COMPARAISON SALES :
- Sales IDENTIQUES : {sales_oui:,}
- Sales DIFFERENTS : {sales_non:,}

CLASSIFICATION CLIENTS :
- CLIENT_PARC : {client_parc_count:,}
  - dont SALES : {client_parc_sales_count:,}
  - dont RESEAU : {client_parc_count - client_parc_sales_count:,}
- CLIENT_BPE : {client_bpe_count:,}
  - dont SALES : {client_bpe_sales_count:,}
  - dont RESEAU : {client_bpe_count - client_bpe_sales_count:,}

PNB TOTAL : {results['PNB_PRGM_TOTAL'].sum():,.2f} EUR

MAPPING PRODUIT -> USAGE :
- Fichier activé : {'Oui' if ludo_mapping else 'Non'}
- Mappings dans fichier : {nb_mappings_ludo if ludo_mapping else 'N/A'}
- Usages trouvés : {usage_found_count:,}"""

        print(stats_msg)

        return save_path


def build_mapping() -> dict:
    """Construit le mapping des colonnes a partir des positions par defaut.

    Vérification colonnes : déléguée à l'UI web.
    """
    return dict(DEFAULT_POSITIONS)


def parse_args(argv=None) -> argparse.Namespace:
    """Definit et parse les arguments de la ligne de commande."""
    parser = argparse.ArgumentParser(
        description=f"Sales & PNB Analyzer [{VERSION_ID}] - analyse croisee PARC/CM360/PRGM/OPTIFLUX -> CSV"
    )
    parser.add_argument("--source-parc", required=False, default=None, type=Path,
                        help="Fichier PARC_CLIENT (CSV) - source principale (obligatoire)")
    parser.add_argument("--source-cm360", required=False, default=None, type=Path,
                        help="Fichier CM360 (CSV) (obligatoire)")
    parser.add_argument("--source-prgm", required=False, default=None, type=Path,
                        help="Fichier PRGM_CONSOLIDE Worldline (CSV) (obligatoire)")
    parser.add_argument("--opti", required=False, default=None, type=Path,
                        help="Fichier OPTIFLUX (CSV) (obligatoire)")
    parser.add_argument("--matching-usage", required=False, type=Path, default=None,
                        help="Fichier MATCHING_USAGE produits -> usages (CSV) (optionnel)")
    parser.add_argument("--output-dir", required=True, type=Path,
                        help="Repertoire de sortie (obligatoire)")
    parser.add_argument("--output-filename", required=True, type=str,
                        help="Nom du fichier de sortie sans extension (obligatoire)")
    return parser.parse_args(argv)


def main(argv=None) -> int:
    """Point d'entree CLI."""
    args = parse_args(argv)

    # Mapping des colonnes (positions par defaut) et colonnes PNB.
    # Vérification colonnes : déléguée à l'UI web.
    mapping = build_mapping()
    pnb_prgm_cols = []

    try:
        if not args.source_parc:
            args.source_parc = resolve_source("PARC_CLIENT", required=True)
        if not args.source_cm360:
            args.source_cm360 = resolve_source("CM360", required=True)
        if not args.source_prgm:
            args.source_prgm = resolve_source("PRGM_AGREGE", required=True)
        if not args.opti:
            args.opti = resolve_source("OPTIFLUX", required=True)
        analyzer = SalesPNBAnalyzer_M3X9R(args)
        save_path = analyzer.worker(mapping, pnb_prgm_cols)
    except FileNotFoundError as e:
        print(f"[ERREUR] Fichier introuvable : {e}", file=sys.stderr)
        return 2
    except Exception as e:
        print(f"[ERREUR] Echec du traitement : {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1

    print(f"[OK] Analyse terminee. Fichier genere : {save_path}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
