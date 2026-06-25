# CONVENTIONS

## Nommage
- Dossiers livrables : `BNP_LIV_032`, ... (préfixe `BNP_LIV_` = `livrable_ref` Supabase ;
  cas particulier `BNP_LIV_000_CONSOLE`).
- Fichiers de code : `0N.{CODE}.py` où `{CODE}` est le code `[XXXXX]` lu sur la **ligne titre du
  docstring** (source de vérité). Mono-script -> `01.{CODE}.py` ; multi-scripts -> `01.`, `02.`...
- Exceptions CONSOLE : `_shared.py` (bibliothèque) et `a_console.py` (lanceur legacy) gardent leur nom.
- Fonctions Python : snake_case. Fonctions communes : voir `catalog/functions.json` + `lib/fonctions.py`.

## Docstring d'en-tête des scripts (format imposé — parsé par la Console)
```
NOM DU SCRIPT [CODE]
====================
DESCRIPTION
-----------
... 2-3 phrases ...
SOURCES REQUISES
----------------
- NOM (csv|xlsx) — rôle, obligatoire|optionnel
OUTPUTS PRODUITS
----------------
- {output_dir}/{filename}.xlsx — ...
ARGUMENTS CLI
-------------
--source-x PATH         (obligatoire) description
--output-dir PATH       (obligatoire) Dossier de sortie
--output-filename NAME  (obligatoire) Nom du fichier de sortie (sans extension)
DECOMPOSITION
-------------
1. Étape
   1.1 Sous-étape
2. ...
```
- La ligne `ARGUMENTS CLI` doit lister un `--flag METAVAR (obligatoire|optionnel) desc` par ligne.
- La section `DECOMPOSITION` est arborescente (`1.`, `1.1`, ...) ; elle alimente l'arbre dans l'UI.

## Code Python (livrables)
- Chaque script est **AUTONOME** : pas d'import de `local/lib/` ni de `_shared` ; on **inline** les
  fonctions communes (section "FONCTIONS UTILITAIRES (inlinées)").
- **Pas de Tkinter/customtkinter, pas de GUI.** Tous les paramètres via `argparse`.
- Type hints sur les fonctions ; `pathlib.Path` pour les chemins d'arguments ; encoding `utf-8`.
- Logs : `print("[i/N] description")` entre étapes, `print("[OK] ...")` à la fin (français, stdout).
- Codes de sortie : `0` succès, `1` erreur fonctionnelle (fichier/colonne manquant…), `2` technique.
- Refactoring = **chirurgical** : on retire la GUI et on ajoute l'entrée CLI ; la logique métier
  (calculs, jointures, colonnes, plages) est préservée à l'identique.
- Pas d'emoji dans le code de production.

## Code Node.js (bnp/server.js)
- Modules natifs uniquement : `http`, `fs`, `path`, `child_process`, `url`. Aucun package.json à dépendances.
- ES6+ (const/let, arrow, async/await, template literals), pas de `var`.
- Erreurs : try/catch + retour JSON `{error: "..."}`. Logs serveur en français.

## HTML (bnp/index.html)
- **Un seul fichier monolithique** (HTML + CSS + JS inline) — par design (envoi par mail). Pas de
  fichier `.css`/`.js` séparé, pas de sous-dossier, pas de CDN, pas de framework.
- Classes en kebab-case. Variables CSS pour la palette. Pas d'`!important` sauf nécessité.
- Dataviz : **SVG natif** écrit à la main (pas de Chart.js/D3).
- Pas d'emoji ; pictogrammes Unicode discrets si besoin (▸ ◆ ⟶ ◇ ▲).

## Charte visuelle (thème CLAIR « shadcn sans shadcn »)
- Fond `#FFFFFF` ; panel `#F9FAFB` ; bordures `#E5E7EB` / hover `#D1D5DB`.
- Texte `#1F2937` ; secondaire `#6B7280` ; muted `#9CA3AF`.
- **Vert BNP `#009A44`** (accent) ; foncé `#006837` (hover) ; clair `#E6F4EC` (fonds subtils).
- Erreur `#DC2626` ; warning `#F59E0B`.
- Police : `-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif` (pas de Google Fonts).
- Rayons 6px (boutons/inputs) / 8px (cards) ; ombre carte `0 1px 2px rgba(0,0,0,.05)` ;
  transitions douces `150ms`. Dense mais aéré. Pas de gros titres ni de textes d'intro inutiles.
- Note : la console de logs est volontairement sombre (lisibilité terminal).

## JSON
- Indentation 2 espaces ; clés en snake_case ; toujours `version` ; `last_updated` si pertinent.
