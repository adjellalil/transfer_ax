# Cash Management Console

Interface web locale unifiée pour les livrables de la mission
**BNP Paribas Cash Management — Direction Monétique** (Ali Djellalil / Arteonys).

Remplace l'écosystème éparpillé de programmes Python à GUI Tkinter par une seule console
web (Node.js natif, **zéro dépendance npm**) qui pilote l'exécution des scripts livrables.

## Démarrage rapide (agents)
Commence par lire `_agents/STARTER.md`, puis `_agents/ARCHITECTURE.md`,
`_agents/CONVENTIONS.md` et `_agents/JOURNAL.md`.

## Structure
- `_agents/` — instructions et journal de bord (référence ; seul `JOURNAL.md` évolue).
- `local/` — atelier. **Reste sur le PC ARTEONYS.** Contient `livrables/` (code source pullé
  depuis Supabase) et `lib/fonctions.py` (bibliothèque de référence).
- `bnp/` — application web livrable, copiée vers le PC BNP. 4 fichiers, zéro npm.

## Deux machines
- **PC ARTEONYS** : atelier de fabrication (ce dossier, accès Claude Code).
- **PC BNP** : atelier de production (`localhost:3000`, données réelles). Transfert par e-mail.

Aucune donnée BNP réelle ne réside ici — uniquement du code.
