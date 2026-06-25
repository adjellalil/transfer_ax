
"""
CONSOLE HUB [ACNSL]
====================
Console centrale qui détecte automatiquement tous les programmes du dossier
courant (b_*.py, c_*.py, ...) et permet de les lancer indépendamment.

PRINCIPE :
- Au démarrage, scan du dossier où se trouve a_console.py.
- Tout fichier .py commençant par une lettre minuscule + underscore (b_, c_, d_, ...)
  est considéré comme un module lançable. Pas de limite sur le nombre.
- Chaque module est lancé dans un SOUS-PROCESSUS séparé (subprocess.Popen) :
  - Évite tout conflit Tkinter (plusieurs racines Tk dans un même process).
  - Chaque app a sa propre fenêtre indépendante.
  - Si un module crash, ça n'affecte pas la console.

AFFICHAGE :
- Cartes uniformes en liste scrollable (CTkScrollableFrame).
- Supporte 20+ modules sans pb (le scroll fait le travail).
- Titre extrait du nom de fichier ; description = première ligne du docstring s'il existe.

BNP Paribas Cash Management - Direction Monétique
Mai 2026
"""

import customtkinter as ctk
import subprocess
import sys
import os
import re
import threading

# === CONFIGURATION ===
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("green")

VERSION_ID = "ACNSL"

# Couleurs BNP
VERT_BNP = "#00915A"
VERT_FONCE = "#003D2E"
VERT_CLAIR = "#E8F5E9"
GRIS = "#666666"
GRIS_CLAIR = "#E0E0E0"
GRIS_TRES_CLAIR = "#F5F5F5"
BLANC = "#FFFFFF"
NOIR = "#333333"
ROUGE = "#E53935"

# Pattern de détection d'un module lançable : lettre minuscule + _ + nom.py
# Exemples : b_txt_to_csv.py, c_xlsx_to_csv.py, z_super_outil.py
MODULE_PATTERN = re.compile(r'^([b-z])_([a-zA-Z0-9_]+)\.py$')

class ConsoleHub(ctk.CTk):
    """Hub principal qui liste et lance les modules détectés."""

    def __init__(self):
        super().__init__()

        self.title(f"CONSOLE HUB [{VERSION_ID}]")
        self.geometry("1100x800")
        self.minsize(900, 600)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)
        self.grid_rowconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=0)

        self.script_dir = self._get_script_dir()
        self.modules = []  # liste de tuples (letter, name, title, description, filepath)
        self.launched_procs = []  # tracker des sous-processus

        self.setup_ui()
        self.scan_modules()

    def _get_script_dir(self):
        """Renvoie le dossier où se trouve a_console.py (ou cwd en mode dev)."""
        try:
            return os.path.dirname(os.path.abspath(__file__))
        except NameError:
            return os.getcwd()

    def setup_ui(self):
        # === HEADER ===
        header = ctk.CTkFrame(self, corner_radius=0, fg_color=VERT_BNP, height=70)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)

        hc = ctk.CTkFrame(header, fg_color="transparent")
        hc.pack(fill="x", padx=20, pady=12)

        title_box = ctk.CTkFrame(hc, fg_color="transparent")
        title_box.pack(side="left")

        ctk.CTkLabel(
            title_box,
            text=f"CONSOLE HUB [{VERSION_ID}]",
            font=("Segoe UI", 22, "bold"),
            text_color="white"
        ).pack(anchor="w")

        ctk.CTkLabel(
            title_box,
            text="Lanceur centralisé des outils Data Management",
            font=("Segoe UI", 11),
            text_color="#C8E6C9"
        ).pack(anchor="w")

        # Bouton refresh
        ctk.CTkButton(
            hc, text="Rafraîchir",
            command=self.scan_modules,
            width=120, height=34,
            fg_color="white", text_color=VERT_BNP,
            hover_color=GRIS_CLAIR,
            font=("Segoe UI", 11, "bold")
        ).pack(side="right", padx=(5, 0))

        ctk.CTkButton(
            hc, text="Ouvrir dossier",
            command=self.open_folder,
            width=140, height=34,
            fg_color=VERT_FONCE, text_color="white",
            hover_color=VERT_FONCE,
            font=("Segoe UI", 11, "bold")
        ).pack(side="right", padx=(5, 0))

        # === CONTENU ===
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=20, pady=15)
        content.grid_columnconfigure(0, weight=1)
        content.grid_rowconfigure(1, weight=1)

        # Titre section
        title_frame = ctk.CTkFrame(content, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        ctk.CTkLabel(
            title_frame,
            text="Modules disponibles",
            font=("Segoe UI", 16, "bold"),
            text_color=VERT_BNP
        ).pack(side="left")

        self.lbl_count = ctk.CTkLabel(
            title_frame,
            text="(scan en cours...)",
            font=("Segoe UI", 11),
            text_color=GRIS
        )
        self.lbl_count.pack(side="left", padx=10)

        # Liste scrollable
        self.scroll = ctk.CTkScrollableFrame(content, fg_color=GRIS_TRES_CLAIR, corner_radius=8)
        self.scroll.grid(row=1, column=0, sticky="nsew")
        self.scroll.grid_columnconfigure(0, weight=1)

        # === FOOTER ===
        footer = ctk.CTkFrame(self, corner_radius=0, height=40)
        footer.grid(row=2, column=0, sticky="ew")
        footer.grid_propagate(False)

        self.lbl_status = ctk.CTkLabel(
            footer,
            text=f"Dossier : {self.script_dir}",
            font=("Segoe UI", 10),
            text_color=GRIS
        )
        self.lbl_status.pack(side="left", padx=20, pady=10)

        self.lbl_procs = ctk.CTkLabel(
            footer,
            text="0 fenêtre(s) ouverte(s)",
            font=("Segoe UI", 10),
            text_color=VERT_BNP
        )
        self.lbl_procs.pack(side="right", padx=20, pady=10)

    def open_folder(self):
        """Ouvre le dossier du hub dans l'explorateur Windows / Finder / Linux."""
        try:
            if sys.platform == "win32":
                os.startfile(self.script_dir)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", self.script_dir])
            else:
                subprocess.Popen(["xdg-open", self.script_dir])
        except Exception as e:
            self.lbl_status.configure(text=f"Erreur ouverture dossier : {e}")

    def extract_title_and_desc(self, filepath, name):
        """Extrait titre lisible + description du docstring du fichier."""
        # Titre lisible à partir du nom : txt_to_csv -> TXT TO CSV
        title = name.replace('_', ' ').upper()

        # Tentative d'extraction du docstring (en-tête de fichier)
        description = ""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read(3000)  # premiers 3000 chars suffisent

            # Cherche un docstring """...""" ou '''...'''
            m = re.search(r'^"""(.*?)"""', content, re.DOTALL | re.MULTILINE)
            if not m:
                m = re.search(r"^'''(.*?)'''", content, re.DOTALL | re.MULTILINE)

            if m:
                doc = m.group(1).strip()
                lines = [l.strip() for l in doc.split('\n') if l.strip()]
                # Première ligne = souvent un titre/version (ex: "CSV MASTER TOOL V3 [CDBDM]")
                # On essaie de prendre la 1ère ligne SI elle ressemble à un titre,
                # sinon la 1ère ligne descriptive.
                if lines:
                    # Skip lignes de séparation (=== ou ---)
                    first_real = next((l for l in lines if not re.fullmatch(r'[=\-_*]+', l)), '')
                    title_candidate = first_real

                    # Description = 2-3 premières lignes descriptives (skip séparateurs et titre)
                    desc_lines = []
                    for l in lines:
                        if re.fullmatch(r'[=\-_*]+', l):
                            continue
                        if l == title_candidate:
                            continue
                        if l.startswith(('BNP', 'Mai 2026', 'Février', 'PRINCIPE')):
                            continue
                        desc_lines.append(l)
                        if len(desc_lines) >= 2:
                            break

                    if title_candidate and 5 <= len(title_candidate) <= 80:
                        title = title_candidate

                    description = ' '.join(desc_lines)[:200]
        except Exception:
            pass

        if not description:
            description = "Aucune description disponible."

        return title, description

    def scan_modules(self):
        """Scan du dossier pour détecter les modules b_*.py, c_*.py, etc."""
        self.modules = []

        try:
            files = sorted(os.listdir(self.script_dir))
        except Exception as e:
            self.lbl_status.configure(text=f"Erreur scan : {e}")
            return

        for fname in files:
            m = MODULE_PATTERN.match(fname)
            if not m:
                continue
            letter = m.group(1)
            name = m.group(2)
            filepath = os.path.join(self.script_dir, fname)
            title, desc = self.extract_title_and_desc(filepath, name)
            self.modules.append((letter, name, title, desc, filepath))

        # Tri par lettre (b, c, d, ...) — le tri natural fait déjà ça
        self.modules.sort(key=lambda x: (x[0], x[1]))

        self.render_modules()

    def render_modules(self):
        """Affiche les modules dans la liste scrollable."""
        # Vide la liste
        for w in self.scroll.winfo_children():
            w.destroy()

        self.lbl_count.configure(text=f"({len(self.modules)} module(s) détecté(s))")

        if not self.modules:
            empty = ctk.CTkLabel(
                self.scroll,
                text="Aucun module détecté.\n\n"
                     "Placez des fichiers nommés b_xxx.py, c_xxx.py, ... dans le dossier\n"
                     f"{self.script_dir}\n\n"
                     "Puis cliquez sur Rafraîchir.",
                font=("Segoe UI", 11),
                text_color=GRIS,
                justify="center"
            )
            empty.pack(pady=50)
            return

        for idx, (letter, name, title, desc, filepath) in enumerate(self.modules):
            self.create_module_row(idx, letter, name, title, desc, filepath)

    def create_module_row(self, index, letter, name, title, description, filepath):
        """Crée une ligne pour un module."""
        row = ctk.CTkFrame(
            self.scroll,
            fg_color=BLANC,
            corner_radius=8,
            border_width=1,
            border_color=GRIS_CLAIR
        )
        row.grid(row=index, column=0, sticky="ew", pady=4, padx=5)
        row.grid_columnconfigure(2, weight=1)

        # Pastille lettre
        badge = ctk.CTkFrame(row, fg_color=VERT_BNP, corner_radius=20, width=42, height=42)
        badge.grid(row=0, column=0, rowspan=2, padx=(12, 10), pady=12)
        badge.grid_propagate(False)
        ctk.CTkLabel(
            badge, text=letter.upper(),
            font=("Segoe UI", 18, "bold"),
            text_color="white"
        ).place(relx=0.5, rely=0.5, anchor="center")

        # Index numérique
        ctk.CTkLabel(
            row, text=f"{index + 1:02d}",
            font=("Segoe UI", 10),
            text_color=GRIS,
            width=30
        ).grid(row=0, column=1, rowspan=2, padx=(0, 10))

        # Titre
        ctk.CTkLabel(
            row, text=title,
            font=("Segoe UI", 13, "bold"),
            text_color=NOIR,
            anchor="w"
        ).grid(row=0, column=2, sticky="w", pady=(10, 0))

        # Sous-ligne : description + nom de fichier
        sub_frame = ctk.CTkFrame(row, fg_color="transparent")
        sub_frame.grid(row=1, column=2, sticky="w", pady=(0, 10))

        ctk.CTkLabel(
            sub_frame, text=description,
            font=("Segoe UI", 10),
            text_color=GRIS,
            anchor="w",
            wraplength=600,
            justify="left"
        ).pack(side="left")

        ctk.CTkLabel(
            sub_frame, text=f"  ·  {os.path.basename(filepath)}",
            font=("Consolas", 9),
            text_color=VERT_BNP
        ).pack(side="left")

        # Bouton lancer
        ctk.CTkButton(
            row, text="Lancer",
            command=lambda p=filepath, t=title: self.launch_module(p, t),
            width=110, height=36,
            fg_color=VERT_BNP, hover_color=VERT_FONCE,
            font=("Segoe UI", 11, "bold")
        ).grid(row=0, column=3, rowspan=2, padx=12, pady=12)

    def launch_module(self, filepath, title):
        """Lance un module dans un sous-processus séparé."""
        try:
            # sys.executable garantit qu'on utilise le même interpréteur Python
            # cwd=script_dir pour que les modules trouvent leurs fichiers relatifs
            proc = subprocess.Popen(
                [sys.executable, filepath],
                cwd=self.script_dir
            )
            self.launched_procs.append((proc, title))
            self.lbl_status.configure(text=f"Lancé : {title} (PID {proc.pid})")
            self.update_procs_count()
            # Surveille la fin du process en background pour MAJ le compteur
            threading.Thread(target=self._watch_proc, args=(proc,), daemon=True).start()
        except Exception as e:
            self.lbl_status.configure(text=f"Erreur lancement : {e}")

    def _watch_proc(self, proc):
        """Attend la fin d'un sous-processus pour mettre à jour le compteur."""
        proc.wait()
        # Nettoie les procs terminés
        self.launched_procs = [(p, t) for p, t in self.launched_procs if p.poll() is None]
        try:
            self.after(0, self.update_procs_count)
        except Exception:
            pass  # fenêtre fermée

    def update_procs_count(self):
        """Met à jour le compteur de fenêtres ouvertes."""
        alive = [(p, t) for p, t in self.launched_procs if p.poll() is None]
        self.launched_procs = alive
        n = len(alive)
        self.lbl_procs.configure(text=f"{n} fenêtre(s) ouverte(s)")

if __name__ == "__main__":
    app = ConsoleHub()
    app.mainloop()

