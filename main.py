# ======================================================
# ✅ Data Analyst / AI / AFC / Cybersecurity Desktop App
# Application moderne avec interface unique et sophistiquée
# ======================================================
# Auteur: Fatima Zahra Tayebi (étudiante)
#
# Fichiers attendus dans le même dossier que ce script :
#   - dataBiblio.xlsx  (vos données actuelles)
#
# Dépendances:
#   pip install pandas numpy matplotlib seaborn scikit-learn openpyxl
# ======================================================

import os
import warnings
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("TkAgg")

import matplotlib.pyplot as plt
import seaborn as sns
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.patches import Ellipse
from scipy.spatial import ConvexHull
from scipy.stats import chi2_contingency

from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score, accuracy_score, confusion_matrix, classification_report
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.ensemble import IsolationForest
from sklearn.neighbors import LocalOutlierFactor

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from datetime import datetime


# ======================================================
# 🎨 STYLE MODERNE
# ======================================================

COLORS = {
    # Palette claire et moins "noire"
    "primary": "#2563EB",       # Blue 600
    "secondary": "#1D4ED8",     # Blue 700
    "accent": "#14B8A6",        # Teal 500
    "light": "#F8FAFC",         # Slate 50
    "white": "#FFFFFF",
    "dark": "#0F172A",          # Slate 900
    "gray": "#64748B",          # Slate 500
    "success": "#22C55E",
    "warning": "#F59E0B",
    "error": "#EF4444",
    # UI spécifiques
    "nav_bg": "#EEF2FF",        # Indigo 50
    "nav_btn": "#1E40AF",       # Blue 800
    "nav_btn_hover": "#1D4ED8", # Blue 700
}


FONT_TITLE = ("Segoe UI", 20, "bold")
FONT_SUBTITLE = ("Segoe UI", 14, "bold")
FONT_BTN = ("Segoe UI", 12, "bold")
FONT_SMALL = ("Segoe UI", 10)
FONT_TEXT = ("Segoe UI", 11)
FONT_TEXT_BOLD = ("Segoe UI", 11, "bold")

sns.set_style("whitegrid")


# ======================================================
# 📦 HELPERS
# ======================================================

def app_dir() -> str:
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()

DEFAULT_EXCEL = os.path.join(app_dir(), "dataBiblio.xlsx")

def now_tag() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")

def safe_read_excel(path: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(path)
        return df
    except Exception as e:
        raise RuntimeError(f"Impossible de lire le fichier Excel: {e}")

def numeric_columns(df: pd.DataFrame) -> list[str]:
    num_cols = []
    for c in df.columns:
        if pd.api.types.is_numeric_dtype(df[c]):
            num_cols.append(c)
        else:
            try:
                tmp = pd.to_numeric(df[c].astype(str).str.replace(",", ".", regex=False), errors="coerce")
                if tmp.notna().mean() > 0.8:
                    df[c] = tmp
                    num_cols.append(c)
            except Exception:
                pass
    return num_cols


# ======================================================
# 📱 CLASSE PRINCIPALE DE L'APPLICATION
# ======================================================

class ModernApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Analyst, AI and Cybersecurity Project")
        self.root.geometry("1200x800")
        self.root.configure(bg=COLORS["light"])
        
        
        # Initialisation du contexte de données
        self.ctx = DataContext()
        
        # Variables pour les contrôles
        self.k_var = tk.IntVar(value=4)
        
        # Configuration du style
        self.setup_styles()
        
        # Création de l'interface
        self.setup_ui()
        
        # Chargement initial des données
        self.load_initial_data()
    def create_start_page(self):
        """Page 1 (PDF): Titre + Start + Prof + Étudiant"""
        page = tk.Frame(self.pages_container, bg=COLORS["light"])
        page.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.pages["start"] = page
    
        # Grand cadre comme le PDF
        outer = tk.Frame(page, bg=COLORS["white"], bd=4, relief="solid")
        outer.place(relx=0.12, rely=0.18, relwidth=0.76, relheight=0.64)
    
        # Cadre intérieur
        inner = tk.Frame(outer, bg=COLORS["white"], bd=3, relief="solid")
        inner.place(relx=0.07, rely=0.12, relwidth=0.86, relheight=0.76)
    
        # Titre central (dans un rectangle)
        title_box = tk.Frame(inner, bg=COLORS["white"], bd=2, relief="solid")
        title_box.place(relx=0.22, rely=0.10, relwidth=0.56, relheight=0.22)
    
        tk.Label(
            title_box,
            text="Data Analyst, AI and Cybersecurity\nProject [ IN .... ]",
            font=("Segoe UI", 16, "bold"),
            bg=COLORS["white"],
            fg=COLORS["primary"],
            justify="center"
        ).place(relx=0.5, rely=0.5, anchor="center")
    
        # Bouton Start (rectangle)
        start_btn = tk.Button(
            inner,
            text="Start",
            font=("Segoe UI", 18, "bold"),
            bg=COLORS["white"],
            fg="#D946EF",  # rose/violet comme le PDF
            bd=2,
            relief="solid",
            cursor="hand2",
            command=lambda: self.show_page("menu")
        )
        start_btn.place(relx=0.28, rely=0.42, relwidth=0.44, relheight=0.14)
    
        # Bas: Prof + Étudiant
        prof_box = tk.Frame(inner, bg=COLORS["white"], bd=2, relief="solid")
        prof_box.place(relx=0.08, rely=0.78, relwidth=0.36, relheight=0.14)
        tk.Label(prof_box, text="Dr. EL MKHALET MOUNA", font=FONT_SMALL,
                 bg=COLORS["white"], fg=COLORS["primary"]).place(relx=0.5, rely=0.5, anchor="center")
    
        student_box = tk.Frame(inner, bg=COLORS["white"], bd=2, relief="solid")
        student_box.place(relx=0.56, rely=0.78, relwidth=0.36, relheight=0.14)
        tk.Label(student_box, text="Name of engineering student", font=FONT_SMALL,
                 bg=COLORS["white"], fg=COLORS["primary"]).place(relx=0.5, rely=0.5, anchor="center")
    
    
    def create_menu_page(self):
        """Page 2 (PDF): 4 gros boutons"""
        page = tk.Frame(self.pages_container, bg=COLORS["light"])
        page.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.pages["menu"] = page
    
        # Grand cadre
        outer = tk.Frame(page, bg=COLORS["white"], bd=4, relief="solid")
        outer.place(relx=0.18, rely=0.12, relwidth=0.64, relheight=0.76)
    
        inner = tk.Frame(outer, bg=COLORS["white"], bd=3, relief="solid")
        inner.place(relx=0.06, rely=0.08, relwidth=0.88, relheight=0.86)
    
        # 4 blocs (positions comme PDF)
        def big_btn(parent, text, cmd):
            return tk.Button(
                parent,
                text=text,
                font=("Segoe UI", 14, "bold"),
                bg=COLORS["white"],
                fg="#E11D48" if "DATA-ANALYST" in text else "#DC2626",
                bd=2,
                relief="solid",
                cursor="hand2",
                wraplength=220,
                justify="center",
                command=cmd
            )
    
        b1 = big_btn(inner, "DATA-ANALYST ACP", lambda: self.show_page("acp"))
        b2 = big_btn(inner, "Artificial Intelligence\nClustering and Forecasting", lambda: self.show_page("ai"))
        b3 = big_btn(inner, "DATA-ANALYST AFC", lambda: self.show_page("afc"))
        b4 = big_btn(inner, "CYBER-SECURITY", lambda: self.show_page("cyber"))
    
        # Placement en grille (2x2)
        b1.place(relx=0.10, rely=0.10, relwidth=0.36, relheight=0.32)
        b2.place(relx=0.54, rely=0.10, relwidth=0.36, relheight=0.32)
        b3.place(relx=0.10, rely=0.54, relwidth=0.36, relheight=0.32)
        b4.place(relx=0.54, rely=0.54, relwidth=0.36, relheight=0.32)
    
        # Petit bouton retour (optionnel)
        back = tk.Button(
            page, text="⟵ Retour",
            font=FONT_BTN, bg=COLORS["nav_btn"], fg="white",
            bd=0, cursor="hand2",
            command=lambda: self.show_page("start")
        )
        back.place(relx=0.03, rely=0.03)
    def create_acp_page(self):
        acp_frame = tk.Frame(self.pages_container, bg=COLORS["light"])
        acp_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.pages["acp"] = acp_frame
    
        # Bouton retour menu
        tk.Button(acp_frame, text="⟵ Menu", font=FONT_BTN, bg=COLORS["nav_btn"], fg="white",
                  bd=0, cursor="hand2", command=lambda: self.show_page("menu")).pack(anchor="nw", padx=10, pady=10)
    
        # ---- COPIE du contenu de create_acp_tab (sans self.notebook.add) ----
        _outer, control_frame = self.make_scrollable_sidebar(acp_frame, "Analyse en Composantes Principales")
    
        k_frame = tk.Frame(control_frame, bg=COLORS["nav_bg"])
        k_frame.pack(pady=10, padx=20, fill='x')
    
        tk.Label(k_frame, text="Nombre de clusters (k):",
                 font=FONT_TEXT, bg=COLORS["nav_bg"], fg=COLORS["dark"]).pack(side='left')
    
        k_spin = tk.Spinbox(k_frame, from_=3, to=7, width=5,
                            textvariable=self.k_var, font=FONT_TEXT)
        k_spin.pack(side='right')
    
        ttk.Button(control_frame, text="Appliquer k",
                   command=self.update_k, style='Modern.TButton').pack(pady=10, padx=20, fill='x')
    
        buttons = [
            ("📋 Tableau Moyenne et Écart-Type", self.acp_table_moyenne_ecart_type),
            ("🧮 Matrice Centrée-Réduite", self.acp_matrice_centree_reduite),
            ("🔥 Heat-Map de Corrélation", self.acp_heatmap_correlation),
            ("📈 Calcul des Inerties", self.acp_inerties),
            ("📍 Plan Factoriel des Individus", self.acp_plan_factoriel_individus),
            ("⚫ Cercle de Corrélation", self.acp_cercle_correlation),
            ("⭐ Qualité de Représentation", self.acp_qualite_representation),
            ("🎯 Contributions", self.acp_contributions),
        ]
        for text, command in buttons:
            self.nav_button(control_frame, text, command).pack(pady=5, padx=20, fill="x")
    
        display_frame = tk.Frame(acp_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
    
        self.acp_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.acp_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        self.acp_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.acp_table_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    
    def create_ai_page(self):
        ai_frame = tk.Frame(self.pages_container, bg=COLORS["light"])
        ai_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.pages["ai"] = ai_frame
    
        tk.Button(ai_frame, text="⟵ Menu", font=FONT_BTN, bg=COLORS["nav_btn"], fg="white",
                  bd=0, cursor="hand2", command=lambda: self.show_page("menu")).pack(anchor="nw", padx=10, pady=10)
    
        _outer, control_frame = self.make_scrollable_sidebar(ai_frame, "Clustering et Prédiction")
    
        buttons = [
            ("🔵 Affichage des Clusters", self.ai_affichage_clusters),
            ("📊 Pourcentage des Clusters", self.ai_pourcentage_clusters),
            ("📈 Métriques Random Forest", self.ai_metriques_random_forest),
            ("🔮 Prédiction Nouvel Individu", self.ai_prediction_nouvel_individu),
        ]
        for text, command in buttons:
            self.nav_button(control_frame, text, command).pack(pady=5, padx=20, fill="x")
    
        display_frame = tk.Frame(ai_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
    
        self.ai_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.ai_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        self.ai_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.ai_table_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    
    def create_afc_page(self):
        afc_frame = tk.Frame(self.pages_container, bg=COLORS["light"])
        afc_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.pages["afc"] = afc_frame
    
        tk.Button(afc_frame, text="⟵ Menu", font=FONT_BTN, bg=COLORS["nav_btn"], fg="white",
                  bd=0, cursor="hand2", command=lambda: self.show_page("menu")).pack(anchor="nw", padx=10, pady=10)
    
        _outer, control_frame = self.make_scrollable_sidebar(afc_frame, "Analyse Factorielle des Correspondances")
    
        buttons = [
            ("📥 Import Tableau de Contingence", self.afc_import_data_table),
            ("📊 Matrice de Fréquences", self.afc_matrice_frequences),
            ("χ² Dépendance et Inertie", self.afc_chi2_dependance_inertie),
            ("📏 Distance du χ²", self.afc_distance_chi2),
            ("📈 Plan Factoriel AFC", self.afc_plan_factoriel_associations),
            ("🧪 Test χ² et Interprétation", self.afc_test_chi2_interpretation),
        ]
        for text, command in buttons:
            self.nav_button(control_frame, text, command).pack(pady=5, padx=20, fill="x")
    
        display_frame = tk.Frame(afc_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
    
        self.afc_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.afc_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        self.afc_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.afc_table_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    
    def create_cyber_page(self):
        cyber_frame = tk.Frame(self.pages_container, bg=COLORS["light"])
        cyber_frame.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.pages["cyber"] = cyber_frame
    
        tk.Button(cyber_frame, text="⟵ Menu", font=FONT_BTN, bg=COLORS["nav_btn"], fg="white",
                  bd=0, cursor="hand2", command=lambda: self.show_page("menu")).pack(anchor="nw", padx=10, pady=10)
    
        _outer, control_frame = self.make_scrollable_sidebar(cyber_frame, "Détection d'Anomalies")
    
        buttons = [
            ("📥 IMPORT DATA", self.cyber_import_data),
            ("🌲 Isolation Forest", self.cyber_isolation_forest),
            ("📊 LOF Algorithm", self.cyber_lof),
            ("🛡️ Interprétation et Actions", self.cyber_interpretation_actions),
        ]
        for text, command in buttons:
            self.nav_button(control_frame, text, command).pack(pady=5, padx=20, fill="x")
    
        display_frame = tk.Frame(cyber_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
    
        self.cyber_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.cyber_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
        self.cyber_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.cyber_table_frame.pack(fill='both', expand=True, padx=10, pady=10)

        
    def reset_acp_display(self):
        """Réinitialise complètement l'affichage ACP"""
        # Efface TOUT le contenu
        for widget in self.acp_canvas_frame.winfo_children():
            widget.destroy()
        for widget in self.acp_table_frame.winfo_children():
            widget.destroy()
        
        # Recrée une structure vide mais fonctionnelle
        # Pour le canvas
        self.acp_canvas_placeholder = tk.Frame(self.acp_canvas_frame, bg=COLORS["white"])
        self.acp_canvas_placeholder.pack(fill='both', expand=True)
        
        # Pour les tableaux - crée un notebook pour gérer plusieurs tableaux
        self.acp_table_notebook = ttk.Notebook(self.acp_table_frame)
        self.acp_table_notebook.pack(fill='both', expand=True)
        
        # Frame principal pour les tableaux simples
        self.acp_table_main_frame = tk.Frame(self.acp_table_notebook, bg=COLORS["white"])
        self.acp_table_notebook.add(self.acp_table_main_frame, text="Résultats")
    
    def setup_styles(self):
        """Configure les styles de l'interface"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Style pour les boutons
        style.configure('Modern.TButton',
                        font=FONT_BTN,
                        padding=10,
                        relief='flat',
                        background=COLORS["primary"],
                        foreground=COLORS["white"])
        style.map('Modern.TButton',
                 background=[('active', COLORS["secondary"])])
        
        # Style pour les onglets
        style.configure('Modern.TNotebook',
                       background=COLORS["light"],
                       borderwidth=0)
        style.configure('Modern.TNotebook.Tab',
                       font=FONT_BTN,
                       padding=[20, 10],
                       background=COLORS["light"],
                       foreground=COLORS["dark"])
        style.map('Modern.TNotebook.Tab',
                 background=[('selected', COLORS["white"])],
                 foreground=[('selected', COLORS["primary"])])


    def nav_button(self, parent, text, command):
        """Bouton de navigation avec retour à la ligne (wrap)."""
        return tk.Button(
            parent,
            text=text,
            command=command,
            font=FONT_BTN,
            bg=COLORS["nav_btn"],
            fg=COLORS["white"],
            activebackground=COLORS["nav_btn_hover"],
            activeforeground=COLORS["white"],
            bd=0,
            relief="flat",
            padx=10,
            pady=10,
            wraplength=240,
            justify="center",
            cursor="hand2"
        )

    def make_scrollable_sidebar(self, parent, title_text: str, width: int = 300):
        """Sidebar gauche scrollable (pour accéder aux derniers boutons)."""
        outer = tk.Frame(parent, bg=COLORS["nav_bg"], width=width)
        outer.pack(side="left", fill="y", padx=(0, 10))
        outer.pack_propagate(False)

        canvas = tk.Canvas(outer, bg=COLORS["nav_bg"], highlightthickness=0, bd=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)

        vsb.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner = tk.Frame(canvas, bg=COLORS["nav_bg"])
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_configure(_evt=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(evt):
            canvas.itemconfigure(win, width=evt.width)

        inner.bind("<Configure>", _on_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(evt):
            if evt.delta:
                canvas.yview_scroll(int(-1 * (evt.delta / 120)), "units")

        def _bind_wheel(_evt=None):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
            canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        def _unbind_wheel(_evt=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)

        tk.Label(inner, text=title_text, font=FONT_SUBTITLE,
                 bg=COLORS["nav_bg"], fg=COLORS["dark"]).pack(pady=20, padx=10)

        return outer, inner

    def setup_ui(self):
        """Interface conforme au PDF: Start -> Menu 4 boutons -> Pages modules"""
        # Fenêtre principale
        self.root.configure(bg=COLORS["light"])
    
        # Conteneur des pages
        self.pages_container = tk.Frame(self.root, bg=COLORS["light"])
        self.pages_container.pack(fill="both", expand=True)
    
        self.pages = {}
    
        # Créer pages
        self.create_start_page()
        self.create_menu_page()
        self.create_acp_page()
        self.create_ai_page()
        self.create_afc_page()
        self.create_cyber_page()
    
        # Afficher la 1ère page
        self.show_page("start")
    
    
    def show_page(self, name: str):
        """Affiche une page (frame)"""
        page = self.pages.get(name)
        if page is None:
            return
        page.tkraise()

    
    def create_home_tab(self):
        """Crée l'onglet d'accueil"""
        home_frame = tk.Frame(self.notebook, bg=COLORS["light"])
        self.notebook.add(home_frame, text="🏠 Accueil")
        
        # Conteneur principal
        main_container = tk.Frame(home_frame, bg=COLORS["white"])
        main_container.pack(fill='both', expand=True, padx=50, pady=50)
        
        # Titre
        tk.Label(main_container, text="Bienvenue dans l'Application d'Analyse de Données",
                font=FONT_TITLE, bg=COLORS["white"], fg=COLORS["primary"]).pack(pady=(0, 20))
        
        # Description
        desc_text = """Cette application combine l'analyse de données, l'intelligence artificielle 
et la cybersécurité en une seule plateforme sophistiquée.

Fonctionnalités principales:
• Analyse en Composantes Principales (ACP)
• Clustering et Forêts Aléatoires (AI)
• Analyse Factorielle des Correspondances (AFC)
• Détection d'Anomalies en Cybersécurité

Sélectionnez un module dans les onglets ci-dessus pour commencer."""
        
        tk.Label(main_container, text=desc_text, font=FONT_TEXT, bg=COLORS["white"], 
                fg=COLORS["dark"], justify='left', wraplength=800).pack(pady=20)
        
        # Cadre pour les boutons d'action
        action_frame = tk.Frame(main_container, bg=COLORS["white"])
        action_frame.pack(pady=30)
        
        # Bouton pour changer le dataset
        ttk.Button(action_frame, text="📁 Changer le Dataset Principal", 
                  command=self.change_dataset, style='Modern.TButton').pack(side='left', padx=10)

        self.nav_button(action_frame, "▶ Start", lambda: self.show_page("acp")).pack(side="left", padx=10)
        
        # Information sur le dataset
        info_frame = tk.Frame(main_container, bg=COLORS["light"], relief='groove', borderwidth=2)
        info_frame.pack(fill='x', pady=20, padx=50)
        
        self.dataset_info_label = tk.Label(info_frame, text="Dataset: Non chargé", 
                                          font=FONT_TEXT, bg=COLORS["light"], fg=COLORS["dark"])
        self.dataset_info_label.pack(pady=10)
        
        # Signature
        signature_frame = tk.Frame(main_container, bg=COLORS["white"])
        signature_frame.pack(side='bottom', fill='x', pady=20)
        
        tk.Label(signature_frame, text="Dr. EL MKHALET MOUNA", 
                font=FONT_SMALL, bg=COLORS["white"], fg=COLORS["primary"]).pack(side='left')
        tk.Label(signature_frame, text="Fatima Zahra Tayebi (Étudiante)", 
                font=FONT_SMALL, bg=COLORS["white"], fg=COLORS["primary"]).pack(side='right')
    
    def create_acp_tab(self):
        """Crée l'onglet ACP"""
        acp_frame = tk.Frame(self.notebook, bg=COLORS["light"])
        self.notebook.add(acp_frame, text="📊 ACP")
        
        # Panneau de contrôle à gauche (scrollable)
        _outer, control_frame = self.make_scrollable_sidebar(acp_frame, "Analyse en Composantes Principales")
        
        # Contrôle pour K
        k_frame = tk.Frame(control_frame, bg=COLORS["nav_bg"])
        k_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(k_frame, text="Nombre de clusters (k):", 
                font=FONT_TEXT, bg=COLORS["nav_bg"], fg=COLORS["dark"]).pack(side='left')
        
        k_spin = tk.Spinbox(k_frame, from_=3, to=7, width=5, 
                           textvariable=self.k_var, font=FONT_TEXT)
        k_spin.pack(side='right')
        
        ttk.Button(control_frame, text="Appliquer k", 
                  command=self.update_k, style='Modern.TButton').pack(pady=10, padx=20, fill='x')
        
        # Boutons ACP
        buttons = [
            ("📋 Tableau Moyenne et Écart-Type", self.acp_table_moyenne_ecart_type),
            ("🧮 Matrice Centrée-Réduite", self.acp_matrice_centree_reduite),
            ("🔥 Heat-Map de Corrélation", self.acp_heatmap_correlation),
            ("📈 Calcul des Inerties", self.acp_inerties),
            ("📍 Plan Factoriel des Individus", self.acp_plan_factoriel_individus),
            ("⚫ Cercle de Corrélation", self.acp_cercle_correlation),
            ("⭐ Qualité de Représentation", self.acp_qualite_representation),
            ("🎯 Contributions", self.acp_contributions),
        ]
        
        for text, command in buttons:
            btn = self.nav_button(control_frame, text, command)
            btn.pack(pady=5, padx=20, fill="x")
        
        # Zone d'affichage à droite
        display_frame = tk.Frame(acp_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
        
        # Canvas pour les graphiques
        self.acp_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.acp_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Zone pour les tableaux
        self.acp_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.acp_table_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    def create_ai_tab(self):
        """Crée l'onglet AI"""
        ai_frame = tk.Frame(self.notebook, bg=COLORS["light"])
        self.notebook.add(ai_frame, text="🤖 Intelligence Artificielle")
        
        # Panneau de contrôle (scrollable)
        _outer, control_frame = self.make_scrollable_sidebar(ai_frame, "Clustering et Prédiction")
        
        # Boutons AI
        buttons = [
            ("🔵 Affichage des Clusters", self.ai_affichage_clusters),
            ("📊 Pourcentage des Clusters", self.ai_pourcentage_clusters),
            ("📈 Métriques Random Forest", self.ai_metriques_random_forest),
            ("🔮 Prédiction Nouvel Individu", self.ai_prediction_nouvel_individu),
        ]
        
        for text, command in buttons:
            btn = self.nav_button(control_frame, text, command)
            btn.pack(pady=5, padx=20, fill="x")
        
        # Zone d'affichage
        display_frame = tk.Frame(ai_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
        
        # Canvas pour les graphiques
        self.ai_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.ai_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Zone pour les tableaux
        self.ai_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.ai_table_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    def create_afc_tab(self):
        """Crée l'onglet AFC"""
        afc_frame = tk.Frame(self.notebook, bg=COLORS["light"])
        self.notebook.add(afc_frame, text="📐 AFC")
        
        # Panneau de contrôle (scrollable)
        _outer, control_frame = self.make_scrollable_sidebar(afc_frame, "Analyse Factorielle des Correspondances")
        
        # Boutons AFC
        buttons = [
            ("📥 Import Tableau de Contingence", self.afc_import_data_table),
            ("📊 Matrice de Fréquences", self.afc_matrice_frequences),
            ("χ² Dépendance et Inertie", self.afc_chi2_dependance_inertie),
            ("📏 Distance du χ²", self.afc_distance_chi2),
            ("📈 Plan Factoriel AFC", self.afc_plan_factoriel_associations),
            ("🧪 Test χ² et Interprétation", self.afc_test_chi2_interpretation),
        ]
        
        for text, command in buttons:
            btn = self.nav_button(control_frame, text, command)
            btn.pack(pady=5, padx=20, fill="x")
        
        # Zone d'affichage
        display_frame = tk.Frame(afc_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
        
        # Canvas pour les graphiques
        self.afc_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.afc_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Zone pour les tableaux
        self.afc_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.afc_table_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    def create_cyber_tab(self):
        """Crée l'onglet Cybersécurité"""
        cyber_frame = tk.Frame(self.notebook, bg=COLORS["light"])
        self.notebook.add(cyber_frame, text="🛡️ Cybersécurité")
        
        # Panneau de contrôle (scrollable)
        _outer, control_frame = self.make_scrollable_sidebar(cyber_frame, "Détection d'Anomalies")
        
        # Boutons Cybersécurité
        buttons = [
            ("📥 IMPORT DATA", self.cyber_import_data),
            ("🌲 Isolation Forest", self.cyber_isolation_forest),
            ("📊 LOF Algorithm", self.cyber_lof),
            ("🛡️ Interprétation et Actions", self.cyber_interpretation_actions),
        ]
        
        for text, command in buttons:
            btn = self.nav_button(control_frame, text, command)
            btn.pack(pady=5, padx=20, fill="x")
        
        # Zone d'affichage
        display_frame = tk.Frame(cyber_frame, bg=COLORS["white"])
        display_frame.pack(side='right', fill='both', expand=True)
        
        # Canvas pour les graphiques
        self.cyber_canvas_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.cyber_canvas_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Zone pour les tableaux
        self.cyber_table_frame = tk.Frame(display_frame, bg=COLORS["white"])
        self.cyber_table_frame.pack(fill='both', expand=True, padx=10, pady=10)
    
    def load_initial_data(self):
        """Charge les données initiales"""
        try:
            if os.path.exists(DEFAULT_EXCEL):
                self.ctx.load(DEFAULT_EXCEL)
                self.update_status("✅ Dataset chargé avec succès")
                self.update_dataset_info()
            else:
                self.update_status("⚠️ Dataset par défaut non trouvé")
        except Exception as e:
            self.update_status(f"❌ Erreur de chargement: {str(e)}")
    
    def change_dataset(self):
        """Permet de changer le dataset principal"""
        path = filedialog.askopenfilename(
            title="Choisir votre fichier Excel principal",
            filetypes=[("Excel", "*.xlsx"), ("Tous", "*.*")]
        )
        if path:
            try:
                self.ctx.load(path)
                self.update_status(f"✅ Dataset chargé: {os.path.basename(path)}")
                self.update_dataset_info()
                messagebox.showinfo("Succès", f"Dataset chargé avec succès!\n{len(self.ctx.num_cols)} variables numériques détectées.")
            except Exception as e:
                self.update_status(f"❌ Erreur: {str(e)}")
                messagebox.showerror("Erreur", str(e))
    
    def update_k(self):
        """Met à jour la valeur de k pour le clustering"""
        try:
            self.ctx.set_k(self.k_var.get())
            self.update_status(f"✅ k mis à jour: {self.ctx.k}")
        except Exception as e:
            self.update_status(f"❌ Erreur: {str(e)}")
            messagebox.showerror("Erreur", str(e))
    
    def update_status(self, message):
        """Met à jour le message de statut (safe mode)"""
        # Si la barre de statut existe encore
        if hasattr(self, "status_label"):
            self.status_label.config(text=message)
            self.root.update()
        else:
            # Fallback : afficher dans la console (debug)
            print("[STATUS]", message)

    
    def update_dataset_info(self):
        """Met à jour l'information sur le dataset"""
        if self.ctx.df is not None:
            info = f"Dataset: {os.path.basename(self.ctx.data_path)} | "
            info += f"Lignes: {len(self.ctx.df)} | "
            info += f"Variables: {len(self.ctx.num_cols)} | "
            info += f"k: {self.ctx.k}"
            
            self.dataset_info_label.config(text=info)
            self.dataset_label.config(text=info)
    
    def clear_canvas(self, frame):
        """Efface le contenu d'un frame de canvas"""
        for widget in frame.winfo_children():
            widget.destroy()
    
    def clear_table(self, frame):
        """Efface le contenu d'un frame de tableau"""
        for widget in frame.winfo_children():
            widget.destroy()
    
    def show_plot(self, frame, fig):
        """Affiche une figure matplotlib dans un frame"""
        self.clear_canvas(frame)
        
        canvas = FigureCanvasTkAgg(fig, master=frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill='both', expand=True)
        
        # Ajouter une barre d'outils
        toolbar_frame = tk.Frame(frame)
        toolbar_frame.pack(fill='x')
        
        ttk.Button(toolbar_frame, text="📥 Exporter l'image", 
                  command=lambda: self.export_figure(fig)).pack(side='right', padx=5)
        ttk.Button(toolbar_frame, text="🔄 Actualiser", 
                  command=lambda: canvas.draw()).pack(side='right', padx=5)
    
    
    def show_table(self, frame, df, title):
        """Affiche un dataframe dans un frame (alignement propre + scroll)."""
        self.clear_table(frame)

        # Titre
        tk.Label(frame, text=title, font=FONT_SUBTITLE,
                 bg=COLORS["white"], fg=COLORS["primary"]).pack(pady=(0, 10))

        # Frame Treeview
        tree_frame = tk.Frame(frame, bg=COLORS["white"])
        tree_frame.pack(fill='both', expand=True)

        # Colonnes (on inclut l'index comme vraie colonne pour éviter le décalage)
        cols = ["Index"] + list(df.columns)

        tree = ttk.Treeview(tree_frame, columns=cols, show="headings")

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Headings + columns
        for col in cols:
            tree.heading(col, text=str(col))
            tree.column(col, width=140, anchor="center")

        tree.column("Index", width=120, anchor="center")

        # Données
        for i, row in df.iterrows():
            tree.insert("", "end", values=[i] + list(row))

        tree.pack(side='left', fill='both', expand=True)
        vsb.pack(side='right', fill='y')
        hsb.pack(side='bottom', fill='x')

        # Bouton d'export
        btn_frame = tk.Frame(frame, bg=COLORS["white"])
        btn_frame.pack(fill='x', pady=10)

        ttk.Button(btn_frame, text="📥 Exporter vers Excel",
                   command=lambda: self.export_excel(df, title)).pack(side='right')


    def export_figure(self, fig):
        """Exporte une figure matplotlib"""
        path = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG", "*.png"), ("PDF", "*.pdf"), ("Tous", "*.*")],
            initialfile=f"figure_{now_tag()}.png"
        )
        if path:
            try:
                fig.savefig(path, dpi=300, bbox_inches='tight')
                self.update_status(f"✅ Figure exportée: {path}")
                messagebox.showinfo("Export", f"Figure exportée avec succès!")
            except Exception as e:
                self.update_status(f"❌ Erreur d'export: {str(e)}")
                messagebox.showerror("Erreur", f"Erreur lors de l'export: {str(e)}")
    
    def export_excel(self, df, title):
        """Exporte un dataframe vers Excel"""
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Tous", "*.*")],
            initialfile=f"{title}_{now_tag()}.xlsx".replace(" ", "_")
        )
        if path:
            try:
                df.to_excel(path, index=True)
                self.update_status(f"✅ Données exportées: {path}")
                messagebox.showinfo("Export", f"Données exportées avec succès!")
            except Exception as e:
                self.update_status(f"❌ Erreur d'export: {str(e)}")
                messagebox.showerror("Erreur", f"Erreur lors de l'export: {str(e)}")
    
    # ======================================================
    # 📊 MÉTHODES ACP
    # ======================================================


    def acp_table_moyenne_ecart_type(self):
        """Affiche le tableau des moyennes et écarts-types"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return
    
        self.reset_acp_display()  # <-- AJOUT: Nettoyer l'affichage d'abord
        
        stats = pd.DataFrame({
            "Moyenne": self.ctx.X.mean(),
            "Écart-type": self.ctx.X.std(),
            "Min": self.ctx.X.min(),
            "Max": self.ctx.X.max()
        }).round(3)
        
        self.show_table(self.acp_table_frame, stats, "Tableau Moyenne et Écart-Type")
        self.show_page("acp")
    
    def acp_matrice_centree_reduite(self):
        """Affiche la matrice centrée-réduite"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return
    
        self.reset_acp_display()  # <-- AJOUT: Nettoyer l'affichage d'abord
        
        df_cr = pd.DataFrame(self.ctx.X_scaled, columns=self.ctx.num_cols).round(3)
        self.show_table(self.acp_table_frame, df_cr.head(30), "Matrice Centrée-Réduite (30 premières lignes)")
        self.show_page("acp")
    
    def acp_heatmap_correlation(self):
        """Affiche la heatmap de corrélation"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return
    
        self.reset_acp_display()  # <-- AJOUT: Nettoyer l'affichage d'abord
        
        fig = Figure(figsize=(10, 8), dpi=100)
        ax = fig.add_subplot(111)
        
        corr = self.ctx.X.corr().round(3)
        sns.heatmap(corr, annot=True, fmt=".2f", cmap="RdBu_r", center=0, ax=ax, 
                   cbar_kws={'label': 'Coefficient de corrélation'})
        ax.set_title("Heat-Map de la Matrice de Corrélation", fontsize=14, fontweight='bold')
        
        self.show_plot(self.acp_canvas_frame, fig)
        self.show_page("acp")
    def acp_inerties(self):
        """Affiche les graphiques d'inertie"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return

        # Réinitialiser la zone d'affichage pour l'ACP
        self.reset_acp_display()
        
        explained = self.ctx.explained
        eig = self.ctx.eigvals
        cum = np.cumsum(explained)
        
        fig = Figure(figsize=(12, 5), dpi=100)
        ax1 = fig.add_subplot(121)
        ax2 = fig.add_subplot(122)
        
        # Scree plot
        bars = ax1.bar(range(1, len(eig)+1), eig, color=COLORS["primary"], alpha=0.7)
        ax1.plot(range(1, len(eig)+1), eig, marker="o", color=COLORS["accent"], linewidth=2)
        ax1.set_title("Scree Plot - Valeurs propres", fontweight='bold')
        ax1.set_xlabel("Composante principale")
        ax1.set_ylabel("Valeur propre")
        ax1.grid(True, alpha=0.3)
        
        # Variance cumulée
        ax2.plot(range(1, len(cum)+1), cum*100, marker="s", color=COLORS["success"], linewidth=2)
        ax2.axhline(80, linestyle="--", color=COLORS["warning"], alpha=0.7, label="Seuil 80%")
        ax2.fill_between(range(1, len(cum)+1), 0, cum*100, alpha=0.2, color=COLORS["success"])
        ax2.set_title("Variance cumulée expliquée", fontweight='bold')
        ax2.set_xlabel("Nombre de composantes")
        ax2.set_ylabel("Variance expliquée (%)")
        ax2.set_ylim(0, 105)
        ax2.legend()
        ax2.grid(True, alpha=0.3)
        
        fig.tight_layout()
        self.show_plot(self.acp_canvas_frame, fig)
        self.show_page("acp")
    
    def acp_plan_factoriel_individus(self):
        """Affiche le plan factoriel des individus"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return

        # Réinitialiser la zone d'affichage pour l'ACP
        self.reset_acp_display()
        
        scores = self.ctx.scores[:, :2]
        ids = self.ctx.df[self.ctx.id_col].astype(str) if self.ctx.id_col else self.ctx.df.index.astype(str)
        
        fig = Figure(figsize=(10, 8), dpi=100)
        ax = fig.add_subplot(111)
        
        # Nuage de points
        scatter = ax.scatter(scores[:, 0], scores[:, 1], s=80, alpha=0.8, 
                           color=COLORS["primary"], edgecolors='white', linewidth=1)
        
        # Lignes de référence
        ax.axhline(0, linewidth=0.8, color='gray', alpha=0.5, linestyle='--')
        ax.axvline(0, linewidth=0.8, color='gray', alpha=0.5, linestyle='--')
        
        # Annotations des points
        for i in range(len(ids)):
            ax.annotate(ids.iloc[i], (scores[i, 0], scores[i, 1]), 
                       fontsize=8, alpha=0.7, 
                       xytext=(5, 5), textcoords='offset points',
                       bbox=dict(boxstyle='round,pad=0.3', alpha=0.2, facecolor='white'))
        
        ax.set_xlabel(f"PC1 ({self.ctx.explained[0]*100:.1f}%)", fontweight='bold')
        ax.set_ylabel(f"PC2 ({self.ctx.explained[1]*100:.1f}%)", fontweight='bold')
        ax.set_title("Plan Factoriel des Individus", fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        
        fig.tight_layout()
        self.show_plot(self.acp_canvas_frame, fig)
        self.show_page("acp")
    
    def acp_cercle_correlation(self):
        """Affiche le cercle de corrélation"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return

        # Réinitialiser la zone d'affichage pour l'ACP
        self.reset_acp_display()
        
        corr_axes = self.ctx.corr_axes
        vars_ = self.ctx.num_cols
        x = corr_axes[:, 0]
        y = corr_axes[:, 1]
        
        fig = Figure(figsize=(9, 9), dpi=100)
        ax = fig.add_subplot(111, aspect='equal')
        
        # Cercle unité
        circle = plt.Circle((0, 0), 1, fill=False, linestyle='--', 
                           color=COLORS["gray"], alpha=0.5, linewidth=1.5)
        ax.add_artist(circle)
        
        # Flèches et variables
        for i, v in enumerate(vars_):
            # Flèche
            ax.arrow(0, 0, x[i], y[i], head_width=0.03, head_length=0.03, 
                    length_includes_head=True, color=COLORS["accent"], alpha=0.8)
            
            # Texte
            ax.text(x[i]*1.15, y[i]*1.15, v, 
                   ha="center", va="center", 
                   fontsize=10, fontweight='bold', 
                   color=COLORS["primary"],
                   bbox=dict(boxstyle='round,pad=0.3', alpha=0.2, facecolor='white'))
        
        # Lignes de référence
        ax.axhline(0, linewidth=0.8, color='gray', alpha=0.5)
        ax.axvline(0, linewidth=0.8, color='gray', alpha=0.5)
        
        ax.set_xlim(-1.2, 1.2)
        ax.set_ylim(-1.2, 1.2)
        ax.set_xlabel("PC1", fontweight='bold')
        ax.set_ylabel("PC2", fontweight='bold')
        ax.set_title("Cercle de Corrélation", fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3)
        
        fig.tight_layout()
        self.show_plot(self.acp_canvas_frame, fig)
        self.show_page("acp")
    
    
    def acp_qualite_representation(self):
        """Affiche la qualité de représentation des individus et variables"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return
    
        self.reset_acp_display()  # <-- CORRECTION: Utiliser reset_acp_display()
    
        ids = self.ctx.df[self.ctx.id_col].astype(str) if self.ctx.id_col else self.ctx.df.index.astype(str)
    
        scores_all = self.ctx.scores
        dist2 = np.sum(scores_all**2, axis=1, keepdims=True) + 1e-12
        cos2_ind = (scores_all[:, :2]**2) / dist2
    
        cos2_ind_df = pd.DataFrame(
            cos2_ind,
            columns=["COS² PC1", "COS² PC2"]
        ).round(4)
        cos2_ind_df.insert(0, "ID", ids)
    
        corr_axes = self.ctx.corr_axes
        cos2_var_df = pd.DataFrame(
            corr_axes[:, :2]**2,
            columns=["COS² PC1", "COS² PC2"],
            index=self.ctx.num_cols
        ).reset_index()
        cos2_var_df.rename(columns={"index": "Variable"}, inplace=True)
        cos2_var_df = cos2_var_df.round(4)
    
        # Créer un container pour les deux tableaux
        container = tk.Frame(self.acp_table_frame, bg=COLORS["white"])
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Premier tableau
        frame1 = tk.Frame(container, bg=COLORS["white"])
        frame1.pack(fill="both", expand=True, pady=(0, 10))
        self.show_table(frame1, cos2_ind_df.head(20),
                        "Qualité de représentation – Individus (Top 20)")
        
        # Deuxième tableau
        frame2 = tk.Frame(container, bg=COLORS["white"])
        frame2.pack(fill="both", expand=True)
        self.show_table(frame2, cos2_var_df,
                        "Qualité de représentation – Variables")
    
        self.show_page("acp")
    
    def acp_contributions(self):
        """Affiche les contributions (individus + variables)"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return
        
        self.reset_acp_display()  # <-- CORRECTION: Utiliser reset_acp_display()
    
        ids = self.ctx.df[self.ctx.id_col].astype(str) if self.ctx.id_col else self.ctx.df.index.astype(str)
    
        # Contributions variables
        corr_axes = self.ctx.corr_axes
        contrib_var_df = pd.DataFrame(
            (corr_axes[:, :2]**2) * 100,
            columns=["Contribution PC1 (%)", "Contribution PC2 (%)"],
            index=self.ctx.num_cols
        ).reset_index()
        contrib_var_df.rename(columns={"index": "Variable"}, inplace=True)
        contrib_var_df = contrib_var_df.round(2)
    
        # Contributions individus
        scores2 = self.ctx.scores[:, :2]
        ss = np.sum(scores2**2, axis=0) + 1e-12
        contrib_ind = (scores2**2) / ss * 100
    
        contrib_ind_df = pd.DataFrame(
            contrib_ind,
            columns=["Contribution PC1 (%)", "Contribution PC2 (%)"]
        ).round(2)
        contrib_ind_df.insert(0, "ID", ids)
    
        # Créer un container pour les deux tableaux
        container = tk.Frame(self.acp_table_frame, bg=COLORS["white"])
        container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Premier tableau
        frame1 = tk.Frame(container, bg=COLORS["white"])
        frame1.pack(fill="both", expand=True, pady=(0, 10))
        self.show_table(
            frame1,
            contrib_ind_df.head(20),
            "Contributions – Individus (Top 20)"
        )
        
        # Deuxième tableau
        frame2 = tk.Frame(container, bg=COLORS["white"])
        frame2.pack(fill="both", expand=True)
        self.show_table(
            frame2,
            contrib_var_df,
            "Contributions – Variables"
        )
    
        self.show_page("acp")

    
    # ======================================================
    # 🤖 MÉTHODES AI
    # ======================================================
    
    def ai_affichage_clusters(self):
        """Affiche les clusters avec enveloppes convexes"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return

        self.clear_table(self.ai_table_frame)
        
        scores2 = self.ctx.scores[:, :2]
        labels = self.ctx.df["Cluster"].values
        ids = self.ctx.df[self.ctx.id_col].astype(str) if self.ctx.id_col else self.ctx.df.index.astype(str)
        
        colors = [COLORS["primary"], COLORS["accent"], COLORS["success"], 
                 COLORS["warning"], "#6B46C1", "#ED64A6", "#0EA5E9"]
        
        fig = Figure(figsize=(11, 9), dpi=100)
        ax = fig.add_subplot(111)
        
        # Tracer chaque cluster avec son enveloppe convexe
        for idx, cl in enumerate(sorted(np.unique(labels), key=lambda s: int(s.split()[-1]))):
            mask = labels == cl
            color = colors[idx % len(colors)]
            
            # Points du cluster
            ax.scatter(scores2[mask, 0], scores2[mask, 1], 
                      s=100, alpha=0.8, label=cl, 
                      color=color, edgecolors='white', linewidth=1.5, zorder=3)
            
            # Enveloppe convexe
            if len(scores2[mask]) >= 3:
                try:
                    hull = ConvexHull(scores2[mask])
                    hull_points = scores2[mask][hull.vertices]
                    
                    # Remplir l'enveloppe convexe
                    ax.fill(hull_points[:, 0], hull_points[:, 1], 
                           alpha=0.1, color=color, zorder=1)
                    
                    # Tracer le contour
                    ax.plot(hull_points[:, 0], hull_points[:, 1], 
                           color=color, alpha=0.5, linewidth=2, zorder=2)
                    ax.plot([hull_points[-1, 0], hull_points[0, 0]], 
                           [hull_points[-1, 1], hull_points[0, 1]], 
                           color=color, alpha=0.5, linewidth=2, zorder=2)
                except:
                    pass
            
            # Annotations des points (premiers points seulement pour la lisibilité)
            n_annotations = min(10, len(np.where(mask)[0]))
            for i in np.where(mask)[0][:n_annotations]:
                ax.annotate(ids.iloc[i], (scores2[i, 0], scores2[i, 1]), 
                           fontsize=8, alpha=0.8, 
                           xytext=(5, 5), textcoords='offset points',
                           bbox=dict(boxstyle='round,pad=0.3', alpha=0.2, facecolor='white'))
        
        ax.axhline(0, linewidth=0.8, color='gray', alpha=0.5, linestyle='--')
        ax.axvline(0, linewidth=0.8, color='gray', alpha=0.5, linestyle='--')
        
        ax.set_title(f"Clusters K-means (k={self.ctx.k}) dans le plan factoriel", 
                    fontsize=14, fontweight='bold')
        ax.set_xlabel("PC1")
        ax.set_ylabel("PC2")
        ax.legend(loc='best', fontsize=10)
        ax.grid(True, alpha=0.3)
        
        fig.tight_layout()
        self.show_plot(self.ai_canvas_frame, fig)
        self.show_page("ai")  # Sélectionne l'onglet AI
    
    def ai_pourcentage_clusters(self):
        """Affiche le pourcentage de chaque cluster"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return

        # Nettoyer l\'affichage pour permettre de re-cliquer sans bug
        self.clear_canvas(self.ai_canvas_frame)
        self.clear_table(self.ai_table_frame)

        if "Cluster" not in self.ctx.df.columns:
            self.ctx.set_k(self.ctx.k)

        counts = self.ctx.df["Cluster"].value_counts().sort_index()
        perc = (counts / counts.sum() * 100).round(2)
        
        # Créer un dataframe
        out = pd.DataFrame({"Effectif": counts, "Pourcentage (%)": perc})
        
        # Créer un graphique camembert
        fig = Figure(figsize=(10, 8), dpi=100)
        ax1 = fig.add_subplot(121)
        ax2 = fig.add_subplot(122)
        
        # Camembert
        colors_pie = [COLORS["primary"], COLORS["accent"], COLORS["success"], 
                     COLORS["warning"], "#6B46C1", "#ED64A6", "#0EA5E9"]
        
        wedges, texts, autotexts = ax1.pie(perc, labels=counts.index, 
                                          colors=colors_pie[:len(perc)],
                                          autopct='%1.1f%%', startangle=90)
        
        for autotext in autotexts:
            autotext.set_color('white')
            autotext.set_fontweight('bold')
        
        ax1.set_title(f"Distribution des Clusters (k={self.ctx.k})", 
                     fontweight='bold')
        
        # Bar chart
        x_pos = np.arange(len(counts))
        bars = ax2.bar(x_pos, counts.values, color=colors_pie[:len(counts)], alpha=0.8)
        ax2.set_xlabel("Cluster")
        ax2.set_ylabel("Effectif")
        ax2.set_title("Effectif par Cluster", fontweight='bold')
        ax2.set_xticks(x_pos)
        ax2.set_xticklabels(counts.index)
        ax2.grid(True, alpha=0.3, axis='y')
        
        # Ajouter les valeurs sur les barres
        for idx, (bar, count) in enumerate(zip(bars, counts.values)):
            height = bar.get_height()
            pct = float(perc.iloc[idx])
            ax2.text(bar.get_x() + bar.get_width()/2., height + 0.5,
                    f"{count}\\n({pct:.1f}%)",
                    ha='center', va='bottom', fontsize=9)
        
        fig.tight_layout()
        self.show_plot(self.ai_canvas_frame, fig)
        
        # Afficher aussi le tableau
        self.show_table(self.ai_table_frame, out, "Répartition des Clusters")
        self.show_page("ai")
    
    def ai_metriques_random_forest(self):
        """Affiche les métriques du Random Forest"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return

        self.clear_canvas(self.ai_canvas_frame)
        self.clear_table(self.ai_table_frame)

        if self.ctx.rf is None or self.ctx.X_test is None:
            # Recalcul sécurité
            self.ctx.set_k(self.ctx.k)

        y_pred = self.ctx.rf.predict(self.ctx.X_test)
        acc = accuracy_score(self.ctx.y_test, y_pred)
        cm = confusion_matrix(self.ctx.y_test, y_pred, labels=self.ctx.rf.classes_)
        cr = classification_report(self.ctx.y_test, y_pred, output_dict=True)
        
        # Créer les graphiques
        fig = Figure(figsize=(12, 5), dpi=100)
        ax1 = fig.add_subplot(121)
        ax2 = fig.add_subplot(122)
        
        # Matrice de confusion
        sns.heatmap(cm, annot=True, fmt="d", cmap="Blues", ax=ax1,
                   xticklabels=self.ctx.rf.classes_, 
                   yticklabels=self.ctx.rf.classes_,
                   cbar_kws={'label': 'Nombre d\'individus'})
        ax1.set_title("Matrice de Confusion - Random Forest", fontweight='bold')
        ax1.set_xlabel("Prédiction")
        ax1.set_ylabel("Réel")
        
        # Importance des variables
        fi = pd.DataFrame({"Variable": self.ctx.num_cols, 
                          "Importance": self.ctx.rf.feature_importances_})
        fi = fi.sort_values("Importance", ascending=False)
        
        colors_bar = plt.cm.Blues(np.linspace(0.5, 1, len(fi)))
        bars = ax2.barh(fi["Variable"], fi["Importance"], color=colors_bar)
        ax2.set_title("Importance des Variables", fontweight='bold')
        ax2.set_xlabel("Importance")
        ax2.invert_yaxis()
        ax2.grid(True, alpha=0.3, axis='x')
        
        # Ajouter les valeurs
        for bar, imp in zip(bars, fi["Importance"]):
            ax2.text(bar.get_width() + 0.01, bar.get_y() + bar.get_height()/2,
                    f'{imp:.3f}', va='center', fontsize=9)
        
        fig.tight_layout()
        self.show_plot(self.ai_canvas_frame, fig)
        
        # Afficher les métriques numériques
        metrics_df = pd.DataFrame([
            {"Métrique": "Accuracy", "Valeur": f"{acc:.4f}"},
            {"Métrique": "Macro avg F1", "Valeur": f"{cr.get('macro avg', {}).get('f1-score', float('nan')):.3f}"},
            {"Métrique": "Weighted avg F1", "Valeur": f"{cr.get('weighted avg', {}).get('f1-score', float('nan')):.3f}"},
        ])

        for cls in cr:
            if cls not in ['accuracy', 'macro avg', 'weighted avg']:

                new_row = pd.DataFrame({
                    "Métrique": [f"Précision ({cls})", f"Rappel ({cls})", f"F1-score ({cls})"],
                    "Valeur": [f"{cr[cls]['precision']:.3f}", 
                              f"{cr[cls]['recall']:.3f}", 
                              f"{cr[cls]['f1-score']:.3f}"]
                })
                metrics_df = pd.concat([metrics_df, new_row], ignore_index=True)
        
        self.show_table(self.ai_table_frame, metrics_df, "Métriques du Random Forest")
        self.show_page("ai")
    
    def ai_prediction_nouvel_individu(self):
        """Ouvre une fenêtre pour prédire un nouvel individu"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return
        
        # Créer une fenêtre modale
        pred_window = tk.Toplevel(self.root)
        pred_window.title("Prédiction de Nouvel Individu")
        pred_window.geometry("600x700")
        pred_window.configure(bg=COLORS["light"])
        pred_window.transient(self.root)
        pred_window.grab_set()
        
        # Titre
        tk.Label(pred_window, text="Prédiction avec Random Forest", 
                font=FONT_SUBTITLE, bg=COLORS["light"], fg=COLORS["primary"]).pack(pady=20)
        
        # Description
        desc = "Remplissez les valeurs pour chaque variable, puis cliquez sur 'Prédire'."
        tk.Label(pred_window, text=desc, font=FONT_TEXT, 
                bg=COLORS["light"], fg=COLORS["dark"]).pack(pady=(0, 20))
        
        # Cadre pour les entrées
        input_frame = tk.Frame(pred_window, bg=COLORS["white"], relief='groove', borderwidth=2)
        input_frame.pack(fill='both', expand=True, padx=30, pady=10)
        
        # Canvas et scrollbar pour les nombreuses variables
        canvas = tk.Canvas(input_frame, bg=COLORS["white"])
        scrollbar = ttk.Scrollbar(input_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=COLORS["white"])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Entrées pour chaque variable
        entries = {}
        for i, col in enumerate(self.ctx.num_cols):
            row_frame = tk.Frame(scrollable_frame, bg=COLORS["white"])
            row_frame.pack(fill='x', padx=20, pady=8)
            
            tk.Label(row_frame, text=col, font=FONT_TEXT, 
                    bg=COLORS["white"], fg=COLORS["dark"], width=25, anchor='w').pack(side='left')
            
            # Valeur par défaut (moyenne de la variable)
            default_val = self.ctx.X[col].mean()
            var = tk.StringVar(value=f"{default_val:.2f}")
            
            entry = tk.Entry(row_frame, textvariable=var, font=FONT_TEXT, 
                           width=15, justify='right')
            entry.pack(side='right')
            
            entries[col] = (entry, var)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Fonction de prédiction
        def predict():
            try:
                # Récupérer les valeurs
                vals = []
                for c in self.ctx.num_cols:
                    entry, var = entries[c]
                    v = float(entry.get().replace(",", "."))
                    vals.append(v)
                
                # Faire la prédiction
                new_df = pd.DataFrame([vals], columns=self.ctx.num_cols)
                pred = self.ctx.rf.predict(new_df)[0]
                proba = self.ctx.rf.predict_proba(new_df)[0]
                proba_map = dict(zip(self.ctx.rf.classes_, proba))
                top = sorted(proba_map.items(), key=lambda x: x[1], reverse=True)
                
                # Afficher les résultats
                result_text = f"✅ Cluster prédit: {pred}\n\n"
                result_text += "Probabilités par cluster:\n"
                for cl, p in top:
                    result_text += f"• {cl}: {p:.1%}\n"
                
                messagebox.showinfo("Résultat de la prédiction", result_text)
                
            except ValueError as e:
                messagebox.showerror("Erreur", f"Valeur invalide: {str(e)}\nVeuillez entrer des nombres valides.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la prédiction: {str(e)}")
        
        # Boutons
        btn_frame = tk.Frame(pred_window, bg=COLORS["light"])
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="🔮 Prédire", command=predict, 
                  style='Modern.TButton').pack(side='left', padx=10)
        ttk.Button(btn_frame, text="❌ Fermer", 
                  command=pred_window.destroy, style='Modern.TButton').pack(side='left', padx=10)
    
    # ======================================================
    # 📐 MÉTHODES AFC
    # ======================================================
    
    def afc_import_data_table(self):
        """Importe et affiche le tableau de contingence AFC"""
        if self.ctx.df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord charger un dataset.")
            return
        
        try:
            # Construire le tableau 8x8
            base = self.ctx.df.copy()
            if self.ctx.id_col:
                base["_ID_"] = base[self.ctx.id_col].astype(str)
            else:
                base["_ID_"] = base.index.astype(str)
            
            # Prendre 8 individus
            base8 = base.head(8).reset_index(drop=True)
            
            # Prendre 8 variables
            vars_ = self.ctx.num_cols.copy()
            if len(vars_) >= 8:
                vars_ = vars_[:8]
            else:
                # Ajouter des variables dérivées
                while len(vars_) < 8:
                    name = f"Var_{len(vars_)+1}"
                    base8[name] = base8[self.ctx.num_cols].mean(axis=1) + np.random.randn(len(base8)) * 0.1
                    vars_.append(name)
            
            vars_ = vars_[:8]
            
            # Calculer les poids 1-10
            W = pd.DataFrame(index=base8["_ID_"].tolist(), columns=vars_, dtype=int)
            for v in vars_:
                x = base8[v].astype(float).values
                if np.std(x) == 0:
                    W[v] = 5
                else:
                    x_norm = (x - np.min(x)) / (np.max(x) - np.min(x) + 1e-10)
                    W[v] = np.round(1 + x_norm * 9).astype(int)
            
            # Stocker dans le contexte
            self.ctx.afc_table = W
            self.ctx.afc_rows = W.index.tolist()
            self.ctx.afc_cols = vars_
            
            # Afficher
            self.show_table(self.afc_table_frame, W, "Tableau de Contingence AFC (8x8)")
            self.show_page("afc")  # Sélectionne l'onglet AFC
            self.update_status("✅ Tableau AFC importé avec succès")
            
        except Exception as e:
            self.update_status(f"❌ Erreur AFC: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur lors de l'import AFC: {str(e)}")
    
    def afc_matrice_frequences(self):
        """Affiche la matrice de fréquences AFC"""
        if not hasattr(self.ctx, 'afc_table') or self.ctx.afc_table is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer le tableau de contingence.")
            return
        
        try:
            N = self.ctx.afc_table.values.astype(float)
            N_sum = N.sum()
            
            if N_sum == 0:
                messagebox.showerror("Erreur", "La somme du tableau est nulle.")
                return
            
            P = N / N_sum
            dfP = pd.DataFrame(P, index=self.ctx.afc_rows, columns=self.ctx.afc_cols).round(4)
            
            self.show_table(self.afc_table_frame, dfP, "Matrice de Fréquences AFC")
            self.show_page("afc")
            
        except Exception as e:
            self.update_status(f"❌ Erreur AFC: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans le calcul des fréquences: {str(e)}")
    
    def afc_chi2_dependance_inertie(self):
        """Calcule et affiche le χ² et l'inertie"""
        if not hasattr(self.ctx, 'afc_table') or self.ctx.afc_table is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer le tableau de contingence.")
            return
        
        try:
            N = self.ctx.afc_table.values.astype(float)
            n = N.sum()
            
            if n == 0:
                messagebox.showerror("Erreur", "La somme du tableau est nulle.")
                return
            
            P = N / n
            r = P.sum(axis=1, keepdims=True)
            c = P.sum(axis=0, keepdims=True)
            E = r @ c
            chi2 = n * np.sum((P - E)**2 / (E + 1e-12))
            inertia_total = chi2 / n
            
            out = pd.DataFrame({
                "Chi-deux (χ²)": [chi2],
                "Inertie totale": [inertia_total],
                "n total": [n],
            }).round(6)
            
            self.show_table(self.afc_table_frame, out, "Test χ² et Inertie")
            self.show_page("afc")
            
        except Exception as e:
            self.update_status(f"❌ Erreur AFC: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans le calcul du χ²: {str(e)}")
    
    def afc_distance_chi2(self):
        """Calcule et affiche la distance du χ²"""
        if not hasattr(self.ctx, 'afc_table') or self.ctx.afc_table is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer le tableau de contingence.")
            return
        
        try:
            N = self.ctx.afc_table.values.astype(float)
            n = N.sum()
            
            if n == 0:
                messagebox.showerror("Erreur", "La somme du tableau est nulle.")
                return
            
            P = N / n
            r = P.sum(axis=1)
            c = P.sum(axis=0)
            
            row_prof = (P.T / (r + 1e-12)).T
            Dr_inv = np.diag(1/(c + 1e-12))
            D = np.zeros((row_prof.shape[0], row_prof.shape[0]))
            
            for i in range(row_prof.shape[0]):
                for j in range(row_prof.shape[0]):
                    diff = row_prof[i] - row_prof[j]
                    D[i, j] = diff @ Dr_inv @ diff.T
            
            dfD = pd.DataFrame(D, index=self.ctx.afc_rows, columns=self.ctx.afc_rows).round(4)
            
            self.show_table(self.afc_table_frame, dfD, "Distance du χ² entre individus")
            self.show_page("afc")
            
        except Exception as e:
            self.update_status(f"❌ Erreur AFC: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans le calcul des distances: {str(e)}")
    
    def afc_plan_factoriel_associations(self):
        """Affiche le plan factoriel AFC"""
        if not hasattr(self.ctx, 'afc_table') or self.ctx.afc_table is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer le tableau de contingence.")
            return
        
        try:
            N = self.ctx.afc_table.values.astype(float)
            
            # Analyse des correspondances
            n = N.sum()
            P = N / n
            r = P.sum(axis=1, keepdims=True)
            c = P.sum(axis=0, keepdims=True)
            
            Dr_inv_sqrt = np.diag((1/np.sqrt(r.flatten() + 1e-12)))
            Dc_inv_sqrt = np.diag((1/np.sqrt(c.flatten() + 1e-12)))
            S = Dr_inv_sqrt @ (P - r @ c) @ Dc_inv_sqrt
            
            U, s, VT = np.linalg.svd(S, full_matrices=False)
            eig = s**2
            explained = eig / eig.sum() if eig.sum() > 0 else eig
            
            F = Dr_inv_sqrt @ U @ np.diag(s)
            G = Dc_inv_sqrt @ VT.T @ np.diag(s)
            
            # Stocker les résultats
            self.ctx.afc_result = {
                "eig": eig,
                "explained": explained,
                "F": F[:, :2],
                "G": G[:, :2],
            }
            
            # Créer le graphique
            fig = Figure(figsize=(11, 9), dpi=100)
            ax = fig.add_subplot(111)
            
            # Individus
            for i, (x, y) in enumerate(F[:, :2]):
                ax.scatter(x, y, s=100, color=COLORS["primary"], alpha=0.7, zorder=3)
                ax.annotate(self.ctx.afc_rows[i], (x, y), 
                          fontsize=9, fontweight='bold',
                          xytext=(5, 5), textcoords='offset points',
                          bbox=dict(boxstyle='round,pad=0.3', alpha=0.3, facecolor='white'),
                          zorder=4)
            
            # Variables
            for j, (x, y) in enumerate(G[:, :2]):
                ax.scatter(x, y, s=150, marker='s', color=COLORS["accent"], alpha=0.7, zorder=3)
                ax.annotate(self.ctx.afc_cols[j], (x, y), 
                          fontsize=9, fontweight='bold',
                          xytext=(5, 5), textcoords='offset points',
                          bbox=dict(boxstyle='round,pad=0.3', alpha=0.3, facecolor='white'),
                          zorder=4)
            
            # Lignes de référence
            ax.axhline(0, linewidth=0.8, color='gray', alpha=0.5, linestyle='--')
            ax.axvline(0, linewidth=0.8, color='gray', alpha=0.5, linestyle='--')
            
            ax.set_xlabel("Dimension 1", fontweight='bold')
            ax.set_ylabel("Dimension 2", fontweight='bold')
            ax.set_title("Plan Factoriel AFC - Individus et Variables", 
                        fontsize=14, fontweight='bold')
            ax.grid(True, alpha=0.3)
            
            # Légende
            from matplotlib.lines import Line2D
            legend_elements = [
                Line2D([0], [0], marker='o', color='w', markerfacecolor=COLORS["primary"], 
                      markersize=10, label='Individus'),
                Line2D([0], [0], marker='s', color='w', markerfacecolor=COLORS["accent"], 
                      markersize=10, label='Variables')
            ]
            ax.legend(handles=legend_elements, loc='best')
            
            fig.tight_layout()
            self.show_plot(self.afc_canvas_frame, fig)
            self.show_page("afc")
            
            # Afficher aussi les valeurs propres
            eig_df = pd.DataFrame({
                "Dimension": [f"Dim{i+1}" for i in range(len(eig))],
                "Valeur propre": eig.round(4),
                "% Variance": (explained * 100).round(2),
                "% Cumulé": (np.cumsum(explained) * 100).round(2)
            })
            
            self.show_table(self.afc_table_frame, eig_df, "Valeurs propres AFC")
            
        except Exception as e:
            self.update_status(f"❌ Erreur AFC: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans l'analyse AFC: {str(e)}")

    def afc_test_chi2_interpretation(self):
        """Test du χ² (scipy) + interprétation"""
        if not hasattr(self.ctx, "afc_table") or self.ctx.afc_table is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer le tableau de contingence.")
            return
        
        try:
            # 1. NETTOYER COMPLÈTEMENT L'AFFICHAGE
            for widget in self.afc_canvas_frame.winfo_children():
                widget.destroy()
            for widget in self.afc_table_frame.winfo_children():
                widget.destroy()
            
            # 2. CRÉER UN FRAME SCROLLABLE POUR TOUT LE CONTENU
            canvas = tk.Canvas(self.afc_table_frame, bg=COLORS["white"])
            scrollbar = ttk.Scrollbar(self.afc_table_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg=COLORS["white"])
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # 3. CALCULER LES RÉSULTATS
            N = self.ctx.afc_table.values.astype(float)
            chi2, pval, dof, expected = chi2_contingency(N)
            
            # Interprétation
            alpha = 0.05
            reject = (pval < alpha)
            interpretation = "Dépendance significative (rejet de H0)" if reject else "Pas de dépendance significative (non-rejet de H0)"
            conclusion_color = COLORS["success"] if reject else COLORS["warning"]
            
            # 4. AFFICHER LES RÉSULTATS DU TEST
            result_title = tk.Label(
                scrollable_frame, 
                text="Test d'Indépendance du χ²", 
                font=FONT_SUBTITLE,
                bg=COLORS["white"], 
                fg=COLORS["primary"]
            )
            result_title.pack(pady=(10, 10))
            
            # Créer un DataFrame avec les résultats
            result_df = pd.DataFrame({
                "Statistique": ["Chi-deux (χ²)", "Degrés de liberté", "p-value", "Seuil α"],
                "Valeur": [f"{chi2:.4f}", f"{dof}", f"{pval:.6f}", f"{alpha}"]
            })
            
            # Afficher le premier tableau (résultats)
            self._create_simple_table(
                scrollable_frame, 
                result_df, 
                "Résultats du test"
            )
            
            # Afficher l'interprétation
            interp_frame = tk.Frame(scrollable_frame, bg=COLORS["light"], relief='groove', borderwidth=2)
            interp_frame.pack(fill='x', pady=20, padx=10)
            
            interp_label = tk.Label(
                interp_frame, 
                text=f"Interprétation: {interpretation}", 
                font=FONT_TEXT_BOLD,
                bg=COLORS["light"], 
                fg=conclusion_color,
                wraplength=800
            )
            interp_label.pack(pady=10, padx=10)
            
            # 5. AFFICHER LES EFFECTIFS ATTENDUS
            exp_title = tk.Label(
                scrollable_frame, 
                text="Effectifs attendus sous H0", 
                font=FONT_SUBTITLE,
                bg=COLORS["white"], 
                fg=COLORS["primary"]
            )
            exp_title.pack(pady=(20, 10))
            
            exp_df = pd.DataFrame(
                expected, 
                index=self.ctx.afc_table.index, 
                columns=self.ctx.afc_table.columns
            ).round(3)
            
            # Afficher le deuxième tableau (effectifs attendus)
            self._create_simple_table(
                scrollable_frame, 
                exp_df, 
                "Tableau des effectifs attendus"
            )
            
            # 6. BOUTONS D'EXPORT
            btn_frame = tk.Frame(scrollable_frame, bg=COLORS["white"])
            btn_frame.pack(fill='x', pady=20)
            
            ttk.Button(
                btn_frame, 
                text="📥 Exporter tous les résultats", 
                command=lambda: self._export_chi2_all(chi2, pval, dof, interpretation, exp_df)
            ).pack(side='right', padx=5)
            
            self.show_page("afc")
            self.update_status(f"✅ Test khi-deux terminé: {interpretation}")
            
        except Exception as e:
            self.update_status(f"❌ Erreur dans le test χ²: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans le test χ²: {str(e)}")
    
    def _create_simple_table(self, parent_frame, df, title):
        """Crée un tableau simple sans effacer le parent"""
        # Frame pour ce tableau spécifique
        table_frame = tk.Frame(parent_frame, bg=COLORS["white"])
        table_frame.pack(fill='x', pady=(0, 10))
        
        # Titre
        title_label = tk.Label(
            table_frame, 
            text=title, 
            font=FONT_TEXT_BOLD,
            bg=COLORS["white"], 
            fg=COLORS["primary"]
        )
        title_label.pack(pady=(0, 5))
        
        # Créer le Treeview
        tree_frame = tk.Frame(table_frame, bg=COLORS["white"])
        tree_frame.pack(fill='x')
        
        # Colonnes
        if isinstance(df, pd.DataFrame):
            cols = list(df.columns)
        else:
            cols = ["Colonne"]
        
        tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=min(8, len(df)))
        
        # Configurer les colonnes
        for col in cols:
            tree.heading(col, text=str(col))
            tree.column(col, width=120, anchor="center")
        
        # Ajouter les données
        if isinstance(df, pd.DataFrame):
            if df.index.name is not None:
                # Inclure l'index comme première colonne
                all_cols = [df.index.name] + cols
                tree["columns"] = all_cols
                for col in all_cols:
                    tree.heading(col, text=str(col))
                    tree.column(col, width=120, anchor="center")
                
                for idx, row in df.iterrows():
                    tree.insert("", "end", values=[idx] + list(row))
            else:
                for idx, row in df.iterrows():
                    tree.insert("", "end", values=list(row))
        else:
            # Si ce n'est pas un DataFrame
            tree.insert("", "end", values=[str(df)])
        
        tree.pack(fill='x')
    
    def _export_chi2_all(self, chi2, pval, dof, interpretation, exp_df):
        """Exporte tous les résultats du test khi-deux"""
        # Créer un DataFrame avec les résultats
        results_df = pd.DataFrame({
            "Statistique": ["Chi-deux (χ²)", "Degrés de liberté", "p-value", "Interprétation"],
            "Valeur": [f"{chi2:.4f}", f"{dof}", f"{pval:.6f}", interpretation]
        })
        
        # Demander où sauvegarder
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("Tous", "*.*")],
            initialfile=f"khi2_results_{now_tag()}.xlsx"
        )
        
        if path:
            try:
                # Écrire dans un fichier Excel avec deux onglets
                with pd.ExcelWriter(path, engine='openpyxl') as writer:
                    results_df.to_excel(writer, sheet_name='Résultats', index=False)
                    exp_df.to_excel(writer, sheet_name='Effectifs attendus')
                
                self.update_status(f"✅ Résultats exportés: {path}")
                messagebox.showinfo("Export", "Résultats exportés avec succès!")
            except Exception as e:
                self.update_status(f"❌ Erreur d'export: {str(e)}")
                messagebox.showerror("Erreur", f"Erreur lors de l'export: {str(e)}")
    # ======================================================
    # 🛡️ MÉTHODES CYBERSÉCURITÉ
    # ======================================================
    
    
    def cyber_import_data(self):
        """Importe un dataset pour la cybersécurité (sans casser les colonnes texte)."""
        path = filedialog.askopenfilename(
            title="Choisir un fichier pour la cybersécurité",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv"), ("Tous", "*.*")]
        )

        if not path:
            return

        try:
            if path.lower().endswith(".csv"):
                df = pd.read_csv(path)
            else:
                df = pd.read_excel(path)

            # Détecter colonnes numériques sans convertir les colonnes texte (ex: Bibliothèque)
            num_cols = []
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    num_cols.append(col)
                else:
                    tmp = pd.to_numeric(df[col].astype(str).str.replace(",", ".", regex=False), errors="coerce")
                    if tmp.notna().mean() > 0.8:
                        df[col] = tmp
                        num_cols.append(col)

            if len(num_cols) < 2:
                raise ValueError("Au moins 2 colonnes numériques sont requises.")

            # Nettoyer: ne dropna que sur les colonnes numériques
            df = df.dropna(subset=num_cols).reset_index(drop=True)

            # Stocker dans le contexte
            self.ctx.cyber_df = df
            self.ctx.cyber_num_cols = num_cols
            self.ctx.cyber_X = df[num_cols].values.astype(float)

            # Affichage
            self.show_table(self.cyber_table_frame, df.head(20), "Données Cybersécurité (20 premières lignes)")
            self.show_page("cyber")
            self.update_status(f"✅ Dataset cybersécurité importé: {os.path.basename(path)}")

        except Exception as e:
            self.update_status(f"❌ Erreur d'import: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur lors de l'import: {str(e)}")

    def cyber_isolation_forest(self):
        """Exécute et affiche l'Isolation Forest"""
        if not hasattr(self.ctx, 'cyber_df') or self.ctx.cyber_df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer un dataset cybersécurité.")
            return
        
        try:
            X = self.ctx.cyber_X
            iso = IsolationForest(contamination=0.08, random_state=42, n_estimators=200)
            pred = iso.fit_predict(X)
            
            # Stocker les résultats
            self.ctx.cyber_iso = iso
            self.ctx.cyber_pred_iso = pred
            
            # Réduction de dimension pour la visualisation
            Z = PCA(n_components=2, random_state=42).fit_transform(StandardScaler().fit_transform(X))
            
            # Identifier les noms des individus (première colonne non numérique)
            id_col = None
            for col in self.ctx.cyber_df.columns:
                if col not in self.ctx.cyber_num_cols:
                    id_col = col
                    break
            
            if id_col:
                ids = self.ctx.cyber_df[id_col].astype(str)
            else:
                ids = self.ctx.cyber_df.index.astype(str)
            
            # Créer le graphique
            fig = Figure(figsize=(11, 9), dpi=100)
            ax = fig.add_subplot(111)
            
            # Séparer normal et anomalies
            normal_mask = pred == 1
            anomaly_mask = pred == -1
            
            # Points normaux
            if np.any(normal_mask):
                ax.scatter(Z[normal_mask, 0], Z[normal_mask, 1], 
                          s=60, alpha=0.7, label='Normal',
                          color=COLORS["success"], edgecolors='white', linewidth=1)
            
            # Anomalies
            if np.any(anomaly_mask):
                anomaly_scatter = ax.scatter(Z[anomaly_mask, 0], Z[anomaly_mask, 1], 
                                            s=100, alpha=0.9, label='Anomalie',
                                            color=COLORS["error"], edgecolors='white', linewidth=1.5)
                
                # Annoter les anomalies
                for i in np.where(anomaly_mask)[0]:
                    ax.annotate(ids.iloc[i], (Z[i, 0], Z[i, 1]), 
                              fontsize=8, fontweight='bold',
                              xytext=(5, 5), textcoords='offset points',
                              bbox=dict(boxstyle='round,pad=0.3', alpha=0.3, facecolor='white'))
            
            ax.set_xlabel("Dimension 1", fontweight='bold')
            ax.set_ylabel("Dimension 2", fontweight='bold')
            ax.set_title("Isolation Forest - Détection d'Anomalies", 
                        fontsize=14, fontweight='bold')
            ax.legend(loc='best')
            ax.grid(True, alpha=0.3)
            
            fig.tight_layout()
            self.show_plot(self.cyber_canvas_frame, fig)
            self.show_page("cyber")
            
            # Afficher les scores d'anomalie
            scores = -iso.score_samples(X)
            df_out = self.ctx.cyber_df.copy()
            df_out["AnomalyScore"] = scores
            df_out["IF_Label"] = np.where(pred == -1, "Anomalie", "Normal")
            
            # Trier par score
            top_anomalies = df_out.sort_values("AnomalyScore", ascending=False).head(15)
            
            # Sélectionner les colonnes à afficher
            display_cols = []
            if id_col:
                display_cols.append(id_col)
            display_cols += self.ctx.cyber_num_cols[:3]  # Premières 3 variables numériques
            display_cols += ["AnomalyScore", "IF_Label"]
            
            self.show_table(self.cyber_table_frame, 
                          top_anomalies[display_cols].round(4), 
                          "Top 15 Anomalies - Isolation Forest")
            
        except Exception as e:
            self.update_status(f"❌ Erreur Isolation Forest: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans Isolation Forest: {str(e)}")
    
    def cyber_lof(self):
        """Exécute et affiche le LOF avec noms des individus"""
        if not hasattr(self.ctx, 'cyber_df') or self.ctx.cyber_df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer un dataset cybersécurité.")
            return
        
        try:
            X = self.ctx.cyber_X
            Xs = StandardScaler().fit_transform(X)
            
            lof = LocalOutlierFactor(n_neighbors=20, contamination=0.08)
            pred = lof.fit_predict(Xs)
            
            # Stocker les résultats
            self.ctx.cyber_lof = lof
            self.ctx.cyber_pred_lof = pred
            
            # Réduction de dimension
            Z = PCA(n_components=2, random_state=42).fit_transform(Xs)
            
            # Identifier les noms des individus
            id_col = None
            for col in self.ctx.cyber_df.columns:
                if col not in self.ctx.cyber_num_cols:
                    id_col = col
                    break
            
            if id_col:
                ids = self.ctx.cyber_df[id_col].astype(str)
            else:
                ids = self.ctx.cyber_df.index.astype(str)
            
            # Créer le graphique
            fig = Figure(figsize=(11, 9), dpi=100)
            ax = fig.add_subplot(111)
            
            # Séparer normal et anomalies
            normal_mask = pred == 1
            anomaly_mask = pred == -1
            
            # Points normaux
            if np.any(normal_mask):
                ax.scatter(Z[normal_mask, 0], Z[normal_mask, 1], 
                          s=60, alpha=0.7, label='Normal',
                          color=COLORS["success"], edgecolors='white', linewidth=1)
            
            # Anomalies
            if np.any(anomaly_mask):
                anomaly_scatter = ax.scatter(Z[anomaly_mask, 0], Z[anomaly_mask, 1], 
                                            s=100, alpha=0.9, label='Anomalie',
                                            color=COLORS["error"], edgecolors='white', linewidth=1.5)
                
                # Annoter TOUTES les anomalies avec leurs noms
                for i in np.where(anomaly_mask)[0]:
                    ax.annotate(ids.iloc[i], (Z[i, 0], Z[i, 1]), 
                              fontsize=8, fontweight='bold',
                              xytext=(5, 5), textcoords='offset points',
                              bbox=dict(boxstyle='round,pad=0.3', alpha=0.3, facecolor='white'))
            
            ax.set_xlabel("Dimension 1", fontweight='bold')
            ax.set_ylabel("Dimension 2", fontweight='bold')
            ax.set_title("Local Outlier Factor (LOF) - Détection d'Anomalies", 
                        fontsize=14, fontweight='bold')
            ax.legend(loc='best')
            ax.grid(True, alpha=0.3)
            
            fig.tight_layout()
            self.show_plot(self.cyber_canvas_frame, fig)
            self.show_page("cyber")
            
            # Afficher les scores LOF
            scores = -lof.negative_outlier_factor_
            df_out = self.ctx.cyber_df.copy()
            df_out["AnomalyScore_LOF"] = scores
            df_out["LOF_Label"] = np.where(pred == -1, "Anomalie", "Normal")
            
            # Trier par score
            top_anomalies = df_out.sort_values("AnomalyScore_LOF", ascending=False).head(15)
            
            # Sélectionner les colonnes à afficher
            display_cols = []
            if id_col:
                display_cols.append(id_col)
            display_cols += self.ctx.cyber_num_cols[:3]
            display_cols += ["AnomalyScore_LOF", "LOF_Label"]
            
            self.show_table(self.cyber_table_frame, 
                          top_anomalies[display_cols].round(4), 
                          "Top 15 Anomalies - LOF")
            
        except Exception as e:
            self.update_status(f"❌ Erreur LOF: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans LOF: {str(e)}")
    
    def cyber_interpretation_actions(self):
        """Affiche l'interprétation et les actions recommandées"""
        if not hasattr(self.ctx, 'cyber_df') or self.ctx.cyber_df is None:
            messagebox.showwarning("Erreur", "Veuillez d'abord importer un dataset cybersécurité.")
            return
        
        try:
            # Identifier les individus à risque
            risk_mask = np.zeros(len(self.ctx.cyber_df), dtype=bool)
            
            if hasattr(self.ctx, 'cyber_pred_iso') and self.ctx.cyber_pred_iso is not None:
                risk_mask |= (self.ctx.cyber_pred_iso == -1)
            
            if hasattr(self.ctx, 'cyber_pred_lof') and self.ctx.cyber_pred_lof is not None:
                risk_mask |= (self.ctx.cyber_pred_lof == -1)
            
            # Créer le dataframe de résultats
            df_out = self.ctx.cyber_df.copy()
            
            # Identifier les colonnes
            id_col = None
            for col in df_out.columns:
                if col not in self.ctx.cyber_num_cols:
                    id_col = col
                    break
            
            # Ajouter les colonnes de risque et d'actions
            df_out["Risque"] = np.where(risk_mask, "Élevé", "Normal")
            
            actions = []
            for i in range(len(df_out)):
                if risk_mask[i]:
                    actions.append("🔴 Isoler - Analyser logs - Changer credentials - Scanner malware")
                else:
                    actions.append("🟢 Surveillance normale - Vérifier logs périodiquement")
            
            df_out["Actions Recommandées"] = actions
            
            # Filtrer les individus à haut risque
            high_risk = df_out[df_out["Risque"] == "Élevé"]
            
            if len(high_risk) == 0:
                messagebox.showinfo("Résultat", "✅ Aucune anomalie détectée dans le dataset.")
                return
            
            # Sélectionner les colonnes à afficher
            display_cols = []
            if id_col:
                display_cols.append(id_col)
            
            # Ajouter quelques variables numériques importantes
            if hasattr(self.ctx, 'cyber_iso'):
                # Utiliser les features importantes de l'Isolation Forest
                display_cols += self.ctx.cyber_num_cols[:min(3, len(self.ctx.cyber_num_cols))]
            
            display_cols += ["Risque", "Actions Recommandées"]
            
            # Afficher le tableau
            self.show_table(self.cyber_table_frame, 
                          high_risk[display_cols].head(20), 
                          f"Individus à Haut Risque ({len(high_risk)} détectés)")
            
            self.show_page("cyber")
            
            # Créer un graphique de synthèse
            fig = Figure(figsize=(10, 6), dpi=100)
            ax = fig.add_subplot(111)
            
            # Statistiques
            total = len(df_out)
            high_risk_count = len(high_risk)
            normal_count = total - high_risk_count
            
            labels = ['Normal', 'Haut Risque']
            sizes = [normal_count, high_risk_count]
            colors = [COLORS["success"], COLORS["error"]]
            explode = (0, 0.1)
            
            wedges, texts, autotexts = ax.pie(sizes, explode=explode, labels=labels, colors=colors,
                                             autopct='%1.1f%%', shadow=True, startangle=90)
            
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
            
            ax.set_title(f"Répartition des Risques\n{high_risk_count}/{total} anomalies détectées", 
                        fontsize=14, fontweight='bold')
            
            fig.tight_layout()
            self.show_plot(self.cyber_canvas_frame, fig)
            
        except Exception as e:
            self.update_status(f"❌ Erreur d'interprétation: {str(e)}")
            messagebox.showerror("Erreur", f"Erreur dans l'interprétation: {str(e)}")


# ======================================================
# 🗂️ CLASSE DE CONTEXTE DE DONNÉES
# ======================================================

class DataContext:
    def __init__(self):
        self.data_path = DEFAULT_EXCEL
        self.df = None
        self.id_col = None
        self.num_cols = []
        self.X = None
        self.X_scaled = None
        self.scaler = None
        
        # PCA
        self.pca = None
        self.scores = None
        self.explained = None
        self.eigvals = None
        self.corr_axes = None
        
        # Clustering / ML
        self.k = 4
        self.kmeans = None
        self.cluster_labels = None
        self.rf = None
        self.X_train = self.X_test = self.y_train = self.y_test = None
        
        # AFC
        self.afc_table = None
        self.afc_rows = None
        self.afc_cols = None
        self.afc_result = None
        
        # Cybersécurité
        self.cyber_df = None
        self.cyber_num_cols = None
        self.cyber_X = None
        self.cyber_iso = None
        self.cyber_lof = None
        self.cyber_pred_iso = None
        self.cyber_pred_lof = None
    
    
    def load(self, path: str | None = None):
        if path:
            self.data_path = path
        
        if not os.path.exists(self.data_path):
            raise FileNotFoundError(f"Fichier introuvable: {self.data_path}")
        
        df = safe_read_excel(self.data_path)
        
        # Colonne ID: préférer 'Bibliothèque'
        if "Bibliothèque" in df.columns:
            self.id_col = "Bibliothèque"
        else:
            non_num = [c for c in df.columns if not pd.api.types.is_numeric_dtype(df[c])]
            self.id_col = non_num[0] if non_num else None
        
        self.df = df.copy()
        self.num_cols = numeric_columns(self.df)
        
        if self.id_col in self.num_cols:
            self.num_cols = [c for c in self.num_cols if c != self.id_col]
        
        if len(self.num_cols) < 2:
            raise ValueError("Il faut au moins 2 variables numériques pour l'ACP/Clustering.")
        
        self.X = self.df[self.num_cols].copy()
        self.df = self.df.dropna(subset=self.num_cols).reset_index(drop=True)
        self.X = self.df[self.num_cols].copy()
        
        self.scaler = StandardScaler()
        self.X_scaled = self.scaler.fit_transform(self.X)
        
        self._compute_pca()
        self._compute_kmeans_and_rf()
    
    def _compute_pca(self):
        p = self.X_scaled.shape[1]
        self.pca = PCA(n_components=p, random_state=42)
        self.scores = self.pca.fit_transform(self.X_scaled)
        self.explained = self.pca.explained_variance_ratio_
        self.eigvals = self.pca.explained_variance_
        
        corr = np.corrcoef(self.X_scaled.T, self.scores.T)
        self.corr_axes = corr[:p, p:]
    
    def _compute_kmeans_and_rf(self):
        scores_2 = self.scores[:, :2]
        self.kmeans = KMeans(n_clusters=self.k, random_state=42, n_init=10)
        self.cluster_labels = self.kmeans.fit_predict(scores_2)
        self.df["Cluster"] = [f"Cluster {c+1}" for c in self.cluster_labels]
        
        X_rf = self.X.copy()
        y = self.df["Cluster"]
        self.X_train, self.X_test, self.y_train, self.y_test = train_test_split(
            X_rf, y, test_size=0.25, random_state=42, stratify=y
        )
        self.rf = RandomForestClassifier(n_estimators=200, random_state=42)
        self.rf.fit(self.X_train, self.y_train)
    
    def set_k(self, k: int):
        if not (3 <= k <= 7):
            raise ValueError("k doit être entre 3 et 7.")
        self.k = k
        self._compute_kmeans_and_rf()


# ======================================================
# 🚀 POINT D'ENTRÉE DE L'APPLICATION
# ======================================================

if __name__ == "__main__":
    root = tk.Tk()
    app = ModernApp(root)
    root.mainloop()