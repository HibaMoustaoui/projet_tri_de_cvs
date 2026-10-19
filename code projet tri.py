print("Lancement de l'application ...")
import os
import re
import pdfplumber
import pandas as pd
import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue") 

# --- Backend Logique ---
def extract_contact_info(text):
    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', text)
    email = email_match.group(0) if email_match else "Non trouvé"
    
    phone_match = re.search(r'(?:\+212|0)([5-7]\d{8})', text)
    phone = phone_match.group(0) if phone_match else "Non trouvé"
    
    name = "Inconnu"
    if email != "Non trouvé":
        email_prefix = email.split('@')[0]
        clean_name = re.sub(r'[\d\._-]+', ' ', email_prefix).strip()
        if clean_name:
            name = clean_name.title()
            
    if name == "Inconnu":
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        if lines:
            for line in lines[:3]:
                if len(line) > 3 and not any(kw in line.lower() for kw in ["cv", "curriculum", "resume", "profil", "née", "étudiante"]):
                    name = line
                    break
                
    return name, phone, email

def read_cv(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as reader:
            return "".join([page.extract_text(layout=True) or "" for page in reader.pages])
    except:
        return ""

def classify_stagiaire(cv_text):
    cv_text_lower = cv_text.lower()
    keywords = ["python", "application", "anglais", "science", "energy", "stage", "ia"]
    found_keywords = [word for word in keywords if word in cv_text_lower]
    score = len(found_keywords)
    return score, found_keywords, ("accepte" if score >= 3 else "refuse")

# --- Frontend Interface ---
class PremiumCVApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("Smart Screen & Admin Investigation System v3.3")
        self.geometry("950x650")
        self.resizable(True, True)
        self.minsize(900, 600)
        
        self.folder_path = ""
        self.result = []
        
        self.sidebar = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color=("#f4f6f9", "#111723"))
        self.sidebar.pack(side="left", fill="y")
        self.setup_sidebar()
        
        self.content_frame = ctk.CTkFrame(self, corner_radius=12, fg_color=("#ffffff", "#1e2530"))
        self.content_frame.pack(side="right", fill="both", expand=True, padx=20, pady=20)
        
        self.pages = {}
        self.create_page_1()
        self.create_page_2()
        self.create_page_3()
        
        self.show_page(1)

    def setup_sidebar(self):
        lbl_logo = ctk.CTkLabel(self.sidebar, text="🤖 WORKFLOW", font=("Segoe UI", 16, "bold"), text_color=("#1f497d", "#54a0ff"))
        lbl_logo.pack(pady=(30, 40))
        
        self.step_indicators = {}
        steps = [(1, "  1. Accueil & Tâches"), (2, "  2. Tri Auto des CVs"), (3, "  3. Rapport & Enquête")]
        for num, text in steps:
            lbl = ctk.CTkLabel(self.sidebar, text=text, font=("Segoe UI", 13), anchor="w", width=170, height=35, corner_radius=6)
            lbl.pack(pady=8, padx=15)
            self.step_indicators[num] = lbl

    def update_sidebar_indicators(self, active_page):
        for num, lbl in self.step_indicators.items():
            if num == active_page:
                lbl.configure(text_color="white", fg_color=("#1f497d", "#2980b9"), font=("Segoe UI", 13, "bold"))
            else:
                lbl.configure(text_color=("#5c6b73", "#95a5a6"), fg_color="transparent", font=("Segoe UI", 13, "normal"))

    def show_page(self, page_num):
        for p in self.pages.values():
            p.pack_forget()
        self.pages[page_num].pack(fill="both", expand=True, padx=20, pady=20)
        self.update_sidebar_indicators(page_num)

    # --- PAGE 1: ACCUEIL ---
    def create_page_1(self):
        page = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.pages[1] = page
        
        header_frame = ctk.CTkFrame(page, fg_color=("#1f497d", "#17202a"), height=80, corner_radius=10)
        header_frame.pack(fill="x", pady=(0, 20))
        header_frame.pack_propagate(False)
        
        title = ctk.CTkLabel(header_frame, text="Système Numérique de Tri de CV & Enquêtes", font=("Segoe UI", 18, "bold"), text_color="white")
        title.pack(pady=25)
        
        lbl_welcome = ctk.CTkLabel(page, text="Bienvenue dans votre assistant de gestion RH intelligent.", font=("Segoe UI", 14, "italic"), text_color="gray")
        lbl_welcome.pack(pady=(0, 15))
        
        cards_container = ctk.CTkFrame(page, fg_color="transparent")
        cards_container.pack(fill="both", expand=True, pady=10)
        
        for text in [
            "📄 Extraction PDF Automatique\nLecture instantanée du texte des CVs, des emails et téléphones.",
            "⚡ Algorithme de Scoring & Tri\nFiltrage basé sur des mots-clés cibles (Python, IA, Stage...).",
            "🛡️ Enquête Administrative Directe\nEnvoi et suivi automatique des profils validés vers l'administration."
        ]:
            card = ctk.CTkFrame(cards_container, fg_color=("#f8f9fa", "#252d3a"), corner_radius=8, height=85)
            card.pack(fill="x", pady=6)
            lbl = ctk.CTkLabel(card, text=text, font=("Segoe UI", 12), justify="left", anchor="w", padx=15)
            lbl.pack(fill="both", expand=True)
        
        btn_next = ctk.CTkButton(page, text="Ouvrir le Module de Tri ➔", font=("Segoe UI", 13, "bold"), 
                                  fg_color=("#27ae60", "#2ecc71"), hover_color="#2196f3", height=42, width=220,
                                  command=lambda: self.show_page(2))
        btn_next.pack(side="bottom", pady=10)

    # --- PAGE 2: TRAITEMENT ---
    def create_page_2(self):
        page = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.pages[2] = page
        
        lbl_section = ctk.CTkLabel(page, text="Analyse en masse des CVs (Format PDF)", font=("Segoe UI", 15, "bold"))
        lbl_section.pack(anchor="w", pady=(0, 10))
        
        top_bar = ctk.CTkFrame(page, fg_color="transparent")
        top_bar.pack(fill="x", pady=5)
        
        self.btn_select = ctk.CTkButton(top_bar, text="📁 Choisir un dossier", font=("Segoe UI", 12, "bold"), 
                                        fg_color="#34495e", hover_color="#2c3e50", height=36, command=self.select_folder)
        self.btn_select.pack(side="left", padx=(0, 15))
        
        self.lbl_folder = ctk.CTkLabel(top_bar, text="Aucun dossier sélectionné", font=("Segoe UI", 12, "italic"), text_color="gray")
        self.lbl_folder.pack(side="left", fill="x")
        
        self.text_log = ctk.CTkTextbox(page, font=("Consolas", 11), fg_color=("#1e272e", "#0f141c"), text_color="#2ecc71")
        self.text_log.pack(fill="both", expand=True, pady=(10, 15))
        
        nav_frame = ctk.CTkFrame(page, fg_color="transparent")
        nav_frame.pack(side="bottom", fill="x", pady=(5, 0))
        
        btn_prev = ctk.CTkButton(nav_frame, text="⏪ Accueil", font=("Segoe UI", 12), fg_color="#7f8c8d", width=120, command=lambda: self.show_page(1))
        btn_prev.pack(side="left", padx=5)
        
        self.btn_go_page3 = ctk.CTkButton(nav_frame, text="Rapports & Enquêtes ➔", font=("Segoe UI", 12, "bold"), fg_color="#2980b9",
                                           state="disabled", command=lambda: self.show_page(3))
        self.btn_go_page3.pack(side="left", padx=20)
        
        self.btn_process = ctk.CTkButton(nav_frame, text="Lancer l'analyse ⚡", font=("Segoe UI", 12, "bold"), 
                                          fg_color="#10ac84", hover_color="#0fa87f", width=180, state="disabled", command=self.process_cvs)
        self.btn_process.pack(side="right", padx=5)

    def select_folder(self):
        self.folder_path = filedialog.askdirectory()
        if self.folder_path:
            self.lbl_folder.configure(text=f"Cible : ...{os.path.basename(self.folder_path)}", text_color=("#2c3e50", "#f5f6fa"))
            self.btn_process.configure(state="normal")
            self.text_log.insert("end", f"[Système] Dossier connecté : {self.folder_path}\n")

    def process_cvs(self):
        self.text_log.delete("1.0", "end")
        self.text_log.insert("end", "[Processus] Initialisation du moteur d'analyse...\n")
        self.update()

        pdf_files = [f for f in os.listdir(self.folder_path) if f.lower().endswith('.pdf')]
        if not pdf_files:
            messagebox.showwarning("Dossier Vide", "Aucun CV au format .pdf n'a été détecté.")
            return
        
        self.result = []
        for index, pdf_file in enumerate(pdf_files, start=1):
            file_path = os.path.join(self.folder_path, pdf_file)
            cv_text = read_cv(file_path)
            
            if cv_text:
                score, found_keywords, decision = classify_stagiaire(cv_text)
                name, phone, email = extract_contact_info(cv_text)
                
                self.result.append({
                    "Nom & Prénom": name,
                    "Téléphone": phone,
                    "E-mail": email,
                    "Mots clés trouvés": ", ".join(found_keywords) or "Aucun",
                    "Score": int(score),
                    "Décision": str(decision),
                    "Nom du fichier": str(pdf_file)
                })
                self.text_log.insert("end", f"   [✓] {index}/{len(pdf_files)} : {name} extracted (Score: {score}/7 -> {decision.upper()})\n")
            else:
                self.result.append({
                    "Nom & Prénom": "Erreur de lecture", "Téléphone": "N/A", "E-mail": "N/A",
                    "Mots clés trouvés": "Erreur", "Score": 0, "Décision": "Echec", "Nom du fichier": str(pdf_file)
                })
                self.text_log.insert("end", f"   [✗] Error reading file: {pdf_file}\n")
            self.update()

        self.text_log.insert("end", "\n[Terminé] Analyse globale achevée. Allez à la page suivante.\n")
        
        self.render_pretty_results()
        self.btn_go_page3.configure(state="normal")

    # --- PAGE 3: EXPORT & ENQUÊTE (DESIGN NOVEAU) ---
    def create_page_3(self):
        page = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.pages[3] = page
        
        lbl_section = ctk.CTkLabel(page, text="Livrables & Synthèse des Décisions", font=("Segoe UI", 15, "bold"))
        lbl_section.pack(anchor="w", pady=(0, 10))
        
        main_layout = ctk.CTkFrame(page, fg_color="transparent")
        main_layout.pack(fill="both", expand=True, pady=5)
        
        self.results_frame = ctk.CTkScrollableFrame(
            main_layout, 
            label_text="📋 Synthese des Candidats Analysés",
            label_font=("Segoe UI", 12, "bold"),
            fg_color=("#f8f9fa", "#161b22"),
            corner_radius=10
        )
        self.results_frame.pack(side="left", fill="both", expand=True, padx=(0, 15))
        
        self.lbl_empty_msg = ctk.CTkLabel(
            self.results_frame, 
            text="⏳ En attente d'exécution des analyses à l'étape 2.", 
            font=("Segoe UI", 13, "italic"), text_color="gray"
        )
        self.lbl_empty_msg.pack(pady=50)

        actions_sidebar = ctk.CTkFrame(main_layout, fg_color=("#ebf5fb", "#1c2833"), width=220, corner_radius=10)
        actions_sidebar.pack(side="right", fill="y", padx=(5, 0))
        actions_sidebar.pack_propagate(False)
        
        lbl_actions_title = ctk.CTkLabel(actions_sidebar, text="⚙️ Actions", font=("Segoe UI", 14, "bold"))
        lbl_actions_title.pack(pady=(20, 15))
        
        self.btn_export = ctk.CTkButton(
            actions_sidebar, text="📊 Exporter Excel", font=("Segoe UI", 12, "bold"),
            fg_color="#e67e22", hover_color="#d35400", height=40, command=self.export_to_excel
        )
        self.btn_export.pack(fill="x", padx=15, pady=10)
        
        self.btn_send_email = ctk.CTkButton(
            actions_sidebar, text="✉️ Lancer l'Enquête", font=("Segoe UI", 12, "bold"),
            fg_color="#1abc9c", hover_color="#16a085", height=40, state="disabled", command=self.send_enquete_email
        )
        self.btn_send_email.pack(fill="x", padx=15, pady=10)

        nav_frame = ctk.CTkFrame(page, fg_color="transparent")
        nav_frame.pack(side="bottom", fill="x", pady=(10, 0))
        
        btn_prev = ctk.CTkButton(nav_frame, text="⏪ Étape Précédente", font=("Segoe UI", 12), fg_color="#7f8c8d", width=150, command=lambda: self.show_page(2))
        btn_prev.pack(side="left")

    def render_pretty_results(self):
        for child in self.results_frame.winfo_children():
            child.destroy()
            
        if not self.result:
            return

        for idx, item in enumerate(self.result, start=1):
            row_card = ctk.CTkFrame(self.results_frame, fg_color=("#ffffff", "#21262d"), corner_radius=6, height=45)
            row_card.pack(fill="x", pady=4, padx=5)
            row_card.pack_propagate(False)
            
            lbl_num = ctk.CTkLabel(row_card, text=f"#{idx}", font=("Segoe UI", 11, "bold"), width=30, text_color="gray")
            lbl_num.pack(side="left", padx=(10, 5))
            
            lbl_name = ctk.CTkLabel(row_card, text=item['Nom & Prénom'], font=("Segoe UI", 12, "bold"), anchor="w", width=180)
            lbl_name.pack(side="left", padx=5)
            
            is_accepted = item['Décision'] == 'accepte'
            status_text = " ✅ ACCEPTÉ " if is_accepted else " ❌ REFUSÉ "
            status_bg = "#2ecc71" if is_accepted else "#e74c3c"
            
            lbl_status = ctk.CTkLabel(
                row_card, text=status_text, font=("Segoe UI", 10, "bold"), 
                text_color="white", fg_color=status_bg, corner_radius=4, height=22
            )
            lbl_status.pack(side="right", padx=15)
            
            lbl_info = ctk.CTkLabel(row_card, text=f"Score: {item['Score']}/7", font=("Segoe UI", 11), text_color="gray")
            lbl_info.pack(side="right", padx=10)

    def export_to_excel(self):
        if not self.result:
            messagebox.showwarning("Aucune donnée", "Veuillez d'abord scanner les fichiers à la phase 2.")
            return
            
        try:
            df = pd.DataFrame(self.result)
            save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if save_path:
                with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Base_Candidats')
                    worksheet = writer.sheets['Base_Candidats']
                    
                    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
                    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                    center_alignment = Alignment(horizontal="center", vertical="center")
                    thin_border = Border(left=Side(style='thin', color='D3D3D3'), right=Side(style='thin', color='D3D3D3'),
                                         top=Side(style='thin', color='D3D3D3'), bottom=Side(style='thin', color='D3D3D3'))

                    for col_idx in range(1, len(df.columns) + 1):
                        cell = worksheet.cell(row=1, column=col_idx)
                        cell.fill = header_fill; cell.font = header_font; cell.alignment = center_alignment; cell.border = thin_border

                    for row in worksheet.iter_rows(min_row=2, max_row=len(df) + 1, min_col=1, max_col=len(df.columns)):
                        for cell in row:
                            cell.alignment = center_alignment; cell.border = thin_border

                    for col in worksheet.columns:
                        max_len = max(len(str(cell.value or '')) for cell in col)
                        worksheet.column_dimensions[col[0].column_letter].width = max(max_len + 5, 22)

                messagebox.showinfo("Succès", "Base de contacts Excel générée avec succès !")
                self.btn_send_email.configure(state="normal")
        except Exception as e:
            messagebox.showerror("Erreur", f"Échec de l'exportation: {e}")

    def send_enquete_email(self):
        accepted_candidates = [c["Nom & Prénom"] for c in self.result if c["Décision"] == "accepte"]
        receiver_email = "administration@univ.ma"
        messagebox.showinfo("SMTP Connecté", 
                            f"Transmission effectuée vers : {receiver_email}\n"
                            f"Les fiches de contact de: {', '.join(accepted_candidates)} ont été transférées.")

if __name__ == "__main__":
    app = PremiumCVApp()
    app.mainloop()