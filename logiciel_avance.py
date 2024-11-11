import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox,PhotoImage
import pandas as pd

# Importation des fonctions de calcul depuis calcul.py
from calcul_avance import (
    calculer_et_tracer_gap_de_taux,
    calculer_et_tracer_gap_avec_couverture_neutre,
    impact_mni,
    calculer_van_par_choc,
    calculer_impact_sur_bilan,
    calculer_var_99
)

resultats = {}  
# Dictionnaire pour stocker les résultats pour l'exportation Excel

def create_gui():
    root = tk.Tk()
    root.title('ALM - RTIG')
    root.configure(bg='#f0f0f0')

    button_style = {'font': ('Helvetica', 12), 'bg': '#4CAF50', 'fg': 'white', 'padx': 10, 'pady': 5, 'relief': tk.RAISED}

    # Chargement de l'image de fond
    bg_image = PhotoImage(file="image2.png")

    # Création d'un Label pour afficher l'image et le positionner comme fond
    bg_label = tk.Label(root, image=bg_image)
    bg_label.place(x=0, y=0, relwidth=1, relheight=1)


    #Base de données
    global base_de_donnees
    base_de_donnees = None

    def load_file():
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            global base_de_donnees
            base_de_donnees = pd.read_excel(file_path)
            messagebox.showinfo("Information", "Fichier chargé avec succès!")

    def exporter_resultats():
        if resultats:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if file_path:
                with pd.ExcelWriter(file_path) as writer:
                    for nom, df in resultats.items():
                        df.to_excel(writer, sheet_name=nom)
                messagebox.showinfo("Exportation réussie", "Les résultats ont été exportés avec succès.")
        else:
            messagebox.showerror("Erreur", "Aucun résultat à exporter.")

    def run_calculer_et_tracer_gap_de_taux():
        if base_de_donnees is not None:
            gap_de_taux_annuel = calculer_et_tracer_gap_de_taux(base_de_donnees)
            resultats['Gap de Taux'] = pd.DataFrame({'Gap de Taux': gap_de_taux_annuel})
            messagebox.showinfo("Succès", "Calcul du gap de taux effectué.")
        else:
            messagebox.showerror("Erreur", "Veuillez charger un fichier de données d'abord.")

    def run_calculer_et_tracer_gap_avec_couverture_neutre():
        if base_de_donnees is not None:
            gap_de_taux_avant_couverture, gap_de_taux_apres_couverture, swaps_neutres = calculer_et_tracer_gap_avec_couverture_neutre(base_de_donnees)
            # Stockage des résultats pour l'exportation
            resultats['Gap avec Couverture Neutre'] = pd.DataFrame({
                'Gap Avant Couverture': gap_de_taux_avant_couverture,
                'Gap Après Couverture': gap_de_taux_apres_couverture,
                'Swaps Neutres': swaps_neutres
            })
            messagebox.showinfo("Succès", "Calcul du gap avec couverture neutre effectué.")
        else:
            messagebox.showerror("Erreur", "Veuillez charger un fichier de données d'abord.")
    
    def run_impact_mni():
        if base_de_donnees is not None:
            variation_taux = simpledialog.askfloat("Input", "Entrez la variation de taux d'intérêt (%):", minvalue=-100, maxvalue=100)
            if variation_taux is not None:
                impact_mni_result = impact_mni(base_de_donnees, variation_taux)
                # Stockage des résultats pour l'exportation
                resultats['Impact MNI'] = pd.DataFrame({'Impact MNI': impact_mni_result})
                messagebox.showinfo("Succès", "Calcul de l'impact MNI effectué.")
        else:
            messagebox.showerror("Erreur", "Veuillez charger un fichier de données d'abord.")
        
    def run_calculer_van_par_choc():
            if base_de_donnees is not None:
                choc_haut = simpledialog.askfloat("Input", "Entrez le choc haut (%):", minvalue=-100, maxvalue=100)
                choc_bas = simpledialog.askfloat("Input", "Entrez le choc bas (%):", minvalue=-100, maxvalue=100)
                delta_court = simpledialog.askfloat("Input", "Entrez le delta court (%):", minvalue=-100, maxvalue=100)
                delta_long = simpledialog.askfloat("Input", "Entrez le delta long (%):", minvalue=-100, maxvalue=100)
                if None not in (choc_haut, choc_bas, delta_court, delta_long):
                    van_par_choc = calculer_van_par_choc(base_de_donnees, choc_haut / 100, choc_bas / 100, delta_court / 100, delta_long / 100)
                    resultats['VAN par Choc'] = pd.DataFrame([van_par_choc], index=["Valeur"])
                    messagebox.showinfo("VAN par Choc", "Calcul de VAN par choc effectué.")
                else:
                    messagebox.showerror("Erreur", "Toutes les valeurs doivent être saisies.")
            else:
                messagebox.showerror("Erreur", "Veuillez charger un fichier de données d'abord.")
    
    def run_calculer_impact_sur_bilan():
        if base_de_donnees is not None:
            choc_haut = simpledialog.askfloat("Input", "Entrez le choc haut (%):", minvalue=-100, maxvalue=100)
            choc_bas = simpledialog.askfloat("Input", "Entrez le choc bas (%):", minvalue=-100, maxvalue=100)
            delta_court = simpledialog.askfloat("Input", "Entrez le delta court (%):", minvalue=-100, maxvalue=100)
            delta_long = simpledialog.askfloat("Input", "Entrez le delta long (%):", minvalue=-100, maxvalue=100)
            if None not in (choc_haut, choc_bas, delta_court, delta_long):
                impacts_sur_bilan = calculer_impact_sur_bilan(base_de_donnees, choc_haut / 100, choc_bas / 100, delta_court / 100, delta_long / 100)
                resultats['Impact sur Bilan'] = pd.DataFrame([impacts_sur_bilan], index=["Impact"])
                messagebox.showinfo("Impact sur Bilan", "Calcul de l'impact sur le bilan effectué.")
            else:
                messagebox.showerror("Erreur", "Toutes les valeurs doivent être saisies.")
        else:
            messagebox.showerror("Erreur", "Veuillez charger un fichier de données d'abord.")

    def run_calculer_var_99():
        if base_de_donnees is not None:
            delta_taux = simpledialog.askfloat("Input", "Entrez la variation de taux d'intérêt (%):", minvalue=-100, maxvalue=100)
            if delta_taux is not None:
                var_99_hausse, var_99_baisse, besoin_en_fonds_propres = calculer_var_99(base_de_donnees, delta_taux / 100)
                resultats['VaR 99'] = pd.DataFrame({
                    'VaR 99 Hausse': [var_99_hausse],
                    'VaR 99 Baisse': [var_99_baisse],
                    'Besoin en fonds propres': [besoin_en_fonds_propres]
                })
                messagebox.showinfo("RTIG (VaR 99)", f"Calcul de VaR 99 effectué.")
            else:
                messagebox.showerror("Erreur", "La variation de taux doit être saisie.")
        else:
            messagebox.showerror("Erreur", "Veuillez charger un fichier de données d'abord.")

    
    # Boutons pour les actions de l'utilisateur
    load_button = tk.Button(root, text="Charger Fichier Excel", command=load_file, **button_style)
    load_button.pack(pady=10)

    bouton_gap_de_taux = tk.Button(root, text="Gap de Taux", command=run_calculer_et_tracer_gap_de_taux, **button_style)
    bouton_gap_de_taux.pack(pady=10)

    bouton_gap_avec_couverture = tk.Button(root, text="Gap avec Couverture Neutre", command=run_calculer_et_tracer_gap_avec_couverture_neutre, **button_style)
    bouton_gap_avec_couverture.pack(pady=10)

    bouton_impact_mni = tk.Button(root, text="Impact MNI", command=run_impact_mni, **button_style)
    bouton_impact_mni.pack(pady=10)

    bouton_van_par_choc = tk.Button(root, text="VAN par Choc", command=run_calculer_van_par_choc, **button_style)
    bouton_van_par_choc.pack(pady=10)

    bouton_impact_sur_bilan = tk.Button(root, text="Impact sur Bilan", command=run_calculer_impact_sur_bilan, **button_style)
    bouton_impact_sur_bilan.pack(pady=10)

    bouton_var_99 = tk.Button(root, text="Calcul RTIG", command=run_calculer_var_99, **button_style)
    bouton_var_99.pack(pady=10)

    bouton_exporter = tk.Button(root, text="Exporter Résultats Excel", command=exporter_resultats, **button_style)
    bouton_exporter.pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()