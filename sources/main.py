import customtkinter as ctk
from tkinter import filedialog, messagebox
import convert_to_json

def selectionnerFichierEntree():
    nom_fichier = filedialog.askopenfilename(title = "Sélectionnez le fichier à traiter",
                                          filetypes = (("Document Excel",
                                                        "*.xlsx"),))    
    fichier_entree.configure(text = nom_fichier)

def selectionnerDossierSortie():
    nom_dossier = filedialog.askdirectory(title = "Sélectionnez le dossier où exporter le json")   
    fichier_sortie.configure(text = nom_dossier)

def ConvertirFichier(f_in, d_out):
    try:
        nom_fichier_cree = convert_to_json.convert_to_json(f_in, d_out)
    except:
        messagebox.showerror(title="Erreur", message="Erreur lors de la création du fichier")

    messagebox.showinfo(title="Traitement terminé", message=f"Le fichier {nom_fichier_cree} a été généré avec succès")
                                        


ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

root = ctk.CTk()
root.geometry("800x550")
root.title("Convertisseur Xlsx vers Json")

entree_frame = ctk.CTkFrame(master=root)
entree_frame.pack(pady=20, padx=40, fill="both", expand=True)

entree_label = ctk.CTkLabel(master=entree_frame, text="Fichier en entrée", fg_color="grey80", padx=30, corner_radius=6)
entree_label.pack(pady=12, padx=10)

fichier_entree = ctk.CTkLabel(master=entree_frame, text="Selectionnez un fichier", anchor="w", fg_color="grey95", corner_radius=6, width=550)
fichier_entree.pack(pady=12, padx=10, side=ctk.LEFT)

bouton_explorer_entree = ctk.CTkButton(master=entree_frame, text="Explorer...", command=selectionnerFichierEntree)
bouton_explorer_entree.pack(pady=12, padx=10, side=ctk.RIGHT)

sortie_frame = ctk.CTkFrame(master=root)
sortie_frame.pack(pady=(0, 20), padx=40, fill="both", expand=True)

sortie_label = ctk.CTkLabel(master=sortie_frame, text="Fichier en sortie", fg_color="grey80", padx=30, corner_radius=6)
sortie_label.pack(pady=12, padx=10)

fichier_sortie = ctk.CTkLabel(master=sortie_frame, text="Selectionnez le dossier où exporter le json", anchor="w", fg_color="grey95", corner_radius=6, width=550)
fichier_sortie.pack(pady=12, padx=10, side=ctk.LEFT)

bouton_explorer_sortie = ctk.CTkButton(master=sortie_frame, text="Explorer...", command=selectionnerDossierSortie)
bouton_explorer_sortie.pack(pady=12, padx=10, side=ctk.RIGHT)

bouton_exporter = ctk.CTkButton(master=root, text="Convertir en Json", font=("Arial", 28),
                                command=lambda:ConvertirFichier(fichier_entree.cget("text"), fichier_sortie.cget("text"))
                                )
bouton_exporter.pack(pady=(0, 20), padx=40, fill="both", expand=True)


if __name__ == "__main__":
    root.mainloop() 