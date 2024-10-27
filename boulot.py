import tkinter as tk
from tkinter import messagebox, filedialog, Scrollbar
import pandas as pd
from rapidfuzz import fuzz, process


def comparer_listes():
    try:
        # Récupération des listes entrées par l'utilisateur
        liste1 = entry_liste1.get("1.0", "end-1c").splitlines()
        liste2 = entry_liste2.get("1.0", "end-1c").splitlines()
        exclusions = exclusion_entry.get("1.0", "end-1c").splitlines()

        # Vérification des entrées
        if not liste1 or not liste2 or liste1 == [""] or liste2 == [""]:
            messagebox.showwarning("Attention", "Veuillez entrer les deux listes.")
            return

        # Récupération du seuil de similarité
        seuil_similarite = int(seuil_entry.get())

        # Comparaison des listes avec exclusion
        resultats = []
        for nom in liste1:
            if nom not in exclusions:  # Vérification si le nom est exclu
                meilleure_correspondance, score, _ = process.extractOne(nom, liste2, scorer=fuzz.ratio)
                if score >= seuil_similarite:
                    resultats.append({
                        "Nom dans Liste 1": nom,
                        "Correspondance dans Liste 2": meilleure_correspondance,
                        "Score de Similarité": score
                    })

        # Affichage des résultats dans le Text widget
        resultats_text.delete("1.0", tk.END)
        for res in resultats:
            resultats_text.insert(
                tk.END,
                f"{res['Nom dans Liste 1']} -> {res['Correspondance dans Liste 2']} (Score: {res['Score de Similarité']})\n"
            )

        # Sauvegarde des résultats dans un DataFrame pour exportation
        global df_resultats
        df_resultats = pd.DataFrame(resultats)
        messagebox.showinfo("Résultats", "La comparaison est terminée !")

    except Exception as e:
        messagebox.showerror("Erreur", f"Une erreur est survenue : {e}")


def exporter_excel():
    try:
        if df_resultats.empty:
            messagebox.showwarning("Attention", "Aucun résultat à exporter.")
            return

        # Demander à l'utilisateur où sauvegarder le fichier Excel
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            df_resultats.to_excel(file_path, index=False)
            messagebox.showinfo("Exportation", f"Les résultats ont été exportés avec succès vers '{file_path}'.")

    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'exportation : {e}")


def exporter_csv():
    try:
        if df_resultats.empty:
            messagebox.showwarning("Attention", "Aucun résultat à exporter.")
            return

        # Demander à l'utilisateur où sauvegarder le fichier CSV
        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if file_path:
            df_resultats.to_csv(file_path, index=False)
            messagebox.showinfo("Exportation", f"Les résultats ont été exportés avec succès vers '{file_path}'.")

    except Exception as e:
        messagebox.showerror("Erreur", f"Erreur lors de l'exportation : {e}")


# Création de la fenêtre principale
root = tk.Tk()
root.title("Comparateur de Listes d'Entreprises")
root.geometry("700x800")

# Widgets pour la première liste
label_liste1 = tk.Label(root, text="Liste 1 (un nom d'entreprise par ligne):")
label_liste1.pack(pady=5)
frame1 = tk.Frame(root)
frame1.pack(pady=5)
entry_liste1 = tk.Text(frame1, height=8, width=50)
entry_liste1.pack(side=tk.LEFT)
scrollbar1 = Scrollbar(frame1, command=entry_liste1.yview)
scrollbar1.pack(side=tk.RIGHT, fill=tk.Y)
entry_liste1.config(yscrollcommand=scrollbar1.set)

# Widgets pour la deuxième liste
label_liste2 = tk.Label(root, text="Liste 2 (un nom d'entreprise par ligne):")
label_liste2.pack(pady=5)
frame2 = tk.Frame(root)
frame2.pack(pady=5)
entry_liste2 = tk.Text(frame2, height=8, width=50)
entry_liste2.pack(side=tk.LEFT)
scrollbar2 = Scrollbar(frame2, command=entry_liste2.yview)
scrollbar2.pack(side=tk.RIGHT, fill=tk.Y)
entry_liste2.config(yscrollcommand=scrollbar2.set)

# Zone pour la liste des exclusions
label_exclusions = tk.Label(root, text="Exclusions (un nom de dossier par ligne):")
label_exclusions.pack(pady=5)
frame3 = tk.Frame(root)
frame3.pack(pady=5)
exclusion_entry = tk.Text(frame3, height=5, width=50)
exclusion_entry.pack(side=tk.LEFT)
scrollbar3 = Scrollbar(frame3, command=exclusion_entry.yview)
scrollbar3.pack(side=tk.RIGHT, fill=tk.Y)
exclusion_entry.config(yscrollcommand=scrollbar3.set)

# Champ pour définir le seuil de similarité
label_seuil = tk.Label(root, text="Seuil de similarité (0-100):")
label_seuil.pack(pady=5)
seuil_entry = tk.Entry(root)
seuil_entry.insert(0, "80")  # Valeur par défaut
seuil_entry.pack(pady=5)

# Bouton pour lancer la comparaison
button_comparer = tk.Button(root, text="Comparer", command=comparer_listes)
button_comparer.pack(pady=10)

# Zone d'affichage des résultats
label_resultats = tk.Label(root, text="Résultats:")
label_resultats.pack(pady=5)
frame4 = tk.Frame(root)
frame4.pack(pady=5)
resultats_text = tk.Text(frame4, height=10, width=50)
resultats_text.pack(side=tk.LEFT)
scrollbar4 = Scrollbar(frame4, command=resultats_text.yview)
scrollbar4.pack(side=tk.RIGHT, fill=tk.Y)
resultats_text.config(yscrollcommand=scrollbar4.set)

# Boutons pour exporter les résultats
button_exporter_excel = tk.Button(root, text="Exporter vers Excel", command=exporter_excel)
button_exporter_excel.pack(pady=5)

button_exporter_csv = tk.Button(root, text="Exporter vers CSV", command=exporter_csv)
button_exporter_csv.pack(pady=5)

# Variable globale pour stocker le DataFrame des résultats
df_resultats = pd.DataFrame()

# Boucle principale de l'application
root.mainloop()
