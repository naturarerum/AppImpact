import tkinter as tk
from tkinter import HORIZONTAL
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from tkinter.ttk import Progressbar

import openpyxl


class MainGui:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("AppImpact")
        self.root.iconbitmap(r'C:\Python34\DLLs\pyc.ico')
        self.root.geometry("280x125")
        self.create_widgets()
        #self.progress = Progressbar(self, orient=HORIZONTAL, length=100, mode='indeterminate')
        self.root.mainloop()

        # TODO : Séparer en plusieurs classes

    def create_widgets(self):
        # creer un Frame principal
        self.release = tk.StringVar()
        main_frame = ttk.LabelFrame(self.root)
        main_frame.grid(column=0, row=0)

        # Frame boutons
        btn_Frame = ttk.LabelFrame(self.root)
        btn_Frame.grid(column=2, row=0)

        # Création du label Saisie de la release
        lbl_release = ttk.Label(main_frame, text="Release             : ").grid(column=0, row=0)

        # Création du widget saisie de la release
        saisie_release = ttk.Entry(main_frame, width=15, textvariable=self.release)
        saisie_release.grid(column=2, row=0)
        # TODO : Remplacer par une liste déroulante auto-populée

        # Création d'un widget bouton pour lancer le traitement
        btn_start = ttk.Button(btn_Frame, text="Start", command=self.traitement).grid(column=2, row=0)

        # Création progressbar


    def traitement(self):
        # TODO : Ajouter une exception si release pas saisie
        nomfic = self.askopenfile()
        self.traite_impact(nomfic)

    def traite_impact(self, filename):
        fichier = filename
        rel = self.release.get()  # récupere la saisie de l'utilisateur
        type_service = ['New', 'Update', 'Reuse']
        try:
            wb = openpyxl.load_workbook(fichier, data_only=True)
        except openpyxl.utils.exceptions.InvalidFileException:
            messagebox.showerror("Erreur", "le fichier doit etre au format xlsx")
        sheet = wb.get_sheet_by_name("Statut des services")
        maxrow = sheet.max_row
        for row_index in range(2, maxrow):
            flag_type = False
            flag_mois = False
            mois_release = ''
            print('Row: ' + str(row_index))
            projet = sheet.cell(row=row_index, column=1).value
            type_s = sheet.cell(row=row_index, column=2).value
            if type_s in type_service:
                flag_type = True
            nom_service = sheet.cell(row=row_index, column=3).value
            mois_release = sheet.cell(row=row_index, column=10).value
            if mois_release == rel:
                flag_mois = True
            application = sheet.cell(row=row_index, column=25).value
            if flag_mois and flag_type:
                print(projet, type_s, nom_service,mois_release,application)

    def askopenfile(self):
        # get filename
        filename = filedialog.askopenfilename()
        # return filename
        return filename
        # TODO : Ajouter exception si pas de fichier choisi (filename laissé vide)

    def _quit(self):
        self.root.quit()
        self.root.destroy()
        exit()


# Main
if __name__ == '__main__':
    # Create an Application and run it
    app = MainGui()
