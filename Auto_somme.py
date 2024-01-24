import openpyxl
from openpyxl.styles import Protection

# Charger le fichier Excel
fichier_excel = openpyxl.load_workbook('votre_tableau.xlsx')

# Sélectionner la feuille de calcul (assurez-vous de remplacer 'Feuille1' par le nom réel de votre feuille)
feuille = fichier_excel['Feuille1']

# Sélectionner la cellule somme (assurez-vous de remplacer 'A1' par l'emplacement réel de votre cellule somme)
cellule_somme = feuille['A1']

# Verrouiller la cellule
cellule_somme.protection = Protection(locked=True)

# Protéger la feuille de calcul (vous pouvez définir un mot de passe si nécessaire)
feuille.protection.sheet = True

# Sauvegarder les modifications dans le fichier Excel
fichier_excel.save('tableau_protege.xlsx')

# Fermer le fichier Excel
fichier_excel.close()
