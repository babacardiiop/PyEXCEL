import openpyxl
from openpyxl.styles import Protection
#Les deux premières lignes nous permettent d'importer les librairies pour la bonne exécution du script. C'est à ne surtout pas oublier.......
# Je load en premier temps le fichier Excel que je souhaite modifier
fichier_excel = openpyxl.load_workbook('testme.xlsx')

# Je sélectionne la feuille de calcul sur laquelle je veux effectuer les modifications
feuille = fichier_excel['Feuille1']

# Je sélectionne la cellule somme 
cellule_somme = feuille['A1']

# Je verrouille la cellule
cellule_somme.protection = Protection(locked=True)

# Protégeons la feuille de calcul
feuille.protection.sheet = True

# Je sauvegarder les modifications dans le fichier Excel
fichier_excel.save('testme_protege.xlsx')

# Fermons le fichier Excel
fichier_excel.close()
