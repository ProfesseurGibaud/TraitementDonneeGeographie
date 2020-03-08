
import os

import xlrd
from xlwt import Workbook, Formula
from wordcloud import WordCloud
import matplotlib.pyplot as plt


if __name__ == '__main__':
    def dossier():
        os.chdir(r"C:\Users\Utilisateur\Desktop\Marion")
    dossier()

Book = xlrd.open_workbook("questionnaire projets.xlsx")
SheetDuruy = Book.sheet_by_name("2NDE7 DURUY")
SheetFaure = Book.sheet_by_name("2NDE3 FAURE")

def ListeVersListeValeur(Liste):
    LL = []
    for cell in Liste:
        LL.append(cell.value)
    del LL[0]
    return LL

def TraitementFloat(Liste):
    LL= []
    NA = 0
    for item in Liste:
        if type(item) == float and item > 0:
            LL.append(item)
        elif item == "NA":
            NA = NA + 1
    return LL,NA

def TraitementString(Liste):
    LL = []
    NA = 0
    for item in Liste:
        if item == "NA":
            NA = NA + 1
        elif item == "":
            pass
        else:
            LLL = item.split(',')
            for iitem in LLL:
                if iitem[0] == " ":
                    LL.append(iitem[1:])
                else:
                    LL.append(iitem)
    return LL,NA



ListeLieuDuruy,NA_Lieu_Dury = TraitementFloat(ListeVersListeValeur(SheetDuruy.col(1)))

ListeDecoDuruy,NA_Deco_Dury = TraitementFloat(ListeVersListeValeur(SheetDuruy.col(2)))

ListeLoisirDuruy,NA_Loisir_Dury = TraitementFloat(ListeVersListeValeur(SheetDuruy.col(3)))

ListeArrondissementDuruy,NA_Arrondissement_Dury = TraitementFloat(ListeVersListeValeur(SheetDuruy.col(4)))

ListeTravailDuruy,NA_Travail_Dury =TraitementFloat( ListeVersListeValeur(SheetDuruy.col(5)))

ListeEmotionLyceeDuruy, NA_EmotionLycee_Dury =TraitementString(ListeVersListeValeur(SheetDuruy.col(6)))

ListeTypeLoisirDuruy,NA_TypeLoisir_Dury = TraitementString(ListeVersListeValeur(SheetDuruy.col(7)))

ListeEmotionLoisirDuruy,NA_EmotionLoisir_Dury = TraitementString(ListeVersListeValeur(SheetDuruy.col(8)))

ListeEmotionMaisonDuruy,NA_EmotionMaison_Dury = TraitementString(ListeVersListeValeur(SheetDuruy.col(9)))

ListeMoyenDeplacementDuruy,NA_MoyenDeplacement_Dury = TraitementString(ListeVersListeValeur(SheetDuruy.col(10)))

ListeEmotionDeplacementDuruy,NA_EmotionDeplacement_Dury = TraitementString(ListeVersListeValeur(SheetDuruy.col(11)))


DicoDuruy = {"Lieu": ListeLieuDuruy , "Deco": ListeDecoDuruy, "Loisir" : ListeLoisirDuruy,"Arrondissement" : ListeArrondissementDuruy, "Emotion Lycée" : ListeEmotionLyceeDuruy, "Travail" : ListeTravailDuruy, "Type de Loisir" : ListeTypeLoisirDuruy, "Emotion Loisir" : ListeEmotionLoisirDuruy, "Emotion Maison" : ListeEmotionMaisonDuruy, "Moyen Déplacement" : ListeMoyenDeplacementDuruy, "Emotion Déplacement" : ListeEmotionDeplacementDuruy}



def ListeToString(Liste):
    string = ""
    for item in Liste:
        string = string + " " + str(item)
    return string


def NuageDePoint(CleDico,Dico):
    Texte = ListeToString(Dico[CleDico])
    wordcloud = WordCloud().generate(Texte)
    image = wordcloud.to_image()
    image.show()
    image.save(CleDico + ".png")
    L = Dico[CleDico]
    dico = {}
    for item in L:
        if item in dico:
            dico[item] = dico[item] + 1
        else:
            dico[item] = 0
    for item in dico:
        print(item, dico[item])
    return dico


def MinMaxMoyEcarType(CleDico,Dico):
    L = Dico[CleDico]
    print(CleDico + "\n")
    print("Moyenne : " + str(np.mean(L)) + "\n")
    print("Minimum : " +str(np.min(L)) + "\n")
    print("Maximum : " + str(np.max(L)) + "\n")
    print("Ecart Type : " + str(np.std(L)) + "\n")











"""

Utilisation


"""


ListeTag = ["Lieu", "Deco", "Loisir", "Arrondissement", "Emotion Lycée", "Travail", "Type de Loisir", "Emotion Loisir", "Emotion Maison", "Moyen Deplacement", "Emotion Deplacement"]