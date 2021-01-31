#-*-coding:UTF-8 -*
import os

if __name__ == "__main__":
    annee = input("Saisi une année : ")
    annee = int(annee)

    if annee % 400 == 0 or (annee % 4 == 0 and annee % 100 != 0):
        print(f"l'année {annee} est bissextile !")
    else :
        print(f"l'année {annee} n'est pas bissextile !")        



def table_par_mumu(mumu, max=10):
    nb = 0
    while nb < max :
        print(f"{nb} * {mumu} =", nb *mumu)
        nb += 1

os.system("pause")

def afficher_virgule(flottant):
    if type(flottant) != float:
        raise TypeError("le paramètre attendu doit être flottant...")
    flottant = str(flottant)
    premier, deuxieme = flottant.split(".")
    return ",".join([premier,deuxieme[:3]])


def afficher(*parametres, sep=' ', fin='\n'):
    parametres = list(parametres)
    for i, parametre in enumerate(parametres):
        parametres[i] = str(parametre)
    chaine = sep.join(parametres)
    chaine += fin
    print(chaine)

    