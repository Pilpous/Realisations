#-*-coding:UTF-8 -*
from random import randrange
from math import ceil
import os


#croupier = randrange(50)
print("Tu dispose de 50 crédits")
cred = 50
on_continue = True

while on_continue == True:
    choix_joueur = -1
    while choix_joueur < 0 or choix_joueur > 49:
        choix_joueur = input("Sur quoi tu mises (entre 0 et 49) :")
        try:
            choix_joueur = int(choix_joueur)
        except ValueError:
            print("Mise sur un chiffre, connard !")
            choix_joueur = -1
            continue
        if choix_joueur < 0:
            print("ça, c'est négatif...")
        elif choix_joueur > 49:
            print("ça, c'est au dessus de 49...")
    
    mise = 0
    while mise <= 0 or mise > cred:
        mise = input("Combien tu mises : ")
        try:
            mise = int(mise)
        except ValueError:
            print("Mise un chiffre, connard !")
            mise = -1
            continue
        if mise <= 0:
            print("la mise est négative ou nulle...")
        elif mise > cred:
            print("T'as pas assez de blé !")


    croupier = randrange(50)
    print(f"le numéro gagnant est le {croupier}")     
    if choix_joueur == croupier:
        gain = mise * 3
        cred = cred + gain
        print("Bingo !")
        print(f"Tu gagnes 3 fois ta mise : {gain} ")
    elif choix_joueur % 2 == 0 and croupier % 2 == 0:
        gain = mise / 2 
        cred = cred + ceil(gain)
        print("Numéro pair")
        print(f"Tu récupères la moitié de ta mise, arrondi : {ceil(gain)}")
    elif choix_joueur % 2 == 1 and croupier % 2 == 1:
        gain = mise / 2 
        cred = cred + ceil(gain)
        print("Numéro impair")
        print(f"Tu récupères la moitié de ta mise, arrondi : {ceil(gain)}")
    else:
        cred = cred - mise
        print(f"T'as perdu {mise} !")


    if cred <= 0:
        print("t'as plus de blé... fin de partie")
        on_continue = False
    else:
        print(f"Tu dispose de {cred} après ce tour")
        ca_joue = input("Tu continues ? O/N")
        if ca_joue.upper != "O" and ca_joue.upper != "N":
            print("ah ouais ? ben tiens !")
            on_continue = False
        elif ca_joue.upper == "O":
            on_continue = True
        elif ca_joue.upper == "N":
            on_continue = False


os.system("pause")