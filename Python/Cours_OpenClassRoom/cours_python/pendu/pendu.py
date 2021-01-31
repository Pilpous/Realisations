#-*-coding:UTF-8 -*
#import donnees
from random import choice
import os
import pickle
from fonctions import *
from donnees import *

on_continue = "z"
verif_fic_score()

while on_continue.upper() != "Q":
    mot_a_trouver = choice(liste_mots)
    longueur_mot_a_trouver = "*"*len(mot_a_trouver)
    trouvee = {}
    tentative = 0
    
    print("C'est parti :")
    print(longueur_mot_a_trouver)

    while tentative != nb_a_jouer and longueur_mot_a_trouver != mot_a_trouver:
        lettre = "zz"
        while lettre.upper() not in alphabet :
            lettre = input("Saisis une lettre :")

        print(f"La lettre {lettre.upper()} se trouve-t-elle dans le mot mystère ?")

        compteur = 0
        for i, l in enumerate(mot_a_trouver):
            if l == lettre:
                trouvee[i] = l
                compteur =+ 1
        print(f"Cette lettre est {compteur} fois dans le mot mystère !")

        for cle in trouvee:
            longueur_mot_a_trouver = longueur_mot_a_trouver[0:cle] + longueur_mot_a_trouver[cle].replace(longueur_mot_a_trouver[cle],trouvee[cle]) + longueur_mot_a_trouver[cle+1:]

        print(f"Mot mystère : {longueur_mot_a_trouver} ")
        tentative = tentative + 1
        #nb_a_jouer = nb_a_jouer - tentative
        print(f"tentative = {tentative}")
        print(f"nb a jouer = {nb_a_jouer}")
        print(f"Plus que {nb_a_jouer - tentative} tentatives pour trouver le mot mystère...")

    score = nb_a_jouer - tentative
    if tentative == nb_a_jouer:
        print("LOOOOOOOOSER !")
        print(f"Le mot à deviner était : {mot_a_trouver}")
        print(f"Tu as marqué {score} point sur ce tour.")

    elif longueur_mot_a_trouver == mot_a_trouver:
        print("Félicitations, tu as trouvé le mot !")
        print(f"Tu as marqué {score} point(s) sur ce tour.")

    maj_score(score,nom_joueur)
    recup_score(nom_joueur)

    on_continue = input("Vous voulez arrêter ? Tapez Q...")

afficher_scores()

os.system("pause")