#-*-coding:UTF-8 -*
import os
import pickle
from donnees import *



#### GESTION SCORES ####
def recup_score(joueur):
    with open("D:/ut/mbg17590/Scripts/Python_2/pendu/score", "rb") as score:
            recup = pickle.Unpickler(score)
            score_recup = recup.load()
    score_joueur = score_recup[joueur]
    return joueur, score_joueur

def afficher_scores():
    with open("D:/ut/mbg17590/Scripts/Python_2/pendu/score", "rb") as score:
        recup = pickle.Unpickler(score)
        score_recup = recup.load()
    print("SCORES")    
    for cle, valeur in score_recup.items():
        print(f"{cle} a {valeur} point(s)")
    return score_recup.items()

def verif_fic_score():
    if os.access("D:/ut/mbg17590/Scripts/Python_2/pendu/score",1) == True:
        print("Fichier de score trouvé ! Vérification")
        try:
            with open("D:/ut/mbg17590/Scripts/Python_2/pendu/score", "rb") as score:
                recup = pickle.Unpickler(score)
                score_recup = recup.load()
        except EOFError:
            print("Fichier vide")
            score_recup = {}
        if not score_recup:
             score_recup[nom_joueur] = 0
             with open("D:/ut/mbg17590/Scripts/Python_2/pendu/score", "wb") as score:
                picki = pickle.Pickler(score)
                picki.dump(score_recup)
        else :        
            print("Vérification si partie en cours :")
            if (nom_joueur in score_recup):
                print("Même joueur joue encore")
            else:
                print("Première partie : inscription du joueur")
                score_recup[nom_joueur] = 0
                with open("D:/ut/mbg17590/Scripts/Python_2/pendu/score", "wb") as score:
                    picki = pickle.Pickler(score)
                    picki.dump(score_recup)
        print(score_recup)


def maj_score(parti_score,nom_joueur):
    with open("D:/ut/mbg17590/Scripts/Python_2/pendu/score", "rb") as score:
        recup = pickle.Unpickler(score)
        score_recup = recup.load()
    try:
        old_score = score_recup[nom_joueur]
        old_score = int(old_score)
        parti_score = int(parti_score)
        new_score = old_score + parti_score
        score_recup[nom_joueur] = new_score
        with open("D:/ut/mbg17590/Scripts/Python_2/pendu/score", "wb") as score:
            picki = pickle.Pickler(score)
            picki.dump(score_recup)
        print("Scores mis à jour")
    except:
        print("Pb pendant la maj des scores...")