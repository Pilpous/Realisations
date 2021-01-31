#-*-coding:UTF-8 -*
import os
from exercice import *

#table_par_mumu(5, 15)
#os.system("pause")

try:
    annee = input("Saisi une année : ")
    annee = int(annee)

    if annee % 400 == 0 or (annee % 4 == 0 and annee % 100 != 0):
        result = "l'année  est bissextile !"
    else :
        result = "l'année n'est pas bissextile !"
except ValueError:
    print("Erreur lors de la saisie")
else:
    print(result)


fl = 3.99999999998
fl = str(fl)
res = fl.split(".")[0] , fl.split(".")[1][0:3]
",".join(res)



chaine = "bonjour"
i = 0 
while i < len(chaine):
    print(chaine[i])
    i += 1


list_fuits =[12,48,85,14,9,7]
a_virer = 7
[nb - a_virer for nb in list_fuits if nb>a_virer] 

inventaire = [
    ("pommes", 22),
    ("kiwis", 13),
    ("bananes", 48),
    ("fraises", 12),
    ("abricots", 52)
]
inverse = [(qtt,fruit) for fruit,qtt in inventaire]
inventaire = [(fruit,qtt) for qtt, fruit in sorted(inverse, reverse=True)]



class Personne:
    def __init__(self,nom,prenom,age,lieu_residence):
        self.nom = nom
        self.prenom = prenom
        self._age = age
        self.lieu_residence = lieu_residence
    def _get_age(self):
        print("Accès _age !")    
        return self._age
    def _set_age(self, nvl_age):
        print("Modification _age")
        print("{} a en fait {} ans".format(self.prenom, nvl_age))
        self._age = nvl_age
    age = property(_get_age, _set_age)
    def __getattr__(self,nom):
        print(f"Attention ! Pas d'attribut {nom} dans l'objet !")
        


class TableauNoir:
    def __init__(self,surface):
        self.surface = ""
    def ecrire(self,message):
        if self.surface != "":
            self.surface += "\n"
        self.surface += message
    def lire(self):
        print(self.surface)
    def effacer(self):
        self.surface = ""


class ZDict:
    def __init__(self):
        self._dictionnaire = {}
        self.nom = "test"
    def __getitem__(self,index):
        return self._dictionnaire[index]
    def __setitem__(self,index,valeur):
        print(f"Modification de l'objet {index} - nouvelle valeur à {valeur}")
        self._dictionnaire[index] = valeur
    def __delitem__(self,index):
        print("Suppression de l'objet ZDict")
        del self._dictionnaire[index]
    def __repr__(self):
            balance = "{"
            passage = True
            for cle, valeur in self.items():
                if not passage :
                    balance += " , "
                else:
                    passage = False
                balance += repr(cle) + ": " + repr(valeur)
            balance += "}"
            return balance

    def __str__(self):
        return repr(self)
    def __contains__(self,obj_in):
        print("Vérification : l'objet contient {}".format(obj_in))
        if obj_in in self._dictionnaire.keys():
            return True
        else:
            return False
    def __getattr__(self,nom):
        print("Alerte - Pas d'attribut {} dans cet objet".format(nom))
    def __setattr__(self, nom_att, val_attr):
        print("Modif de l'objet via setattr")
        object.__setattr__(self, nom_att, val_attr)
    def __delattr__(self, nom_att):
        raise AttributeError("On supprime pas ! Popopop !")
    def __len__(self):
        print("Taille de l'objet :")
        return len(self._dictionnaire)
    dictionnaire = property(__getitem__,__setitem__)



class Duree():
    def __init__(self, min=0, sec=0):
        self.min = min
        self.sec = sec
    def __str__(self):
        return "{0:02}:{1:02}".format(self.min, self.sec)
    def __add__(self,a_rajouter):
        nouvelle_duree = Duree()
        nouvelle_duree.min = self.min
        nouvelle_duree.sec = self.sec
        nouvelle_duree.sec += a_rajouter
        if nouvelle_duree.sec >= 60:
            nouvelle_duree.min += nouvelle_duree.sec // 60
            nouvelle_duree.sec = nouvelle_duree.sec % 60
        return nouvelle_duree
    def __radd__(self, a_rajouter):
        return self + a_rajouter
    def __iadd__(self, a_rajouter):
        self.sec += a_rajouter
        if self.sec >= 60:
            self.min += self.sec // 60
            self.sec = self.sec % 60
        return self
    def __sub__(self, a_enlever):
        nouvelle_duree = Duree()
        nouvelle_duree.min = self.min
        nouvelle_duree.sec = self.sec
        nouvelle_duree.sec -= a_enlever
        while nouvelle_duree.sec < 0:
            nouvelle_duree.min = nouvelle_duree.min -1
            nouvelle_duree.sec = 60 - abs(nouvelle_duree.sec)
        return nouvelle_duree



etudiants = [
    ("Bernard", 18, 12),
    ("Lionel", 22, 9),
    ("Fréderic", 17, 15),
    ("Stéphane", 19, 5),
    ("Olivia", 14, 18),
]
sorted(etudiants, key=lambda colonne: colonne[2])


class Etudiants():
    def __init__(self, nom, age, note):
        self.nom = nom
        self.age = age
        self.note = note
    def __repr__(self):
        return "<Etudiant {} (age={}, moyenne={})>".format(self.nom, self.age, self.note)

etudiants = [
    Etudiants("Bernard", 48, 12),
    Etudiants("Lionel", 54, 8),
    Etudiants("Fréderic", 45, 15),
    Etudiants("Stéphane", 52, 6),
    Etudiants("Olivia", 50, 18),
]

sorted(etudiants, key= lambda colonne: colonne.note)
sorted(etudiants, key=lambda colonne: colonne.age, reverse=True)


class Vivant():
    def __init__(self,etat):
        self.etat = etat
    def __str__(self):
        if self.etat == "Oui":
            return True
        else :
            return False

class Animal(Vivant):
    def __init__(self, regime, cri, habitat, etat):
        Vivant.__init__(self,etat)
        self.regime = regime
        self.cri = cri
        self.habitat = habitat

    def __str__(self):
        return "{0} - {1} - {2} - Vivant : {3}".format(self.regime,self.cri,self.habitat, self.etat)

class Canides(Animal):
    def __init__(self, surnom, regime, cri, habitat,etat):
        Animal.__init__(self,regime,cri,habitat,etat)
        self.surnom = surnom
    def __str__(self):
        return "Le chien s'appelle {0} - c'est un {1} - il {2} et il habite dans une {3} - {4}".format(self.surnom,self.regime,self.cri,self.habitat,self.etat)

chien = Canides("Medor","Omnivore","aboie","niche","Oui")


def intervalle(deb,fin):
    yield deb + 1


def Afficher(*valeurs, sep=' ', end='\n'):
    valeurs = list(valeurs)
    for i, valeur in enumerate(valeurs):
        valeurs[i] = str(valeur)
    chaine = sep.join(valeurs)
    chaine += end
    print(chaine, end=' ')

inventaire = [
    ("pommes", 22),
    ("melons", 4),
    ("poires", 18),
    ("fraises", 76),
    ("prunes", 51),
]

def tri_inventaire(inventaire):
    maj_inventaire= [(i, fruits) for fruits, i in inventaire]
    inventaire = [(fruits ,i) for i, fruits in sorted(maj_inventaire, reverse=True)]
    return inventaire

def fonctions_con(*test, **test1):
    print(test , test1)


