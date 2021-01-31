#-*-coding:UTF-8 -*

class DicoRange ():
    """Classe d'exercice : création d'un dictionnaire rangé
    """
    def __init__(self, base={}, **donnees):
        self._list_key = []
        self._list_value = []

        for cle in base:
            self[cle] = base[cle]
        for cle in donnees:
            self[cle] = donnees[cle]
        
    
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

    
    def __getitem__(self, index):
        if index not in self._list_key:
            raise KeyError(f"pas de clef {index} !")
        else :
            indice = self._list_key.index(index)
            return self._list_value[indice]
    
    def __setitem__(self, index, value):
        if index in self._list_key:
            indice = self._list_key.index(index)
            self._list_value[indice] = value
        else:
            self._list_key.append(index)
            self._list_value.append(value)
    
    def __delitem__(self, index):
        del self._list_key[index], self._list_value[index]
    
    def __contains__(self,to_check):
        if to_check in self._list_key or to_check in self._list_value:
            return True
        else:
            return False
    
    def __len__(self):
        return len(self._list_key)

     
    def items(self):
        for i, cle in enumerate(self._list_key):
            value = self._list_value[i]
            yield(cle, value)
    def keys(self):
        return list(self._list_key)
    def values(self):
        return list(self._list_value)

