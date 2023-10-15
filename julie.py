class Animal:
    def __init__(self, n_pattes):
        self.n_pattes = n_pattes

    def eat(self):
        return 'miam'

    def bruit(self):
        return 'GRRRRRR'


class Mouton(Animal):
    def __init__(self, couleur_laine):
        super().__init__(n_pattes=4)
        self.couleur_laine = couleur_laine

    def changer_couleur(self, nouvelle_couleur):
        self.couleur_laine = nouvelle_couleur

    def bruit(self):
        return 'BEEEEE'


sheepie = Mouton("bleu")
sheepie.changer_couleur("rouge")

moutonne = Mouton("vert")

...
