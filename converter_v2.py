import json
import io
import csv
from openpyxl import load_workbook
from typing import TypedDict, Optional
from dataclasses import dataclass, field


########## CONVERSION HELPER FUNCS
def conv_plato(sg):
    return(round((-1 * 616.868) + (1111.14 * sg) - (630.272 * sg**2) + (135.997 * sg**3), 1))

def conv_ebc(srm):
    return round(float(srm*1.97), 1)




########## MAIN RECIPE CLASS
class Recipe:
    def __init__(self, name):
        #ingredients
        self.fermentables = list[Fermentable] = []
        self.hops = list[Hop] = []
        self.yeasts = list[Yeast] = []
        self.misc = list[Misc] = []
        self.water = list[Water] = []

        #process steps
        self.mash = list[MashStep] = []
        self.lauter = Lauter()
        self.boil = Boil()
        self.cool = Cool()
        self.ferm = list[FermStep] = []

        #overall beer properties
        self.properties: BeerProperties()

    def make_brewsheet(self, batchno, brewdate, out_path, excel_layout = 'brewsheet_empty.xlsx'):
        raise NotImplementedError

    def read_BFjson(self, in_path):
        with open(path, 'r') as file:
        data = json.load(file)

        self.properties.name = data["name"]
        self.properties.style = data['style']['name']
        self.properties.og = conv_plato(data["og"])
        self.properties.fg = conv_plato(data["fg"])
        self.properties.abv = round(data['abv'], 1)
        self.properties.color = conv_ebc(data["color"])
        self.properties.ibu = data["ibu"]

        for fermentable in data['data']['mashFermentables']:
            f = Fermentable(
                name = fermentable[name],
                amount = fermentable[amount]
                supplier = fermentable[name]
                )
            self.fermentables.append(f)
        
        for h in data['hops']:
            if h["name"] not in self.hops.name:
                self.hops.append(Hop(name = h["name"], alpha = round(h["alpha"], 1)))
            for i in self.hops.name = h["name"]:
                



        



#DATA CLASSES FOR MAIN RECIPE CLASS 
@dataclass
class Fermentable:
    name:str
    amount:float
    supplier:str


@dataclass 
class HopAddition:
    add_type: str
    time: int
    amount: float

@dataclass
class Hop:
    name: str
    alpha:float
    additions: list[HopAddition]= []

    @property
    def amount(self) -> float: 
        additions.amount.sum()

    def add_addition(self, amount, time, add_type):
        self.additions.append{
            "add_type": add_type
            "time": time,
            "amount": amount}
        self.amount = calc_total_amount(self.additions)


@dataclass
class Misc:
    raise NotImplementedError()


@dataclass
class Yeast:
    raise NotImplementedError()


@dataclass 
class BeerProperties:
    name: str = None
    style: str = None
    og: float = None
    fg: float = None
    abv: float = None
    color: int = None
    ibu: int = None













    




