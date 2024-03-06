import enum
import os
from enum import Enum


class Sectores(Enum):
    PJ7 = enum.auto()
    PJ8 = enum.auto()
    PJ9 = enum.auto()
    PJ10 = enum.auto()
    PJ11 = enum.auto()
    PJ12 = enum.auto()
    PJ13 = enum.auto()
    PJ14 = enum.auto()
    PJ15 = enum.auto()
    PJ16 = enum.auto()
    PJ17 = enum.auto()
    PJ18 = enum.auto()
    PJ19 = enum.auto()
    SJD3 = [PJ7, PJ8, PJ9]
    SJD4 = [PJ10, PJ11, PJ12]
    SJD5 = [PJ13, PJ14, PJ15]
    SJD6 = [PJ16, PJ17, PJ18]
    SJ00 = enum.auto()
    SSAA = enum.auto()
    PJ = [SJD3, SJD4, SJD5, SJD6, SJ00, SSAA]
    pass


infl = open("fl.txt", "r")
otfl = open("out.txt", "w")

project_pth: str = os.getcwd()
res_pth: str = project_pth

dta = []
for line in infl:
    pts = line.split("\t")
    tmp = {
        "tag": pts[0],
        "desde": pts[1],
        "hasta": pts[2],
        "sector": ""
    }
    dta.append(tmp)
    pass


def ub_en_tag(elemento):
    for s in Sectores:
        s: Enum
        if s.name in e["tag"]:
            elemento["sector"] = s.value
            return True
        pass
    return False
    pass


def resolve_ub_org_dest(elemento):
    

    pass


for e in dta:
    if ub_en_tag(e):
        continue
    resolve_ub_org_dest(e)


    pass
