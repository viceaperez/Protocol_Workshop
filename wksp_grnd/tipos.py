import os
import re
from enum import Enum

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# TODO: MODIFICAR EL ORIGEN DE DATOS DE ELEMENTOS PARA DISTINTOS PATIOS
origen = "elementos220.xlsx"

project_pth: str = os.getcwd()
res_pth: str = project_pth + "\\res"


class Cable:
    def __init__(self):
        self.calibre: str = "-"
        self.aislacion: str = "-"
        self.equipos: (str, float) = "-"
        self.estructura: (str, float) = "-"
        self.tierra: (str, float) = "-"
        pass

    pass


class Elemento:

    def __init__(self):
        self.tipo = None
        self.cables: list[Cable] = []
        self.soldaduras: list[int] = []
        self.categorias: list = []

    pass


class Categoria(Enum):
    BANCODUCTO = "Bancoductos"
    ESTRUCTURAL = "Estructurales"
    ESCALERILLA = "Escalerillas"
    CAMARA = "Camaras"
    EQUIPO = "Equipos"
    TRANSFORMADOR_OTROS = "totros"
    pass


class Tipo:
    def __init__(self, nombre, ident):
        self.nombre = nombre
        self.identificador: re.Pattern = ident
        pass


class TipoElemento(Enum):
    MALLA = Tipo("Puesta a Tierra Subterranea", re.compile("malla", re.IGNORECASE))
    CERCO = Tipo("Cerco Perimetral", re.compile("cerco p", re.IGNORECASE))
    AP = Tipo("Aislador de Pedestal", re.compile("ap", re.IGNORECASE))
    DESC_T = Tipo("Desconectador con Puesta a Tierra", re.compile("(89)(.){3,4}(t)", re.IGNORECASE))
    DESCT = Tipo("Desconectador", re.compile("(89)(.){3,4}[^t]", re.IGNORECASE))
    INTERRUPTOR = Tipo("Interruptor", re.compile("52"))
    pass


elementos: dict[str:Elemento] = {}
''''''
wb: Workbook = openpyxl.load_workbook(res_pth + "\\" + origen)
for ws in wb:
    ws: Worksheet
    el = Elemento()
    te: TipoElemento = None
    lk = False
    for tipo in TipoElemento:
        if ws.title == tipo.name:
            el.tipo = tipo
            te = tipo
            break
        pass
    if not te:
        print("Hoja no encontrada en los tipos de elemento: " + ws.title)
        continue

    for i in range(1, 13):
        el.soldaduras.append(ws.cell(i, 2).value)
        pass
    for i in range(15, 20):
        cb = Cable()
        cb.calibre = ws.cell(i, 1).value
        cb.aislacion = ws.cell(i, 2).value
        cb.equipos = ws.cell(i, 3).value
        cb.estructura = ws.cell(i, 4).value
        cb.tierra = ws.cell(i, 5).value
        el.cables.append(cb)
        pass
    for cat in Categoria:
        for i in range(1, 13):
            if cat.name == ws.cell(i, 4).value:
                el.categorias.append(cat)
                pass
            pass
        pass
    elementos[te] = el
    pass

''''''
