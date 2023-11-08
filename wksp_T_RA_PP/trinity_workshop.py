import os
import re

import openpyxl
from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class TrinityWorkshop:
    project_pth: str = os.getcwd()
    res_pth: str = project_pth + "\\res"
    destiny_pth: str = project_pth + "\\destiny_files"
    origin_pth: str = project_pth + "\\origin_files"
    src_pth = res_pth + "\\MATRIZ DE TENDIDO.txt"
    wb = open(src_pth)

    starting_corr = 0
    db: list[dict[str:Cell]] = []

    @classmethod
    def fetch(cls):
        for row in cls.wb:
            pts = row.split("\t")
            fields: dict[str:str] = {
                "tag": pts[0].strip(),
                "tipo_cable": pts[1].strip(),
                "desde": pts[2].strip(),
                "hasta": pts[3].strip(),
                "n_hebras": pts[5].strip(),
                "n_ptas": pts[6].strip(),
                "ubicacion": pts[7].strip(),
                "largo": pts[4].strip(),
                "homologacion": pts[8].strip(),
            }
            cls.db.append(fields)
            pass
        pass

    @classmethod
    def find(cls, tag):
        for f in cls.db:
            if f["tag"] == tag:
                return f
            pass
        pass

    @classmethod
    def elem(cls, idx):
        return cls.db[idx]

    pass

    @classmethod
    def len(cls):
        return len(cls.db)

    planos: dict[str:str] = {
        "Planta": "SNN4008-E-PRN-14-EL-PL-0001-L0001",
        "PK": "SNN4008-E-PRN-14-EL-PL-0002-L0001",
        "SalaPK_1_2": "SNN4008-E-PRN-14-EL-PL-0008-L0001",
        "SalaPK_3_4": "SNN4008-E-PRN-14-EL-PL-0009-L0001",
        "PJ": "SNN4008-E-PRN-14-EL-PL-0003-L0001",
        "SalaPJ_1_2": "SNN4008-E-PRN-14-EL-PL-0011-L0001",
        "SalaPJ_3": "SNN4008-E-PRN-14-EL-PL-0012-L0001",
        "PATR": "SNN4008-E-PRN-14-EL-PL-0004-L0001",
        "PZ": "SNN4008-E-PRN-14-EL-PL-0005-L0001",
        "SSGG": "SNN4008-E-PRN-14-EL-PL-0010-L0001"
    }

    @classmethod
    def resolve_plano(cls, ub, desde, hasta):
        if ub:
            if re.match("500", ub):
                return cls.planos["PK"]
            if re.match("220", ub):
                return cls.planos["PJ"]
            if re.match("casa", ub, re.IGNORECASE):
                return cls.planos["SSGG"]
            pass
        desde_main = desde.split("+")[0]
        hasta_main = hasta.split("+")[0]

        if re.match("PK", desde_main):
            if re.match("PK", hasta_main):
                return cls.planos["PK"]

        if re.match("^SKD1(.)+$", desde_main, re.IGNORECASE):
            if re.match("^SKD1(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["SalaPK_1_2"]
            if re.match("^SKD2(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD3(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD4(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
        if re.match("^SKD2(.)+$", desde_main, re.IGNORECASE):
            if re.match("^SKD1(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["SalaPK_1_2"]
            if re.match("^SKD2(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD3(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD4(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
        if re.match("^SKD3(.)+$", desde_main, re.IGNORECASE):
            if re.match("^SKD1(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD2(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD3(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["SalaPK_3_4"]
            if re.match("^SKD4(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["SalaPK_3_4"]
        if re.match("^SKD4(.)+$", desde_main, re.IGNORECASE):
            if re.match("^SKD1(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD2(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["PK"]
            if re.match("^SKD3(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["SalaPK_3_4"]
            if re.match("^SKD4(.)+$", hasta_main, re.IGNORECASE):
                return cls.planos["SalaPK_3_4"]
        if re.match("SJD1", desde_main):
            if re.match("SJD1", desde_main):
                return cls.planos["SalaPJ_1_2"]
            if re.match("SJD2", desde_main):
                return cls.planos["SalaPJ_1_2"]
            if re.match("SJD3", desde_main):
                return cls.planos["PJ"]

        if re.match("^tdc(.)+$", desde_main, re.IGNORECASE):
            return cls.planos["SSGG"]
        if re.match("^tgc(.)+$", desde_main, re.IGNORECASE):
            return cls.planos["SSGG"]
        if re.match("^tdc(.)+$", hasta_main, re.IGNORECASE):
            return cls.planos["SSGG"]
        if re.match("^tgc(.)+$", hasta_main, re.IGNORECASE):
            return cls.planos["SSGG"]
        if re.match("^PK(.)+$", desde_main, re.IGNORECASE):
            return cls.planos["PK"]
        if re.match("^PJ(.)+$", desde_main, re.IGNORECASE):
            return cls.planos["PJ"]

        if re.match("PZ(.)+", desde_main):
            if re.match("SKD", hasta_main):
                return cls.planos["Planta"]
            if re.match("PCZ(.)+", hasta):
                return cls.planos["PZ"]
        if re.match("PATR(.)+", desde):
            if re.match("SKD(.)+", hasta):
                return cls.planos["Planta"]
            if re.match("PATR(.)+", hasta):
                return cls.planos["PATR"]
        if re.match("(.)*(BAT)(.)*", desde):
            if re.match("(.)*(BAT)(.)*", hasta):
                return cls.planos["SSGG"]
        print("A: " + desde + "\nB: " + hasta + "\nNo coincide con plano")
        # todo

    pass


class TendidoWorkshop:
    res_pth = TrinityWorkshop.res_pth + "\\T"
    working_template: Workbook = openpyxl.load_workbook(res_pth + "\\Protocolo Tendido de conductores electricos.xlsx")
    working_ws: Worksheet = working_template.worksheets[0]

    tags_to_gen = []
    corrs = []
    fl = open(TrinityWorkshop.origin_pth + "\\out.txt")
    for line in fl:
        pts = line.strip().split("\t")
        tags_to_gen.append(pts[0])
        corrs.append(pts[1])
        pass

    @classmethod
    def set(cls, r, c, val):
        cls.working_ws.cell(r, c, val)
        pass

    fields: dict[str:Cell] = {
        "correlativo": working_ws.cell(12, 23),
        "plano": working_ws.cell(13, 7),
        "fecha": working_ws.cell(13, 23),
        "check_control": working_ws.cell(17, 7),
        "check_alumbrado": working_ws.cell(17, 16),
        "check_fuerza": working_ws.cell(17, 23),
        "check_pantalla_y": working_ws.cell(17, 29),
        "check_pantalla_n": working_ws.cell(17, 33),
        "tag": working_ws.cell(19, 9),
        "seccion": working_ws.cell(20, 9),
        "aislacion": working_ws.cell(21, 9),
        "desde": working_ws.cell(19, 23),
        "hasta": working_ws.cell(20, 23),
        "longitud": working_ws.cell(21, 23),
        "elabora_nombre": working_ws.cell(44, 4),
        "elabora_cargo": working_ws.cell(45, 4),
        "elabora_fecha": working_ws.cell(46, 4),
        "revisa_nombre": working_ws.cell(44, 13),
        "revisa_cargo": working_ws.cell(45, 13),
        "revisa_fecha": working_ws.cell(46, 13),
    }

    @classmethod
    def fetch(cls):
        pass

    @classmethod
    def inprint(cls, fields: dict[str:Cell], corr):

        pts = fields["tipo_cable"].split(" ")

        cls.fields["correlativo"].value = corr
        cls.fields["plano"].value = TrinityWorkshop.resolve_plano(fields["ubicacion"], fields["desde"], fields["hasta"])
        cls.fields["fecha"].value = "05/06/2023"
        cls.toggle_field(fields["tag"])
        cls.fields["tag"].value = fields["tag"]
        cls.fields["seccion"].value = pts[0] + " " + pts[1]
        cls.fields["aislacion"].value = pts[2]
        cls.fields["desde"].value = fields["desde"]
        cls.fields["hasta"].value = fields["hasta"]
        cls.fields["longitud"].value = fields["largo"]
        cls.toggle_checks()
        cls.fields["elabora_nombre"].value = "Jose Godoy"
        cls.fields["elabora_cargo"].value = "Supervisor"
        # cls.fields["elabora_fecha"].value=corr
        cls.fields["revisa_nombre"].value = "Claudio Boris Hurtado G."
        cls.fields["revisa_cargo"].value = "Jefe Terreno"
        # cls.fields["revisa_fecha"].value=corr

        pass

    @classmethod
    def gen(cls):
        for i in range(len(cls.tags_to_gen)):
            fields = TrinityWorkshop.find(cls.tags_to_gen[i])
            cls.inprint(fields, cls.corrs[i])
            cls.working_template.save(
                TrinityWorkshop.destiny_pth + "\\" + cls.corrs[i] + "-T_" + fields["tag"] + ".xlsx")
        pass

    pass

    @classmethod
    def toggle_field(cls, tag):
        if re.match("W(.)+$", tag, re.IGNORECASE):
            cls.fields["check_control"].value = "✔"
            cls.fields["check_pantalla_y"].value = "✔"
            cls.fields["check_fuerza"].value = ""
            cls.fields["check_pantalla_n"].value = ""
            return
        cls.fields["check_fuerza"].value = "✔"
        cls.fields["check_pantalla_n"].value = "✔"
        cls.fields["check_control"].value = ""
        cls.fields["check_pantalla_y"].value = ""
        pass

        pass

    @classmethod
    def toggle_checks(cls):
        for i in range(27, 36):
            cls.working_ws.cell(i, 16, "✔")
        pass


class AislacionWorkshop:
    res_pth = TrinityWorkshop.res_pth + "\\RA"
    working_template: Workbook = openpyxl.load_workbook(res_pth + "\\Protocolo Pruebas de aislación de cables.xlsx")
    working_ws: Worksheet = working_template.worksheets[0]

    tags_to_gen = []
    corrs = []
    fl = open(TrinityWorkshop.origin_pth + "\\out.txt")
    for line in fl:
        pts = line.strip().split("\t")
        tags_to_gen.append(pts[0])
        corrs.append(pts[1])
        pass

    field: dict[str: Cell] = {
        "plano": working_ws.cell(11,3),
        "correlativo": working_ws.cell(13,3),
        "fecha":working_ws.cell(15,3),

        "check_control": working_ws.cell(20, 2),
        "check_alumbrado": working_ws.cell(22, 2),
        "check_fuerza": working_ws.cell(24, 2),
        "check_pantalla_y": working_ws.cell(21, 5),
        "check_pantalla_n": working_ws.cell(23, 5),
        "tag": working_ws.cell(26, 3),
        "seccion": working_ws.cell(27, 3),
        "aislacion": working_ws.cell(28, 3),
        "desde": working_ws.cell(30, 3),
        "hasta": working_ws.cell(31, 3),
        "longitud": working_ws.cell(32, 3),
        "elabora_nombre": working_ws.cell(44, 4),
        "elabora_cargo": working_ws.cell(45, 4),
        "elabora_fecha": working_ws.cell(46, 4),
        "revisa_nombre": working_ws.cell(44, 13),
        "revisa_cargo": working_ws.cell(45, 13),
        "revisa_fecha": working_ws.cell(46, 13),

    }
    pass


class PuntoPuntoWorkshop:
    pass


TrinityWorkshop.fetch()
TendidoWorkshop.gen()
pass
