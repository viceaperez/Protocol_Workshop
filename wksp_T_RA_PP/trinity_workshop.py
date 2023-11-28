import os
import random
import re

import openpyxl
import numpy
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

            tipo_cable = pts[1].strip()

            fields: dict[str:str] = {
                "tag": pts[0].strip(),
                "tipo_cable": tipo_cable,
                "desde": pts[2].strip(),
                "hasta": pts[3].strip(),
                "n_hebras": cls.get_hebras(pts[5].strip(), tipo_cable.split(" ")[0], pts[8].strip()),
                "n_ptas": pts[6].strip(),
                "ubicacion": pts[7].strip(),
                "largo": pts[4].strip(),
                "homologacion": pts[8].strip(),
                "uso": cls.resolve_uso(pts[0].strip())
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
        raise Exception("A: " + desde + "\nB: " + hasta + "\nNo coincide con plano")
        # print("A: " + desde + "\nB: " + hasta + "\nNo coincide con plano")

    pass

    @classmethod
    def resolve_uso(cls, tag):
        if re.match("W(.)+$", tag, re.IGNORECASE):
            return "C"
        return "F"
        pass

    @classmethod
    def get_hebras(cls, n_hebras, calibre, homologacion):
        if homologacion:
            tmp = homologacion.split(" ")[0]
            tmp2 = re.sub(re.compile("([^0-9])*(X)(.)*", re.IGNORECASE), "", tmp)
            return int(tmp2)
        if n_hebras:
            return int(n_hebras)
        if "c" in calibre:
            return int(calibre.strip().split("-c-")[0])
        else:
            return int(calibre.strip().split("x")[0])

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

        cls.fields["correlativo"].value = corr
        cls.fields["plano"].value = TrinityWorkshop.resolve_plano(fields["ubicacion"], fields["desde"], fields["hasta"])
        cls.fields["fecha"].value = "05/06/2023"
        cls.toggle_field(fields["tag"])
        cls.fields["tag"].value = fields["tag"]
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

        if fields["homologacion"] == "":
            pts = fields["tipo_cable"].split(" ")
            cls.fields["seccion"].value = pts[0] + " " + pts[1]
            cls.fields["aislacion"].value = pts[2]
        else:
            pts = fields["homologacion"].split(" ")
            cls.fields["seccion"].value = pts[0] + " " + pts[1]
            try:
                cls.fields["aislacion"].value = pts[2]
            except IndexError:
                pts = fields["tipo_cable"].split(" ")
                cls.fields["aislacion"].value = pts[2]

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

    fields: dict[str: Cell] = {
        "plano": working_ws.cell(11, 3),
        "correlativo": working_ws.cell(13, 3),
        "fecha": working_ws.cell(15, 3),
        "check_control": working_ws.cell(20, 2),
        "check_alumbrado": working_ws.cell(22, 2),
        "check_fuerza": working_ws.cell(24, 2),
        "check_pantalla_y": working_ws.cell(21, 5),
        "check_pantalla_n": working_ws.cell(23, 5),
        "tag": working_ws.cell(26, 3),
        "seccion": working_ws.cell(27, 3),
        "aislacion": working_ws.cell(28, 3),
        "tension_serv": working_ws.cell(29, 3),
        "desde": working_ws.cell(30, 3),
        "hasta": working_ws.cell(31, 3),
        "longitud": working_ws.cell(32, 3),
        "capacidad": working_ws.cell(33, 3),
        "instrumento_tipo": working_ws.cell(36, 3),
        "instrumento_marca": working_ws.cell(37, 3),
        "instrumento_modelo": working_ws.cell(38, 3),
        "instrumento_serie": working_ws.cell(39, 3),
        "instrumento_calibracion": working_ws.cell(40, 3),
        "ensayo_tension": working_ws.cell(43, 3),
        "ensayo_tiempo": working_ws.cell(44, 3),
        "ensayo_temperatura": working_ws.cell(45, 3),
        "ensayo_humedad": working_ws.cell(46, 3),
        "ensayo_longitud": working_ws.cell(47, 3),
        "elabora_nombre": working_ws.cell(51, 2),
        "elabora_cargo": working_ws.cell(52, 2),
        "elabora_fecha": working_ws.cell(53, 2),
        "revisa_nombre": working_ws.cell(51, 7),
        "revisa_cargo": working_ws.cell(52, 7),
        "revisa_fecha": working_ws.cell(53, 7)
    }

    capacidades = None

    @classmethod
    def inprint(cls, fields: dict[str:Cell], corr):

        cls.fields["plano"].value = TrinityWorkshop.resolve_plano(fields["ubicacion"], fields["desde"], fields["hasta"])
        cls.fields["correlativo"].value = corr
        cls.fields["fecha"].value = "29/06/2023"
        cls.toggle_checks(fields["tag"])
        cls.fields["tag"].value = fields["tag"]
        cls.fields["tension_serv"].value = "0,6-1kV"
        cls.fields["desde"].value = fields["desde"]
        cls.fields["hasta"].value = fields["hasta"]
        cls.fields["longitud"].value = fields["largo"]
        cls.fields["instrumento_tipo"].value = "Megometro"
        cls.fields["instrumento_marca"].value = "Megger"
        cls.fields["instrumento_modelo"].value = "MIT525"
        cls.fields["instrumento_serie"].value = "101629407"
        cls.fields["instrumento_calibracion"].value = "06-02-2023"
        cls.fields["ensayo_tension"].value = "1kV"
        cls.fields["ensayo_tiempo"].value = "1 Minuto"
        cls.fields["ensayo_temperatura"].value = cls.gen_temp()
        cls.fields["ensayo_humedad"].value = cls.gen_hum()
        cls.fields["ensayo_longitud"].value = fields["largo"] + " m"
        cls.fields["elabora_nombre"].value = "Jose Godoy Espinoza"
        cls.fields["elabora_cargo"].value = "Supervisor Eléctrico"
        cls.fields["elabora_fecha"].value = ""
        cls.fields["revisa_nombre"].value = "Claudio Boris H."
        cls.fields["revisa_cargo"].value = "Jefe Terreno"
        cls.fields["revisa_fecha"].value = ""

        if fields["homologacion"] == "":
            pts = fields["tipo_cable"].split(" ")
            cls.fields["capacidad"].value = cls.resolve_capacidad(pts[0], pts[1]) + " A"
            cls.fields["seccion"].value = pts[0] + " " + pts[1]
            cls.fields["aislacion"].value = pts[2]
        else:
            pts = fields["homologacion"].split(" ")
            cls.fields["capacidad"].value = cls.resolve_capacidad(pts[0], pts[1])
            cls.fields["seccion"].value = pts[0] + " " + re.sub("Â²", "²", pts[1])
            try:
                cls.fields["aislacion"].value = pts[2]
            except IndexError:
                pts = fields["tipo_cable"].split(" ")
                cls.fields["aislacion"].value = pts[2]

        cls.populate_table(fields["n_hebras"], pts[0], fields["uso"])

        pass

    pass

    @classmethod
    def fetch(cls):
        cls.capacidades = []
        tmp = open(cls.res_pth + "\\tabla corrientes.txt")
        for l in tmp:
            t = l.split(" ")
            cls.capacidades.append([t[0], t[1], t[2]])
            pass
        pass

    @classmethod
    def toggle_checks(cls, tag):
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

    @classmethod
    def resolve_capacidad(cls, calibre: str, unidad: str):
        offset = 0
        if re.match("AWG", unidad, re.IGNORECASE):
            offset += 1
        seccion = cls.get_seccion(calibre)
        for e in cls.capacidades:
            if e[offset] == seccion:
                return e[2]
        raise Exception("calibre no encontrado")
        pass

    @classmethod
    def gen_temp(cls):
        num = random.Random().randint(a=150, b=250)
        return str(num / 10)
        pass

    @classmethod
    def gen_hum(cls):
        num = random.Random().randint(a=200, b=400)
        return str(num / 10)
        pass

    @classmethod
    def gen_res(cls):
        num = numpy.random.normal(loc=100, scale=15)
        return str(int(num) / 10)

    @classmethod
    def get_hebras(cls, calibre):
        if "c" in calibre:
            return calibre.strip().split("-c-")[0].replace("-", "/")
        else:
            return calibre.strip().split("x")[0].replace("-", "/")
        pass

    @classmethod
    def get_seccion(cls, calibre):
        if "c" in calibre:
            pts = calibre.strip().split("-c-")
            return pts[1].replace("-", "/")
        else:
            pts = calibre.strip().split("x")
            return pts[len(pts) - 1].replace("-", "/")
        pass

    @classmethod
    def populate_table(cls, n_hebras, calibre, uso):
        if n_hebras:
            hebras = int(n_hebras)
        else:
            hebras = int(cls.get_hebras(calibre))
            pass

        for i in range(0, hebras):
            row = 12 + (2 * i)
            for j in range(i, hebras):
                if i == j:
                    continue
                col = 8 + (2 * j)
                cls.working_ws.cell(row, col).value = cls.gen_res()
                pass
            if uso == "C":
                cls.working_ws.cell(row, 40).value = cls.gen_res()
                cls.working_ws.cell(row, 41).value = cls.gen_res()
                cls.working_ws.cell(row, 42).value = "OK"

        pass

    @classmethod
    def gen(cls):
        for i in range(len(cls.tags_to_gen)):
            cls.flush()
            fields = TrinityWorkshop.find(cls.tags_to_gen[i])

            if fields["homologacion"] != "" and int(cls.get_hebras(fields["homologacion"].split(" ")[0])) <= 1:
                continue
            if fields["n_hebras"] != "":
                if int(fields["n_hebras"]) <= 1:
                    continue
            elif int(cls.get_hebras(fields["tipo_cable"].split(" ")[0])) <= 1:
                continue

            cls.inprint(fields, cls.corrs[i])
            cls.working_template.save(
                TrinityWorkshop.destiny_pth + "\\" + cls.corrs[i] + "-RA_" + fields["tag"] + ".xlsx")
        pass

    @classmethod
    def flush(cls):
        for i in range(0, 16):
            row = 12 + (2 * i)
            for j in range(0, 16):
                if i == j:
                    continue
                col = 8 + (2 * j)
                cls.working_ws.cell(row, col).value = ""
                pass

            cls.working_ws.cell(row, 40).value = ""
            cls.working_ws.cell(row, 41).value = ""
            cls.working_ws.cell(row, 42).value = ""
        pass


class PuntoPuntoWorkshop:
    class Hebra:

        def __init__(self):
            self.desde = ""
            self.desde_bornera = ""
            self.desde_borne = ""
            self.hasta = ""
            self.hasta_bornera = ""
            self.hasta_borne = ""
            pass

        pass

    class Circuito:
        def __init__(self):
            self.tag = ""
            self.puntas: list[PuntoPuntoWorkshop.Hebra] = []

        pass

    capacidades = None
    res_pth = TrinityWorkshop.res_pth + "\\PP"
    working_template: Workbook = openpyxl.load_workbook(res_pth + "\\Protocolo Punto a Punto y Conexionado.xlsx")
    working_ws: Worksheet = working_template.worksheets[0]
    ok = True

    tags_to_gen = []
    corrs = []
    fl = open(TrinityWorkshop.origin_pth + "\\out.txt")
    for line in fl:
        pts = line.strip().split("\t")
        tags_to_gen.append(pts[0])
        corrs.append(pts[1])
        pass

    fields: dict[str: Cell] = {
        "nivel_tension": working_ws.cell(12, 9),
        "plano": working_ws.cell(13, 9),
        "correlativo": working_ws.cell(12, 23),
        "fecha": working_ws.cell(13, 23),
        "check_control": working_ws.cell(17, 7),
        "check_alumbrado": working_ws.cell(17, 16),
        "check_fuerza": working_ws.cell(17, 23),
        "check_pantalla_y": working_ws.cell(17, 29),
        "check_pantalla_n": working_ws.cell(17, 33),
        "tag": working_ws.cell(19, 9),
        "tag_2": working_ws.cell(36, 17),
        "seccion": working_ws.cell(20, 9),
        "aislacion": working_ws.cell(21, 9),
        "tension_serv": working_ws.cell(22, 9),
        "desde": working_ws.cell(19, 23),
        "desde_2": working_ws.cell(25, 4),
        "hasta": working_ws.cell(20, 23),
        "hasta_2": working_ws.cell(25, 32),
        "longitud": working_ws.cell(21, 23),
        "capacidad": working_ws.cell(22, 23),
        "instrumento_tipo": working_ws.cell(50, 26),
        "instrumento_marca": working_ws.cell(51, 26),
        "instrumento_modelo": working_ws.cell(52, 26),
        "instrumento_serie": working_ws.cell(53, 26),
        "instrumento_calibracion": working_ws.cell(54, 26),
        "elabora_nombre": working_ws.cell(62, 4),
        "elabora_cargo": working_ws.cell(63, 4),
        "elabora_fecha": working_ws.cell(64, 4),
        "revisa_nombre": working_ws.cell(62, 12),
        "revisa_cargo": working_ws.cell(63, 12),
        "revisa_fecha": working_ws.cell(64, 12)
    }

    circuitos: list[Circuito] = []

    @classmethod
    def fetch(cls):
        # todo refactor
        cls.capacidades = []
        tmp = open(cls.res_pth + "\\tabla corrientes.txt")
        for l in tmp:
            t = l.split(" ")
            cls.capacidades.append([t[0], t[1], t[2]])
            pass
        tmp.close()
        tmp = open(cls.res_pth + "\\puntas.txt")
        for l in tmp:
            t = l.split("\t")
            circ_nuevo = True
            hebra = PuntoPuntoWorkshop.Hebra()
            hebra.desde = t[1]
            hebra.desde_bornera = t[2]
            hebra.desde_borne = t[3]
            hebra.hasta = t[4]
            hebra.hasta_bornera = t[5]
            hebra.hasta_borne = t[6].strip()
            for circ in cls.circuitos:
                if circ.tag == t[0]:
                    circ.puntas.append(hebra)
                    circ_nuevo = False
                    break
                pass
            if circ_nuevo:
                nuevo = PuntoPuntoWorkshop.Circuito()
                nuevo.tag = t[0]
                nuevo.puntas.append(hebra)
                cls.circuitos.append(nuevo)
            pass
        pass

    @classmethod
    def gen(cls):
        for i in range(len(cls.tags_to_gen)):
            cls.ok = True
            cls.flush()
            fields = TrinityWorkshop.find(cls.tags_to_gen[i])
            cls.inprint(fields, cls.corrs[i])

            if cls.ok:
                name = cls.corrs[i] + "-PP_" + fields["tag"] + ".xlsx"
                pth = TrinityWorkshop.destiny_pth + "\\" + name
                cls.working_template.save(pth)
            else:
                name = "COMPLETAR "+cls.corrs[i] + "-PP_" + fields["tag"] + ".xlsx"
                pth = TrinityWorkshop.destiny_pth + "\\" + name
                cls.working_template.save(pth)
        pass

    @classmethod
    def flush(cls):
        for i in range(0, 3):
            col = 2 + (2 * i)
            col_2 = 30 + (2 * i)
            for j in range(0, 17):
                row = 28 + j
                cls.working_ws.cell(row, col).value = ""
                cls.working_ws.cell(row, col_2).value = ""
        pass

    @classmethod
    def inprint(cls, fields, corr):
        cls.fields["nivel_tension"].value = cls.resolve_tension(fields["uso"])
        cls.fields["tension_serv"].value = cls.resolve_tension(fields["uso"])

        cls.fields["plano"].value = TrinityWorkshop.resolve_plano(fields["ubicacion"], fields["desde"], fields["hasta"])
        cls.fields["correlativo"].value = corr
        cls.fields["fecha"].value = "29/06/2023"
        cls.toggle_checks(fields["tag"])
        cls.fields["tag"].value = fields["tag"]
        cls.fields["tag_2"].value = fields["tag"]
        cls.fields["desde"].value = fields["desde"]
        cls.fields["desde_2"].value = fields["desde"]
        cls.fields["hasta"].value = fields["hasta"]
        cls.fields["hasta_2"].value = fields["hasta"]
        cls.fields["longitud"].value = fields["largo"]
        cls.fields["instrumento_tipo"].value = "Pinza Amperimetrica"
        cls.fields["instrumento_marca"].value = "Fluke"
        cls.fields["instrumento_modelo"].value = "376"
        cls.fields["instrumento_serie"].value = "57194715MV"
        cls.fields["instrumento_calibracion"].value = "13-10-2022"
        cls.fields["elabora_nombre"].value = "Jose Godoy Espinoza"
        cls.fields["elabora_cargo"].value = "Supervisor Eléctrico"
        cls.fields["elabora_fecha"].value = ""
        cls.fields["revisa_nombre"].value = "Claudio Boris H."
        cls.fields["revisa_cargo"].value = "Jefe Terreno"
        cls.fields["revisa_fecha"].value = ""

        if fields["homologacion"] == "":
            pts = fields["tipo_cable"].split(" ")
            cls.fields["capacidad"].value = cls.resolve_capacidad(pts[0], pts[1]) + " A"
            cls.fields["seccion"].value = pts[0] + " " + pts[1]
            cls.fields["aislacion"].value = pts[2]
        else:
            pts = fields["homologacion"].split(" ")
            cls.fields["capacidad"].value = cls.resolve_capacidad(pts[0], pts[1])
            cls.fields["seccion"].value = pts[0] + " " + re.sub("Â²", "²", pts[1])
            try:
                cls.fields["aislacion"].value = pts[2]
            except IndexError:
                pts = fields["tipo_cable"].split(" ")
                cls.fields["aislacion"].value = pts[2]

        cls.populate_pp(fields["tag"], fields["desde"], fields["hasta"], fields["n_hebras"])

        pass

    @classmethod
    def resolve_tension(cls, uso):
        if uso == "C":
            return "125 V"
        else:
            return "220 V"
        pass

    @classmethod
    def toggle_checks(cls, uso):
        if uso == "C":
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

    @classmethod
    def resolve_capacidad(cls, calibre, unidad):
        offset = 0
        if re.match("AWG", unidad, re.IGNORECASE):
            offset += 1
        seccion = cls.get_seccion(calibre)
        for e in cls.capacidades:
            if e[offset] == seccion:
                return e[2]
        raise Exception("calibre no encontrado")
        pass

    @classmethod
    def populate_pp(cls, tag, desde, hasta, n_hebras):
        circ = cls.buscar_circ(tag)
        if not cls.ok:
            return
        lock = 0
        if circ.puntas[0].desde != desde:
            lock = 1
        for i in range(int(len(circ.puntas) / 2)):
            row = 28 + i
            desde_bornera = cls.working_ws.cell(row, 2)
            desde_borne = cls.working_ws.cell(row, 4)
            desde_n_hebra = cls.working_ws.cell(row, 6)
            hasta_bornera = cls.working_ws.cell(row, 34)
            hasta_borne = cls.working_ws.cell(row, 32)
            hasta_n_hebra = cls.working_ws.cell(row, 30)

            idx = i + (int((len(circ.puntas) / 2)) * lock)
            desde_bornera.value = circ.puntas[idx].desde_bornera
            desde_borne.value = circ.puntas[idx].desde_borne
            desde_n_hebra.value = i + 1
            hasta_bornera.value = circ.puntas[idx].hasta_bornera
            hasta_borne.value = circ.puntas[idx].hasta_borne
            hasta_n_hebra.value = i + 1
        pass

    @classmethod
    def get_seccion(cls, calibre):
        if "c" in calibre:
            pts = calibre.strip().split("-c-")
            return pts[1].replace("-", "/")
        else:
            pts = calibre.strip().split("x")
            return pts[len(pts) - 1].replace("-", "/")
        pass

    @classmethod
    def buscar_circ(cls, tag) -> Circuito:
        for circ in cls.circuitos:
            if circ.tag == tag:
                return circ
            pass
        cls.ok = False
        # raise Exception("CIRCUTO " + tag + " no encontrado")
        pass


kinds = [
    "T",
    "RA",
    "PP",
]


def start(kinds: list[str]):
    TrinityWorkshop.fetch()
    if "T" in kinds:
        TendidoWorkshop.gen()
    if "RA" in kinds:
        AislacionWorkshop.fetch()
        AislacionWorkshop.gen()
    if "PP" in kinds:
        PuntoPuntoWorkshop.fetch()
        PuntoPuntoWorkshop.gen()
    pass


start(kinds)

pass
