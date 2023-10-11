import os

import openpyxl
from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class TrinityWorkshop:
    project_pth: str = os.getcwd()
    res_pth: str = project_pth + "\\res"
    destiny_pth: str = project_pth + "\\destiny_files"
    src_pth = res_pth + "\\MATRIZ DE TENDIDO.xlsx"
    circ_source: Worksheet = openpyxl.load_workbook(res_pth, data_only=True, read_only=True).woksheets[0]

    starting_corr = 0

    @classmethod
    def gen(cls):
        max_row = cls.circ_source.max_row
        for row in range(7, max_row):
            fields: dict[str:Cell] = {
                "tag": cls.circ_source.cell(row, 2),
                "tipo_cable": cls.circ_source.cell(row, 3),
                "desde": cls.circ_source.cell(row, 4),
                "hasta": cls.circ_source.cell(row, 5),
                "n_hebras": cls.circ_source.cell(row, 7),
                "n_ptas": cls.circ_source.cell(row, 8),
                "largo": cls.circ_source.cell(row, 13),
                "homologacion": cls.circ_source.cell(row, 20)
            }
            fields["t"]

            pass

        pass

    pass


class TendidoWorkshop:
    res_pth = TrinityWorkshop.res_pth + "\\T"
    working_template: Workbook = openpyxl.load_workbook(res_pth + "\\Protocolo Tendido de conductores electricos.xlsx")
    working_ws: Worksheet = working_template.worksheets[0]

    @classmethod
    def set(cls, r, c, val):
        cls.working_ws.cell(r, c, val)
        pass

    fields: dict[str:Cell] = {
        "correlativo": working_ws.cell(10, 23),
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
    def inprint(cls):
        pass

    @classmethod
    def gen(cls):
        cls.fetch()
        cls.inprint()
        pass

    pass


class AislacionWorkshop:
    pass


class PuntoPuntoWorkshop:
    pass


pass
