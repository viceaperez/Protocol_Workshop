import os

import openpyxl
from openpyxl.cell import Cell
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet


class Table:
    def __init__(self, nw: Cell, ne: Cell, sw: Cell, se: Cell, data: Cell):
        self.title = None
        self.title_cell = None
        self.nw_cell: Cell = nw
        self.ne_cell: Cell = ne
        self.sw_cell: Cell = sw
        self.se_cell: Cell = se
        self.get_title()
        self.get_subtitle()
        self.get_headers()

    def belongs(self, cell: Cell) -> bool:
        if self.nw_cell.column > cell.column:
            return False
        if self.nw_cell.row < cell.row:
            return False
        if self.ne_cell.column < cell.column:
            return False
        if self.sw_cell.row > cell.row:
            return False
        return True

    pass

    def get_title(self):
        self.title_cell: Cell = self.nw_cell
        #TODO: seguir aquí
        self.title_cell.
        self.title: str = self.title_cell.value
        pass


class Alumb:
    correlative_counter = 0

    def __init__(self):
        self.tag: (str, None) = None
        self.patio: (str, None) = None
        self.ubicacion: (str, None) = None
        self.tipo: (str, None) = None
        self.corr: int = -1
        pass

    pass


class AlumbradoWorkshop:
    project_pth: str = os.getcwd()
    res_pth: str = project_pth + "\\res"
    destiny_pth: str = project_pth + "\\destiny_files"
    src_pth = res_pth + "\\matriz.xlsx"

    source: Worksheet = openpyxl.load_workbook(res_pth, data_only=True, read_only=True).woksheets[0]

    src_sheets: dict[str:Worksheet] = {
        "Caminos interiores": source["CAMINOS INTERIORES"],
        "PK": source["P. 500KV"],
        "PATR": source["P.ATR"],
        "PZ": source["P.REACTORES"],
        "PJ": source["P.220KV"],
    }

    working_template: Workbook = openpyxl.load_workbook(
        res_pth + "\\Protocolo de montaje de luminarias  SSEE Parinas.xlsx")

    starting_corr = 0

    data: list[Alumb] = []

    @classmethod
    def fetch(cls):
        for sh in cls.src_sheets:
            sh: Worksheet
            headers: list[Cell] = cls.fetch_headers(sh)

            pass

    @classmethod
    def fetch_headers(cls, sh: Worksheet) -> list[Cell]:
        result: list[Cell] = []
        max_row = sh.max_row
        max_col = sh.max_column
        col_idx = 0
        while col_idx < max_row:
            row_idx = 0
            while row_idx < max_col:
                c: Cell = sh.cell(row_idx, col_idx)
                if c.value in None:
                    continue
                    pass

                result.append(c)

                col_idx += 1
                pass

            row_idx += 1
            pass
        return result
        pass


pass
