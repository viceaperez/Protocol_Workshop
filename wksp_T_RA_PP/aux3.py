import os

import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from docx import Document

project_pth: str = os.getcwd()
res_pth: str = project_pth + "\\res"
destiny_pth: str = project_pth + "\\destiny_files"
origin_pth: str = project_pth + "\\origin_files"

"""
lst = os.listdir(origin_pth)

lst_log: dict[str:str] = {}
out = open(destiny_pth + "\\transmittal.txt", "w")
for e in lst:
    corr = e.split("-")[0]
    tag = e.split("_")[1].strip(".pdf")
    tipo = e.split("-")[1][0]
    out.write(tag + "\t" + str(corr) + "\t" + str(tipo) + "\n")
    pass

out.close()
"""

out = open(destiny_pth + "\\transmittal.txt", "r")
mem: list[dict] = []
for line in out:
    pts = line.strip().split("\t")
    mem.append({
        "tag": pts[0],
        "corr": pts[1],
        "tipo": pts[2]
    })
    pass

doc = Document(destiny_pth + "\\Formato transmittal Elec.docx")
table = doc.tables[0]
i = 0
for e in mem:
    i = i + 1
    row_cells = table.add_row().cells
    row_cells[0].text = str(i)
    if e["tipo"] == "T":
        row_cells[3].text = "Protocolo Tendido " + e["tag"]
        pass
    else:
        row_cells[3].text = "Protocolo Prueba Aislación " + e["tag"]
        pass
    row_cells[1].text = e["corr"]
    row_cells[2].text = "0"
    row_cells[4].text = "Pendiente compaginación con Protocolo Punto-Punto y Plano"
    pass
doc.save(destiny_pth + "\\Formato transmittal Elec.docx")
pass
