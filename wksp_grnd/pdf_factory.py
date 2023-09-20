import os

from win32com import client

project_pth: str = os.getcwd()
res_pth: str = project_pth + "\\res"
destiny_pth: str = project_pth + "\\destiny_files"
origin_pth: str = project_pth + "\\origin_files"

fl_lst = os.listdir(origin_pth)

excel = client.Dispatch("Excel.Application")
for file in fl_lst:
    pth = origin_pth + "\\" + file
    sheets = excel.Workbooks.Open(pth)
    work_sheets = sheets.Worksheets[0]
    out_pth = destiny_pth + "\\" + file.replace(".xlsx", ".pdf")
    work_sheets.ExportAsFixedFormat(0, out_pth)
excel.Quit()
