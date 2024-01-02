import os

project_pth: str = os.getcwd()
res_pth: str = project_pth + "\\res"
destiny_pth: str = project_pth + "\\destiny_files"
origin_pth: str = project_pth + "\\origin_files"

ls = os.listdir(destiny_pth)
for f in ls:
    os.startfile(destiny_pth + "\\" + f, "print")
