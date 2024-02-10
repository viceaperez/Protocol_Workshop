from pyautocad import Autocad

acad = Autocad(create_if_not_exists=True)

# acad.prompt("Hello, Autocad from Python\n")

# switch_caching(True)
a = acad.doc
lineas = []
tipos_cruz = []
tipos_t = []

tramos = []

ignored_layers = [
        "OTROS",
        "Chicotes",
        "0"
    ]

for i in acad.iter_objects():
    if i.Layer in ignored_layers:
        continue
    if hasattr(i, "EntityName"):
        if i.EntityName == "AcDbLine":
            lineas.append(i)
        continue
    pass

largo_cable_4 = 0
for i in lineas:
    largo_cable_4 += i.Length
    pass
largo_cable_4 = largo_cable_4 / 1000  # de mm a m
pass
