from pyautocad import Autocad

acad = Autocad(create_if_not_exists=True)

#acad.prompt("Hello, Autocad from Python\n")

#switch_caching(True)
a = acad.doc
b = a.Blocks
c = []
for block in b:
    c.append(block)
    pass
#d = acad.iter_objects()
d = []
lineas =[]
tipos_cruz = []
tipos_t = []
for i in acad.iter_objects():
    if not hasattr(i,"EffectiveName"):
        lineas.append(i)
        continue
    if i.EffectiveName == "Conex. Tipo Cruz para clable 4.0":
        tipos_cruz.append(i)
        continue
        pass
    if i.EffectiveName == "Conex. Tipo T. para cable 4.0":
        tipos_t.append(i)
        continue
        pass
    d.append(i)
    pass
largo_cable_4 = 0
for i in lineas:
    largo_cable_4 += i.Length
    pass
largo_cable_4 = largo_cable_4/1000 # de mm a m
pass