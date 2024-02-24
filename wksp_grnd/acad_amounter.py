from enum import Enum

from pyautocad import Autocad


class Tramo(Enum):
    PJ07 = "PJ07"
    PJ08 = "PJ08"
    PJ09_12 = "PJ09-12"
    PJ10 = "PJ10"
    PJ11 = "PJ11"
    PJ13 = "PJ13"
    PJ14 = "PJ14"
    PJ15 = "PJ15"
    PJ16 = "PJ16"
    PJ17 = "PJ17"
    PJ18 = "PJ18"
    PF = "PFuturo"
    CERCO = "Cerco"
    CI = "Caminos Interiores"
    pass


class Sector:
    def __int__(self):
        self.tramo = None
        self.pool = []
        self.lineas = []
        self.largo = 0
        self.conectT4_2 = []
        self.conectT4_4 = []
        self.conectCruz = []

        self.lineas_r = []
        self.largo_r = 0
        self.conectT4_2_r = []
        self.conectT4_4_r = []
        self.conectCruz_r = []
        pass

    pass


class AcadMan:
    acad = Autocad(create_if_not_exists=True)

    ignored_layers = [
        "OTROS",
        "Chicotes",
        "0",
        "Anotaciones",
        "Defpoints"
    ]

    objs = []
    sectores: list[Sector] = []

    @classmethod
    def go(cls):
        for i in cls.objs:
            result = i.Layer
            for j in cls.sectores:
                if j.tramo.value == result:
                    cls.go_layer(j, i)
                elif j.tramo.value + "-listo" == result:
                    cls.go_listo(j, i)
                pass
            pass
        cls.consistency()
        pass

    @classmethod
    def purge(cls):
        for i in cls.objs:
            if i.EntityName == "AcDbLine":
                if i.Length <= 2000:
                    i.Delete()
            pass
        pass

    @classmethod
    def fetch(cls):
        for i in cls.acad.iter_objects():
            if i.Layer in cls.ignored_layers:
                continue
            cls.objs.append(i)
            pass
        pass

    pass

    @classmethod
    def init(cls):
        for tipo in Tramo:
            temp = Sector()
            temp.tramo = tipo
            temp.pool = []
            temp.lineas = []
            temp.largo = 0
            temp.conectT4_4 = []
            temp.conectT4_2 = []
            temp.conectCruz = []

            temp.lineas_r = []
            temp.largo_r = 0
            temp.conectT4_2_r = []
            temp.conectT4_4_r = []
            temp.conectCruz_r = []
            cls.sectores.append(temp)
            pass
        cls.fetch()
        pass

    @classmethod
    def consistency(cls):
        for i in cls.sectores:
            if len(i.lineas) + len(i.conectT4_2) + len(i.conectT4_4) + len(i.conectCruz) == len(i.pool):
                i.consistencia = True
                pass
            else:
                i.consistencia = False
                pass
            for j in i.lineas:
                i.largo += j.Length
                pass
            for j in i.lineas_r:
                i.largo_r += j.Length
                pass
            i.largo = i.largo / 1000
            i.largo_r = i.largo_r / 1000
        pass

    @classmethod
    def go_listo(cls, j, i):
        j.pool.append(i)
        if cls.is_line(i):
            j.lineas_r.append(i)
            return
        elif cls.is_conectT4_4(i):
            j.conectT4_4_r.append(i)
            return
        elif cls.is_conectT4_2(i):
            j.conectT4_2_r.append(i)
            return
        elif cls.is_conectCruz(i):
            j.conectCruz_r.append(i)
            return
        pass

    @classmethod
    def is_line(cls, i):
        return i.EntityName == "AcDbLine"

    @classmethod
    def is_conectT4_4(cls, i):
        return i.EntityName == "AcDbHatch" and 190100 >= i.Area >= 189900

    @classmethod
    def is_conectT4_2(cls, i):
        return i.EntityName == "AcDbHatch" and 170100 >= i.Area >= 159900

    @classmethod
    def is_conectCruz(cls, i):
        return i.EntityName == "AcDbHatch" and 240100 >= i.Area >= 239900

    @classmethod
    def go_layer(cls, j, i):
        j.pool.append(i)
        if cls.is_line(i):
            j.lineas.append(i)
            return
        elif cls.is_conectT4_4(i):
            j.conectT4_4.append(i)
            return
        elif cls.is_conectT4_2(i):
            j.conectT4_2.append(i)
            return
        elif cls.is_conectCruz(i):
            j.conectCruz.append(i)
            return
        pass

    @classmethod
    def out(cls):
        fl = open("out.txt", "w")
        fl.write(
            "Sector\tlargo\tconectT4_2\tconectT4_4\tconectCruz\tlargo listo\tconectT4_2 listo\tconectT4_4 listo\tconectCruz listo\n")
        for i in cls.sectores:
            fl.write(i.tramo.value + "\t")
            fl.write(str(i.largo).replace(".", ",") + "\t")
            fl.write(str(len(i.conectT4_2)).replace(".", ",") + "\t")
            fl.write(str(len(i.conectT4_4)).replace(".", ",") + "\t")
            fl.write(str(len(i.conectCruz)).replace(".", ",") + "\t")
            fl.write(str(i.largo_r).replace(".", ",") + "\t")
            fl.write(str(len(i.conectT4_2_r)).replace(".", ",") + "\t")
            fl.write(str(len(i.conectT4_4_r)).replace(".", ",") + "\t")
            fl.write(str(len(i.conectCruz_r)).replace(".", ",") + "\n")
            pass
        fl.close()
        pass


AcadMan.init()
AcadMan.go()
AcadMan.out()
pass
