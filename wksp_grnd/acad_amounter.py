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
        pass

    pass


class AcadMan:
    acad = Autocad(create_if_not_exists=True)

    ignored_layers = [
        "OTROS",
        "Chicotes",
        "0"
    ]

    objs = []
    sectores: list[Sector] = []

    @classmethod
    def go(cls):
        for i in cls.objs:
            result = Tramo(i.Layer)
            for j in cls.sectores:
                if j.tramo != result:
                    continue
                j.pool.append(i)
                if i.EntityName == "AcDbLine":
                    j.lineas.append(i)
                    break
                    pass
                elif i.EntityName == "AcDbHatch":
                    if 190100 >= i.Area >= 189900:
                        j.conectT4_4.append(i)
                        break
                        pass
                    elif 170100 >= i.Area >= 159900:
                        j.conectT4_2.append(i)
                        break
                        pass
                    elif 240100 >= i.Area >= 239900:
                        j.conectCruz.append(i)
                        break
                        pass
                    pass
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
            cls.sectores.append(temp)
            pass
        cls.fetch()
        pass

    @classmethod
    def consistency(cls):
        for i in cls.sectores:
            if len(i.lineas)+len(i.conectT4_2)+len(i.conectT4_4)+len(i.conectCruz) == len(i.pool):
                i.consistencia = True
                pass
            else:
                i.consistencia = False
                pass
            for j in i.lineas:
                i.largo += j.Length
                pass
            i.largo = i.largo/1000
        pass


AcadMan.init()
AcadMan.go()
pass
