from xlrd import open_workbook

class ExcelDosyaIslemleri:
    def __init__(self,dosyaYolu):
        try:
            self.__dosya=open_workbook(dosyaYolu)
            self.__matris=self.__ExcelToMatris()
        except Exception:
            print("Dosya acilamadi veya bulunamadi!")

    def __ExcelToMatris(self):
        sayfa1=self.__dosya.sheets()[0]
        matris=[]
        for satir in range(sayfa1.nrows):
            satirTemp=[]
            for sutun in range(sayfa1.ncols):
                satirTemp.append(sayfa1.cell(satir,sutun).value)
            matris.append(satirTemp)
        return matris
    
    def GetMatris(self):
        return self.__matris

    def GetSatir(self,satirIndex):
        return self.__matris[satirIndex]

    def GetParametreIsimleri(self):
        return self.__matris[0]
    
    def GetSutunVeriTipi(self,sutunIndex):
        return type(self.__matris[0][sutunIndex])

excel=ExcelDosyaIslemleri("deneme.xlsx")
print("---------------Matris----------------------")
print(excel.GetMatris())
print("---------------Matris 1 satir----------------------")
print(excel.GetSatir(1))
print("---------------Parametere isimleri----------------------")
print(excel.GetParametreIsimleri())
print("---------------Matris----------------------")
print(excel.GetParametreIsimleri())
print("---------------Matris----------------------")
print(excel.GetParametreIsimleri())
print("---------------Matris----------------------")
print(excel.GetParametreIsimleri())


    