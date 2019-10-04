import pythoncom
from win32com.client import Dispatch, gencache
import LDefin2D
import MiscellaneousHelpers as MH


class OpenKompas():
    def __init__(self, dmax, dmin, z1):
        self.dmax = dmax
        self.dmin = dmin
        self.z1 = z1 * 1000
        self.width = (float(self.dmax) - float(self.dmin)) * 1000

    def kompas(self):
        #  Подключим константы API Компас
        self.kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

        #  Подключим описание интерфейсов API5
        self.kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        self.kompas_object = self.kompas6_api5_module.KompasObject(
            Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(self.kompas6_api5_module.KompasObject.CLSID,
                                                                     pythoncom.IID_IDispatch))
        MH.iKompasObject = self.kompas_object

        #  Подключим описание интерфейсов API7
        kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        application = kompas_api7_module.IApplication(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                                     pythoncom.IID_IDispatch))
        MH.iApplication = application
        application.Visible = True

        # Определяем версию Компаса (API5)
        version = str(self.kompas_object.ksGetSystemVersion()[1]) + '.' + str(
            self.kompas_object.ksGetSystemVersion()[2])
        print('Версия Компаса: ' + version)

        # Пробуем получить активный 2D-документ (API5)
        self.iDocument2D = self.kompas_object.ActiveDocument2D()

        # Если документа нет, его нужно создать
        if self.iDocument2D is None:
            # Получим коллекцию документов (API7)
            self.documents = application.Documents
            # Добавим в коллекцию новый документ типа "Чертеж" (1)
            self.new_document_2d = self.documents.Add(2, True)
            # Теперь снова устанавливаем на него ссылку для API5
            if self.new_document_2d is not None:
                print('Создан новый документ')
                self.iDocument2D = self.kompas_object.ActiveDocument2D()
            else:
                # Если не удалось создать документ - выходим
                print('Не удалось создать чертеж, работа завершена')
                exit(0)
        else:
            print('Построения в текущем документе')

    def line(self, x1, y1, x2, y2):
        '''Тонкая линия'''
        obj = self.iDocument2D.ksLineSeg(x1, y1, x2, y2, 2)  # 2 - стиль линии

    def rectangle(self, x, y, height):
        '''Прямоугольник'''
        iRectangleParam = self.kompas6_api5_module.ksRectangleParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_RectangleParam))
        iRectangleParam.Init()
        iRectangleParam.x = x
        iRectangleParam.y = y
        iRectangleParam.ang = 0  # Поворот прямоугольника на 0 градусов
        iRectangleParam.height = height
        iRectangleParam.width = self.width  # Ширина прямоугольника
        iRectangleParam.style = 1  # Стиль линии
        obj = self.iDocument2D.ksRectangle(iRectangleParam)

    def arrow(self, leader_x, leader_y, point_x, point_y, text):
        '''Стрелка'''
        iLeaderParam = self.kompas6_api5_module.ksLeaderParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_LeaderParam))
        iLeaderParam.Init()
        iLeaderParam.around = 0
        iLeaderParam.arrowType = 2
        iLeaderParam.cText0 = 0
        iLeaderParam.cText1 = 0
        iLeaderParam.cText2 = 1
        iLeaderParam.cText3 = 0
        iLeaderParam.dirX = 1
        iLeaderParam.signType = 0
        iLeaderParam.x = leader_x
        iLeaderParam.y = leader_y
        iPolylineArray = self.kompas6_api5_module.ksDynamicArray(iLeaderParam.GetpPolyline())
        iMathPointArray = self.kompas_object.GetDynamicArray(LDefin2D.POINT_ARR)
        iMathPointParam = self.kompas6_api5_module.ksMathPointParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_MathPointParam))
        iMathPointParam.Init()
        iMathPointParam.x = point_x
        iMathPointParam.y = point_y
        iMathPointArray.ksAddArrayItem(-1, iMathPointParam)
        iPolylineArray.ksAddArrayItem(-1, iMathPointArray)
        iLeaderParam.SetpPolyline(iPolylineArray)
        iTextLineArray = self.kompas6_api5_module.ksDynamicArray(iLeaderParam.GetpTextline())
        iTextLineParam = self.kompas6_api5_module.ksTextLineParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 7
        iTextItemArray = self.kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = self.kompas6_api5_module.ksTextItemParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = text
        iTextItemParam.type = 0
        iTextItemFont = self.kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "GOST type A"
        iTextItemFont.height = self.width // 6
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)
        iTextLineArray.ksAddArrayItem(-1, iTextLineParam)
        iLeaderParam.SetpTextline(iTextLineArray)
        iCutLine = self.iDocument2D.ksLeader(iLeaderParam)

    def text(self, x, y, text):
        '''Надпись'''
        iParagraphParam = self.kompas6_api5_module.ksParagraphParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_ParagraphParam))
        iParagraphParam.Init()
        iParagraphParam.x = x
        iParagraphParam.y = y
        iParagraphParam.ang = 0
        iParagraphParam.hFormat = 0
        iParagraphParam.vFormat = 0
        iParagraphParam.style = 1
        self.iDocument2D.ksParagraph(iParagraphParam)
        iTextLineParam = self.kompas6_api5_module.ksTextLineParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_TextLineParam))
        iTextLineParam.Init()
        iTextLineParam.style = 1
        iTextItemArray = self.kompas_object.GetDynamicArray(LDefin2D.TEXT_ITEM_ARR)
        iTextItemParam = self.kompas6_api5_module.ksTextItemParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_TextItemParam))
        iTextItemParam.Init()
        iTextItemParam.iSNumb = 0
        iTextItemParam.s = text
        iTextItemParam.type = 0
        iTextItemFont = self.kompas6_api5_module.ksTextItemFont(iTextItemParam.GetItemFont())
        iTextItemFont.Init()
        iTextItemFont.bitVector = 4096
        iTextItemFont.color = 0
        iTextItemFont.fontName = "GOST type A"
        iTextItemFont.height = self.width // 6
        iTextItemFont.ksu = 1
        iTextItemArray.ksAddArrayItem(-1, iTextItemParam)
        iTextLineParam.SetTextItemArr(iTextItemArray)
        self.iDocument2D.ksTextLine(iTextLineParam)
        obj = self.iDocument2D.ksEndObj()

    def hatching(self, x1, y1, x2, y2, ang, width):
        '''Штриховка'''
        iHatchParam = self.kompas6_api5_module.ksHatchParam(
            self.kompas_object.GetParamStruct(self.kompas6_constants.ko_HatchParam))
        iHatchParam.Init()
        iHatchParam.ang = ang
        iHatchParam.color = -16777216
        iHatchParam.sheeting = 1
        iHatchParam.step = self.width // 10
        iHatchParam.width = width
        iHatchParam.x = 0
        iHatchParam.y = 0
        iHatchParam.style = 0
        iHatchParam.boundaries = self.iDocument2D.ksNewGroup(1)
        self.iDocument2D.ksContour(1)
        self.iDocument2D.ksNewGroup(1)
        obj = self.iDocument2D.ksLineSeg(x1, y1, x2, y1, 1)
        obj = self.iDocument2D.ksLineSeg(x2, y1, x2, y2, 1)
        obj = self.iDocument2D.ksLineSeg(x2, y2, x1, y2, 1)
        obj = self.iDocument2D.ksLineSeg(x1, y2, x1, y1, 1)
        self.iDocument2D.ksEndGroup()
        obj = self.iDocument2D.ksEndObj()

        self.iDocument2D.ksEndGroup()
        self.iDocument2D.ksHatchByParam(iHatchParam)

