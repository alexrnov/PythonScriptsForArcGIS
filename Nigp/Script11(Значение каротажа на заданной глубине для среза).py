import sys, arcgisscripting, os
import arcpy as ARCPY
gp = arcgisscripting.create(9.3)

#функция на входе берет точечный слой со значениями каротажа, который формируется в скрипте Script10
#и возвращает массив [название скважины, x,y,глубина измерений, значения измерений]
def read_map_layer(layer):
        def check_unicode(text):#функция преобразует кодировку юникод в utf-8 
                if(type(text)==unicode):text = text.encode('utf-8')
                return text
        rows = gp.SearchCursor(layer)
        row = rows.Next()
        layer = gp.describe(layer)
        arr_wells=[]#массив для записывания скважин
        while row:
                feat = row.GetValue(layer.ShapeFieldName)
                point = feat.getpart(0)
                id_well = check_unicode(row.GetValue("id_well"))#id скважины
                name_well = check_unicode(row.GetValue("name_well"))#название скважины
                x = point.x
                y = point.y
                z = check_unicode(row.GetValue("z"))#abs устья скважины
                depth_well = check_unicode(row.GetValue("depth_well"))#глубина скважины
                name_object = check_unicode(row.GetValue("name_object"))#название объекта
                name_area = check_unicode(row.GetValue("name_area"))#название участка
                depth_logging = check_unicode(row.GetValue("depth_logging"))#строка, которая содержит значения глубин
                value_logging = check_unicode(row.GetValue("value_logging"))#строка, которая содержит измеренные значения ГИС
                recover = check_unicode(row.GetValue("recover"))#мощность перекрывающих отложений
                abs_litostrat_roof = check_unicode(row.GetValue("abs_litostrat_roof"))#abs кровли литостратиграфических пластов
                abs_litostrat_sole = check_unicode(row.GetValue("abs_litostrat_sole"))#abs подошвы литостратиграфических пластов
                stratigraphy = check_unicode(row.GetValue("stratigraphy"))#стратиграфия пластов
                lithology = check_unicode(row.GetValue("lithology"))#литология пластов
                #добавить в выходной массив сведения о текущей скважине
                arr_wells.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,depth_logging,\
                value_logging,recover,abs_litostrat_roof,abs_litostrat_sole,stratigraphy,lithology])
                row = rows.next()
        #Возвращает массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,
        #глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений,
        #глубина кровли пласта/глубина кровли пласта/глубина кровли пласта,Jmh/Jsn/Ol/,алевролит/известняк/доломит/]
        return arr_wells 
#функция на входе берет массив, считанный из картографического слоя и значение глубины
#и возвращает массив в формате [название скважины, x,y, значение скважинной магниторазведки на заданной глубине]
def select_value_depth(arr_map,given_depth):
        k=0
        arr_output=[]
        while(k<len(arr_map)):#перебор элементов массива из картографического слоя
                id_well = str(arr_map[k][0])#id скважины
                z = str(arr_map[k][4])#abs устья скважины
                if(z<>'')and(z<>' ')and(z<>'null'):#если есть данные по abs устья текущей скважины
                        str_depth_logging = str(arr_map[k][8])#строка содержит глубины измерений ГИС для текущей скважины
                        str_value_logging = str(arr_map[k][9])#строка содержит значения измерений ГИС для текущей скважины
                        arr_depth_logging = str_depth_logging.split("/")#получить массив глубин измерений ГИС для текущей скважины
                        arr_value_logging = str_value_logging.split("/")#получить массив значений ГИС для текущей скважины 
                        if(len(arr_depth_logging)==len(arr_value_logging)):#если массивы глубин ГИС и измеренных значений ГИС равны
                                z = round(float(z),2)#преобразовать в вещественный тип с двумя знаками после запятой
                                given_depth = round(float(given_depth),2)#преобразовать в вещественный тип с двумя знаками после запятой
                                #вычислить разницу от устья скважины до заданной глубины в единицах балтийской системы высот
                                depth_to_val = z - given_depth
                                #если значение больше нуля, значит устье скважины находится выше
                                #заданной глубины и следовательно можно проводить вычисление
                                if (depth_to_val>0):
                                        m=0

                                        #перебор значений глубин измерений ГИС
                                        while(m<len(arr_depth_logging)-1):#не обрабытавать последний элемент массива - поскольку '1/2/3/' дает ['1','2','3','']
                                                similar = depth_to_val-round(float(arr_depth_logging[m]),2)
                                                ARCPY.AddMessage(str(depth_to_val)+" - "+str(arr_depth_logging[m])+" = "+str(similar))
                                                if(similar<0):
                                                        ARCPY.AddMessage("Отрицательное значение")
                                                        break
                                                m+=1
             
                                
                        else:
                                ARCPY.AddMessage("Массивы глубин измерений ГИС и самих измеренных \
                                значений не равны, код скважины: "+str(id_well))
                ARCPY.AddMessage("------------------------------")
                k=k+1
        return arr_output #вернуть массив [name_well,x,y,val]
#Функция на входе берет: название выходного точечного слоя, итоговый массив для записи, и пространственную привязку
#и записывет картографический слой 
def write_layer_sm(FC,arr_points,spatial_ref):
        fc_path,fc_name = os.path.split(FC)
        gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",spatial_ref)
        gp.AddField(FC,"name","STRING",15,15)
        gp.AddField(FC,"val_sm","FLOAT",15,15)
        rows = gp.InsertCursor(FC)
        pnt = gp.CreateObject("Point")
        k=0
        while(k<len(arr_points)):#создание точечного слоя
                feat = rows.NewRow()#Новая строка
                pnt.id=k#ID
                pnt.x=arr_points[k][1]
                pnt.y=arr_points[k][2]
                feat.shape=pnt
                feat.SetValue("name",arr_points[k][0])#название скважины
                feat.SetValue("val_sm",float(arr_points[k][3]))#значение скважинной магниторазведки
                rows.InsertRow(feat)
                k=k+1
if __name__== '__main__':#если модуль запущен, а не импортирован
        #-------------------------------------входные параметры----------------------------------------
        InpPFC = ARCPY.GetParameterAsText(0)#входной точечный слой скважин
        OutputFC = arcpy.GetParameterAsText(1)#выходной точечный слой скважин 
        depth = ARCPY.GetParameterAsText(2)#атрибут глубины измеренных значений скважинной магниторазведки
        Full_report = ARCPY.GetParameterAsText(3)#параметр определяет выводить ли полный отчет
        #----------------------------------------------------------------------------------------------
        arr_map_points = read_map_layer(InpPFC)#извлечь поля точечного слоя
        if(len(arr_map_points)==0):#если массив точек - завершить работу программы
                raise SystemExit(1)#завершить работу программы
        ARCPY.AddMessage("Количество точек картографического слоя: "+str(len(arr_map_points)))
        if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате:")
                string_point=""
                m=0
                while(m<len(arr_map_points[0])):
                        string_point=string_point+" "+str(arr_map_points[0][m])
                        m+=1
                ARCPY.AddMessage(string_point)
        arr_select = select_value_depth(arr_map_points,depth)     
        
        
'''
gp.AddMessage("Количество скважин, считанных из слоя = "+str(len(arr_sm)))
arr_select = select_value(arr_sm,depth)#получить массив точек со значениями скважинной магниторазведки на заданной глубине
gp.AddMessage("Количество скважин, с данными магнитной восприимчивости для заданной глубины = "+str(len(arr_select)))
gp.AddMessage("В формате:")
k=0
while(k<10):
        gp.AddMessage(arr_select[k])
        k=k+1
gp.AddMessage("------------------------------------------------------------")
Spat_ref = gp.describe(InpPFC).spatialreference
write_layer_sm(OutputFC,arr_select,Spat_ref)#записать итоговый массив к картографический слой
'''
