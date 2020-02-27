import sys, arcgisscripting, os, math
import arcpy as ARCPY
from numpy import median
gp = arcgisscripting.create(9.3)
#функция на входе берет точечный слой со значениями каротажа, который формируется в скрипте Script10_3
#и возвращает массив точек
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
                id_well = check_unicode(row.GetValue("id_скважины"))#id скважины
                name_well = check_unicode(row.GetValue("название_скважины"))#название скважины
                x = point.x
                y = point.y
                z = round(check_unicode(row.GetValue("z")),1)#abs устья скважины
                depth_well = check_unicode(row.GetValue("глубина_скважины"))#глубина скважины
                name_object = check_unicode(row.GetValue("название_объекта"))#название объекта
                name_area = check_unicode(row.GetValue("название_участка"))#название участка
                type_tn = check_unicode(row.GetValue("тип_точки_наблюдения"))#тип точки наблюдения
                status_doc = check_unicode(row.GetValue("состояние_документирования"))#состояние_документирования
                type_doc = check_unicode(row.GetValue("тип_документирования"))#тип_документирования
                depth_logging = check_unicode(row.GetValue("глубина_измерений_ГИС"))#строка, которая содержит значения глубин
                value_logging = check_unicode(row.GetValue("значения_ГИС"))#строка, которая содержит измеренные значения ГИС
                interval_abs = check_unicode(row.GetValue("интервалы_искомого_подразделения"))#интервалы_искомого_подразделения
                abs_litostrat_roof = check_unicode(row.GetValue("глубина_кровли_пластов"))#abs кровли литостратиграфических пластов
                abs_litostrat_sole = check_unicode(row.GetValue("глубина_подошвы_пластов"))#abs подошвы литостратиграфических пластов
                stratigraphy = check_unicode(row.GetValue("стратиграфия"))#стратиграфия пластов
                lithology = check_unicode(row.GetValue("литология"))#литология пластов
                #добавить в выходной массив сведения о текущей скважине
                arr_wells.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,type_tn,status_doc,type_doc,\
                depth_logging,value_logging,interval_abs,abs_litostrat_roof,abs_litostrat_sole,stratigraphy,lithology])
                row = rows.next()
        #Возвращает массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,тип_точки_наблюдения,
        #состояние_документирования,тип_документирования,глубина ГИС/глубина ГИС...,значение ГИС/значение ГИС...,
        #интервалы искомых стратиграфических подразделений,глубина кровли пластов, глубина подошвы пластов,
        #Jmh/Jsn/Ol/,алевролит/известняк/доломит/]
        return arr_wells
def select_values(arr_map):
        arr_width_average=[]#выходной массив ТН со средними значениями ГИС для интервалов
        for current_point in arr_map:#перебор ТН
                if(current_point[11]=='нет данных')or(current_point[12]=='нет данных'):continue#если нет данных по ГИС
                if(current_point[13]=='-1.0'):continue#если для данной ТН нет искомого пласта - перейти к следующей ТН 
                intervals = current_point[13].split(";")#получить массив с данными интервалов искомых стратиграфических подразделений
                arr_depth_logging = current_point[11].split("/")#получить массив глубин измерений ГИС для текущей скважины
                arr_value_logging = current_point[12].split("/")#получить массив значений ГИС для текущей скважины
                col_dimension=0.0#суммарное количество измерений в искомых стратиграфических подразделениях
                sum_vales=0.0#сумма значений ГИС в искомых стратиграфических подразделениях
                arr_dimension_contain=[]#массив для записи глубин измерений в искомых стратиграфических пластах
                arr_value_contain=[]#массив для записи значений измерений в искомых стратиграфических пластах
                depth_interval=0#суммарное количество измерений в искомых стратиграфических пластах
                value_interval=0.0#суммарное значение каротажных измерений в искомых стратиграфических пластах
                m=0
                while m<len(intervals)-1:#перебор искомых интервалов, не обрабатывать последний элемент массива - поскольку '45;' даст ['45','']
                        #depth_interval=0#суммарное количество измерений в искомых стратиграфических пластах
                        #value_interval=0.0#суммарное значение каротажных измерений в искомых стратиграфических пластах
                        arr_int = intervals[m].split("-")#получить массив текущего интервала формата['10.0','10.5']
                        abs_roof = float(arr_int[0])#абс кровли текущего интервала искомого стратиграфического уровня
                        abs_sole = float(arr_int[1])#абс подошвы текущего интервала искомого стратиграфического уровня
                        if len(arr_depth_logging)<>len(arr_value_logging) :#если массивы глубин ГИС и измеренных значений ГИС не равны
                                ARCPY.AddMessage("Массивы глубин измерений ГИС и самих измеренных \
                                значений не равны, код скважины: "+str(current_point[0]))
                        else:
                                if(abs_sole-abs_roof>=0.1):#если мощность текущего подразделения больше 0.1
                                        bool_interval=False#логическая переменная определяет дошел ли цикл до искомого стратиграфического подразделения
                                        n=0
                                        #перебор глубин измерений ГИС текущей скважины(для текущего стратиграфического подразделения)
                                        while(n<len(arr_depth_logging)-1):#не обрабатывать последний элемент массива, поскольку 1/2/3 даст ['1','2','3','']
                                                depth_float = round(float(arr_depth_logging[n]),2)#преобразовать текущее значение глубины в вещественный тип
                                                val_float = round(float(arr_value_logging[n]),2)#преобразовать текущее значение измерений в вещественный тип
                                                #если цикл ранее дошел до искомого стратиграфического подразделения
                                                if(bool_interval):
                                                        #если текущая глубина измерений меньше подошвы коры выветривания
                                                        if(depth_float<=abs_sole):
                                                                #ARCPY.AddMessage(" 1 текущая глубина = "+str(depth_float)+" кровля пласта = "+str(abs_roof)+" подошва пласта = "+str(abs_sole))
                                                                #если значение - корректно
                                                                if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                        #добавить в массив глубин измерений в искомом стратиграфическом уровне
                                                                        #- текущую глубину измерений
                                                                        arr_dimension_contain.append(depth_float)
                                                                        #добавить в массив значений измерений во искомом стратиграфическом уровне
                                                                        #- текущее значение измерений ГИС
                                                                        arr_value_contain.append(val_float)
                                                                        #Прибавить к сумме каротажных измерений текущее значение
                                                                        value_interval=value_interval+val_float
                                                                        depth_interval+=1#увеличить количество измерений в подразделении
                                                        #цикл уже вышел за пределы коры выветривания - закончить цикл
                                                        else:break#выйти из цикла(перебор глубин для текущего стратиграфического подразделения)
                                                else:#если цикл еще не дошел до искомого подразделения
                                                        #если текущая глубина измерения ГИС равна или первый раз превысила
                                                        #abs кровли коры выветривания, значит цикл дошел до искомого стратиграфического подразделения
                                                        if(depth_float>=abs_roof)and(depth_float<=abs_sole):
                                                                #ARCPY.AddMessage("0 текущая глубина = "+str(depth_float)+" кровля пласта = "+str(abs_roof)+" подошва пласта = "+str(abs_sole))
                                                                bool_interval=True
                                                                #если значение - корректно
                                                                if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                        #добавить в массив глубин измерений в искомом стратиграфическом уровне
                                                                        #- текущую глубину измерений
                                                                        arr_dimension_contain.append(depth_float)
                                                                        #добавить в массив значений измерений во искомом стратиграфическом уровне
                                                                        #- текущее значение измерений ГИС
                                                                        arr_value_contain.append(val_float)
                                                                        #Прибавить к сумме каротажных измерений текущее значение
                                                                        value_interval=value_interval+val_float
                                                                        depth_interval+=1#увеличить количество измерений в подразделении                     
                                                n+=1
                        m+=1
                        
                #для текущей ТН измерения ГИС для интервалов найдены и сумма значений больше 0
                if(depth_interval>0)and(value_interval>0):
                        average_gis=round(value_interval/depth_interval,2)#вычислить среднее 
                        msd=0.0 #среднеквадратическое отклонение
                        standard_error = 0.0 #стандартная ошибка среднего
                        if (len(arr_dimension_contain)==len(arr_value_contain)):
                                i=0
                                while(i<len(arr_value_contain)):
                                        msd=msd+(arr_value_contain[i]-average_gis)**2
                                        i+=1
                                msd=math.sqrt(msd/depth_interval)
                                standard_error = msd/math.sqrt(depth_interval)#вычислить ошибку среднего
                        else:
                                ARCPY.AddMessage("количество элементов в массиве \
                                глубин измерений и массиве со значениями измерений не равно")
                        #ARCPY.AddMessage(str(current_point[0])+", сумма ГИС="+str(value_interval)+\
                        #" количество ГИС="+str(depth_interval)+ ", среднее="+str(round(value_interval/depth_interval,2)))
                        arr_width_average.append([current_point[0],current_point[1],current_point[2],\
                        current_point[3],current_point[4],current_point[5],current_point[6],current_point[7],
                        current_point[8],current_point[9],current_point[10],current_point[11],\
                        current_point[12],current_point[13],current_point[14],current_point[15],\
                        current_point[16],current_point[17],average_gis,depth_interval,msd,standard_error,median(arr_value_contain)])
        #Возвращает массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,тип_точки_наблюдения,
        #состояние_документирования,тип_документирования,глубина ГИС/глубина ГИС...,значение ГИС/значение ГИС...,
        #интервалы искомых стратиграфических подразделений,глубина кровли пластов, глубина подошвы пластов,
        #Jmh/Jsn/Ol/,алевролит/известняк/доломит/,среднее значение ГИС по искомым интервалам,количество измерений,
        #среднеквадратическое отклонение,ошибка среднего,медиана]
        return arr_width_average

#Возвращает массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,тип_точки_наблюдения,
#состояние_документирования,тип_документирования,глубина ГИС/глубина ГИС...,значение ГИС/значение ГИС...,
#интервалы искомых стратиграфических подразделений,глубина кровли пластов, глубина подошвы пластов,
#Jmh/Jsn/Ol/,алевролит/известняк/доломит/,среднее значение ГИС по искомым интервалам,
#среднеквадратическое отклонение,ошибка среднего,медиана]
def write_layer(lay,array_map,spatial_ref):
        fc_path,fc_name = os.path.split(lay)
        gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",spatial_ref)
        gp.AddField(lay,"id_скважины","STRING",15,15)#атрибут id скважин
	gp.AddField(lay,"название_скважины","STRING",15,15)#атрибут названия скважин
	gp.AddField(lay,"x","FLOAT",15,15)#добавить атрибут x
	gp.AddField(lay,"y","FLOAT",15,15)#добавить атрибут y
	gp.AddField(lay,"z","FLOAT",15,15)#добавить атрибут z
	gp.AddField(lay,"глубина_скважины","FLOAT",15,15)#добавить атрибут глубины скважины
	gp.AddField(lay,"название_объекта","STRING","","",100)#атрибут - название объекта
	gp.AddField(lay,"название_участка","STRING","","",50)#атрибут - название участка
        gp.AddField(lay,"тип_точки_наблюдения","STRING","","",100)#атрибут - тип точки наблюдения
        gp.AddField(lay,"состояние_документирования","STRING","","",100)#атрибут - состояние документирования
        gp.AddField(lay,"тип_документирования","STRING","","",100)#атрибут - тип документирования
	gp.AddField(lay,"глубина_измерений_ГИС","STRING","","",1000)#атрибут значения глубин измерения ГИС
	gp.AddField(lay,"значения_ГИС","STRING","","",1000)#атрибут для значений ГИС
        gp.AddField(lay,"интервалы_искомого_подразделения","STRING","","",1000)#атрибут абсолютных отметок искомых стратиграфических подразделений
        gp.AddField(lay,"глубина_кровли_пластов","STRING","","",1000)#атрибут для глубин кровли пластов
        gp.AddField(lay,"глубина_подошвы_пластов","STRING","","",1000)#атрибут для глубин подошвы пластов
        gp.AddField(lay,"стратиграфия","STRING","","",1000)#атрибут для стратиграфии
        gp.AddField(lay,"литология","STRING","","",1000)#атрибут для литологии
        gp.AddField(lay,"среднее_значение_ГИС","FLOAT",15,15)#атрибут среднее значение ГИС
        gp.AddField(lay,"количество_измерений","FLOAT",15,15)#атрибут для количества измерений
        gp.AddField(lay,"среднеквадратическое_отклонение","FLOAT",15,15)#атрибут для среднеквадратического_отклонения
        gp.AddField(lay,"ошибка_среднего","FLOAT",15,15)#атрибут для ошибки среднего
        gp.AddField(lay,"медиана","FLOAT",15,15)#атрибут для медианы
        rows = gp.InsertCursor(lay)
        pnt = gp.CreateObject("Point")
        k=0
        while(k<len(array_map)):#создание точечного слоя
                id_well = float(array_map[k][0])#id скважин
    		name_well = str(array_map[k][1])#названия скважин
    		z = round(float(array_map[k][4]),3)#z
                depth_well = float(array_map[k][5])#глубина скважины
    		name_obj = str(array_map[k][6])#название объекта
    		name_uchastok = str(array_map[k][7])#название участка
                name_type_tn = str(array_map[k][8])#атрибут - тип точки наблюдения
    		name_status_doc = str(array_map[k][9])#атрибут - состояние документирования
    		name_type_doc = str(array_map[k][10])#атрибут - тип документирования
    		depth_logging = str(array_map[k][11])#значения глубин измерения ГИС
    		value_logging = str(array_map[k][12])#значения ГИС
    		abs_strat = array_map[k][13]#атрибут abs подразделений
    		abs_litostrat_roof = str(array_map[k][14])#глубины кровли пластов
                abs_litostrat_sole = str(array_map[k][15])#глубины подошвы пластов
    		stratigraphy = str(array_map[k][16])#стратиграфия
    		lithology = str(array_map[k][17])#литология
    		average_gis = array_map[k][18]#среднее значение ГИС
    		quant_gis = array_map[k][19]#количество измерений
    		msd_atr = array_map[k][20]#среднеквадратическое отклонение
    		standard_error_atr = array_map[k][21]#ошибка среднего
    		mediana = array_map[k][22]#медиана
    		feat = rows.NewRow()#Новая строка
    		pnt.id = k #ID
    		x = round(float(array_map[k][2]),3)
    		y = round(float(array_map[k][3]),3)
    		pnt.x = x
    		pnt.y = y
    		feat.shape = pnt
    		feat.SetValue("id_скважины",id_well)#добавить значения к атрибутивным полям
    		feat.SetValue("название_скважины",name_well)
		feat.SetValue("x",x)
		feat.SetValue("y",y)
    		feat.SetValue("z",z)
    		feat.SetValue("глубина_скважины",depth_well)
    		feat.SetValue("название_объекта",name_obj)
    		feat.SetValue("название_участка",name_uchastok)
                feat.SetValue("тип_точки_наблюдения",name_type_tn)
                feat.SetValue("состояние_документирования",name_status_doc)
                feat.SetValue("тип_документирования",name_type_doc)
    		feat.SetValue("глубина_измерений_ГИС",depth_logging)
    		feat.SetValue("значения_ГИС",value_logging)
    		feat.SetValue("интервалы_искомого_подразделения",abs_strat)
    		feat.SetValue("глубина_кровли_пластов",abs_litostrat_roof)
    		feat.SetValue("глубина_подошвы_пластов",abs_litostrat_sole)
    		feat.SetValue("стратиграфия",stratigraphy)
    		feat.SetValue("литология",lithology)
    		feat.SetValue("среднее_значение_ГИС",average_gis)
    		feat.SetValue("количество_измерений",quant_gis)
    		feat.SetValue("среднеквадратическое_отклонение",msd_atr)
    		feat.SetValue("ошибка_среднего",standard_error_atr)
    		feat.SetValue("медиана",mediana)
    		rows.InsertRow(feat)#добавить строку к таблице слоя
                k=k+1
if __name__== '__main__':#если модуль запущен, а не импортирован
        #-------------------------------------входные параметры----------------------------------------
        InpPFC = ARCPY.GetParameterAsText(0)#входной точечный слой скважин
        OutputFC = arcpy.GetParameterAsText(1)#выходной точечный слой скважин 
        Full_report = ARCPY.GetParameterAsText(2)#параметр определяет выводить ли полный отчет
        #----------------------------------------------------------------------------------------------
        arr_map_points = read_map_layer(InpPFC)#извлечь поля точечного слоя
        
        if(len(arr_map_points)==0):#если массив точек пуст - завершить работу программы
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
                ARCPY.AddMessage("-------------------------------------------------------------------------------")
        #получить среднее измеренное значение ГИС в искомых стратиграфических уровнях для всех скважин
        arr_select = select_values(arr_map_points)     
        if(len(arr_select)==0):#если массив пуст
                raise SystemExit(0)
        ARCPY.AddMessage("Количество скважин с информацией ГИС на заданную глубину для вывода в картографический слой: "+str(len(arr_select)))
        Spat_ref = gp.describe(InpPFC).spatialreference
        write_layer(OutputFC,arr_select,Spat_ref)#записать итоговый массив к картографический слой
        


