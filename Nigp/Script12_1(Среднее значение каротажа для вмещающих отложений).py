import sys, arcgisscripting, os, math
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
#функция на входе берет массив, считанный из картографического слоя
#и возвращает массив в формате
#[id скважины, линия/номер, x,y,z,глубина скважины,название объекта,
#название участка, мощность перекрывающих отложений, cумма каротажных измерений во вмещающих отложениях,
#сумма измерений ГИС во вмещающих отложениях,среднее измеренное значение ГИС во вмещающих отложениях]
def select_value_depth(arr_map):
        k=0
        arr_output=[]
        while(k<len(arr_map)):#перебор элементов массива из картографического слоя
                vhr = round(arr_map[k][10],2)#мощность перекрывающих отложений
                if(vhr==-1.0):pass#Если нет данных по мощности перекрывающих отложений
                else:#Если есть данные о перекрывающих отложениях            
                        id_well = str(arr_map[k][0])#id скважины                  
                        str_depth_logging = str(arr_map[k][8])#строка содержит глубины измерений ГИС для текущей скважины
                        str_value_logging = str(arr_map[k][9])#строка содержит значения измерений ГИС для текущей скважины
                        arr_depth_logging = str_depth_logging.split("/")#получить массив глубин измерений ГИС для текущей скважины
                        arr_value_logging = str_value_logging.split("/")#получить массив значений ГИС для текущей скважины
                        arr_dimension_contain=[]#массив для записи глубин измерений по вмещающим отложениям
                        arr_value_contain=[]#массив для записи значений измерений по вмещающим отложениям
                        if(len(arr_depth_logging)<>len(arr_value_logging)):#если массивы глубин ГИС и измеренных значений ГИС равны
                                ARCPY.AddMessage("Массивы глубин измерений ГИС и самих измеренных \
                                значений не равны, код скважины: "+str(id_well))
                        else:#если массивы глубин ГИС и измеренных значений ГИС равны
                                bool_recover=False#логическая переменная определяет дошел ли цикл до вмещающих отложений
                                depth_recover=0#суммарное количество измерений во вмещающих отложениях для текущей скважины
                                value_recover=0.0#суммарное значение каротажных измерений для вмещающих отложений
                                dep=0.0
                                if(vhr>0.0):#если мощность перекрывающих отложений больше нуля
                                        m=0
                                        #перебор глубин измерений ГИС для текущей скважины
                                        while(m<len(arr_depth_logging)-1):#не обрабатывать последний элемент массива, поскольку 1/2/3 даст ['1','2','3','']
                                                #преобразовать текущее значение глубины в вещественный тип
                                                depth_float = round(float(arr_depth_logging[m]),2)
                                                #преобразовать текущее значение измерений в вещественный тип
                                                val_float = round(float(arr_value_logging[m]),2)
                                                #если цикл дошел до вмещающих отложений
                                                if(bool_recover):
                                                        #если значение - корректно
                                                        if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                #добавить в массив глубин измерений во вмещающих - текущую глубину измерений
                                                                arr_dimension_contain.append(depth_float)
                                                                #добавить в массив значений измерений во вмещающих - текущее значение измерений
                                                                arr_value_contain.append(val_float)
                                                                #Прибавить к сумме каротажных измерений текущее значение
                                                                value_recover=value_recover+val_float
                                                                depth_recover+=1#увеличить количество измерений во вмещающих отложениях
                                                #если цикл еще не дошел до вмещающих отложений 
                                                else:
                                                        #если текущая глубина измерения ГИС равна или первый раз превысила
                                                        #мощность перекрывающих отложений, значит цикл дошел до вмещающих отложений
                                                        if(depth_float==vhr)or(depth_float>vhr):
                                                                bool_recover=True
                                                                dep=depth_float
                                                                #если значение - корректно
                                                                if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                        #добавить в массив глубин измерений во вмещающих - текущую глубину измерений
                                                                        arr_dimension_contain.append(depth_float)
                                                                        #добавить в массив значений измерений во вмещающих - текущее значение измерений
                                                                        arr_value_contain.append(val_float)
                                                                        #добавить текущее значение измерений ГИС к сумме измерений
                                                                        #в перекрывающих отложениях
                                                                        value_recover=value_recover+val_float
                                                                        depth_recover+=1#увеличить количество измерений во вмещающих отложениях
                                                m+=1
                                        #если количество измерений во вмещающих отложениях больше нуля
                                        #и если сумма каротажных измерений больше нуля
                                        # - вычислить среднее измеренное значение ГИС во вмещающих отложениях
                                        if(depth_recover>0)and(value_recover>0.0):
                                                #вычислить среднее значение
                                                sum_value_recover=round(value_recover/depth_recover,2)
                                                msd=0.0 #среднеквадратическое отклонение
                                                standard_error = 0.0 #стандартная ошибка среднего
                                                if (len(arr_dimension_contain)==len(arr_value_contain)):
                                                        i=0
                                                        while(i<len(arr_value_contain)):
                                                                msd=msd+(arr_value_contain[i]-sum_value_recover)**2
                                                                i+=1                                
                                                        msd=math.sqrt(msd/depth_recover)
                                                        standard_error = msd/math.sqrt(depth_recover)#вычислить ошибку среднего
                                                else:
                                                        ARCPY.AddMessage("количество элементов в массиве \
                                                        глубин измерений и массиве со значениями измерений не равно")
                                                
                                                curr_array=[id_well,arr_map[k][1],arr_map[k][2],arr_map[k][3],\
                                                round(arr_map[k][4],2),arr_map[k][5],arr_map[k][6],arr_map[k][7],\
                                                vhr,value_recover,depth_recover,sum_value_recover,msd,standard_error]
                                                arr_output.append(curr_array)
                                elif(vhr==0.0):#если мощность перекрывающих отложений равна 0
                                        ARCPY.AddMessage("Открытая площадь")
                
                k=k+1
        #вернуть массив в формате [id скважины, линия/номер, x,y,z,глубина скважины,название объекта,
        #название участка, мощность перекрывающих отложений, сумма каротажных измерений во вмещающих отложениях,
        #сумма измерений ГИС во вмещающих отложениях, среднее измеренное значение ГИС во вмещающих отложениях,
        #среднеквадратическое отклонение]
        return arr_output 
#Функция на входе берет название выходного точечного слоя, итоговый массив для записи, и пространственную привязку
#и записывет картографический слой 
def write_layer(FC,array_map,spatial_ref):
        fc_path,fc_name = os.path.split(FC)
        gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",spatial_ref)
        gp.AddField(OutputFC,"id_well","STRING",15,15)#атрибут id скважин
	gp.AddField(OutputFC,"name_well","STRING",15,15)#атрибут названия скважин
	gp.AddField(OutputFC,"x","FLOAT",15,15)#добавить атрибут x
	gp.AddField(OutputFC,"y","FLOAT",15,15)#добавить атрибут y
	gp.AddField(OutputFC,"z","FLOAT",15,15)#добавить атрибут z
	gp.AddField(OutputFC,"depth_well","FLOAT",15,15)#добавить атрибут глубины скважины
	gp.AddField(OutputFC,"name_object","STRING","","",100)#атрибут - название объекта
	gp.AddField(OutputFC,"name_area","STRING","","",50)#атрибут - название участка
        gp.AddField(OutputFC,"recover","FLOAT",15,15)#добавить атрибут мощности перекрывающих отложений
        gp.AddField(OutputFC,"value_rec","FLOAT",15,15)#добавить атрибут суммы значений измерений ГИС во вмещающих
        gp.AddField(OutputFC,"depth_rec","FLOAT",15,15)#добавить атрибут каротажных измерений во вмещающих отложений
        gp.AddField(OutputFC,"mean_value_gis","FLOAT",15,15)#добавить атрибут для среднего измеренного значения ГИС во вмещающих отложениях
        gp.AddField(OutputFC,"msd","FLOAT",15,15)#добавить атрибут для среднеквадратического отклонения
        gp.AddField(OutputFC,"error_mean","FLOAT",15,15)#добавить атрибут для ошибки общего среднего
        rows = gp.InsertCursor(FC)
        pnt = gp.CreateObject("Point")
        k=0
        while(k<len(array_map)):#создание точечного слоя
                id_well = float(array_map[k][0])#id скважин
    		name_well = str(array_map[k][1])#названия скважин
    		z = round(float(array_map[k][4]),3)#z
                depth_well = float(array_map[k][5])#глубина скважины
    		name_obj = str(array_map[k][6])#название объекта
    		name_uchastok = str(array_map[k][7])#название участка
    		recover = float(array_map[k][8])#атрибут мощности перекрывающих отложений
    		value_rec = float(array_map[k][9])#сумма значений измерений ГИС во вмещающих
    		depth_rec = float(array_map[k][10])#сумма каротажных измерений во вмещающих отложений
    		mean_value_gis = array_map[k][11]#среднее измеренное значение ГИС во вмещающих отложениях
    		msd_value = array_map[k][12]#среднеквадратическое отклонение
    		error_mean = array_map[k][13]#ошибка общего среднего
                feat = rows.NewRow()#Новая строка
                pnt.id=k#ID
                x = round(float(array_map[k][2]),3)
    		y = round(float(array_map[k][3]),3)
    		pnt.x = x
    		pnt.y = y
                feat.shape = pnt
    		feat.SetValue("id_well",id_well)#добавить значения к атрибутивным полям
    		feat.SetValue("name_well",name_well)
		feat.SetValue("x",x)
		feat.SetValue("y",y)
    		feat.SetValue("z",z)
    		feat.SetValue("depth_well",depth_well)
    		feat.SetValue("name_object",name_obj)
    		feat.SetValue("name_area",name_uchastok)
    		feat.SetValue("recover",recover)
    		feat.SetValue("value_rec",value_rec)
    		feat.SetValue("depth_rec",depth_rec)
    		feat.SetValue("mean_value_gis",mean_value_gis)
    		feat.SetValue("msd",msd_value)
    		feat.SetValue("error_mean",error_mean)
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
        #получить среднее измеренное значение ГИС в кимберлитовмещающих отложениях для всех скважин
        arr_select = select_value_depth(arr_map_points)
        if(len(arr_select)==0):#если массив пуст
                raise SystemExit(0)
        ARCPY.AddMessage("Количество скважин с информацией ГИС на заданную глубину для вывода в картографический слой: "+str(len(arr_select)))
        if(Full_report == "true"):
                ARCPY.AddMessage("В формате:")
                ARCPY.AddMessage("id скважины: "+str(arr_select[0][0]))
                ARCPY.AddMessage("название скважины: "+str(arr_select[0][1]))
                ARCPY.AddMessage("x: "+str(arr_select[0][2]))
                ARCPY.AddMessage("y: "+str(arr_select[0][3]))
                ARCPY.AddMessage("z: "+str(arr_select[0][4]))
                ARCPY.AddMessage("глубина скважины: "+str(arr_select[0][5]))
                ARCPY.AddMessage("название объекта: "+str(arr_select[0][6]))
                ARCPY.AddMessage("название участка: "+str(arr_select[0][7]))
                ARCPY.AddMessage("мощность перекрывающих отложений: "+str(arr_select[0][8]))
                ARCPY.AddMessage("сумма измерений ГИС во вмещающих отложениях: "+str(arr_select[0][9]))
                ARCPY.AddMessage("количество измерений во вмещающих отложениях: "+str(arr_select[0][10]))
                ARCPY.AddMessage("среднее измеренное значение ГИС во вмещающих отложениях: "+str(arr_select[0][11]))
        Spat_ref = gp.describe(InpPFC).spatialreference
        write_layer(OutputFC,arr_select,Spat_ref)#записать итоговый массив к картографический слой
        


