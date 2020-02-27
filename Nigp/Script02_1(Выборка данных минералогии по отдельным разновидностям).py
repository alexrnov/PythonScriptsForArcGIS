import arcpy, sys, arcgisscripting, os
#Функция в качестве входного параметра берет точечный слой с данными минералогии
#и возвращает массив в формате
#[id_well,name_cell,x,y,z,name_object,name_uchastok,number_test,id_test,code_type_test,abs_from,abs_to,mineralogy]
def read_points(layer):
        def check_unicode(text):#функция преобразует кодировку юникод в utf-8 
                if(type(text)==unicode):text = text.encode('utf-8')
                return text
        rows = gp.SearchCursor(layer)
        row = rows.Next()
        layer = gp.describe(layer)
        arr_wells=[]#массив для записывания скважин
        while row:#перебор точек картографического слоя
                feat = row.GetValue(layer.ShapeFieldName)
                point = feat.getpart(0)
                id_well = check_unicode(row.GetValue("id_скважины"))#id скважины
                name_cell = check_unicode(row.GetValue("название_скважины"))#название скважины
                x = point.x
                y = point.y
                z = row.GetValue("z")
                z = round(z,2)#округлить до двух знаков после запятой 
                name_object = check_unicode(row.GetValue("название_объекта"))#название объекта
                name_uchastok = check_unicode(row.GetValue("название_участка"))#название участка
                number_test = check_unicode(row.GetValue("номер_пробы"))#номер пробы
                id_test = check_unicode(row.GetValue("id_пробы"))#id пробы
                code_type_test = check_unicode(row.GetValue("тип_пробы"))#код типа пробы
                abs_from = row.GetValue("интервал_пробы_от")#интервал глубин пробы от
                abs_from = round(abs_from,2)#округлить до двух знаков после запятой 
                abs_to = row.GetValue("интервал_пробы_до")#интервал глубин пробы до
                abs_to = round(abs_to,2)#округлить до двух знаков после запятой 
                mineralogy = row.GetValue("находки_МСА")#данные минералогии
                arr_wells.append([id_well,name_cell,x,y,z,name_object,name_uchastok,number_test,id_test,code_type_test, \
                abs_from,abs_to,mineralogy])
                row = rows.next()
        return arr_wells #вернуть массив точек, считанных из точечного слоя минералогии
#в формате [id_well,name_cell,x,y,z,name_object,name_uchastok,number_test,id_test,code_type_test,abs_from,abs_to,mineralogy]
#Функция в качестве входных параметров берет название минерала, гранулометрию, сохранность
#и возвращает индекс минерала, например pir05_0
def create_name_mineralogy(mineral,gran,safety):
        index_min = ""
        if(mineral=="Пироп"):
                index_min = "pir"+gran+"_"+safety#сформировать индекс минерала название+гранулометрия+сохранность
        if(mineral=="Осколки пиропа"):index_min = "pir"+gran+"_"+"oskl"
        if(mineral=="Корродированный пироп"):index_min = "pir"+gran+"_"+"krd"
        if(mineral=="Пикроильменит"):index_min = "pik"+gran+"_"+safety
        if(mineral=="Осколки пикроильменита"):index_min = "pik"+gran+"_"+"oskl"
        if(mineral=="Корродированный пикроильменит"):index_min = "pik"+gran+"_"+"krd"
        if(mineral=="Агрегатный пикроильменит"):index_min = "pik"+gran+"_"+"agr"
        if(mineral=="Вторичный пикроильменит"):index_min = "pik"+gran+"_"+"vto"
        if(mineral=="Хромшпинелид"):index_min = "hrsh"+gran
        if(mineral=="Оливин"):index_min = "olv"+gran
        if(mineral=="Хромдиопсид"):index_min = "hrds"+gran
        if(mineral=="Циркон"):index_min = "crk"+gran
        if(mineral=="Алмаз"):index_min = "almaz"+gran
        return index_min
#функция в качестве входых параметров берет массив точек наблюдений с информацией по МСА и индекс минерала,
#который указывается в сценарии для поиска
#функция ищет минерал с указанным индексом, и записывает находки в массив формата
#[id_well,name_cell,x,y,z,name_object,name_uchastok,number_test,id_test,code_type_test,
#abs_from,abs_to,mineral_index,amount_mineral]
def read_mineralogy(arr_p,min_index):
        k=0
        find_array = []#массив для записи точек наблюдения с находками МСА, которые указаны в сценарии
        current_find_mineral = []#текущая точка наблюдения с находками МСА, который указан в сценарии
        while(k<len(arr_p)):#перебор всех точек наблюдения с информацией по МСА
                current_arr_string = arr_p[k][12].split("/")#получить найденные МСА для текущей пробы
                m=0
                while(m<len(current_arr_string)-1):#перебор массива индексов МСА для текущей точки наблюдений
                        current_arr_mineral = current_arr_string[m].split("=")#получить индекс минерала и количество найденных кристаллов
                        if (min_index==str(current_arr_mineral[0])):#если текущее название кристалла совпало с указанным в скрипте
                                current_find_mineral = [arr_p[k][0],arr_p[k][1],arr_p[k][2],arr_p[k][3],arr_p[k][4], \
                                arr_p[k][5],arr_p[k][6],arr_p[k][7],arr_p[k][8],arr_p[k][9],arr_p[k][10], \
                                arr_p[k][11],current_arr_mineral[0],current_arr_mineral[1]]
                                find_array.append(current_find_mineral)
                                break#выйти из внутреннего цикла - перейти к следующей ТН
                        m+=1
                k+=1
        #вернуть массив точек наблюдений с находками МСА, которые указаны в сценарии
        #формат массива: [id_well,name_cell,x,y,z,name_object,name_uchastok,number_test,id_test,code_type_test,
        #abs_from,abs_to,mineral_index,amount_mineral]
        return find_array
#функция в качестве входных параметров берет директорию выходного картографического слоя, систему координат, и итоговый массив
#в формате [id_well,name_cell,x,y,z,name_object,name_uchastok,number_test,id_test,code_type_test,
#abs_from,abs_to,mineral_index,amount_mineral]
#и создает картографический слой с соответствующими атрибутами
def write_to_layer(lay, C_System, array_map):
        fc_path, fc_name = os.path.split(lay)#разделить директорию и имя файла
        gp=arcgisscripting.create(9.3)#создать объект геопроцессинга версии 9.3
        if fc_path.count("-") or fc_name.count("-"):
                gp.AddMessage(" '-' Недопустимый символ в названии файла")
                raise SystemExit(1)#завершить работу программы
        try:
                gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",C_System)#создать точечный слой
        except:
                gp.AddMessage("Ошибка создания картографического слоя")
                raise SystemExit(1)#завершить работу программы
        gp.AddField(OutputFC,"id_скважины","STRING",15,15)#атрибут id скважин
	gp.AddField(OutputFC,"название_скважины","STRING",15,15)#атрибут названия скважин
	gp.AddField(OutputFC,"x","FLOAT",15,15)#добавить атрибут x
	gp.AddField(OutputFC,"y","FLOAT",15,15)#добавить атрибут y
	gp.AddField(OutputFC,"z","FLOAT",15,15)#добавить атрибут z
	gp.AddField(OutputFC,"название_объекта","STRING","","",100)#атрибут - название объекта
	gp.AddField(OutputFC,"название_участка","STRING","","",50)#атрибут - название участка
	gp.AddField(OutputFC,"номер_пробы","STRING","","",50)#атрибут - номер пробы
	gp.AddField(OutputFC,"id_пробы","STRING","","",50)#атрибут - id пробы
	gp.AddField(OutputFC,"тип_пробы","STRING","","",50)#атрибут - код типа пробы
	gp.AddField(OutputFC,"интервал_пробы_от","FLOAT","","",50)#атрибут - интервал глубины от
	gp.AddField(OutputFC,"интервал_пробы_до","FLOAT","","",50)#атрибут - интервал глубины до
	gp.AddField(OutputFC,"находки_МСА","STRING","","",100)#атрибут - индекс МСА
	gp.AddField(OutputFC,"количество_кристаллов","FLOAT","","",50)#количество найденных кристаллов указанного МСА
	rows = gp.InsertCursor(OutputFC)
	pnt = gp.CreateObject("Point")
	k=0
	while(k<len(array_map)):
                id_well = float(array_map[k][0])
    		name_well = str(array_map[k][1])
    		z = float(array_map[k][4])
    		name_obj = str(array_map[k][5])
    		name_uchastok = str(array_map[k][6])
    		number_test = str(array_map[k][7])
    		id_test = str(array_map[k][8])
    		code_type_test = str(array_map[k][9])
    		abs_from = float(array_map[k][10])
    		abs_to = float(array_map[k][11])
    		mineralogy = str(array_map[k][12])
    		amount_min = float(array_map[k][13])
    		feat = rows.NewRow()#Новая строка
    		pnt.id = k #ID
    		pnt.x = float(array_map[k][2])
    		pnt.y = float(array_map[k][3])
    		feat.shape = pnt
    		feat.SetValue("id_скважины",id_well)#добавить значения к атрибутивным полям
    		feat.SetValue("название_скважины",name_well)
		feat.SetValue("x",float(array_map[k][3]))
		feat.SetValue("y",float(array_map[k][2]))
    		feat.SetValue("z",z)
    		feat.SetValue("название_объекта",name_obj)
    		feat.SetValue("название_участка",name_uchastok)
                feat.SetValue("номер_пробы",number_test)
                feat.SetValue("id_пробы",id_test)
                feat.SetValue("тип_пробы",code_type_test)
                feat.SetValue("интервал_пробы_от",abs_from)
                feat.SetValue("интервал_пробы_до",abs_to)
                feat.SetValue("находки_МСА",mineralogy)
                feat.SetValue("количество_кристаллов",amount_min)
    		rows.InsertRow(feat)#добавить строку к таблице слоя
                k+=1
if __name__ == '__main__': #если модуль запущен а не импортирован
        gp = arcgisscripting.create(9.3)    
        #-------------------------------входные данные------------------------------------
        InpPFC = arcpy.GetParameterAsText(0)#входной точечный слой скважин с данными минералогии
        name_mineral = arcpy.GetParameterAsText(1).encode('utf-8')#название минера - спутника алмаза
        granulometria = arcpy.GetParameterAsText(2).encode('utf-8')#гранулометрия
        save_mineral = arcpy.GetParameterAsText(3).encode('utf-8')#сохранность МСА   
        OutputFC = arcpy.GetParameterAsText(4)#выходной точечный слой скважин c находками МСА
        #---------------------------------------------------------------------------------
        arr_points = read_points(InpPFC)#извлечь данные минералогии
        arcpy.AddMessage("Количество объектов точечного слоя c данными минералогии: "+str(len(arr_points)))
        if (len(arr_points)>0):
                arcpy.AddMessage("В формате:")
                arcpy.AddMessage(str(arr_points[0][0])+" "+str(arr_points[0][1])+" "+str(arr_points[0][2])+" "+ \
                str(arr_points[0][3])+" "+str(arr_points[0][4])+" "+str(arr_points[0][5])+" "+str(arr_points[0][6])+" "+ \
                str(arr_points[0][7])+" "+str(arr_points[0][8])+" "+str(arr_points[0][9])+" "+str(arr_points[0][10])+" "+ \
                str(arr_points[0][11])+" "+str(arr_points[0][12]))
                arcpy.AddMessage("-----------------------------------------------------------------------------------")
        else:raise SystemExit(1)#завершить работу программы
        index_mineral = create_name_mineralogy(name_mineral,granulometria,save_mineral)#получить индекс минерала                       
        arcpy.AddMessage("Индекс минерала - спутника алмазов: "+index_mineral)
        find_mineral_array = read_mineralogy(arr_points,index_mineral)
        if(len(find_mineral_array)<>0):#если непустой массив точек наблюдей с находками МСА, которые указаны в сценарии
                arcpy.AddMessage("Количество точек наблюдений с находками МСА "+index_mineral+": "+str(len(find_mineral_array)))
                arcpy.AddMessage("В формате:")
                arcpy.AddMessage(str(find_mineral_array[0][0])+" "+str(find_mineral_array[0][1])+ \
                " "+str(find_mineral_array[0][2])+" "+str(find_mineral_array[0][3])+" "+ \
                str(find_mineral_array[0][4])+" "+str(find_mineral_array[0][5])+" "+ \
                str(find_mineral_array[0][6])+" "+str(find_mineral_array[0][7])+" "+ \
                str(find_mineral_array[0][8])+" "+str(find_mineral_array[0][9])+" "+ \
                str(find_mineral_array[0][10])+" "+str(find_mineral_array[0][11])+" "+ \
                str(find_mineral_array[0][12])+" "+str(find_mineral_array[0][13]))
                arcpy.AddMessage("-----------------------------------------------------------------------------------")
        else:
                arcpy.AddMessage("В слое с МСА не найдено минералов "+index_mineral)
                raise SystemExit(1)#завершить работу программы 
        Spat_ref = gp.describe(InpPFC).spatialreference
        #вывести полученный массив в картографический слой
        write_to_layer(OutputFC, Spat_ref, find_mineral_array)















                
