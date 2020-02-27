import arcpy, sys, arcgisscripting, os
#Функция в качестве входного параметра берет точечный слой с данными минералогии(проб)
#и возвращает два массива:
#массив с названиями полей листа минералогии
#и массив точек(проб), считанных из точечного слоя минералогии 
#[id_well,name_cell,x,y,z,глубина скважины,name_object,name_uchastok,number_test,id_test,
#code_type_test,abs_from,abs_to,mineralogy,отметка_кровли_пластов,отметка_подошвы_пластов
#стратиграфия,литология,тип_точки_наблюдения,состояние_документирования,состояние_опробования,тип_документирования]
def read_points(layer):
        def check_unicode(text):#функция преобразует кодировку юникод в utf-8 
                if(type(text)==unicode):text = text.encode('utf-8')
                return text
        arr_wells=[]#массив для записывания скважин
        arr_fields = []#массив с названиями атрибутов входного точечного слоя
        fields = arcpy.ListFields(layer)#получить список названий атрибутов входного точечного слоя
        #перебор атрибутов
        for field in fields:arr_fields.append(field.name.encode('utf-8'))#заполнить массив с названиям атрибутов
        rows = gp.SearchCursor(layer)
        layer = gp.describe(layer)
        for row in rows:#перебор точек картографического слоя
                feat = row.GetValue(layer.ShapeFieldName)
                point = feat.getpart(0)
                id_well = check_unicode(row.GetValue("id_скважины"))#id скважины
                name_cell = check_unicode(row.GetValue("название_скважины"))#название скважины
                x = point.x
                y = point.y
                z = row.GetValue("z")
                z = round(z,2)#округлить до двух знаков после запятой
                depth_well = check_unicode(row.GetValue("глубина_скважины"))#глубина скважины
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
                abs_litostrat_roof = check_unicode(row.GetValue("отметка_кровли_пластов"))
                abs_litostrat_sole = check_unicode(row.GetValue("отметка_подошвы_пластов"))
                stratigraphy = check_unicode(row.GetValue("стратиграфия"))
                lithology = check_unicode(row.GetValue("литология"))
                type_tn = check_unicode(row.GetValue("тип_точки_наблюдения"))
                status_doc = check_unicode(row.GetValue("состояние_документирования"))
                status_test = check_unicode(row.GetValue("состояние_опробования"))
                type_doc = check_unicode(row.GetValue("тип_документирования"))
                arr_wells.append([id_well,name_cell,x,y,z,depth_well,name_object,name_uchastok,number_test,id_test,code_type_test, \
                abs_from,abs_to,mineralogy,abs_litostrat_roof,abs_litostrat_sole,stratigraphy,lithology, \
                type_tn, status_doc, status_test,type_doc])
        #вернуть массив с названиями полей листа минералогии
        #и массив точек(проб), считанных из точечного слоя минералогии
        #в формате [id_well,name_cell,x,y,z,глубина скважины,name_object,name_uchastok,number_test,id_test,
        #code_type_test,abs_from,abs_to,mineralogy,отметка_кровли_пластов,отметка_подошвы_пластов,
        #стратиграфия,литология,тип_точки_наблюдения,состояние_документирования,состояние_опробования,тип_документирования]
        return arr_fields, arr_wells 
#Функция в качестве входных параметров берет массив с пробами из входного слоя и
#и название стратиграфического уровня, в котором следует вести поиск
def search_mineral_in_stratigraphy(arr_p,strat):
        find_array = []
        k=0
        while(k<len(arr_p)):#перебор проб
                mark_from = arr_p[k][11]#интервал глубин пробы от, для текущей пробы
                mark_to = arr_p[k][12]#интервал глубин пробы до, для текущей пробы
                arr_mark_litostrat_roof = arr_p[k][14].split("/")#получить массив отметок кровли пластов для текущей пробы
                arr_mark_litostrat_sole = arr_p[k][15].split("/")#получить массив отметок подошвы пластов для текущей пробы
                arr_strat = arr_p[k][16].split("/")#получить массив стратиграфии пластов для текущей пробы
                if(len(arr_strat)==1):#если нет данных по стратиграфии для текущей пробе
                        #arcpy.AddMessage("нет данных по стратиграфии для текущей пробы")
                        #arcpy.AddMessage("-------------------- "+str(k))
                        k+=1
                        continue#поскольку нет данных по стратиграфии по текущей пробе, перейти к следующей пробе
                arr_lit = arr_p[k][17].split("/")#получить массив литологии пластов для текущей пробы
                m=0
                while(m<len(arr_strat)-1):#перебор стратиграфических пластов для текущей пробы, не обрабатывать последний элемент массива, поскольку 1/2/3 даст ['1','2','3','']
                        #если название стратиграфии пласта совпало с типом стртиграфии, заданным как входной параметр
                        if(arr_strat[m].count(strat)):
                                #arcpy.AddMessage(str(arr_p[k][0])+", "+str(arr_strat[m])+", "+str(arr_lit[m])+", интервалы пласта:"+ \
                                #str(arr_mark_litostrat_roof[m])+"-"+str(arr_mark_litostrat_sole[m])+", интервалы пробы:"+str(mark_from)+"-"+str(mark_to))                      
                                if(mark_from==float(arr_mark_litostrat_roof[m]))and(mark_to==float(arr_mark_litostrat_sole[m])):
                                        #arcpy.AddMessage("интервал пробы и пласта совпадают (1 случай)")
                                        find_array.append([arr_p[k][0],arr_p[k][1],arr_p[k][2],arr_p[k][3],arr_p[k][4], \
                                        arr_p[k][5],arr_p[k][6],arr_p[k][7],arr_p[k][8],arr_p[k][9],arr_p[k][10], \
                                        arr_p[k][11],arr_p[k][12],arr_p[k][13],arr_strat[m],arr_lit[m],arr_p[k][18], \
                                        arr_p[k][19],arr_p[k][20],arr_p[k][21]])
                                        break#выйти из текущего цикла
                                elif(mark_from>=float(arr_mark_litostrat_roof[m]))and(mark_to<=float(arr_mark_litostrat_sole[m])):
                                        #arcpy.AddMessage("совпадение - интервалы пробы находятся внутри пласта (2 случай)")
                                        find_array.append([arr_p[k][0],arr_p[k][1],arr_p[k][2],arr_p[k][3],arr_p[k][4], \
                                        arr_p[k][5],arr_p[k][6],arr_p[k][7],arr_p[k][8],arr_p[k][9],arr_p[k][10], \
                                        arr_p[k][11],arr_p[k][12],arr_p[k][13],arr_strat[m],arr_lit[m],arr_p[k][18], \
                                        arr_p[k][19],arr_p[k][20],arr_p[k][21]])
                                        break#выйти из текущего цикла
                                elif(mark_from>=float(arr_mark_litostrat_roof[m]))and(mark_from<float(arr_mark_litostrat_sole[m])):
                                        #arcpy.AddMessage("совпадение по верхней границе пробы (3 случай)")
                                        find_array.append([arr_p[k][0],arr_p[k][1],arr_p[k][2],arr_p[k][3],arr_p[k][4], \
                                        arr_p[k][5],arr_p[k][6],arr_p[k][7],arr_p[k][8],arr_p[k][9],arr_p[k][10], \
                                        arr_p[k][11],arr_p[k][12],arr_p[k][13],arr_strat[m],arr_lit[m],arr_p[k][18], \
                                        arr_p[k][19],arr_p[k][20],arr_p[k][21]])
                                        break#выйти из текущего цикла
                                elif(mark_to>float(arr_mark_litostrat_roof[m]))and(mark_to<=float(arr_mark_litostrat_sole[m])):
                                        #arcpy.AddMessage("совпадение по нижней границе пробы (4 случай)")
                                        find_array.append([arr_p[k][0],arr_p[k][1],arr_p[k][2],arr_p[k][3],arr_p[k][4], \
                                        arr_p[k][5],arr_p[k][6],arr_p[k][7],arr_p[k][8],arr_p[k][9],arr_p[k][10], \
                                        arr_p[k][11],arr_p[k][12],arr_p[k][13],arr_strat[m],arr_lit[m],arr_p[k][18], \
                                        arr_p[k][19],arr_p[k][20],arr_p[k][21]])
                                        break#выйти из текущего цикла
                                elif(mark_from<=float(arr_mark_litostrat_roof[m]))and(mark_to>=float(arr_mark_litostrat_sole[m])):
                                        #arcpy.AddMessage("совпадение - интервал пробы больше мощности пласта (5 случай)")
                                        find_array.append([arr_p[k][0],arr_p[k][1],arr_p[k][2],arr_p[k][3],arr_p[k][4], \
                                        arr_p[k][5],arr_p[k][6],arr_p[k][7],arr_p[k][8],arr_p[k][9],arr_p[k][10], \
                                        arr_p[k][11],arr_p[k][12],arr_p[k][13],arr_strat[m],arr_lit[m],arr_p[k][18], \
                                        arr_p[k][19],arr_p[k][20],arr_p[k][21]])
                                        break#выйти из текущего цикла
                                else:
                                        pass
                                        #arcpy.AddMessage("нет совпадений")
                        else:
                                pass
                                #arcpy.AddMessage("другая свита")
                        m+=1
                #arcpy.AddMessage("-------------------- "+str(k))
                k+=1
        #вернуть массив в формате
        #[id_well,name_cell,x,y,z,глубина скважины,name_object,name_uchastok,number_test,id_test,
        #type_test,abs_from,abs_to,mineralogy, стратиграфия пласта с которым полностью или частично совпала проба,
        #литология пласта с которым полностью или частично совпала проба,тип_точки_наблюдения,
        #состояние_документирования,состояние_опробования,тип_документирования]
        return find_array
#функция в качестве входных параметров берет директорию выходного картографического слоя, систему координат, и итоговый массив
#в формате
#[id_well,name_cell,x,y,z,глубина скважины,name_object,name_uchastok,number_test,id_test,
#type_test,abs_from,abs_to,mineralogy, стратиграфия пласта с которым полностью или частично совпала проба,
#литология пласта с которым полностью или частично совпала проба,тип_точки_наблюдения,
#состояние_документирования,состояние_опробования,тип_документирования]
#и создает картографический слой с соответствующими атрибутами
def write_to_layer(lay, C_System, array_map, atr_min):
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
	gp.AddField(OutputFC,"глубина_скважины","STRING","","",50)#атрибут - глубина скважины
	gp.AddField(OutputFC,"название_объекта","STRING","","",100)#атрибут - название объекта
	gp.AddField(OutputFC,"название_участка","STRING","","",50)#атрибут - название участка
	gp.AddField(OutputFC,"номер_пробы","STRING","","",50)#атрибут - номер пробы
	gp.AddField(OutputFC,"id_пробы","STRING","","",50)#атрибут - id пробы
	gp.AddField(OutputFC,"тип_пробы","STRING","","",50)#атрибут - код типа пробы
	gp.AddField(OutputFC,"интервал_пробы_от","FLOAT","","",50)#атрибут - интервал глубины от
	gp.AddField(OutputFC,"интервал_пробы_до","FLOAT","","",50)#атрибут - интервал глубины до
	gp.AddField(OutputFC,"находки_МСА","STRING","","",1000)#атрибут - индекс МСА
	gp.AddField(OutputFC,"стратиграфия","STRING","","",1000)#атрибут для стратиграфии
	gp.AddField(OutputFC,"литология","STRING","","",1000)#атрибут для литологии
	gp.AddField(OutputFC,"тип_точки_наблюдения","STRING","","",1000)#атрибут для типа точки наблюдения
        gp.AddField(OutputFC,"состояние_документирования","STRING","","",1000)#атрибут для состояния документирования
        gp.AddField(OutputFC,"состояние_опробования","STRING","","",1000)#атрибут для состояния опробования
        gp.AddField(OutputFC,"тип_документирования","STRING","","",1000)#атрибут для типа документирования  
        #создать атрибутивные поля для находок МСА
        u=24
        while(u<len(atr_min)):
                gp.AddField(OutputFC,atr_min[u],"LONG","","",50)
                u+=1
	rows = gp.InsertCursor(OutputFC)
	pnt = gp.CreateObject("Point")
	k=0
	while(k<len(array_map)):
                id_well = float(array_map[k][0])
    		name_well = str(array_map[k][1])
    		z = float(array_map[k][4])
    		depth_well = str(array_map[k][5])
    		name_object = str(array_map[k][6])
    		name_area = str(array_map[k][7])
    		number_test = str(array_map[k][8])
    		id_test = str(array_map[k][9])
    		code_type_test = str(array_map[k][10])
    		abs_from = float(array_map[k][11])
    		abs_to = float(array_map[k][12])
    		mineralogy = str(array_map[k][13])
                arr_min = mineralogy.split("/")
    		stratigraphy = str(array_map[k][14])
    		lithology = str(array_map[k][15])
    		type_tn = str(array_map[k][16])
    		stat_doc = str(array_map[k][17])
    		stat_test = str(array_map[k][18])
    		type_doc = str(array_map[k][19])
    		feat = rows.NewRow()#Новая строка
    		pnt.id = k #ID
    		pnt.x = float(array_map[k][2])
    		pnt.y = float(array_map[k][3])
    		feat.shape = pnt
    		feat.SetValue("id_скважины",id_well)#добавить значения к атрибутивным полям
    		feat.SetValue("название_скважины",name_well)
		feat.SetValue("x",float(array_map[k][2]))
		feat.SetValue("y",float(array_map[k][3]))
    		feat.SetValue("z",z)
    		feat.SetValue("название_объекта",name_object)
    		feat.SetValue("название_участка",name_area)
    		feat.SetValue("глубина_скважины",depth_well)
                feat.SetValue("номер_пробы",int(float(number_test)))
                feat.SetValue("id_пробы",int(float(id_test)))
                feat.SetValue("тип_пробы",code_type_test)
                feat.SetValue("интервал_пробы_от",abs_from)
                feat.SetValue("интервал_пробы_до",abs_to)
                feat.SetValue("находки_МСА",mineralogy)
                feat.SetValue("стратиграфия",stratigraphy)
                feat.SetValue("литология",lithology)
                feat.SetValue("тип_точки_наблюдения",type_tn)
                feat.SetValue("состояние_документирования",stat_doc)
                feat.SetValue("состояние_опробования",stat_test)
                feat.SetValue("тип_документирования",type_doc)
                #заполнить нулевыми значениями атрибутивные поля для находок МСА
                u=24
                while(u<len(atr_min)):
                        feat.SetValue(atr_min[u],0) 
                        u+=1
                #перебор элементов шапки листа с минералогией
                i=24 #начать перебор с элемента pir05_0
                while(i<len(atr_min)-1):
                        #перебор находок МСА для текущей скважины, не обрабатывать последний элемент массива
                        #он является пустым - при переводе строки '1/2/3/' получается массив ['1','2','3','']          
                        m=0
                        while(m<len(arr_min)-1):
                                curr_min = arr_min[m].split("=")#разделить тип МСА и количество кристаллов
                                if (atr_min[i]==curr_min[0]):#если атрибут из шапки листа минералогии совпал с находкой для текущей скважины
                                        val = int(float(curr_min[1]))
                                        feat.SetValue(curr_min[0],val)#добавить к текущему типу МСА количество найденных кристаллов 
                                        break      
                                m+=1
                        i+=1
                rows.InsertRow(feat)#добавить строку к таблице слоя
                k+=1
if __name__ == '__main__': #если модуль запущен а не импортирован
        gp = arcgisscripting.create(9.3)    
        #-------------------------------входные данные------------------------------------
        InpPFC = arcpy.GetParameterAsText(0)#входной точечный слой скважин с данными минералогии
        Stratigraphy =  arcpy.GetParameterAsText(1).encode('utf-8')#стратиграфический уровень осадочного чехла
        OutputFC = arcpy.GetParameterAsText(2)#выходной точечный слой скважин c находками МСА
        Full_report = arcpy.GetParameterAsText(3)#параметр определяет выводить ли полный отчет
        if(Stratigraphy=="Дяхтарская толща (J1dh)"):
                Stratigraphy = "dh"
        elif(Stratigraphy=="Кора выветривания (T2-3)"):
                Stratigraphy = "T2-3"
        elif(Stratigraphy=="Укугутская свита (J1uk)"):
                Stratigraphy = "uk"
        elif(Stratigraphy=="Тюнгская свита (J1tn)"):
                Stratigraphy = "tn"
        #---------------------------------------------------------------------------------
        fields_arr, arr_points = read_points(InpPFC)#получить массив атрибутов входного слоя и массив с данными входного слоя
        arcpy.AddMessage("Количество объектов точечного слоя c данными минералогии: "+str(len(arr_points)))
        if (len(arr_points)>0):
                if(Full_report=="true"):#выводить полный отчет
                        arcpy.AddMessage("В формате:")
                        i=0
                        s = ""
                        while(i<len(arr_points[0])):
                                s = s+" "+str(arr_points[0][i])
                                i+=1
                        arcpy.AddMessage(s)
                        arcpy.AddMessage("--------------------------------------------")
        else:raise SystemExit(1)#завершить работу программы
        arcpy.AddMessage("Количество атрибутов входного точечного слоя: "+str(len(fields_arr)))
        if (len(fields_arr)>0):
                if(Full_report=="true"):#выводить полный отчет
                        k=0
                        while(k<len(fields_arr)):
                                arcpy.AddMessage(str(k)+": "+str(fields_arr[k]))
                                k+=1
                        arcpy.AddMessage("--------------------------------------------")
        #найти пробы, которые были взяты в стратиграфическом уровне, стратиграфия которого
        #указывается во входном параметре(Stratigraphy); или проба пересекается с этим уровнем
        find_msa_arr = search_mineral_in_stratigraphy(arr_points,Stratigraphy)                       
        arcpy.AddMessage("Количество проб в заданном стратиграфическом уровне: "+str(len(find_msa_arr)))
        if(len(find_msa_arr)>5):
                if(Full_report=="true"):#выводить полный отчет
                        k=0
                        while(k<5):
                                i=0
                                s = ""
                                while(i<len(find_msa_arr[k])):
                                        s = s+" "+str(find_msa_arr[k][i])
                                        i+=1
                                arcpy.AddMessage(s)
                                arcpy.AddMessage("--------------------------------------------")
                                k+=1
        Spat_ref = gp.describe(InpPFC).spatialreference
        #вывести полученный массив в картографический слой
        write_to_layer(OutputFC, Spat_ref, find_msa_arr,fields_arr)
        














                
