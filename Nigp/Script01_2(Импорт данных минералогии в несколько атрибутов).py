import arcpy as ARCPY
import arcgisscripting
import os
try:
	from xlrd import open_workbook #модуль для работы с Excel-файлами
except ImportError: #если модуль xlrd для работы с Excel-файлами не найден
        ARCPY.AddMessage("Модуль xlrd для работы с Excel-файлами не найден, программа остановлена")
        raise SystemExit(1)#завершить работу программы
#функция в качестве входного параметра берет excel-файл, и возвращает список листов excel-файла
def get_excel_list(excel_file,List_point,List_strat_lit,List_code_strat,List_code_lit,List_mineral,List_type_test,List_code_type_TN,List_code_status_document,List_code_status_test,List_code_type_document):
        def check_unicode(text):#функция преобразует кодировку юникод в utf-8 
                if(type(text)==unicode):text = text.encode('utf-8')
                return text
        def get_object_list(L_name_array,list_name):#ищет лист с заданным именем в списке листов excel-файла
                k=0
                list_object = "null"
                while(k<len(L_name_array)):#перебор листов excel-файла
                        #если лист с названием, которое указано во входных параметрах найден
                        if(str(L_name_array[k])==str(list_name.encode('utf-8'))):
                                list_object = excel_book.sheet_by_name(list_name)
                                break 
                        k+=1
                return list_object #вернуть объект листа
        excel_book = open_workbook(excel_file,formatting_info=True)#получить объект книги excel
        array_list = excel_book.sheet_names()#получить список названий листов excel
        k=0
        name_list_array=[]#массив названий листов к кодировке utf-8
        while (k<len(array_list)):#перебор названий excel-листов
                name_list_array.append(check_unicode(array_list[k]))#добавить в массив текущее название в кодировке utf-8
                k+=1
        #получить объекты листов
        L_point_obj = get_object_list(name_list_array,List_point)#лист точки наблюдений
        L_strat_lit_obj = get_object_list(name_list_array,List_strat_lit)#лист стратиграфия литология
        L_code_strat_obj = get_object_list(name_list_array,List_code_strat)#лист коды стратиграфии
        L_code_lit_obj = get_object_list(name_list_array,List_code_lit)#лист коды литологии
        L_mineral_obj = get_object_list(name_list_array,List_mineral)#лист с минералогией
        L_type_test = get_object_list(name_list_array,List_type_test)#лист с минералогией
        L_code_type_TN = get_object_list(name_list_array,List_code_type_TN)#лист с кодами типов ТН
        L_code_status_document = get_object_list(name_list_array,List_code_status_document)#лист с кодами состояния документирования
        L_code_status_test = get_object_list(name_list_array,List_code_status_test)#лист с кодами состояния опробывания
        L_code_type_document = get_object_list(name_list_array,List_code_type_document)#лист с кодами типов документации
        return  L_point_obj, L_strat_lit_obj, L_code_strat_obj, L_code_lit_obj, L_mineral_obj, L_type_test, \
        L_code_type_TN, L_code_status_document, L_code_status_test,L_code_type_document#вернуть обекты excel-листов
#Функция проверяет, все ли объекты excel-листов созданы, если нет - завершить работу программы
def verify_lists(d_point, d_strat_lit, d_code_strat, d_code_lit, d_mineral, d_code_test, d_type_TN, d_document,d_test,d_type_document):
        if((d_point=="null")or(d_strat_lit=="null")or(d_code_strat=="null")or(d_code_lit=="null")or \
        (d_mineral=="null")or(d_code_test=="null")or(d_type_TN=="null")or(d_document=="null")or \
        (d_test=="null")or(d_type_document=="null")): #проверить все ли объекты excel-листов созданы
                ARCPY.AddMessage("Прочитаны не все excel-листы, программа остановлена")#завершить работу программы
                raise SystemExit(1)#Завершить работу программы
        else:ARCPY.AddMessage("Все листы excel-листа найдены")
#функция в качестве параметра берет объект листа Excel-файла, содержащего список точек наблюдений
#и возвращает массив [id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок
#код статуса документирования,код состаяния опробования,тип документации]
def read_points_wells(list_excel):
	title=True#переменная устанавливает является ли строка шапкой документа
	def_wells_array=[]#массив скважин
	for rownum in range(list_excel.nrows):#перебор всех строк листа
		if title: 
			title=False
			continue #Пропустить чтение "шапки" с иназваниями атрибутов
		row_array=[]#массив строк
		current_row=list_excel.row_values(rownum)#текущая строка листа
		name_object = "no_data"#название объекта
		name_area = "no_data"#название участка
		id_well="0" #идентификатор скважины
		depth_well="-1"#глубина скважины 
		x="0"#координата x
		y="0"#координата y
		z="-1"#высотная отметка
		name_well="no_data" #линия/номер скважины
		type_well = "данные отсутствуют"#код типа ТН
		status_doc = "данные отсутствуют"#код статуса документирования
		status_test = "данные отсутствуют"#код состаяния опробования
		type_doc = "данные отсутствуют"#тип документации
		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
			if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
			if((j==1)and(cell<>'')):name_object=cell
			if((j==2)and(cell<>'')):name_area=cell
			if((j==5)and(cell<>'')):id_well=cell
			if((j==6)and(cell<>'')):type_well=cell
			if((j==7))and(cell<>''):depth_well=cell
			if((j==12)and(cell<>'')):x=cell
			if((j==13)and(cell<>'')):y=cell
			if((j==14)and(cell<>'')):z=cell
			if((j==19)and(cell<>'')):status_doc=cell
			if((j==22)and(cell<>'')):status_test=cell
			if((j==25)and(cell<>'')):name_well=cell
			if((j==26)and(cell<>'')):type_doc = cell
		def_wells_array.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,str(type_well), \
                str(status_doc),str(status_test),str(type_doc)])#добавить текущую скважину к списку скважин
	return def_wells_array #вернуть массив точек наблюдений в формате
        #[id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок,код типа ТН, ]
        #код статуса документирования,код состаяния опробования,тип документации]
#функция в качестве параметра берет объект листа Excel-файла, содержащего коды и их расшифровки
#и возвращает массив [код, расшифровка]
def read_code_wells(list_excel):
	title=True#переменная устанавливает является ли строка шапкой документа
	def_array_code=[]#массив с кодами 
	for rownum in range(list_excel.nrows):#перебор всех строк листа
		if title: 
			title=False
			continue #Пропустить чтение "шапки" с названиями атрибутов
		row_array=[]#массив строк
		current_row=list_excel.row_values(rownum)#текущая строка листа
		def_code="no_data"#код 
		def_name = "no_data"#расшифровка
		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
			if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
			if((j==0)and(cell<>'')):def_code=cell #первый столбец с кодом 
			if((j==1)and(cell<>'')):def_name=cell #второй столбец с расшифровкой
		def_array_code.append([def_code,def_name])#добавить текущую строку к общему массиву 
	return def_array_code #возвращает массив в формате [код, расшифровка]
#функция берет на входе массив с точками наблюдения в формате
#[id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок,код типа ТН, ]
#код статуса документирования,код состаяния опробования,код типа документации]
#и заменяет последние четыре атрибута с кодами на их расшифровки 
def replacement_code_lists(wells_array,def_type_tn,def_status_document,def_status_test,def_type_document):
        array_wells_replace = wells_array #массив с точками наблюдений, в котором коды заменяются на расшифрованные значения
        k=0
        while(k<len(wells_array)):#перебор элементов массива точек (из листа Точки наблюдений)
                m=0
                while(m<len(def_type_tn)):#перебор кодов типов точек наблюдений
                        #если код типа пробы из массива с ТН совпал с кодом из листа с кодами типов ТН
                        if(str(wells_array[k][8])==str(def_type_tn[m][0])):
                                array_wells_replace[k][8]=def_type_tn[m][1]#поменять код на расшифровку
                                break#выйти из внутреннего цикла
                        m+=1
                m=0
                while(m<len(def_status_document)):#перебор кодов типов состояния документирования
                        #если код типа пробы из массива с ТН совпал с кодом из листа с кодами состояния документирования
                        if(str(wells_array[k][9])==str(def_status_document[m][0])):
                                array_wells_replace[k][9]=def_status_document[m][1]#поменять код на расшифровку
                                break#выйти из внутреннего цикла
                        m+=1
                m=0
                while(m<len(def_status_test)):#перебор кодов типов состояния опробования
                        #если код типа пробы из массива с ТН совпал с кодом из листа с кодами состояния опробования
                        if(str(wells_array[k][10])==str(def_status_test[m][0])):
                                array_wells_replace[k][10]=def_status_test[m][1]#поменять код на расшифровку
                                break#выйти из внутреннего цикла
                        m+=1
                m=0
                while(m<len(def_type_document)):#перебор кодов типов документирования
                        #если код типа пробы из массива с ТН совпал с кодом из листа с кодами типов документирования
                        if(str(wells_array[k][11])==str(def_type_document[m][0])):
                                array_wells_replace[k][11]=def_type_document[m][1]#поменять код на расшифровку
                                break#выйти из внутреннего цикла
                        m+=1
                k+=1
        #вернуть массив в формате
        #[id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок,расшифровка типа ТН, ]
        #расшифровка статуса документирования,расшифровка состаяния опробования,расшифровка типа документации]
        return array_wells_replace
#Функция в качестве входного параметра берет объект листа Excel-файла, содержащего сведения о мощности и составе пластов
#и возвращает массив [point_id,кровля,подошва,код возвраста,код породы,описание породы,линия/номер скважины]
def read_strat_lit_wells(list_excel):
	title=True#переменная устанавливает является ли строка шапкой документа
	def_arr_strat_lit_wells=[] #массив со значениями мощности пластов
	for rownum in range(list_excel.nrows):#перебор всех строк листа
		if title: 
			title=False
			continue #Пропустить чтение "шапки" с иназваниями атрибутов
		current_row=list_excel.row_values(rownum)#текущая строка листа
		point_id="no_data"
		roof="no_data" #кровля
		sole="no_data" #подошва
		cod_old="no_data" #код возраста
		cod_rock="no_data" #код породы
		about_rock="no_data" #описание породы
		name_point="no_data" #линия/номер скважины
		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
			if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
			if((j==0)and(cell<>'')):point_id=cell
			if((j==1)and(cell<>'')):roof=cell
			if((j==2)and(cell<>'')):sole=cell
			if((j==3)and(cell<>'')):cod_old=cell
			if((j==4)and(cell<>'')):cod_rock=cell
			if((j==5)and(cell<>'')):about_rock=cell
			if((j==6)and(cell<>'')):name_point=cell
		if((point_id<>"no_data")or(roof<>"no_data")or(sole<>"no_data")or(cod_old<>"no_data")or \
                   (cod_rock<>"no_data")or(about_rock<>"no_data")or(name_point<>"no_data")): #Если не считана пустая строка
			def_arr_strat_lit_wells.append([point_id,roof,sole,cod_old,cod_rock,about_rock,name_point])
	#вернуть массив в формате [point_id,кровля,подошва,код возвраста,код породы,описание породы,линия/номер скважины]
	return def_arr_strat_lit_wells

#функция в качестве входного параметра берет массив, содержащий информацию о пластах 
#и возвращает массив в формате [код точки, кровля пластов, коды стратиграфий, коды литологий]
def single_width_blocking(arr_form):
	def_arr_calc_width=[]#общий массив для записи данных
	well_first=arr_form[0][0]#первое значение идентификации
	k=0
	roof=""#ads кровли
	sole=""#abs подошвы
	code_old=""#код породы
	code_lit=""#код литологии
	while(k<len(arr_form)):
		if (arr_form[k][0]==well_first):#если текущий идентификатор равен предыдущему
        		roof=roof+str(arr_form[k][1])+"/" #добавить abs кровли текущего пласта
        		sole=sole+str(arr_form[k][2])+"/" #добавить abs подошвы текущего пласта
			code_old=code_old+str(arr_form[k][3])+"/" #добавить код возраста текущего пласта
			code_lit=code_lit+str(arr_form[k][4])+"/" #добавить код литологии текущего пласта
    		else: #если курсор перешел к новой скважине
			current_array=[well_first,roof,sole,code_old,code_lit] #запомнить данные предыдущей скважины
			def_arr_calc_width.append(current_array)# и добавить их к общему массиву скважин
			well_first=arr_form[k][0]#установить текущий идентификатор
			roof=str(arr_form[k][1])+"/" #начать новую строку abs кровли текущего пласта
			sole=str(arr_form[k][2])+"/" #начать новую строку abs подошвы текущего пласта
			code_old=str(arr_form[k][3])+"/" # начать новую строку c кодами возраста текущего пласта
			code_lit=str(arr_form[k][4])+"/" # начать новую строку с кодами литологии текущего пласта
    		k+=1
    		if (k==len(arr_form)):#добавить в общий массив данный о последней скважине
        		current_array=[well_first,roof,sole,code_old,code_lit]
			def_arr_calc_width.append(current_array)
	#вернуть массив в формате [код точки, кровли пластов, подошвы пластов, коды стратиграфий, коды литологии]
	return def_arr_calc_width
#функция в качестве входных параметров берет массив в формате[код точки, кровли пластов, коды стратиграфий, коды литологий]
#и возвращает массив с текстовым описанием
#в формате [коды точек, abs кровли/abs кровли/abs кровли,
#abs подошвы/abs подошвы/abs подошвы, Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
def view_strat_lit(arr_code_strat_lit,strat_code,lit_code):
        def_arr_text_strat_lit=[]
        k=0
        while(k<len(arr_code_strat_lit)):#перебор массива скважин формата [код точки, кровли пластов, коды стратиграфий, коды литологий]
                str_strat_code=str(arr_code_strat_lit[k][3])#получить строку кодов стратиграфий для текущей скважины
                arr_strat_code=str_strat_code.split("/")#получить массив кодов стратиграфий для текущей скважины
                str_lit_code=str(arr_code_strat_lit[k][4])#получить строку кодов литологий для текущей скважины
                arr_lit_code=str_lit_code.split("/")#получить массив кодов литологий для текущей скважины
                m=0
                stroka_strat=""#строка со стратиграфией в формате Jmh/Jsn/Ol/
                ##перебор кодов стратиграфий текущей скважины, не обрабатывать последний элемент массива кодов стратиграфий поскольку
		#он является пустым - при переводе строки '1/2/3/' получается массив ['1','2','3',''] 
                while(m<len(arr_strat_code)-1):
                        n=0
                        #перебор массива всех кодов стратиграфий, формат [код, стратиграфия]
                        while(n<len(strat_code)):
                                cod1=float(arr_strat_code[m])#преобразовать в число код стратиграфии из массива arr_code_strat_lit
                                cod2=float(strat_code[n][0])#преобразовать в число код из входног массива strat_code
                                if(cod1==cod2):#если коды совпали
                                        #формировать строку со стратиграфией в формате Jmh/Jsn/Ol/
                                        stroka_strat=stroka_strat+strat_code[n][1]+"/"
                                        #если коды совпали - выйти из текущего цикла(перейти к следующей стратиграфии текущей скважины)
                                        break 
                                n+=1
                        m+=1
                m=0
                stroka_lit=""#строка с литологией в формате алевролит/известняк/доломит/
                #перебор кодов литологий текущей скважины, не обрабатывать последний элемент массива кодов литологий поскольку
		#он является пустым - при переводе строки '1/2/3/' получается массив ['1','2','3',''] 
                while(m<len(arr_lit_code)-1):
                        n=0
                        #перебор массива всех кодов литологии, формат [код, литология]
                        while(n<len(lit_code)):
                                cod1=float(arr_lit_code[m])#преобразовать в число код литологии из массива arr_code_strat_lit
                                cod2=float(lit_code[n][0])#преобразовать в число код из входного массива lit_code
                                if(cod1==cod2):#если коды совпали
                                        #формировать строку с литологией в формате алевролит/известняк/доломит/
                                        stroka_lit=stroka_lit+lit_code[n][1]+"/"
                                        #если коды совпали - выйти из текущего цикла(перейти к следующей литологии текущей скважины)
                                        break   
                                n+=1
                        m+=1
                #сформировать текущий элемент массива в формате [код точки,
                #abs кровли/abs кровли/abs кровли,abs подошвы/abs подошвы/abs подошвы,Jmh/Jsn/Ol/,алевролит/известняк/доломит/]
                current_arr=[arr_code_strat_lit[k][0],arr_code_strat_lit[k][1],arr_code_strat_lit[k][2],stroka_strat,stroka_lit]
                #добавить текущий элемент к выходному массиву
                def_arr_text_strat_lit.append(current_arr)
                k+=1
        #вернуть массив в формате [коды точек, abs кровли/abs кровли/abs кровли,abs подошвы/abs подошвы/abs подошвы,
        #Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
        return def_arr_text_strat_lit
#функция берет объект excel-листа с результатами минералогии
#и возвращает массив с названием атрибутов excel-листа(шапки) и массив считанных строк excel-листа 
def read_mineralogy(list_excel):
        title=True#переменная устанавливает является ли строка шапкой документа\
        def_mineral_arr = []#массив для хранения строк, считанных из листа результатов минералогии
        for rownum in range(list_excel.nrows):#перебор строк листа
                if title:#если указатель находится на первой строке таблицы, соответствующей "шапке"
                        title=False
                        current_row = list_excel.row_values(rownum)
                        arr_name_cell = []#массив с именами атрибутов листа результатов минералогии
                        for j, cell in enumerate(current_row):#перебор ячеек шапки листа результатов минералогии
                                if(type(cell)==unicode):cell=cell.encode('utf-8')
                                arr_name_cell.append(cell)#добавить в массив название атрибута из текущей ячейки
                        continue #начать новую итерацию общего цикла
                current_row = list_excel.row_values(rownum)#текущая строка листа
                curr_string = []#массив для завписи ячееек текущей строки
                for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
                        if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
                        curr_string.append(cell)#добавить в массив значение текущей ячейки
                #если количество ячеек текущей строки равно количесвту ячеек 'шапки'  
                #добавить в массив строк массив текущей строки 
                if(len(curr_string)==len(arr_name_cell)):def_mineral_arr.append(curr_string)
        #вернуть массив с названием атрибутов excel-листа(шапки) и массив считанных строк excel-листа 
        return arr_name_cell,def_mineral_arr
#функция на входе берет массив атрибутов excel-листа и массив со всеми строками excel-листа
#и возвращает массив строк с непустыми значениями
def calc_non_empty(atributes,values):
        k=0
        arr_non_empty = []#массив для записи всех непустых строк 
        while(k<len(values)):#перебор строк листа
                l=0
                stroka_mineral=''
                id_points = 'no_data'#id точек наблюдений
                number_test = 'no_data'#номер пробы
                id_test = 'no_data'#id пробы
                code_type_test = 'no_data'#код типа пробы
                interval_from = 'no_data'#интервал глубины от
                interval_to = 'no_data'#интервал глубины до
                curr_arr = []#массив для записи атрибутов текущей строки
                while(l<len(atributes)):#перебор атрибутов текущей строки
                        #если считываются данные о миниралах
                        #значение с индексом len(atributes)-1 это значение атрибута UIN
                        if(l==0):id_points = str(values[k][l])#id точек наблюдений
                        if(l==1):number_test = str(values[k][l])#номер пробы
                        if(l==2):id_test = str(values[k][l])#id пробы
                        if(l==3):code_type_test = str(values[k][l])#код типа пробы
                        if(l==4):interval_from = str(values[k][l])#интервал глубины от
                        if(l==5):interval_to = str(values[k][l])#интервал глубины до
                        if((l>=8)and(l<len(atributes)-1)):
                                #Если значение атрибута минерала не пустое, значит это находка минерала
                                if(values[k][l]<>''):
                                        stroka_mineral=stroka_mineral+str(atributes[l])+"="+str(values[k][l])+"/"
                                        curr_arr = [id_points,number_test,id_test,code_type_test,interval_from,interval_to,stroka_mineral]
                        l+=1
                #если массив текущей строки не пустой
                #добавить в общий массив непустых строк
                if(len(curr_arr)<>0):arr_non_empty.append(curr_arr)
                k+=1
        #вернуть массив в формате [id скважины, номер пробы, id пробы, код типа пробы, интервал глубин от,
        #интервал глубин до, находки минералов]
        return arr_non_empty

#функция в качестве параметра берет объект листа Excel-файла, содержащего коды типов проб
#и возвращает массив [код пробы, наименование пробы]
def read_code_type(list_excel):
	title=True#переменная устанавливает является ли строка шапкой документа
	def_array_type_test=[]#массив с кодами проб
	for rownum in range(list_excel.nrows):#перебор всех строк листа
		if title: 
			title=False
			continue #Пропустить чтение "шапки" с иназваниями атрибутов
		row_array=[]#массив строк
		current_row=list_excel.row_values(rownum)#текущая строка листа
		def_code_type_test="no_data"#код пробы
		def_name_type_test = "no_data"#название пробы
		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
			if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
			if((j==0)and(cell<>'')):def_code_type_test=cell #первый столбец с кодом пробы
			if((j==1)and(cell<>'')):def_name_type_test=cell #второй столбец с наименованием пробы
		def_array_type_test.append([def_code_type_test,def_name_type_test])#добавить текущую строку к общему массиву 
	return def_array_type_test #возвращает массив в формате [код литологии, нименование литологии]

def replacement_code_test(def_non_empty,def_code_type_test):
        k=0
        while(k<len(def_non_empty)):#перебор элементов массива точек с минералогией
                m=0
                while(m<len(def_code_type_test)):#перебор кодов типов пород
                        #если код типа пробы из массива с минералогоией совпал с кодом из листа с кодами типов проб
                        if(def_non_empty[k][3]==str(def_code_type_test[m][0])):
                                def_non_empty[k][3]=def_code_type_test[m][1]#поменять код типа пробы на название типа пробы(шлам, шлих и т.п.)
                                break#выйти из внутреннего цикла
                        m+=1
                k+=1
        return def_non_empty
        
#функция на входе берет массив точек наблюдений с координатами и массив результатов минералогии
#и возвращает объединенный массив в формате
#[id_скважины,линия/имя скважины,x,y,z,объект,участок,номер пробы,id пробы,код типа пробы,интервал глубины от, интервал глубины до, находки МСА]
def united_two_arrays(arr_wells, arr_min):
	#общий массив, который формируется на основе массива листа точек наблюдений и листа с минералогией
	#с первого листа берутся сведения о координатах, номере скважины, 
	#идентификатора, из второго - данные о пробах и находкам минералов спутников алмазов
	def_united_array=[]
	k=0
	while(k<len(arr_min)):#перебор элементов первого листа (массива точек)
    		m=0
    		while(m<len(arr_wells)):#перебор элементов второго листа (массива с данными минералогии) 
        		if(arr_min[k][0]==str(arr_wells[m][0])):#если идентификаторы совпали
				#значит это данные одной и той же скважины,
				#поэтому сформировать один общий массив
				curr_array = [arr_wells[m][0],arr_wells[m][1],arr_wells[m][2],arr_wells[m][3],arr_wells[m][4], \
                                arr_wells[m][5],arr_wells[m][6],arr_wells[m][7],arr_min[k][1],arr_min[k][2],arr_min[k][3], \
                                arr_min[k][4],arr_min[k][5],arr_min[k][6],arr_wells[m][8],arr_wells[m][9],arr_wells[m][10], \
                                arr_wells[m][11]]
            			def_united_array.append(curr_array) #добавить в общий массив точек
	    			break
        		m=m+1
    		k=k+1
    	#вернуть объединенный массив с координатами и данными по находкам МСА
    	#[id_скважины,линия/имя скважины,x,y,z,объект,участок,номер пробы,id пробы,код типа пробы,интервал глубины от,
    	#интервал глубины до, находки МСА,тип скважины, статус документирования,состаяние опробования,
    	#тип документации]
	return def_united_array


# функция на входе берет массив в формате[id_скважины,линия/имя скважины,x,y,z,объект,участок,
#номер пробы,id пробы,тип пробы,интервал глубины от, интервал глубины до, находки МСА,
#тип скважины, статус документирования,состаяние опробования,тип документации]
#и массив в формате[коды точек, abs кровли/abs кровли/abs кровли,abs подошвы/abs подошвы/abs подошвы,
#Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
#и объединяет эти массивы
def create_final_array(arr_min,arr_strat_lit):
	def_united_array=[]#объединенный массив
	if(len(arr_strat_lit)==0):#если массив с данными по литостратиграфии пуст, вернуть массив с пустыми значениями по литостратиграфии
                k=0
                while(k<len(arr_min)):#перебор элементов первого листа (массива точек c минералогией)
                        #все равно добавить к формируемому массиву данные по текущей скважине без данных по литостратиграфии
                        curr_array = [arr_min[k][0],arr_min[k][1],arr_min[k][2], \
                        arr_min[k][3],arr_min[k][4],arr_min[k][5],arr_min[k][6], \
                        arr_min[k][7],arr_min[k][8],arr_min[k][9],arr_min[k][10], \
                        arr_min[k][11],arr_min[k][12],arr_min[k][13],"","","","", \
                        arr_min[k][14],arr_min[k][15],arr_min[k][16],arr_min[k][17]]
            		def_united_array.append(curr_array) #добавить в общий массив точек
                        k+=1
                #вернуть массив с пустыми значениями литостратиграфии
                return def_united_array
	k=0
	while(k<len(arr_min)):#перебор элементов первого листа (массива точек c минералогией)
                find_stratolit=False#логическая переменная определяет, есть ли для текущей скважины информация по литостратиграфии пластов
    		m=0
    		while(m<len(arr_strat_lit)):#перебор элементов второго листа (массива с данными литостратиграфии) 
        		if(str(arr_min[k][0])==str(arr_strat_lit[m][0])):#если идентификаторы совпали
				#значит это данные одной и той же скважины,
				#поэтому сформировать один общий массив
				curr_array = [arr_min[k][0],arr_min[k][1],arr_min[k][2], \
                                arr_min[k][3],arr_min[k][4],arr_min[k][5],arr_min[k][6], \
                                arr_min[k][7],arr_min[k][8],arr_min[k][9],arr_min[k][10], \
                                arr_min[k][11],arr_min[k][12],arr_min[k][13],arr_strat_lit[m][1],arr_strat_lit[m][2], \
                                arr_strat_lit[m][3],arr_strat_lit[m][4], \
                                arr_min[k][14],arr_min[k][15],arr_min[k][16],arr_min[k][17]]
            			def_united_array.append(curr_array) #добавить в общий массив точек
            			find_stratolit=True#для текущей скважины есть информация по литостратиграфии пластов
        		m=m+1
        		#если достигнут конец цикла и при этом не найдено совпадения кодов(т.е. для текущей скважины отсутствуют данные по литостратиграфии пластов)
        		if(m==len(arr_strat_lit))and(not find_stratolit):
                                #все равно добавить к формируемому массиву данные по текущей скважине без данных по литостратиграфии
                                curr_array = [arr_min[k][0],arr_min[k][1],arr_min[k][2], \
                                arr_min[k][3],arr_min[k][4],arr_min[k][5],arr_min[k][6], \
                                arr_min[k][7],arr_min[k][8],arr_min[k][9],arr_min[k][10], \
                                arr_min[k][11],arr_min[k][12],arr_min[k][13],"","","","", \
                                arr_min[k][14],arr_min[k][15],arr_min[k][16],arr_min[k][17]]
            			def_united_array.append(curr_array) #добавить в общий массив точек
    		k=k+1
    	#вернуть массив в формате [id_скважины,линия/имя скважины,x,y,z,объект,участок,
    	#номер пробы,id пробы,тип пробы,интервал глубины от, интервал глубины до, находки МСА,
    	#abs кровли/abs кровли/abs кровли,abs подошвы/abs подошвы/abs подошвы,
    	#Jmh/Jsn/Ol/, алевролит/известняк/доломит/, тип скважины, статус документирования,состаяние опробования,
    	#тип документации]
    	return def_united_array
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
        gp.AddField(lay,"id_скважины","STRING",15,15)#атрибут id скважин
	gp.AddField(lay,"название_скважины","STRING",15,15)#атрибут названия скважин
	gp.AddField(lay,"x","FLOAT",15,15)#добавить атрибут x
	gp.AddField(lay,"y","FLOAT",15,15)#добавить атрибут y
	gp.AddField(lay,"z","FLOAT",15,15)#добавить атрибут z
	gp.AddField(lay,"глубина_скважины","STRING","","",50)#атрибут - глубина скважины
	gp.AddField(lay,"название_объекта","STRING","","",100)#атрибут - название объекта
	gp.AddField(lay,"название_участка","STRING","","",100)#атрибут - название участка
	gp.AddField(lay,"номер_пробы","LONG","","",50)#атрибут - номер пробы
	gp.AddField(lay,"id_пробы","LONG","","",50)#атрибут - id пробы
	gp.AddField(lay,"тип_пробы","STRING","","",50)#атрибут - код типа пробы
	gp.AddField(lay,"интервал_пробы_от","FLOAT","","",50)#атрибут - интервал глубины от
	gp.AddField(lay,"интервал_пробы_до","FLOAT","","",50)#атрибут - интервал глубины до
	gp.AddField(lay,"находки_МСА","STRING","","",1000)#атрибут - информация по МСА
	gp.AddField(lay,"отметка_кровли_пластов","STRING","","",1000)#атрибут для глубин кровли пластов
        gp.AddField(lay,"отметка_подошвы_пластов","STRING","","",1000)#атрибут для глубин подошвы пластов
        gp.AddField(lay,"стратиграфия","STRING","","",1000)#атрибут для стратиграфии
        gp.AddField(lay,"литология","STRING","","",1000)#атрибут для литологии
        gp.AddField(lay,"тип_точки_наблюдения","STRING","","",1000)#атрибут для типа точки наблюдения
        gp.AddField(lay,"состояние_документирования","STRING","","",1000)#атрибут для состояния документирования
        gp.AddField(lay,"состояние_опробования","STRING","","",1000)#атрибут для состояния опробования
        gp.AddField(lay,"тип_документирования","STRING","","",1000)#атрибут для типа документирования
        #создать атрибутивные поля для находок МСА
        u=8
        while(u<len(atr_min)-1):#не учитывать последний элемент, поскольку он не относится к минералогии(UIN)
                gp.AddField(lay,atr_min[u],"LONG","","",50)
                u+=1
	rows = gp.InsertCursor(lay)
	pnt = gp.CreateObject("Point")
	k=0
	#[id_скважины,линия/имя скважины,x,y,z,объект,участок,
    	#номер пробы,id пробы,тип пробы,интервал глубины от, интервал глубины до, находки МСА,
    	#abs кровли/abs кровли/abs кровли,abs подошвы/abs подошвы/abs подошвы,
    	#Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
	while(k<len(array_map)): #создание точечного слоя
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
    		abs_litostrat_roof = str(array_map[k][14])
    		abs_litostrat_sole = str(array_map[k][15])
    		stratigraphy = str(array_map[k][16])
    		lithology = str(array_map[k][17])
    		type_tn = str(array_map[k][18])
    		stat_doc = str(array_map[k][19])
    		stat_test = str(array_map[k][20])
    		type_doc = str(array_map[k][21])
    		feat = rows.NewRow()#Новая строка
    		pnt.id = k #ID
    		pnt.x = float(array_map[k][3])
    		pnt.y = float(array_map[k][2])
    		feat.shape = pnt
    		feat.SetValue("id_скважины",id_well)#добавить значения к атрибутивным полям
    		feat.SetValue("название_скважины",name_well)
		feat.SetValue("x",float(array_map[k][3]))
		feat.SetValue("y",float(array_map[k][2]))
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
                feat.SetValue("отметка_кровли_пластов",abs_litostrat_roof)
                feat.SetValue("отметка_подошвы_пластов",abs_litostrat_sole)
                feat.SetValue("стратиграфия",stratigraphy)
                feat.SetValue("литология",lithology)
                feat.SetValue("тип_точки_наблюдения",type_tn)
                feat.SetValue("состояние_документирования",stat_doc)
                feat.SetValue("состояние_опробования",stat_test)
                feat.SetValue("тип_документирования",type_doc)
                #заполнить нулевыми значениями атрибутивные поля для находок МСА
                u=8
                while(u<len(atr_min)-1):#не учитывать последний элемент, поскольку он не относится к минералогии(UIN)
                        feat.SetValue(atr_min[u],0) 
                        u+=1
                #перебор элементов шапки листа с минералогией
                i=8 #начать перебор с элемента pir05_0
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
if __name__== '__main__':#если модуль запущен, а не импортирован
        #-------------------------------входные данные------------------------------------
        InputExcelFile = ARCPY.GetParameterAsText(0)#входной excel-файл
        Coordinate_System = ARCPY.GetParameterAsText(1)#координатная система выходного точечного слоя
        OutputFC = ARCPY.GetParameterAsText(2)#выходной точечный слой скважин
        Full_report = ARCPY.GetParameterAsText(3)#параметр определяет выводить ли полный отчет
        List_point = ARCPY.GetParameterAsText(4)#Название листа с точками наблюдений 
        List_strat_lit = ARCPY.GetParameterAsText(5)#Название листа с описанием стратиграфии и литологии
        List_code_strat = ARCPY.GetParameterAsText(6)#Название листа с кодами стратиграфии
        List_code_lit = ARCPY.GetParameterAsText(7)#Название листа с кодами литологии
        List_mineral = ARCPY.GetParameterAsText(8)#Название листа с данными минералогии
        List_code_type_test = ARCPY.GetParameterAsText(9)#Название листа с кодами типов проб
        List_code_type_TN = ARCPY.GetParameterAsText(10)#Название листа с кодами типов ТН
        List_code_status_document = ARCPY.GetParameterAsText(11)#Название листа с кодами состояния документирования
        List_code_status_test = ARCPY.GetParameterAsText(12)#Название листа с кодами состояния опробывания
        List_code_type_document = ARCPY.GetParameterAsText(13)#Название листа с кодами типов документирования
        #---------------------------------------------------------------------------------
        #получить объект excel-листов
        obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_mineral_list,obj_code_type_test, \
        obj_code_type_TN, obj_code_status_document, obj_code_status_test, obj_code_type_document = \
        get_excel_list(InputExcelFile, List_point,List_strat_lit,List_code_strat,List_code_lit,List_mineral, \
        List_code_type_test,List_code_type_TN,List_code_status_document,List_code_status_test,List_code_type_document)
        #проверить все ли excel-листы с именами, указанными как входные параметры, найдены
        verify_lists(obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_mineral_list,obj_code_type_test, \
        obj_code_type_TN,obj_code_status_document,obj_code_status_test,obj_code_type_document)
        wells_array = read_points_wells(obj_point)#получить массив точек из листа excеl-файла с точками наблюдений
        if(len(wells_array)==0):#если массив точек наблюдений пуст - завершить работу программы
                ARCPY.AddMessage("Лист точек наблюдений пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        ARCPY.AddMessage("Общее количество точек наблюдений, прочитанных из Excel-файла: "+str(len(wells_array)))
        ARCPY.AddMessage(wells_array[0])
        if(Full_report == "true"):#выводить полный отчет
                if(len(wells_array)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(wells_array[k][0])+" "+str(wells_array[k][1])+" "+ \
                                str(wells_array[k][2])+" "+str(wells_array[k][3])+" "+str(wells_array[k][4])+" "+ \
                                str(wells_array[k][5])+" "+str(wells_array[k][6])+" "+str(wells_array[k][7])+" "+ \
                                str(wells_array[k][8])+" "+str(wells_array[k][9])+" "+str(wells_array[k][10])+" "+ \
                                str(wells_array[k][11]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_tn = read_code_wells(obj_code_type_TN)#получить массив с кодами типов точек наблюдений
        ARCPY.AddMessage("Количество кодов с типами точек наблюдения: "+str(len(code_type_tn)))
        if(len(code_type_tn)==0): #если массив кодов типов ТН пуст - завершить работу программы
                ARCPY.AddMessage("Лист с кодами типов ТН пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):
                if(len(code_type_tn)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_type_tn[k][0])+" "+str(code_type_tn[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_status_document = read_code_wells(obj_code_status_document)#получить массив с кодами состояния документирования
        ARCPY.AddMessage("Количество кодов состояния документирования: "+str(len(code_status_document)))
        if(len(code_status_document)==0):#если массив кодов состояния документирования пуст - завершить работу программы
                ARCPY.AddMessage("Лист с кодами состояния документирования пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):
                ARCPY.AddMessage("В формате:")
                k=0
                while(k<len(code_status_document)):
                        ARCPY.AddMessage(str(code_status_document[k][0])+" "+str(code_status_document[k][1]))
                        k+=1
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_status_test = read_code_wells(obj_code_status_test)# получить массив с кодами состояния опробования
        ARCPY.AddMessage("Количество кодов состояния опробования: "+str(len(code_status_test)))
        if(len(code_status_test)==0):#если массив кодов состояния опробования пуст - завершить работу программы
                ARCPY.AddMessage("Лист с кодами состояния опробования пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):
                ARCPY.AddMessage("В формате:")
                k=0
                while(k<len(code_status_test)):
                        ARCPY.AddMessage(str(code_status_test[k][0])+" "+str(code_status_test[k][1]))
                        k+=1
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_document = read_code_wells(obj_code_type_document)#получить массив с кодами типов документирования
        ARCPY.AddMessage("Количество кодов типов документирования: "+str(len(code_type_document)))
        if(len(code_type_document)==0):#если массив кодов типов документирования пуст - завершить работу программы
                ARCPY.AddMessage("Лист с кодами типов документирования пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):
                ARCPY.AddMessage("В формате:")
                k=0
                while(k<len(code_type_document)):
                        ARCPY.AddMessage(str(code_type_document[k][0])+" "+str(code_type_document[k][1]))
                        k+=1
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить массив - копию массива ТН, где заменены четыре поля с кодами на их расшифровки
        wells_array_final = replacement_code_lists(wells_array,code_type_tn,code_status_document,code_status_test, \
        code_type_document)
        ARCPY.AddMessage("Общее количество точек наблюдений, с расшифрованными кодами общей информации по скважинам: "+ \
        str(len(wells_array_final)))
        if(Full_report == "true"):#выводить полный отчет
                if(len(wells_array)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(wells_array_final[k][0])+" "+str(wells_array_final[k][1])+" "+ \
                                str(wells_array_final[k][2])+" "+str(wells_array_final[k][3])+" "+str(wells_array_final[k][4])+" "+ \
                                str(wells_array_final[k][5])+" "+str(wells_array_final[k][6])+" "+str(wells_array_final[k][7])+" "+ \
                                str(wells_array_final[k][8])+" "+str(wells_array_final[k][9])+" "+str(wells_array_final[k][10])+" "+ \
                                str(wells_array_final[k][11]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_strat = read_code_wells(obj_code_strat)#получить массив с кодами стратиграфии
        if(len(code_strat)==0):#если массив кодов стратиграфии пуст - завершить работу программы
                ARCPY.AddMessage("Лист кодов стратиграфии пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        ARCPY.AddMessage("Количество кодов стратиграфии: "+str(len(code_strat)))
        if(Full_report == "true"):#выводить полный отчет
                if(len(code_strat)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_strat[k][0])+" "+str(code_strat[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_lit = read_code_wells(obj_code_lit)#получить массив с кодами литологии
        ARCPY.AddMessage("Количество кодов литологии: "+str(len(code_lit)))
        if(len(code_lit)==0):#если массив кодов литологии пуст - завершить работу программы
                ARCPY.AddMessage("Лист кодов литологии пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):#выводить полный отчет
                if(len(code_lit)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_lit[k][0])+" "+str(code_lit[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        arr_formation = read_strat_lit_wells(obj_strat_lit) #получить массив с данными пластов скважин
        ARCPY.AddMessage("Количество пластов, считанных из файла: "+str(len(arr_formation)))
        if(Full_report == "true"):#выводить полный отчет
                if(len(arr_formation)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(arr_formation[k][0])+" "+str(arr_formation[k][1])+" "+ \
                                str(arr_formation[k][2])+" "+str(arr_formation[k][3])+" "+str(arr_formation[k][4])+ \
                                " "+str(arr_formation[k][5])+" "+str(arr_formation[k][6]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")       
        if(len(arr_formation)==0):#если массив литологии пластов пуст
                arr_single_width = []#создать пустой массив с сгруппированными пластами для каждой скважины
                ARCPY.AddMessage("Информация о литологии отсутствует")
        else:#если массив литологии не пуст
                arr_single_width = single_width_blocking(arr_formation) #получить массив с сгруппированными пластами для каждой скважины
        ARCPY.AddMessage("Количество скважин с информацией о вскрытых пластах: "+str(len(arr_single_width)))
        if(len(arr_single_width)==0):#если массив с сгруппированными пластами для каждой скважины пуст 
                arr_view_strat_lit=[]#создать пустой массив с текстовым описанием стратиграфии и литологии в формате [коды точек, глубина/глубина/глубина,Jmh/Jsn/Ol/
                ARCPY.AddMessage("Массив с группированными пластами для каждой скважины пуст")
        else:
                if(Full_report == "true"):#выводить полный отчет
                        ARCPY.AddMessage("В формате:")
                        ARCPY.AddMessage(str(arr_single_width[0][0])+" "+str(arr_single_width[0][1])+ \
                        " "+str(arr_single_width[0][2])+" "+str(arr_single_width[0][3]))
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
                        #получить массив с текстовым описанием стратиграфии и литологии в формате [коды точек, глубина/глубина/глубина,Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
                arr_view_strat_lit = view_strat_lit(arr_single_width,code_strat,code_lit)
        ARCPY.AddMessage("Количество скважин с текстовым описанием стратиграфии и литологии: "+str(len(arr_view_strat_lit)))
        if(len(arr_view_strat_lit)==0):#если нет скважин со стратолитологическим описанием
                ARCPY.AddMessage("Нет информации по стратиграфии и литологии")
        else:#если есть скважины со стратолитологическим описанием
                if(Full_report == "true"):#выводить полный отчет
                        ARCPY.AddMessage("В формате:")
                        ARCPY.AddMessage(str(arr_view_strat_lit[0][0])+" "+str(arr_view_strat_lit[0][1])+" "+str(arr_view_strat_lit[0][2])+\
                        " "+str(arr_view_strat_lit[0][3])+" "+str(arr_view_strat_lit[0][4]))
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")         
        #получить массив с названием атрибутов excel-листа(шапки) минералогии и массив считанных строк excel-листа 
        list_atributes,value_atributes=read_mineralogy(obj_mineral_list)   
        #вернуть массив непустык строк листа минералогии
        non_empty = calc_non_empty(list_atributes,value_atributes)
        if (len(non_empty)==0):
                ARCPY.AddMessage("Нет данных по минералогии")
                raise SystemExit(1)#завершить работу программы
        ARCPY.AddMessage("Общее количество точек наблюдей с находками минералов - спутников алмазов: "+str(len(non_empty)))
        if(Full_report=="true"):#выводить полный отчет
                if (len(non_empty)>10):
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(non_empty[k]))
                                k+=1
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_test = read_code_type(obj_code_type_test)#получить массив с кодами типов проб
        ARCPY.AddMessage("Количество кодов c кодами проб: "+str(len(code_type_test)))
        if(len(code_type_test)==0):#если массив кодов с типами проб пуст - завершить работу программы
                ARCPY.AddMessage("Лист кодов типов проб, пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):#выводить полный отчет
                if(len(code_type_test)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_type_test[k][0])+" "+str(code_type_test[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить массив, где вместо кодов типов пород - реальные названия, такие как шлих, шлам и т.д.
        arr_mineral = replacement_code_test(non_empty,code_type_test)
        ARCPY.AddMessage("Общее количество точек наблюдей с расшифрованными кодами проб: "+str(len(arr_mineral)))
        if(Full_report == "true"):#выводить полный отчет
                if(len(arr_mineral)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                m=0
                                stroka=''
                                while(m<7):
                                        stroka = stroka+' '+str(arr_mineral[k][m])
                                        m+=1
                                ARCPY.AddMessage(stroka)
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")        
        #получить объедененный массив с координатами и данными минералогии
        united_array = united_two_arrays(wells_array_final, arr_mineral)
        ARCPY.AddMessage("Количество элементов объединенного массива данных скважин и минералогии: "+str(len(united_array)))
        if(len(united_array)==0):#если объединенный массив пуст
                ARCPY.AddMessage("Массив данных скважин и кмв пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате:")
                if(len(arr_mineral)>5):
                        k=0
                        while(k<5):
                                ARCPY.AddMessage(str(united_array[k][0])+" "+str(united_array[k][1])+" "+str(united_array[k][2])+ \
                                " "+str(united_array[k][3])+" "+str(united_array[k][4])+" "+str(united_array[k][5])+ \
                                " "+str(united_array[k][6])+" "+str(united_array[k][7])+" "+str(united_array[k][8])+ \
                                " "+str(united_array[k][9])+" "+str(united_array[k][10])+" "+str(united_array[k][11])+ \
                                " "+str(united_array[k][12])+" "+str(united_array[k][13])+" "+str(united_array[k][14])+ \
                                " "+str(united_array[k][15])+" "+str(united_array[k][16])+" "+str(united_array[k][17]))
                                k+=1
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить финальный массив для выввода в картографический слой
        #этот финальный массив формируется на основе объединения массива с минералогией и координатами
        #и массива с литостратиграфией
        final_array = create_final_array(united_array,arr_view_strat_lit)
        ARCPY.AddMessage("Количество элементов объединенного массива для вывода в картографический слой "+str(len(final_array)))
        if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате:")
                m=0
                stroka=''
                while(m<22):
                        stroka = stroka+' '+str(final_array[0][m])
                        m+=1
                ARCPY.AddMessage(stroka)
        #вывести полученный массив в картографический слой 
        write_to_layer(OutputFC, Coordinate_System, final_array, list_atributes)

