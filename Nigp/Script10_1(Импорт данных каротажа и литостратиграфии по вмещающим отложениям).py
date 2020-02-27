#Скрипт в качесвтве входных параметров берет excel-файл импортированный из ИСИХОГИ и содержащий данные
#по точкам наблюдений, извлекает из него эти данные и записывает их в картографический слой в формате:
#[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,кмв/кмв...,
#мощность перекрывающих отложений, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
#Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
import arcpy as ARCPY
import arcgisscripting
import os
try:
	from xlrd import open_workbook #модуль для работы с Excel-файлами
except ImportError: #если модуль xlrd для работы с Excel-файлами не найден
        ARCPY.AddMessage("Модуль xlrd для работы с Excel-файлами не найден, программа остановлена")
        raise SystemExit(1)#завершить работу программы
#функция в качестве входного параметра берет excel-файл, и возвращает список листов excel-файла
def read_excel_list(excel_file,List_point,List_strat_lit,List_code_strat,List_code_lit,List_GIS):
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
        L_GIS_obj = get_object_list(name_list_array,List_GIS)#лист ГИС
        return L_point_obj, L_strat_lit_obj, L_code_strat_obj, L_code_lit_obj, L_GIS_obj
#Функция проверяет, все ли объекты excel-листов созданы, если нет - завершить работу программы
def verify_lists(d_point, d_strat_lit, d_code_strat, d_code_lit, d_GIS):
        if((d_point=="null")or(d_strat_lit=="null")or(d_code_strat=="null")or(d_code_lit=="null")or \
        (d_GIS=="null")): #проверить все ли объекты excel-листов созданы
                ARCPY.AddMessage("Прочитаны не все excel-листы, программа остановлена")#завершить работу программы
                raise SystemExit(1)#Завершить работу программы
#функция в качестве параметра берет объект листа Excel-файла, содержащего список точек наблюдений
#и возвращает массив [id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок]
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
		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
			if (type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
			if ((j==1)and(cell<>'')):name_object=cell
			if ((j==2)and(cell<>'')):name_area=cell
			if ((j==5)and(cell<>'')):id_well=cell
			if ((j==7))and(cell<>''):depth_well=cell
			if ((j==12)and(cell<>'')):x=cell
			if ((j==13)and(cell<>'')):y=cell
			if ((j==14)and(cell<>'')):z=cell
			if ((j==25)and(cell<>'')):name_well=cell
		def_wells_array.append([id_well,name_well,x,y,z,depth_well,name_object,name_area])#добавить текущую скважину к списку скважин
	return def_wells_array #вернуть массив точек наблюдений в формате [id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок]
#функция в качестве параметра берет объект листа Excel-файла, содержащего коды стратиграфии
#и возвращает массив [код стратиграфии, наименование стратиграфии]
def read_strat_wells(list_excel):
	title=True#переменная устанавливает является ли строка шапкой документа
	def_array_strat=[]#массив с кодами стратиграфии
	for rownum in range(list_excel.nrows):#перебор всех строк листа
		if title: 
			title=False
			continue #Пропустить чтение "шапки" с названиями атрибутов
		row_array=[]#массив строк
		current_row=list_excel.row_values(rownum)#текущая строка листа
		def_code_strat="no_data"#код стратиграфии
		def_name_strat="no_data" #название стратиграфии
		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
			if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
			if((j==0)and(cell<>'')):def_code_strat=cell #первый столбец с кодом стратиграфии
			if((j==1)and(cell<>'')):def_name_strat=cell #второй столбец с наименованием стратиграфии
		def_array_strat.append([def_code_strat,def_name_strat])#добавить текущую строку к общему массиву 
	return def_array_strat #возвращает массив в формате [код стратиграфии, именование стратиграфии]
#Функция в качестве входных данных берет массив кодов стратиграфии и массив с символами кимберлитовмещающих пород
#и возвращает массив в формате [код стратиграфии, наименование стратиграфии, код перекрывающих(P) или вмещающих пород(V)]
def calc_type_rock(arr_strat, country):
	def_code_strat_new=[]#общий массив для записи
	k = 0
	while (k<len(arr_strat)):#перебор массива с кодами стратиграфии
		m=0
		b=False #переменная определяет были ли найден символы геологического периода, соответствующего вмещающим породам
		while (m<len(country)):#перебор массива с символами <<вмещающих>> геологических периодов
			if(arr_strat[k][1].count(country[m])):#если была обнаружена вмещающая порода
				#добавить в массив вмещающую породу
				def_code_strat_new.append([arr_strat[k][0],arr_strat[k][1],"V"])
				b=True #были найдены вмещающие породы
				break#выйти из внутреннего цикла
			m+=1
		if (b==False):#если не было совпадений с вмещающими породами
			#добавить в массив <<перекрывающую>> породу
			def_code_strat_new.append([arr_strat[k][0],arr_strat[k][1],"P"])
		k+=1
	#вернуть массив в формате [код стратиграфии, обозначение стратиграфии, код отложений (перекрывающие или вмещающие)]
	return def_code_strat_new
#функция в качестве параметра берет объект листа Excel-файла, содержащего коды литологии
#и возвращает массив [код литологии, наименование литологии]
def read_lit_wells(list_excel):
	title=True#переменная устанавливает является ли строка шапкой документа
	def_array_lit=[]#массив с кодами стратиграфии
	for rownum in range(list_excel.nrows):#перебор всех строк листа
		if title: 
			title=False
			continue #Пропустить чтение "шапки" с иназваниями атрибутов
		row_array=[]#массив строк
		current_row=list_excel.row_values(rownum)#текущая строка листа
		def_code_lit="no_data"#код литологии
		def_name_lit = "no_data"#название литологии
		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
			if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
			if((j==0)and(cell<>'')):def_code_lit=cell #первый столбец с кодом литологии
			if((j==1)and(cell<>'')):def_name_lit=cell #второй столбец с наименованием литологии
		def_array_lit.append([def_code_lit,def_name_lit])#добавить текущую строку к общему массиву 
	return def_array_lit #возвращает массив в формате [код литологии, нименование литологии]
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
                ##перебор кодов литологий текущей скважины, не обрабатывать последний элемент массива кодов литологий поскольку
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
#функция берет в качестве параметров массив [код точки, кровли пластов, подощвы пластов, коды стратиграфий, коды литологий] 
#и массив [код стратиграфии, наименование стратиграфии, код перекрывающих(P) или вмещающих пород(V)]
#и возвращает массив [код скважины, мощность перекрывающих]
def calc_width_blocking(arr_width, code_strat):
	def_arr_width_blocking=[]
	k=0
	while(k<len(arr_width)):#перебор массива скважин формата [код точки, кровли пластов, коды стратиграфий, коды литологий]
		find_fraction=False#логическая переменная определяет найдена ли кимберлитовмещающая порода
		str_roof = str(arr_width[k][1])#получить строку abs кровель текущей скважины
		arr_roof = str_roof.split("/")#получить массив abs кровель для текущей скважины
		str_code = str(arr_width[k][3])#получить строку кодов стратиграфий для текущей скважины
		arr_code = str_code.split("/")#получить массив кодов стратиграфий для текущей скважины
		m=0
		##перебор кодов стратиграфий текущей скважины, не обрабатывать последний элемент массива кодов стратиграфий поскольку
		#он является пустым - при переводе строки '1/2/3/' получается массив ['1','2','3',''] 
		while(m<len(arr_code)-1):
			n=0
			if(find_fraction):break#выйти из этого цикла если кимберлитовмещающая порода уже была обнаружена. Перейти к следующей скважине
			#перебор всех кодов стратиграфий массива [код стратиграфии, наименование стратиграфии, код перекрывающих(P) или вмещающих пород(V)]
			while(n<len(code_strat)):
				#если код стратиграфии текущего пласта совпадает с кодом из массива типов пород(перекрывающие/вмещающие)
				cod1=float(arr_code[m])#преобразовать в число код стратиграфии из массива arr_width
				cod2=float(code_strat[n][0])#преобразовать в число код из массива code_strat_new 
				if(cod1==cod2):#если коды совпали
					if(code_strat[n][2]=="V"):#если этому коду соответствует вмещающая порода
						current_array=[arr_width[k][0],arr_roof[m]]#массив для текущей скважины [код скважины, мощность перекрывающих]
						def_arr_width_blocking.append(current_array)#добавить текущую скважину в общий массив
						find_fraction=True#перейти к следующей скважине
						break#выйти из внутреннего цикла
				n+=1
			m+=1
			#если все пласты текущей скважины проанализированы, но при этом не обнаружены вмещающие породы
			#вывести код этой скважены на экран
			if((m==len(arr_code)-1)and(not find_fraction)):
				ARCPY.AddMessage("Мощность перекрывающих отложений не определена. Код точки: "+str(arr_width[k][0]))
		k+=1
	return def_arr_width_blocking #вернуть массив [код скважины, мощность перекрывающих]
#Функция в качестве входных параметров берет объект листа Excel-файла, содержащего список измерений ГИС,
#а также название столбца со значениями измерений для конкретного метода скважинной магниторазведки
#и возвращает массив [id скважины, глубина измерений, значение измерений ГИС]
def read_gis_wells(list_excel,method_gis):
	title=True#переменная устанавливает является ли строка шапкой документа
	number_gis=0#номер столбца, в котором содержатся значения измерений ГИС
	def_gis_arrays = []#массив для хранения массивов [id_скважины, глубина измерений, значения измерений ГИС]
	find_list = False #переменная определяет, найден ли столбец со значениями ГИС
	for rownum in range(list_excel.nrows):#перебор всех строк листа
		if title:#если указатель находится на первой строке таблице, соответствующей "шапке"
			title=False
        		current_row=list_excel.row_values(rownum)#текущая строка листа
			for j, cell in enumerate(current_row):#Просмотреть ячейки заголовка для поиска метода ГИС
            			if(type(cell)==unicode):
					cell=cell.encode('utf-8')
                                #Если название столбец совпадает с названием столбца со значениями измерений каротажа 
				if (cell==method_gis):
					ARCPY.AddMessage('Есть данные по каротажу, столбец в excel: '+str(j+1))
					number_gis=int(j)#номер столбца, в котором содержатся значения измерений ГИС
					find_list = True #столбец со значениями ГИС найден
                			break#Выйти из внутреннего цикла
                	if not find_list:#если столбец со значениями ГИС не найден
                                ARCPY.AddMessage("Нет данных по выбранному методу каротажа")
                                return def_gis_arrays #вернуть пустой массив ГИС
			continue #начать новую итерацию общего цикла
    		current_row=list_excel.row_values(rownum)#текущая строка листа
    		id_well="no_data"#идентификатор скважины
    		dep="no_data"#глубина измеренных значений
    		gis="no_data" #измеренное значение ГИС
    		for j, cell in enumerate(current_row):#перебор ячеек текущей строки листа
        		if(type(cell)==unicode):cell=cell.encode('utf-8')#если тип ячейки unicode - преобразовать в utf-8
        		if((j==0)and(cell<>'')):id_well=cell #идентификатор скважины
			if((j==1)and(cell<>'')):dep=cell #глубина измерений
			if((j==number_gis)and(j<>0)):gis=cell #измеренное значение ГИС
    		def_gis_array=[id_well,dep,gis]#массив, содержащий данные по текущему измерению
    		def_gis_arrays.append(def_gis_array)#Добавить текущее измерение в общий массив 
	return def_gis_arrays #[id скважины, глубина измерений,измеренное значение ГИС]
#функция в качестве параметра берет массив считанный с листа ГИС т.е. массив всех строк листа
#и возвращает массив в формате [id скважины, глубина/глубина/глубина..., значение/значение/значение...]
def single_gis_data(gis_arr): 
	well_first=gis_arr[0][0]#первое значение идентификации
	h=""#строка для записи значений глубин
	v=""#строка для записи значений ГИС
	#массив для записи массивов скважин с данными о глубине и каротажу [id, глубина, значение ГИС]
	def_single_gis_array=[]
	k=0
	while k<len(gis_arr):# перебор массива, считанного из excel-файла
    		if (gis_arr[k][0]==well_first):#если текущий идентификатор равен предыдущему
        		h=h+str(gis_arr[k][1])+"/" #строка, содержащая значения глубин измерений ГИС
			v=v+str(gis_arr[k][2])+"/" #строка, содержащая значения ГИС
    		else: #если курсор перешел к новой скважине
			current_array=[well_first,h,v] #запомнить данные предыдущей скважины
			def_single_gis_array.append(current_array)# и добавить их к общему массиву скважин
			well_first=gis_arr[k][0]#установить текущий идентификатор
			h=str(gis_arr[k][1])+"/" # начать новую строку глубин
			v=str(gis_arr[k][2])+"/" # начать новую строку значений ГИС
    		k=k+1
    		if (k==len(gis_arr)):#добавить в общий массив данный о последней скважине
        		current_array=[well_first,h,v]
			def_single_gis_array.append(current_array)
	return def_single_gis_array #вернуть массив в формате [id скважины, depth/depth..., value/value...]
#функция в качестве входных параметров берет массив точек с координатами и прочей информацией
#и массив с измеренными значениями ГИС
#возвращает объединенный массив в формате [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...]
def united_two_arrays(arr_wells, arr_gis):
	#общий массив, который формируется на основе массивом листа точек наблюдений и листа с ГИС
	#с первого листа берутся сведения о координатах, номере скважины, 
	#идентификатора, из второго - данные глубины и значений измерений ГИС
	def_united_array=[]
	if(len(arr_gis)==0):#если массив с данными ГИС пуст, вернуть массив элементов первого листа с пустыми параметрами глубины и значеней ГИС
                k=0
                while(k<len(arr_wells)):#перебор элементов первого листа (массива точек)
                        curr_array = [arr_wells[k][0],arr_wells[k][1],arr_wells[k][2],arr_wells[k][3],arr_wells[k][4], \
                        arr_wells[k][5],arr_wells[k][6],arr_wells[k][7],"no_data","no_data"]
                        def_united_array.append(curr_array) #добавить в общий массив точек
                        k+=1
                #вернуть массив точек наблюдений с пустыми значениями для атрибутов ГИС(глубина и значение измерений)
                #[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,"-","-"]
                return def_united_array
	k=0
	while(k<len(arr_wells)):#перебор элементов первого листа (массива точек)
                find_gis=False#логическая переменная определяет, найдены ли данные ГИС для текущей скважины
    		m=0
    		while(m<len(arr_gis)):#перебор элементов второго листа (массива со значениями скважинной мгниторазведки) 
        		if(arr_wells[k][0]==arr_gis[m][0]):#если идентификаторы совпали
				#значит это данные одной и той же скважины,
				#поэтому сформировать один общий массив
				curr_array = [arr_wells[k][0],arr_wells[k][1],arr_wells[k][2],arr_wells[k][3],arr_wells[k][4], \
                                arr_wells[k][5],arr_wells[k][6],arr_wells[k][7],arr_gis[m][1],arr_gis[m][2]]
            			def_united_array.append(curr_array) #добавить в общий массив точек
	    			find_gis=True#для текущей скважины найдены данные ГИС
        		m=m+1
        		#если достигнут конец цикла и при этом не найдено совпадения кодов(т.е. для текущей скважины отсутствуют данные ГИС
        		if(m==len(arr_gis)and(not find_gis)):
                                #Все равно добавить к финальному массиву данные по текущей скважине без данных о ГИС
                                curr_array = [arr_wells[k][0],arr_wells[k][1],arr_wells[k][2],arr_wells[k][3],arr_wells[k][4], \
                                arr_wells[k][5],arr_wells[k][6],arr_wells[k][7],"no_data","no_data"]
                                def_united_array.append(curr_array) #добавить в общий массив точек
    		k=k+1	
	return def_united_array #вернуть массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...]
#функция в качестве входных параметров берет массив [id точки,линия/номер,x,y,z,глубина,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...]
#и массив [код скважины, мощность перекрывающих]
#возвращает массив [id точки,линия/номер,x,y,z,глубина,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений]
def united_width_blocking(un_array, block_array):
	def_united_final_arr=[]#общий массив
	if(len(block_array)==0):#если массив с данными о мощности перекрывающих отложений пуст, вернуть массив с пустыми значениями мощности
                k=0
                while(k<len(un_array)):#перебор элементов объекдиненного массива с точками наблюдений и ГИС
                        curr_array=[un_array[k][0],un_array[k][1],un_array[k][2],un_array[k][3],un_array[k][4],\
                        un_array[k][5],un_array[k][6],un_array[k][7],un_array[k][8],un_array[k][9],"-1.0"]
                        def_united_final_arr.append(curr_array)#записать в общий массив информацию о текущей скважине
                        k+=1
                #возвращает массив с пустыми значениями по мощности перекрывающих отложений [id точки,линия/номер,x,y,z,глубина скважины,
                #объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений]
                return def_united_final_arr
	k=0
	while(k<len(un_array)):#перебор элементов массива [id точки,линия/номер,x,y,z,глубина,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...]
                find_m_bool=False#логическая переменная определяет, вычислена ли для текущей скважины мощность перекрывающих отложений
		curr_array=[]
		cod1=float(un_array[k][0])#код точки наблюдений первого массива
		m=0
		while(m<len(block_array)):#перебор элементов массива [код скважины, мощность перекрывающих]
			cod2=float(block_array[m][0])#код точки наблюдений второго массива
			if(cod1==cod2):#если найдено совпадений кодов
				#объединить информацию массивов
				curr_array=[un_array[k][0],un_array[k][1],un_array[k][2],un_array[k][3],un_array[k][4],un_array[k][5],un_array[k][6],un_array[k][7],un_array[k][8],un_array[k][9],block_array[m][1]]
				def_united_final_arr.append(curr_array)#записать в общий массив информацию о текущей скважине
				find_m_bool=True#для текущей скважины вычислена мощность перекрывающих отложений
			m+=1
			#если достигнут конец цикла и при этом не найдено совпадения кодов(т.е. для текущей скважины отсутствуют данные по мощности перекрывающих отложений)
			if(m==len(block_array))and(not find_m_bool):
                                #Все равно добавить к финальному массиву данные по текущей скважине без данных о мощности перекрывающих отложений 
                                curr_array=[un_array[k][0],un_array[k][1],un_array[k][2],un_array[k][3],un_array[k][4],un_array[k][5],un_array[k][6],un_array[k][7],un_array[k][8],un_array[k][9],"-1.0"]
                                def_united_final_arr.append(curr_array)#записать в общий массив информацию о текущей скважине без сведений о мощности перекрывающих отложений
		k+=1
	#возвращает массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений]
	return def_united_final_arr
#функция в качестве входных параметров берет массив в формате
#[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений]
#и массив в формате [id точки, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
#Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
#и на основе идентификатора точек объединяет эти два массива
def united_data_litostrat(fin_array,arr_stratolit):
        un_array_litostrat = []#общий массив для записи данных
        if(len(arr_stratolit)==0):#если массив с данными по литостратиграфии пуст, вернуть массив с пустыми значениями по литостратиграфии
                k=0
                #перебор элементов массива [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,
                #глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений]
                while(k<len(fin_array)):
                        curr_array=[fin_array[k][0],fin_array[k][1],fin_array[k][2],fin_array[k][3],fin_array[k][4],\
                        fin_array[k][5],fin_array[k][6],fin_array[k][7],fin_array[k][8],fin_array[k][9],fin_array[k][10],\
                        "","","",""]
                        un_array_litostrat.append(curr_array)
                        k+=1
                #вернуть массив с пустыми значениями литостратиграфии
                #[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,
                #мощность перекрывающих отложений, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
                #Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
                return un_array_litostrat
        k=0
        #перебор элементов массива [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,
        #глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений]
        while(k<len(fin_array)):
                find_stratolit=False#логическая переменная определяет, есть ли для текущей скважины информация по литостратиграфии пластов
                curr_array=[]
                cod1=float(fin_array[k][0])#код точки наблюдений первого массива
                m=0
                #перебор элементов массива [id точки, abs кровли/abs кровли/abs кровли,
                #abs подошвы/abs подошвы/abs подошвы, Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
                while(m<len(arr_stratolit)):
                        cod2=float(arr_stratolit[m][0])#код точки наблюдений второго массива
                        if(cod1==cod2):#если найдено совпадение кодов
                                #объединить информацию массивов
                                curr_array=[fin_array[k][0],fin_array[k][1],fin_array[k][2],fin_array[k][3],fin_array[k][4],\
                                fin_array[k][5],fin_array[k][6],fin_array[k][7],fin_array[k][8],fin_array[k][9],fin_array[k][10],\
                                arr_stratolit[m][1],arr_stratolit[m][2],arr_stratolit[m][3],arr_stratolit[m][4]]
                                un_array_litostrat.append(curr_array)
                                find_stratolit=True#для текущей скважины есть информация по литостратиграфии пластов
                        m+=1
                        #если достигнут конец цикла и при этом не найдено совпадения кодов(т.е. для текущей скважины отсутствуют данные по литостратиграфии пластов)
                        if(m==len(arr_stratolit))and(not find_stratolit):
                                #все равно добавить к формируемому массиву данные по текущей скважине без данных по литостратиграфии
                                curr_array=[fin_array[k][0],fin_array[k][1],fin_array[k][2],fin_array[k][3],fin_array[k][4],\
                                fin_array[k][5],fin_array[k][6],fin_array[k][7],fin_array[k][8],fin_array[k][9],fin_array[k][10],\
                                "","","",""]
                                un_array_litostrat.append(curr_array)   
                k+=1
        #вернуть массив в формате
        #[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,
        #мощность перекрывающих отложений, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
        #Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
        return un_array_litostrat
#функция в качестве входных параметров берет директорию выходного картографического слоя, систему координат, и итоговый массив
#в формате [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,
#мощность перекрывающих отложений, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
#Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
#и создает картографический слой с соответствующими атрибутами
def write_to_layer(lay, C_System, array_map):
	fc_path, fc_name = os.path.split(lay)#разделить директорию и имя файла
	gp=arcgisscripting.create(9.3)#создать объект геопроцессинга версии 9.3
	gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",C_System)#создать точечный слой
	gp.AddField(OutputFC,"id_well","STRING",15,15)#атрибут id скважин
	gp.AddField(OutputFC,"name_well","STRING",15,15)#атрибут названия скважин
	gp.AddField(OutputFC,"x","FLOAT",15,15)#добавить атрибут x
	gp.AddField(OutputFC,"y","FLOAT",15,15)#добавить атрибут y
	gp.AddField(OutputFC,"z","FLOAT",15,15)#добавить атрибут z
	gp.AddField(OutputFC,"depth_well","FLOAT",15,15)#добавить атрибут глубины скважины
	gp.AddField(OutputFC,"name_object","STRING","","",100)#атрибут - название объекта
	gp.AddField(OutputFC,"name_area","STRING","","",50)#атрибут - название участка
	gp.AddField(OutputFC,"depth_logging","STRING","","",1000)#атрибут значения глубин измерения ГИС
	gp.AddField(OutputFC,"value_logging","STRING","","",1000)#атрибут для значений ГИС
        gp.AddField(OutputFC,"recover","FLOAT",15,15)#добавить атрибут мощности перекрывающих отложений
        gp.AddField(OutputFC,"abs_litostrat_roof","STRING","","",1000)#атрибут для глубин кровли пластов
        gp.AddField(OutputFC,"abs_litostrat_sole","STRING","","",1000)#атрибут для глубин подошвы пластов
        gp.AddField(OutputFC,"stratigraphy","STRING","","",1000)#атрибут для стратиграфии
        gp.AddField(OutputFC,"lithology","STRING","","",1000)#атрибут для литологии
	rows = gp.InsertCursor(OutputFC)
	pnt = gp.CreateObject("Point")
	k=0
	while(k<len(array_map)): #создание точечного слоя
    		id_well = float(array_map[k][0])#id скважин
    		name_well = str(array_map[k][1])#названия скважин
    		z = round(float(array_map[k][4]),3)#z
                depth_well = float(array_map[k][5])#глубина скважины
    		name_obj = str(array_map[k][6])#название объекта
    		name_uchastok = str(array_map[k][7])#название участка
    		depth_logging = str(array_map[k][8])#значения глубин измерения ГИС
    		value_logging = str(array_map[k][9])#значения ГИС
    		recover = float(array_map[k][10])#атрибут мощности перекрывающих отложений
    		abs_litostrat_roof = str(array_map[k][11])#глубины кровли пластов
                abs_litostrat_sole = str(array_map[k][12])#глубины подошвы пластов
    		stratigraphy = str(array_map[k][13])#стратиграфия
    		lithology = str(array_map[k][14])#литология
    		feat = rows.NewRow()#Новая строка
    		pnt.id = k #ID
    		x = round(float(array_map[k][3]),3)
    		y = round(float(array_map[k][2]),3)
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
    		feat.SetValue("depth_logging",depth_logging)
    		feat.SetValue("value_logging",value_logging)
    		feat.SetValue("recover",recover)
    		feat.SetValue("abs_litostrat_roof",abs_litostrat_roof)
    		feat.SetValue("abs_litostrat_sole",abs_litostrat_sole)
    		feat.SetValue("stratigraphy",stratigraphy)
    		feat.SetValue("lithology",lithology)
    		rows.InsertRow(feat)#добавить строку к таблице слоя
    		k=k+1
if __name__== '__main__':#если модуль запущен, а не импортирован
        #-------------------------------входные данные------------------------------------
        InputExcelFile = ARCPY.GetParameterAsText(0)#входной excel-файл
        List_point = ARCPY.GetParameterAsText(1)#Название листа с точками наблюдений
        List_strat_lit = ARCPY.GetParameterAsText(2)#Название листа с описанием стратиграфии и литологии
        List_code_strat = ARCPY.GetParameterAsText(3)#Название листа с кодами стратиграфии
        List_code_lit = ARCPY.GetParameterAsText(4)#Название листа с кодами литологии
        List_GIS = ARCPY.GetParameterAsText(5)#Название листа с ГИС
        code_method_gis = ARCPY.GetParameterAsText(6).encode('utf-8')#Название столбца с данными каротажа для метода ГИС
        code_contain = ARCPY.GetParameterAsText(7).encode('utf-8')#Индексы возрастов кимберлитовмещающих отложений
        Coordinate_System = ARCPY.GetParameterAsText(8)#координатная система выходного точечного слоя    
        OutputFC = ARCPY.GetParameterAsText(9)#выходной точечный слой скважин
        Full_report = ARCPY.GetParameterAsText(10)#параметр определяет выводить ли полный отчет
        #---------------------------------------------------------------------------------
        #получить объекты excel-листов
        obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_GIS = \
        read_excel_list(InputExcelFile,List_point,List_strat_lit,List_code_strat,List_code_lit,List_GIS)
        #проверить все ли excel-листы с именами, указанными как входные параметры, найдены
        verify_lists(obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_GIS)
        arr_country_rock = code_contain.split(";")#получить массив индексов возрастов кимберлитовмещающих отложений
        wells_array = read_points_wells(obj_point)#получить массив точек из листа excеl-файла с точками наблюдений
        if(len(wells_array)==0):#если массив точек наблюдений пуст - завершить работу программы
                ARCPY.AddMessage("Лист точек наблюдений пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        ARCPY.AddMessage("Общее количество точек наблюдений, прочитанных из Excel-файла: "+str(len(wells_array)))
        if(Full_report == "true"):#выводить полный отчет
                if(len(wells_array)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(wells_array[k][0])+" "+str(wells_array[k][1])+" "+ \
                                str(wells_array[k][2])+" "+str(wells_array[k][3])+" "+str(wells_array[k][4])+" "+ \
                                str(wells_array[k][5])+" "+str(wells_array[k][6])+" "+str(wells_array[k][7]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_strat = read_strat_wells(obj_code_strat)#получить массив с кодами стратиграфии
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
        code_strat_new = calc_type_rock(code_strat, arr_country_rock)#получить массив с кодами стратиграфии и типами пород(вмещающими/перекрывающими)
        ARCPY.AddMessage("Количество кодов стратиграфии с типами пород: "+str(len(code_strat_new)))
        if(Full_report == "true"):#выводить полный отчет
                if(len(code_strat_new)>10):
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_strat_new[k][0])+" "+str(code_strat_new[k][1])+" "+code_strat_new[k][2])
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_lit = read_lit_wells(obj_code_lit)#получить массив с кодами литологии
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
        if(len(arr_formation)==0):#если массив литологии пластов пуст - завершить работу программы
                arr_single_width = []#создать пустой массив с сгруппированными пластами для каждой скважины
                ARCPY.AddMessage("Информация о литологии отсутствует")
        else:#если массив литологии не пуст
                arr_single_width = single_width_blocking(arr_formation) #получить массив с сгруппированными пластами для каждой скважины
        ARCPY.AddMessage("Количество скважин с информацией о вскрытых пластах: "+str(len(arr_single_width)))        
        if(len(arr_single_width)==0):#если массив с сгруппированными пластами для каждой скважины пуст - завершить работу программы
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
                arr_width_blocking=[]#создать пустой массив с мощностью перекрывающих отложений
                ARCPY.AddMessage("Нет информации по стратиграфии и литологии")
        else:#если есть скважины со стратолитологическим описанием
                if(Full_report == "true"):#выводить полный отчет
                        ARCPY.AddMessage("В формате:")
                        ARCPY.AddMessage(str(arr_view_strat_lit[0][0])+" "+str(arr_view_strat_lit[0][1])+" "+str(arr_view_strat_lit[0][2])+\
                        " "+str(arr_view_strat_lit[0][3])+" "+str(arr_view_strat_lit[0][4]))
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")     
                arr_width_blocking=calc_width_blocking(arr_single_width, code_strat_new)#вычислить мощность перекрывающих отложений
        ARCPY.AddMessage("Количество скважин с вычисленной мощностью перекрывающих отложений: "+str(len(arr_width_blocking)))
        if(len(arr_width_blocking)==0):#если нет скважин с вычисленной мощностью перекрывающих отложений
                ARCPY.AddMessage("Массив скважин с вычисленной мощностью перекрывающих отложений пуст")
        else:#если есть скважины с вычисленной мощностью перекрывающих отложений
                if(Full_report == "true"):#выводить полный отчет
                        ARCPY.AddMessage("В формате:")
                        k=0
                        while(k<len(arr_width_blocking)):
                                ARCPY.AddMessage(str(arr_width_blocking[k][0])+" "+str(arr_width_blocking[k][1]))
                                k+=1
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        gis_arrays = read_gis_wells(obj_GIS,code_method_gis)#получить массив измеренных значений каротажа из листа ГИС Excel-файла
        if(len(gis_arrays)==0):#если нет данных по каротажу выбранным методом
                single_gis_array=[]#создать пустой массив в формате [id точки,глубина/глубина..,значение/значение..]
                ARCPY.AddMessage("Нет данных по каротажу для выбранного метода")
        else:#если есть данные по каротажу выбранным методом
                if(Full_report == "true"):#выводить полный отчет
                        if(len(gis_arrays)>10):
                                ARCPY.AddMessage("В формате:")
                                k=0
                                while(k<10):
                                        ARCPY.AddMessage(str(gis_arrays[k][0])+" "+str(gis_arrays[k][1])+" "+str(gis_arrays[k][2]))
                                        k+=1
                                ARCPY.AddMessage("...........")
                                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
                single_gis_array = single_gis_data(gis_arrays)#получить массив в формате [id точки,глубина/глубина..,значение/значение..]
        ARCPY.AddMessage("Количество скважин со значениями ГИС: "+ str(len(single_gis_array)))
        if(len(single_gis_array)==0):#если массив со значениями ГИС пуст
                ARCPY.AddMessage("Массив со значениями ГИС не сформирован")
        else:
                if(Full_report == "true"):#выводить полный отчет
                        ARCPY.AddMessage("В формате:")
                        ARCPY.AddMessage(str(single_gis_array[0][0])+" "+str(single_gis_array[0][1])+" "+str(single_gis_array[0][2]))
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить объедененный массив со значениями ГИС, координатами, глубинами и т.п.
        united_array = united_two_arrays(wells_array, single_gis_array)
        ARCPY.AddMessage("Количество элементов объединенного массива точек наблюдений и каротажа: "+str(len(united_array)))
        if(len(united_array)==0):#если объединенный массив пуст
                ARCPY.AddMessage("Массив данных скважин и ГИС пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        else:#если объединенный массив не пуст
                if(Full_report == "true"):#выводить полный отчет
                        ARCPY.AddMessage("В формате:")
                        ARCPY.AddMessage(str(united_array[0][0])+" "+str(united_array[0][1])+" "+str(united_array[0][2])+ \
                        " "+str(united_array[0][3])+" "+str(united_array[0][4])+" "+str(united_array[0][5])+ \
                        " "+str(united_array[0][6])+" "+str(united_array[0][7])+" "+str(united_array[0][8])+" "+str(united_array[0][9]))
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить объединенный массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений]
        final_array=united_width_blocking(united_array, arr_width_blocking)
        ARCPY.AddMessage("Количество элементов объединенного массива с мощностью перекрывающих отложений: "+str(len(final_array)))
        if(len(final_array)==0):#если массив пуст
                ARCPY.AddMessage("Массив с мощностью перекрывающих отложений пуст, программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате:")
                ARCPY.AddMessage(str(final_array[0][0])+" "+str(final_array[0][1])+" "+str(final_array[0][2])+ \
                " "+str(final_array[0][3])+" "+str(final_array[0][4])+" "+str(final_array[0][5])+" "+ \
                str(final_array[0][6])+" "+str(final_array[0][7])+" "+str(final_array[0][8])+" "+str(final_array[0][9])+" "+str(final_array[0][10]))
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить объединенный массив, на основе которого создается картографический слой 
        #[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,мощность перекрывающих отложений,
        #глубина кровли пласта/глубина кровли пласта/глубина кровли пласта,Jmh/Jsn/Ol/,алевролит/известняк/доломит/]
        map_array=united_data_litostrat(final_array,arr_view_strat_lit)
        ARCPY.AddMessage("Количество элементов объединенного массива для создания картографического слоя: "+str(len(map_array)))
        if(len(map_array)==0):#если массив для вывода в картографический слой пуст
                ARCPY.AddMessage("Массив для вывода в картографический слой пуст, - программа остановлена")
                raise SystemExit(1)#завершить работу программы
        if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате:")
                ARCPY.AddMessage("id скважины: "+str(map_array[0][0]))
                ARCPY.AddMessage("линия/номер: "+str(map_array[0][1]))
                ARCPY.AddMessage("x: "+str(map_array[0][2]))
                ARCPY.AddMessage("y: "+str(map_array[0][3]))
                ARCPY.AddMessage("z: "+str(map_array[0][4]))
                ARCPY.AddMessage("глубина скважины: "+str(map_array[0][5]))
                ARCPY.AddMessage("название объекта: "+str(map_array[0][6]))
                ARCPY.AddMessage("название участка: "+str(map_array[0][7]))
                ARCPY.AddMessage("глубина каротажных измерений: "+str(map_array[0][8]))
                ARCPY.AddMessage("значения каротажных измерений: "+str(map_array[0][9]))
                ARCPY.AddMessage("мощность перекрывающих отложений: "+str(map_array[0][10]))
                ARCPY.AddMessage("глубины залегания кровли пластов: "+str(map_array[0][11]))
                ARCPY.AddMessage("глубины залегания подошвы пластов: "+str(map_array[0][12]))
                ARCPY.AddMessage("стратиграфия пластов: "+str(map_array[0][13]))
                ARCPY.AddMessage("литология пластов: "+str(map_array[0][14]))
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #вывести полученный массив в картографический слой
        write_to_layer(OutputFC, Coordinate_System, map_array)
        del wells_array,code_strat,code_strat_new,code_lit,arr_formation,arr_single_width,arr_width_blocking, \
            gis_arrays,single_gis_array,united_array,final_array
        
