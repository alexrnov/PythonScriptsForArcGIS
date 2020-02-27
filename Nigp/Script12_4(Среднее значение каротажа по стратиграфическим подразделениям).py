#Скрипт в качесвтве входных параметров берет excel-файл импортированный из ИСИХОГИ и содержащий данные
#по точкам наблюдений, извлекает из него эти данные и записывает их в картографический слой в формате:
#[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,код статуса документирования,тип документации,
#глубина/глубина...,значение ГИС/значение ГИС...,
#abs подразделения, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы
#Jmh/Jsn/Ol/, алевролит/известняк/доломит/,среднее_значение_ГИС,количество_измерений,среднеквадратическое_отклонение,
#ошибка_среднего,медиана]
import arcpy as ARCPY
import arcgisscripting, ConversionUtils, os
from numpy import median
try:
    from xlrd import open_workbook #модуль для работы с Excel-файлами
except ImportError: #если модуль xlrd для работы с Excel-файлами не найден
    ARCPY.AddMessage("Модуль xlrd для работы с Excel-файлами не найден, программа остановлена")
    raise SystemExit(1)#завершить работу программы
#функция в качестве входного параметра берет excel-файл, и возвращает список листов excel-файла
def read_excel_list(excel_file,List_point,List_strat_lit,List_code_strat,List_code_lit,List_GIS,List_code_type_TN,List_code_status_document,List_code_type_document):
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
    L_code_type_TN = get_object_list(name_list_array,List_code_type_TN)#лист с кодами типов ТН
    L_code_status_document = get_object_list(name_list_array,List_code_status_document)#лист с кодами состояния документирования
    L_code_type_document = get_object_list(name_list_array,List_code_type_document)#лист с кодами типов документации
    return L_point_obj, L_strat_lit_obj, L_code_strat_obj, L_code_lit_obj, L_GIS_obj,\
    L_code_type_TN, L_code_status_document,L_code_type_document#вернуть обекты excel-листов
#Функция проверяет, все ли объекты excel-листов созданы, если нет - завершить работу программы
def verify_lists(d_point, d_strat_lit, d_code_strat, d_code_lit, d_GIS, d_type_TN, d_document, d_type_document):
    if((d_point=="null")or(d_strat_lit=="null")or(d_code_strat=="null")or(d_code_lit=="null")or \
    (d_GIS=="null")or(d_type_TN=="null")or(d_document=="null")or(d_type_document=="null")): #проверить все ли объекты excel-листов созданы
        return False
    else:return True
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
        name_object = "нет данных"#название объекта
        name_area = "нет данных"#название участка
        id_well="0" #идентификатор скважины
        depth_well="-1"#глубина скважины 
        x="0"#координата x
        y="0"#координата y
        z="-1"#высотная отметка
        name_well="no_data" #линия/номер скважины
        type_well = "данные отсутствуют"#код типа ТН
        status_doc = "данные отсутствуют"#код статуса документирования
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
            if((j==25)and(cell<>'')):name_well=cell
            if((j==26)and(cell<>'')):type_doc = cell
        def_wells_array.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,str(type_well), \
        str(status_doc),str(type_doc)])#добавить текущую скважину к списку скважин
    return def_wells_array #вернуть массив точек наблюдений в формате
    #[id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок,код типа ТН,
    #код статуса документирования,тип документации]
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
#код статуса документирования,код типа документации]
#и заменяет последние четыре атрибута с кодами на их расшифровки 
def replacement_code_lists(wells_array,def_type_tn,def_status_document,def_type_document):
    def replacement_element(code_interpretation,i):#меняет коды на расшифровку для типов ТН, состояния документирования, типов документирования
        m=0
        while(m<len(code_interpretation)):#перебор кодов типов точек наблюдений
            #если код типа пробы из массива совпал с кодом из листа с кодами типов 
            if(str(wells_array[k][i])==str(code_interpretation[m][0])):
                wells_array[k][i]=code_interpretation[m][1]#поменять код на расшифровку
                break#выйти из внутреннего цикла
            m+=1
    k=0
    while(k<len(wells_array)):#перебор элементов массива точек (из листа Точки наблюдений)
        replacement_element(def_type_tn,8)
        replacement_element(def_status_document,9)
        replacement_element(def_type_document,10)
        k+=1
    #вернуть массив в формате
    #[id скважины, линия/номер скважины, x, y, z, глубина скважины, объект, участок,расшифровка типа ТН, ]
    #расшифровка статуса документирования,расшифровка типа документации]
    return wells_array
#Функция в качестве входных данных берет массив кодов стратиграфии и массив с символами возраста для искомых стратиграфических подразделений
#и возвращает массив в формате [код стратиграфии, наименование стратиграфии, код искомого подразделения(Y/N)]
def calc_type_rock(arr_strat, country):
    def_code_strat_new=[]#общий массив для записи
    k = 0
    while (k<len(arr_strat)):#перебор массива с кодами стратиграфии
        m=0
        b=False #переменная определяет были ли найден символы геологического периода, соответствующего искомому стратиграфическому подразделению
        while (m<len(country)):#перебор массива с символами для искомого стратиграфического подразделения
            if(arr_strat[k][1].count(country[m])):#если была обнаружена вмещающая порода
                #добавить в массив искомое стратиграфическое подразделение
                def_code_strat_new.append([arr_strat[k][0],arr_strat[k][1],"Y"])
                b=True #были найдены породы для искомого стратиграфического подразделения
                break#выйти из внутреннего цикла
            m+=1
        if (b==False):#если не было совпадений с искомым стратиграфическим подразделением
            #добавить в массив <<не искомое стратиграфическое подразделение>> 
            def_code_strat_new.append([arr_strat[k][0],arr_strat[k][1],"N"])
        k+=1
    #вернуть массив в формате [код стратиграфии, обозначение стратиграфии, код отложений)]
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
    def form_string(def_arr_code,def_code):#функция формирует строку с информацией по стратиграфии или литологии
        return_string = ""#строка со стратиграфией(литологией) 
        m=0
        while(m<len(def_arr_code)-1):
            n=0
            #перебор массива всех кодов стратиграфий(литологий), формат [код, наименование]
            while(n<len(def_code)):
                cod1=float(def_arr_code[m])#преобразовать в число код стратиграфии(литологии) 
                cod2=float(def_code[n][0])#преобразовать в число код из входного массива def_code
                if(cod1==cod2):#если коды совпали
                    #формировать строку со стратиграфией(литологией) в формате Jmh/Jsn/Ol/
                    return_string=return_string+def_code[n][1]+"/"
                    #если коды совпали - выйти из текущего цикла(перейти к следующей стратиграфии(литологии) текущей скважины)
                    break 
                n+=1
            m+=1
        return return_string
    def_arr_text_strat_lit=[]
    k=0
    while(k<len(arr_code_strat_lit)):#перебор массива скважин формата [код точки, кровли пластов, коды стратиграфий, коды литологий]
        str_strat_code=str(arr_code_strat_lit[k][3])#получить строку кодов стратиграфий для текущей скважины
        arr_strat_code=str_strat_code.split("/")#получить массив кодов стратиграфий для текущей скважины
        str_lit_code=str(arr_code_strat_lit[k][4])#получить строку кодов литологий для текущей скважины
        arr_lit_code=str_lit_code.split("/")#получить массив кодов литологий для текущей скважины
        stroka_strat=form_string(arr_strat_code,strat_code)#строка со стратиграфией в формате Jmh/Jsn/Ol/
        stroka_lit=form_string(arr_lit_code,lit_code)#строка с литологией в формате алевролит/известняк/доломит/
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
#и массив [код стратиграфии, наименование стратиграфии, код входного возраста(Y/N)]
#и возвращает массив [код скважины, найденные интервалы abs from - abs to]
def calc_with_strat(arr_width, code_strat):
    def_arr_with_strat=[]
    k=0
    while(k<len(arr_width)):#перебор массива скважин формата [код точки, кровли пластов, коды стратиграфий, коды литологий]
        str_roof = str(arr_width[k][1])#получить строку abs кровель текущей скважины
        arr_roof = str_roof.split("/")#получить массив abs кровель для текущей скважины
        str_sole = str(arr_width[k][2])#получить строку abs подошв текущей скважины
        arr_sole = str_sole.split("/")#получить массив abs подошв для текущей скважины
        str_code = str(arr_width[k][3])#получить строку кодов стратиграфий для текущей скважины
        arr_code = str_code.split("/")#получить массив кодов стратиграфий для текущей скважины
        m=0
        str_intervals=''#строка содержит интервалы (abs from-abs to) которые совпали с искомым возрастом
        ##перебор кодов стратиграфий текущей скважины, не обрабатывать последний элемент массива кодов стратиграфий поскольку
        #он является пустым - при переводе строки '1/2/3/' получается массив ['1','2','3','']
        while(m<len(arr_code)-1):
            n=0		
            #перебор всех кодов стратиграфий массива [код стратиграфии, наименование стратиграфии, код входного возраста(Y/N)]
            while(n<len(code_strat)):
                #если код стратиграфии текущего пласта совпадает с кодом из массива типов пород(Y/N)
                cod1=float(arr_code[m])#преобразовать в число код стратиграфии из массива arr_width
                cod2=float(code_strat[n][0])#преобразовать в число код из массива code_strat_new 
                if(cod1==cod2):#если коды совпали
                    #если код является искомым возрастом
                    #добавить к строке абсолютные отметки кровли и подошвы текущего пласта
                    if(code_strat[n][2]=="Y"):str_intervals+=str(arr_roof[m])+"-"+str(arr_sole[m])+";"
                n+=1
            m+=1
            #если достигнут конец цикла и найдено совпадение искомого возраста и стратиграфических подразделений
            #добавить в выходной массив данные по текущей скважине
            if (m==len(arr_code)-1)and(str_intervals<>''):def_arr_with_strat.append([arr_width[k][0],str_intervals])
        k+=1
    #вернуть массив [код скважины, найденные интервалы]
    return def_arr_with_strat 
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
#возвращает объединенный массив в формате [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,
#код типа ТН,код статуса документирования,тип документации,глубина/глубина...,значение ГИС/значение ГИС...]
def united_two_arrays(arr_wells, arr_gis):
    #общий массив, который формируется на основе массивом листа точек наблюдений и листа с ГИС
    #с первого листа берутся сведения о координатах, номере скважины, 
    #идентификатора, из второго - данные глубины и значений измерений ГИС
    def_united_array=[]
    if(len(arr_gis)==0):#если массив с данными ГИС пуст, вернуть массив элементов первого листа с пустыми параметрами глубины и значеней ГИС
        k=0
        while(k<len(arr_wells)):#перебор элементов первого листа (массива точек)
            curr_array = [arr_wells[k][0],arr_wells[k][1],arr_wells[k][2],arr_wells[k][3],arr_wells[k][4], \
            arr_wells[k][5],arr_wells[k][6],arr_wells[k][7],arr_wells[k][8],arr_wells[k][9],\
            arr_wells[k][10],"нет данных","нет данных"]
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
                curr_array = [arr_wells[k][0],arr_wells[k][1],arr_wells[k][2],arr_wells[k][3],\
                arr_wells[k][4],arr_wells[k][5],arr_wells[k][6],arr_wells[k][7],\
                arr_wells[k][8],arr_wells[k][9],arr_wells[k][10],arr_gis[m][1],arr_gis[m][2]]
                def_united_array.append(curr_array) #добавить в общий массив точек
                find_gis=True#для текущей скважины найдены данные ГИС
            m=m+1
            #если достигнут конец цикла и при этом не найдено совпадения кодов(т.е. для текущей скважины отсутствуют данные ГИС
            if(m==len(arr_gis)and(not find_gis)):
                #Все равно добавить к финальному массиву данные по текущей скважине без данных о ГИС
                curr_array = [arr_wells[k][0],arr_wells[k][1],arr_wells[k][2],arr_wells[k][3],\
                arr_wells[k][4],arr_wells[k][5],arr_wells[k][6],arr_wells[k][7],arr_wells[k][8],\
                arr_wells[k][9],arr_wells[k][10],"нет данных","нет данных"]
                def_united_array.append(curr_array) #добавить в общий массив точек
        k=k+1	
    return def_united_array
    #вернуть массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,
    #код статуса документирования,тип документации, глубина/глубина...,значение ГИС/значение ГИС...]
#функция в качестве входных параметров берет массив [id точки,линия/номер,x,y,z,глубина,объект,участок,
#код типа ТН,код статуса документирования,тип документации,глубина/глубина...,значение ГИС/значение ГИС...]
#и массив [код скважины, "abs кровли-abs подошвы;abs кровли1-abs подошвы1;..."]
#возвращает массив [id точки,линия/номер,x,y,z,глубина,объект,участок,
#код типа ТН,код статуса документирования,тип документации,глубина/глубина...,значение ГИС/значение ГИС...,
#"abs кровли-abs подошвы;abs кровли1-abs подошвы1;..."]
def united_width_strat(un_array, strat_array):
    def_united_final_arr=[]#общий массив
    if(len(strat_array)==0):#если массив с данными о найденных стратиграфических подразделениях пуст, вернуть массив с пустыми значениями для подразделений
        k=0
        while(k<len(un_array)):#перебор элементов объекдиненного массива с точками наблюдений и ГИС
            curr_array=[un_array[k][0],un_array[k][1],un_array[k][2],un_array[k][3],un_array[k][4],\
            un_array[k][5],un_array[k][6],un_array[k][7],un_array[k][8],un_array[k][9],un_array[k][10],\
            un_array[k][11],un_array[k][12],"-1.0"]
            def_united_final_arr.append(curr_array)#записать в общий массив информацию о текущей скважине
            k+=1
        #возвращает массив с пустыми значениями по подразделению[id точки,линия/номер,x,y,z,глубина скважины,
        #объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,"-1.0"]
        return def_united_final_arr
    k=0
    while(k<len(un_array)):#перебор элементов массива [id точки,линия/номер,x,y,z,глубина,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...]
        find_m_bool=False#логическая переменная определяет, вычислены ли для текущей скважины данные по подразделению
        curr_array=[]
        cod1=float(un_array[k][0])#код точки наблюдений первого массива
        m=0
        while(m<len(strat_array)):#перебор элементов массива [код скважины, "интервалы подразделений"]
            cod2=float(strat_array[m][0])#код точки наблюдений второго массива
            if(cod1==cod2):#если найдено совпадений кодов
                #объединить информацию массивов
                curr_array=[un_array[k][0],un_array[k][1],un_array[k][2],un_array[k][3],un_array[k][4],\
                un_array[k][5],un_array[k][6],un_array[k][7],un_array[k][8],un_array[k][9],un_array[k][10],\
                un_array[k][11],un_array[k][12],strat_array[m][1]]
                def_united_final_arr.append(curr_array)#записать в общий массив информацию о текущей скважине
                find_m_bool=True#для текущей скважины вычислены параметры стратиграфических подразделений
            m+=1
            #если достигнут конец цикла и при этом не найдено совпадения кодов(т.е. для текущей скважины отсутствуют данные по искомому стратиграфическому подразделению)
            if(m==len(strat_array))and(not find_m_bool):
                #Все равно добавить к финальному массиву данные по текущей скважине без данных о стратиграфическом подразделении
                curr_array=[un_array[k][0],un_array[k][1],un_array[k][2],un_array[k][3],un_array[k][4],\
                un_array[k][5],un_array[k][6],un_array[k][7],un_array[k][8],un_array[k][9],un_array[k][10],\
                un_array[k][11],un_array[k][12],"-1.0"]
                def_united_final_arr.append(curr_array)#записать в общий массив информацию о текущей скважине без сведений о искомом стратиграфическом подразделении
        k+=1
    #возвращает массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,код статуса документирования,тип документации,
    #глубина/глубина...,значение ГИС/значение ГИС...,abs кровли-abs подошвы]
    return def_united_final_arr
#функция в качестве входных параметров берет итоговый массив для вывода в картографический слой, массив в формате
#[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,код статуса документирования,тип документации,
#глубина/глубина...,значение ГИС/значение ГИС...,abs кровли-abs подошвы]
#и массив в формате [id точки, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
#Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
#и на основе идентификатора точек объединяет эти два массива
def united_data_litostrat(un_array_litostrat,fin_array,arr_stratolit):
    if(len(arr_stratolit)==0):#если массив с данными по литостратиграфии пуст, вернуть массив с пустыми значениями по литостратиграфии
        k=0
        #перебор элементов массива [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,
        #код типа ТН,код статуса документирования,тип документации,
        #глубина/глубина...,значение ГИС/значение ГИС...,abs подразделения]
        while(k<len(fin_array)):
            curr_array=[fin_array[k][0],fin_array[k][1],fin_array[k][2],fin_array[k][3],fin_array[k][4],\
            fin_array[k][5],fin_array[k][6],fin_array[k][7],fin_array[k][8],fin_array[k][9],fin_array[k][10],\
            fin_array[k][11],fin_array[k][12],fin_array[k][13],"нет данных","нет данных","нет данных","нет данных"]
            un_array_litostrat.append(curr_array)
            k+=1
        #вернуть массив с пустыми значениями литостратиграфии
        #[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,глубина/глубина...,значение ГИС/значение ГИС...,
        #abs подразделения, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
        #Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
        return un_array_litostrat
    k=0
    #перебор элементов массива [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,
    #глубина/глубина...,значение ГИС/значение ГИС...,abs подразделения]
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
                fin_array[k][11],fin_array[k][12],fin_array[k][13],\
                arr_stratolit[m][1],arr_stratolit[m][2],arr_stratolit[m][3],arr_stratolit[m][4]]
                un_array_litostrat.append(curr_array)
                find_stratolit=True#для текущей скважины есть информация по литостратиграфии пластов
            m+=1
            #если достигнут конец цикла и при этом не найдено совпадения кодов(т.е. для текущей скважины отсутствуют данные по литостратиграфии пластов)
            if(m==len(arr_stratolit))and(not find_stratolit):
                #все равно добавить к формируемому массиву данные по текущей скважине без данных по литостратиграфии
                curr_array=[fin_array[k][0],fin_array[k][1],fin_array[k][2],fin_array[k][3],fin_array[k][4],\
                fin_array[k][5],fin_array[k][6],fin_array[k][7],fin_array[k][8],fin_array[k][9],fin_array[k][10],\
                fin_array[k][11],fin_array[k][12],fin_array[k][13],"нет данных","нет данных","нет данных","нет данных"]
                un_array_litostrat.append(curr_array)   
        k+=1
    #вернуть массив в формате
    #[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,код статуса документирования,тип документации,
    #глубина/глубина...,значение ГИС/значение ГИС...,
    #abs подразделения, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
    #Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
    return un_array_litostrat
#функция в качестве входных параметров берет директорию выходного картографического слоя, систему координат, и итоговый массив
#в формате
#[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,код статуса документирования,тип документации,
#глубина/глубина...,значение ГИС/значение ГИС...,
#abs подразделения, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
#Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
#и создает картографический слой с соответствующими атрибутами
def write_to_layer(lay, C_System, array_map, def_recovery):
    fc_path, fc_name = os.path.split(lay)#разделить директорию и имя файла
    gp=arcgisscripting.create(9.3)#создать объект геопроцессинга версии 9.3
    gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",C_System)#создать точечный слой
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
    x_rec = 0.0
    y_rec = 0.0
    if def_recovery == "true":#если необходимо учитывать поправку БГРЭ по координатам
        x_rec = 20000.0
        y_rec = 10000.0
    rows = gp.InsertCursor(lay)
    pnt = gp.CreateObject("Point")
    k=0
    while(k<len(array_map)): #создание точечного слоя
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
        feat = rows.NewRow()#Новая строка
        pnt.id = k #ID
        x = round(float(array_map[k][3])-x_rec,3)
        y = round(float(array_map[k][2])-y_rec,3)
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
        rows.InsertRow(feat)#добавить строку к таблице слоя
        k=k+1
#функция вычисляет среднее арефмитическое и медиану по заданным интервалам
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
            current_point[16],current_point[17],average_gis,depth_interval,msd,standard_error,round(median(arr_value_contain),2)])
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
def write_layer_average(lay,array_map,spatial_ref,def_recovery):
    fc_path,fc_name = os.path.split(lay)
    gp=arcgisscripting.create(9.3)#создать объект геопроцессинга версии 9.3
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
    x_rec = 0.0
    y_rec = 0.0
    if def_recovery == "true":#если необходимо учитывать поправку БГРЭ по координатам
        x_rec = 20000.0
        y_rec = 10000.0
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
        x = round(float(array_map[k][3])-x_rec,3)
        y = round(float(array_map[k][2])-y_rec,3)
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
    #-------------------------------входные данные------------------------------------
    InputExcelFile = ARCPY.GetParameterAsText(0)#входной набор excel-файлов
    code_method_gis = ARCPY.GetParameterAsText(1).encode('utf-8')#Название столбца с данными каротажа для метода ГИС
    code_contain = ARCPY.GetParameterAsText(2).encode('utf-8')#Индекс возраста(стратиграфии)
    Coordinate_System = ARCPY.GetParameterAsText(3)#координатная система выходного точечного слоя    
    output_FC = ConversionUtils.gp.GetParameterAsText(4)#выходной картографический слой с данными ГИС и интервалами
    OutputFC_average = ConversionUtils.gp.GetParameterAsText(5)#выходной картографический слой с вычисленными средними значениями по ГИС
    Full_report = ARCPY.GetParameterAsText(6)#параметр определяет выводить ли полный отчет
    recovery = ARCPY.GetParameterAsText(7)#параметр определяет учитывать ли поправку к системе координат в БГРЭ
    List_point = ARCPY.GetParameterAsText(8)#Название листа с точками наблюдений
    List_strat_lit = ARCPY.GetParameterAsText(9)#Название листа с описанием стратиграфии и литологии
    List_code_strat = ARCPY.GetParameterAsText(10)#Название листа с кодами стратиграфии
    List_code_lit = ARCPY.GetParameterAsText(11)#Название листа с кодами литологии
    List_GIS = ARCPY.GetParameterAsText(12)#Название листа с ГИС
    List_code_type_TN = ARCPY.GetParameterAsText(13)#Название листа с кодами типов ТН
    List_code_status_document = ARCPY.GetParameterAsText(14)#Название листа с кодами состояния документирования
    List_code_type_document = ARCPY.GetParameterAsText(15)#Название листа с кодами типов документирования  
    #---------------------------------------------------------------------------------
    excel_list_files = ConversionUtils.SplitMultiInputs(InputExcelFile)#получить массив excel-файлов
    ARCPY.AddMessage("-----------------------------------------------------------------------------------")
    ARCPY.AddMessage("Количество Excel-файлов: "+str(len(excel_list_files))+" шт.")
    map_array=[]#массив для вывода всех ТН в выходной картографический слой
    for input in excel_list_files:#перебор excel-файлов
        #outdata = ConversionUtils.GenerateOutputName(input, output_workspace)#путь выходного картографического слоя
        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        ARCPY.AddMessage("Excel-файл: "+str(input.encode('utf-8')))
        arr_country_rock = code_contain.split(";")#получить массив индексов искомых возрастов 
        #получить объекты excel-листов
        obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_GIS, obj_code_type_TN, \
        obj_code_status_document,obj_code_type_document = \
        read_excel_list(input,List_point,List_strat_lit,List_code_strat,List_code_lit,List_GIS,\
        List_code_type_TN,List_code_status_document,List_code_type_document)
        #проверить все ли excel-листы с именами, указанными как входные параметры, найдены
        ver_lists = verify_lists(obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_GIS, obj_code_type_TN,\
        obj_code_status_document,obj_code_type_document)
        if not ver_lists:
            ARCPY.AddMessage("Прочитаны не все excel-листы")#завершить работу программы
            continue#перейти к обработке следующего excel-файла
        else:
            ARCPY.AddMessage("Все листы excel-листа найдены")
        wells_array = read_points_wells(obj_point)#получить массив точек из листа excеl-файла с точками наблюдений
        if(len(wells_array)==0):#если массив точек наблюдений пуст - завершить работу программы
            ARCPY.AddMessage("Лист точек наблюдений текущего excel-файла пуст")
            continue#перейти к обработке следующего excel-файла
        ARCPY.AddMessage("Общее количество точек наблюдений, прочитанных из Excel-файла: "+str(len(wells_array)))
        if(Full_report == "true"):#выводить полный отчет
            if(len(wells_array)>10):
                ARCPY.AddMessage("В формате:")
                k=0
                while(k<10):
                    ARCPY.AddMessage(str(wells_array[k][0])+" "+str(wells_array[k][1])+" "+ \
                    str(wells_array[k][2])+" "+str(wells_array[k][3])+" "+str(wells_array[k][4])+" "+ \
                    str(wells_array[k][5])+" "+str(wells_array[k][6])+" "+str(wells_array[k][7])+" "+ \
                    str(wells_array[k][8])+" "+str(wells_array[k][9])+" "+str(wells_array[k][10]))
                    k+=1
                ARCPY.AddMessage("...........")
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_tn = read_code_wells(obj_code_type_TN)#получить массив с кодами типов точек наблюдений
        ARCPY.AddMessage("Количество кодов с типами точек наблюдения: "+str(len(code_type_tn)))
        if(len(code_type_tn)==0): #если массив кодов типов ТН пуст - завершить работу программы
            ARCPY.AddMessage("Лист с кодами типов ТН пуст для текущего excel-файла")
            continue#перейти к обработке следующего excel-файла
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
            ARCPY.AddMessage("Лист с кодами состояния документирования пуст для текущего excel-файла")
            continue#перейти к обработке следующего excel-файла
        if(Full_report == "true"):
            ARCPY.AddMessage("В формате:")
            k=0
            while(k<len(code_status_document)):
                ARCPY.AddMessage(str(code_status_document[k][0])+" "+str(code_status_document[k][1]))
                k+=1
            ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_document = read_code_wells(obj_code_type_document)#получить массив с кодами типов документирования
        ARCPY.AddMessage("Количество кодов типов документирования: "+str(len(code_type_document)))
        if(len(code_type_document)==0):#если массив кодов типов документирования пуст - завершить работу программы
            ARCPY.AddMessage("Лист с кодами типов документирования пуст для текущего excel-файла")
            continue#перейти к обработке следующего excel-файла
        if(Full_report == "true"):
            ARCPY.AddMessage("В формате:")
            k=0
            while(k<len(code_type_document)):
                ARCPY.AddMessage(str(code_type_document[k][0])+" "+str(code_type_document[k][1]))
                k+=1
            ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить массив - копию массива ТН, где заменены четыре поля с кодами на их расшифровки
        wells_array = replacement_code_lists(wells_array,code_type_tn,code_status_document,code_type_document)
        ARCPY.AddMessage("Общее количество точек наблюдений, с расшифрованными кодами общей информации по скважинам: "+ \
        str(len(wells_array)))
        if(Full_report == "true"):#выводить полный отчет
            if(len(wells_array)>10):
                ARCPY.AddMessage("В формате:")
                k=0
                while(k<10):
                    ARCPY.AddMessage(str(wells_array[k][0])+" "+str(wells_array[k][1])+" "+ \
                    str(wells_array[k][2])+" "+str(wells_array[k][3])+" "+str(wells_array[k][4])+" "+ \
                    str(wells_array[k][5])+" "+str(wells_array[k][6])+" "+str(wells_array[k][7])+" "+ \
                    str(wells_array[k][8])+" "+str(wells_array[k][9])+" "+str(wells_array[k][10]))
                    k+=1
                ARCPY.AddMessage("...........")
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_strat = read_code_wells(obj_code_strat)#получить массив с кодами стратиграфии
        if(len(code_strat)==0):#если массив кодов стратиграфии пуст - завершить работу программы
            ARCPY.AddMessage("Лист кодов стратиграфии пуст для текущего excel-файла")
            continue#перейти к обработке следующего excel-файла
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
        #получить массив с кодами стратиграфии и типами пород(стратиграфия совпала/не совпала)
        code_strat_new = calc_type_rock(code_strat, arr_country_rock)
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
            ARCPY.AddMessage("Лист кодов литологии пуст для текущего excel-файла")
            continue#перейти к обработке следующего excel-файла
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
            arr_with_strat=[]#с искомым стратиграфическим подразделением(и)
            ARCPY.AddMessage("Нет информации по стратиграфии и литологии")
        else:#если есть скважины со стратолитологическим описанием
            if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате:")
                ARCPY.AddMessage(str(arr_view_strat_lit[0][0])+" "+str(arr_view_strat_lit[0][1])+" "+str(arr_view_strat_lit[0][2])+\
                " "+str(arr_view_strat_lit[0][3])+" "+str(arr_view_strat_lit[0][4]))
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")     
            #вычислить abs кровли и подошвы стратиграфического подразделения(подразделений)
            arr_with_strat=calc_with_strat(arr_single_width, code_strat_new)
        ARCPY.AddMessage("Количество скважин с стратиграфическим подразделением(и): "+str(len(arr_with_strat)))
        if(len(arr_with_strat)==0):#с искомым стратиграфическим подразделением(и)
            ARCPY.AddMessage("Массив скважин с стратиграфическим подразделением(и) пуст")
        else:#если есть скважины с искомым стратиграфическим подразделением(и)
            if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате: id скважины | найденные искомые стратиграфические пласты")
                k=0
                while(k<len(arr_with_strat)):
                    ARCPY.AddMessage(str(arr_with_strat[k][0])+" "+str(arr_with_strat[k][1]))
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
            ARCPY.AddMessage("Массив данных скважин и ГИС пуст для текущего excel-файла")
            continue#перейти к обработке следующего excel-файла
        else:#если объединенный массив не пуст
            if(Full_report == "true"):#выводить полный отчет
                ARCPY.AddMessage("В формате:")
                ARCPY.AddMessage(str(united_array[0][0])+" "+str(united_array[0][1])+" "+str(united_array[0][2])+ \
                " "+str(united_array[0][3])+" "+str(united_array[0][4])+" "+str(united_array[0][5])+ \
                " "+str(united_array[0][6])+" "+str(united_array[0][7])+" "+str(united_array[0][8])+ \
                " "+str(united_array[0][9])+" "+str(united_array[0][10])+" "+str(united_array[0][11])+\
                " "+str(united_array[0][12]))
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить объединенный массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,код статуса документирования,тип документации,
        #глубина/глубина...,значение ГИС/значение ГИС...,abs подразделения]
        final_array=united_width_strat(united_array, arr_with_strat)
        ARCPY.AddMessage("Количество элементов объединенного массива с данными по найденным стратиграфическим подразделениям: "+str(len(final_array)))
        if(len(final_array)==0):#если массив пуст
            ARCPY.AddMessage("Объединенный массив с данными по найденным стратиграфическим подразделениям пуст")
            continue#перейти к обработке следующего excel-файла
        if(Full_report == "true"):#выводить полный отчет
            ARCPY.AddMessage("В формате:")
            ARCPY.AddMessage(str(final_array[0][0])+" "+str(final_array[0][1])+" "+str(final_array[0][2])+ \
            " "+str(final_array[0][3])+" "+str(final_array[0][4])+" "+str(final_array[0][5])+" "+ \
            str(final_array[0][6])+" "+str(final_array[0][7])+" "+str(final_array[0][8])+" "+str(final_array[0][9])+ \
            " "+str(final_array[0][10])+" "+str(final_array[0][11])+" "+str(final_array[0][12])+" "+ \
            str(final_array[0][13]))
            ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #получить объединенный массив, на оснARCPY.AddMessage("Количество кодов стратиграфии: "+str(len(code_strat)))ове которого создается картографический слой 
        #[id точки,линия/номер,x,y,z,глубина скважины,объект,участок,код типа ТН,код статуса документирования,тип документации,
        #глубина/глубина...,значение ГИС/значение ГИС...,
        #abs подразделения, abs кровли/abs кровли/abs кровли, abs подошвы/abs подошвы/abs подошвы,
        #Jmh/Jsn/Ol/, алевролит/известняк/доломит/]
        map_array=united_data_litostrat(map_array,final_array,arr_view_strat_lit)
        ARCPY.AddMessage("Количество элементов объединенного массива для создания картографического слоя: "+str(len(map_array)))
        if(len(map_array)==0):#если массив для вывода в картографический слой пуст
            ARCPY.AddMessage("Массив для вывода в картографический слой пуст")
            continue#перейти к обработке следующего excel-файла
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
            ARCPY.AddMessage("тип точки наблюдения: "+str(map_array[0][8]))
            ARCPY.AddMessage("статус документирования: "+str(map_array[0][9]))
            ARCPY.AddMessage("тип документации: "+str(map_array[0][10]))
            ARCPY.AddMessage("глубина каротажных измерений: "+str(map_array[0][11]))
            ARCPY.AddMessage("значения каротажных измерений: "+str(map_array[0][12]))
            ARCPY.AddMessage("abs кровли и подошвы искомого стратиграфического подразделения: "+str(map_array[0][13]))            
            ARCPY.AddMessage("глубины залегания кровли пластов: "+str(map_array[0][14]))
            ARCPY.AddMessage("глубины залегания подошвы пластов: "+str(map_array[0][15]))
            ARCPY.AddMessage("стратиграфия пластов: "+str(map_array[0][16]))
            ARCPY.AddMessage("литология пластов: "+str(map_array[0][17]))
            ARCPY.AddMessage("-----------------------------------------------------------------------------------")
    write_to_layer(output_FC, Coordinate_System, map_array,recovery)
    arr_select=select_values(map_array)#вычислить среднее значение и медиану
    ARCPY.AddMessage("Количество ТН, для которых вычислены среднее арефметическое и медиана: "+str(len(arr_select)))
    #вывести полученный массив в картографический слой
    write_layer_average(OutputFC_average,arr_select,Coordinate_System,recovery)#записать итоговый массив к картографический слой
    del wells_array,code_strat,code_strat_new,code_lit,arr_formation,arr_single_width,arr_with_strat,\
    gis_arrays,single_gis_array,united_array,final_array
        
        
