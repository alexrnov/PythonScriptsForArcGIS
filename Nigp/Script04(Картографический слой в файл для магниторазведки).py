#скрипт в качестве входного параметра берет картографический слой с данными магнитной восприимчивости,
#размер моделируемого кимберлитового тела и его магнитную восприимчивость
#и записывает прочитанные данные в текстовый файл, который использует программа GravMag3D
#для физико-геологического моделирования для магниторазведки
import arcpy, sys, arcgisscripting, os
#функция определяет корректно ли значение размера кимберлитового тела
#s_kimb - значение размера кимберлитового тела и возвращает часть значения размера кимберлитового тела 
def true_size(s_kimb):
        s = False#логическая переменная определяет корректно ли значение размера кимберлитового тела
        int1 = 0 #размер кимберлитового тела
        if(s_kimb.count("x")):#если символ x найден
                size_kimb_arr = s_kimb.split("x")#разделить значение, например 100x100 будет [100,100]
                if(size_kimb_arr[0]==size_kimb_arr[1]):#если две части совпадают (например 100x100)
                        try:#если эти части содержат только символы цифр
                                int1 = int(size_kimb_arr[0])#преобразовать часть в число
                                if ((int1>0)and(int1<3000)):s = True#если размер больше 0 и меньше 3000                
                        except ValueError:pass
        return s, int1
#функция в качестве входного параметра берет точечный слой с точками наблюдений,
#который содержит данные по магнитной восприимчивости вчр и мощности перекрывающих отложений
#и возвращает массив в формате
# [id скважины,линия/точка,x,y,abs устья скважин,название объекта,название участка,глубины измерений магнитной восприимчивости,мощность перекрывающих отложений]
def read_points(layer):
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
        name_cell = check_unicode(row.GetValue("name"))#название скважины
        x = point.x
        y = point.y
        z = check_unicode(row.GetValue("z"))#высотная отметка устья скважин
	z = float(z)
	z = round(z,2)#округлить до двух знаков после запятой 
        name_object = check_unicode(row.GetValue("name_object"))#название объекта
	name_uchastok = check_unicode(row.GetValue("name_area"))#название участка
	depth = check_unicode(row.GetValue("depth"))#глубины измерения магнитной восприимчивости
	value_kmv = check_unicode(row.GetValue("value_kmv"))#значения измерений магнитной восприимчивости
        block_m = check_unicode(row.GetValue("m"))#мощность перекрывающих отложений
	block_m = float(block_m)
	block_m = round(block_m,1)#округлить до одного знака после запятой
	arr_wells.append([id_well,name_cell,x,y,z,name_object,name_uchastok,depth,value_kmv,block_m])
        row = rows.next()
    return arr_wells #вернуть массив [id_well,name_cell,x,y,z,name_object,name_uchastok,depth,value_kmv,block_m]
#Функция в качестве входных параметров берет название текстового файла для записи данных по точкам наблюдений
#и сам массив, содержащий информацию по точкам наблюдений
#и записывает текстовый файл
def write_txt_file(text_file,array_points,size_k,mv_k,ac):
	if not(text_file==""):#если указан путь для текстового файла 
    		fc_path, fc_name = os.path.split(text_file)
    		if(fc_name<>"magn.txt"):#файл должен иметь название magn.txt - это необходимо для программы GravMag3D
                        gp.AddMessage("Некорректное название текстового файла, он должен иметь имя magn.txt")
                        raise SystemExit(1)#завершить работу программы
    		text_file_open = open(text_file,"w")
    		k=0
    		text_file_open.write("size_kimberlite="+str(size_k)+"\n")#размер кимберлитового тела
    		text_file_open.write("magnetizability_of_kimberlite="+str(mv_k)+"\n")#магнитная восприимчивость кимберлитового тела
    		text_file_open.write("shooting_accuracy="+str(ac)+"\n")#точность магнитометрической съемки
    		while (k<len(array_points)):#записать данные в текстовый файл
        		text_file_open.write(str(array_points[k][0])+"$"+str(array_points[k][1])+"$"+ \
                        str(array_points[k][3])+"$"+str(array_points[k][2])+"$"+str(array_points[k][4])+ \
                        "$"+str(array_points[k][5])+"$"+str(array_points[k][6])+"$"+str(array_points[k][7])+ \
                        "$"+str(array_points[k][8])+"$"+str(array_points[k][9])+"\n")
        		k=k+1
	text_file_open.close()
if __name__ == '__main__':#если модуль запущен а не импортирован
        gp = arcgisscripting.create(9.3)
        #-------------------------------входные данные------------------------------------
        InpPFC = arcpy.GetParameterAsText(0)#входной точечный слой скважин
        size_kimb = arcpy.GetParameterAsText(1)#входное значение размера кимберлитового тела
        mv_kimb = arcpy.GetParameterAsText(2)#входное значение магнитной восприимчивости кимберлитового тела
        accuracy = arcpy.GetParameterAsText(3)#Входное значение точности магнитометрической съемки
        output_text_file = gp.GetParameterAsText(4)#Выходной текстовый файл с данными магнитной восприимчивости
        #---------------------------------------------------------------------------------
        bool_size, value_size = true_size(size_kimb) 
        if (not(bool_size)or(value_size==0)):#если данные размера кимберлитовых тел некорректны
                gp.AddMessage("Некорректное значение в текстовом поле ввода 'входное значение размера кимберлитового тела'")
                raise SystemExit(1)#завершить работу программы
        try:
                float_mv_kimb = float(mv_kimb)#преобразовать в вещественный тип
                if not((float_mv_kimb>0)and(float_mv_kimb<1000)):#если значение магнитной восприимчивости больше 0 и меньше 1000 ед.СИ
                        gp.AddMessage("Некорректное значение в текстовом поле ввода 'входное значение магнитной восприимчивости кимберлитового тела'")              
                        gp.AddMessage("Программа остановлена")
                        raise SystemExit(1)#завершить работу программы
        except ValueError:#если значение содержит нечисловые символы 
                gp.AddMessage("Некорректное значение в текстовом поле ввода 'входное значение магнитной восприимчивости кимберлитового тела'")
                raise SystemExit(1)#завершить работу программы
        try:
                float_accuracy = float(accuracy)
        except ValueError:#если значение содержит нечисловые символы
                gp.AddMessage("Некоректное значение в текстовом поле ввода 'точность магнитометрической съемки'")
                gp.AddMessage("Программа остановлена")
                raise SystemExit(1)#завершить работу программы
        arr_points = read_points(InpPFC)#извлечь поля точечного слоя
        gp.AddMessage("Количество точек входного точечного слоя: "+str(len(arr_points)))
        write_txt_file(output_text_file, arr_points, value_size, mv_kimb, accuracy)#записать массив тоек наблюдений в текстовый файл
