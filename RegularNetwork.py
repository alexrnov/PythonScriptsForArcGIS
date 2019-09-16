import sys, arcgisscripting, os, math
import arcpy as ARCPY
gp = arcgisscripting.create(9.3)
#функция на входе берет точечный слой со значениями каротажа, который формируется в скрипте Script12_3
#и возвращает массив точек
def read_map_layer(layer,atr_evaluate):
    def check_unicode(text):#функция преобразует кодировку юникод в utf-8 
        if(type(text)==unicode):text = text.encode('utf-8')
        return text
    rows = gp.SearchCursor(layer)
    row = rows.Next()
    layer = gp.describe(layer)
    arr_wells=[]#массив для записывания скважин
    try:
        while row:
            feat = row.GetValue(layer.ShapeFieldName)
            point = feat.getpart(0)
            x = point.x
            y = point.y
            z = float(check_unicode(row.getValue(atr_evaluate)))
            arr_wells.append([x,y,z])
            row = rows.next()
    except ValueError:
        ARCPY.AddMessage("Нечисловое значение атрибута")
    #Возвращает массив [id точки,линия/номер,x,y,z,глубина скважины,объект,участок,тип_точки_наблюдения,
    #состояние_документирования,тип_документирования,глубина ГИС/глубина ГИС...,значение ГИС/значение ГИС...,
    #интервалы искомых стратиграфических подразделений,глубина кровли пластов, глубина подошвы пластов,
    #Jmh/Jsn/Ol/,алевролит/известняк/доломит/,среднее_значение_ГИС,количество_измерений,
    #среднеквадратическое_отклонение,ошибка_среднего]
    return arr_wells
def excent(arrXY,st):      #Функция вычисляет эксент входного точечного слоя и возвращает координаты эксента
    x_min = (1.0e400)
    x_max = - x_min
    y_min = x_min
    y_max = -y_min
    k=0
    while(k<len(arrXY)):#Вычисление минимальных и максимальных координат x и y
        x_min = min(x_min,arrXY[k][0])
        x_max = max(x_max,arrXY[k][0])
        y_min = min(y_min,arrXY[k][1])
        y_max = max(y_max,arrXY[k][1])
        k+=1
    k = None
    return x_min-st,x_max+st,y_min-st,y_max+st #Немного расширить границы матрицы
#создать матрицу для регулярной сети
def matrix_points(x_min,x_max,y_min,y_max,st):#Функция возвращает массив координат x,y, расположенных через равный интервал interval  
    #st - интервал между точками                 
    arr_points=[]
    x_min2 = x_min #в пределах вычесленного эксента
    y_min2 = y_min
    while (y_min2 < y_max):         #Приращение по оси y
        while (x_min2 < x_max):         #Приращение по оси x 
            x_min2 = x_min2+st
            arr_points.append([int(x_min2*st)/st,int(y_min2*st)/st,0])
        y_min2 = y_min2+st
        x_min2 = x_min
    del x_min2,y_min2 
    return arr_points #вернуть массив [x,y] через интервал st в пределах эксента
#функция вычисляет среднее значение параметра, по точкам, попавшим в окружность
#функция берет массив ТН, массив-матрицу, и размер окна (диаметр окружности)
def search_in_a_circle(def_arr_map_points,def_arr_matrix,a):
    arr_return=[]
    a/=2#радиус окружности
    k=0
    while(k<len(def_arr_matrix)):#перебор элементов матрицы
        x1 = def_arr_matrix[k][0]#координаты центра для текущей окружности
        y1 = def_arr_matrix[k][1]
        num_find_points=0#количество ТН попавших в текущий квадрат
        sum_find_points=0.0#сумма средних значений точек попавших в текущий квадрат
        m=0
        while(m<len(def_arr_map_points)):#перебор скважин
            x2 = def_arr_map_points[m][0]#координаты текущей ТН
            y2 = def_arr_map_points[m][1]
            distance = math.sqrt((x2-x1)**2+(y2-y1)**2)#расстояние между текущей точкой матрицы и скважиной
            if(distance<=a):#если расстояние между точками меньше расстояния радиуса искомого объекта
                num_find_points+=1#увеличить количество попавших ТН
                sum_find_points+=def_arr_map_points[m][2]#увеличить сумму значений ГИС ТН, попавших в    
            m+=1
        #Если в окружность попала хотя бы одна точка 
        if(num_find_points>0)and(sum_find_points>0.0):arr_return.append([x1,y1,num_find_points,sum_find_points/num_find_points])
        else:arr_return.append([x1,y1,0,0.0])
        k+=1
    return arr_return
#функция вычисляет среднее значение параметра, по точкам, попавшим в квадрат
#функция берет массив ТН, массив-матрицу, и размер окна (ширина квадрата) 
def search_in_a_square(def_arr_map_points,def_arr_matrix,a):
    arr_return=[]
    a/=2 #высота(или пол квадрата)
    k=0
    while(k<len(def_arr_matrix)):#перебор точек матрицы
        x=def_arr_matrix[k][0]#координаты центра для текущего квадрата
        y=def_arr_matrix[k][1]
        num_find_points=0#количество ТН попавших в текущий квадрат
        sum_find_points=0.0#сумма средних значений точек попавших в текущий квадрат
        for current_point in def_arr_map_points:#перебор скважин со средним значением
            x1=current_point[0]#координаты текущей ТН 
            y1=current_point[1]
            #ARCPY.AddMessage(str(x)+' '+str(y)+' '+str(x1)+' '+str(y1))
            if((x+a>=x1)and(x-a<=x1)and(y+a>=y1)and(y-a<=y1))or((x==x1)and(y==y1)):#если ТН попапа в квадрат
                num_find_points+=1#увеличить количество попавших ТН
                sum_find_points+=current_point[2]#увеличить сумму значений ГИС ТН, попавших в квадрат 
        #Если в квадрат попала хотя бы одна точка 
        if(num_find_points>0)and(sum_find_points>0.0):arr_return.append([x,y,num_find_points,sum_find_points/num_find_points])
        else:arr_return.append([x,y,0,0.0])
        k+=1
    return arr_return
def write_layer(lay,array_map,spatial_ref):
    fc_path,fc_name = os.path.split(lay)
    gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",spatial_ref)
    gp.AddField(lay,"количество_скважин","DOUBLE",15,15)#атрибут количества ТН попавших в квадрат сети
    gp.AddField(lay,"среднее_значение","FLOAT",15,15)#атрибут для среднего значения ГИС по ТН, попавшим в квадрат
    rows = gp.InsertCursor(lay)
    pnt = gp.CreateObject("Point")
    k=0
    while(k<len(array_map)):#создание точечного слоя
        feat = rows.NewRow()#Новая строка
        pnt.id = k#ID
        x = round(float(array_map[k][0]),3)
        y = round(float(array_map[k][1]),3)
        col_tn=array_map[k][2]#количество ТН попавших в квадрат сети
        average_tn_gis = round(array_map[k][3],2)#среднее значение ГИС по ТН, попавшим в квадрат сети
        pnt.x = x
        pnt.y = y
        feat.shape = pnt
        feat.SetValue("количество_скважин",col_tn)
        feat.SetValue("среднее_значение",average_tn_gis)
        rows.InsertRow(feat)#добавить строку к таблице слоя
        k=k+1
if __name__ == '__main__':#если модуль запущен, а не импортирован
    InpPFC = ARCPY.GetParameterAsText(0)#входной точечный слой скважин
    #название поля в таблице, где содержатся значения, по которым нужно вычислять
    #среднее значение точек, попавших в зону влияния
    atribute_value = ARCPY.GetParameterAsText(1)
    try:size_network = float(ARCPY.GetParameterAsText(2))#шаг регулярной сети
    except ValueError:ARCPY.AddMessage("Некорректное значение в поле 'шаг регулярной сети'")
    try:size_window = float(ARCPY.GetParameterAsText(3))#размер окна
    except ValueError:ARCPY.AddMessage("Некорректное значение в поле 'размер окна'")
    type_window = ARCPY.GetParameterAsText(4)#тип окна - окружность или квадрат
    OutputFC = arcpy.GetParameterAsText(5)#выходной точечный слой скважин
    arr_map_points = read_map_layer(InpPFC,atribute_value)#извлечь поля точечного слоя
    if(len(arr_map_points)==0):#если массив точек пуст - завершить работу программы
        ARCPY.AddMessage("Массив ТН пуст, программа остановлена")
        raise SystemExit(1)#завершить работу программы
    ARCPY.AddMessage("Количество точек картографического слоя: "+str(len(arr_map_points)))
    ARCPY.AddMessage("В формате:")
    string_point=""
    m=0
    while(m<len(arr_map_points[0])):
        string_point=string_point+" "+str(arr_map_points[0][m])
        m+=1
    ARCPY.AddMessage(string_point)
    ARCPY.AddMessage("-------------------------------------------------------------------------------")
    x_min1, x_max1, y_min1, y_max1 = excent(arr_map_points,size_network)#вычислить координаты эксента
    arr_matrix = matrix_points(x_min1, x_max1, y_min1, y_max1,size_network)#генерация матрицы точек с заданным интервалом
    arcpy.AddMessage("количество точек матрицы: "+str(len(arr_matrix)))
    average_points = []
    if type_window==u"Окружность":
        average_points = search_in_a_circle(arr_map_points,arr_matrix,size_window)#поиск точек, попавших в окружности матрицы
    elif type_window==u"Квадрат":
        average_points = search_in_a_square(arr_map_points,arr_matrix,size_window)#поиск точек, попавших в квадраты матрицы
    Spat_ref = gp.describe(InpPFC).spatialreference
    if len(average_points)>0:
        write_layer(OutputFC,average_points,Spat_ref)
    else:
        ARCPY.AddMessage("Массив для выходного точечного слоя пуст - слой не создан")
