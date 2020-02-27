import os, arcpy, arcgisscripting
def read_text_file(txt_file):
    arr_p=[]
    f = open(txt_file,'r')
    for line in f:
        arr=line.split(" ")
        curr_arr=[]
        for element_line in arr:
            if element_line!='':
                curr_arr.append(element_line)
        #arcpy.AddMessage(str(curr_arr))
        try:
            arr_p.append([float(curr_arr[0]),float(curr_arr[1]),float(curr_arr[2])])
        except ValueError:
            arcpy.AddMessage(str(arr_p))
        
    '''
    for line in f:
        arr=line.split(" ")
        try:
            arr_p.append([str(arr[0]),float(arr[1]),float(arr[2]),float(arr[3][:len(arr[3])-1])])#для z не считывать \n
        except ValueError:
            arr_p.append([str(arr[0]),float(arr[1]),float(arr[2]),0.0])#для z не считывать \n
    '''
    f.close()
    return arr_p
def write_map_layer(lay,CS,map_arr):
    fc_path,fc_name = os.path.split(lay)
    gp=arcgisscripting.create(9.3)#создать объект геопроцессинга версии 9.3
    gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",CS)
    gp.AddField(lay,"id","STRING",15,15)#атрибут id скважин
    gp.AddField(lay,"x","FLOAT",15,15)#добавить атрибут x
    gp.AddField(lay,"y","FLOAT",15,15)#добавить атрибут y
    gp.AddField(lay,"z","FLOAT",15,15)#добавить атрибут z
    rows = gp.InsertCursor(lay)
    pnt = gp.CreateObject("Point")
    k=0
    while(k<len(map_arr)):
        feat = rows.NewRow()#Новая строка
        pnt.id = k
        pnt.x = map_arr[k][0]-20000.0
        pnt.y = map_arr[k][1]-10000.0
        feat.shape = pnt
        feat.SetValue("id",str(k))#добавить значения к атрибутивным полям
        feat.SetValue("x",map_arr[k][0]-20000.0)
        feat.SetValue("y",map_arr[k][1]-10000.0)
        feat.SetValue("z",map_arr[k][2])
        rows.InsertRow(feat)#добавить строку к таблице слоя
        k+=1
if __name__=='__main__':
    InputTextFile = arcpy.GetParameterAsText(0)
    OutputFC = arcpy.GetParameterAsText(1)
    C_system = arcpy.GetParameterAsText(2)
    arcpy.AddMessage(str(InputTextFile.encode('utf-8')))
    arr_points=read_text_file(InputTextFile)
    k=0
    while(k<10):
        arcpy.AddMessage(str(arr_points[k]))
        k+=1
    write_map_layer(OutputFC,C_system,arr_points)
    
    
