import sys, arcgisscripting, os
import arcpy as ARCPY
gp = arcgisscripting.create(9.3)

#������� �� ����� ����� �������� ���� �� ���������� ��������, ������� ����������� � ������� Script10
#� ���������� ������ [�������� ��������, x,y,������� ���������, �������� ���������]
def read_map_layer(layer):
        def check_unicode(text):#������� ����������� ��������� ������ � utf-8 
                if(type(text)==unicode):text = text.encode('utf-8')
                return text
        rows = gp.SearchCursor(layer)
        row = rows.Next()
        layer = gp.describe(layer)
        arr_wells=[]#������ ��� ����������� �������
        while row:
                feat = row.GetValue(layer.ShapeFieldName)
                point = feat.getpart(0)
                id_well = check_unicode(row.GetValue("id_well"))#id ��������
                name_well = check_unicode(row.GetValue("name_well"))#�������� ��������
                x = point.x
                y = point.y
                z = check_unicode(row.GetValue("z"))#abs ����� ��������
                depth_well = check_unicode(row.GetValue("depth_well"))#������� ��������
                name_object = check_unicode(row.GetValue("name_object"))#�������� �������
                name_area = check_unicode(row.GetValue("name_area"))#�������� �������
                depth_logging = check_unicode(row.GetValue("depth_logging"))#������, ������� �������� �������� ������
                value_logging = check_unicode(row.GetValue("value_logging"))#������, ������� �������� ���������� �������� ���
                recover = check_unicode(row.GetValue("recover"))#�������� ������������� ���������
                abs_litostrat_roof = check_unicode(row.GetValue("abs_litostrat_roof"))#abs ������ ��������������������� �������
                abs_litostrat_sole = check_unicode(row.GetValue("abs_litostrat_sole"))#abs ������� ��������������������� �������
                stratigraphy = check_unicode(row.GetValue("stratigraphy"))#������������ �������
                lithology = check_unicode(row.GetValue("lithology"))#��������� �������
                #�������� � �������� ������ �������� � ������� ��������
                arr_wells.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,depth_logging,\
                value_logging,recover,abs_litostrat_roof,abs_litostrat_sole,stratigraphy,lithology])
                row = rows.next()
        #���������� ������ [id �����,�����/�����,x,y,z,������� ��������,������,�������,
        #�������/�������...,�������� ���/�������� ���...,�������� ������������� ���������,
        #������� ������ ������/������� ������ ������/������� ������ ������,Jmh/Jsn/Ol/,���������/���������/�������/]
        return arr_wells 
#������� �� ����� ����� ������, ��������� �� ����������������� ���� � �������� �������
#� ���������� ������ � ������� [�������� ��������, x,y, �������� ���������� ��������������� �� �������� �������]
def select_value_depth(arr_map,given_depth):
        k=0
        arr_output=[]
        while(k<len(arr_map)):#������� ��������� ������� �� ����������������� ����
                id_well = str(arr_map[k][0])#id ��������
                z = str(arr_map[k][4])#abs ����� ��������
                if(z<>'')and(z<>' ')and(z<>'null'):#���� ���� ������ �� abs ����� ������� ��������
                        str_depth_logging = str(arr_map[k][8])#������ �������� ������� ��������� ��� ��� ������� ��������
                        str_value_logging = str(arr_map[k][9])#������ �������� �������� ��������� ��� ��� ������� ��������
                        arr_depth_logging = str_depth_logging.split("/")#�������� ������ ������ ��������� ��� ��� ������� ��������
                        arr_value_logging = str_value_logging.split("/")#�������� ������ �������� ��� ��� ������� �������� 
                        if(len(arr_depth_logging)==len(arr_value_logging)):#���� ������� ������ ��� � ���������� �������� ��� �����
                                z = round(float(z),2)#������������� � ������������ ��� � ����� ������� ����� �������
                                given_depth = round(float(given_depth),2)#������������� � ������������ ��� � ����� ������� ����� �������
                                #��������� ������� �� ����� �������� �� �������� ������� � �������� ���������� ������� �����
                                depth_to_val = z - given_depth
                                #���� �������� ������ ����, ������ ����� �������� ��������� ����
                                #�������� ������� � ������������� ����� ��������� ����������
                                if (depth_to_val>0):
                                        m=0

                                        #������� �������� ������ ��������� ���
                                        while(m<len(arr_depth_logging)-1):#�� ������������ ��������� ������� ������� - ��������� '1/2/3/' ���� ['1','2','3','']
                                                similar = depth_to_val-round(float(arr_depth_logging[m]),2)
                                                ARCPY.AddMessage(str(depth_to_val)+" - "+str(arr_depth_logging[m])+" = "+str(similar))
                                                if(similar<0):
                                                        ARCPY.AddMessage("������������� ��������")
                                                        break
                                                m+=1
             
                                
                        else:
                                ARCPY.AddMessage("������� ������ ��������� ��� � ����� ���������� \
                                �������� �� �����, ��� ��������: "+str(id_well))
                ARCPY.AddMessage("------------------------------")
                k=k+1
        return arr_output #������� ������ [name_well,x,y,val]
#������� �� ����� �����: �������� ��������� ��������� ����, �������� ������ ��� ������, � ���������������� ��������
#� ��������� ���������������� ���� 
def write_layer_sm(FC,arr_points,spatial_ref):
        fc_path,fc_name = os.path.split(FC)
        gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",spatial_ref)
        gp.AddField(FC,"name","STRING",15,15)
        gp.AddField(FC,"val_sm","FLOAT",15,15)
        rows = gp.InsertCursor(FC)
        pnt = gp.CreateObject("Point")
        k=0
        while(k<len(arr_points)):#�������� ��������� ����
                feat = rows.NewRow()#����� ������
                pnt.id=k#ID
                pnt.x=arr_points[k][1]
                pnt.y=arr_points[k][2]
                feat.shape=pnt
                feat.SetValue("name",arr_points[k][0])#�������� ��������
                feat.SetValue("val_sm",float(arr_points[k][3]))#�������� ���������� ���������������
                rows.InsertRow(feat)
                k=k+1
if __name__== '__main__':#���� ������ �������, � �� ������������
        #-------------------------------------������� ���������----------------------------------------
        InpPFC = ARCPY.GetParameterAsText(0)#������� �������� ���� �������
        OutputFC = arcpy.GetParameterAsText(1)#�������� �������� ���� ������� 
        depth = ARCPY.GetParameterAsText(2)#������� ������� ���������� �������� ���������� ���������������
        Full_report = ARCPY.GetParameterAsText(3)#�������� ���������� �������� �� ������ �����
        #----------------------------------------------------------------------------------------------
        arr_map_points = read_map_layer(InpPFC)#������� ���� ��������� ����
        if(len(arr_map_points)==0):#���� ������ ����� - ��������� ������ ���������
                raise SystemExit(1)#��������� ������ ���������
        ARCPY.AddMessage("���������� ����� ����������������� ����: "+str(len(arr_map_points)))
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                string_point=""
                m=0
                while(m<len(arr_map_points[0])):
                        string_point=string_point+" "+str(arr_map_points[0][m])
                        m+=1
                ARCPY.AddMessage(string_point)
        arr_select = select_value_depth(arr_map_points,depth)     
        
        
'''
gp.AddMessage("���������� �������, ��������� �� ���� = "+str(len(arr_sm)))
arr_select = select_value(arr_sm,depth)#�������� ������ ����� �� ���������� ���������� ��������������� �� �������� �������
gp.AddMessage("���������� �������, � ������� ��������� ��������������� ��� �������� ������� = "+str(len(arr_select)))
gp.AddMessage("� �������:")
k=0
while(k<10):
        gp.AddMessage(arr_select[k])
        k=k+1
gp.AddMessage("------------------------------------------------------------")
Spat_ref = gp.describe(InpPFC).spatialreference
write_layer_sm(OutputFC,arr_select,Spat_ref)#�������� �������� ������ � ���������������� ����
'''
