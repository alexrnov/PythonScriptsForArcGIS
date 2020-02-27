import sys, arcgisscripting, os, math
import arcpy as ARCPY
gp = arcgisscripting.create(9.3)

#������� �� ����� ����� �������� ���� �� ���������� ��������, ������� ����������� � ������� Script10_2
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

                abs_roof_bark = round(check_unicode(row.GetValue("abs_roof_bark")),1)#������ ���� ������������
                abs_sole_bark = round(check_unicode(row.GetValue("abs_sole_bark")),1)#������� ���� ������������
                depth_bark = round(check_unicode(row.GetValue("depth_bark")),1)#�������� ���� ������������
                
                abs_litostrat_roof = check_unicode(row.GetValue("abs_litostrat_roof"))#abs ������ ��������������������� �������
                abs_litostrat_sole = check_unicode(row.GetValue("abs_litostrat_sole"))#abs ������� ��������������������� �������
                stratigraphy = check_unicode(row.GetValue("stratigraphy"))#������������ �������
                lithology = check_unicode(row.GetValue("lithology"))#��������� �������
                #�������� � �������� ������ �������� � ������� ��������
                arr_wells.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,depth_logging,\
                value_logging,abs_roof_bark,abs_sole_bark,depth_bark,abs_litostrat_roof,abs_litostrat_sole,stratigraphy,lithology])
                row = rows.next()
        #���������� ������ [id �����,�����/�����,x,y,z,������� ��������,������,�������,
        #�������/�������...,�������� ���/�������� ���...,abs ������ ���� ������������, abs ������� ���� ������������, �������� ���� ������������
        #������� ������ ������/������� ������ ������/������� ������ ������,Jmh/Jsn/Ol/,���������/���������/�������/]
        return arr_wells 
#������� �� ����� ����� ������, ��������� �� ����������������� ����
#� ���������� ������ � �������
#[id ��������, �����/�����, x,y,z,������� ��������,�������� �������,
#�������� �������, �������� ������������� ���������, c���� ���������� ��������� �� ��������� ����������,
#����� ��������� ��� �� ��������� ����������,������� ���������� �������� ��� �� ��������� ����������]
def select_value_depth(arr_map):
        k=0
        arr_output=[]#�������� ������
        while(k<len(arr_map)):#������� ��������� ������� �� ����������������� ����
                bark = float(arr_map[k][12])#�������� ���� ������������
                if(bark==-1.0):pass#���� ��� ������ �� ���� - ���������� ������� ��������
                else:#���� ���� ������ � ������������� ����������
                        id_well = str(arr_map[k][0])#id ��������
                        str_depth_logging = str(arr_map[k][8])#������ �������� ������� ��������� ��� ��� ������� ��������
                        str_value_logging = str(arr_map[k][9])#������ �������� �������� ��������� ��� ��� ������� ��������
                        arr_depth_logging = str_depth_logging.split("/")#�������� ������ ������ ��������� ��� ��� ������� ��������
                        arr_value_logging = str_value_logging.split("/")#�������� ������ �������� ��� ��� ������� ��������
                        arr_dimension_contain=[]#������ ��� ������ ������ ��������� � ���� ������������
                        arr_value_contain=[]#������ ��� ������ �������� ��������� � ���� ������������
                        roof_bark = float(arr_map[k][10])#abs ������ ���� ������������
                        sole_bark = float(arr_map[k][11])#��s ������� ���� ������������
                        if(len(arr_depth_logging)<>len(arr_value_logging)):#���� ������� ������ ��� � ���������� �������� ��� �� �����
                                ARCPY.AddMessage("������� ������ ��������� ��� � ����� ���������� \
                                �������� �� �����, ��� ��������: "+str(id_well))
                        else:
                                bool_bark=False#���������� ���������� ���������� ����� �� ���� �� ���� ������������
                                depth_bark=0#��������� ���������� ��������� � ���� ������������
                                value_bark=0.0#��������� �������� ���������� ��������� ��� ���� ������������
                                dep=0.0
                                if(bark>=0.1):#���� �������� ���� ������������ ������ ����
                                        m=0
                                        #������� ������ ��������� ��� ��� ������� ��������
                                        while(m<len(arr_depth_logging)-1):#�� ������������ ��������� ������� �������, ��������� 1/2/3 ���� ['1','2','3','']
                                                #������������� ������� �������� ������� � ������������ ���
                                                depth_float = round(float(arr_depth_logging[m]),2)
                                                #������������� ������� �������� ��������� � ������������ ���
                                                val_float = round(float(arr_value_logging[m]),2)
                                                #���� ���� ����� ����� �� ���� ������������
                                                if(bool_bark):
                                                        #���� ������� ������� ��������� ������ ������� ���� ������������
                                                        if(depth_float<sole_bark):
                                                                #���� �������� - ���������
                                                                if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                        #�������� � ������ ������ ��������� � ���� ������������ - ������� ������� ���������
                                                                        arr_dimension_contain.append(depth_float)
                                                                        #�������� � ������ �������� ��������� �� ��������� - ������� �������� ���������
                                                                        arr_value_contain.append(val_float)
                                                                        #��������� � ����� ���������� ��������� ������� ��������
                                                                        value_bark=value_bark+val_float
                                                                        depth_bark+=1#��������� ���������� ��������� �� ��������� ����������
                                                        #���� ��� ����� �� ������� ���� ������������ - ��������� ����
                                                        else:break#����� �� �����
                                                #���� ���� ��� �� ����� �� ���� ������������
                                                else:
                                                        #���� ������� ������� ��������� ��� ����� ��� ������ ��� ���������
                                                        #abs ������ ���� ������������, ������ ���� ����� �� ���� ������������
                                                        if(depth_float==roof_bark)or(depth_float>roof_bark):
                                                                bool_bark=True
                                                                dep=depth_float
                                                                #���� �������� - ���������
                                                                if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                        #�������� � ������ ������ ��������� � ���� ������������ - ������� ������� ���������
                                                                        arr_dimension_contain.append(depth_float)
                                                                        #�������� � ������ �������� ��������� � ���� ������������- ������� �������� ���������
                                                                        arr_value_contain.append(val_float)
                                                                        #�������� ������� �������� ��������� ��� � ����� ���������
                                                                        #� ���� ������������
                                                                        value_bark=value_bark+val_float
                                                                        depth_bark+=1#��������� ���������� ��������� � ���� ������������
                                                m+=1
                                        #���� ���������� ��������� � ���� ������������ ������ ����
                                        #� ���� ����� ���������� ��������� ������ ����
                                        # - ��������� ������� ���������� �������� ��� � ���� ������������
                                        #���� ���������� ��������� � ����� ��������� ������ ����
                                        #��������� ������� ��������
                                        if(depth_bark>0)and(value_bark>0.0):
                                                sum_value_bark=round(value_bark/depth_bark,2)#��������� ������� 
                                                msd=0.0 #�������������������� ����������
                                                standard_error = 0.0 #����������� ������ ��������
                                                if (len(arr_dimension_contain)==len(arr_value_contain)):
                                                        i=0
                                                        while(i<len(arr_value_contain)):
                                                                msd=msd+(arr_value_contain[i]-sum_value_bark)**2
                                                                i+=1
                                                        msd=math.sqrt(msd/depth_bark)
                                                        standard_error = msd/math.sqrt(depth_bark)#��������� ������ ��������
                                                else:
                                                        ARCPY.AddMessage("���������� ��������� � ������� \
                                                        ������ ��������� � ������� �� ���������� ��������� �� �����")
                                            
                                                curr_array=[id_well,arr_map[k][1],arr_map[k][2],arr_map[k][3],\
                                                round(arr_map[k][4],2),arr_map[k][5],arr_map[k][6],arr_map[k][7],\
                                                bark,value_bark,depth_bark,sum_value_bark,msd,standard_error]
                                                arr_output.append(curr_array)                                
                k=k+1
        #������� ������ � ������� [id ��������, �����/�����, x,y,z,������� ��������,�������� �������,
        #�������� �������, �������� ���� ������������, ����� ���������� ��������� �� ��������� ����������,
        #����� ��������� ��� �� ��������� ����������, ������� ���������� �������� ��� �� ��������� ����������,
        #�������������������� ����������, ����������� ������ ��������]
        return arr_output 
#������� �� ����� ����� �������� ��������� ��������� ����, �������� ������ ��� ������, � ���������������� ��������
#� ��������� ���������������� ���� 
def write_layer(FC,array_map,spatial_ref):
        fc_path,fc_name = os.path.split(FC)
        gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",spatial_ref)
        gp.AddField(OutputFC,"id_well","STRING",15,15)#������� id �������
	gp.AddField(OutputFC,"name_well","STRING",15,15)#������� �������� �������
	gp.AddField(OutputFC,"x","FLOAT",15,15)#�������� ������� x
	gp.AddField(OutputFC,"y","FLOAT",15,15)#�������� ������� y
	gp.AddField(OutputFC,"z","FLOAT",15,15)#�������� ������� z
	gp.AddField(OutputFC,"depth_well","FLOAT",15,15)#�������� ������� ������� ��������
	gp.AddField(OutputFC,"name_object","STRING","","",100)#������� - �������� �������
	gp.AddField(OutputFC,"name_area","STRING","","",50)#������� - �������� �������
        gp.AddField(OutputFC,"depth_bark","FLOAT",15,15)#�������� ������� �������� ���� ������������
        gp.AddField(OutputFC,"value_rec","FLOAT",15,15)#�������� ������� ����� �������� ��������� ��� �� ���������
        gp.AddField(OutputFC,"depth_rec","FLOAT",15,15)#�������� ������� ���������� ��������� �� ��������� ���������
        gp.AddField(OutputFC,"mean_value_gis","FLOAT",15,15)#�������� ������� ��� �������� ����������� �������� ��� �� ��������� ����������
        gp.AddField(OutputFC,"msd","FLOAT",15,15)#�������� ������� ��� ��������������������� ����������
        gp.AddField(OutputFC,"error_mean","FLOAT",15,15)#�������� ������� ��� ������ ������ ��������
        rows = gp.InsertCursor(FC)
        pnt = gp.CreateObject("Point")
        k=0
        while(k<len(array_map)):#�������� ��������� ����
                id_well = float(array_map[k][0])#id �������
    		name_well = str(array_map[k][1])#�������� �������
    		z = round(float(array_map[k][4]),3)#z
                depth_well = float(array_map[k][5])#������� ��������
    		name_obj = str(array_map[k][6])#�������� �������
    		name_uchastok = str(array_map[k][7])#�������� �������
    		depth_bark = float(array_map[k][8])#������� �������� ���� ������������
    		value_rec = float(array_map[k][9])#����� �������� ��������� ��� �� ���������
    		depth_rec = float(array_map[k][10])#����� ���������� ��������� �� ��������� ���������
    		mean_value_gis = array_map[k][11]#������� ���������� �������� ��� �� ��������� ����������
    		msd_value = array_map[k][12]#�������������������� ����������
    		error_mean = array_map[k][13]#������ ������ ��������
                feat = rows.NewRow()#����� ������
                pnt.id=k#ID
                x = round(float(array_map[k][2]),3)
    		y = round(float(array_map[k][3]),3)
    		pnt.x = x
    		pnt.y = y
                feat.shape = pnt
    		feat.SetValue("id_well",id_well)#�������� �������� � ������������ �����
    		feat.SetValue("name_well",name_well)
		feat.SetValue("x",x)
		feat.SetValue("y",y)
    		feat.SetValue("z",z)
    		feat.SetValue("depth_well",depth_well)
    		feat.SetValue("name_object",name_obj)
    		feat.SetValue("name_area",name_uchastok)
    		feat.SetValue("depth_bark",depth_bark)
    		feat.SetValue("value_rec",value_rec)
    		feat.SetValue("depth_rec",depth_rec)
    		feat.SetValue("mean_value_gis",mean_value_gis)
    		feat.SetValue("msd",msd_value)
    		feat.SetValue("error_mean",error_mean)
    		rows.InsertRow(feat)#�������� ������ � ������� ����
                k=k+1
if __name__== '__main__':#���� ������ �������, � �� ������������
        #-------------------------------------������� ���������----------------------------------------
        InpPFC = ARCPY.GetParameterAsText(0)#������� �������� ���� �������
        OutputFC = arcpy.GetParameterAsText(1)#�������� �������� ���� ������� 
        Full_report = ARCPY.GetParameterAsText(2)#�������� ���������� �������� �� ������ �����
        #----------------------------------------------------------------------------------------------
        arr_map_points = read_map_layer(InpPFC)#������� ���� ��������� ����

        if(len(arr_map_points)==0):#���� ������ ����� ���� - ��������� ������ ���������
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
                ARCPY.AddMessage("-------------------------------------------------------------------------------")

        #�������� ������� ���������� �������� ��� � ������������������� ���������� ��� ���� �������
        arr_select = select_value_depth(arr_map_points)
        if(len(arr_select)==0):#���� ������ ����
                raise SystemExit(0)
        ARCPY.AddMessage("���������� ������� � ����������� ��� �� �������� ������� ��� ������ � ���������������� ����: "+str(len(arr_select)))
        if(Full_report == "true"):
                ARCPY.AddMessage("� �������:")
                ARCPY.AddMessage("id ��������: "+str(arr_select[0][0]))
                ARCPY.AddMessage("�������� ��������: "+str(arr_select[0][1]))
                ARCPY.AddMessage("x: "+str(arr_select[0][2]))
                ARCPY.AddMessage("y: "+str(arr_select[0][3]))
                ARCPY.AddMessage("z: "+str(arr_select[0][4]))
                ARCPY.AddMessage("������� ��������: "+str(arr_select[0][5]))
                ARCPY.AddMessage("�������� �������: "+str(arr_select[0][6]))
                ARCPY.AddMessage("�������� �������: "+str(arr_select[0][7]))
                ARCPY.AddMessage("�������� ���� ������������: "+str(arr_select[0][8]))
                ARCPY.AddMessage("����� ��������� ��� �� ��������� ����������: "+str(arr_select[0][9]))
                ARCPY.AddMessage("���������� ��������� �� ��������� ����������: "+str(arr_select[0][10]))
                ARCPY.AddMessage("������� ���������� �������� ��� � ���� ������������: "+str(arr_select[0][11]))
        
        Spat_ref = gp.describe(InpPFC).spatialreference
        write_layer(OutputFC,arr_select,Spat_ref)#�������� �������� ������ � ���������������� ����



