import sys, arcgisscripting, os, math
import arcpy as ARCPY
from numpy import median
gp = arcgisscripting.create(9.3)
#������� �� ����� ����� �������� ���� �� ���������� ��������, ������� ����������� � ������� Script10_3
#� ���������� ������ �����
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
                id_well = check_unicode(row.GetValue("id_��������"))#id ��������
                name_well = check_unicode(row.GetValue("��������_��������"))#�������� ��������
                x = point.x
                y = point.y
                z = round(check_unicode(row.GetValue("z")),1)#abs ����� ��������
                depth_well = check_unicode(row.GetValue("�������_��������"))#������� ��������
                name_object = check_unicode(row.GetValue("��������_�������"))#�������� �������
                name_area = check_unicode(row.GetValue("��������_�������"))#�������� �������
                type_tn = check_unicode(row.GetValue("���_�����_����������"))#��� ����� ����������
                status_doc = check_unicode(row.GetValue("���������_����������������"))#���������_����������������
                type_doc = check_unicode(row.GetValue("���_����������������"))#���_����������������
                depth_logging = check_unicode(row.GetValue("�������_���������_���"))#������, ������� �������� �������� ������
                value_logging = check_unicode(row.GetValue("��������_���"))#������, ������� �������� ���������� �������� ���
                interval_abs = check_unicode(row.GetValue("���������_��������_�������������"))#���������_��������_�������������
                abs_litostrat_roof = check_unicode(row.GetValue("�������_������_�������"))#abs ������ ��������������������� �������
                abs_litostrat_sole = check_unicode(row.GetValue("�������_�������_�������"))#abs ������� ��������������������� �������
                stratigraphy = check_unicode(row.GetValue("������������"))#������������ �������
                lithology = check_unicode(row.GetValue("���������"))#��������� �������
                #�������� � �������� ������ �������� � ������� ��������
                arr_wells.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,type_tn,status_doc,type_doc,\
                depth_logging,value_logging,interval_abs,abs_litostrat_roof,abs_litostrat_sole,stratigraphy,lithology])
                row = rows.next()
        #���������� ������ [id �����,�����/�����,x,y,z,������� ��������,������,�������,���_�����_����������,
        #���������_����������������,���_����������������,������� ���/������� ���...,�������� ���/�������� ���...,
        #��������� ������� ����������������� �������������,������� ������ �������, ������� ������� �������,
        #Jmh/Jsn/Ol/,���������/���������/�������/]
        return arr_wells
def select_values(arr_map):
        arr_width_average=[]#�������� ������ �� �� �������� ���������� ��� ��� ����������
        for current_point in arr_map:#������� ��
                if(current_point[11]=='��� ������')or(current_point[12]=='��� ������'):continue#���� ��� ������ �� ���
                if(current_point[13]=='-1.0'):continue#���� ��� ������ �� ��� �������� ������ - ������� � ��������� �� 
                intervals = current_point[13].split(";")#�������� ������ � ������� ���������� ������� ����������������� �������������
                arr_depth_logging = current_point[11].split("/")#�������� ������ ������ ��������� ��� ��� ������� ��������
                arr_value_logging = current_point[12].split("/")#�������� ������ �������� ��� ��� ������� ��������
                col_dimension=0.0#��������� ���������� ��������� � ������� ����������������� ��������������
                sum_vales=0.0#����� �������� ��� � ������� ����������������� ��������������
                arr_dimension_contain=[]#������ ��� ������ ������ ��������� � ������� ����������������� �������
                arr_value_contain=[]#������ ��� ������ �������� ��������� � ������� ����������������� �������
                depth_interval=0#��������� ���������� ��������� � ������� ����������������� �������
                value_interval=0.0#��������� �������� ���������� ��������� � ������� ����������������� �������
                m=0
                while m<len(intervals)-1:#������� ������� ����������, �� ������������ ��������� ������� ������� - ��������� '45;' ���� ['45','']
                        #depth_interval=0#��������� ���������� ��������� � ������� ����������������� �������
                        #value_interval=0.0#��������� �������� ���������� ��������� � ������� ����������������� �������
                        arr_int = intervals[m].split("-")#�������� ������ �������� ��������� �������['10.0','10.5']
                        abs_roof = float(arr_int[0])#��� ������ �������� ��������� �������� ������������������ ������
                        abs_sole = float(arr_int[1])#��� ������� �������� ��������� �������� ������������������ ������
                        if len(arr_depth_logging)<>len(arr_value_logging) :#���� ������� ������ ��� � ���������� �������� ��� �� �����
                                ARCPY.AddMessage("������� ������ ��������� ��� � ����� ���������� \
                                �������� �� �����, ��� ��������: "+str(current_point[0]))
                        else:
                                if(abs_sole-abs_roof>=0.1):#���� �������� �������� ������������� ������ 0.1
                                        bool_interval=False#���������� ���������� ���������� ����� �� ���� �� �������� ������������������ �������������
                                        n=0
                                        #������� ������ ��������� ��� ������� ��������(��� �������� ������������������ �������������)
                                        while(n<len(arr_depth_logging)-1):#�� ������������ ��������� ������� �������, ��������� 1/2/3 ���� ['1','2','3','']
                                                depth_float = round(float(arr_depth_logging[n]),2)#������������� ������� �������� ������� � ������������ ���
                                                val_float = round(float(arr_value_logging[n]),2)#������������� ������� �������� ��������� � ������������ ���
                                                #���� ���� ����� ����� �� �������� ������������������ �������������
                                                if(bool_interval):
                                                        #���� ������� ������� ��������� ������ ������� ���� ������������
                                                        if(depth_float<=abs_sole):
                                                                #ARCPY.AddMessage(" 1 ������� ������� = "+str(depth_float)+" ������ ������ = "+str(abs_roof)+" ������� ������ = "+str(abs_sole))
                                                                #���� �������� - ���������
                                                                if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                        #�������� � ������ ������ ��������� � ������� ����������������� ������
                                                                        #- ������� ������� ���������
                                                                        arr_dimension_contain.append(depth_float)
                                                                        #�������� � ������ �������� ��������� �� ������� ����������������� ������
                                                                        #- ������� �������� ��������� ���
                                                                        arr_value_contain.append(val_float)
                                                                        #��������� � ����� ���������� ��������� ������� ��������
                                                                        value_interval=value_interval+val_float
                                                                        depth_interval+=1#��������� ���������� ��������� � �������������
                                                        #���� ��� ����� �� ������� ���� ������������ - ��������� ����
                                                        else:break#����� �� �����(������� ������ ��� �������� ������������������ �������������)
                                                else:#���� ���� ��� �� ����� �� �������� �������������
                                                        #���� ������� ������� ��������� ��� ����� ��� ������ ��� ���������
                                                        #abs ������ ���� ������������, ������ ���� ����� �� �������� ������������������ �������������
                                                        if(depth_float>=abs_roof)and(depth_float<=abs_sole):
                                                                #ARCPY.AddMessage("0 ������� ������� = "+str(depth_float)+" ������ ������ = "+str(abs_roof)+" ������� ������ = "+str(abs_sole))
                                                                bool_interval=True
                                                                #���� �������� - ���������
                                                                if(val_float<>-999999.0)and(val_float<>-999.75)and(val_float<>-995.75):
                                                                        #�������� � ������ ������ ��������� � ������� ����������������� ������
                                                                        #- ������� ������� ���������
                                                                        arr_dimension_contain.append(depth_float)
                                                                        #�������� � ������ �������� ��������� �� ������� ����������������� ������
                                                                        #- ������� �������� ��������� ���
                                                                        arr_value_contain.append(val_float)
                                                                        #��������� � ����� ���������� ��������� ������� ��������
                                                                        value_interval=value_interval+val_float
                                                                        depth_interval+=1#��������� ���������� ��������� � �������������                     
                                                n+=1
                        m+=1
                        
                #��� ������� �� ��������� ��� ��� ���������� ������� � ����� �������� ������ 0
                if(depth_interval>0)and(value_interval>0):
                        average_gis=round(value_interval/depth_interval,2)#��������� ������� 
                        msd=0.0 #�������������������� ����������
                        standard_error = 0.0 #����������� ������ ��������
                        if (len(arr_dimension_contain)==len(arr_value_contain)):
                                i=0
                                while(i<len(arr_value_contain)):
                                        msd=msd+(arr_value_contain[i]-average_gis)**2
                                        i+=1
                                msd=math.sqrt(msd/depth_interval)
                                standard_error = msd/math.sqrt(depth_interval)#��������� ������ ��������
                        else:
                                ARCPY.AddMessage("���������� ��������� � ������� \
                                ������ ��������� � ������� �� ���������� ��������� �� �����")
                        #ARCPY.AddMessage(str(current_point[0])+", ����� ���="+str(value_interval)+\
                        #" ���������� ���="+str(depth_interval)+ ", �������="+str(round(value_interval/depth_interval,2)))
                        arr_width_average.append([current_point[0],current_point[1],current_point[2],\
                        current_point[3],current_point[4],current_point[5],current_point[6],current_point[7],
                        current_point[8],current_point[9],current_point[10],current_point[11],\
                        current_point[12],current_point[13],current_point[14],current_point[15],\
                        current_point[16],current_point[17],average_gis,depth_interval,msd,standard_error,median(arr_value_contain)])
        #���������� ������ [id �����,�����/�����,x,y,z,������� ��������,������,�������,���_�����_����������,
        #���������_����������������,���_����������������,������� ���/������� ���...,�������� ���/�������� ���...,
        #��������� ������� ����������������� �������������,������� ������ �������, ������� ������� �������,
        #Jmh/Jsn/Ol/,���������/���������/�������/,������� �������� ��� �� ������� ����������,���������� ���������,
        #�������������������� ����������,������ ��������,�������]
        return arr_width_average

#���������� ������ [id �����,�����/�����,x,y,z,������� ��������,������,�������,���_�����_����������,
#���������_����������������,���_����������������,������� ���/������� ���...,�������� ���/�������� ���...,
#��������� ������� ����������������� �������������,������� ������ �������, ������� ������� �������,
#Jmh/Jsn/Ol/,���������/���������/�������/,������� �������� ��� �� ������� ����������,
#�������������������� ����������,������ ��������,�������]
def write_layer(lay,array_map,spatial_ref):
        fc_path,fc_name = os.path.split(lay)
        gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",spatial_ref)
        gp.AddField(lay,"id_��������","STRING",15,15)#������� id �������
	gp.AddField(lay,"��������_��������","STRING",15,15)#������� �������� �������
	gp.AddField(lay,"x","FLOAT",15,15)#�������� ������� x
	gp.AddField(lay,"y","FLOAT",15,15)#�������� ������� y
	gp.AddField(lay,"z","FLOAT",15,15)#�������� ������� z
	gp.AddField(lay,"�������_��������","FLOAT",15,15)#�������� ������� ������� ��������
	gp.AddField(lay,"��������_�������","STRING","","",100)#������� - �������� �������
	gp.AddField(lay,"��������_�������","STRING","","",50)#������� - �������� �������
        gp.AddField(lay,"���_�����_����������","STRING","","",100)#������� - ��� ����� ����������
        gp.AddField(lay,"���������_����������������","STRING","","",100)#������� - ��������� ����������������
        gp.AddField(lay,"���_����������������","STRING","","",100)#������� - ��� ����������������
	gp.AddField(lay,"�������_���������_���","STRING","","",1000)#������� �������� ������ ��������� ���
	gp.AddField(lay,"��������_���","STRING","","",1000)#������� ��� �������� ���
        gp.AddField(lay,"���������_��������_�������������","STRING","","",1000)#������� ���������� ������� ������� ����������������� �������������
        gp.AddField(lay,"�������_������_�������","STRING","","",1000)#������� ��� ������ ������ �������
        gp.AddField(lay,"�������_�������_�������","STRING","","",1000)#������� ��� ������ ������� �������
        gp.AddField(lay,"������������","STRING","","",1000)#������� ��� ������������
        gp.AddField(lay,"���������","STRING","","",1000)#������� ��� ���������
        gp.AddField(lay,"�������_��������_���","FLOAT",15,15)#������� ������� �������� ���
        gp.AddField(lay,"����������_���������","FLOAT",15,15)#������� ��� ���������� ���������
        gp.AddField(lay,"��������������������_����������","FLOAT",15,15)#������� ��� ���������������������_����������
        gp.AddField(lay,"������_��������","FLOAT",15,15)#������� ��� ������ ��������
        gp.AddField(lay,"�������","FLOAT",15,15)#������� ��� �������
        rows = gp.InsertCursor(lay)
        pnt = gp.CreateObject("Point")
        k=0
        while(k<len(array_map)):#�������� ��������� ����
                id_well = float(array_map[k][0])#id �������
    		name_well = str(array_map[k][1])#�������� �������
    		z = round(float(array_map[k][4]),3)#z
                depth_well = float(array_map[k][5])#������� ��������
    		name_obj = str(array_map[k][6])#�������� �������
    		name_uchastok = str(array_map[k][7])#�������� �������
                name_type_tn = str(array_map[k][8])#������� - ��� ����� ����������
    		name_status_doc = str(array_map[k][9])#������� - ��������� ����������������
    		name_type_doc = str(array_map[k][10])#������� - ��� ����������������
    		depth_logging = str(array_map[k][11])#�������� ������ ��������� ���
    		value_logging = str(array_map[k][12])#�������� ���
    		abs_strat = array_map[k][13]#������� abs �������������
    		abs_litostrat_roof = str(array_map[k][14])#������� ������ �������
                abs_litostrat_sole = str(array_map[k][15])#������� ������� �������
    		stratigraphy = str(array_map[k][16])#������������
    		lithology = str(array_map[k][17])#���������
    		average_gis = array_map[k][18]#������� �������� ���
    		quant_gis = array_map[k][19]#���������� ���������
    		msd_atr = array_map[k][20]#�������������������� ����������
    		standard_error_atr = array_map[k][21]#������ ��������
    		mediana = array_map[k][22]#�������
    		feat = rows.NewRow()#����� ������
    		pnt.id = k #ID
    		x = round(float(array_map[k][2]),3)
    		y = round(float(array_map[k][3]),3)
    		pnt.x = x
    		pnt.y = y
    		feat.shape = pnt
    		feat.SetValue("id_��������",id_well)#�������� �������� � ������������ �����
    		feat.SetValue("��������_��������",name_well)
		feat.SetValue("x",x)
		feat.SetValue("y",y)
    		feat.SetValue("z",z)
    		feat.SetValue("�������_��������",depth_well)
    		feat.SetValue("��������_�������",name_obj)
    		feat.SetValue("��������_�������",name_uchastok)
                feat.SetValue("���_�����_����������",name_type_tn)
                feat.SetValue("���������_����������������",name_status_doc)
                feat.SetValue("���_����������������",name_type_doc)
    		feat.SetValue("�������_���������_���",depth_logging)
    		feat.SetValue("��������_���",value_logging)
    		feat.SetValue("���������_��������_�������������",abs_strat)
    		feat.SetValue("�������_������_�������",abs_litostrat_roof)
    		feat.SetValue("�������_�������_�������",abs_litostrat_sole)
    		feat.SetValue("������������",stratigraphy)
    		feat.SetValue("���������",lithology)
    		feat.SetValue("�������_��������_���",average_gis)
    		feat.SetValue("����������_���������",quant_gis)
    		feat.SetValue("��������������������_����������",msd_atr)
    		feat.SetValue("������_��������",standard_error_atr)
    		feat.SetValue("�������",mediana)
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
        #�������� ������� ���������� �������� ��� � ������� ����������������� ������� ��� ���� �������
        arr_select = select_values(arr_map_points)     
        if(len(arr_select)==0):#���� ������ ����
                raise SystemExit(0)
        ARCPY.AddMessage("���������� ������� � ����������� ��� �� �������� ������� ��� ������ � ���������������� ����: "+str(len(arr_select)))
        Spat_ref = gp.describe(InpPFC).spatialreference
        write_layer(OutputFC,arr_select,Spat_ref)#�������� �������� ������ � ���������������� ����
        


