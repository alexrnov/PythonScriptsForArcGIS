#������ � �������� �������� ��������� ����� ���������������� ���� � ������� ��������� ���������������,
#������ ������������� �������������� ���� � ��� ��������� ���������������
#� ���������� ����������� ������ � ��������� ����, ������� ���������� ��������� GravMag3D
#��� ������-�������������� ������������� ��� ���������������
import arcpy, sys, arcgisscripting, os
#������� ���������� ��������� �� �������� ������� �������������� ����
#s_kimb - �������� ������� �������������� ���� � ���������� ����� �������� ������� �������������� ���� 
def true_size(s_kimb):
        s = False#���������� ���������� ���������� ��������� �� �������� ������� �������������� ����
        int1 = 0 #������ �������������� ����
        if(s_kimb.count("x")):#���� ������ x ������
                size_kimb_arr = s_kimb.split("x")#��������� ��������, �������� 100x100 ����� [100,100]
                if(size_kimb_arr[0]==size_kimb_arr[1]):#���� ��� ����� ��������� (�������� 100x100)
                        try:#���� ��� ����� �������� ������ ������� ����
                                int1 = int(size_kimb_arr[0])#������������� ����� � �����
                                if ((int1>0)and(int1<3000)):s = True#���� ������ ������ 0 � ������ 3000                
                        except ValueError:pass
        return s, int1
#������� � �������� �������� ��������� ����� �������� ���� � ������� ����������,
#������� �������� ������ �� ��������� ��������������� ��� � �������� ������������� ���������
#� ���������� ������ � �������
# [id ��������,�����/�����,x,y,abs ����� �������,�������� �������,�������� �������,������� ��������� ��������� ���������������,�������� ������������� ���������]
def read_points(layer):
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
        name_cell = check_unicode(row.GetValue("name"))#�������� ��������
        x = point.x
        y = point.y
        z = check_unicode(row.GetValue("z"))#�������� ������� ����� �������
	z = float(z)
	z = round(z,2)#��������� �� ���� ������ ����� ������� 
        name_object = check_unicode(row.GetValue("name_object"))#�������� �������
	name_uchastok = check_unicode(row.GetValue("name_area"))#�������� �������
	depth = check_unicode(row.GetValue("depth"))#������� ��������� ��������� ���������������
	value_kmv = check_unicode(row.GetValue("value_kmv"))#�������� ��������� ��������� ���������������
        block_m = check_unicode(row.GetValue("m"))#�������� ������������� ���������
	block_m = float(block_m)
	block_m = round(block_m,1)#��������� �� ������ ����� ����� �������
	arr_wells.append([id_well,name_cell,x,y,z,name_object,name_uchastok,depth,value_kmv,block_m])
        row = rows.next()
    return arr_wells #������� ������ [id_well,name_cell,x,y,z,name_object,name_uchastok,depth,value_kmv,block_m]
#������� � �������� ������� ���������� ����� �������� ���������� ����� ��� ������ ������ �� ������ ����������
#� ��� ������, ���������� ���������� �� ������ ����������
#� ���������� ��������� ����
def write_txt_file(text_file,array_points,size_k,mv_k,ac):
	if not(text_file==""):#���� ������ ���� ��� ���������� ����� 
    		fc_path, fc_name = os.path.split(text_file)
    		if(fc_name<>"magn.txt"):#���� ������ ����� �������� magn.txt - ��� ���������� ��� ��������� GravMag3D
                        gp.AddMessage("������������ �������� ���������� �����, �� ������ ����� ��� magn.txt")
                        raise SystemExit(1)#��������� ������ ���������
    		text_file_open = open(text_file,"w")
    		k=0
    		text_file_open.write("size_kimberlite="+str(size_k)+"\n")#������ �������������� ����
    		text_file_open.write("magnetizability_of_kimberlite="+str(mv_k)+"\n")#��������� ��������������� �������������� ����
    		text_file_open.write("shooting_accuracy="+str(ac)+"\n")#�������� ������������������ ������
    		while (k<len(array_points)):#�������� ������ � ��������� ����
        		text_file_open.write(str(array_points[k][0])+"$"+str(array_points[k][1])+"$"+ \
                        str(array_points[k][3])+"$"+str(array_points[k][2])+"$"+str(array_points[k][4])+ \
                        "$"+str(array_points[k][5])+"$"+str(array_points[k][6])+"$"+str(array_points[k][7])+ \
                        "$"+str(array_points[k][8])+"$"+str(array_points[k][9])+"\n")
        		k=k+1
	text_file_open.close()
if __name__ == '__main__':#���� ������ ������� � �� ������������
        gp = arcgisscripting.create(9.3)
        #-------------------------------������� ������------------------------------------
        InpPFC = arcpy.GetParameterAsText(0)#������� �������� ���� �������
        size_kimb = arcpy.GetParameterAsText(1)#������� �������� ������� �������������� ����
        mv_kimb = arcpy.GetParameterAsText(2)#������� �������� ��������� ��������������� �������������� ����
        accuracy = arcpy.GetParameterAsText(3)#������� �������� �������� ������������������ ������
        output_text_file = gp.GetParameterAsText(4)#�������� ��������� ���� � ������� ��������� ���������������
        #---------------------------------------------------------------------------------
        bool_size, value_size = true_size(size_kimb) 
        if (not(bool_size)or(value_size==0)):#���� ������ ������� ������������� ��� �����������
                gp.AddMessage("������������ �������� � ��������� ���� ����� '������� �������� ������� �������������� ����'")
                raise SystemExit(1)#��������� ������ ���������
        try:
                float_mv_kimb = float(mv_kimb)#������������� � ������������ ���
                if not((float_mv_kimb>0)and(float_mv_kimb<1000)):#���� �������� ��������� ��������������� ������ 0 � ������ 1000 ��.��
                        gp.AddMessage("������������ �������� � ��������� ���� ����� '������� �������� ��������� ��������������� �������������� ����'")              
                        gp.AddMessage("��������� �����������")
                        raise SystemExit(1)#��������� ������ ���������
        except ValueError:#���� �������� �������� ���������� ������� 
                gp.AddMessage("������������ �������� � ��������� ���� ����� '������� �������� ��������� ��������������� �������������� ����'")
                raise SystemExit(1)#��������� ������ ���������
        try:
                float_accuracy = float(accuracy)
        except ValueError:#���� �������� �������� ���������� �������
                gp.AddMessage("����������� �������� � ��������� ���� ����� '�������� ������������������ ������'")
                gp.AddMessage("��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        arr_points = read_points(InpPFC)#������� ���� ��������� ����
        gp.AddMessage("���������� ����� �������� ��������� ����: "+str(len(arr_points)))
        write_txt_file(output_text_file, arr_points, value_size, mv_kimb, accuracy)#�������� ������ ���� ���������� � ��������� ����
