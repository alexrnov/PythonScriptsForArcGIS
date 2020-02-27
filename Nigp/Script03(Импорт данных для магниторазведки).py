#�������� ������������ ������ ������ �� excel-����� �������, � ������� ���������������� ����,
#������������ ������� �������� �������� ������ ��������� ��������� ���������������, �������� ������������� ��������� � �.�.
#������ ������������ ������� ����������������� ����:
#[OBJECTID*,Shape*,id_well(id ��������),name(�����/�������� ��������),x,y,z,name_object(�������� �������),
#name_area(�������� �������),depth(������� ��������� �������� ��������� ���������������),
#value_kmv(�������� ���������� ��������� ���������������),m(�������� ������������� ���������)]
#���������� ���������������� ���� ����� ���� �������������� � ��������� ������
#��� ������-�������������� ������������� ������������� � ���������������
import arcpy as ARCPY
import arcgisscripting
import os
try:
	from xlrd import open_workbook #������ ��� ������ � Excel-�������
except ImportError: #���� ������ xlrd ��� ������ � Excel-������� �� ������
        ARCPY.AddMessage("������ xlrd ��� ������ � Excel-������� �� ������, ��������� �����������")
        raise SystemExit(1)#��������� ������ ���������
#������� � �������� �������� ��������� ����� excel-����, � ���������� ������ ������ excel-�����.
#����������� ��� ������ ������������� ���������������
def read_excel_list_magn(excel_file,List_point,List_strat_lit,List_code_strat,List_code_lit,List_GIS):#������� ���������� ����� 
        def check_unicode(text):#������� ����������� ��������� ������ � utf-8 
                if(type(text)==unicode):text = text.encode('utf-8')
                return text
        def get_object_list(L_name_array,list_name):#���� ���� � �������� ������ � ������ ������ excel-�����
                k=0
                list_object = "null"
                while(k<len(L_name_array)):#������� ������ excel-�����
                        #���� ���� � ���������, ������� ������� �� ������� ���������� ������
                        if(str(L_name_array[k])==str(list_name.encode('utf-8'))):
                                list_object = excel_book.sheet_by_name(list_name)
                                break 
                        k+=1
                return list_object #������� ������ �����
	excel_book = open_workbook(excel_file,formatting_info=True)#�������� ������ ����� excel
	array_list = excel_book.sheet_names()#�������� ������ �������� ������ excel
        k=0
        name_list_array=[]#������ �������� ������ � ��������� utf-8
        while (k<len(array_list)):#������� �������� excel-������
                name_list_array.append(check_unicode(array_list[k]))#�������� � ������ ������� �������� � ��������� utf-8
                k+=1
        #�������� ������� ������
        L_point_obj = get_object_list(name_list_array,List_point)#���� ����� ����������
        L_strat_lit_obj = get_object_list(name_list_array,List_strat_lit)#���� ������������ ���������
        L_code_strat_obj = get_object_list(name_list_array,List_code_strat)#���� ���� ������������
        L_code_lit_obj = get_object_list(name_list_array,List_code_lit)#���� ���� ���������
        L_GIS_obj = get_object_list(name_list_array,List_GIS)#���� ���
        return L_point_obj, L_strat_lit_obj, L_code_strat_obj, L_code_lit_obj, L_GIS_obj
#������� ���������, ��� �� ������� excel-������ �������, ���� ��� - ��������� ������ ���������
def verify_lists(d_point, d_strat_lit, d_code_strat, d_code_lit, d_GIS):
        if((d_point=="null")or(d_strat_lit=="null")or(d_code_strat=="null")or(d_code_lit=="null")or \
        (d_GIS=="null")): #��������� ��� �� ������� excel-������ �������
                ARCPY.AddMessage("��������� �� ��� excel-�����, ��������� �����������")#��������� ������ ���������
                raise SystemExit(1)#��������� ������ ���������
#������� � �������� ��������� ����� ������ ����� Excel-�����, ����������� ������ ����� ����������
#� ���������� ������ [id ��������, �����/����� ��������, x, y, z, ������, �������]
def read_points_wells(list_excel):
	title=True#���������� ������������� �������� �� ������ ������ ���������
	def_wells_array=[]#������ �������
	for rownum in range(list_excel.nrows):#������� ���� ����� �����
		if title: 
			title=False
			continue #���������� ������ "�����" � ����������� ���������
		row_array=[]#������ �����
		current_row=list_excel.row_values(rownum)#������� ������ �����
		name_object = "no_data"#�������� �������
		name_area = "no_data"#�������� �������
		id_well="no_data" #������������� ��������
		x="no_data"#���������� x
		y="no_data"#���������� y
		z="no_data"#�������� �������
		name_well="no_data" #�����/����� ��������
		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
			if (type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
			if ((j==1)and(cell<>'')):name_object=cell
			if ((j==2)and(cell<>'')):name_area=cell
			if ((j==5)and(cell<>'')):id_well=cell
			if ((j==12)and(cell<>'')):x=cell
			if ((j==13)and(cell<>'')):y=cell
			if ((j==14)and(cell<>'')):z=cell
			if ((j==25)and(cell<>'')):name_well=cell
		def_wells_array.append([id_well,name_well,x,y,z,name_object,name_area])#�������� ������� �������� � ������ �������
	return def_wells_array #������� ������ ����� ���������� � ������� [id ��������, �����/����� ��������, x, y, z, ������, �������]
#������� � �������� ��������� ����� ������ ����� Excel-�����, ����������� ���� ������������
#� ���������� ������ [��� ������������, ����������� ������������]
def read_strat_wells(list_excel):
	title=True#���������� ������������� �������� �� ������ ������ ���������
	def_array_strat=[]#������ � ������ ������������
	for rownum in range(list_excel.nrows):#������� ���� ����� �����
		if title: 
			title=False
			continue #���������� ������ "�����" � ����������� ���������
		row_array=[]#������ �����
		current_row=list_excel.row_values(rownum)#������� ������ �����
		def_code_strat="no_data"#��� ������������
		def_name_strat="no_data" #�������� ������������
		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
			if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
			if((j==0)and(cell<>'')):def_code_strat=cell #������ ������� � ����� ������������
			if((j==1)and(cell<>'')):def_name_strat=cell #������ ������� � ������������� ������������
		def_array_strat.append([def_code_strat,def_name_strat])#�������� ������� ������ � ������ ������� 
	return def_array_strat #���������� ������ � ������� [��� ������������, ����������� ������������]
#������� � �������� ������� ������ ����� ������ ����� ������������ � ������ � ��������� ������������������� �����
#� ���������� ������ � ������� [��� ������������, ������������ ������������, ��� �������������(P) ��� ��������� �����(V)]
def calc_type_rock(arr_strat, country):
	def_code_strat_new=[]#����� ������ ��� ������
	k = 0
	while (k<len(arr_strat)):#������� ������� � ������ ������������
		m=0
		b=False #���������� ���������� ���� �� ������ ������� �������������� �������, ���������������� ��������� �������
		while (m<len(country)):#������� ������� � ��������� <<���������>> ������������� ��������
			if(arr_strat[k][1].count(country[m])):#���� ���� ���������� ��������� ������
				#�������� � ������ ��������� ������
				def_code_strat_new.append([arr_strat[k][0],arr_strat[k][1],"V"])
				b=True #���� ������� ��������� ������
				break#����� �� ����������� �����
			m+=1
		if (b==False):#���� �� ���� ���������� � ���������� ��������
			#�������� � ������ ������������� ������
			def_code_strat_new.append([arr_strat[k][0],arr_strat[k][1],"P"])
		k+=1
	#������� ������ � ������� [��� ������������, ����������� ������������, ��� ��������� (������������� ��� ���������)]
	return def_code_strat_new
#������� � �������� ��������� ����� ������ ����� Excel-�����, ����������� ���� ���������
#� ���������� ������ [��� ���������, ����������� ���������]
def read_lit_wells(list_excel):
	title=True#���������� ������������� �������� �� ������ ������ ���������
	def_array_lit=[]#������ � ������ ������������
	for rownum in range(list_excel.nrows):#������� ���� ����� �����
		if title: 
			title=False
			continue #���������� ������ "�����" � ����������� ���������
		row_array=[]#������ �����
		current_row=list_excel.row_values(rownum)#������� ������ �����
		def_code_lit="no_data"#��� ���������
		def_name_lit = "no_data"#�������� ���������
		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
			if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
			if((j==0)and(cell<>'')):def_code_lit=cell #������ ������� � ����� ���������
			if((j==1)and(cell<>'')):def_name_lit=cell #������ ������� � ������������� ���������
		def_array_lit.append([def_code_lit,def_name_lit])#�������� ������� ������ � ������ ������� 
	return def_array_lit #���������� ������ � ������� [��� ���������, ����������� ���������]
#������� � �������� �������� ��������� ����� ������ ����� Excel-�����, ����������� �������� � �������� � ������� �������
#� ���������� ������ [point_id,������,�������,��� ���������,��� ������,�������� ������,�����/����� ��������]
def read_strat_lit_wells(list_excel):
	title=True#���������� ������������� �������� �� ������ ������ ���������
	def_arr_strat_lit_wells=[] #������ �� ���������� �������� �������
	for rownum in range(list_excel.nrows):#������� ���� ����� �����
		if title: 
			title=False
			continue #���������� ������ "�����" � ����������� ���������
		current_row=list_excel.row_values(rownum)#������� ������ �����
		point_id="no_data"
		roof="no_data" #������
		sole="no_data" #�������
		cod_old="no_data" #��� ��������
		cod_rock="no_data" #��� ������
		about_rock="no_data" #�������� ������
		name_point="no_data" #�����/����� ��������
		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
			if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
			if((j==0)and(cell<>'')):point_id=cell
			if((j==1)and(cell<>'')):roof=cell
			if((j==2)and(cell<>'')):sole=cell
			if((j==3)and(cell<>'')):cod_old=cell
			if((j==4)and(cell<>'')):cod_rock=cell
			if((j==5)and(cell<>'')):about_rock=cell
			if((j==6)and(cell<>'')):name_point=cell
		if((point_id<>"no_data")or(roof<>"no_data")or(sole<>"no_data")or(cod_old<>"no_data")or \
                   (cod_rock<>"no_data")or(about_rock<>"no_data")or(name_point<>"no_data")): #���� �� ������� ������ ������
			def_arr_strat_lit_wells.append([point_id,roof,sole,cod_old,cod_rock,about_rock,name_point])
	#������� ������ � ������� [point_id,������,�������,��� ���������,��� ������,�������� ������,�����/����� ��������]
	return def_arr_strat_lit_wells
#������� � �������� �������� ��������� ����� ������, ���������� ���������� � ������� 
#� ���������� ������ � ������� [��� �����, ������ �������, ���� ������������]
def single_width_blocking(arr_form):
	def_arr_calc_width=[]#����� ������ ��� ������ ������
	well_first=arr_form[0][0]#������ �������� �������������
	k=0
	roof=""#�������
	code_old=""#��� ������
	while(k<len(arr_form)):
		if (arr_form[k][0]==well_first):#���� ������� ������������� ����� �����������
        		roof=roof+str(arr_form[k][1])+"/" #�������� abs ������ �������� ������
			code_old=code_old+str(arr_form[k][3])+"/" #�������� ��� �������� �������� ������
    		else: #���� ������ ������� � ����� ��������
			current_array=[well_first,roof,code_old] #��������� ������ ���������� ��������
			def_arr_calc_width.append(current_array)# � �������� �� � ������ ������� �������
			well_first=arr_form[k][0]#���������� ������� �������������
			roof=str(arr_form[k][1])+"/" # ������ ����� ������ abs ������ �������� ������
			code_old=str(arr_form[k][3])+"/" # ������ ����� ������ c ������ �������� �������� ������
    		k+=1
    		if (k==len(arr_form)):#�������� � ����� ������ ������ � ��������� ��������
        		current_array=[well_first,roof,code_old]
			def_arr_calc_width.append(current_array)
	#������� ������ � ������� [��� �����, ��� ������ �������, ���� ������������]
	return def_arr_calc_width
#������� ����� � �������� ���������� ������ [��� �����, ������ �������, ���� ������������] 
#� ������ [��� ������������, ������������ ������������, ��� �������������(P) ��� ��������� �����(V)]
#� ���������� ������ [��� ��������, �������� �������������]
def calc_width_blocking(arr_width, code_strat):
	def_arr_width_blocking=[]
	k=0
	while(k<len(arr_width)):#������� ������� ������� ������� [��� �����, ������ �������, ���� ������������]
		find_fraction=False#���������� ���������� ���������� ������� �� ������������������� ������
		str_roof = str(arr_width[k][1])#�������� ������ abs ������� ������� ��������
		arr_roof = str_roof.split("/")#�������� ������ abs ������� ��� ������� ��������
		str_code = str(arr_width[k][2])#�������� ������ ����� ������������ ��� ������� ��������
		arr_code = str_code.split("/")#�������� ������ ����� ������������ ��� ������� ��������
		m=0
		##������� ����� ������������ ������� ��������, �� ������������ ��������� ������� ������� ����� ������������ ���������
		#�� �������� ������ - ��� �������� ������ '1/2/3/' ���������� ������ ['1','2','3',''] 
		while(m<len(arr_code)-1):
			n=0
			if(find_fraction):break#����� �� ����� ����� ���� ������������������� ������ ��� ���� ����������. ������� � ��������� ��������
			#������� ���� ����� ������������ ������� [��� ������������, ������������ ������������, ��� �������������(P) ��� ��������� �����(V)]
			while(n<len(code_strat)):
				#���� ��� ������������ �������� ������ ��������� � ����� �� ������� ����� �����(�������������/���������)
				cod1=float(arr_code[m])#������������� � ����� ��� ������������ �� ������� arr_width
				cod2=float(code_strat[n][0])#������������� � ����� ��� �� ������� code_strat 
				if(cod1==cod2):#���� ���� �������
					if(code_strat[n][2]=="V"):#���� ����� ���� ������������� ��������� ������
						current_array=[arr_width[k][0],arr_roof[m]]#������ ��� ������� �������� [��� ��������, �������� �������������]
						def_arr_width_blocking.append(current_array)#�������� ������� �������� � ����� ������
						find_fraction=True#������� � ��������� ��������
						break#����� �� ����������� �����
				n+=1
			m+=1
			#���� ��� ������ ������� �������� ����������������, �� ��� ���� �� ���������� ��������� ������
			#������� ��� ���� �������� �� �����
			if((m==len(arr_code)-1)and(not find_fraction)):
				ARCPY.AddMessage("�������� ������������� ��������� �� ����������. ��� �����: "+str(arr_width[k][0]))
		k+=1
	return def_arr_width_blocking #������� ������ [��� ��������, �������� �������������]
#������� � �������� ������� ���������� ����� ������ ����� Excel-�����, ����������� ������ ��������� ���,
#� ����� �������� ������� �� ���������� ��������� �������� ��������� ���������������
#� ���������� ������ [id ��������, ������� ���������, �������� ���������� ��������� ���������������]
def read_gis_wells(list_excel,method_kmv):
	title=True#���������� ������������� �������� �� ������ ������ ���������
	number_kmv=0#����� �������, � ������� ���������� �������� ��������� ������� ���
	def_kmv_arrays = []#������ ��� �������� �������� [id_��������, ������� ���������, �������� ��������� ������� ���]
	for rownum in range(list_excel.nrows):#������� ���� ����� �����
		if title:#���� ��������� ��������� �� ������ ������ �������, ��������������� "�����"
			title=False
        		current_row=list_excel.row_values(rownum)#������� ������ �����
			for j, cell in enumerate(current_row):#����������� ������ ��������� ��� ������ ������ ���
            			if(type(cell)==unicode):
					cell=cell.encode('utf-8')
                                #���� �������� ������� ��������� � ��������� ������� �� ���������� ��������� �������� ��������� ���������������
				if (cell==method_kmv):
					ARCPY.AddMessage('���� ������ �� �������� ���, ������� � excel: '+str(j+1))
					number_kmv=int(j)#����� �������, � ������� ���������� �������� ��������� ������� ���
                			break#����� �� ����������� �����
			continue #������ ����� �������� ������ �����
    		current_row=list_excel.row_values(rownum)#������� ������ �����
    		id_well="no_data"#������������� ��������
    		dep="no_data"#������� ���������� ��������
    		kmv="no_data" #���������� �������� ��������� ���������������
    		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
        		if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
        		if((j==0)and(cell<>'')):id_well=cell #������������� ��������
			if((j==1)and(cell<>'')):dep=cell #������� ���������
			if((j==number_kmv)and(j<>0)):kmv=cell #���������� �������� ��� 
    		def_kmv_array=[id_well,dep,kmv]#������, ���������� ������ �� �������� ���������
    		def_kmv_arrays.append(def_kmv_array)#�������� ������� ��������� � ����� ������ 
	return def_kmv_arrays #[id ��������, ������� ���������, �������� ���������� ��������� ���������������]
#������� � �������� ��������� ����� ������ ��������� � ����� ��� �.�. ������ ���� ����� �����
#� ���������� ������ � ������� [id ��������, �������/�������/�������..., ��������/��������/��������...]
def single_kmv_data(kmv_arr): 
	well_first=kmv_arr[0][0]#������ �������� �������������
	h=""#������ ��� ������ �������� ������
	v=""#������ ��� ������ �������� ��������� ���������������
	#������ ��� ������ �������� ������� � ������� � ������� � ��������� ��������������� [id, glubina, vospriimchivost]
	def_single_kmv_array=[]
	k=0
	while k<len(kmv_arr):# ������� �������, ���������� �� excel-�����
    		if (kmv_arr[k][0]==well_first):#���� ������� ������������� ����� �����������
        		h=h+str(kmv_arr[k][1])+"/" #������, ���������� �������� ������ ��������� ��������� ���������������
			v=v+str(kmv_arr[k][2])+"/" #������, ���������� �������� ��������� ���������������
    		else: #���� ������ ������� � ����� ��������
			current_array=[well_first,h,v] #��������� ������ ���������� ��������
			def_single_kmv_array.append(current_array)# � �������� �� � ������ ������� �������
			well_first=kmv_arr[k][0]#���������� ������� �������������
			h=str(kmv_arr[k][1])+"/" # ������ ����� ������ ������
			v=str(kmv_arr[k][2])+"/" # ������ ����� ������ �������� ���������������
    		k=k+1
    		if (k==len(kmv_arr)):#�������� � ����� ������ ������ � ��������� ��������
        		current_array=[well_first,h,v]
			def_single_kmv_array.append(current_array)
	return def_single_kmv_array #������� ������ � ������� [id ��������, depth/depth..., value/value...]
#������� � �������� ������� ���������� ����� ������ ����� � ������������ � ������ �����������
#� ������ � ����������� ���������� ��������� ��������������� 
#���������� ������������ ������ � ������� [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...]
def united_two_arrays(arr_wells, arr_kmv):
	#����� ������, ������� ����������� �� ������ �������� ����� ����� ���������� � ����� � ���
	#� ������� ����� ������� �������� � �����������, ������ ��������, 
	#��������������, �� ������� - ������ ������� � �������� ���������
	def_united_array=[]
	k=0
	while(k<len(arr_wells)):#������� ��������� ������� ����� (������� �����)
    		m=0
    		while(m<len(arr_kmv)):#������� ��������� ������� ����� (������� �� ���������� ��������� ���������������) 
        		if(arr_wells[k][0]==arr_kmv[m][0]):#���� �������������� �������
				#������ ��� ������ ����� � ��� �� ��������,
				#������� ������������ ���� ����� ������
				curr_array = [arr_wells[k][0],arr_wells[k][1],arr_wells[k][2],arr_wells[k][3],arr_wells[k][4], \
                                arr_wells[k][5],arr_wells[k][6],arr_kmv[m][1],arr_kmv[m][2]]
            			def_united_array.append(curr_array) #�������� � ����� ������ �����
	    			break
        		m=m+1
    		k=k+1	
	return def_united_array #������� ������ [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...]
#������� � �������� ������� ���������� ����� ������ [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...]
#� ������ [��� ��������, �������� �������������]
#���������� ������ [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...,�������� ������������� ���������]
def united_width_blocking(un_array, block_array):
	def_united_final_arr=[]#����� ������ ��� ������ ������
	k=0
	curr_array=[]
	while(k<len(un_array)):#������� ��������� ������� [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...]
		curr_array=[]
		cod1=float(un_array[k][0])#��� ����� ���������� ������� �������
		m=0
		while(m<len(block_array)):#������� ��������� ������� [��� ��������, �������� �������������]
			cod2=float(block_array[m][0])#��� ����� ���������� ������� �������
			if(cod1==cod2):#���� ������� ���������� �����
				#���������� ���������� ��������
				curr_array=[un_array[k][0],un_array[k][1],un_array[k][2],un_array[k][3],un_array[k][4],un_array[k][5],un_array[k][6],un_array[k][7],un_array[k][8],block_array[m][1]]
				def_united_final_arr.append(curr_array)#�������� � ����� ������ ���������� � ������� ��������
				break
			m+=1
		k+=1
	#���������� ������ [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...,�������� ������������� ���������]
	return def_united_final_arr
#������� � �������� ������� ���������� ����� ���������� ��������� ����������������� ����, ������� ���������, � �������� ������
#� ������� [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...]
#� ������� ���������������� ���� � ���������������� ����������
def write_to_layer(lay, C_System, array_map):
	fc_path, fc_name = os.path.split(lay)#��������� ���������� � ��� �����
	gp=arcgisscripting.create(9.3)#������� ������ �������������� ������ 9.3
	gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",C_System)#������� �������� ����
	gp.AddField(OutputFC,"id_well","STRING",15,15)#������� id �������
	gp.AddField(OutputFC,"name","STRING",15,15)#������� �������� �������
	gp.AddField(OutputFC,"x","FLOAT",15,15)#�������� ������� x
	gp.AddField(OutputFC,"y","FLOAT",15,15)#�������� ������� y
	gp.AddField(OutputFC,"z","FLOAT",15,15)#�������� ������� z
	gp.AddField(OutputFC,"name_object","STRING","","",100)#������� - �������� �������
	gp.AddField(OutputFC,"name_area","STRING","","",50)#������� - �������� �������
	gp.AddField(OutputFC,"depth","STRING","","",1000)#������� �������� ������ ��������� ��������� ��������������
	gp.AddField(OutputFC,"value_kmv","STRING","","",1000)#������� ��� �������� �����
	gp.AddField(OutputFC,"m","FLOAT",15,15)#�������� ������� �������� ������������� ���������
	rows = gp.InsertCursor(OutputFC)
	pnt = gp.CreateObject("Point")
	k=0
	while(k<len(array_map)): #�������� ��������� ����
    		id_well = float(array_map[k][0])
    		name_well = str(array_map[k][1])
    		z = float(array_map[k][4])
    		name_obj = str(array_map[k][5])
    		name_uchastok = str(array_map[k][6])
    		depth = str(array_map[k][7])
    		value_kmv = str(array_map[k][8])
		value_width = float(array_map[k][9])
    		feat = rows.NewRow()#����� ������
    		pnt.id = k #ID
    		pnt.x = float(array_map[k][3])
    		pnt.y = float(array_map[k][2])
    		feat.shape = pnt
    		feat.SetValue("id_well",id_well)#�������� �������� � ������������ �����
    		feat.SetValue("name",name_well)
		feat.SetValue("x",float(array_map[k][3]))
		feat.SetValue("y",float(array_map[k][2]))
    		feat.SetValue("z",z)
    		feat.SetValue("name_object",name_obj)
    		feat.SetValue("name_area",name_uchastok)
    		feat.SetValue("depth",depth)
    		feat.SetValue("value_kmv",value_kmv)
		feat.SetValue("m",value_width)
    		rows.InsertRow(feat)#�������� ������ � ������� ����
    		k=k+1
if __name__== '__main__':#���� ������ �������, � �� ������������
        #-------------------------------������� ������------------------------------------
        InputExcelFile = ARCPY.GetParameterAsText(0)#������� excel-����
        List_point = ARCPY.GetParameterAsText(1)#�������� ����� � ������� ����������
        List_strat_lit = ARCPY.GetParameterAsText(2)#�������� ����� � ��������� ������������ � ���������
        List_code_strat = ARCPY.GetParameterAsText(3)#�������� ����� � ������ ������������
        List_code_lit = ARCPY.GetParameterAsText(4)#�������� ����� � ������ ���������
        List_GIS = ARCPY.GetParameterAsText(5)#�������� ����� � ���
        code_method_kmv = ARCPY.GetParameterAsText(6).encode('utf-8')#�������� ������� � ������� �������� ��������� ���������������
        code_contain = ARCPY.GetParameterAsText(7).encode('utf-8')#������� ��������� ������������������� ���������
        Coordinate_System = ARCPY.GetParameterAsText(8)#������������ ������� ��������� ��������� ����    
        OutputFC = ARCPY.GetParameterAsText(9)#�������� �������� ���� �������
        Full_report = ARCPY.GetParameterAsText(10)#�������� ���������� �������� �� ������ �����
        #---------------------------------------------------------------------------------
        #�������� ������� excel-������
        obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_GIS = \
        read_excel_list_magn(InputExcelFile,List_point,List_strat_lit,List_code_strat,List_code_lit,List_GIS)
        #��������� ��� �� excel-����� � �������, ���������� ��� ������� ���������, �������
        verify_lists(obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_GIS)
        arr_country_rock = code_contain.split(";")#�������� ������ �������� ��������� ������������������� ���������
        wells_array = read_points_wells(obj_point)#�������� ������ ����� �� ����� exc�l-����� � ������� ����������
        if(len(wells_array)==0):#���� ������ ����� ���������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� ����� ���������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        ARCPY.AddMessage("����� ���������� ����� ����������, ����������� �� Excel-�����: "+str(len(wells_array)))
        if(Full_report == "true"):#�������� ������ �����
                if(len(wells_array)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(wells_array[k][0])+" "+str(wells_array[k][1])+" "+ \
                                str(wells_array[k][2])+" "+str(wells_array[k][3])+" "+str(wells_array[k][4])+" "+ \
                                str(wells_array[k][5])+" "+str(wells_array[k][6]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_strat = read_strat_wells(obj_code_strat)#�������� ������ � ������ ������������
        if(len(code_strat)==0):#���� ������ ����� ������������ ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� ����� ������������ ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        ARCPY.AddMessage("���������� ����� ������������: "+str(len(code_strat)))
        if(Full_report == "true"):#�������� ������ �����
                if(len(code_strat)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_strat[k][0])+" "+str(code_strat[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_strat_new = calc_type_rock(code_strat, arr_country_rock)#�������� ������ � ������ ������������ � ������ �����(����������/��������������)
        ARCPY.AddMessage("���������� ����� ������������ � ������ �����: "+str(len(code_strat_new)))
        if(Full_report == "true"):#�������� ������ �����
                if(len(code_strat_new)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_strat_new[k][0])+" "+str(code_strat_new[k][1])+" "+code_strat_new[k][2])
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_lit = read_lit_wells(obj_code_lit)#�������� ������ � ������ ���������
        ARCPY.AddMessage("���������� ����� ���������: "+str(len(code_lit)))
        if(len(code_lit)==0):#���� ������ ����� ��������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� ����� ��������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                if(len(code_lit)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_lit[k][0])+" "+str(code_lit[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        arr_formation = read_strat_lit_wells(obj_strat_lit) #�������� ������ � ������� ������� �������
        ARCPY.AddMessage("���������� �������, ��������� �� �����: "+str(len(arr_formation)))
        if(len(arr_formation)==0):#���� ������ ��������� ������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���������� � ��������� �����������, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                if(len(arr_formation)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(arr_formation[k][0])+" "+str(arr_formation[k][1])+" "+ \
                                str(arr_formation[k][2])+" "+str(arr_formation[k][3])+" "+str(arr_formation[k][4])+ \
                                " "+str(arr_formation[k][5])+" "+str(arr_formation[k][6]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        arr_single_width = single_width_blocking(arr_formation) #�������� ������ � ���������������� �������� ��� ������ ��������
        ARCPY.AddMessage("���������� ������� � ����������� � �������� �������: "+str(len(arr_single_width)))
        if(len(arr_single_width)==0):#���� ������ � ���������������� �������� ��� ������ �������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("������ � ��������������� �������� ��� ������ �������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                ARCPY.AddMessage(str(arr_single_width[0][0])+" "+str(arr_single_width[0][1])+" "+str(arr_single_width[0][2]))
                ARCPY.AddMessage("...........")
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        arr_width_blocking=calc_width_blocking(arr_single_width, code_strat_new)#��������� �������� ������������� ���������
        ARCPY.AddMessage("���������� ������� � ����������� ��������� ������������� ���������: "+str(len(arr_width_blocking)))
        if(len(arr_width_blocking)==0):#���� ��� ������� � ����������� ��������� ������������� ��������� - ��������� ������ ���������
                ARCPY.AddMessage("������ ������� � ����������� ��������� ������������� ��������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                k=0
                while(k<len(arr_width_blocking)):
                        ARCPY.AddMessage(str(arr_width_blocking[k][0])+" "+str(arr_width_blocking[k][1]))
                        k+=1
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        kmv_arrays = read_gis_wells(obj_GIS,code_method_kmv)#�������� ������ ���������� �������� ��� �� ����� ��� Excel-�����
        if(len(kmv_arrays)==0):#���� ��� ������ �� �������� ��������� ��������������� - ��������� ������ ���������
                ARCPY.AddMessage("��� ������ �� �������� ��������� ���������������, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                if(len(kmv_arrays)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(kmv_arrays[k][0])+" "+str(kmv_arrays[k][1])+" "+str(kmv_arrays[k][2]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        single_kmv_array = single_kmv_data(kmv_arrays)#�������� ������ � ������� [id �����,�������/�������..,��������/��������..]
        ARCPY.AddMessage("���������� ��������� ��������������� ������� �� ���������� ��������� ���������������: "+ str(len(single_kmv_array)))        
        if(len(single_kmv_array)==0):#���� ������ �� ���������� ��� ���� - ��������� ������ ���������
                ARCPY.AddMessage("������ �� ���������� ��������� ��������������� �� �����������, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                ARCPY.AddMessage(str(single_kmv_array[0][0])+" "+str(single_kmv_array[0][1])+" "+str(single_kmv_array[0][2]))
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #�������� ������������ ������ �� ���������� ��������� ���������������, ������������, ��������� � �.�.
        united_array = united_two_arrays(wells_array, single_kmv_array)
        ARCPY.AddMessage("���������� ��������� ������������� ������� ������ �������� � ���: "+str(len(united_array)))
        if(len(united_array)==0):#���� ������������ ������ ����
                ARCPY.AddMessage("������ ������ ������� � ��� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                ARCPY.AddMessage(str(united_array[0][0])+" "+str(united_array[0][1])+" "+str(united_array[0][2])+ \
                " "+str(united_array[0][3])+" "+str(united_array[0][4])+" "+str(united_array[0][5])+ \
                " "+str(united_array[0][6])+" "+str(united_array[0][7])+" "+str(united_array[0][8]))
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #�������� ������������ ������ [id �����,�����/�����,x,y,z,������,�������,�������/�������...,���/���...,�������� ������������� ���������]
        final_array=united_width_blocking(united_array, arr_width_blocking)
        ARCPY.AddMessage("���������� ��������� ������������� ������� ��� ������ � ���������������� ���� � ��������� ������������� ���������: "+str(len(final_array)))
        if(len(final_array)==0):#���� ������ ��� ������ � ���������������� ���� ����
                ARCPY.AddMessage("������ ��� ������ � ���������������� ���� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                ARCPY.AddMessage(str(final_array[0][0])+" "+str(final_array[0][1])+" "+str(final_array[0][2])+ \
                " "+str(final_array[0][3])+" "+str(final_array[0][4])+" "+str(final_array[0][5])+" "+ \
                str(final_array[0][6])+" "+str(final_array[0][7])+" "+str(final_array[0][8])+" "+str(final_array[0][9]))
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #������� ���������� ������ � ���������������� ���� 
        write_to_layer(OutputFC, Coordinate_System, final_array)
        del wells_array,code_strat,code_strat_new,code_lit,arr_formation,arr_single_width,arr_width_blocking, \
            kmv_arrays,single_kmv_array,united_array,final_array
