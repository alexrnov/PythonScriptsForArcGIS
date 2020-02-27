import arcpy as ARCPY
import arcgisscripting
import os
try:
	from xlrd import open_workbook #������ ��� ������ � Excel-�������
except ImportError: #���� ������ xlrd ��� ������ � Excel-������� �� ������
        ARCPY.AddMessage("������ xlrd ��� ������ � Excel-������� �� ������, ��������� �����������")
        raise SystemExit(1)#��������� ������ ���������
#������� � �������� �������� ��������� ����� excel-����, � ���������� ������ ������ excel-�����
def get_excel_list(excel_file,List_point,List_strat_lit,List_code_strat,List_code_lit,List_mineral,List_type_test,List_code_type_TN,List_code_status_document,List_code_status_test,List_code_type_document):
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
        L_mineral_obj = get_object_list(name_list_array,List_mineral)#���� � ������������
        L_type_test = get_object_list(name_list_array,List_type_test)#���� � ������������
        L_code_type_TN = get_object_list(name_list_array,List_code_type_TN)#���� � ������ ����� ��
        L_code_status_document = get_object_list(name_list_array,List_code_status_document)#���� � ������ ��������� ����������������
        L_code_status_test = get_object_list(name_list_array,List_code_status_test)#���� � ������ ��������� �����������
        L_code_type_document = get_object_list(name_list_array,List_code_type_document)#���� � ������ ����� ������������
        return  L_point_obj, L_strat_lit_obj, L_code_strat_obj, L_code_lit_obj, L_mineral_obj, L_type_test, \
        L_code_type_TN, L_code_status_document, L_code_status_test,L_code_type_document#������� ������ excel-������
#������� ���������, ��� �� ������� excel-������ �������, ���� ��� - ��������� ������ ���������
def verify_lists(d_point, d_strat_lit, d_code_strat, d_code_lit, d_mineral, d_code_test, d_type_TN, d_document,d_test,d_type_document):
        if((d_point=="null")or(d_strat_lit=="null")or(d_code_strat=="null")or(d_code_lit=="null")or \
        (d_mineral=="null")or(d_code_test=="null")or(d_type_TN=="null")or(d_document=="null")or \
        (d_test=="null")or(d_type_document=="null")): #��������� ��� �� ������� excel-������ �������
                ARCPY.AddMessage("��������� �� ��� excel-�����, ��������� �����������")#��������� ������ ���������
                raise SystemExit(1)#��������� ������ ���������
        else:ARCPY.AddMessage("��� ����� excel-����� �������")
#������� � �������� ��������� ����� ������ ����� Excel-�����, ����������� ������ ����� ����������
#� ���������� ������ [id ��������, �����/����� ��������, x, y, z, ������� ��������, ������, �������
#��� ������� ����������������,��� ��������� �����������,��� ������������]
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
		id_well="0" #������������� ��������
		depth_well="-1"#������� �������� 
		x="0"#���������� x
		y="0"#���������� y
		z="-1"#�������� �������
		name_well="no_data" #�����/����� ��������
		type_well = "������ �����������"#��� ���� ��
		status_doc = "������ �����������"#��� ������� ����������������
		status_test = "������ �����������"#��� ��������� �����������
		type_doc = "������ �����������"#��� ������������
		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
			if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
			if((j==1)and(cell<>'')):name_object=cell
			if((j==2)and(cell<>'')):name_area=cell
			if((j==5)and(cell<>'')):id_well=cell
			if((j==6)and(cell<>'')):type_well=cell
			if((j==7))and(cell<>''):depth_well=cell
			if((j==12)and(cell<>'')):x=cell
			if((j==13)and(cell<>'')):y=cell
			if((j==14)and(cell<>'')):z=cell
			if((j==19)and(cell<>'')):status_doc=cell
			if((j==22)and(cell<>'')):status_test=cell
			if((j==25)and(cell<>'')):name_well=cell
			if((j==26)and(cell<>'')):type_doc = cell
		def_wells_array.append([id_well,name_well,x,y,z,depth_well,name_object,name_area,str(type_well), \
                str(status_doc),str(status_test),str(type_doc)])#�������� ������� �������� � ������ �������
	return def_wells_array #������� ������ ����� ���������� � �������
        #[id ��������, �����/����� ��������, x, y, z, ������� ��������, ������, �������,��� ���� ��, ]
        #��� ������� ����������������,��� ��������� �����������,��� ������������]
#������� � �������� ��������� ����� ������ ����� Excel-�����, ����������� ���� � �� �����������
#� ���������� ������ [���, �����������]
def read_code_wells(list_excel):
	title=True#���������� ������������� �������� �� ������ ������ ���������
	def_array_code=[]#������ � ������ 
	for rownum in range(list_excel.nrows):#������� ���� ����� �����
		if title: 
			title=False
			continue #���������� ������ "�����" � ���������� ���������
		row_array=[]#������ �����
		current_row=list_excel.row_values(rownum)#������� ������ �����
		def_code="no_data"#��� 
		def_name = "no_data"#�����������
		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
			if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
			if((j==0)and(cell<>'')):def_code=cell #������ ������� � ����� 
			if((j==1)and(cell<>'')):def_name=cell #������ ������� � ������������
		def_array_code.append([def_code,def_name])#�������� ������� ������ � ������ ������� 
	return def_array_code #���������� ������ � ������� [���, �����������]
#������� ����� �� ����� ������ � ������� ���������� � �������
#[id ��������, �����/����� ��������, x, y, z, ������� ��������, ������, �������,��� ���� ��, ]
#��� ������� ����������������,��� ��������� �����������,��� ���� ������������]
#� �������� ��������� ������ �������� � ������ �� �� ����������� 
def replacement_code_lists(wells_array,def_type_tn,def_status_document,def_status_test,def_type_document):
        array_wells_replace = wells_array #������ � ������� ����������, � ������� ���� ���������� �� �������������� ��������
        k=0
        while(k<len(wells_array)):#������� ��������� ������� ����� (�� ����� ����� ����������)
                m=0
                while(m<len(def_type_tn)):#������� ����� ����� ����� ����������
                        #���� ��� ���� ����� �� ������� � �� ������ � ����� �� ����� � ������ ����� ��
                        if(str(wells_array[k][8])==str(def_type_tn[m][0])):
                                array_wells_replace[k][8]=def_type_tn[m][1]#�������� ��� �� �����������
                                break#����� �� ����������� �����
                        m+=1
                m=0
                while(m<len(def_status_document)):#������� ����� ����� ��������� ����������������
                        #���� ��� ���� ����� �� ������� � �� ������ � ����� �� ����� � ������ ��������� ����������������
                        if(str(wells_array[k][9])==str(def_status_document[m][0])):
                                array_wells_replace[k][9]=def_status_document[m][1]#�������� ��� �� �����������
                                break#����� �� ����������� �����
                        m+=1
                m=0
                while(m<len(def_status_test)):#������� ����� ����� ��������� �����������
                        #���� ��� ���� ����� �� ������� � �� ������ � ����� �� ����� � ������ ��������� �����������
                        if(str(wells_array[k][10])==str(def_status_test[m][0])):
                                array_wells_replace[k][10]=def_status_test[m][1]#�������� ��� �� �����������
                                break#����� �� ����������� �����
                        m+=1
                m=0
                while(m<len(def_type_document)):#������� ����� ����� ����������������
                        #���� ��� ���� ����� �� ������� � �� ������ � ����� �� ����� � ������ ����� ����������������
                        if(str(wells_array[k][11])==str(def_type_document[m][0])):
                                array_wells_replace[k][11]=def_type_document[m][1]#�������� ��� �� �����������
                                break#����� �� ����������� �����
                        m+=1
                k+=1
        #������� ������ � �������
        #[id ��������, �����/����� ��������, x, y, z, ������� ��������, ������, �������,����������� ���� ��, ]
        #����������� ������� ����������������,����������� ��������� �����������,����������� ���� ������������]
        return array_wells_replace
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
#� ���������� ������ � ������� [��� �����, ������ �������, ���� ������������, ���� ���������]
def single_width_blocking(arr_form):
	def_arr_calc_width=[]#����� ������ ��� ������ ������
	well_first=arr_form[0][0]#������ �������� �������������
	k=0
	roof=""#ads ������
	sole=""#abs �������
	code_old=""#��� ������
	code_lit=""#��� ���������
	while(k<len(arr_form)):
		if (arr_form[k][0]==well_first):#���� ������� ������������� ����� �����������
        		roof=roof+str(arr_form[k][1])+"/" #�������� abs ������ �������� ������
        		sole=sole+str(arr_form[k][2])+"/" #�������� abs ������� �������� ������
			code_old=code_old+str(arr_form[k][3])+"/" #�������� ��� �������� �������� ������
			code_lit=code_lit+str(arr_form[k][4])+"/" #�������� ��� ��������� �������� ������
    		else: #���� ������ ������� � ����� ��������
			current_array=[well_first,roof,sole,code_old,code_lit] #��������� ������ ���������� ��������
			def_arr_calc_width.append(current_array)# � �������� �� � ������ ������� �������
			well_first=arr_form[k][0]#���������� ������� �������������
			roof=str(arr_form[k][1])+"/" #������ ����� ������ abs ������ �������� ������
			sole=str(arr_form[k][2])+"/" #������ ����� ������ abs ������� �������� ������
			code_old=str(arr_form[k][3])+"/" # ������ ����� ������ c ������ �������� �������� ������
			code_lit=str(arr_form[k][4])+"/" # ������ ����� ������ � ������ ��������� �������� ������
    		k+=1
    		if (k==len(arr_form)):#�������� � ����� ������ ������ � ��������� ��������
        		current_array=[well_first,roof,sole,code_old,code_lit]
			def_arr_calc_width.append(current_array)
	#������� ������ � ������� [��� �����, ������ �������, ������� �������, ���� ������������, ���� ���������]
	return def_arr_calc_width
#������� � �������� ������� ���������� ����� ������ � �������[��� �����, ������ �������, ���� ������������, ���� ���������]
#� ���������� ������ � ��������� ���������
#� ������� [���� �����, abs ������/abs ������/abs ������,
#abs �������/abs �������/abs �������, Jmh/Jsn/Ol/, ���������/���������/�������/]
def view_strat_lit(arr_code_strat_lit,strat_code,lit_code):
        def_arr_text_strat_lit=[]
        k=0
        while(k<len(arr_code_strat_lit)):#������� ������� ������� ������� [��� �����, ������ �������, ���� ������������, ���� ���������]
                str_strat_code=str(arr_code_strat_lit[k][3])#�������� ������ ����� ������������ ��� ������� ��������
                arr_strat_code=str_strat_code.split("/")#�������� ������ ����� ������������ ��� ������� ��������
                str_lit_code=str(arr_code_strat_lit[k][4])#�������� ������ ����� ��������� ��� ������� ��������
                arr_lit_code=str_lit_code.split("/")#�������� ������ ����� ��������� ��� ������� ��������
                m=0
                stroka_strat=""#������ �� ������������� � ������� Jmh/Jsn/Ol/
                ##������� ����� ������������ ������� ��������, �� ������������ ��������� ������� ������� ����� ������������ ���������
		#�� �������� ������ - ��� �������� ������ '1/2/3/' ���������� ������ ['1','2','3',''] 
                while(m<len(arr_strat_code)-1):
                        n=0
                        #������� ������� ���� ����� ������������, ������ [���, ������������]
                        while(n<len(strat_code)):
                                cod1=float(arr_strat_code[m])#������������� � ����� ��� ������������ �� ������� arr_code_strat_lit
                                cod2=float(strat_code[n][0])#������������� � ����� ��� �� ������� ������� strat_code
                                if(cod1==cod2):#���� ���� �������
                                        #����������� ������ �� ������������� � ������� Jmh/Jsn/Ol/
                                        stroka_strat=stroka_strat+strat_code[n][1]+"/"
                                        #���� ���� ������� - ����� �� �������� �����(������� � ��������� ������������ ������� ��������)
                                        break 
                                n+=1
                        m+=1
                m=0
                stroka_lit=""#������ � ���������� � ������� ���������/���������/�������/
                #������� ����� ��������� ������� ��������, �� ������������ ��������� ������� ������� ����� ��������� ���������
		#�� �������� ������ - ��� �������� ������ '1/2/3/' ���������� ������ ['1','2','3',''] 
                while(m<len(arr_lit_code)-1):
                        n=0
                        #������� ������� ���� ����� ���������, ������ [���, ���������]
                        while(n<len(lit_code)):
                                cod1=float(arr_lit_code[m])#������������� � ����� ��� ��������� �� ������� arr_code_strat_lit
                                cod2=float(lit_code[n][0])#������������� � ����� ��� �� �������� ������� lit_code
                                if(cod1==cod2):#���� ���� �������
                                        #����������� ������ � ���������� � ������� ���������/���������/�������/
                                        stroka_lit=stroka_lit+lit_code[n][1]+"/"
                                        #���� ���� ������� - ����� �� �������� �����(������� � ��������� ��������� ������� ��������)
                                        break   
                                n+=1
                        m+=1
                #������������ ������� ������� ������� � ������� [��� �����,
                #abs ������/abs ������/abs ������,abs �������/abs �������/abs �������,Jmh/Jsn/Ol/,���������/���������/�������/]
                current_arr=[arr_code_strat_lit[k][0],arr_code_strat_lit[k][1],arr_code_strat_lit[k][2],stroka_strat,stroka_lit]
                #�������� ������� ������� � ��������� �������
                def_arr_text_strat_lit.append(current_arr)
                k+=1
        #������� ������ � ������� [���� �����, abs ������/abs ������/abs ������,abs �������/abs �������/abs �������,
        #Jmh/Jsn/Ol/, ���������/���������/�������/]
        return def_arr_text_strat_lit
#������� ����� ������ excel-����� � ������������ �����������
#� ���������� ������ � ��������� ��������� excel-�����(�����) � ������ ��������� ����� excel-����� 
def read_mineralogy(list_excel):
        title=True#���������� ������������� �������� �� ������ ������ ���������\
        def_mineral_arr = []#������ ��� �������� �����, ��������� �� ����� ����������� �����������
        for rownum in range(list_excel.nrows):#������� ����� �����
                if title:#���� ��������� ��������� �� ������ ������ �������, ��������������� "�����"
                        title=False
                        current_row = list_excel.row_values(rownum)
                        arr_name_cell = []#������ � ������� ��������� ����� ����������� �����������
                        for j, cell in enumerate(current_row):#������� ����� ����� ����� ����������� �����������
                                if(type(cell)==unicode):cell=cell.encode('utf-8')
                                arr_name_cell.append(cell)#�������� � ������ �������� �������� �� ������� ������
                        continue #������ ����� �������� ������ �����
                current_row = list_excel.row_values(rownum)#������� ������ �����
                curr_string = []#������ ��� ������� ������ ������� ������
                for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
                        if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
                        curr_string.append(cell)#�������� � ������ �������� ������� ������
                #���� ���������� ����� ������� ������ ����� ���������� ����� '�����'  
                #�������� � ������ ����� ������ ������� ������ 
                if(len(curr_string)==len(arr_name_cell)):def_mineral_arr.append(curr_string)
        #������� ������ � ��������� ��������� excel-�����(�����) � ������ ��������� ����� excel-����� 
        return arr_name_cell,def_mineral_arr
#������� �� ����� ����� ������ ��������� excel-����� � ������ �� ����� �������� excel-�����
#� ���������� ������ ����� � ��������� ����������
def calc_non_empty(atributes,values):
        k=0
        arr_non_empty = []#������ ��� ������ ���� �������� ����� 
        while(k<len(values)):#������� ����� �����
                l=0
                stroka_mineral=''
                id_points = 'no_data'#id ����� ����������
                number_test = 'no_data'#����� �����
                id_test = 'no_data'#id �����
                code_type_test = 'no_data'#��� ���� �����
                interval_from = 'no_data'#�������� ������� ��
                interval_to = 'no_data'#�������� ������� ��
                curr_arr = []#������ ��� ������ ��������� ������� ������
                while(l<len(atributes)):#������� ��������� ������� ������
                        #���� ����������� ������ � ���������
                        #�������� � �������� len(atributes)-1 ��� �������� �������� UIN
                        if(l==0):id_points = str(values[k][l])#id ����� ����������
                        if(l==1):number_test = str(values[k][l])#����� �����
                        if(l==2):id_test = str(values[k][l])#id �����
                        if(l==3):code_type_test = str(values[k][l])#��� ���� �����
                        if(l==4):interval_from = str(values[k][l])#�������� ������� ��
                        if(l==5):interval_to = str(values[k][l])#�������� ������� ��
                        if((l>=8)and(l<len(atributes)-1)):
                                #���� �������� �������� �������� �� ������, ������ ��� ������� ��������
                                if(values[k][l]<>''):
                                        stroka_mineral=stroka_mineral+str(atributes[l])+"="+str(values[k][l])+"/"
                                        curr_arr = [id_points,number_test,id_test,code_type_test,interval_from,interval_to,stroka_mineral]
                        l+=1
                #���� ������ ������� ������ �� ������
                #�������� � ����� ������ �������� �����
                if(len(curr_arr)<>0):arr_non_empty.append(curr_arr)
                k+=1
        #������� ������ � ������� [id ��������, ����� �����, id �����, ��� ���� �����, �������� ������ ��,
        #�������� ������ ��, ������� ���������]
        return arr_non_empty

#������� � �������� ��������� ����� ������ ����� Excel-�����, ����������� ���� ����� ����
#� ���������� ������ [��� �����, ������������ �����]
def read_code_type(list_excel):
	title=True#���������� ������������� �������� �� ������ ������ ���������
	def_array_type_test=[]#������ � ������ ����
	for rownum in range(list_excel.nrows):#������� ���� ����� �����
		if title: 
			title=False
			continue #���������� ������ "�����" � ����������� ���������
		row_array=[]#������ �����
		current_row=list_excel.row_values(rownum)#������� ������ �����
		def_code_type_test="no_data"#��� �����
		def_name_type_test = "no_data"#�������� �����
		for j, cell in enumerate(current_row):#������� ����� ������� ������ �����
			if(type(cell)==unicode):cell=cell.encode('utf-8')#���� ��� ������ unicode - ������������� � utf-8
			if((j==0)and(cell<>'')):def_code_type_test=cell #������ ������� � ����� �����
			if((j==1)and(cell<>'')):def_name_type_test=cell #������ ������� � ������������� �����
		def_array_type_test.append([def_code_type_test,def_name_type_test])#�������� ������� ������ � ������ ������� 
	return def_array_type_test #���������� ������ � ������� [��� ���������, ����������� ���������]

def replacement_code_test(def_non_empty,def_code_type_test):
        k=0
        while(k<len(def_non_empty)):#������� ��������� ������� ����� � ������������
                m=0
                while(m<len(def_code_type_test)):#������� ����� ����� �����
                        #���� ��� ���� ����� �� ������� � ������������� ������ � ����� �� ����� � ������ ����� ����
                        if(def_non_empty[k][3]==str(def_code_type_test[m][0])):
                                def_non_empty[k][3]=def_code_type_test[m][1]#�������� ��� ���� ����� �� �������� ���� �����(����, ���� � �.�.)
                                break#����� �� ����������� �����
                        m+=1
                k+=1
        return def_non_empty
        
#������� �� ����� ����� ������ ����� ���������� � ������������ � ������ ����������� �����������
#� ���������� ������������ ������ � �������
#[id_��������,�����/��� ��������,x,y,z,������,�������,����� �����,id �����,��� ���� �����,�������� ������� ��, �������� ������� ��, ������� ���]
def united_two_arrays(arr_wells, arr_min):
	#����� ������, ������� ����������� �� ������ ������� ����� ����� ���������� � ����� � ������������
	#� ������� ����� ������� �������� � �����������, ������ ��������, 
	#��������������, �� ������� - ������ � ������ � �������� ��������� ��������� �������
	def_united_array=[]
	k=0
	while(k<len(arr_min)):#������� ��������� ������� ����� (������� �����)
    		m=0
    		while(m<len(arr_wells)):#������� ��������� ������� ����� (������� � ������� �����������) 
        		if(arr_min[k][0]==str(arr_wells[m][0])):#���� �������������� �������
				#������ ��� ������ ����� � ��� �� ��������,
				#������� ������������ ���� ����� ������
				curr_array = [arr_wells[m][0],arr_wells[m][1],arr_wells[m][2],arr_wells[m][3],arr_wells[m][4], \
                                arr_wells[m][5],arr_wells[m][6],arr_wells[m][7],arr_min[k][1],arr_min[k][2],arr_min[k][3], \
                                arr_min[k][4],arr_min[k][5],arr_min[k][6],arr_wells[m][8],arr_wells[m][9],arr_wells[m][10], \
                                arr_wells[m][11]]
            			def_united_array.append(curr_array) #�������� � ����� ������ �����
	    			break
        		m=m+1
    		k=k+1
    	#������� ������������ ������ � ������������ � ������� �� �������� ���
    	#[id_��������,�����/��� ��������,x,y,z,������,�������,����� �����,id �����,��� ���� �����,�������� ������� ��,
    	#�������� ������� ��, ������� ���,��� ��������, ������ ����������������,��������� �����������,
    	#��� ������������]
	return def_united_array


# ������� �� ����� ����� ������ � �������[id_��������,�����/��� ��������,x,y,z,������,�������,
#����� �����,id �����,��� �����,�������� ������� ��, �������� ������� ��, ������� ���,
#��� ��������, ������ ����������������,��������� �����������,��� ������������]
#� ������ � �������[���� �����, abs ������/abs ������/abs ������,abs �������/abs �������/abs �������,
#Jmh/Jsn/Ol/, ���������/���������/�������/]
#� ���������� ��� �������
def create_final_array(arr_min,arr_strat_lit):
	def_united_array=[]#������������ ������
	if(len(arr_strat_lit)==0):#���� ������ � ������� �� ���������������� ����, ������� ������ � ������� ���������� �� ����������������
                k=0
                while(k<len(arr_min)):#������� ��������� ������� ����� (������� ����� c ������������)
                        #��� ����� �������� � ������������ ������� ������ �� ������� �������� ��� ������ �� ����������������
                        curr_array = [arr_min[k][0],arr_min[k][1],arr_min[k][2], \
                        arr_min[k][3],arr_min[k][4],arr_min[k][5],arr_min[k][6], \
                        arr_min[k][7],arr_min[k][8],arr_min[k][9],arr_min[k][10], \
                        arr_min[k][11],arr_min[k][12],arr_min[k][13],"","","","", \
                        arr_min[k][14],arr_min[k][15],arr_min[k][16],arr_min[k][17]]
            		def_united_array.append(curr_array) #�������� � ����� ������ �����
                        k+=1
                #������� ������ � ������� ���������� ����������������
                return def_united_array
	k=0
	while(k<len(arr_min)):#������� ��������� ������� ����� (������� ����� c ������������)
                find_stratolit=False#���������� ���������� ����������, ���� �� ��� ������� �������� ���������� �� ���������������� �������
    		m=0
    		while(m<len(arr_strat_lit)):#������� ��������� ������� ����� (������� � ������� ����������������) 
        		if(str(arr_min[k][0])==str(arr_strat_lit[m][0])):#���� �������������� �������
				#������ ��� ������ ����� � ��� �� ��������,
				#������� ������������ ���� ����� ������
				curr_array = [arr_min[k][0],arr_min[k][1],arr_min[k][2], \
                                arr_min[k][3],arr_min[k][4],arr_min[k][5],arr_min[k][6], \
                                arr_min[k][7],arr_min[k][8],arr_min[k][9],arr_min[k][10], \
                                arr_min[k][11],arr_min[k][12],arr_min[k][13],arr_strat_lit[m][1],arr_strat_lit[m][2], \
                                arr_strat_lit[m][3],arr_strat_lit[m][4], \
                                arr_min[k][14],arr_min[k][15],arr_min[k][16],arr_min[k][17]]
            			def_united_array.append(curr_array) #�������� � ����� ������ �����
            			find_stratolit=True#��� ������� �������� ���� ���������� �� ���������������� �������
        		m=m+1
        		#���� ��������� ����� ����� � ��� ���� �� ������� ���������� �����(�.�. ��� ������� �������� ����������� ������ �� ���������������� �������)
        		if(m==len(arr_strat_lit))and(not find_stratolit):
                                #��� ����� �������� � ������������ ������� ������ �� ������� �������� ��� ������ �� ����������������
                                curr_array = [arr_min[k][0],arr_min[k][1],arr_min[k][2], \
                                arr_min[k][3],arr_min[k][4],arr_min[k][5],arr_min[k][6], \
                                arr_min[k][7],arr_min[k][8],arr_min[k][9],arr_min[k][10], \
                                arr_min[k][11],arr_min[k][12],arr_min[k][13],"","","","", \
                                arr_min[k][14],arr_min[k][15],arr_min[k][16],arr_min[k][17]]
            			def_united_array.append(curr_array) #�������� � ����� ������ �����
    		k=k+1
    	#������� ������ � ������� [id_��������,�����/��� ��������,x,y,z,������,�������,
    	#����� �����,id �����,��� �����,�������� ������� ��, �������� ������� ��, ������� ���,
    	#abs ������/abs ������/abs ������,abs �������/abs �������/abs �������,
    	#Jmh/Jsn/Ol/, ���������/���������/�������/, ��� ��������, ������ ����������������,��������� �����������,
    	#��� ������������]
    	return def_united_array
def write_to_layer(lay, C_System, array_map, atr_min):
        
        fc_path, fc_name = os.path.split(lay)#��������� ���������� � ��� �����
        gp=arcgisscripting.create(9.3)#������� ������ �������������� ������ 9.3
        if fc_path.count("-") or fc_name.count("-"):
                gp.AddMessage(" '-' ������������ ������ � �������� �����")
                raise SystemExit(1)#��������� ������ ���������
        try:
                gp.CreateFeatureClass(fc_path,fc_name,"POINT","","","ENABLED",C_System)#������� �������� ����
        except:
                gp.AddMessage("������ �������� ����������������� ����")
                raise SystemExit(1)#��������� ������ ���������
        gp.AddField(lay,"id_��������","STRING",15,15)#������� id �������
	gp.AddField(lay,"��������_��������","STRING",15,15)#������� �������� �������
	gp.AddField(lay,"x","FLOAT",15,15)#�������� ������� x
	gp.AddField(lay,"y","FLOAT",15,15)#�������� ������� y
	gp.AddField(lay,"z","FLOAT",15,15)#�������� ������� z
	gp.AddField(lay,"�������_��������","STRING","","",50)#������� - ������� ��������
	gp.AddField(lay,"��������_�������","STRING","","",100)#������� - �������� �������
	gp.AddField(lay,"��������_�������","STRING","","",100)#������� - �������� �������
	gp.AddField(lay,"�����_�����","LONG","","",50)#������� - ����� �����
	gp.AddField(lay,"id_�����","LONG","","",50)#������� - id �����
	gp.AddField(lay,"���_�����","STRING","","",50)#������� - ��� ���� �����
	gp.AddField(lay,"��������_�����_��","FLOAT","","",50)#������� - �������� ������� ��
	gp.AddField(lay,"��������_�����_��","FLOAT","","",50)#������� - �������� ������� ��
	gp.AddField(lay,"�������_���","STRING","","",1000)#������� - ���������� �� ���
	gp.AddField(lay,"�������_������_�������","STRING","","",1000)#������� ��� ������ ������ �������
        gp.AddField(lay,"�������_�������_�������","STRING","","",1000)#������� ��� ������ ������� �������
        gp.AddField(lay,"������������","STRING","","",1000)#������� ��� ������������
        gp.AddField(lay,"���������","STRING","","",1000)#������� ��� ���������
        gp.AddField(lay,"���_�����_����������","STRING","","",1000)#������� ��� ���� ����� ����������
        gp.AddField(lay,"���������_����������������","STRING","","",1000)#������� ��� ��������� ����������������
        gp.AddField(lay,"���������_�����������","STRING","","",1000)#������� ��� ��������� �����������
        gp.AddField(lay,"���_����������������","STRING","","",1000)#������� ��� ���� ����������������
        #������� ������������ ���� ��� ������� ���
        u=8
        while(u<len(atr_min)-1):#�� ��������� ��������� �������, ��������� �� �� ��������� � �����������(UIN)
                gp.AddField(lay,atr_min[u],"LONG","","",50)
                u+=1
	rows = gp.InsertCursor(lay)
	pnt = gp.CreateObject("Point")
	k=0
	#[id_��������,�����/��� ��������,x,y,z,������,�������,
    	#����� �����,id �����,��� �����,�������� ������� ��, �������� ������� ��, ������� ���,
    	#abs ������/abs ������/abs ������,abs �������/abs �������/abs �������,
    	#Jmh/Jsn/Ol/, ���������/���������/�������/]
	while(k<len(array_map)): #�������� ��������� ����
                id_well = float(array_map[k][0])
    		name_well = str(array_map[k][1])
    		z = float(array_map[k][4])
    		depth_well = str(array_map[k][5])
    		name_object = str(array_map[k][6])
    		name_area = str(array_map[k][7])
    		number_test = str(array_map[k][8])
    		id_test = str(array_map[k][9])
    		code_type_test = str(array_map[k][10])
    		abs_from = float(array_map[k][11])
    		abs_to = float(array_map[k][12])
    		mineralogy = str(array_map[k][13])
    		arr_min = mineralogy.split("/")
    		abs_litostrat_roof = str(array_map[k][14])
    		abs_litostrat_sole = str(array_map[k][15])
    		stratigraphy = str(array_map[k][16])
    		lithology = str(array_map[k][17])
    		type_tn = str(array_map[k][18])
    		stat_doc = str(array_map[k][19])
    		stat_test = str(array_map[k][20])
    		type_doc = str(array_map[k][21])
    		feat = rows.NewRow()#����� ������
    		pnt.id = k #ID
    		pnt.x = float(array_map[k][3])
    		pnt.y = float(array_map[k][2])
    		feat.shape = pnt
    		feat.SetValue("id_��������",id_well)#�������� �������� � ������������ �����
    		feat.SetValue("��������_��������",name_well)
		feat.SetValue("x",float(array_map[k][3]))
		feat.SetValue("y",float(array_map[k][2]))
    		feat.SetValue("z",z)
    		feat.SetValue("��������_�������",name_object)
    		feat.SetValue("��������_�������",name_area)
    		feat.SetValue("�������_��������",depth_well)
                feat.SetValue("�����_�����",int(float(number_test)))
                feat.SetValue("id_�����",int(float(id_test)))
                feat.SetValue("���_�����",code_type_test)
                feat.SetValue("��������_�����_��",abs_from)
                feat.SetValue("��������_�����_��",abs_to)
                feat.SetValue("�������_���",mineralogy)
                feat.SetValue("�������_������_�������",abs_litostrat_roof)
                feat.SetValue("�������_�������_�������",abs_litostrat_sole)
                feat.SetValue("������������",stratigraphy)
                feat.SetValue("���������",lithology)
                feat.SetValue("���_�����_����������",type_tn)
                feat.SetValue("���������_����������������",stat_doc)
                feat.SetValue("���������_�����������",stat_test)
                feat.SetValue("���_����������������",type_doc)
                #��������� �������� ���������� ������������ ���� ��� ������� ���
                u=8
                while(u<len(atr_min)-1):#�� ��������� ��������� �������, ��������� �� �� ��������� � �����������(UIN)
                        feat.SetValue(atr_min[u],0) 
                        u+=1
                #������� ��������� ����� ����� � ������������
                i=8 #������ ������� � �������� pir05_0
                while(i<len(atr_min)-1):
                        #������� ������� ��� ��� ������� ��������, �� ������������ ��������� ������� �������
                        #�� �������� ������ - ��� �������� ������ '1/2/3/' ���������� ������ ['1','2','3','']          
                        m=0
                        while(m<len(arr_min)-1):
                                curr_min = arr_min[m].split("=")#��������� ��� ��� � ���������� ����������
                                if (atr_min[i]==curr_min[0]):#���� ������� �� ����� ����� ����������� ������ � �������� ��� ������� ��������
                                        val = int(float(curr_min[1]))
                                        feat.SetValue(curr_min[0],val)#�������� � �������� ���� ��� ���������� ��������� ���������� 
                                        break      
                                m+=1
                        i+=1
    		rows.InsertRow(feat)#�������� ������ � ������� ����
                k+=1
if __name__== '__main__':#���� ������ �������, � �� ������������
        #-------------------------------������� ������------------------------------------
        InputExcelFile = ARCPY.GetParameterAsText(0)#������� excel-����
        Coordinate_System = ARCPY.GetParameterAsText(1)#������������ ������� ��������� ��������� ����
        OutputFC = ARCPY.GetParameterAsText(2)#�������� �������� ���� �������
        Full_report = ARCPY.GetParameterAsText(3)#�������� ���������� �������� �� ������ �����
        List_point = ARCPY.GetParameterAsText(4)#�������� ����� � ������� ���������� 
        List_strat_lit = ARCPY.GetParameterAsText(5)#�������� ����� � ��������� ������������ � ���������
        List_code_strat = ARCPY.GetParameterAsText(6)#�������� ����� � ������ ������������
        List_code_lit = ARCPY.GetParameterAsText(7)#�������� ����� � ������ ���������
        List_mineral = ARCPY.GetParameterAsText(8)#�������� ����� � ������� �����������
        List_code_type_test = ARCPY.GetParameterAsText(9)#�������� ����� � ������ ����� ����
        List_code_type_TN = ARCPY.GetParameterAsText(10)#�������� ����� � ������ ����� ��
        List_code_status_document = ARCPY.GetParameterAsText(11)#�������� ����� � ������ ��������� ����������������
        List_code_status_test = ARCPY.GetParameterAsText(12)#�������� ����� � ������ ��������� �����������
        List_code_type_document = ARCPY.GetParameterAsText(13)#�������� ����� � ������ ����� ����������������
        #---------------------------------------------------------------------------------
        #�������� ������ excel-������
        obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_mineral_list,obj_code_type_test, \
        obj_code_type_TN, obj_code_status_document, obj_code_status_test, obj_code_type_document = \
        get_excel_list(InputExcelFile, List_point,List_strat_lit,List_code_strat,List_code_lit,List_mineral, \
        List_code_type_test,List_code_type_TN,List_code_status_document,List_code_status_test,List_code_type_document)
        #��������� ��� �� excel-����� � �������, ���������� ��� ������� ���������, �������
        verify_lists(obj_point, obj_strat_lit, obj_code_strat, obj_code_lit, obj_mineral_list,obj_code_type_test, \
        obj_code_type_TN,obj_code_status_document,obj_code_status_test,obj_code_type_document)
        wells_array = read_points_wells(obj_point)#�������� ������ ����� �� ����� exc�l-����� � ������� ����������
        if(len(wells_array)==0):#���� ������ ����� ���������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� ����� ���������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        ARCPY.AddMessage("����� ���������� ����� ����������, ����������� �� Excel-�����: "+str(len(wells_array)))
        ARCPY.AddMessage(wells_array[0])
        if(Full_report == "true"):#�������� ������ �����
                if(len(wells_array)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(wells_array[k][0])+" "+str(wells_array[k][1])+" "+ \
                                str(wells_array[k][2])+" "+str(wells_array[k][3])+" "+str(wells_array[k][4])+" "+ \
                                str(wells_array[k][5])+" "+str(wells_array[k][6])+" "+str(wells_array[k][7])+" "+ \
                                str(wells_array[k][8])+" "+str(wells_array[k][9])+" "+str(wells_array[k][10])+" "+ \
                                str(wells_array[k][11]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_tn = read_code_wells(obj_code_type_TN)#�������� ������ � ������ ����� ����� ����������
        ARCPY.AddMessage("���������� ����� � ������ ����� ����������: "+str(len(code_type_tn)))
        if(len(code_type_tn)==0): #���� ������ ����� ����� �� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� � ������ ����� �� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):
                if(len(code_type_tn)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_type_tn[k][0])+" "+str(code_type_tn[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_status_document = read_code_wells(obj_code_status_document)#�������� ������ � ������ ��������� ����������������
        ARCPY.AddMessage("���������� ����� ��������� ����������������: "+str(len(code_status_document)))
        if(len(code_status_document)==0):#���� ������ ����� ��������� ���������������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� � ������ ��������� ���������������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):
                ARCPY.AddMessage("� �������:")
                k=0
                while(k<len(code_status_document)):
                        ARCPY.AddMessage(str(code_status_document[k][0])+" "+str(code_status_document[k][1]))
                        k+=1
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_status_test = read_code_wells(obj_code_status_test)# �������� ������ � ������ ��������� �����������
        ARCPY.AddMessage("���������� ����� ��������� �����������: "+str(len(code_status_test)))
        if(len(code_status_test)==0):#���� ������ ����� ��������� ����������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� � ������ ��������� ����������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):
                ARCPY.AddMessage("� �������:")
                k=0
                while(k<len(code_status_test)):
                        ARCPY.AddMessage(str(code_status_test[k][0])+" "+str(code_status_test[k][1]))
                        k+=1
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_document = read_code_wells(obj_code_type_document)#�������� ������ � ������ ����� ����������������
        ARCPY.AddMessage("���������� ����� ����� ����������������: "+str(len(code_type_document)))
        if(len(code_type_document)==0):#���� ������ ����� ����� ���������������� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� � ������ ����� ���������������� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):
                ARCPY.AddMessage("� �������:")
                k=0
                while(k<len(code_type_document)):
                        ARCPY.AddMessage(str(code_type_document[k][0])+" "+str(code_type_document[k][1]))
                        k+=1
                ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #�������� ������ - ����� ������� ��, ��� �������� ������ ���� � ������ �� �� �����������
        wells_array_final = replacement_code_lists(wells_array,code_type_tn,code_status_document,code_status_test, \
        code_type_document)
        ARCPY.AddMessage("����� ���������� ����� ����������, � ��������������� ������ ����� ���������� �� ���������: "+ \
        str(len(wells_array_final)))
        if(Full_report == "true"):#�������� ������ �����
                if(len(wells_array)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(wells_array_final[k][0])+" "+str(wells_array_final[k][1])+" "+ \
                                str(wells_array_final[k][2])+" "+str(wells_array_final[k][3])+" "+str(wells_array_final[k][4])+" "+ \
                                str(wells_array_final[k][5])+" "+str(wells_array_final[k][6])+" "+str(wells_array_final[k][7])+" "+ \
                                str(wells_array_final[k][8])+" "+str(wells_array_final[k][9])+" "+str(wells_array_final[k][10])+" "+ \
                                str(wells_array_final[k][11]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_strat = read_code_wells(obj_code_strat)#�������� ������ � ������ ������������
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
        code_lit = read_code_wells(obj_code_lit)#�������� ������ � ������ ���������
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
        if(len(arr_formation)==0):#���� ������ ��������� ������� ����
                arr_single_width = []#������� ������ ������ � ���������������� �������� ��� ������ ��������
                ARCPY.AddMessage("���������� � ��������� �����������")
        else:#���� ������ ��������� �� ����
                arr_single_width = single_width_blocking(arr_formation) #�������� ������ � ���������������� �������� ��� ������ ��������
        ARCPY.AddMessage("���������� ������� � ����������� � �������� �������: "+str(len(arr_single_width)))
        if(len(arr_single_width)==0):#���� ������ � ���������������� �������� ��� ������ �������� ���� 
                arr_view_strat_lit=[]#������� ������ ������ � ��������� ��������� ������������ � ��������� � ������� [���� �����, �������/�������/�������,Jmh/Jsn/Ol/
                ARCPY.AddMessage("������ � ��������������� �������� ��� ������ �������� ����")
        else:
                if(Full_report == "true"):#�������� ������ �����
                        ARCPY.AddMessage("� �������:")
                        ARCPY.AddMessage(str(arr_single_width[0][0])+" "+str(arr_single_width[0][1])+ \
                        " "+str(arr_single_width[0][2])+" "+str(arr_single_width[0][3]))
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
                        #�������� ������ � ��������� ��������� ������������ � ��������� � ������� [���� �����, �������/�������/�������,Jmh/Jsn/Ol/, ���������/���������/�������/]
                arr_view_strat_lit = view_strat_lit(arr_single_width,code_strat,code_lit)
        ARCPY.AddMessage("���������� ������� � ��������� ��������� ������������ � ���������: "+str(len(arr_view_strat_lit)))
        if(len(arr_view_strat_lit)==0):#���� ��� ������� �� �������������������� ���������
                ARCPY.AddMessage("��� ���������� �� ������������ � ���������")
        else:#���� ���� �������� �� �������������������� ���������
                if(Full_report == "true"):#�������� ������ �����
                        ARCPY.AddMessage("� �������:")
                        ARCPY.AddMessage(str(arr_view_strat_lit[0][0])+" "+str(arr_view_strat_lit[0][1])+" "+str(arr_view_strat_lit[0][2])+\
                        " "+str(arr_view_strat_lit[0][3])+" "+str(arr_view_strat_lit[0][4]))
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")         
        #�������� ������ � ��������� ��������� excel-�����(�����) ����������� � ������ ��������� ����� excel-����� 
        list_atributes,value_atributes=read_mineralogy(obj_mineral_list)   
        #������� ������ �������� ����� ����� �����������
        non_empty = calc_non_empty(list_atributes,value_atributes)
        if (len(non_empty)==0):
                ARCPY.AddMessage("��� ������ �� �����������")
                raise SystemExit(1)#��������� ������ ���������
        ARCPY.AddMessage("����� ���������� ����� �������� � ��������� ��������� - ��������� �������: "+str(len(non_empty)))
        if(Full_report=="true"):#�������� ������ �����
                if (len(non_empty)>10):
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(non_empty[k]))
                                k+=1
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        code_type_test = read_code_type(obj_code_type_test)#�������� ������ � ������ ����� ����
        ARCPY.AddMessage("���������� ����� c ������ ����: "+str(len(code_type_test)))
        if(len(code_type_test)==0):#���� ������ ����� � ������ ���� ���� - ��������� ������ ���������
                ARCPY.AddMessage("���� ����� ����� ����, ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                if(len(code_type_test)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                ARCPY.AddMessage(str(code_type_test[k][0])+" "+str(code_type_test[k][1]))
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #�������� ������, ��� ������ ����� ����� ����� - �������� ��������, ����� ��� ����, ���� � �.�.
        arr_mineral = replacement_code_test(non_empty,code_type_test)
        ARCPY.AddMessage("����� ���������� ����� �������� � ��������������� ������ ����: "+str(len(arr_mineral)))
        if(Full_report == "true"):#�������� ������ �����
                if(len(arr_mineral)>10):
                        ARCPY.AddMessage("� �������:")
                        k=0
                        while(k<10):
                                m=0
                                stroka=''
                                while(m<7):
                                        stroka = stroka+' '+str(arr_mineral[k][m])
                                        m+=1
                                ARCPY.AddMessage(stroka)
                                k+=1
                        ARCPY.AddMessage("...........")
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")        
        #�������� ������������ ������ � ������������ � ������� �����������
        united_array = united_two_arrays(wells_array_final, arr_mineral)
        ARCPY.AddMessage("���������� ��������� ������������� ������� ������ ������� � �����������: "+str(len(united_array)))
        if(len(united_array)==0):#���� ������������ ������ ����
                ARCPY.AddMessage("������ ������ ������� � ��� ����, ��������� �����������")
                raise SystemExit(1)#��������� ������ ���������
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                if(len(arr_mineral)>5):
                        k=0
                        while(k<5):
                                ARCPY.AddMessage(str(united_array[k][0])+" "+str(united_array[k][1])+" "+str(united_array[k][2])+ \
                                " "+str(united_array[k][3])+" "+str(united_array[k][4])+" "+str(united_array[k][5])+ \
                                " "+str(united_array[k][6])+" "+str(united_array[k][7])+" "+str(united_array[k][8])+ \
                                " "+str(united_array[k][9])+" "+str(united_array[k][10])+" "+str(united_array[k][11])+ \
                                " "+str(united_array[k][12])+" "+str(united_array[k][13])+" "+str(united_array[k][14])+ \
                                " "+str(united_array[k][15])+" "+str(united_array[k][16])+" "+str(united_array[k][17]))
                                k+=1
                        ARCPY.AddMessage("-----------------------------------------------------------------------------------")
        #�������� ��������� ������ ��� ������� � ���������������� ����
        #���� ��������� ������ ����������� �� ������ ����������� ������� � ������������ � ������������
        #� ������� � �����������������
        final_array = create_final_array(united_array,arr_view_strat_lit)
        ARCPY.AddMessage("���������� ��������� ������������� ������� ��� ������ � ���������������� ���� "+str(len(final_array)))
        if(Full_report == "true"):#�������� ������ �����
                ARCPY.AddMessage("� �������:")
                m=0
                stroka=''
                while(m<22):
                        stroka = stroka+' '+str(final_array[0][m])
                        m+=1
                ARCPY.AddMessage(stroka)
        #������� ���������� ������ � ���������������� ���� 
        write_to_layer(OutputFC, Coordinate_System, final_array, list_atributes)

