def remove(string): 
    return "".join(string.split()) 

def remove_space(string):
    return " ".join(string.split())

def add_to_dict(dictionary,Id,sec,name):
    #(x,y)= dictionary[Id][sec]
    dictionary[Id][sec]=(0,set())

import xlrd
import os
import sys 

cwd=(os.getcwd())
loc = (cwd)+'/timetable.xlsx'

wb = xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(0)

course_id=-1
section=-1
class_pattern=-1
name=-1
title=-1
course_code=-1
cap=-1
class_instructor=-1
course_catalog=-1

for i in range(sheet.ncols):
    if(remove_space(sheet.cell_value(1, i).lower())=='course id'):
        course_id=i
    if(remove_space(sheet.cell_value(1, i).lower())=='tot enrl'):
        cap=i
    if(remove_space(sheet.cell_value(1, i).lower())=='course title'):
        title=i
    if(remove_space(sheet.cell_value(1, i).lower())=='section'):
        section=i
    if(remove_space(sheet.cell_value(1, i).lower())=='class_instructor'):
        class_instructor=i
    if(remove_space(sheet.cell_value(1, i).lower())=='class pattern'):
        class_pattern=i
    if(remove_space(sheet.cell_value(1, i).lower())=='subject'):
        course_code=i
    if(remove_space(sheet.cell_value(1, i).lower())=='catalog'):
        course_catalog=i
    if(remove_space(sheet.cell_value(1, i).lower())=='instructor name'):
        name=i

if(course_id==-1):
    sys.exit("'course id' column not found in sheet0 of course.xlsx sheet....\n")
if(section==-1):
    sys.exit("'section' column not found in sheet0 of course.xlsx sheet....\n")
if(class_pattern==-1):
    sys.exit("'class pattern' column not found in sheet0 of course.xlsx sheet....\n")
if(course_code==-1):
    sys.exit("'subject' column not found in sheet0 of course.xlsx sheet....\n")
if(title==-1):
    sys.exit("'course title' column not found in sheet0 of course.xlsx sheet....\n")
if(course_id==-1):
    sys.exit("'course id' column not found in sheet0 of course.xlsx sheet....\n")
if(cap==-1):
    sys.exit("'tot enrl' column not found in sheet0 of course.xlsx sheet....\n")
if(class_instructor==-1):
    sys.exit("'class_instructor' column not found in sheet0 of course.xlsx sheet....\n")

Dict={}
Dict_title={}
Dict_coursecode={}
for i in range(2 ,sheet.nrows):
    Dict[sheet.cell_value(i,course_id)]={}
    #print(sheet.cell_value(i,course_id))

for i in range(2 ,sheet.nrows):
    Dict_title[sheet.cell_value(i,course_id)]=sheet.cell_value(i,title)

for i in range(2 ,sheet.nrows):
    Dict_coursecode[sheet.cell_value(i,course_id)]=sheet.cell_value(i,course_code)+sheet.cell_value(i,course_catalog)

for i in range(2 ,sheet.nrows):
    s=(sheet.cell_value(i,section))[0]
    add_to_dict(Dict,sheet.cell_value(i,course_id),s,sheet.cell_value(i,name))

Dict_cap={}
for i in range(2 ,sheet.nrows):
    y=set()
    (x,y)=Dict[sheet.cell_value(i,course_id)][sheet.cell_value(i,section)[0]]
    s=(sheet.cell_value(i,name))
    if(((sheet.cell_value(i,name))[-1])=='.'):
        s=sheet.cell_value(i,name)[0:-2]
    y.add((s,(sheet.cell_value(i,class_instructor))))
    if (remove(sheet.cell_value(i,course_id)),remove(sheet.cell_value(i,section))) not in Dict_cap:
        x=x+((sheet.cell_value(i,cap)))
        Dict_cap[((sheet.cell_value(i,course_id)),sheet.cell_value(i,section))]=1
    Dict[sheet.cell_value(i,course_id)][sheet.cell_value(i,section)[0]]=(x,y)

import xlsxwriter 
workbook = xlsxwriter.Workbook('output2.xlsx')
workbook.add_worksheet('output')
cell_format = workbook.add_format({'bold': True, 'align': 'center'})
cell_format3 =workbook.add_format ({'align': 'center'})
worksheet=workbook.get_worksheet_by_name('output')
worksheet.set_column(0,200, 50)
worksheet.write(0,0,'Course Id',cell_format)
worksheet.write(0,1,'Course Title',cell_format)
worksheet.write(0,2,'Course Code',cell_format)
worksheet.write(0,3,'Section',cell_format)
worksheet.write(0,4,'Instructor Name Faculty',cell_format)
worksheet.write(0,5,'Instructor Name PHD',cell_format)
worksheet.write(0,6,'Capacity',cell_format)

cnt=1
for i,j in Dict.items():
    for k,l in j.items():
        y=set()
        (x,y)=l
        s_fac=""
        s_phd=""
        for (val,nm) in y:
            if(nm[0]=='G'):
                s_fac=s_fac+val+" ,"
            else:
                s_phd=s_phd+val+" ,"
        s_fac=s_fac[0:-2]+'.'
        s_phd=s_phd[0:-2]+'.'
        #print(s)
        worksheet.write(cnt,0,i,cell_format3)
        worksheet.write(cnt,1,Dict_title[i],cell_format3)
        worksheet.write(cnt,2,Dict_coursecode[i],cell_format3)
        worksheet.write(cnt,3,k,cell_format3)
        worksheet.write(cnt,4,s_fac,cell_format3)
        worksheet.write(cnt,5,s_phd,cell_format3)
        worksheet.write(cnt,6,x,cell_format3)
        cnt=cnt+1
workbook.close() 
