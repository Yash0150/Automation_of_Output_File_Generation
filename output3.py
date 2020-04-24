def remove(string): 
    return "".join(string.split()) 

def remove_space(string):
    return " ".join(string.split())

def add_to_dict(dictionary,Id,name):
    #(x,y)= dictionary[Id][sec]
    dictionary[(Id,name)]=((((0,0),0),((0,0),0)),(((0,0),0),((0,0),0)))
    
def add_to_dict_instr(dictionary,Id,name):
    #(x,y)= dictionary[Id][sec]
    dictionary[(Id,name)]=False

import xlrd
import os
import sys 

cwd=(os.getcwd())
loc = (cwd)+'/timetable.xlsx'

wb = xlrd.open_workbook(loc)
sheet=wb.sheet_by_index(0)

import pandas as pd
import numpy
from datetime import time
df = pd.read_excel(loc,'sheet1')
df.columns = df.iloc[0]
df=df.iloc[1:]
df=df.replace(numpy.nan,time(0,0,0))

course_id=-1
section=-1
class_pattern=-1
name=-1
title=-1
course_code=-1
cap=-1
class_instructor=-1
rol=-1
strt=-1

for i in range(sheet.ncols):
    if(remove_space(sheet.cell_value(1, i).lower())=='mtg start'):
        strt=i
    if(remove_space(sheet.cell_value(1, i).lower())=='course id'):
        course_id=i
    if(remove_space(sheet.cell_value(1, i).lower())=='role'):
        rol=i
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
name=class_instructor+1

if(course_id==-1):
    sys.exit("'course id' column not found in sheet0 of course.xlsx sheet....\n")
if(strt==-1):
    sys.exit("'mtg start' column not found in sheet0 of course.xlsx sheet....\n")
if(rol==-1):
    sys.exit("'role' column not found in sheet0 of course.xlsx sheet....\n")
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
for i in range(2 ,sheet.nrows):
    Dict[(sheet.cell_value(i,course_id),sheet.cell_value(i,name))]={}
    #print(sheet.cell_value(i,course_id))

for i in range(2 ,sheet.nrows):
    add_to_dict(Dict,sheet.cell_value(i,course_id),sheet.cell_value(i,name))


#print(Dict)    
Dict_visit={}
for i in range(2 ,sheet.nrows):
    ((((w1,w2),w3),((x1,x2),x3)),(((y1,y2),y3),((z1,z2),z3)))=(Dict[(sheet.cell_value(i,course_id),sheet.cell_value(i,name))])
    s=(sheet.cell_value(i,section))[0]
    q=remove(sheet.cell_value(i,course_id)).lower()+remove(sheet.cell_value(i,name)).lower()+(remove(sheet.cell_value(i,section))[0]).lower()
                                +remove((sheet.cell_value(i,class_pattern)).lower())+str(sheet.cell_value(i,strt))
    start1=df['Mtg Start'][i-1].hour
    endd1=df['End Time'][i-1].hour
    start2=df['Mtg Start'][i-1].minute
    endd2=df['End Time'][i-1].minute
    capacty=(sheet.cell_value(i,cap))
    if q in Dict_visit.keys():
        if(s.lower()=='p'):
            y2+=capacty
        if(s.lower()=='l'):
            w2+=capacty
        if(s.lower()=='t'):
            x2+=capacty
        if(s.lower()=='r'):
            z2+=capacty
        Dict[(sheet.cell_value(i,course_id),sheet.cell_value(i,name))]=((((w1,w2),w3),((x1,x2),x3)),(((y1,y2),y3),((z1,z2),z3)))
        continue
    else:
        Dict_visit[q]=1
        if(s.lower()=='p'):
            y3+=1
            y1+=((endd1-start1)*60+(endd2-start2))
            y2+=capacty
        if(s.lower()=='l'):
            w3+=1
            w1+=((endd1-start1)*60+(endd2-start2))
            w2+=capacty
        if(s.lower()=='t'):
            x3+=1
            x1+=((endd1-start1)*60+(endd2-start2))
            x2+=capacty
        if(s.lower()=='r'):
            z3+=1
            z1+=((endd1-start1)*60+(endd2-start2))
            z2+=capacty
        Dict[(sheet.cell_value(i,course_id),sheet.cell_value(i,name))]=((((w1,w2),w3),((x1,x2),x3)),(((y1,y2),y3),((z1,z2),z3)))
    

        
Dict_course_id={}
Dict_code={}
for i in range(2, sheet.nrows):
    Dict_course_id[sheet.cell_value(i,course_id)]=sheet.cell_value(i,title)
    Dict_code[sheet.cell_value(i,course_id)]=sheet.cell_value(i,course_code)

Dict_isInst={}
for i in range(2 ,sheet.nrows):
    add_to_dict_instr(Dict_isInst,sheet.cell_value(i,course_id),sheet.cell_value(i,name))
for i in range(2 ,sheet.nrows):
    #s=sheet.cell_value(i,title)
    if(sheet.cell_value(i,rol)=='PI'):
        Dict_isInst[(sheet.cell_value(i,course_id),sheet.cell_value(i,name))]=True

import xlsxwriter 
workbook = xlsxwriter.Workbook('output3.xlsx')
workbook.add_worksheet('output')
cell_format = workbook.add_format({'bold': True, 'align': 'center'})
cell_format3 =workbook.add_format ({'align': 'center'})
worksheet=workbook.get_worksheet_by_name('output')
worksheet.set_column(0,200, 50)
worksheet.write(0,0,'Course Id',cell_format)
worksheet.write(0,1,'Course Title',cell_format)
worksheet.write(0,2,'Course Code',cell_format)
worksheet.write(0,3,'Instructor Name',cell_format)
worksheet.write(0,4,'Lecture',cell_format)
worksheet.write(0,5,'Lecture Time',cell_format)
worksheet.write(0,6,'Lecture Cap',cell_format)
worksheet.write(0,7,'Tut',cell_format)
worksheet.write(0,8,'Tut Time',cell_format)
worksheet.write(0,9,'Tut Cap',cell_format)
worksheet.write(0,10,'Practical',cell_format)
worksheet.write(0,11,'Practical Time',cell_format)
worksheet.write(0,12,'Practical Cap',cell_format)
worksheet.write(0,13,'R',cell_format)
worksheet.write(0,14,'R Time',cell_format)
worksheet.write(0,15,'R Cap',cell_format)
worksheet.write(0,16,'Is IC',cell_format)
cnt=1
for (k,i),j in Dict.items():
    ((((w1,w2),w3),((x1,x2),x3)),(((y1,y2),y3),((z1,z2),z3)))=j
    a='YES'
    if(Dict_isInst[(k,i)]==False):
        a="NO"
    worksheet.write(cnt,0,k,cell_format3)
    worksheet.write(cnt,1,Dict_course_id[k],cell_format3)
    worksheet.write(cnt,2,Dict_code[k],cell_format3)
    worksheet.write(cnt,3,i,cell_format3)
    worksheet.write(cnt,4,w3,cell_format3)
    worksheet.write(cnt,5,w1,cell_format3)
    worksheet.write(cnt,6,w2,cell_format3)
    worksheet.write(cnt,7,x3,cell_format3)
    worksheet.write(cnt,8,x1,cell_format3)
    worksheet.write(cnt,9,x2,cell_format3)
    worksheet.write(cnt,10,y3,cell_format3)
    worksheet.write(cnt,11,y1,cell_format3)
    worksheet.write(cnt,12,y2,cell_format3)
    worksheet.write(cnt,13,z3,cell_format3)
    worksheet.write(cnt,14,z1,cell_format3)
    worksheet.write(cnt,15,z2,cell_format3)
    worksheet.write(cnt,16,a,cell_format3)
    cnt=cnt+1
workbook.close() 
