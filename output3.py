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
days=-1
course_catalog=-1

for i in range(sheet.ncols):
    if(remove_space(sheet.cell_value(1, i).lower())=='mtg start'):
        strt=i
    if(remove_space(sheet.cell_value(1, i).lower())=='course id'):
        course_id=i
    if(remove_space(sheet.cell_value(1, i).lower())=='class pattern'):
        days=i
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
    if(remove_space(sheet.cell_value(1, i).lower())=='catalog'):
        course_catalog=i
    if(remove_space(sheet.cell_value(1, i).lower())=='instructor name'):
        name=i
#name=class_instructor+1

print(df['Class_Instructor'][1220])

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

Set_D=set()
for i in range(2 ,sheet.nrows):
    Set_D.add(sheet.cell_value(i,days))

Dict_Days={'F':1,
 'M':1,
 'MW':2,
 'MWF':3,
 'S':1,
 'T':1,
 'TF':2,
 'TH':1,
 'THF':2,
 'TS':2,
 'TTH':2,
 'TTHS':3,
 'W':1,
 'WF':2}

#print(Dict)    
Dict_visit={}
for i in range(2 ,sheet.nrows):
    ((((w1,w2),w3),((x1,x2),x3)),(((y1,y2),y3),((z1,z2),z3)))=(Dict[(sheet.cell_value(i,course_id),sheet.cell_value(i,name))])
    s=(sheet.cell_value(i,section))[0]
    q=remove(sheet.cell_value(i,course_id)).lower()+remove(sheet.cell_value(i,name)).lower()+(remove(sheet.cell_value(i,section))[0]).lower()+remove((sheet.cell_value(i,class_pattern)).lower())+str(sheet.cell_value(i,strt))
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
        b=1
        if(sheet.cell_value(i,days) != ''):
            b=Dict_Days[sheet.cell_value(i,days)]
        tym=(endd2-start2)
        if(tym==30):
            #print(1)
            c=((endd1-start1)+((endd2-start2)/60));
            #print(q)
        else:
            c=((endd1-start1)+(int)((endd2-start2+30)/60));
        if(s.lower()=='p'):
            y3+=1
            y1+=c*b
            y2+=capacty
        if(s.lower()=='l'):
            w3+=1
            w1+=c*b
            w2+=capacty
        if(s.lower()=='t'):
            x3+=1
            x1+=c*b
            x2+=capacty
        if(s.lower()=='r'):
            z3+=1
            z1+=c*b
            z2+=capacty
        Dict[(sheet.cell_value(i,course_id),sheet.cell_value(i,name))]=((((w1,w2),w3),((x1,x2),x3)),(((y1,y2),y3),((z1,z2),z3)))
    
Dict_isphd={}
for i in range(2, sheet.nrows):
    Dict_isphd[sheet.cell_value(i,name)]=((sheet.cell_value(i,class_instructor))[0]!='G')

Dict_PSRN={}
for i in range(2, sheet.nrows):
    Dict_PSRN[sheet.cell_value(i,name)]=((sheet.cell_value(i,class_instructor)))

        
Dict_course_id={}
Dict_code={}
for i in range(2, sheet.nrows):
    Dict_course_id[sheet.cell_value(i,course_id)]=sheet.cell_value(i,title)
    Dict_code[sheet.cell_value(i,course_id)]=sheet.cell_value(i,course_code)+" "+sheet.cell_value(i,course_catalog)

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
worksheet.write(0,4,'Instructor Name',cell_format)
worksheet.write(0,5,'Is_PHD',cell_format)
worksheet.write(0,6,'Lecture',cell_format)
worksheet.write(0,7,'Lecture Time',cell_format)
worksheet.write(0,8,'Lecture Cap',cell_format)
worksheet.write(0,9,'Tut',cell_format)
worksheet.write(0,10,'Tut Time',cell_format)
worksheet.write(0,11,'Tut Cap',cell_format)
worksheet.write(0,12,'Practical',cell_format)
worksheet.write(0,13,'Practical Time',cell_format)
worksheet.write(0,14,'Practical Cap',cell_format)
worksheet.write(0,15,'R',cell_format)
worksheet.write(0,16,'R Time',cell_format)
worksheet.write(0,17,'R Cap',cell_format)
worksheet.write(0,18,'Is IC',cell_format)
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
    worksheet.write(cnt,4,Dict_PSRN[i],cell_format3)
    worksheet.write(cnt,5,str(Dict_isphd[i]),cell_format3)
    worksheet.write(cnt,6,w3,cell_format3)
    worksheet.write(cnt,7,w1,cell_format3)
    worksheet.write(cnt,8,w2,cell_format3)
    worksheet.write(cnt,9,x3,cell_format3)
    worksheet.write(cnt,10,x1,cell_format3)
    worksheet.write(cnt,11,x2,cell_format3)
    worksheet.write(cnt,12,y3,cell_format3)
    worksheet.write(cnt,13,y1,cell_format3)
    worksheet.write(cnt,14,y2,cell_format3)
    worksheet.write(cnt,15,z3,cell_format3)
    worksheet.write(cnt,16,z1,cell_format3)
    worksheet.write(cnt,17,z2,cell_format3)
    worksheet.write(cnt,18,a,cell_format3)
    cnt=cnt+1
workbook.close() 
