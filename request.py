#姓名正则表达式：r'\w+\??\s?\w*'
#Zone2正则表达式：r'\(\D+\d+\'?\w*\)'


import re
import os
import openpyxl

#从文本文档中读取名字
pathd = os.getcwd() + '\\rawdata.txt'
rawdata = open(pathd)
namecontent = rawdata.read()

#将所有的y i and 换行都用逗号代替
namecontent = re.sub('\\s[yi]\\s',',',namecontent)
namecontent = re.sub('\\s and \\s',',',namecontent)
namecontent = re.sub('\\n',',',namecontent)
#将所有的min前加上逗号
namecontent = re.sub('\.',' ',namecontent)
namecontent = re.sub(r'\bmin',',min',namecontent)

#所有逗号之间的内容都保存在compzone中
compzonere = re.compile(r'\w+\??\s?\w*\(\D+\d+\'?\w*\)')
compzone = compzonere.findall(namecontent)
#定义姓名正则
namere = re.compile(r'\w+\??\s?\w*')

#所有的括号中内容保存在zone2中
zone2re = re.compile(r'\(\D+\d+\'?\w*\)')   
zone2 = zone2re.findall(namecontent)
#提取所有的替补队员和他们的上场时间
#替补姓名存储在namechtable[]
#上场时间存储在namechnum[]
#出场次数储存在nametimes[]里面
fstname = []
fstnum = []
fsttime = []
sndname = []
sndnum = []
sndtime = []
fieldtime1 = 0
fieldtime2 = 0
for i in range(len(compzone)):
    current = 0
    current = compzone[i]
    fieldtime1 = re.search(r'\d+',current)
    fieldtime1 = int(fieldtime1.group(0))
    fieldtime2 = 90 - fieldtime1
    zone2temp = re.search(r'\(\D+\d+\'?\w*\)',current)
    current = re.sub(r'\(\D+\d+\'?\w*\)','',current)
    zone1temp = re.search(r'\w+\??\s?\w*',current)
    zone2temp = re.search(r'\w+\??\s?\w*',zone2temp.group(0))
    fstname.append(zone1temp.group(0))
    sndname.append(zone2temp.group(0))
    fsttime.append(fieldtime1)
    sndtime.append(fieldtime2)

#删除所有compzone
namecontent = re.sub(compzonere,'',namecontent)
#print(namecontent)

#建立zone1
zone1re = re.compile(r'\w+\??\s?\w*')
zone1 = zone1re.findall(namecontent)

#建立完整的fstname与sndname
for i in range(len(zone1)):
    fsttime.append(90)
fstname = fstname + zone1

for i in range(len(fstname)):
    fstname[i] = fstname[i].strip()
    if i<len(sndname):
        sndname[i] = sndname[i].strip()


#剔除重复的名字并计数
set_fstname = []
for element in fstname :
    if(element not in set_fstname):
        set_fstname.append(element)
set_fstnum = []
set_fsttime = []
for i in range(len(set_fstname)):
    count = 0
    tim = 0
    for j in range(len(fstname)):
        if(set_fstname[i] == fstname[j]):
            count = count + 1
            tim = tim + fsttime[j]
    set_fstnum.append(count)
    set_fsttime.append(tim)

fsttime = set_fsttime
fstnum = set_fstnum
fstname = set_fstname


#剔除重复的名字并计数
set_sndname = []
for element in sndname :
    if(element not in set_sndname):
        set_sndname.append(element)
set_sndnum = []
set_sndtime = []
for i in range(len(set_sndname)):
    count = 0
    tim = 0
    for j in range(len(sndname)):
        if(set_sndname[i] == sndname[j]):
            count = count + 1
            tim = tim + sndtime[j]
    set_sndnum.append(count)
    set_sndtime.append(tim)

sndtime = set_sndtime
sndnum = set_sndnum
sndname = set_sndname


#将名字写入xlsx
patha = os.getcwd() + '\\Results.xlsx'
wb = openpyxl.load_workbook(patha)
sheet = wb.get_active_sheet()

compname = []
compfstnum = []
compsndnum = []
compfsttime = []
compsndtime = []
comptime = []
for i in range(len(sndname)):

        for j in range(len(fstname)):
            if(i<len(sndname)):
                if((j<len(fstname))):
                    if(sndname[i]==fstname[j]):
                        compname.append(fstname[j])
                        compfstnum.append(fstnum[j])
                        compsndnum.append(sndnum[i])
                        compfsttime.append(fsttime[j])
                        compsndtime.append(sndtime[i])
                        comptime.append(sndtime[i] + fsttime[j])
                        del fstname[j],fstnum[j],fsttime[j],sndnum[i],sndtime[i],sndname[i]

#同时有首发和替补
for i in range(len(compname)):
    sheet['A'+str(i+2)] = compname[i]
    sheet['B'+str(i+2)] = compfstnum[i]
    sheet['C'+str(i+2)] = compsndnum[i]
    sheet['D'+str(i+2)] = compfsttime[i]
    sheet['E'+str(i+2)] = compsndtime[i]
    sheet['F'+str(i+2)] = comptime[i]
    c = i+1

#仅有首发
for i in range(len(fstname)):
    sheet['A'+str(c+i+2)] = fstname[i]
    sheet['B'+str(c+i+2)] = fstnum[i]
    sheet['C'+str(c+i+2)] = 0
    sheet['D'+str(c+i+2)] = fsttime[i]
    sheet['E'+str(c+i+2)] = 0
    sheet['F'+str(c+i+2)] = fsttime[i]
    b=c+i+1

#仅有替补
for i in range(len(sndname)):
    sheet['A'+str(b+i+2)] = sndname[i]
    sheet['B'+str(b+i+2)] = 0
    sheet['C'+str(b+i+2)] = sndnum[i]
    sheet['D'+str(b+i+2)] = 0
    sheet['E'+str(b+i+2)] = sndtime[i]
    sheet['F'+str(b+i+2)] = sndtime[i]

wb.save(patha)
