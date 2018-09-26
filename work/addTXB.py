import time
import datetime
import sched
import os
from tkinter import messagebox
# 今天日期
##############################################此文档已乱，需要重新搞一下了
today = datetime.date.today()
# 昨天时间
yesterday = today - datetime.timedelta(days=1)
yesterday_end_time = int(time.mktime(time.strptime(str(today), '%Y-%m-%d'))) - 1
schedule = sched.scheduler(time.time, time.sleep)
my_date=time.strftime('%Y%m%d',time.localtime(time.time()))
last_day=time.strftime('%Y%m%d',time.localtime(time.time()))

filePath="D:\\程序\\AddTXB\\"

txbName='EndAgm_'+(my_date)+'.TXT'
xjbName='EndAgmRtn_'+(last_day)+'.TXT'
newName='EndAgmRtn_'+my_date+'.txt'

n=1
'''
while os.path.exists(xjbName) is False:
    n+=1
    last_day = time.strftime('%Y%m%d', time.localtime(time.time()))
    xjbName = 'EndAgmRtn_' + (last_day) + '.TXT'
'''
messagebox.showinfo("现金宝解约文件日期提示", '现金宝解约确认文件的日期为'+last_day)
xjbFile = open(filePath+xjbName)
print()
txbFile=  open(filePath+txbName)
print(xjbFile,txbFile)
xjbLine = xjbFile.readlines()
txbLine = txbFile.readlines()

#计算得到第一行：总行数
lens_xjb=int(xjbLine[0])
lens_txb=int(txbLine[0])
lens=lens_xjb+lens_txb
str_lens=str(lens).zfill(8)

#写第一行
f=open(filePath+newName,'w+')
f.write(str_lens+'\n')

#新文件写入xjb数据
for x in xjbLine:
    if len(x)<10:
        r=1
    else:
        f.write(x)

#新文件写入txb数据，并加入“1成功”
for s in txbLine:
    if len(s)<10:
        r=1
    else:
        index=s.index('\n')
        y=s[0:index]
        f.write(y+'1成功\n')

txbFile.close()
xjbFile.close()
print('END')
f.write('END')
f.close()



