# -*- coding: utf-8  -*-
import time
import  os
import sched
from tkinter import *
from tkinter import messagebox


schedule = sched.scheduler(time.time, time.sleep)

my_date=time.strftime('%Y%m%d',time.localtime(time.time()))

myStr= '文件OFD_602_98_'

content=myStr+my_date+'_03.TXT接收完毕'

day=my_date[4:8]

djgsName = "C:\\Users\\wangweixx\\Documents\\Python\\djgsFile.xml"

djgsFile = open(djgsName)

djgsLines = djgsFile.readlines()

def perform_command(cmd, inc,j):

    djgsLines=cmd
    print("this is perform_command")

    #fileName = "D:\\temp\\exchangelog\\"+my_date+"\\jijin"+day+".log"
    fileName="C:\\Users\\wangweixx\\Documents\\Python\\jijin1025.log"
    # 读取配置文件

    file = open(fileName)

    line = file.readlines()
    #把list换成字符串，方便检测是否包含
    str_line=''.join(line)

    #设置一个参数，如果i ，如果i=djgsLines的行数，则证明
    i=0
    for x in djgsLines:
        content=x.replace('YYYYMMDD',my_date)
        print(content)
        if content in str_line:
            i=i+1
            print("_________________________!!!!!!!!!!!!!!!!!!!!!!!!!!________________________")


    if i== len(djgsLines):

        messagebox.showinfo("congratulate", '中登文件已发送到深圳通')

        exit(0)

    j=j+1
    if j% 4==0:
        messagebox.showinfo("注意", "已经过了10分钟，中登文件尚未全部发送！")

    schedule.enter(inc, 0, perform_command, (cmd, inc,j))


def timming_exe(cmd, inc ,i):

    print("this id timming_exe")
    schedule.enter(1, 0, perform_command, (cmd, inc,i))
    # 持续运行，直到计划时间队列变成空为止
    schedule.run()


try:
    messagebox.showinfo("congratulate", '即将开始')
    timming_exe(djgsLines, 30,1)
except Exception as e :
    messagebox.showerror("ERROR", e)

finally:

    messagebox.showinfo("congratulate", '结束！')




#
# 开始接收文件 OFD_602_98_20161102_23.TXT.ACC !
# 18:04:28  文件OFD_602_98_20161102_23.TXT.ACC接收完毕!
# 18:04:28  开始接收文件 OFD_602_99_20161102_23.TXT.ACC !
# 18:04:29  文件OFD_602_99_20161102_23.TXT.ACC接收完毕!
# 18:04:29  开始接收文件 OFI_602_99_20161102.TXT.ACC !
# 18:04:29  文件OFI_602_99_20161102.TXT.ACC接收完毕!
# 18:04:29  开始接收文件 OFI_602_98_20161102.TXT.ACC !
# 18:04:30  文件OFI_602_98_20161102.TXT.ACC接收完毕!
# 18:04:30  开始接收文件 OFD_602_98_20161102_03.TXT.ACC !
# 18:04:30  文件OFD_602_98_20161102_03.TXT.ACC接收完毕!
# 18:04:30  开始接收文件 OFD_602_99_20161102_03.TXT.ACC !
# 18:04:31  文件OFD_602_99_20161102_03.TXT.ACC接收完毕!
# 18:04:31  开始接收文件 OFD_602_98_20161102_13.TXT.ACC !
# 18:04:31  文件OFD_602_98_20161102_13.TXT.ACC接收完毕!
# 18:04:31  开始接收文件 OFD_602_98_20161102_01.TXT.ACC !
# 18:04:32  文件OFD_602_98_20161102_01.TXT.ACC接收完毕!
# 18:04:32  开始接收文件 OFD_602_99_20161102_01.TXT.ACC !
# 18:04:32  文件OFD_602_99_20161102_01.TXT.ACC接收完毕!
# 18:04:32  开始接收文件 OFD_602_99_20161102_13.TXT.ACC !
# 18:04:33  文件OFD_602_99_20161102_13.TXT.ACC接收完毕!
