from MSSQLConnect import MSSQL
import xlrd
import os
import sys
#连接数据库
ms = MSSQL(host="10.101.223.9", user="sa", pwd="sa", db="run")



# 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file,by_index,colnameindex=0):
    data = xlrd.open_workbook(file)
    list = []
    table = data.sheets()[by_index]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    # 存放第一行的数据作为表头
    colnames = table.row_values(colnameindex)  # 某一行数据
    colnames = [1,2,3,4]
    for rownum in range(0, nrows):

        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):

                app[colnames[i]] = str(row[i])
            list.append(app)
    return list




dir="D:\\开放式基金\\测试\\折扣率\\公募基金定投费率优惠汇总20161227"
propFile = open('D:\\prop.txt', 'w+')
files=[]
sum=0
n=0
for root,dirs,files in os.walk(dir):
    for file in files:
        detailTable = excel_table_byindex(os.path.join(root,file),0)
        #detailTable = excel_table_byindex("D:\开放式基金\\测试\\折扣率\\公募基金定投费率优惠汇总20161227\\鹏华基金核对.xlsx",0)
        sum+=detailTable.__len__()
        n+=1
        #记录每个文件的条数
        k=0
        #记录每个文件的基金名称
        file_name=file


        for i in detailTable:
            # 将基金代码都规范为6位数
            j = i[2].split('.')[0]
            i[2] = j.zfill(6)
            if(i[4]=='0.0'):
                i[4]='0'
            elif i[4]=='1.0':
                i[4]='1'
            if ('代码' not in i[2]):

                propSQL="UPDATE run..ofcodeprop  set timeropen = '%s' where ofcode='%s' \n"  % (i[4],i[2])
                print(propSQL)
                propFile.write(propSQL)
                k+=1
                # 为每一段SQl输入表头，标明文件与个数
        propFile.write('-- %s,行数：%d \n\n\n' % (file_name,k))
        propFile.write('--=====================分割线=====================\n\n')
print(n,sum)
propFile.write('--共%d个文件，%d 条语句' % ( len(files),sum))