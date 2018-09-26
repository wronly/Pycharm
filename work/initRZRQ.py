import xlrd
import xlwt
import numpy as np
import datetime

now=datetime.datetime.now()

delta=datetime.timedelta(days=1)
d_zero=datetime.timedelta(days=0)

t_str='20160101'

#print(d.strftime('%Y%m%d'))

'''



'''

# 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file,by_index=1,colnameindex=0):
    data = xlrd.open_workbook(file)
    list = []
    for i in range(by_index):
        table = data.sheets()[i]
        nrows = table.nrows  # 行数
        ncols = table.ncols  # 列数
        #存放第一行的数据作为表头
        colnames = table.row_values(colnameindex)  # 某一行数据

        for rownum in range(1, nrows):

            row = table.row_values(rownum)
            if row:
                app = {}
                for i in range(len(colnames)):
                    app[colnames[i]] = row[i]
                list.append(app)
                #print(app)

    return list


# qxTable=excel_table_byindex("D:\\2\\1QX.xls")
initTable = excel_table_byindex("D:\\2\\RZRQ\\initnum.xls")
changeTable = excel_table_byindex("D:\\2\\RZRQ\\detail.xls")
stkcode=0;
rownum=1;
wTable = xlwt.Workbook()
sheet1 = wTable.add_sheet('sheet1')
#bizdate = datetime.datetime.strptime(t_str, '%Y%m%d')
sheet1.write(0, 0, '备份时间')
sheet1.write(0, 1, '证券代码')
sheet1.write(0, 2, '当前持仓')
sheet1.write(0, 3, '分红送股')

for row in initTable:
    stkcode=row['stkcode']
    t_str=row['bizdate']
    daytime = datetime.datetime.strptime(t_str, '%Y%m%d')
    endtime = datetime.datetime.strptime('20171231', '%Y%m%d')
    bal=float(row['bal'])
    while daytime-endtime<=d_zero:
        if (daytime.weekday()==5 or daytime.weekday()==6):
            daytime = daytime + delta
            continue
        for row1 in changeTable:

            changetime=datetime.datetime.strptime(row1['bizdate'], '%Y%m%d')
            nstk=row1['stkcode']
            if(changetime-daytime==d_zero and float(stkcode)==float(nstk)):
                if (float(row1['digestid'])==float(551002)):
                    bal=bal+0
                    break
                else:
                    bal=bal+float(row1['stkeffect'])
                    break

        sheet1.write(rownum, 0, daytime.strftime('%Y%m%d'))
        sheet1.write(rownum, 1, stkcode)
        sheet1.write(rownum, 2, bal)
        sheet1.write(rownum, 3, 0)
        daytime=daytime+delta
        rownum=rownum+1
wTable.save("D:\\2\\RZRQ\\rzrq.xls")