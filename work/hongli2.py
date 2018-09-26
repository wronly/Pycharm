import xlrd
import numpy as np

#根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file,colnameindex=0,by_index=0):
    data = xlrd.open_workbook(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    colnames =  table.row_values(colnameindex) #某一行数据
    list =[]
    for rownum in range(1,nrows):

         row = table.row_values(rownum)
         if row:
             app = {}
             for i in range(len(colnames)):
                app[colnames[i]] = row[i]
             list.append(app)
    return list

qxTable=excel_table_byindex("D:\\2\\1QX.xls")
detailTable=excel_table_byindex("D:\\2\\1600016_detail.xls")
fhTable=excel_table_byindex("D:\\2\\1600016fh.xls")

#结果表，用于存放结果
result=[]


#分红总共条数，用于显示输出
fhLen=fhTable.__len__()
for row in fhTable:
    stkcode=row['stkcode']
    fhDetail={}
    fhDetail['证券代码']=stkcode
    fhDetail['登记日']=row['交易日期']
    fhDetail['登记股份']=row['对应持仓数量']
    fhDetail['红利金额'] = row['红利金额']
    fhDetail['最晚持有日期'] = row['交易日期']
    #得到该分红股票的所有持仓数据
    stkList=[]
    for row in detailTable:
        if row['证券代码']==stkcode  :
            stkList.append(row)
    print('应该与detail数量一致')
    print(stkList.__len__())
    #分布根据分红日期来得到每次分红一年内的数据
    #逻辑判断分支，如果不行，就跳入else执行
    if (stkList.__len__()==0):
        break
    if(int(stkList[0]['备份时间']))<(int(fhDetail['登记日'])-10000):
        #print(int(stkList[0]['备份时间']))
        #print(int(fhDetail['登记日'])-10000)
        print('这是我们要处理的逻辑')
        #获得免税股数、最早持有日期、最晚持有日期
        yearList=[]
        djr=(int(fhDetail['登记日'])-10000)
       #获取每次分后一年内的数据
        for row in stkList:
            if(int(row['备份时间']))>=djr:
                yearList.append(float(row['当前持仓']))

        min1 = min(yearList)

        print(min1)
        for row in stkList:
            if min1==float(row['当前持仓']):
                #print(row)
                fhDetail['免税股数'] = row['当前持仓']
                fhDetail['最早持有日期'] = row['备份时间']





    else:
        print('首次持仓就不足一年，另外一个逻辑，需要请教代总')


    result.append(fhDetail)

#输出结果=+++++++++++++++++++++
for row in result:
    print(row)


'''
for row in fhTable:
    print(row)
for row in detailTable:
    print(row)

for row in qxTable:
    print(row['zqdm1'])
'''


