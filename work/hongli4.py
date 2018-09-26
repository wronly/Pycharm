import xlrd
import xlwt
import numpy as np


# 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的所以  ，by_index：表的索引
def excel_table_byindex(file, colnameindex=0, by_index=0):
    data = xlrd.open_workbook(file)
    table = data.sheets()[by_index]
    nrows = table.nrows  # 行数
    ncols = table.ncols  # 列数
    colnames = table.row_values(colnameindex)  # 某一行数据
    list = []
    for rownum in range(1, nrows):

        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):
                app[colnames[i]] = row[i]
            list.append(app)
    return list


# qxTable=excel_table_byindex("D:\\2\\1QX.xls")
detailTable = excel_table_byindex("D:\\2\\dai\\1-0chichang.xls")
fhTable = excel_table_byindex("D:\\2\\dai\\1-0liushui.xls")
#detailTable = excel_table_byindex("D:\\2\\factory\\new2.xls")
#fhTable = excel_table_byindex("D:\\2\\factory\\new1.xls")

# 比较此点在一年内，是否是最小值
def compareYear(row1,detailTable, dateBegin, dateEnd, minNum,djr1):

    min1 = minNum
    minRow = row1
    if int(row1['备份时间'])==djr1:
        min1=float(row1['当前持仓'])-float(row1['分红送股'])
    scode=row1['证券代码']
    for row2 in detailTable:
        if int(row2['备份时间']) >=int(dateBegin)  and int(row2['备份时间']) < int(dateEnd )   and row2['证券代码']==scode:
            if int(row2['备份时间'])==djr1:
                if min1 >  float(row2['当前持仓'])-float(row2['分红送股']):
                    min1 = float(row2['当前持仓'])-float(row2['分红送股'])
                    minRow = row2
            else:
                if min1 > float(row2['当前持仓']):
                    min1 = float(row2['当前持仓'])
                    minRow = row2
                    #print(minRow)
                    #print(min1,minRow['备份时间'],dateBegin,dateEnd)
    #print('------------------------------',minRow)
    return minRow


# 结果表，用于存放结果
result = []

# 分红总共条数，用于显示输出
fhLen = fhTable.__len__()

for row in fhTable:
    if  float(row['对应持仓数量']) == float(0) :
        fhDetail = {}
        fhDetail['证券代码'] = row['stkcode']
        fhDetail['登记日'] = row['交易日期']
        fhDetail['登记股份'] = 0

        fhDetail['红利金额'] = row['红利金额']
        fhDetail['最晚持有日期'] = row['交易日期']
        fhDetail['免税股数'] = 0
        fhDetail['最早持有日期'] = '持仓为0，无意义'
        fhDetail['退税金额'] = 0

        result.append(fhDetail)
        continue

    stkcode = row['stkcode']

    print(stkcode)
    fhDetail = {}
    fhDetail['证券代码'] = stkcode
    fhDetail['登记日'] = row['交易日期']
    for temp in detailTable:
        if (temp['证券代码'] == stkcode) and (int(temp['备份时间']) == int(row['交易日期'])):
            fhDetail['登记股份'] = float(temp['当前持仓'])-float(temp['分红送股'])
    #djNum=fhDetail['登记股份']
    fhDetail['红利金额'] = row['红利金额']
    fhDetail['最晚持有日期'] = row['交易日期']
    ksr = (int(fhDetail['登记日']) - 10000)
    djr = int(fhDetail['登记日'])
    # 得到该分红股票的所有持仓数据
    stkList = []
    #=============================================
    fhDetail['退税金额'] = 0
    fhDetail['免税股数'] = 0
    #print(ksr,djr)
    #把所有的数量低于等于登记日的数量的都记下来。
    for row in detailTable:
        if (row['证券代码'] == stkcode) and (int(row['备份时间']) > ksr) and (int(row['备份时间']) <= djr ) :
            stkList.append(row)

    # 获得最理想的那个点
    resultList = {}
    resultList['date'] = 20110000
    resultList['num'] = -1

    if stkList.__len__()==0:
        resultList['date'] = '>=20160101'
        fhDetail['最晚持有日期'] = '<=20161231'
        fhDetail['最早持有日期'] = '>=20160101'
        fhDetail['登记股份']=0
        fhDetail['免税股数']=0
        fhDetail['免税金额']=0
        fhDetail['退税金额']=0
        fhDetail['红利金额']=0
        fhDetail['登记日']  =0
        result.append(fhDetail)
        continue
    #开始时间

    # 把所有stkList里面的都往后推一年，找每个点的最低持仓
    for row in stkList:
        minnum=float(row['当前持仓'])
        dateBegin=int(row['备份时间'])
        dateEnd=dateBegin+10000
        djr1=int(fhDetail['登记日'])
        #resultList['end'] = dateEnd
        #resultList['begin'] = dateBegin
        #print(row)
        res=compareYear(row,detailTable, dateBegin, dateEnd, minnum,djr1)


        if int(row['备份时间'])==djr1:
            if resultList['num'] <  float(res['当前持仓'])-float(res['分红送股']):
                resultList['num'] = float(res['当前持仓'])-float(res['分红送股'])
                resultList['begin'] = dateBegin
                resultList['end'] = dateEnd
        else:
            if resultList['num']< float(res['当前持仓']):
                resultList['num'] = float(res['当前持仓'])
                resultList['begin'] = dateBegin
                resultList['end'] = dateEnd

                print(resultList,res)
            #print(res, dateBegin)
            #print(row)

    fhDetail['免税股数'] = resultList['num']
    fhDetail['最晚持有日期'] = resultList['end']
    fhDetail['最早持有日期'] = resultList['begin']
    if fhDetail['最晚持有日期']>20170000:
        fhDetail['最晚持有日期']=20161231
    #fhDetail['最早持有日期'] = resultList['date']
    if float(fhDetail['登记股份'])==0:
        fhDetail['退税金额']=0
        fhDetail['免税股数']=0
        fhDetail['免税金额'] =0
    else:
        fhDetail['退税金额'] = round((resultList['num'] / float(fhDetail['登记股份'])) * 0.25 * float(fhDetail['红利金额']), 2)
        fhDetail['免税金额'] = round((resultList['num']/float(fhDetail['登记股份']))*float(fhDetail['红利金额']),2)
    result.append(fhDetail)


# 输出结果=+++++++++++++++++++++================================================================================================================
for row in result:
    #print('证券代码'':,row['证券代码'],row['登记股份'],row['免税股数'],row['退税金额'],row['最早持有日期'],row['最晚持有日期'],row['红利金额'],row['登记日'])
    print(row)

wTable=xlwt.Workbook()
sheet1=wTable.add_sheet('sheet1')
for i in range(result.__len__()):


    sheet1.write(i, 0, result[i]['证券代码'])
    sheet1.write(i, 1, result[i]['登记股份'])
    sheet1.write(i, 2, result[i]['免税股数'])
    sheet1.write(i, 3, result[i]['免税金额'])
    sheet1.write(i, 4, result[i]['退税金额'])
    sheet1.write(i, 5, result[i]['最早持有日期'])
    sheet1.write(i, 6, result[i]['最晚持有日期'])
    sheet1.write(i, 7, result[i]['红利金额'])
    sheet1.write(i, 8, result[i]['登记日'])

    wTable.save("D:\\2\\dai\\res1-0.xls")

'''
excel_init_file = xlrd.Workbook("D:\\2\\factory\\res0.xls")
table = excel_init_file.add_worksheet('bas_info')
title_bold = excel_init_file.add_format({'bold': True, 'border': 2, 'bg_color':'blue'})
border = excel_init_file.add_format({ 'border': 1})
excel_title = ['证券代码', '登记股份', '免税股数', '退税金额', '最早持有日期', '最晚持有日期', '红利金额', '登记日']
for i,j in enumerate(excel_title):
    table.set_column(i,i,len(j)+1)
    table.write_string(0,i,j,title_bold)
for k,v in result.items():
    for i in range(len(v)):
        j = v.get(excel_title[i])
        table.write_string(k,i,j,border)
table.set_column(1,1,16)
table.set_column(0,0,16)
excel_init_file.close()

for row in fhTable:
    print(row)
for row in detailTable:
    print(row)

for row in qxTable:
    print(row['zqdm1'])

    for row in result:
    sum+=float(row['退税金额'])
    #print('证券代码'':,row['证券代码'],row['登记股份'],row['免税股数'],row['退税金额'],row['最早持有日期'],row['最晚持有日期'],row['红利金额'],row['登记日'])
print(sum)
'''


