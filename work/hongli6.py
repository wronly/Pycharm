import xlrd
import xlwt
import numpy as np
'''
业务规则：
  例如：分红登记日是20160901，我们要找到从一年前（20150902）到年末20161231.这个段里面，持有够一年的股份数，
    进行退税。如果最大值不够一年，但是假设到2017年一直持有，我们默认持有够一年。
  需要注意：
    数据一定要对。分红登记日对应的持仓可能有送股，要去掉。如果有股票融资融券，也要去掉。
    分红金额要等于 持仓*每股分红。每股分红可以去通达信查公告。
  涉及到红股：
    如果一只股票一年多次分红送股。比如20160901 20161001都送股了。
    那么，在算20150902-20160901的时候，20160901当天持仓不应该包含送股，但是在20161001计算的时候就要加上送股。
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
detailTable = excel_table_byindex("D:\\2\\20171227\\0-1 chichang.xls")
fhTable = excel_table_byindex("D:\\2\\20171227\\1-0 liushui.xls")
'''
这里是一对测试表格。如果有个别数据有问题，我们可以把个别股票放入里面单独跑，很快就能定位
'''
#detailTable = excel_table_byindex("D:\\2\\factory\\new2.xls")
#fhTable = excel_table_byindex("D:\\2\\factory\\new1.xls")

# 比较此点在一年内，是否是最小值
'''
思路：
如果分红登记日是20160901，那么我们从20150902开始，每一天都往后数一年的时间。查找：在这一个时间段里面，最小的
持仓的日期，把这个最小值所在的列（仅仅是这一段时间里面的最小值）返回。
'''
def compareYear(row1,towYearHold, dateBegin, dateEnd, minNum,djr1):

    min1 = minNum
    minRow = row1
    if int(row1['备份时间'])==djr1:
        min1=float(row1['当前持仓'])-float(row1['分红送股'])
    scode=row1['证券代码']
    for row2 in towYearHold:
        if int(row2['备份时间']) >=int(dateBegin)  and int(row2['备份时间']) < int(dateEnd ):
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
'''
    从这里开始，逐行遍历分红数据。每一行分红数据，对应一个股票。
    然后，我们把这个股票从20150831到 20160901这段时间里面，最大的退税股票找到。

'''
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

    #print(stkcode)
    '''
    fhDetail 用于存放每一条分红记录对应的股票的对应的最终计算结果。
    '''
    fhDetail = {}
    fhDetail['证券代码'] = stkcode
    fhDetail['登记日'] = row['交易日期']
    for temp in detailTable:
        if (int(temp['证券代码']) ==int( stkcode)) and (int(temp['备份时间']) == int(row['交易日期'])):
            fhDetail['登记股份'] = float(temp['当前持仓'])-float(temp['分红送股'])
            print(fhDetail['登记股份'] )
    #djNum=fhDetail['登记股份']
    fhDetail['红利金额'] = row['红利金额']
    fhDetail['最晚持有日期'] = row['交易日期']
    ksr = (int(fhDetail['登记日']) - 10000)
    djr =  int(fhDetail['登记日'])
    #===================新增加==================
    #print(fhDetail['登记股份'])
    fhDetail['自算分红金额']=fhDetail['登记股份']*float(row['税前派现金额'])
    fhDetail['自算分红金额']=round(fhDetail['自算分红金额'],2)
    if fhDetail['自算分红金额']==float(row['红利金额']) or fhDetail['自算分红金额']==float(row['红利金额'])*10:
        fhDetail['一致性']=''
    else:
        fhDetail['一致性'] = '不一致'

    # 得到该分红股票的所有持仓数据
    stkList = []
    #=============================================
    fhDetail['退税金额'] = 0
    fhDetail['免税股数'] = 0
    #print(ksr,djr)
    '''
      stkList 里面存放从20150902-20160901每一天的持仓数据
    '''
    for row in detailTable:
        if (int(row['证券代码']) == int(stkcode)) and (int(row['备份时间']) > ksr) and (int(row['备份时间']) <= djr ) :
            stkList.append(row)

    # 获得最理想的那个点
    resultList = {}
    resultList['date'] = 20110000
    resultList['num'] = -1
    '''
    如果没有持仓，len=0，我们就直接输出结果
    有持仓，就把stkList里面的每一天往后数一年，去找最小值。
    再找之前，我们需要把这只股票的2015-16年的持仓数据单独保存下来，提高效率
    '''
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
    '''
    保存有关股票的两年的持仓20150902---20161231
    '''
    towYearHold=[]
    for row in detailTable:
        if (int(row['证券代码']) == int(stkcode)) and (int(row['备份时间']) > ksr)  :
            towYearHold.append(row)

    '''
       有持仓，就把stkList里面的每一天往后数一年，去找最小值。
    '''
    for row in stkList:
        minnum=float(row['当前持仓'])
        dateBegin=int(row['备份时间'])
        dateEnd=dateBegin+10000
        djr1=int(fhDetail['登记日'])
        #resultList['end'] = dateEnd
        #resultList['begin'] = dateBegin
        #print(row)
        res=compareYear(row,towYearHold, dateBegin, dateEnd, minnum,djr1)

        '''
        这里得到了每一行的最小值，然后跟最终的退税股份比较：
        因为这里是从20160901到20160902 365天，365个最小值。我们要从里面找一个最大值，才是
        利益最大化
        '''
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

                #print(resultList,res)
            #print(res, dateBegin)
            #print(row)

    fhDetail['免税股数'] = resultList['num']
    fhDetail['最晚持有日期'] = resultList['end']
    fhDetail['最早持有日期'] = resultList['begin']
    if fhDetail['最晚持有日期']>20180000:
        fhDetail['最晚持有日期']=20171231
    #fhDetail['最早持有日期'] = resultList['date']
    if float(fhDetail['登记股份'])==0:
        fhDetail['退税金额']=0
        fhDetail['免税股数']=0
        fhDetail['免税金额'] =0
    else:
        fhDetail['退税金额'] = round((resultList['num'] / float(fhDetail['登记股份']))  * 0.25 * float(fhDetail['红利金额']), 2)
        fhDetail['免税金额'] = round((resultList['num']/float(fhDetail['登记股份']))*float(fhDetail['红利金额']),2)
    result.append(fhDetail)


# 输出结果=+++++++++++++++++++++================================================================================================================
for row in result:
    print('证券代码:',row['证券代码'],row['登记股份'],row['免税股数'],row['退税金额'],row['最早持有日期'],row['最晚持有日期'],row['红利金额'],row['自算分红金额'],row['登记日'],row['一致性'])
    #print(row)

wTable=xlwt.Workbook()
sheet1=wTable.add_sheet('sheet1')
sheet1.write(0, 0, '证券代码')
sheet1.write(0, 1, '登记股份')
sheet1.write(0, 2, '免税股数')
sheet1.write(0, 3, '免税金额')
sheet1.write(0, 4, '退税金额')
sheet1.write(0, 5, '最早持有日期')
sheet1.write(0, 6, '最晚持有日期')
sheet1.write(0, 7, '红利金额')
sheet1.write(0, 8, '登记日')
sheet1.write(0, 9, '一致性')
sheet1.write(0, 10,'自算分红金额')
for i in range(result.__len__()):


    sheet1.write(i+1, 0, result[i]['证券代码'])
    sheet1.write(i+1, 1, result[i]['登记股份'])
    sheet1.write(i+1, 2, result[i]['免税股数'])
    sheet1.write(i+1, 3, result[i]['免税金额'])
    sheet1.write(i+1, 4, result[i]['退税金额'])
    sheet1.write(i+1, 5, result[i]['最早持有日期'])
    sheet1.write(i+1, 6, result[i]['最晚持有日期'])
    sheet1.write(i+1, 7, result[i]['红利金额'])
    sheet1.write(i+1, 8, result[i]['登记日'])
    sheet1.write(i+1, 9, result[i]['一致性'])
    sheet1.write(i+1, 10, result[i]['自算分红金额'])

    wTable.save("D:\\2\\dai\\res0-1.xls")


