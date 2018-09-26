from MSSQLConnect import MSSQL
import xlrd
import os
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
    colnames = [1,2,3,4,5,6,7,8,9,10]
    last_code=''
    last_operway=''
    last_type=''
    num=0
    for rownum in range(0, nrows):

        row = table.row_values(rownum)
        if row:
            app = {}
            for i in range(len(colnames)):

                if i<9:
                    app[colnames[i]] = str(row[i])
                else :
                    num+=1
                    app[colnames[i]] = str(num)
                if(i==1):
                    #print(colnames[i],i)
                    if last_code!=app[colnames[i]] :
                        num=0
                        last_code=app[colnames[i]]
                elif i==3:
                    if last_type!=app[colnames[i]] :
                        num=0
                        last_type = app[colnames[i]]
                elif i==4:
                    if last_operway!=app[colnames[i]] :
                        num=0
                        last_operway = app[colnames[i]]

            list.append(app)
    return list




dir="D:\\开放式基金\\测试\\折扣率\\公募基金定投费率优惠汇总20161227"
files=[]
sum=0
n=0
propFile = open('D:\\disctratio.txt', 'w+')
for root,dirs,files in os.walk(dir):

    for file in files:
        n = 0
        detailTable = excel_table_byindex(os.path.join(root,file),1)
        #detailTable = excel_table_byindex("D:\开放式基金\\测试\\折扣率\\公募基金定投费率优惠汇总20161227\\中融基金修订.xlsx",1)
        sum+=detailTable.__len__()
        # 将基金代码都规范为6位数

        for i in detailTable:
            j=i[2].split('.')[0]
            i[2]=j.zfill(6)

            if i[4]=='申购':
                i[4]='240022'
            elif i[4]=='定投':
                i[4]='240039'
            if i[5]=='0.0':
                i[5]='0'
            print(i)
            #if ('代码' in i[2]):
            #    print(i)
            if ('代码' not in i[2]):

                #discSQL="INSERT INTO run..ofhqdiscratioctrl ([serverid],[orgid],[tacode],[ofcode],[trdid],[pt_operway],[subsection],[toplimit],[lowlimit],[topdiscratio],[lowdiscratio],[begindate],[enddate],[remark])  VALUES"\
                #"(1 ,'0000','777',%s,%s,%s,%s,%s,%s,%s,%s,'','','');" % (i[2] ,i[4],i[5],i[10],i[7],i[6],i[9],i[8])
                discSQL = "INSERT INTO workdb.dbo.ofhqdiscratioctrl_0206 ([serverid],[orgid],[tacode],[ofcode],[trdid],[pt_operway],[subsection],[toplimit],[lowlimit],[topdiscratio],[lowdiscratio],[begindate],[enddate],[remark])  VALUES" \
                          "(1 ,'0000','777','%s','%s','%s','%s','%s','%s','%s','%s','','','');" % (i[2] ,i[4],i[5],i[10],i[7],i[6],i[9],i[8])
                #print(discSQL)
                #ms.ExecNonQuery(discSQL)
                n += 1
                propFile.write(discSQL)
                propFile.write('\n')
            else:
                print(i[2]+'--=====================分割线=====================')

        propFile.write('-- %s,行数：%d \n\n\n' % (file, n))
        propFile.write('--=====================分割线=====================\n\n')


print(n)
