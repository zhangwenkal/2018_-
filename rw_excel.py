import xlrd
from xlrd import open_workbook
from datetime import datetime
import os,re
import xlwt
from xlutils.copy import copy

path="E:\D"
filename="test.xlsx"

filepath=os.path.join(path,filename)

xl=open_workbook(filepath)

table=xl.sheet_by_index(0)
#print(table.ncols,table.nrows)
create_time=table.col_values(-3)[1:]
update_time=table.col_values(-2)[1:]
#print(create_time[0])


# d1 = datetime.strptime('2012-03-05 00:41:20', '%Y-%m-%d %H:%M:%S')
# d2 = datetime.strptime('2012-03-01 00:41:20', '%Y-%m-%d %H:%M:%S')
# delta = d1 - d2
#print(delta.days)

"""
参数格式:"14/六月/18 9:56 下午"
格式装换，将带汉字的时间 转换成 2018-6-04 20:27:00 格式
传参类型序列
"""


def change_time(args):
    # 拼接正则表达式
    sy=';./+ :'
    symbol = "[" + sy + "]+"
    # 一次性分割字符串
    a=[]
    for ar in args:
        result = re.split(symbol, ar)
        # 去除空字符
        a1=[x for x in result if x]  #['14', '六月', '18', '9', '56', '下午']  日期/月份/年
        a.append(a1)

    #将汉字月份适配成数字月份
    month=['一月','二月','三月','四月','五月','六月','七月','八月','九月','十月','十一月','十二月']
    tt=[]
    for aa in a:
        j=0
        for i in month:
            j += 1
            if i==aa[1]:
                break

    #适配成年月日格式月份
        if aa[-1]=="上午":
            if j>9:   #给分钟补上0
                t='20'+aa[2]+'-'+str(j)+'-'+aa[0]+' '+aa[3]+':'+aa[4]+':'+'00'
            else:
                t = '20' + aa[2] + '-' + '0'+str(j) + '-' + aa[0] + ' ' + aa[3] + ':' + aa[4] + ':' + '00'

        elif aa[-1]=="下午" and int(aa[-3])==12:
            if j>9:   #给分钟补上0
                t='20'+aa[2]+'-'+str(j)+'-'+aa[0]+' '+aa[3]+':'+aa[4]+':'+'00'
            else:
                t = '20' + aa[2] + '-' + '0'+str(j) + '-' + aa[0] + ' ' + aa[3] + ':' + aa[4] + ':' + '00'
        else:
            if j>9:  #给分钟补上0
                t = '20' + aa[2] + '-' + str(j) + '-' + aa[0] + ' ' +str(int(aa[3])+12) + ':' + aa[4] + ':' + '00'
            else:
                t = '20' + aa[2] + '-' + '0'+str(j) + '-' + aa[0] + ' ' + str(int(aa[3]) + 12) + ':' + aa[4] + ':' + '00'
        tt.append(t)
    return tt

def delta(t1,t2):
    date1=change_time(t1)
    date2=change_time(t2)
    # 将xlrd的对象转化为xlwt的对象
    xl_w=open_workbook(filepath)
    excel_w = copy(xl_w)
    # 获取要操作的sheet
    table_w = excel_w.get_sheet(0)

    delta_l=[]
    j=0
    for d1,d2 in zip(date1,date2):
        delta=datetime.strptime(d1, '%Y-%m-%d %H:%M:%S')-datetime.strptime(d2, '%Y-%m-%d %H:%M:%S')
        j+=1
        # 写入excel文件
        table_w.write(j,table.ncols-1,str(delta))
    table_w.write(0, table.ncols - 1, "bug生命周期")
    #os.remove(filepath)
    excel_w.save(r"E:\D\\test1.xls")

dd=delta(update_time,create_time)

