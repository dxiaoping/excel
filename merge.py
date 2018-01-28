#下面这些变量需要您根据自己的具体情况选择
#coding:utf8
biaotou=['学号','姓名','联系电话']
#在哪里搜索多个表格
filelocation = "C:\\Users\\16935\\Desktop\\merge\\"
#当前文件夹下搜索的文件名后缀
fileform="xls"
#将合并后的表格存放到的位置
filedestination="C:\\Users\\16935\\Desktop\\merge1\\"
#合并后的表格命名为file
file="test"
#首先查找默认文件夹下有多少文档需要整合
import glob
from numpy import *
filearray=[]
for filename in glob.glob(filelocation+"*."+fileform):
    filearray.append(filename)
    print(filearray)
#以上是从pythonscripts文件夹下读取所有excel表格，并将所有的名字存储到列表filearray
print("在默认文件夹下有%d个文档哦"%len(filearray))
ge=len(filearray)
matrix = [None]*ge
#实现读写数据
#下面是将所有文件读数据到三维列表cell[][][]中（不包含表头）
import xlrd
for i in range(ge):
    fname=filearray[i]
    bk=xlrd.open_workbook(fname)
    try:
        sh=bk.sheet_by_name("Sheet1")
    except:
        print ("在文件%s中没有找到sheet1，读取文件数据失败,要不你换换表格的名字？" %fname)
    # sh = bk.sheet_by_name("sheet1")
    nrows=sh.nrows
    ncols = sh.ncols
    matrix[i] = [0]*(ncols-1)

    for m in range(ncols-1):
        matrix[i][m] = ["0"]*nrows
    for k in range(1,ncols):
        for j in range(0,nrows):
            matrix[i][k-1][j]=sh.cell(j,k).value
            print(matrix)
#下面是写数据到新的表格test.xls中哦
import xlwt
filename=xlwt.Workbook()
sheet=filename.add_sheet("Sheet1")
#下面是把表头写上
for i in range(0,len(biaotou)):
    sheet.write(i,0,biaotou[i])
#求和前面的文件一共写了多少行
zh=1
for i in range(ge):
    for j in range(len(matrix[i])):
        for k in range(len(matrix[i][j])):
            sheet.write(k,zh,matrix[i][j][k])
        zh=zh+1
print("我已经将%d个文件合并成1个文件，并命名为%s.xls.快打开看看正确不？"%(ge,file))
filename.save(filedestination+file+".xls")

