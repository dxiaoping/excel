#下面这些变量需要您根据自己的具体情况选择
#coding:utf8
biaotou=['学号','姓名','联系电话']
#在哪里搜索多个表格
filelocation = "C:\\Users\\16935\\Desktop\\merge2\\"
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

a=['14计科','14软件','14物联','14数学','14信科',
       '15计科','15软件','15数学','15信科',
       '16计科','16软件','16物联','16数学','16信科',
       '17计科','17软件','17物联','17数学','17信科']
sheet=['14计科','14软件','14物联','14数学','14信科',
       '15计科','15软件','15数学','15信科',
       '16计科','16软件','16物联','16数学','16信科',
       '17计科','17软件','17物联','17数学','17信科']
sh=sheet
#实现读写数据
#下面是将所有文件读数据到三维列表cell[][][]中（不包含表头）

import xlrd
import xlwt
filename=xlwt.Workbook()
'''top = filearray[3]
filearray[3] = filearray[2]
filearray[2] = filearray[1]
filearray[1] = filearray[0]
filearray[0] = top'''
for sheets in range(0,len(a)):
    for i in range(ge):
        fname = filearray[i]
        bk = xlrd.open_workbook(fname)
        try:
            sh[sheets]=bk.sheet_by_name(a[sheets])
        except:
            print ("在文件%s中没有找到sheet1，读取文件数据失败,要不你换换表格的名字？" %fname)
        # sh = bk.sheet_by_name("sheet1")
        nrows=sh[sheets].nrows
        ncols = sh[sheets].ncols
        matrix[i] = [0]*(ncols)
        for m in range(ncols):
            matrix[i][m] = ["0"]*nrows
        for k in range(0,ncols):
            for j in range(0,nrows):
                matrix[i][k][j]=sh[sheets].cell(j,k).value
                # print(matrix)
    # 下面是写数据到新的表格test.xls中哦
    sheet[sheets]=filename.add_sheet(a[sheets])
    #下面是把表头写上
    # for i in range(0,len(biaotou)):
    #     sheet.write(i,0,biaotou[i])
    # 求和前面的文件一共写了多少行
    zh=3
    z=0
    for j in range(0, 3):
        for k in range(len(matrix[i][j])):
            sheet[sheets].write(k, z, matrix[i][j][k])
        z=z+1
    for i in range(ge):
        for j in range(3,len(matrix[i])):#列号
            if((matrix[i][j][0] == '总分') | (matrix[i][j][0] == '寝室总分')):
                for k in range(len(matrix[i][j])):#行号
                    # ----------改表头----------#
                    # if(matrix[i][j][0] == '总分'):
                    #     matrix[i][j][0] = '寝室总分'
                    # -------加总分--------#
                    # if((matrix[i][j][0] == '总分') & (k > 0)):
                    #     matrix[i][j][0] == '寝室总分'
                    #     sun = 0.0
                    #     for q in range(3,8):
                    #         if(matrix[i][q][k] == '优差寝'):
                    #             sun ='寝室总分'
                    #             break
                    #         if(type(matrix[i][q][k]) != type(0.0)):
                    #             matrix[i][q][k] =0.0
                    #         sun = matrix[i][q][k] + sun
                    #     matrix[i][j][k] = sun
                    sheet[sheets].write(k,zh,matrix[i][j][k])
                zh=zh+1
        #---------新建列计算百分比------#
        # matrix[i][4][0] = '百分比'
        # sheet[sheets].write(0, 5, matrix[i][4][0])
        # for q in range(1,len(matrix[i][2])):
        #     if(type(matrix[i][3][q]) == type(0.0)):
        #         if(matrix[i][3][q] > 0):
        #             if (type(matrix[i][4][q]) == type(0.0)):
        #                 if(matrix[i][4][q] > 0):
        #                     matrix[i][4][q] = matrix[i][3][q]/matrix[i][4][q]
        #                     sheet[sheets].write(q, 5, matrix[i][4][q])
    print("我已经将%d个文件合并成1个文件，并命名为%s.xls.快打开看看正确不？"%(ge,file))
filename.save(filedestination+file+".xls")
