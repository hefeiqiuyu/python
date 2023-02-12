import time
import os
import xlrd
import copy
import pathlib
import PyQt5


def printFinder(val):
    print(val)


def getusefile():
    # 查当前目录下所有xls xlsx文件，返回文件名列表
    usefile = []
    excelfile = sorted(pathlib.Path('.').glob('**/*.xls*'))
    usefile = [str(tpfile) for tpfile in excelfile]
    return copy.deepcopy(usefile)


def rdusefile(fileName, checkvalue):
    # 读一个文件，并在文件单元格中查找目标数据，如果找到就返回文件名及数据
    data = xlrd.open_workbook(fileName)  # 打开当前目录下名为 fileName 的文档

    worksheets = data.sheet_names()  # 返回book中所有工作表的名字
    findout = []

    for filenum in range(len(worksheets)):
        # 打开excel文件的第filenum张表
        sheet_1 = data.sheets()[filenum]  # 通过索引顺序获取sheet表
        nrows = sheet_1.nrows  # 获取该sheet中的有效行数
        ncols = sheet_1.ncols  # 获取该sheet中的有效列数
        getdata = []

        # 读取文件数据
        for rowNum in range(0, nrows):
            tep1 = []
            for colNum in range(0, ncols):
                # tep1.append(sheet_1.row(rowNum)[colNum].value)
                if checkvalue in str(sheet_1.row(rowNum)[colNum].value):
                    result = []
                    local = fileName.split('.')
                    result.append("文件:" + fileName + " 的表 " + worksheets[filenum] + " 找到了 ")
                    for cnt in range(0, ncols):
                        result.append(str(sheet_1.row(rowNum)[cnt].value))
                    printFinder(result)

    return copy.deepcopy(findout)


def checkvalue(val):
    # 在当前目录的所有Excel表里找一个字符的位置
    # 获取当前目录内所有Excel 文件列表
    print("咔哒咔哒,工作拉 ^。^  开始找  " + val)
    filelist = getusefile()
    check = []
    # 在每一个文件中查找目标数据
    if filelist:
        for filetp in filelist:
            findout = rdusefile(filetp, val)
            if findout:
                check.extend(findout)
    return copy.deepcopy(check)


# 查字符在哪里
while (1):
    print("\n将要找的文件放在同一个文件夹里哦 =。=")
    findVal = input("请输入要找的字:")
    if findVal != "":
        checkall = checkvalue(findVal)
        print(str(checkall))
        print

