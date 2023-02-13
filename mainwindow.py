from PyQt5 import QtWidgets
from ui_mainwindow import Ui_MainWindow
import xlrd
import copy
import pathlib
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QListView, QMessageBox, QTableWidgetItem



class MainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setupUi(self)


        # self.Title.setText("hello Python")

        self.pushButton.setShortcut('Return')
        self.pushButton.clicked.connect(self.onpushButtonClicked)
        # self.World.clicked.connect(self.onWorldClicked)
        # self.China.clicked.connect(self.onChinaClicked)
        # self.lineEdit.textChanged.connect(self.onlineEditTextChanged)


    def onpushButtonClicked(self):
        findVal = str(self.lineEdit_3.text())
        # print(findVal)
        if findVal != "":
            self.checkvalue(findVal)
        else:
            QMessageBox.information(self, '抱歉', '请输入需要查询的IP地址')

    # def onWorldClicked(self, remark):
    #     self.Title.setText("Hello World")


    # def onChinaClicked(self):
    #     # self.Title.setText("Hello China")
    #     findVal = self.lineEdit_3.text()
    #     # print(findVal)
    #     if findVal != "":
    #         self.checkvalue(findVal)


    # def onlineEditTextChanged(self, p_str):
    #     self.Title.setText(p_str)

    def printFinder(val):
        print(val)

    def getusefile(self):
        # 查当前目录下所有xls xlsx文件，返回文件名列表
        usefile = []
        excelfile = sorted(pathlib.Path('.').glob('**/*.xls*'))
        usefile = [str(tpfile) for tpfile in excelfile]
        return copy.deepcopy(usefile)

    def rdusefile(self,fileName, checkvalue):
        # 读一个文件，并在文件单元格中查找目标数据，如果找到就返回文件名及数据
        data = xlrd.open_workbook(fileName)  # 打开当前目录下名为 fileName 的文档

        worksheets = data.sheet_names()  # 返回book中所有工作表的名字
        findout = []
        result = []
        for filenum in range(len(worksheets)):
            # 打开excel文件的第filenum张表
            sheet_1 = data.sheets()[filenum]  # 通过索引顺序获取sheet表
            nrows = sheet_1.nrows  # 获取该sheet中的有效行数
            ncols = sheet_1.ncols  # 获取该sheet中的有效列数

            # 读取文件数据
            for rowNum in range(0, nrows):
                for colNum in range(0, ncols):
                    if checkvalue in str(sheet_1.row(rowNum)[colNum].value):
                        result = []
                        local = fileName.split('.')
                        result.append(fileName + "表" + worksheets[filenum])
                        for cnt in range(0, ncols):
                            if sheet_1.row(rowNum)[cnt].ctype == 2:
                                result.append(str(int(sheet_1.row(rowNum)[cnt].value)))
                            else:
                                result.append(str(sheet_1.row(rowNum)[cnt].value))


        # return copy.deepcopy(findout)
        # print(result)
        return result

    def checkvalue(self,val):
        # 在当前目录的所有Excel表里找一个字符的位置
        # 获取当前目录内所有Excel 文件列表
        # print("咔哒咔哒,工作拉 ^。^  开始找  " + val)
        filelist = self.getusefile()
        check = []
        items=[]
        # 在每一个文件中查找目标数据
        if filelist:
            for filetp in filelist:
                # findout = self.rdusefile(filetp, val)
                # print(self.rdusefile(filetp, val))
                if self.rdusefile(filetp, val):
                    check.extend(self.rdusefile(filetp, val))

                    items.insert(2,self.rdusefile(filetp, val)[:])
                    for i in reversed(range(self.tableWidget.rowCount())):
                        self.tableWidget.removeRow(i)

                    for i in range(len(items)):
                        item = items[i]

                        row = self.tableWidget.rowCount()
                        self.tableWidget.insertRow(row)
                        for j in range(len(item)):
                            item = QTableWidgetItem(str(items[i][j]))
                            self.tableWidget.setItem(row, j, item)

                else:
                    QMessageBox.information(self, '抱歉', '没有找到！')
        # self.lineEdit.setText(str(findout))
        # if findout:
        #     self.add(findout)
        # else:
        #     QMessageBox.information(self, '抱歉', '没有找到！')





    # # 查字符在哪里
    # while (1):
    #     print("\n将要找的文件放在同一个文件夹里哦 =。=")
    #     findVal = input("请输入要找的字:")
    #     if findVal != "":
    #         checkall = checkvalue(findVal)
    #         print(str(checkall))
    #         print


