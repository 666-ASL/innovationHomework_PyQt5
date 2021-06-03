import sys
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import *
from PyQt5 import QtCore
import pandas as pd
import numpy as np
# import test_changeText
class test(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # self.setGeometry(300, 300, 800, 800)
        self.resize(1000,1000)
        self.setWindowTitle('折旧折耗一体化系统')
        self.centralWidget = QWidget(self)
        # self.grid = QGridLayout()
        # 测试信号与槽
        # self.btn1 = QPushButton('按钮1',self)
        # self.btn1.move(20,60)
        # self.label1 = QLabel('按一下试试',self)
        # self.label1.move(20,30)
        # self.btn1.clicked.connect(self.btnclick)
        # self.btn1.clicked.connect(lambda:self.label1.setText('按键1被点击'))

        # 测试打开文件
        # self.btn2 = QPushButton('打开本地文件',self)
        # self.btn2.move(20,100)
        # self.btn2.clicked.connect(self.openfile)

        # 菜单栏 “导入数据” 设置

        ## 选项卡界面
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        self.tab4 = QWidget()
        self.tabWidget = QTabWidget(self.centralWidget)
        self.tabWidget.setGeometry(QtCore.QRect(0,60,self.width(), self.height()-20))
        self.tabWidget.addTab(self.tab1, "")
        self.tabWidget.addTab(self.tab2, "")
        self.tabWidget.addTab(self.tab3, "")
        self.tabWidget.addTab(self.tab4, "")
        self.tabWidget.setVisible(False)

        self.setCentralWidget(self.centralWidget)
        # self.setLayout()
        # self.setCentralWidget(self)
        ## 表格设置
        ### 表一
        self.table1 = QTableWidget(self.tab1)
        self.table1.setGeometry(QtCore.QRect(0, 0, self.width(), self.height()-60))
        # self.table1.move(20,150)
        self.table1.setRowCount(0)
        self.table1.setColumnCount(0)
        self.table1.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table1.raise_()
        # self.table1.setVisible(False)
        ### 表二
        self.table2 = QTableWidget(self.tab2)
        self.table2.setGeometry(QtCore.QRect(0, 0, self.width(),self.height()-60))
        # self.table2.move(20,150)
        self.table2.setRowCount(0)
        self.table2.setColumnCount(0)
        self.table2.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table2.raise_()
        # self.table2.setVisible(False)
        ### 表三
        self.table3 = QTableWidget(self.tab3)
        self.table3.setGeometry(QtCore.QRect(0, 0, self.width(),self.height()-60))
        # self.table3.move(20,150)
        self.table3.setRowCount(0)
        self.table3.setColumnCount(0)
        self.table3.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table3.raise_()
        # self.table3.setVisible(False)
        ### 表四
        self.table4 = QTableWidget(self.tab4)
        self.table4.setGeometry(QtCore.QRect(0, 0, self.width(),self.height()-60))
        # self.table4.move(20,150)
        self.table4.setRowCount(0)
        self.table4.setColumnCount(0)
        self.table4.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table4.raise_()
        # self.table4.setVisible(False)

        ## 宫格布局，使表格随窗口缩放
        self.grid = QGridLayout(self.centralWidget) # 中心窗口网格布局
        self.grid.addWidget(self.tabWidget, 0, 0, 1, 1)
        self.grid_tab1 = QGridLayout(self.tab1) # 表格窗口网格布局
        self.grid_tab2 = QGridLayout(self.tab2)
        self.grid_tab3 = QGridLayout(self.tab3)
        self.grid_tab4 = QGridLayout(self.tab4)
        self.grid_tab1.addWidget(self.table1, 0, 0, 1, 1)
        self.grid_tab2.addWidget(self.table2, 0, 0, 1, 1)
        self.grid_tab3.addWidget(self.table3, 0, 0, 1, 1)
        self.grid_tab4.addWidget(self.table4, 0, 0, 1, 1)

        ## 菜单栏设置与事件触发
        self.mainMenu = QMenuBar(self)
        self.mainMenu.setGeometry(QtCore.QRect(0, 0, 800, 22))
        self.setMenuBar(self.mainMenu)
        # self.menu1 = QMenu(self.mainMenu)
        # self.menu1.setTitle('导入数据')
        action = QAction(self)
        action.triggered.connect(self.openfile)
        action.setText('导入数据')
        self.mainMenu.addAction(action) #直接在菜单栏上设置事件响应
        # self.mainMenu.addMenu(self.menu1)
        # self.menu1.addAction(action)
        self.tabWidget.tabBarClicked.connect(self.tab1Click)
        self.tabWidget.tabBarClicked.connect(self.tab2Click)
        self.tabWidget.tabBarClicked.connect(self.tab3Click)
        self.tabWidget.tabBarClicked.connect(self.tab4Click)
        # self.changeEvent()
    def tab1Click(self):
        self.show_data_table(self.data[0],self.table1)
    def tab2Click(self):
        self.show_data_table(self.data[1],self.table2)
    def tab3Click(self):
        self.show_data_table(self.data[2],self.table3)
    def tab4Click(self):
        self.show_data_table(self.data[3],self.table4)

    def openfile(self):
        openfile_name = QFileDialog.getOpenFileName(self, '选择文件', '', 'Excel files(*.xlsx , *.xls)')
        sheet = pd.read_excel(openfile_name[0],sheet_name=None)
        self.data=list(sheet.values())
        self.name=list(sheet.keys())
        self.show_data_table(self.data[0],self.table1)
        # print(df)

    def show_data_table(self,data,table):
        '''
        :param data: 表格数据
        :param table: 所在表格
        '''
        # sheet = pd.read_excel(path,sheet_name=None)
        # data = pd.read_excel(path)
        # sheet_names, datas,=[], [] # 表单名，数据集
        # heads, rows, columns=[], [], [] # 表头，行，列
        # sheet_cnt = len(sheet)
        # for name, data in sheet.items():
        #     heads.append(list(data.columns))
        #     rows.append(data.values.shape[0])
        #     columns.append(data.values.shape[1])
        #     data_dict=data.to_dict(orient='records')
        #     sheet_names.append(name)
        #     datas.append(data_dict)
        # table = self.table1

        input_table_rows = data.shape[0]
        input_table_colunms = data.shape[1]
        input_table_header = data.columns.values.tolist()
        # print(input_table_rows)
        # print(input_table_colunms)
        # print(input_table_header)
        table.setColumnCount(input_table_colunms)
        table.setRowCount(input_table_rows)
        table.setHorizontalHeaderLabels(input_table_header)
        # 遍历data，将data中的每一项作为table_item加入到表格中
        for i in range(data.values.shape[0]):
            data_row = data.iloc[[i]]
            data_row_array = np.array(data_row)
            data_row_list = data_row_array.tolist()[0]
            for j in range(data.values.shape[1]):
                data_row_item = str(data_row_list[j]) # 数据必须为字符串，若是数字则在表格中不显示
                newItem = QTableWidgetItem(data_row_item)
                newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                table.setItem(i,j,newItem)
        self.tabWidget.setTabText(0, self.name[0])
        self.tabWidget.setTabText(1, self.name[1])
        self.tabWidget.setTabText(2, self.name[2])
        self.tabWidget.setTabText(3, self.name[3])
        self.tabWidget.setVisible(True)
        # table.setVisible(True)
        table.verticalHeader().setVisible(False) # 取消行号

        # self.table1.setColumnCount(input_table_colunms)
        # self.table1.setRowCount(input_table_rows)
        # self.table1.setHorizontalHeaderLabels(input_table_header)
        # for i in range(data.values.shape[0]):
        #     data_row = data.iloc[[i]]
        #     data_row_array = np.array(data_row)
        #     data_row_list = data_row_array.tolist()[0]
        #     for j in range(data.values.shape[1]):
        #         data_row_item = str(data_row_list[j])
        #         newItem = QTableWidgetItem(data_row_item)
        #         newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        #         self.table1.setItem(i, j, newItem)
        # self.table1.setVisible(True)


        # print(name,data)
    # 窗口改变事件不符合对于任意事件均有响应，不符合需求
    # def changeEvent(self,e):
    #     if e.type()==QtCore.QEvent.WindowStateChange:
    #         self.tabWidget.setGeometry(QtCore.QRect(0,30,self.width(),self.height()))
    #         self.table1.setGeometry(QtCore.QRect(0, 0, self.width(), self.height() - 60))
    #         self.table2.setGeometry(QtCore.QRect(0, 0, self.width(), self.height() - 60))
    #         self.table3.setGeometry(QtCore.QRect(0, 0, self.width(), self.height() - 60))
    #         self.table4.setGeometry(QtCore.QRect(0, 0, self.width(), self.height() - 60))
        # self.changeEvent(e)


    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.close()



if __name__ == '__main__':
    app = QApplication(sys.argv)

    # 新手教学测试
    # w = QWidget()
    # w.resize(400,400)
    # w.move(300,300)
    # w.setWindowTitle('测试窗口')
    # w.show()

    # .ui --> .py 测试
    # MainWindow = QtWidgets.QMainWindow()
    # test = test_changeText.Ui_MainWindow()
    # test.setupUi(MainWindow)
    # MainWindow.show()

    # 按键重写测试（手写窗口）
    MainWindow = QMainWindow()
    # click = test_changeText.Ui_MainWindow()
    # click.setupUi(MainWindow)
    click = test()
    click.show()


    sys.exit(app.exec_())