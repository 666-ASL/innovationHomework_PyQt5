import sys
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import *
from PyQt5 import QtCore
from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon,QFont
import pyqtgraph as pg
import pandas as pd
import numpy as np
import random
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
import matplotlib.pyplot as plt
from textwrap import fill
import xlwt
# import test_changeText
class mySystem(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        # self.setGeometry(300, 300, 800, 800)
        self.resize(1500,1000)
        self.setWindowTitle('折旧折耗一体化系统')
        self.setWindowIcon(QIcon('./Icon.jpg'))
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
        self.tab5 = QWidget()
        self.tab6 = QWidget()
        self.tabWidget = QTabWidget(self.centralWidget)
        self.tabWidget.setGeometry(QtCore.QRect(0,60,self.width(), self.height()-20))
        self.tabWidget.addTab(self.tab1, "")
        self.tabWidget.addTab(self.tab5, "")
        self.tabWidget.addTab(self.tab2, "")
        self.tabWidget.addTab(self.tab3, "")
        self.tabWidget.addTab(self.tab4, "")
        self.tabWidget.addTab(self.tab6, "")
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

        ### 表五
        self.table5 = QTableWidget(self.tab5)
        self.table5.setGeometry(QtCore.QRect(0, 0, self.width(), self.height()-60))
        # self.table1.move(20,150)
        self.table5.setRowCount(0)
        self.table5.setColumnCount(0)
        self.table5.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table5.raise_()
        # self.table5.setVisible(False)

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

        # self.centralWidget_2 = QWidget(self.tab6)
        # self.centralWidget_2.setObjectName("centralWidget_2")

        ###在tab6中添加布局
        self.gridLayoutWidget = QWidget(self.tab6)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(0, 0, self.width(), self.height()))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")
        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        ####选择区域
        self.groupBox = QGroupBox(self.gridLayoutWidget)
        self.groupBox.setTitle("选择参数")
        self.groupBox.setGeometry(QtCore.QRect(0, 0, 300, 300))
        self.gridLayout.addWidget(self.groupBox, 0, 0, 1, 1)
        # 画图区域
        self.groupBox_Draw = QGroupBox(self.gridLayoutWidget)
        self.groupBox_Draw.setTitle("画图区域")
        self.groupBox_Draw.setGeometry(QtCore.QRect(310, 0, self.width() - 300, 300))
        self.draw_grid = QGridLayout(self.groupBox_Draw)
        self.gridLayout.addWidget(self.groupBox_Draw, 0, 1, 1, 2)

        # 输出区域
        self.groupBox_Calculate = QGroupBox(self.gridLayoutWidget)
        self.groupBox_Calculate.setTitle("数据输出区域")
        self.groupBox_Calculate.setGeometry(QtCore.QRect(0, 310, self.width(), self.height()))
        self.gridLayout.addWidget(self.groupBox_Calculate, 1, 0, 2, 3)

        self.out_data_table = QTableWidget(self.groupBox_Calculate)
        self.out_data_table.setGeometry(QtCore.QRect(0, 0, self.width(), self.height() - 60))
        # self.table4.move(20,150)
        self.out_data_table.setRowCount(0)
        self.out_data_table.setColumnCount(0)
        self.out_data_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.out_data_table.raise_()


        ###groupBox中添加布局
        self.formLayoutWidget = QWidget(self.groupBox)
        self.formLayoutWidget.setGeometry(QtCore.QRect(20, 20, self.groupBox.width() , self.groupBox.height() ))
        self.formLayoutWidget.setObjectName("formLayoutWidget")
        self.formLayout = QtWidgets.QFormLayout(self.formLayoutWidget)
        self.formLayout.setContentsMargins(0, 0, 0, 0)
        self.formLayout.setObjectName("formLayout")
        ##在formLayoutWidget中添加控件
        # self.lab_type = QtWidgets.QLabel(self.formLayoutWidget)
        # self.lab_type.setText('曲线类型')
        # self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.lab_line)

        self.lab_line = QtWidgets.QLabel(self.formLayoutWidget)
        self.lab_line.setObjectName("lab_line")
        self.lab_line.setText('曲线类别:')
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.LabelRole, self.lab_line)

        self.lab_output = QLabel(self.formLayoutWidget)
        self.lab_output.setObjectName("lab_output")
        self.lab_output.setText('输出类型:')
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.LabelRole, self.lab_output)

        self.lab_choose = QLabel(self.formLayoutWidget)
        self.lab_choose.setObjectName("lab_choose")
        self.lab_choose.setText('选择:')
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.LabelRole, self.lab_choose)

        # self.cb_type = QComboBox(self.formLayoutWidget)
        # self.cb_type = QComboBox(self.formLayoutWidget)
        # self.cb_type.addItems(['折线图', '曲线图'])
        # self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.cb_type)

        self.cb_line = QComboBox(self.formLayoutWidget)
        self.cb_line.setObjectName("cb_line")
        self.cb_line.addItems(['储量曲线', '折耗率曲线', '产量曲线', '新增资产曲线'])
        self.formLayout.setWidget(0, QtWidgets.QFormLayout.FieldRole, self.cb_line)

        # 输出的类型
        self.cb_output = QComboBox(self.formLayoutWidget)
        self.cb_output.setObjectName("cb_output")
        self.cb_output.addItems(['按年', '按半年', '按季度', '按月'])
        self.cb_output.currentIndexChanged.connect(self.selectedActions)
        self.formLayout.setWidget(1, QtWidgets.QFormLayout.FieldRole, self.cb_output)

        # 输出的按什么划分
        self.cb_choose = QComboBox(self.formLayoutWidget)
        self.cb_choose.setObjectName("cb_choose")
        self.cb_choose.addItems(['一整年'])
        # self.cb_choose.activated.connect(self.selectedActions)
        self.formLayout.setWidget(2, QtWidgets.QFormLayout.FieldRole, self.cb_choose)

        ##画图
        self.btn_calculate = QPushButton(self.formLayoutWidget)
        self.btn_calculate.setObjectName("btn_calculate")
        self.btn_calculate.setText("开始画图")
        self.formLayout.setWidget(3, QtWidgets.QFormLayout.SpanningRole, self.btn_calculate)
        self.btn_calculate.clicked.connect(self.btn_draw_click)
        plt.rcParams['font.sans-serif'] = ['SimHei']
        self.fig = plt.Figure()
        self.canvas = FigureCanvas(self.fig)
        self.axes = self.fig.add_subplot(111)
        self.fig.subplots_adjust(right=0.7)

        # 保存图像
        self.btn_save_fig = QPushButton(self.formLayoutWidget)
        self.btn_save_fig.setText('保存图像')
        self.formLayout.setWidget(4, QtWidgets.QFormLayout.SpanningRole, self.btn_save_fig)
        self.btn_save_fig.clicked.connect(self.save_fig)

        ##输出计算
        self.btn_OutPut = QPushButton(self.formLayoutWidget)
        self.btn_OutPut.setObjectName("btn_OutPut")
        self.btn_OutPut.setText("输出计算")
        self.formLayout.setWidget(5, QtWidgets.QFormLayout.SpanningRole, self.btn_OutPut)
        self.btn_OutPut.clicked.connect(self.btn_output_click)

        # 保存输出表格
        self.btn_save = QPushButton(self.formLayoutWidget)
        self.btn_save.setText('保存表格')
        self.formLayout.setWidget(6, QtWidgets.QFormLayout.SpanningRole, self.btn_save)
        self.btn_save.clicked.connect(self.outputfile)

        ############################################################################################################

        ## 宫格布局，使表格随窗口缩放
        self.grid = QGridLayout(self.centralWidget) # 中心窗口网格布局
        self.grid.addWidget(self.tabWidget, 0, 0, 1, 1)
        self.grid_tab1 = QGridLayout(self.tab1) # 表格窗口网格布局
        self.grid_tab5 = QGridLayout(self.tab5)
        self.grid_tab2 = QGridLayout(self.tab2)
        self.grid_tab3 = QGridLayout(self.tab3)
        self.grid_tab4 = QGridLayout(self.tab4)
        self.grid_tab6 = QGridLayout(self.tab6)
        self.grid_out_data = QGridLayout(self.groupBox_Calculate)

        self.grid_tab1.addWidget(self.table1, 0, 0, 1, 1)
        self.grid_tab5.addWidget(self.table5, 0, 0, 1, 1)
        self.grid_tab2.addWidget(self.table2, 0, 0, 1, 1)
        self.grid_tab3.addWidget(self.table3, 0, 0, 1, 1)
        self.grid_tab4.addWidget(self.table4, 0, 0, 1, 1)
        self.grid_tab6.addWidget(self.gridLayoutWidget,0,0,1,1)
        self.grid_out_data.addWidget(self.out_data_table,0,0,1,1)

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
        self.tabWidget.tabBarClicked.connect(self.tab5Click)
        self.tabWidget.tabBarClicked.connect(self.tab2Click)
        self.tabWidget.tabBarClicked.connect(self.tab3Click)
        self.tabWidget.tabBarClicked.connect(self.tab4Click)
        # self.changeEvent()
    def tab1Click(self):
        self.show_data_table(self.data[0],self.table1)
    def tab5Click(self):
        self.show_data_table(self.resevers,self.table5)
    def tab2Click(self):
        self.show_data_table(self.data[1],self.table2)
    def tab3Click(self):
        self.show_data_table(self.data[2],self.table3)
    def tab4Click(self):
        self.show_data_table(self.data[3],self.table4)


    ##下拉框的索引改变
    def selectedActions(self):
        i = self.cb_output.currentIndex()
        if i == 0:
            self.cb_choose.clear()
            self.lab_choose.setText('选择年份')
            self.cb_choose.addItem('一整年')
        elif i == 1:
            self.cb_choose.clear()
            self.lab_choose.setText('选择半年')
            self.cb_choose.addItems(['上半年', '下半年'])
        elif i == 2:
            self.cb_choose.clear()
            self.lab_choose.setText('选择季度')
            self.cb_choose.addItems(['第一季度', '第二季度', '第三季度', '第四季度'])
        else:
            self.cb_choose.clear()
            self.lab_choose.setText('选择月份')
            self.cb_choose.addItems(
                ['1月份', '2月份', '3月份', '4月份', '5月份', '6月份', '7月份', '8月份', '9月份', '10月份', '11月份', '12月份', '全部'])

    def openfile(self):
        openfile_name,type = QFileDialog.getOpenFileName(self, '选择文件', '', 'Excel files(*.xlsx , *.xls)')
        if openfile_name == '':
            return
        self.process_data(openfile_name)
        self.show_data_table(self.data[0],self.table1)
        self.show_data_table(self.resevers,self.table5)
        # print(df)

    def process_data(self,openfile_name):
        sheet = pd.read_excel(openfile_name, sheet_name=None)
        self.data = list(sheet.values())
        self.name = list(sheet.keys())
        self.name.insert(1, '初始储量')
        self.name.append('折耗测算')
        finance = set(self.data[0][self.data[0].columns.values[1]])
        self.row0_items = self.data[0].values
        self.congruent_relationship = {}  # 对应关系字典
        for row in self.row0_items:
            self.congruent_relationship[row[2]] = row[1]
        self.row1_items = self.data[1].values
        # oil, gas = 0, 0
        self.init_resverses = {}
        for row in self.row1_items:
            key = self.congruent_relationship[row[1]]
            if key not in self.init_resverses.keys():
                self.init_resverses[key] = row[2] + row[3] / 0.1255
            else:
                self.init_resverses[key] += row[2]
                self.init_resverses[key] += row[3] / 0.1255
        # self.init_resverses=sorted(self.init_resverses)
        # finance=sorted(finance)
        column = {'序号': [], '财务单元': [], '储量(万吨)': []}
        for i, fina in enumerate(finance):
            column['序号'].append(i + 1)
            column['财务单元'].append(fina)
            column['储量(万吨)'].append(self.init_resverses[fina])
        self.resevers = pd.DataFrame(column)
        # 按月统计标定储量
        self.calibration_reserves = {}  # 标定储量
        for res in self.data[2].values:
            if res[2] not in self.calibration_reserves.keys():
                self.calibration_reserves[res[2]] = [res[3]]
                for i in range(5):
                    self.calibration_reserves[res[2]].append(res[3])
            else:
                for i in range(6):
                    self.calibration_reserves[res[2]].append(res[3])
        pass

        # 按月统计产量和新增资产
        self.output = {}  # 产量
        self.new_assets = {}  # 新增资产
        for info in self.data[3].values:
            if info[2] not in self.output.keys():
                self.output[info[2]] = [info[3]]
            else:
                self.output[info[2]].append(info[3])
            if info[2] not in self.new_assets.keys():
                self.new_assets[info[2]] = [info[4]]
            else:
                self.new_assets[info[2]].append(info[4])
        pass

        # 单位折耗率=月产量/标定储量
        # 综合折耗率=各月份折耗率之和
        self.depletion_rate = {}
        self.comprehensive_depletion_rate = {}
        for key in self.output.keys():
            out = np.array(self.output[key])
            reservers = np.array(self.calibration_reserves[key])
            result = out / reservers
            rate_sum = result.sum()
            # result = list(result)
            # result.append(rate_sum)

            self.depletion_rate[key] = result  # 一维数组
            self.comprehensive_depletion_rate[key] = rate_sum
            # self.depletion_rate[key] = np.array(self.output[key])/np.array([self.calibration_reserves[key]]) # 二维数组[1x12]

        pass


    def show_data_table(self,data,table):
        '''
        :param data: 表格数据
        :param table: 所在表格
        '''
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
        self.tabWidget.setTabText(4, self.name[4])
        self.tabWidget.setTabText(5, self.name[5])
        self.tabWidget.setVisible(True)
        # table.setVisible(True)
        table.verticalHeader().setVisible(False) # 取消行号

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

    def btn_draw_click(self):
        line_type = self.cb_line.currentText()
        output_type = self.cb_output.currentText()
        choose = self.cb_choose.currentText()
        self.plot_chart(line_type,output_type,choose)

    def plot_chart(self,line_type,output_type,choose):
        '''
        :param line_type: 曲线类型
        :param output_type: 输出类型，按年，按半年，按季度，按月
        :param choose: 根据输出类型做选择
        :return:
        '''
        self.axes.cla() # 清空画布
        if line_type=='储量曲线':
            data = self.calibration_reserves
            self.axes.set_ylabel('储量')
            pass
        elif line_type == '折耗曲线':
            data = self.init_resverses
            self.axes.set_ylabel('折耗')
            pass
        elif line_type == '产量曲线':
            data = self.output
            self.axes.set_ylabel('产量')
            pass
        elif line_type == '折耗率曲线':
            data = self.depletion_rate
            self.axes.set_ylabel('折耗率')
            pass
        elif line_type == '新增资产曲线':
            data = self.new_assets
            self.axes.set_ylabel('新增资产')
            pass

        # 使用pyqtgraph作图，不美观
        # self.fig = pg.PlotWidget()
        # r_symbol = random.choice(['o', 's', 't', 't1', 't2', 't3', 'd', '+', 'x', 'p', 'h', 'star'])
        # r_color = random.choice(['b', 'g', 'r', 'c', 'm', 'y', 'k', 'd', 'l', 's'])
        # for key in data:
        #     self.fig.plot(data[key],pen=pg.mkPen(color="r",width=5),symbol=r_symbol, symbolBrush=r_color)
        # self.draw_grid.addWidget(self.fig)

        if output_type == '按半年':
            if choose == '上半年':
                data = self.dict_slice(data, 0, 6)
                x_ticks = np.arange(1, 7, 1)
            elif choose == '下半年':
                data = self.dict_slice(data, 6, 12)
                x_ticks = np.arange(7, 13, 1)
        elif output_type == '按季度':
            if choose == '第一季度':
                data = self.dict_slice(data, 0, 3)
                x_ticks = np.arange(1, 4, 1)
            elif choose == '第二季度':
                data = self.dict_slice(data, 3, 6)
                x_ticks = np.arange(4, 7, 1)
            elif choose == '第三季度':
                data = self.dict_slice(data, 6, 9)
                x_ticks = np.arange(7, 10, 1)
            elif choose == '第四季度':
                data = self.dict_slice(data, 9, 12)
                x_ticks = np.arange(10, 13, 1)
        elif output_type == '按月':
            if choose == '1月份':
                data = self.dict_slice(data, 0, 1)
                x_ticks = np.arange(1, 2, 1)
            elif choose == '2月份':
                data = self.dict_slice(data, 1, 2)
                x_ticks = np.arange(2, 3, 1)
            elif choose == '3月份':
                data = self.dict_slice(data, 2, 3)
                x_ticks = np.arange(3, 4, 1)
            elif choose == '4月份':
                data = self.dict_slice(data, 3, 4)
                x_ticks = np.arange(4, 5, 1)
            elif choose == '5月份':
                data = self.dict_slice(data, 4, 5)
                x_ticks = np.arange(5, 6, 1)
            elif choose == '6月份':
                data = self.dict_slice(data, 5, 6)
                x_ticks = np.arange(6, 7, 1)
            elif choose == '7月份':
                data = self.dict_slice(data, 6, 7)
                x_ticks = np.arange(7, 8, 1)
            elif choose == '8月份':
                data = self.dict_slice(data, 7, 8)
                x_ticks = np.arange(8, 9, 1)
            elif choose == '9月份':
                data = self.dict_slice(data, 8, 9)
                x_ticks = np.arange(9, 10, 1)
            elif choose == '10月份':
                data = self.dict_slice(data, 9, 10)
                x_ticks = np.arange(10, 11, 1)
            elif choose == '11月份':
                data = self.dict_slice(data, 10, 11)
                x_ticks = np.arange(11, 12, 1)
            elif choose == '12月份':
                data = self.dict_slice(data, 11, 12)
                x_ticks = np.arange(12, 13, 1)
            elif choose == '全部':
                x_ticks = np.arange(1, 13, 1)
        else:
            x_ticks = np.arange(1, 13, 1)
        self.axes.set_title(f'{line_type}({choose})')
        # self.out_data = data
        legends = []
        barw=1
        data_show = False
        xt = x_ticks
        val = np.array(list(data.values())).flatten()
        for key in data:
            legends.append(key)
            if len(data[key]) == 1:
                self.axes.bar(x_ticks, data[key])
                x_ticks = x_ticks + barw
                self.axes.set_xlabel(choose)
                self.axes.set_xticks([])
                data_show = True
            else:
                self.axes.plot(x_ticks,data[key])
        legends = [fill(legend,10) for legend in legends]
        self.axes.legend(legends,loc=3,bbox_to_anchor=(1.02, 0),borderaxespad=0)
        if choose[-2::] != '月份':
            self.axes.set_xlabel('月份')
            self.axes.set_xticks(x_ticks)
        if data_show:
            for x,y in zip(np.arange(xt,x_ticks,barw),val):
                self.axes.text(x,y+1,y,ha='center', va='bottom', fontsize=10)
        self.axes.grid(alpha=0.5)
        self.canvas.draw()
        self.draw_grid.addWidget(self.canvas)

    def dict_slice(self,dict,start,end):
        keys = dict.keys()
        dict_slice = {}
        for k in list(keys):
            dict_slice[k] = dict[k][start:end]
        return dict_slice

    def btn_output_click(self):
        output_type = self.cb_output.currentText()
        choose = self.cb_choose.currentText()
        self.show_final_data(output_type,choose)
        pass

    def show_final_data(self,output_type,choose):
        self.out_data_table.setRowCount(0)
        self.out_data_table.clearContents()
        table_num = 1 # 根据output_type和choose确定表的个数
        column_num = 4 # 每个表的列数
        keys = self.output.keys()
        rate = np.array(list(self.depletion_rate.values()))
        assets = np.array(list(self.new_assets.values()))
        output = np.array(list(self.output.values()))
        rows = []
        if output_type == '按年':
            table_num = 1
            headers = ['一整年']
            self.row_data(table_num,keys,rate,assets,output,rows)

        elif output_type == '按半年':
            table_num = 2
            #
            # rate_up = self.dict_slice(self.depletion_rate,0,6) # 上半年单位折耗率
            # rate_down = self.dict_slice(self.depletion_rate,6,12) # 下半年单位折耗率
            # sum_rate_up = np.array(list(rate_up.values())).sum(1) # 上半年综合折耗率
            # sum_rate_up = dict(zip(keys,sum_rate_up))
            # sum_rate_down = np.array(list(rate_down.values())).sum(1) # 下半年综合折耗率
            # sum_rate_down = dict(zip(keys, sum_rate_down))
            #
            # new_assets_up = self.dict_slice(self.new_assets,0,6)
            # new_assets_down = self.dict_slice(self.new_assets,6,12)
            # sum_assets_up = np.array(list(new_assets_up.values())).sum(1) # 上半年新增资产合计
            # sum_assets_up = dict(zip(keys, sum_assets_up))
            # sum_assets_down = np.array(list(new_assets_down.values())).sum(1) # 下半年新增资产合计
            # sum_assets_down = dict(zip(keys, sum_assets_down))
            #
            # out_up = self.dict_slice(self.output,0,6)
            # out_down = self.dict_slice(self.output, 6, 12)
            # sum_out_up = np.array(list(out_up.values())).sum(1) # 上半年产量合计
            # sum_out_up = dict(zip(keys, sum_out_up))
            # sum_out_down = np.array(list(out_down.values())).sum(1) # 下半年产量合计
            # sum_out_down = dict(zip(keys, sum_out_down))

            # 组织数据
            # self.out_data = {}
            headers = ['上半年', '下半年']
            self.row_data(table_num,keys,rate,assets,output,rows)

        elif output_type == '按季度':
            table_num = 4
            headers = ['第一季度', '第二季度', '第三季度', '第四季度']
            self.row_data(table_num,keys,rate,assets,output,rows)

        elif output_type == '按月':
            table_num = 12
            column_num = 5
            headers = ['一月', '二月', '三月', '四月','五月', '六月', '七月', '八月','九月', '十月', '十一月', '十二月']
            self.row_data(table_num,keys,rate,assets,output,rows)

        finace = ['财务单元','综合折耗率','新增资产合计','产量合计','单位折耗']

        self.out_data_table.setColumnCount(table_num*column_num)
        self.out_data_table.setRowCount(13)
        font = QFont()
        font.setBold(True)
        for i in range(0,13):
            for j in range(0,table_num*column_num):

                if i == 0 and j % column_num != 0:
                    dataItem = ''
                    newItem = QTableWidgetItem(dataItem)
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    newItem.setFont(font)
                    self.out_data_table.setColumnWidth(j,240)
                    self.out_data_table.setItem(i, j, newItem)
                elif i == 0 and j % column_num == 0:
                    dataItem = headers[j//column_num]
                    newItem = QTableWidgetItem(dataItem)
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    newItem.setFont(font)
                    self.out_data_table.setSpan(0, j, 1, column_num)
                    self.out_data_table.setColumnWidth(j, 240)
                    self.out_data_table.setItem(i, j, newItem)
                elif i == 1:
                    dataItem = finace[j%column_num]
                    newItem = QTableWidgetItem(dataItem)
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    newItem.setFont(font)
                    self.out_data_table.setItem(i,j,newItem)
                else:
                    dataItem = str(rows[i-2][j])
                    newItem = QTableWidgetItem(dataItem)
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    self.out_data_table.setItem(i,j,newItem)

        self.out_data_table.verticalHeader().setVisible(False)
        self.out_data_table.horizontalHeader().setVisible(False)

    def row_data(self,table_num,keys,rate,assets,output,rows):
        m = int(12 / table_num)
        for i, key in enumerate(keys):
            n = 0
            row = []
            for j in range(0, table_num):
                row.append(key)
                if table_num == 12:
                    row.append(rate[i][0:n + m].sum())
                else:
                    row.append(rate[i][n:n + m].sum())
                row.append(assets[i][n:n + m].sum())
                row.append(output[i][n:n + m].sum())
                if table_num == 12:
                    row.append(rate[i][j])
                n = n + m
            rows.append(row)

    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Escape:
            self.close()

    def outputfile(self):
        text=self.cb_output.currentText()
        filename, type = QFileDialog.getSaveFileName(None, '保存文件', f'测算结果({text})', 'Excel 97-2003 工作簿(*.xls)')
        if filename == '' or type == '':
            return
        wbk = xlwt.Workbook()
        self.sheet = wbk.add_sheet("sheet", cell_overwrite_ok=True)
        self.add2()
        wbk.save(filename)

    def add2(self):
        row = 0
        col = 0
        for i in range(self.out_data_table.columnCount()):
            for j in range(self.out_data_table.rowCount()):
                try:
                    text = str(self.out_data_table.item(row, col).text())
                    self.sheet.write(row, col, text)
                    row += 1
                except AttributeError:
                    row += 1
            row = 0
            col += 1

    def save_fig(self):
        screen = QApplication.primaryScreen()
        pix = screen.grabWindow(self.canvas.winId())
        text = self.cb_line.currentText()
        text_choose = self.cb_choose.currentText()
        fd, type = QFileDialog.getSaveFileName(None, "保存图片", f"{text}({text_choose})", "*.jpg;;*.png;;All Files(*)")
        if fd == '':
            return
        pix.save(fd)


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
    click = mySystem()
    click.show()


    sys.exit(app.exec_())