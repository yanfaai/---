import os

from PyQt5.QtCore import Qt, QUrl
from PyQt5.QtGui import QFont, QIcon, QPalette, QBrush, QPixmap, QStandardItem, QStandardItemModel
from PyQt5.QtWidgets import *

from PyQt5.QtWebEngineWidgets import QWebEngineView
import pandas as pd
import re
from PyQt5 import QtGui, QtCore
from pyecharts.charts import Bar, Gauge, Grid, Tab, Liquid
from pyecharts import options as opts
from pyecharts.globals import SymbolType

font = QFont()
font.setBold(True)
font.setPointSize(10)

qlist = []
jdlist=[]

class Window_(QStandardItemModel):

    def data(self, index, role=None):
        if role == QtCore.Qt.TextAlignmentRole:
            return QtCore.Qt.AlignCenter
        return QStandardItemModel.data(self, index, role)


class Ui_MainWindow(QTabWidget):
    font = QFont()
    font.setBold(True)
    font.setPointSize(11)

    def __init__(self):
        super(Ui_MainWindow, self).__init__()
        self.setWindowTitle('天工矿业统计数据视觉系统')
        self.setToolTip('开发人员，闫发爱')
        self.setWindowIcon(QIcon('jiugang.jpg'))
        self.setStyleSheet('background-color:#afb4db')
        self.setWindowOpacity(0.9)

        self.setTabsClosable(True)
        self.setStyleSheet("background-color:darkGray")
        QTabbar = QTabBar()
        QTabbar.setStyleSheet("background-color:#afb4db")
        QTabbar.setFont(font)
        self.setTabBar(QTabbar)
        self.tab_0 = QWidget()
        self.tab_0.setStyleSheet('background-color:#afb4db')
        self.tab_1 = QWidget()
        self.tab_1.setStyleSheet('background-color:#afb4db')
        self.tab_2 = QWidget()
        self.tab_2.setStyleSheet('background-color:#afb4db')

        self.addTab(self.tab_0, '天工矿业公司经营指标')
        self.addTab(self.tab_1, '天工矿业公司费用汇总(不含新洲)')
        self.addTab(self.tab_2, '其余各单位经营指标')
        scree = QDesktopWidget()
        scree.screenGeometry()
        self.setFixedSize(scree.width() - 200, scree.height() - 150)  # 固定写法


    def setUi_1(self):  # 天工费用汇总表
        self.tab_1.setGeometry(self.width() - 1850, self.height() - 1050, self.width() - 200, self.height() - 250)
        self.setTabText(1, '天工费用汇总(不含新洲)')
        self.tab_1TableView = QTableView(self.tab_1)
        self.tab_1TableView.setGeometry(0, 0, self.width(), 300)
        web = QWebEngineView(self.tab_1)
        web.setGeometry(0, 320, self.width(), 450)

        self.QHB = QHBoxLayout()
        self.Widget_01 = QWidget(self.tab_1)
        self.Widget_01.setGeometry(10, 600, self.width(), 50)

        QComboBox_0 = QComboBox(self.tab_1)
        QComboBox_0.addItems(['一季度', '二季度', '半年度', '三季度', '四季度'])
        QComboBox_1 = QComboBox(self.tab_1)
        QComboBox_1.addItems(['预算', '实际', '预算与实际'])
        QComboBox_2 = QComboBox(self.tab_1)
        QComboBox_2.addItems(['职工薪酬', '办公费用', '差旅费', '业务用车费', '管理费用', '期间费用'])
        QComboBox_4 = QComboBox(self.tab_1)
        QComboBox_4.addItems(['Bar_1', 'Pie'])
        QComboBox_3 = QComboBox(self.tab_1)
        QComboBox_3.addItems(['同比', '环比', '全年预算与实际累计'])
        QComboBox_5 = QComboBox(self.tab_1)
        QComboBox_5.addItems(['月份'])

        self.QHB.addWidget(QComboBox_5)
        self.QHB.addWidget(QComboBox_1)
        self.QHB.addWidget(QComboBox_2)
        self.QHB.addWidget(QComboBox_3)
        self.QHB.addWidget(QComboBox_0)
        self.QHB.addWidget(QComboBox_4)
        self.QHB.setSpacing(50)
        self.Widget_01.setLayout(self.QHB)

    def setUi_0(self):  # 天工经营指标
        self.tab_0.setGeometry(self.width() - 1850, self.height() - 1050, self.width() - 200, self.height() - 250)
        self.setTabText(0, '天工矿业公司经营指标')
        self.tab_0TableView = QTableView(self.tab_0)
        self.tab_0TableView.setGeometry(0, 0, self.width(), 400)
        self.tab_0web = QWebEngineView(self.tab_0)
        self.tab_0web.setGeometry(0, 400, self.width(), 250)

        self.QHB = QHBoxLayout()
        self.Widget_0 = QWidget(self.tab_0)
        self.Widget_0.setGeometry(10, 650, 1050, 50)

        self.tab_0QComboBox_1 = QComboBox(self.tab_0)
        self.tab_0QComboBox_1.addItems(['2021月预算与实际', '实际', '同期'])
        self.tab_0QComboBox_2 = QComboBox(self.tab_0)
        self.tab_0QComboBox_2.addItems(['天工费用', '天工利润','新洲产值', '钨矿产量', '钨矿销量','钨矿生产成本', '新洲收入','新洲利润', '新洲存货资金'])
        self.tab_0QComboBox_0 = QComboBox(self.tab_0)
        self.tab_0QComboBox_0.addItems(['一季度', '二季度', '半年度', '三季度', '四季度','累计'])
        self.tab_0PushButton = QPushButton('系统退出',self.tab_0)
        self.tab_0QComboBox_3 = QComboBox(self.tab_0)
        self.tab_0QComboBox_3.addItems(['同比', '环比', '全年预算与实际累计'])
        self.tab_0QComboBox_5 = QComboBox(self.tab_0)
        date_ = pd.read_excel('天工矿业公司2021年预算目标完成情况统计报表.xlsx', sheet_name=0, header=1, usecols='e:ab', nrows=16)
        self.tab_0QComboBox_5.addItems(
            [x.replace('实际', '').replace('预算', '') for x in date_.columns[::2]])  # 设定了步长，若用集合去重会无序
        self.QHB.addWidget(self.tab_0QComboBox_5)
        self.QHB.addWidget(self.tab_0QComboBox_1)
        self.QHB.addWidget(self.tab_0QComboBox_2)
        self.QHB.addWidget(self.tab_0QComboBox_3)
        self.QHB.addWidget(self.tab_0QComboBox_0)
        self.QHB.addWidget(self.tab_0PushButton)
        self.QHB.setSpacing(50)
        self.Widget_0.setLayout(self.QHB)

        self.tab_0PushButton.clicked.connect(self.close)
        self.tab_0QComboBox_5.currentTextChanged.connect(self.m)
        self.tab_0QComboBox_0.currentTextChanged.connect(self.jd)

        self.tab_0QComboBox_1.currentTextChanged.connect(self.sj)
        # self.tab_0QComboBox_2.currentTextChanged.connect(self.sj)

    def setUi_2(self):  # 各单位
        self.tab_2.setGeometry(self.width() - 1850, self.height() - 1050, self.width() - 200, self.height() - 250)

        TableView = QTableView(self.tab_2)
        TableView.setGeometry(0, 0, self.width(), 300)
        web = QWebEngineView(self.tab_2)

        web.setGeometry(0, 320, self.width(), 450)

        self.QHB = QHBoxLayout()
        self.Widget_01 = QWidget(self.tab_2)
        self.Widget_01.setGeometry(10, 600, 1050, 50)
        self.Widget_01.setStyleSheet('background-color:#afb4db')

        QComboBox_0 = QComboBox(self.tab_2)
        QComboBox_0.addItems(['一季度', '二季度', '累计', '三季度', '四季度'])
        QComboBox_1 = QComboBox(self.tab_2)
        QComboBox_1.addItems(['预算', '实际', '预算与实际'])
        QComboBox_2 = QComboBox(self.tab_2)
        QComboBox_2.addItems(['职工薪酬', '办公费用', '差旅费', '业务用车费', '管理费用', '期间费用', '利润总额'])
        QComboBox_5 = QComboBox(self.tab_2)
        QComboBox_5.addItems(['Bar_1', 'Pie'])
        QComboBox_3 = QComboBox(self.tab_2)
        QComboBox_3.addItems(['同比', '环比', '全年预算与实际累计'])
        QComboBox_4 = QComboBox(self.tab_2)
        QComboBox_4.addItems(['各单位名字（re）'])
        QComboBox_6 = QComboBox(self.tab_2)
        QComboBox_6.addItems(['月份'])

        self.QHB.addWidget(QComboBox_4)
        self.QHB.addWidget(QComboBox_6)
        self.QHB.addWidget(QComboBox_1)
        self.QHB.addWidget(QComboBox_2)
        self.QHB.addWidget(QComboBox_0)
        self.QHB.addWidget(QComboBox_3)

        self.QHB.addWidget(QComboBox_5)
        self.QHB.setSpacing(50)
        self.Widget_01.setLayout(self.QHB)

    def uidata_0(self):  # 天工费用汇总数据
        num = []
        df = pd.read_excel('天工矿业公司2020年预算目标完成情况统计报表（最终）.xlsx',
                           sheet_name=1, usecols='A,D:AS', header=1, nrows=114)

        df.columns = map(lambda x: x.replace('\n', ''), df.columns)
        df = df.drop([i for i in df.columns if re.match('.*2019.*', i)], axis=1)
        df = df.drop('2020年预算', axis=1)
        df = df.round(2)

        data = pd.read_excel('天工矿业公司2021年预算目标完成情况统计报表.xlsx',
                             sheet_name=1, usecols='A,D:AC'
                             , header=1, nrows=114)
        data = data.round(2)
        data = data.drop('2021年预算', axis=1)
        data.columns = map(lambda x: '2021年' + x, data.columns)

        data = data.loc[:, (data != 0).any(axis=0)]

        for i, k in zip(df.columns, data.columns):

            if not all([i == '费用项目', k == '费用项目']):
                num.append(df[[i]])
                num.append(data[[k]])

        df_1 = pd.concat(num, axis=1)

        df_1 = df_1.drop([x for x in df_1.columns if re.match(".*2020.*预算.*", x)], axis=1)

        df_1 = df_1.drop(["2021年预算累计", '2021年费用项目'], axis=1)

        df_1 = df_1.loc[:, (df_1 != 0).any(axis=0)]

        df_1 = df_1.set_index('费用项目')

        df_1 = df_1.loc[(df_1 != 0).any(axis=1), :]

        df_1.insert(0, '费用项目', df_1.index)


        columns = [id for id in df_1.columns]
        # model = QStandardItemModel(len(df_1), len(columns))  # 设置数据层次结构，rows行cols列
        model = Window_(len(df_1), len(columns))  # 设置数据层次结构，rows行cols列
        model.setHorizontalHeaderLabels([str(i) for i in columns])  # 设置列名
        for row in range(len(df_1)):
            for column in range(len(columns)):
                if column == 0:
                    item = QStandardItem(str(df_1.iloc[row, column]))
                    item.setFont(font)
                    item.setForeground(QBrush(QtGui.QColor('red')))  # 设置文本颜色
                    model.setItem(row, column, item)  # 设置每个位置的文本值
                else:
                    item = QStandardItem(str(df_1.iloc[row, column]))
                    item.setForeground(QBrush(QtGui.QColor('blue')))  # 设置文本颜色
                    model.setItem(row, column, item)  # 设置每个位置的文本值

        self.tab_1TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)  # 设置不可编辑
        self.tab_1TableView.setSortingEnabled(False)  # 设置是否排序
        self.tab_1TableView.verticalHeader().setVisible(False)  # 影藏垂直标题
        self.tab_1TableView.setAutoFillBackground(True)
        self.tab_1TableView.setModel(model)
        self.tab_1TableView.setColumnWidth(0, 240)  # 必须在模型设置了以后，再设置列宽与行高
        # self.tab_1TableView.setStyleSheet('background-color:#afb4db')，# 父对象设置了就别在设置样式了

        self.tab_1TableView.horizontalHeader().setFont(font)
        self.tab_1TableView.setShowGrid(True)

    def setdata_0(self):  # 天工经营指标（数据）
        data = pd.read_excel('天工矿业公司2020年预算目标完成情况统计报表（最终）.xlsx',
                             sheet_name=0, header=1, usecols='B,F:AO', nrows=13)

        data.columns = map(lambda x: x.replace('\n', '').replace(' ', ''), data.columns)

        # data.项目[4] = '费用合计' # 链式索引不能或不建议这样复值，
        data.iloc[4,0] = '天工费用'

        data.dropna(axis=0, how='any', inplace=True)

        data.drop([x for x in data.columns if re.match('.*2019.*', x)], axis=1, inplace=True)

        xp = pd.read_excel('天工矿业公司2020年预算目标完成情况统计报表（最终）.xlsx',
                           sheet_name=1, header=117, usecols='A,F:AO', nrows=38)

        xp.columns = map(lambda x: x.replace(' ', '').replace('\n', ''), xp.columns)

        # xp.项目[37] = '天工利润'
        xp.iloc[36,0] = '天工利润'

        xp.drop([n for n in xp.columns if re.match('.*2019', n)], axis=1, inplace=True)

        data = data.T
        xp_r=xp[37:]
        xp_d=xp_r.T
        data.insert(4, '5', xp_d[37])


        data = data.T
        # print(data)

        df = pd.read_excel('天工矿业公司2021年预算目标完成情况统计报表.xlsx',
                           header=1, usecols='B,E:AB', nrows=14, sheet_name=0)

        # df.项目[4] = '费用合计'
        df.iloc[4,0] = '天工费用'

        # df.项目[5] = '天工利润'
        df.iloc[5,0] = '天工利润'

        df = df.fillna(0)
        # jdlist.append(df)
        df = df.loc[:, (df != 0).any(axis=0)]

        df.columns = ['2021年' + k for k in df.columns]

        df = df.T

        df.drop([6, 0], axis=1, inplace=True)

        df = df.T

        df.index = data.index

        lis = []

        for key, value in zip(data.columns, df.columns):
            lis.append(data[[key]])
            lis.append(df[[value]])

        DF = pd.concat(lis, axis=1)

        DF = DF.drop([m for m in DF.columns if re.match('.*2020.*预算.*', m)], axis=1)

        DF = DF.drop(columns='项目')

        DF.insert(0, '项目', df['2021年项目'])
        


        DF.fillna(0, inplace=True)

        DF = DF.loc[:, (DF != float(0)).any(axis=0)]  # 这一点要仔细理解

        # DF.项目[6] = '新洲收入' # 链式索引不能或不建议这样复值，
        DF.iloc[5,0] = '新洲收入'
        DF.iloc[6,0] = '新洲利润'
        DF.iloc[7,0] = '新洲产值'
        DF.iloc[8,0] = '钨矿产量'
        DF.iloc[9,0] = '钨矿销量'
        DF.iloc[10,0] = '钨矿生产成本'
        DF.iloc[11,0] = '新洲存货资金'

        DF = DF.round(2)
        DF.drop('2021年项目', axis=1, inplace=True)

        DF_=DF.T

        DF_.insert(0,'0',DF.columns)
        DF=DF_.T

        columns = [id for id in DF.columns]
        model = Window_(len(DF), len(columns))  # 设置数据层次结构，rows行cols列
        model.setHorizontalHeaderLabels([str(i) for i in columns])  # 设置列名
        for row in range(len(DF)):
            for column in range(len(columns)):
                if column == 0 or row==0:
                    item = QStandardItem(str(DF.iloc[row, column]))
                    item.setForeground(QBrush(QtGui.QColor('green')))  # 设置文本颜色
                    item.setFont(font)
                    model.setItem(row, column, item)  # 设置每个位置的文本值
                else:
                    item = QStandardItem(str(DF.iloc[row, column]))
                    item.setForeground(QBrush(QtGui.QColor('BLUE')))  # 设置文本颜色
                    model.setItem(row, column, item)  # 设置每个位置的文本值

        self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)  # 设置不可编辑
        self.tab_0TableView.verticalHeader().setVisible(False)  # 影藏垂直标题
        self.tab_0TableView.horizontalHeader().setVisible(False) # 影藏行标题
        self.tab_0TableView.setModel(model)
        # self.tab_0TableView.setColumnWidth(0, 120)
        self.tab_0TableView.horizontalHeader().setFont(font)
        self.tab_0TableView.resizeColumnsToContents()
        self.tab_0TableView.setColumnWidth(0, 120)
        DF.rename(index={k: v for k, v in zip(DF.index, DF.项目)}, inplace=True)
      
        qlist.append(DF)
        

    def m(self):  # 月份
        # self.tab_0QComboBox_1.hide()

        try:
            DF_1 = qlist[0]
            col_2 = '2021年' + self.tab_0QComboBox_5.currentText() + '实际'
            col_1 = '2021年' + self.tab_0QComboBox_5.currentText() + '预算'

            DF_ = DF_1.loc[:, col_1:col_2]
            DF_1=DF_.T
            DF_1.drop('项目',axis=1,inplace=True)
            DF_2=DF_1.T
            DF_2['同比'] = DF_2.iloc[:, 2] - DF_2.iloc[:, 1]
            DF_2['同比']=DF_2['同比'].astype('float')
            DF_2['同比'] = DF_2['同比'].round(2)

            DF_2['较计划'] = DF_2.iloc[:, 2] - DF_2.iloc[:, 0]
            DF_2['较计划']= DF_2['较计划'].astype('float')
            DF_2['较计划'] = DF_2['较计划'].round(2)

            DF_2.insert(0, '项目', DF_2.index)
            DF_3=DF_2.T
            DF_3.insert(0,'0',DF_2.columns)
            DF_3=DF_3.T
            DF_3=DF_3.round(2)

            columns = [id for id in DF_3.columns]
            model = Window_(len(DF_3), len(columns))  # 设置数据层次结构，rows行cols列
            model.setHorizontalHeaderLabels([str(i) for i in columns])  # 设置列名
            for row in range(len(DF_3)):
                for column in range(len(columns)):
                    if column == 0 or row==0:
                        item = QStandardItem(str(DF_3.iloc[row, column]))
                        item.setForeground(QBrush(QtGui.QColor('green')))  # 设置文本颜色
                        item.setFont(font)
                        model.setItem(row, column, item)  # 设置每个位置的文本值
                    else:
                        item = QStandardItem(str(DF_3.iloc[row, column]))
                        item.setForeground(QBrush(QtGui.QColor('BLUE')))  # 设置文本颜色
                        model.setItem(row, column, item)  # 设置每个位置的文本值

            self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)  # 设置不可编辑
            self.tab_0TableView.verticalHeader().setVisible(False)  # 影藏垂直标题
            self.tab_0TableView.setModel(model)
            self.tab_0TableView.resizeColumnsToContents()
            self.tab_0TableView.setColumnWidth(0, 120)
            self.tab_0TableView.horizontalHeader().setFont(font)
            DF_4=DF_3.T
            DF_4.drop('0',axis=1,inplace=True)
            DF_5=DF_4.T

            def bar_(self):
                bar=Bar(init_opts=opts.InitOpts('950px','280px',bg_color='skyblue'))
                bar.add_xaxis(list(DF_5.columns)[1:])
                bar.add_yaxis(series_name='2021年'+self.tab_0QComboBox_5.currentText()+self.tab_0QComboBox_2.currentText(),
                              y_axis=list(DF_5.loc[self.tab_0QComboBox_2.currentText()])[1:])

                bar.set_global_opts(datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                    title_opts=opts.TitleOpts(
                                        'Bar_' + self.tab_0QComboBox_2.currentText() + '->指标单位[万元:元:标吨]'),
                                    tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='cross'),
                                    toolbox_opts=opts.ToolboxOpts())

                return bar
            def garu(self):

                try:
                    df=DF_5.iloc[:,3]/DF_5.iloc[:,1]

                    df=df[self.tab_0QComboBox_2.currentText()]

                    df=df.__round__(3)
                    print(df)

                    C=Liquid(init_opts=opts.InitOpts('950px','320px',bg_color='skyblue'))
                    C.add(self.tab_0QComboBox_2.currentText()+'完成率', [df, 1-df], is_outline_show=False, shape=SymbolType.DIAMOND)
                    C.set_global_opts(
                        title_opts=opts.TitleOpts(title='2021年'+self.tab_0QComboBox_5.currentText()+self.tab_0QComboBox_2.currentText()+'完成率'))
                    return C
                except:
                    QMessageBox.information(self.tab_0, '温馨提示', '不能除0！！！', QMessageBox.Yes,
                                            QMessageBox.No)


            tab = Tab()
            tab.add(bar_(self),'2021年'+self.tab_0QComboBox_5.currentText()+self.tab_0QComboBox_2.currentText())
            tab.add(garu(self),'2021年'+self.tab_0QComboBox_5.currentText()+self.tab_0QComboBox_2.currentText()+'预算完成率')
            tab.render('m_1.html')
            path=os.getcwd()+'/m_1.html'
            self.tab_0web.load(QUrl.fromLocalFile(path))


        except:
            QMessageBox.information(self.tab_0, '温馨提示', '所选择的月份没有数据...，请选择合适月份或及时更新数据！！！', QMessageBox.Yes, QMessageBox.No)


    def sj(self):
        # self.tab_0QComboBox_5.hide()
        # self.tab_0TableView.close()
        df_s=qlist[0]
        df_item=df_s.项目
        df_s1= df_s.drop('项目', axis=1)
        if self.tab_0QComboBox_1.currentText()=='实际':
            dfs_=df_s1.drop([x for x in df_s1.columns if re.match('.*2020.*实际',x)],axis=1)
            dfs_1 =dfs_.drop([x for x in df_s1.columns if re.match('.*2021.*预算', x)],axis=1)
            dfs_1.insert(0, '项目', df_item)

            mod = Window_(dfs_1.shape[0], dfs_1.shape[1])
            mod.setHorizontalHeaderLabels(dfs_1.columns)
            for row in range(dfs_1.shape[0]):
                for col in range(dfs_1.shape[1]):
                    if row == 0 or col == 0:
                        item = QStandardItem(dfs_1.iloc[row, col])
                        item.setForeground(QBrush(QtGui.QColor('dark')))
                        item.setFont(font)
                        mod.setItem(row, col, item)
                    else:
                        item = QStandardItem(str(dfs_1.iloc[row, col]))
                        item.setForeground(QBrush(QtGui.QColor('BLUE')))
                        mod.setItem(row, col, item)
            self.tab_0TableView.setModel(mod)
            self.tab_0TableView.horizontalHeader().setVisible(False)
            self.tab_0TableView.verticalHeader().setVisible(False)
            self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tab_0TableView.resizeColumnsToContents()
            self.tab_0TableView.setColumnWidth(0, 120)

            bar = Bar(init_opts=opts.InitOpts('1200px', '280px', bg_color='Skyblue'))
            bar.set_global_opts(toolbox_opts=opts.ToolboxOpts(is_show=True))
            bar.add_yaxis(series_name='2021年'+self.tab_0QComboBox_2.currentText() + '走势图',
                          y_axis=list(dfs_1.loc[self.tab_0QComboBox_2.currentText()])[1:])
            bar.add_xaxis(list(dfs_1.columns)[1:])
            bar.set_global_opts(datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                title_opts=opts.TitleOpts('Bar_'+self.tab_0QComboBox_2.currentText()+'->指标单位[万元:元:标吨]'),
                                tooltip_opts=opts.TooltipOpts(trigger='axis',axis_pointer_type='cross')
                                ,toolbox_opts=opts.ToolboxOpts())
            bar.set_series_opts(label_opts=opts.LabelOpts(is_show=True),

                                markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
                                                                      opts.MarkLineItem(type_="min", name="最小值"),
                                                                      opts.MarkLineItem(type_="max", name="最大值")
                                                                      ]))
            bar.render('2021_sj.html')
            path=os.getcwd()+'/2021_sj.html'
            self.tab_0web.load(QUrl.fromLocalFile(path))


        elif self.tab_0QComboBox_1.currentText()=='同期':
            dfs_ = df_s1.drop([x for x in df_s1.columns if re.match('.*2021.*实际', x)],axis=1)
            dfs_1 = dfs_.drop([x for x in df_s1.columns if re.match('.*2021.*预算', x)],axis=1)
            dfs_1.insert(0, '项目', df_item)

            mod=Window_(dfs_1.shape[0],dfs_1.shape[1])
            mod.setHorizontalHeaderLabels(dfs_1.columns)
            for row in range(dfs_1.shape[0]):
                for col in range(dfs_1.shape[1]):
                    if row==0 or col==0:
                        item=QStandardItem(dfs_1.iloc[row,col])
                        item.setForeground(QBrush(QtGui.QColor('dark')))
                        item.setFont(font)
                        mod.setItem(row,col,item)
                    else:
                        item = QStandardItem(str(dfs_1.iloc[row, col]))
                        item.setForeground(QBrush(QtGui.QColor('BLUE')))
                        mod.setItem(row,col,item)
            self.tab_0TableView.setModel(mod)
            self.tab_0TableView.horizontalHeader().setVisible(False)
            self.tab_0TableView.verticalHeader().setVisible(False)
            self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tab_0TableView.resizeColumnsToContents()
            self.tab_0TableView.setColumnWidth(0,120)

            bar=Bar(init_opts=opts.InitOpts('1200px','280px',bg_color='Skyblue'))
            bar.set_global_opts(toolbox_opts=opts.ToolboxOpts(is_show=True))
            bar.add_yaxis(series_name='2020年'+self.tab_0QComboBox_2.currentText()+'走势图',
                          y_axis=list(dfs_1.loc[self.tab_0QComboBox_2.currentText()])[1:])
            bar.add_xaxis(list(dfs_1.columns)[1:])
            bar.set_global_opts(tooltip_opts=opts.TooltipOpts(trigger='axis',axis_pointer_type='cross'),
                                datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                title_opts=opts.TitleOpts(title='Bar_'+self.tab_0QComboBox_2.currentText()+'->指标单位[万元:元:标吨]'),
                                toolbox_opts=opts.ToolboxOpts())
            bar.set_series_opts(label_opts=opts.LabelOpts(is_show=True),
                                markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
                                                                      opts.MarkLineItem(type_="min", name="最小值"),
                                                                      opts.MarkLineItem(type_="max", name="最大值")
                                                                      ]))
            bar.render('2020sj.html')
            path=os.getcwd()+'/2020sj.html'
            self.tab_0web.load(QUrl.fromLocalFile(path))

        elif self.tab_0QComboBox_1.currentText()=='2021月预算与实际':
            dfs_=df_s1.drop([x for x in df_s1.columns if re.match('.*2020.*实际',x)],axis=1)

            dfs_.insert(0, '项目', df_item)

            mod = Window_(dfs_.shape[0], dfs_.shape[1])
            mod.setHorizontalHeaderLabels(dfs_.columns)
            for row in range(dfs_.shape[0]):
                for col in range(dfs_.shape[1]):
                    if row == 0 or col == 0:
                        item = QStandardItem(dfs_.iloc[row, col])
                        item.setForeground(QBrush(QtGui.QColor('dark')))
                        item.setFont(font)
                        mod.setItem(row, col, item)
                    else:
                        item = QStandardItem(str(dfs_.iloc[row, col]))
                        item.setForeground(QBrush(QtGui.QColor('BLUE')))
                        mod.setItem(row, col, item)
            self.tab_0TableView.setModel(mod)
            self.tab_0TableView.horizontalHeader().setVisible(False)
            self.tab_0TableView.verticalHeader().setVisible(False)
            self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
            self.tab_0TableView.resizeColumnsToContents()
            self.tab_0TableView.setColumnWidth(0, 120)

            dy=dfs_.drop([x for x in dfs_.columns if re.match('.*2021.*实际',x)],axis=1)
            # print(dy)
            dj=dfs_.drop([x for x in dfs_.columns if re.match('.*2021.*预算',x)],axis=1)
            # print(dj)
            dj_1=dj.T
            # print(dj_1)
            bar = Bar(init_opts=opts.InitOpts('1200px', '280px', bg_color='Skyblue'))
            bar.set_global_opts(toolbox_opts=opts.ToolboxOpts(is_show=True),
                                title_opts=opts.TitleOpts(title='Bar'+self.tab_0QComboBox_2.currentText()+'->指标单位[万元,元,标吨]',
                                                          subtitle='指标单位:万元|标吨'),
                                datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                yaxis_opts=opts.AxisOpts(type_='value',
                                                         # axislabel_opts=opts.LabelOpts(formatter='{value}万元|标吨'),
                                                         axistick_opts=opts.AxisTickOpts(is_show=True),
                                                         name_gap=20,
                                                         axisline_opts=opts.AxisLineOpts(is_on_zero=True,on_zero_axis_index=0)),
                                # tooltip_opts=opts.TooltipOpts(trigger='axis',axis_pointer_type="cross"),

                                )

            bar.add_yaxis(series_name='2021年月预算走势-->'+self.tab_0QComboBox_2.currentText(),
                          y_axis=list(dy.loc[self.tab_0QComboBox_2.currentText()])[1:],is_selected=False)

            bar.add_yaxis(series_name='2021年月实际走势-->'+self.tab_0QComboBox_2.currentText(),
                          y_axis=list(dj.loc[self.tab_0QComboBox_2.currentText()])[1:],gap='0%')

            bar.add_xaxis(list(map(lambda x : x.replace('2021年','').replace('实际','').replace('预算',''),list(dfs_.columns)[1::2])))

            bar.set_series_opts(label_opts=opts.LabelOpts(is_show=True),

                                markline_opts=opts.MarkLineOpts(data=[opts.MarkLineItem(type_="average", name="平均值"),
                                                                      opts.MarkLineItem(type_="min", name="最小值"),
                                                                      opts.MarkLineItem(type_="max", name="最大值")
                                                                      ]))
            # bar.reversal_axis()
            bar.render('2021msj.html')
            path=os.getcwd()+r'\2021msj.html'


            self.tab_0web.load(QUrl.fromLocalFile(path))

    def jd(self):
        # print(jdlist[0].columns)
        df=qlist[0]
        df_=df.drop('项目',axis=0)
        month=self.tab_0QComboBox_0.currentText()
        try:
            if month=='一季度':
                df_['2021年'+month+'预算']=df_['2021年1月预算']+df_['2021年2月预算']+df_['2021年3月预算']

                cb=df_['2021年1月预算'].钨矿产量*df_['2021年1月预算'].钨矿生产成本+\
                                         df_['2021年2月预算'].钨矿产量*df_['2021年2月预算'].钨矿生产成本+\
                                         df_['2021年3月预算'].钨矿产量*df_['2021年3月预算'].钨矿生产成本
                cb_1=cb/df_['2021年'+month+'预算'].钨矿产量
                df_['2021年'+month+'预算'].钨矿生产成本=round(cb_1,0)
                df_['2021年'+month+'预算'].新洲存货资金=df_['2021年3月预算'].新洲存货资金

                df_['2021年'+month+'实际'] = df_['2021年1月实际'] + df_['2021年2月实际'] + df_['2021年3月实际']
                cb = df_['2021年1月实际'].钨矿产量 * df_['2021年1月实际'].钨矿生产成本 + \
                     df_['2021年2月实际'].钨矿产量 * df_['2021年2月实际'].钨矿生产成本 + \
                     df_['2021年3月实际'].钨矿产量 * df_['2021年3月实际'].钨矿生产成本
                cb_1 = cb / df_['2021年'+month+'实际'].钨矿产量
                df_['2021年'+month+'实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年'+month+'实际'].新洲存货资金 = df_['2021年3月实际'].新洲存货资金

                df_['2020年'+month+'实际'] = df_['2020年1月实际'] + df_['2020年2月实际'] + df_['2020年3月实际']
                cb = df_['2020年1月实际'].钨矿产量 * df_['2020年1月实际'].钨矿生产成本 + \
                     df_['2020年2月实际'].钨矿产量 * df_['2020年2月实际'].钨矿生产成本 + \
                     df_['2020年3月实际'].钨矿产量 * df_['2020年3月实际'].钨矿生产成本
                cb_1 = cb / df_['2020年'+month+'实际'].钨矿产量
                df_['2020年' + month + '实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2020年'+month+'实际'].新洲存货资金 = df_['2020年3月实际'].新洲存货资金

                dajd = pd.DataFrame([df_['2021年' + month + '预算'], df_['2021年' + month + '实际'],
                                    df_['2020年' + month + '实际']])
                dajd=dajd.T
                dajd['同比'] = df_['2021年' + month + '实际'] - df_['2020年' + month + '实际']
                dajd['较计划'] = df_['2021年' + month + '实际'] - df_['2021年' + month + '预算']
                dajd=dajd.applymap(lambda x:round(x,2))
                m_df = dajd
                m_df = m_df.applymap(lambda x: round(x, 2))

                m_df.insert(0, '项目', m_df.index)
                ml_df = m_df.T

                ml_df.insert(0, '项目', m_df.columns)
                modf = ml_df.T

                mod = Window_(len(modf), len(modf.columns))
                mod.setHorizontalHeaderLabels(modf.columns)
                for row in range(len(modf)):
                    for col in range(len(modf.columns)):
                        if row == 0 or col == 0:
                            item = QStandardItem(modf.iloc[row, col])
                            item.setForeground(QBrush(QtGui.QColor('dark')))
                            item.setFont(font)
                            mod.setItem(row, col, item)
                        else:
                            item = QStandardItem(str(modf.iloc[row, col]))
                            item.setForeground(QBrush(QtGui.QColor('BLUE')))
                            mod.setItem(row, col, item)

                        self.tab_0TableView.setModel(mod)
                        self.tab_0TableView.horizontalHeader().setVisible(False)
                        self.tab_0TableView.verticalHeader().setVisible(False)
                        self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
                        self.tab_0TableView.resizeColumnsToContents()
                        self.tab_0TableView.setColumnWidth(0, 120)

                def bar_(self):
                    bar = Bar(init_opts=opts.InitOpts('950px', '280px', bg_color='skyblue'))
                    bar.add_xaxis(list(dajd.columns))
                    bar.add_yaxis(
                        series_name='2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText(),
                        y_axis=list(dajd.loc[self.tab_0QComboBox_2.currentText()]))

                    bar.set_global_opts(datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                        title_opts=opts.TitleOpts(
                                            'Bar_' + self.tab_0QComboBox_0.currentText()+self.tab_0QComboBox_2.currentText() + '->指标单位[万元:元:标吨]'),
                                        tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='cross'),
                                        toolbox_opts=opts.ToolboxOpts())

                    return bar

                def garu(self):

                    df = dajd.iloc[:, 1] / dajd.iloc[:, 0]
                    df = df[self.tab_0QComboBox_2.currentText()]
                    # df=df['天工利润']*100
                    df = df.__round__(3)


                    C = Liquid(init_opts=opts.InitOpts(bg_color='skyblue'))
                    C.add(self.tab_0QComboBox_0.currentText() +self.tab_0QComboBox_2.currentText()+ '完成率', [df, 1 - df], is_outline_show=False,
                          shape=SymbolType.DIAMOND)
                    C.set_global_opts(
                        title_opts=opts.TitleOpts(
                            title='2021年'  + self.tab_0QComboBox_0.currentText()+self.tab_0QComboBox_2.currentText() + '完成率'))

                    return C

                tab = Tab()
                tab.add(bar_(self), '2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText())
                tab.add(garu(self),
                        '2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText() + '预算完成率')

                tab.render('jdm_1.html')
                path = os.getcwd() + '/jdm_1.html'
                self.tab_0web.load(QUrl.fromLocalFile(path))

            elif month=='二季度':

                df_['2021年'+month+'预算'] = df_['2021年4月预算'] + df_['2021年5月预算'] + df_['2021年6月预算']
                cb = df_['2021年4月预算'].钨矿产量 * df_['2021年4月预算'].钨矿生产成本 + \
                     df_['2021年5月预算'].钨矿产量 * df_['2021年5月预算'].钨矿生产成本 + \
                     df_['2021年6月预算'].钨矿产量 * df_['2021年6月预算'].钨矿生产成本
                cb_1 = cb / df_['2021年'+month+'预算'].钨矿产量
                df_['2021年'+month+'预算'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年'+month+'预算'].新洲存货资金 = df_['2021年6月预算'].新洲存货资金
                # df_.i
                df_['2021年'+month+'实际'] = df_['2021年4月实际'] + df_['2021年5月实际'] + df_['2021年6月实际']
                cb = df_['2021年4月实际'].钨矿产量 * df_['2021年4月实际'].钨矿生产成本 + \
                     df_['2021年5月实际'].钨矿产量 * df_['2021年5月实际'].钨矿生产成本 + \
                     df_['2021年6月实际'].钨矿产量 * df_['2021年6月实际'].钨矿生产成本
                cb_1 = cb / df_['2021年'+month+'实际'].钨矿产量
                df_['2021年'+month+'实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年'+month+'实际'].新洲存货资金 = df_['2021年6月实际'].新洲存货资金

                df_['2020年' + month + '实际'] = df_['2020年4月实际'] + df_['2020年5月实际'] + df_['2020年6月实际']
                cb = df_['2020年4月实际'].钨矿产量 * df_['2020年4月实际'].钨矿生产成本 + \
                     df_['2020年5月实际'].钨矿产量 * df_['2020年5月实际'].钨矿生产成本 + \
                     df_['2020年6月实际'].钨矿产量 * df_['2020年6月实际'].钨矿生产成本
                cb_1 = cb / df_['2020年' + month + '实际'].钨矿产量
                df_['2020年' + month + '实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2020年' + month + '实际'].新洲存货资金 = df_['2020年6月实际'].新洲存货资金

                dajd = pd.DataFrame([df_['2021年' + month + '预算'], df_['2021年' + month + '实际'],
                                    df_['2020年' + month + '实际']])
                dajd=dajd.T

                dajd['同比'] = df_['2021年' + month + '实际'] - df_['2020年' + month + '实际']
                dajd['较计划'] = df_['2021年' + month + '实际'] - df_['2021年' + month + '预算']
                dajd=dajd.applymap(lambda x:round(x,2))
                # dajd = dajd.applymap(lambda x: "%.2f" % x) #不能这个，有副作用，返回的是字符串，不利于数据分析
                m_df = dajd
                #
                m_df=m_df.applymap(lambda x:round(x,2))
                # print(m_df)
                m_df.insert(0, '项目', m_df.index)
                ml_df = m_df.T

                ml_df.insert(0, '项目', m_df.columns)
                modf=ml_df.T

                mod=Window_(len(modf),len(modf.columns))
                mod.setHorizontalHeaderLabels(modf.columns)
                for row in range(len(modf)):
                    for col in range(len(modf.columns)):
                        if row==0 or col==0:
                            item = QStandardItem(modf.iloc[row, col])
                            item.setForeground(QBrush(QtGui.QColor('dark')))
                            item.setFont(font)
                            mod.setItem(row, col, item)
                        else:
                            item = QStandardItem(str(modf.iloc[row, col]))
                            item.setForeground(QBrush(QtGui.QColor('BLUE')))
                            mod.setItem(row, col, item)

                        self.tab_0TableView.setModel(mod)
                        self.tab_0TableView.horizontalHeader().setVisible(False)
                        self.tab_0TableView.verticalHeader().setVisible(False)
                        self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
                        self.tab_0TableView.resizeColumnsToContents()
                        self.tab_0TableView.setColumnWidth(0, 120)

                def bar_(self):
                    bar=Bar(init_opts=opts.InitOpts('950px','280px',bg_color='skyblue'))
                    bar.add_xaxis(list(dajd.columns))
                    bar.add_yaxis(series_name='2021年'+self.tab_0QComboBox_0.currentText()+self.tab_0QComboBox_2.currentText(),
                                  y_axis=list(dajd.loc[self.tab_0QComboBox_2.currentText()]))

                    bar.set_global_opts(datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                        title_opts=opts.TitleOpts(
                                            'Bar_' + self.tab_0QComboBox_0.currentText() + '->指标单位[万元:元:标吨]'),
                                        tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='cross'),
                                        toolbox_opts=opts.ToolboxOpts())

                    return bar
                def garu(self):

                    df=dajd.iloc[:,1]/dajd.iloc[:,0]
                    df=df[self.tab_0QComboBox_2.currentText()]

                    df=df.__round__(3)

                    C=Liquid(init_opts=opts.InitOpts(bg_color='skyblue'))
                    C.add(self.tab_0QComboBox_0.currentText()+'完成率', [df, 1-df], is_outline_show=False, shape=SymbolType.DIAMOND)
                    C.set_global_opts(
                        title_opts=opts.TitleOpts(title='2021年'+self.tab_0QComboBox_0.currentText()+self.tab_0QComboBox_2.currentText()+'完成率'))

                    return C

                tab=Tab()
                tab.add(bar_(self),'2021年'+self.tab_0QComboBox_0.currentText()+self.tab_0QComboBox_2.currentText())
                tab.add(garu(self),'2021年'+self.tab_0QComboBox_0.currentText()+self.tab_0QComboBox_2.currentText()+'预算完成率')
                 
                tab.render('jdm_1.html')
                path=os.getcwd()+'/jdm_1.html'
                self.tab_0web.load(QUrl.fromLocalFile(path))


            elif month=='三季度':

                df_['2021年' + month + '预算'] = df_['2021年7月预算'] + df_['2021年8月预算'] + df_['2021年9月预算']
                cb = df_['2021年7月预算'].钨矿产量 * df_['2021年7月预算'].钨矿生产成本 + \
                     df_['2021年8月预算'].钨矿产量 * df_['2021年8月预算'].钨矿生产成本 + \
                     df_['2021年9月预算'].钨矿产量 * df_['2021年9月预算'].钨矿生产成本
                cb_1 = cb / df_['2021年'+month+'预算'].钨矿产量
                df_['2021年' + month + '预算'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年' + month + '预算'].新洲存货资金 = df_['2021年9月预算']
                # df_.i
                df_['2021年' + month + '实际'] = df_['2021年7月实际'] + df_['2021年8月实际'] + df_['2021年9月实际']
                cb = df_['2021年7月实际'].钨矿产量 * df_['2021年7月实际'].钨矿生产成本 + \
                     df_['2021年8月实际'].钨矿产量 * df_['2021年8月实际'].钨矿生产成本 + \
                     df_['2021年9月实际'].钨矿产量 * df_['2021年9月实际'].钨矿生产成本
                cb_1 = cb / df_['2021年'+month+'实际'].钨矿产量
                df_['2021年' + month + '实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年' + month + '实际'].新洲存货资金 = df_['2021年6月实际']

                df_['2020年' + month + '实际'] = df_['2020年7月实际'] + df_['2020年8月实际'] + df_['2020年9月实际']
                cb = df_['2020年7月实际'].钨矿产量 * df_['2020年7月实际'].钨矿生产成本 + \
                     df_['2020年8月实际'].钨矿产量 * df_['2020年8月实际'].钨矿生产成本 + \
                     df_['2020年9月实际'].钨矿产量 * df_['2020年9月实际'].钨矿生产成本
                cb_1 = cb / df_['2020年' + month + '实际'].钨矿产量
                df_['2020年' + month + '实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2020年' + month + '实际'].新洲存货资金 = df_['2021年9月实际']

                dajd = pd.DataFrame([df_['2021年' + month + '预算'], df_['2021年' + month + '实际'],
                                    df_['2020年' + month + '实际']])
                dajd = dajd.T
                dajd['同比'] = df_['2021年' + month + '实际'] - df_['2020年' + month + '实际']
                dajd['较计划'] = df_['2021年' + month + '实际'] - df_['2021年' + month + '预算']
                dajd = dajd.T

                dajd['同比'] = df_['2021年' + month + '实际'] - df_['2020年' + month + '实际']
                dajd['较计划'] = df_['2021年' + month + '实际'] - df_['2021年' + month + '预算']
                dajd=dajd.applymap(lambda x:round(x,2))
                m_df = dajd

                m_df = m_df.applymap(lambda x: round(x, 2))

                m_df.insert(0, '项目', m_df.index)
                ml_df = m_df.T

                ml_df.insert(0, '项目', m_df.columns)
                modf = ml_df.T

                mod = Window_(len(modf), len(modf.columns))
                mod.setHorizontalHeaderLabels(modf.columns)
                for row in range(len(modf)):
                    for col in range(len(modf.columns)):
                        if row == 0 or col == 0:
                            item = QStandardItem(modf.iloc[row, col])
                            item.setForeground(QBrush(QtGui.QColor('dark')))
                            item.setFont(font)
                            mod.setItem(row, col, item)
                        else:
                            item = QStandardItem(str(modf.iloc[row, col]))
                            item.setForeground(QBrush(QtGui.QColor('BLUE')))
                            mod.setItem(row, col, item)

                        self.tab_0TableView.setModel(mod)
                        self.tab_0TableView.horizontalHeader().setVisible(False)
                        self.tab_0TableView.verticalHeader().setVisible(False)
                        self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
                        self.tab_0TableView.resizeColumnsToContents()
                        self.tab_0TableView.setColumnWidth(0, 120)

                def bar_(self):
                    bar = Bar(init_opts=opts.InitOpts('950px', '280px', bg_color='skyblue'))
                    bar.add_xaxis(list(dajd.columns))
                    bar.add_yaxis(
                        series_name='2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText(),
                        y_axis=list(dajd.loc[self.tab_0QComboBox_2.currentText()])) # loc的语法

                    bar.set_global_opts(datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                        title_opts=opts.TitleOpts(
                                            'Bar_' + self.tab_0QComboBox_0.currentText() +self.tab_0QComboBox_2.currentText()+ '->指标单位[万元:元:标吨]'),
                                        tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='cross'),
                                        toolbox_opts=opts.ToolboxOpts())

                    return bar

                def garu(self):

                    df = dajd.iloc[:, 1] / dajd.iloc[:, 0]
                    df = df[self.tab_0QComboBox_2.currentText()]
                    df = df.__round__(3)

                    C = Liquid(init_opts=opts.InitOpts(bg_color='skyblue'))
                    C.add(self.tab_0QComboBox_0.currentText() + '完成率', [df, 1 - df], is_outline_show=False,
                          shape=SymbolType.DIAMOND)
                    C.set_global_opts(
                        title_opts=opts.TitleOpts(
                            title='2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText() + '完成率'))

                    return C

                tab = Tab()
                tab.add(bar_(self), '2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText())
                tab.add(garu(self),
                        '2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText() + '预算完成率')

                tab.render('jdm_1.html')
                path = os.getcwd() + '/jdm_1.html'
                self.tab_0web.load(QUrl.fromLocalFile(path))

            elif month=='四季度':
                df_['2021年' + month + '预算'] = df_['2021年10月预算'] + df_['2021年11月预算'] + df_['2021年12月预算']
                cb = df_['2021年10月预算'].钨矿产量 * df_['2021年10月预算'].钨矿生产成本 + \
                     df_['2021年11月预算'].钨矿产量 * df_['2021年11月预算'].钨矿生产成本 + \
                     df_['2021年12月预算'].钨矿产量 * df_['2021年12月预算'].钨矿生产成本
                cb_1 = cb / df_['2021年' + month + '预算'].钨矿产量
                df_['2021年' + month + '预算'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年' + month + '预算'].新洲存货资金 = df_['2021年9月预算']
                # df_.i
                df_['2021年' + month + '实际'] = df_['2021年10月实际'] + df_['2021年11月实际'] + df_['2021年12月实际']
                cb = df_['2021年10月实际'].钨矿产量 * df_['2021年10月实际'].钨矿生产成本 + \
                     df_['2021年11月实际'].钨矿产量 * df_['2021年11月实际'].钨矿生产成本 + \
                     df_['2021年12月实际'].钨矿产量 * df_['2021年12月实际'].钨矿生产成本
                cb_1 = cb / df_['2021年' + month + '实际'].钨矿产量
                df_['2021年' + month + '实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年' + month + '实际'].新洲存货资金 = df_['2021年6月实际']

                df_['2020年' + month + '实际'] = df_['2020年7月实际'] + df_['2020年8月实际'] + df_['2020年9月实际']
                cb = df_['2020年10月实际'].钨矿产量 * df_['2020年10月实际'].钨矿生产成本 + \
                     df_['2020年11月实际'].钨矿产量 * df_['2020年11月实际'].钨矿生产成本 + \
                     df_['2020年12月实际'].钨矿产量 * df_['2020年12月实际'].钨矿生产成本
                cb_1 = cb / df_['2020年' + month + '实际'].钨矿产量
                df_['2021年' + month + '实际'].钨矿生产成本 = round(cb_1, 0)
                df_['2021年' + month + '实际'].新洲存货资金 = df_['2021年9月实际']

                dajd=pd.DataFrame([df_['2021年' + month + '预算'],df_['2021年' + month + '实际'],df_['2020年' + month + '实际']])
                dajd = dajd.T
                dajd['同比'] = df_['2021年' + month + '实际'] - df_['2020年' + month + '实际']
                dajd['较计划'] = df_['2021年' + month + '实际'] - df_['2021年' + month + '预算']
                dajd=dajd.applymap(lambda x:round(x,2))
                m_df = dajd # 数据分析用,m_df对象为mvc用

                m_df = m_df.applymap(lambda x: round(x, 2))
                # print(m_df)
                m_df.insert(0, '项目', m_df.index)
                ml_df = m_df.T
                # print(m_df)
                ml_df.insert(0, '项目', m_df.columns)
                modf = ml_df.T

                mod = Window_(len(modf), len(modf.columns))
                mod.setHorizontalHeaderLabels(modf.columns)
                for row in range(len(modf)):
                    for col in range(len(modf.columns)):
                        if row == 0 or col == 0:
                            item = QStandardItem(modf.iloc[row, col])
                            item.setForeground(QBrush(QtGui.QColor('dark')))
                            item.setFont(font)
                            mod.setItem(row, col, item)
                        else:
                            item = QStandardItem(str(modf.iloc[row, col]))
                            item.setForeground(QBrush(QtGui.QColor('BLUE')))
                            mod.setItem(row, col, item)

                        self.tab_0TableView.setModel(mod)
                        self.tab_0TableView.horizontalHeader().setVisible(False)
                        self.tab_0TableView.verticalHeader().setVisible(False)
                        self.tab_0TableView.setEditTriggers(QAbstractItemView.NoEditTriggers)
                        self.tab_0TableView.resizeColumnsToContents()
                        self.tab_0TableView.setColumnWidth(0, 120)

                def bar_(self):
                    bar = Bar(init_opts=opts.InitOpts('950px', '280px', bg_color='skyblue'))
                    bar.add_xaxis(list(dajd.columns))
                    bar.add_yaxis(
                        series_name='2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText(),
                        y_axis=list(dajd.loc[self.tab_0QComboBox_2.currentText()]))

                    bar.set_global_opts(datazoom_opts=opts.DataZoomOpts(type_='inside'),
                                        title_opts=opts.TitleOpts(
                                            'Bar_' + self.tab_0QComboBox_0.currentText() +self.tab_0QComboBox_2.currentText()+'->指标单位[万元:元:标吨]'),
                                        tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='cross'),
                                        toolbox_opts=opts.ToolboxOpts())

                    return bar

                def garu(self):

                    df = dajd.iloc[:, 1] / dajd.iloc[:, 0]
                    df = df[self.tab_0QComboBox_2.currentText()]
                    # df=df['天工利润']*100
                    df = df.__round__(3)

                    C = Liquid(init_opts=opts.InitOpts(bg_color='skyblue'))
                    C.add(self.tab_0QComboBox_0.currentText() + '完成率', [df, 1 - df], is_outline_show=False,
                          shape=SymbolType.DIAMOND)
                    C.set_global_opts(
                        title_opts=opts.TitleOpts(
                            title='2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText() + '完成率'))

                    return C

                tab = Tab()
                tab.add(bar_(self), '2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText())
                tab.add(garu(self),
                        '2021年' + self.tab_0QComboBox_0.currentText() + self.tab_0QComboBox_2.currentText() + '预算完成率')

                tab.render('jdm_1.html')
                path = os.getcwd() + '/jdm_1.html'
                self.tab_0web.load(QUrl.fromLocalFile(path))
        except:
            QMessageBox.information(self.tab_0,'提示','季度时间不对,请核实季度所占的月份!',QMessageBox.Yes,QMessageBox.No)

    # 以下是事件函数
    def selected_tb_text(self, a):
        try:
            indexes = a.selectedIndexes()  # 获取表格对象中被选中的数据索引列表
            indexes_dict = {}
            for index in indexes:  # 遍历每个单元格
                row, column = index.row(), index.column()  # 获取单元格的行号，列号
                if row in indexes_dict.keys():
                    indexes_dict[row].append(column)
                else:
                    indexes_dict[row] = [column]
            # 将数据表数据用制表符(\t)和换行符(\n)连接，使其可以复制到excel文件中(难点，卖点)
            text = ''
            for row, columns in indexes_dict.items():
                row_data = ''
                for column in columns:
                    data = a.model().item(row, column).text()
                    if row_data:
                        row_data = row_data + '\t' + data
                    else:
                        row_data = data

                if text:
                    text = text + '\n' + row_data
                else:
                    text = row_data
            print(text)
            return text
        except BaseException as e:
            print(e)
            return ''

    def keyReleaseEvent(self, Event):  # 重写键盘监听事件
        # 监听 CTRL+C 组合键，实现复制数据到粘贴板
        if Event.key() == Qt.Key_C and QApplication.keyboardModifiers() == Qt.ControlModifier:
            Event.key()

            text = self.selected_tb_text(self.tab_0TableView)  # 获取当前表格选中的数据
            if text:
                try:
                    clipboard = QApplication.clipboard()
                    data = clipboard.mimeData()
                    print(data.formats())
                    clipboard.setText(text)  # 复制到粘贴板

                except BaseException as e:
                    print(e)

    def keyPressEvent(self, QKeyEvent):  # 重写键盘监听事件
        # 监听 CTRL+ALT 组合键，实现复制数据到粘贴板
        if QKeyEvent.key() == Qt.Key_Alt and QApplication.keyboardModifiers() == Qt.ControlModifier:
            text = self.selected_tb_text(self.tab_1TableView)  # 获取当前表格选中的数据
            if text:
                try:
                    clipboard = QApplication.clipboard()
                    data = clipboard.mimeData()
                    print(data.formats())
                    clipboard.setText(text)  # 复制到粘贴板

                except BaseException as e:
                    print(e)


if __name__ == '__main__':
    import sys
    QApplication.setAttribute(Qt.AA_UseSoftwareOpenGL)
    APP = QApplication(sys.argv)
    Ui = Ui_MainWindow()
    Ui.setUi_0()
    Ui.setUi_1()
    Ui.setUi_2()
    Ui.uidata_0()
    Ui.setdata_0()
    Ui.show()
    sys.exit(APP.exec_())
