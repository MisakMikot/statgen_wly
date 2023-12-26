import matplotlib.pyplot as plt
from PySide6 import QtWidgets, QtCore, QtGui
from PySide6.QtWidgets import QApplication, QMainWindow, QTreeWidgetItem, QLabel, QFileDialog
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QPixmap, QDoubleValidator, QIntValidator, QMouseEvent
# from Ui_statgen import Ui_MainWindow
import sys
import os
import xlrd


class MyQLabel(QLabel):
    def __int__(self):
        super().__init__()
    clicked = Signal()

    def mousePressEvent(self, QMouseEvent):
        if QMouseEvent.buttons() == QtCore.Qt.LeftButton:
            self.clicked.emit()

    def enterEvent(self, QMouseEvent):
        self.setToolTip("将图表保存到本地")


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(749, 552)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout.setObjectName("verticalLayout")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.horizontalLayout_7 = QtWidgets.QHBoxLayout(self.tab)
        self.horizontalLayout_7.setObjectName("horizontalLayout_7")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.zx_list = QtWidgets.QTreeWidget(self.tab)
        self.zx_list.setObjectName("zx_list")
        self.verticalLayout_2.addWidget(self.zx_list)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        self.zx_del = QtWidgets.QPushButton(self.tab)
        self.zx_del.setObjectName("zx_del")
        self.horizontalLayout_3.addWidget(self.zx_del)
        self.zx_clr = QtWidgets.QPushButton(self.tab)
        self.zx_clr.setObjectName("zx_clr")
        self.horizontalLayout_3.addWidget(self.zx_clr)
        self.verticalLayout_2.addLayout(self.horizontalLayout_3)
        self.horizontalLayout_7.addLayout(self.verticalLayout_2)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.horizontalLayout_5 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_5.setObjectName("horizontalLayout_5")
        self.label_7 = QtWidgets.QLabel(self.tab)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout_5.addWidget(self.label_7)
        self.zx_x = QtWidgets.QLineEdit(self.tab)
        self.zx_x.setObjectName("zx_x")
        self.horizontalLayout_5.addWidget(self.zx_x)
        self.label_4 = QtWidgets.QLabel(self.tab)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout_5.addWidget(self.label_4)
        self.zx_y = QtWidgets.QLineEdit(self.tab)
        self.zx_y.setObjectName("zx_y")
        self.horizontalLayout_5.addWidget(self.zx_y)
        self.verticalLayout_3.addLayout(self.horizontalLayout_5)
        self.horizontalLayout_6 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_6.setObjectName("horizontalLayout_6")
        self.label_5 = QtWidgets.QLabel(self.tab)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout_6.addWidget(self.label_5)
        self.zx_name = QtWidgets.QLineEdit(self.tab)
        self.zx_name.setObjectName("zx_name")
        self.horizontalLayout_6.addWidget(self.zx_name)
        self.verticalLayout_3.addLayout(self.horizontalLayout_6)
        self.horizontalLayout_4 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_4.setObjectName("horizontalLayout_4")
        self.zx_item = QtWidgets.QLineEdit(self.tab)
        self.zx_item.setObjectName("zx_item")
        self.horizontalLayout_4.addWidget(self.zx_item)
        self.zx_value = QtWidgets.QLineEdit(self.tab)
        self.zx_value.setObjectName("zx_value")
        self.horizontalLayout_4.addWidget(self.zx_value)
        self.zx_add = QtWidgets.QPushButton(self.tab)
        self.zx_add.setObjectName("zx_add")
        self.horizontalLayout_4.addWidget(self.zx_add)
        self.zx_load = QtWidgets.QPushButton(self.tab)
        self.zx_load.setObjectName("zx_load")
        self.horizontalLayout_4.addWidget(self.zx_load)
        self.verticalLayout_3.addLayout(self.horizontalLayout_4)
        self.zx_gen = QtWidgets.QPushButton(self.tab)
        self.zx_gen.setObjectName("zx_gen")
        self.verticalLayout_3.addWidget(self.zx_gen)
        self.horizontalLayout_7.addLayout(self.verticalLayout_3)
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.horizontalLayout_12 = QtWidgets.QHBoxLayout(self.tab_2)
        self.horizontalLayout_12.setObjectName("horizontalLayout_12")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.zz_list = QtWidgets.QTreeWidget(self.tab_2)
        self.zz_list.setObjectName("zz_list")
        self.verticalLayout_4.addWidget(self.zz_list)
        self.horizontalLayout_8 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_8.setObjectName("horizontalLayout_8")
        self.zz_del = QtWidgets.QPushButton(self.tab_2)
        self.zz_del.setObjectName("zz_del")
        self.horizontalLayout_8.addWidget(self.zz_del)
        self.zz_clr = QtWidgets.QPushButton(self.tab_2)
        self.zz_clr.setObjectName("zz_clr")
        self.horizontalLayout_8.addWidget(self.zz_clr)
        self.verticalLayout_4.addLayout(self.horizontalLayout_8)
        self.horizontalLayout_12.addLayout(self.verticalLayout_4)
        self.verticalLayout_5 = QtWidgets.QVBoxLayout()
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.horizontalLayout_10 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_10.setObjectName("horizontalLayout_10")
        self.label = QtWidgets.QLabel(self.tab_2)
        self.label.setObjectName("label")
        self.horizontalLayout_10.addWidget(self.label)
        self.zz_xname = QtWidgets.QLineEdit(self.tab_2)
        self.zz_xname.setObjectName("zz_xname")
        self.horizontalLayout_10.addWidget(self.zz_xname)
        self.label_3 = QtWidgets.QLabel(self.tab_2)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_10.addWidget(self.label_3)
        self.zz_yname = QtWidgets.QLineEdit(self.tab_2)
        self.zz_yname.setObjectName("zz_yname")
        self.horizontalLayout_10.addWidget(self.zz_yname)
        self.verticalLayout_5.addLayout(self.horizontalLayout_10)
        self.horizontalLayout_11 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_11.setObjectName("horizontalLayout_11")
        self.label_6 = QtWidgets.QLabel(self.tab_2)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout_11.addWidget(self.label_6)
        self.zz_itemname = QtWidgets.QLineEdit(self.tab_2)
        self.zz_itemname.setObjectName("zz_itemname")
        self.horizontalLayout_11.addWidget(self.zz_itemname)
        self.verticalLayout_5.addLayout(self.horizontalLayout_11)
        self.horizontalLayout_9 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_9.setObjectName("horizontalLayout_9")
        self.zz_item = QtWidgets.QLineEdit(self.tab_2)
        self.zz_item.setObjectName("zz_item")
        self.horizontalLayout_9.addWidget(self.zz_item)
        self.zz_value = QtWidgets.QLineEdit(self.tab_2)
        self.zz_value.setObjectName("zz_value")
        self.horizontalLayout_9.addWidget(self.zz_value)
        self.zz_add = QtWidgets.QPushButton(self.tab_2)
        self.zz_add.setObjectName("zz_add")
        self.horizontalLayout_9.addWidget(self.zz_add)
        self.zz_load = QtWidgets.QPushButton(self.tab_2)
        self.zz_load.setObjectName("zz_load")
        self.horizontalLayout_9.addWidget(self.zz_load)
        self.verticalLayout_5.addLayout(self.horizontalLayout_9)
        self.zz_gen = QtWidgets.QPushButton(self.tab_2)
        self.zz_gen.setObjectName("zz_gen")
        self.verticalLayout_5.addWidget(self.zz_gen)
        self.horizontalLayout_12.addLayout(self.verticalLayout_5)
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.horizontalLayout_15 = QtWidgets.QHBoxLayout(self.tab_3)
        self.horizontalLayout_15.setObjectName("horizontalLayout_15")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout()
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.bt_list = QtWidgets.QTreeWidget(self.tab_3)
        self.bt_list.setAnimated(False)
        self.bt_list.setObjectName("bt_list")
        self.verticalLayout_6.addWidget(self.bt_list)
        self.horizontalLayout_13 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_13.setObjectName("horizontalLayout_13")
        self.bt_del = QtWidgets.QPushButton(self.tab_3)
        self.bt_del.setObjectName("bt_del")
        self.horizontalLayout_13.addWidget(self.bt_del)
        self.bt_clr = QtWidgets.QPushButton(self.tab_3)
        self.bt_clr.setObjectName("bt_clr")
        self.horizontalLayout_13.addWidget(self.bt_clr)
        self.verticalLayout_6.addLayout(self.horizontalLayout_13)
        self.horizontalLayout_15.addLayout(self.verticalLayout_6)
        self.verticalLayout_7 = QtWidgets.QVBoxLayout()
        self.verticalLayout_7.setObjectName("verticalLayout_7")
        self.horizontalLayout_14 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_14.setObjectName("horizontalLayout_14")
        self.bt_item = QtWidgets.QLineEdit(self.tab_3)
        self.bt_item.setObjectName("bt_item")
        self.horizontalLayout_14.addWidget(self.bt_item)
        self.bt_value = QtWidgets.QLineEdit(self.tab_3)
        self.bt_value.setObjectName("bt_value")
        self.horizontalLayout_14.addWidget(self.bt_value)
        self.bt_out = QtWidgets.QLineEdit(self.tab_3)
        self.bt_out.setObjectName("bt_out")
        self.horizontalLayout_14.addWidget(self.bt_out)
        self.bt_add = QtWidgets.QPushButton(self.tab_3)
        self.bt_add.setObjectName("bt_add")
        self.horizontalLayout_14.addWidget(self.bt_add)
        self.bt_load = QtWidgets.QPushButton(self.tab_3)
        self.bt_load.setObjectName("bt_load")
        self.horizontalLayout_14.addWidget(self.bt_load)
        self.verticalLayout_7.addLayout(self.horizontalLayout_14)
        self.bt_gen = QtWidgets.QPushButton(self.tab_3)
        self.bt_gen.setObjectName("bt_gen")
        self.verticalLayout_7.addWidget(self.bt_gen)
        self.horizontalLayout_15.addLayout(self.verticalLayout_7)
        self.horizontalLayout_15.setStretch(0, 2)
        self.horizontalLayout_15.setStretch(1, 3)
        self.tabWidget.addTab(self.tab_3, "")
        self.verticalLayout.addWidget(self.tabWidget)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.img = MyQLabel(self.centralwidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.img.sizePolicy().hasHeightForWidth())
        self.img.setSizePolicy(sizePolicy)
        self.img.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.img.setText("")
        self.img.setScaledContents(True)
        self.img.setObjectName("img")
        self.horizontalLayout.addWidget(self.img)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem1)
        self.horizontalLayout.setStretch(0, 1)
        self.horizontalLayout.setStretch(1, 3)
        self.horizontalLayout.setStretch(2, 1)
        self.verticalLayout.addLayout(self.horizontalLayout)
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout_2.addWidget(self.label_2)
        self.head = QtWidgets.QLineEdit(self.centralwidget)
        self.head.setObjectName("head")
        self.horizontalLayout_2.addWidget(self.head)
        self.verticalLayout.addLayout(self.horizontalLayout_2)
        self.verticalLayout.setStretch(0, 1)
        self.verticalLayout.setStretch(1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 749, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "统计表生成器-汪刘洋"))
        self.zx_list.headerItem().setText(0, _translate("MainWindow", "项"))
        self.zx_list.headerItem().setText(1, _translate("MainWindow", "值"))
        self.zx_del.setText(_translate("MainWindow", "删除项"))
        self.zx_clr.setText(_translate("MainWindow", "清空"))
        self.label_7.setText(_translate("MainWindow", "X轴名称："))
        self.zx_x.setText(_translate("MainWindow", "X"))
        self.label_4.setText(_translate("MainWindow", "Y轴名称："))
        self.zx_y.setText(_translate("MainWindow", "Y"))
        self.label_5.setText(_translate("MainWindow", "项名称："))
        self.zx_name.setText(_translate("MainWindow", "item"))
        self.zx_item.setPlaceholderText(_translate("MainWindow", "项"))
        self.zx_value.setPlaceholderText(_translate("MainWindow", "值"))
        self.zx_add.setText(_translate("MainWindow", "添加项"))
        self.zx_load.setText(_translate("MainWindow", "从表格导入"))
        self.zx_gen.setText(_translate("MainWindow", "生成"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "折线图"))
        self.zz_list.headerItem().setText(0, _translate("MainWindow", "项"))
        self.zz_list.headerItem().setText(1, _translate("MainWindow", "值"))
        self.zz_del.setText(_translate("MainWindow", "删除项"))
        self.zz_clr.setText(_translate("MainWindow", "清空"))
        self.label.setText(_translate("MainWindow", "X轴名称："))
        self.zz_xname.setText(_translate("MainWindow", "X"))
        self.label_3.setText(_translate("MainWindow", "Y轴名称："))
        self.zz_yname.setText(_translate("MainWindow", "Y"))
        self.label_6.setText(_translate("MainWindow", "项名称："))
        self.zz_itemname.setText(_translate("MainWindow", "item"))
        self.zz_item.setPlaceholderText(_translate("MainWindow", "项"))
        self.zz_value.setPlaceholderText(_translate("MainWindow", "值"))
        self.zz_add.setText(_translate("MainWindow", "添加项"))
        self.zz_load.setText(_translate("MainWindow", "从表格导入"))
        self.zz_gen.setText(_translate("MainWindow", "生成"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "柱状图"))
        self.bt_list.headerItem().setText(0, _translate("MainWindow", "项"))
        self.bt_list.headerItem().setText(1, _translate("MainWindow", "值"))
        self.bt_list.headerItem().setText(2, _translate("MainWindow", "离心"))
        self.bt_del.setText(_translate("MainWindow", "删除项"))
        self.bt_clr.setText(_translate("MainWindow", "清空"))
        self.bt_item.setPlaceholderText(_translate("MainWindow", "项"))
        self.bt_value.setPlaceholderText(_translate("MainWindow", "值"))
        self.bt_out.setPlaceholderText(_translate("MainWindow", "离心（1-100）"))
        self.bt_add.setText(_translate("MainWindow", "添加项"))
        self.bt_load.setText(_translate("MainWindow", "从表格导入"))
        self.bt_gen.setText(_translate("MainWindow", "生成"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "饼图"))
        self.label_2.setText(_translate("MainWindow", "表头："))
        self.head.setText(_translate("MainWindow", "统计表"))


class myWindow(QMainWindow):
    def closeEvent(self, event):
        sys.exit()


def make_autopct(values):
    def my_autopct(pct):
        total = sum(values)
        val = int(round(pct*total/100.0))
        # 同时显示数值和占比的饼图
        return '{p:.2f}%  ({v:d})'.format(p=pct, v=val)
    return my_autopct


def bt_gen():
    plt.clf()
    title = ui.head.text()
    items = []
    values = []
    out = []
    for i in range(0, ui.bt_list.topLevelItemCount()):
        item = ui.bt_list.topLevelItem(i)
        items.append(item.text(0))
        values.append(float(item.text(1)))
        out.append(int(item.text(2))/100)
    plt.pie(values, out, items, autopct=make_autopct(
        values), pctdistance=0.5, labeldistance=1.2)
    plt.title(title)
    plt.savefig('temp.png')
    pix = QPixmap('temp.png')
    os.remove('temp.png')
    ui.img.setPixmap(pix)


def bt_clr():
    ui.bt_list.clear()


def bt_add():
    item = ui.bt_item.text()
    value = ui.bt_value.text()
    out = ui.bt_out.text()
    if out == '':
        out = '0'
    if item == '' or value == '':
        return
    a = QTreeWidgetItem(ui.bt_list.invisibleRootItem())
    a.setText(0, item)
    a.setText(1, value)
    a.setText(2, out)
    ui.bt_item.setText('')
    ui.bt_value.setText('')
    ui.bt_out.setText('')


def bt_del():
    ui.bt_list.invisibleRootItem().removeChild(ui.bt_list.currentItem())


def zz_gen():
    plt.clf()
    title = ui.head.text()
    xname = ui.zz_xname.text()
    yname = ui.zz_yname.text()
    itemname = ui.zz_itemname.text()
    items = []
    values = []
    for i in range(0, ui.zz_list.topLevelItemCount()):
        item = ui.zz_list.topLevelItem(i)
        items.append(item.text(0))
        values.append(float(item.text(1)))
    plt.bar(range(len(values)), values, tick_label=items, label=itemname)
    plt.title(title)
    plt.xlabel(xname)
    plt.ylabel(yname)
    plt.legend(loc='best')
    c = 0
    for a, b in zip(items, values):
        plt.text(c, b, str(b), ha='center', va='bottom')
        c += 1
    plt.savefig('temp.png')
    pix = QPixmap('temp.png')
    os.remove('temp.png')
    ui.img.setPixmap(pix)


def zz_del():
    ui.zz_list.invisibleRootItem().removeChild(ui.zz_list.currentItem())


def zz_clr():
    ui.zz_list.clear()


def zz_add():
    item = ui.zz_item.text()
    value = ui.zz_value.text()
    if item == '' or value == '':
        return
    a = QTreeWidgetItem(ui.zz_list.invisibleRootItem())
    a.setText(0, item)
    a.setText(1, value)
    ui.zz_item.setText('')
    ui.zz_value.setText('')


def zx_gen():
    plt.clf()
    title = ui.head.text()
    xname = ui.zx_x.text()
    yname = ui.zx_y.text()
    itemname = ui.zx_name.text()
    items = []
    values = []
    for i in range(0, ui.zx_list.topLevelItemCount()):
        item = ui.zx_list.topLevelItem(i)
        items.append(item.text(0))
        values.append(float(item.text(1)))
    plt.plot(items, values, c='red', label=itemname)
    plt.scatter(items,values, c='red')
    plt.legend(loc='best')
    plt.grid(True, linestyle='--', alpha=0.5)
    plt.title(title)
    plt.xlabel(xname)
    plt.ylabel(yname)
    for x1, y1 in zip(items, values):
        plt.text(x1, y1, str(y1), ha='center', va='bottom', fontsize=10)
    plt.savefig('temp.png')
    pix = QPixmap('temp.png')
    os.remove('temp.png')
    ui.img.setPixmap(pix)


def zx_clr():
    ui.zx_list.clear()


def zx_del():
    ui.zx_list.invisibleRootItem().removeChild(ui.zx_list.currentItem())


def zx_add():
    item = ui.zx_item.text()
    value = ui.zx_value.text()
    if item == '' or value == '':
        return
    a = QTreeWidgetItem(ui.zx_list.invisibleRootItem())
    a.setText(0, item)
    a.setText(1, value)
    ui.zx_item.setText('')
    ui.zx_value.setText('')


def zx_load():
    path = QFileDialog.getOpenFileName(mw, '打开表格', filter='XLS表格 (*.xls)')[0]
    if path == '':
        return
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    for row in range(sh.nrows):
        try:
            a = QTreeWidgetItem(ui.zx_list.invisibleRootItem())
            a.setText(0, sh.cell_value(row, 0))
            a.setText(1, str(float(sh.cell_value(row, 1))))
        except:
            ui.zx_list.invisibleRootItem().removeChild(a)


def zz_load():
    path = QFileDialog.getOpenFileName(mw, '打开表格', filter='XLS表格 (*.xls)')[0]
    if path == '':
        return
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    for row in range(sh.nrows):
        try:
            a = QTreeWidgetItem(ui.zz_list.invisibleRootItem())
            a.setText(0, sh.cell_value(row, 0))
            a.setText(1, str(float(sh.cell_value(row, 1))))
        except:
            ui.zz_list.invisibleRootItem().removeChild(a)


def bt_load():
    path = QFileDialog.getOpenFileName(mw, '打开表格', filter='XLS表格 (*.xls)')[0]
    if path == '':
        return
    sh = xlrd.open_workbook(path).sheet_by_index(0)
    for row in range(sh.nrows):
        try:
            a = QTreeWidgetItem(ui.bt_list.invisibleRootItem())
            a.setText(0, sh.cell_value(row, 0))
            a.setText(1, str(float(sh.cell_value(row, 1))))
            a.setText(2, '0')
        except:
            ui.zz_list.invisibleRootItem().removeChild(a)


def save():
    path = QFileDialog.getSaveFileName(mw, '保存统计图', filter='PNG图片 (*.png)')[0]
    if path == '':
        return
    plt.savefig(path, dpi=300)


def slotConn():
    ui.zx_gen.clicked.connect(zx_gen)
    ui.zx_add.clicked.connect(zx_add)
    ui.zx_clr.clicked.connect(zx_clr)
    ui.zx_del.clicked.connect(zx_del)
    ui.zx_load.clicked.connect(zx_load)
    ui.zz_add.clicked.connect(zz_add)
    ui.zz_del.clicked.connect(zz_del)
    ui.zz_clr.clicked.connect(zz_clr)
    ui.zz_gen.clicked.connect(zz_gen)
    ui.zz_load.clicked.connect(zz_load)
    ui.bt_add.clicked.connect(bt_add)
    ui.bt_del.clicked.connect(bt_del)
    ui.bt_clr.clicked.connect(bt_clr)
    ui.bt_gen.clicked.connect(bt_gen)
    ui.bt_load.clicked.connect(bt_load)
    ui.img.clicked.connect(save)


def main():
    plt.rcParams['font.sans-serif'] = ['Microsoft YaHei']  # 用来正常显示中文标签
    plt.rcParams['axes.unicode_minus'] = False  # 用来正常显示负号
    plt.figure(dpi=100)
    plt.clf()
    plt.plot()
    plt.savefig('temp.png')
    pix = QPixmap('temp.png')
    os.remove('temp.png')
    ui.img.setPixmap(pix)
    dval = QDoubleValidator()
    ui.zx_value.setValidator(dval)
    ui.zz_value.setValidator(dval)
    ui.bt_value.setValidator(dval)
    ui.bt_out.setValidator(QIntValidator(0, 100))
    # root = QTreeWidgetItem(ui.zz_list)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    mw = myWindow()
    ui = Ui_MainWindow()
    ui.setupUi(mw)
    mw.show()
    main()
    slotConn()
    sys.exit(app.exec())
