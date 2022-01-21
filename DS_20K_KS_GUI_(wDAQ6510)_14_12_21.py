from PyQt5 import QtCore, QtGui, QtWidgets
from datetime import datetime
from PyQt5.QtCore import QMutex, QObject, QRect, QThread, pyqtSignal, pyqtSlot
from PyQt5.QtGui import QPainter, QColor, QBrush
from scipy.interpolate import interp1d
import pandas as pd
import seaborn as sns
import holoviews as hv
import hvplot.pandas
import numpy as np
import pyvisa
import time
import math
import ctypes
import glob
import os
import threading
#
import ukri_logo_small # UKRI Logo Resource File

# timeKeeper Class - Background thread polling time and date and updating signal to main thread and worker thread
class timeKeeper(QObject):
    finished_time = pyqtSignal()
    timeDate = pyqtSignal(str)          #signal for main GUI
    timeDate_0 = pyqtSignal(str)        #signal for worker thread

    def __init__(self) -> None:
        super().__init__(parent=None)
    
    def run(self):
        now = datetime.now()
        print("Time Keeper Running...")
        self.pastTime = now.strftime("%H:%M:%S")
        while True:
            now = datetime.now()
            self.time   = now.strftime("%H:%M:%S")
            #self.date   = now.strftime("%d/%m/%Y")
            self.date   = now.strftime("%Y/%m/%d")
            if self.time != self.pastTime:
                self.pastTime = now.strftime("%H:%M:%S")
                self.concat = (self.time+","+self.date)
                self.timeDate.emit(self.concat)
                self.timeDate_0.emit(self.concat)
            time.sleep(0.1)

'''End Class'''

# Main GUI Window Class:
class Ui_MainWindow(object):
    # Setup Function:
    def setupUi(self, MainWindow):
        myappid = u'DS_20K_QCC_.V1.0'
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        #
        self.startTimeKeeper()
        self.workerStartedQuery = False
        now = datetime.now()
        #
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1281, 971)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        MainWindow.setWindowIcon(QtGui.QIcon("ukri_square256.png"))
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(10, 90, 1261, 841))
        self.tabWidget.setObjectName("tabWidget")
        self.tab = QtWidgets.QWidget()
        self.tab.setObjectName("tab")
        self.label_7 = QtWidgets.QLabel(self.tab)
        self.label_7.setGeometry(QtCore.QRect(130, 470, 241, 21))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_7.setFont(font)
        self.label_7.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_7.setObjectName("label_7")
        self.pushButton = QtWidgets.QPushButton(self.tab)
        self.pushButton.setGeometry(QtCore.QRect(70, 380, 431, 41))
        self.pushButton.setObjectName("pushButton")
        #
        #self.plainTextEdit = QtWidgets.QPlainTextEdit(self.tab)
        #self.plainTextEdit.setGeometry(QtCore.QRect(70, 530, 1121, 261))
        #self.plainTextEdit.setObjectName("plainTextEdit")
        #
        self.plainTextEdit = QtWidgets.QTextBrowser(self.tab)
        self.plainTextEdit.setGeometry(QRect(70, 530, 1120, 260))
        self.plainTextEdit.setObjectName("plainTextEdit")
        #
        self.tabWidget_2 = QtWidgets.QTabWidget(self.tab)
        self.tabWidget_2.setGeometry(QtCore.QRect(580, 20, 611, 421))
        self.tabWidget_2.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.tabWidget_2.setIconSize(QtCore.QSize(16, 16))
        self.tabWidget_2.setElideMode(QtCore.Qt.ElideNone)
        self.tabWidget_2.setUsesScrollButtons(True)
        self.tabWidget_2.setTabsClosable(False)
        self.tabWidget_2.setMovable(False)
        self.tabWidget_2.setTabBarAutoHide(True)
        self.tabWidget_2.setObjectName("tabWidget_2")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.tab_3 = TileWidget()
        self.widget = QtWidgets.QWidget(self.tab_3)
        self.widget.setGeometry(QtCore.QRect(0, 0, 601, 401))
        self.widget.setObjectName("widget")
        self.tabWidget_2.addTab(self.tab_3, "")
        self.tab_4 = QtWidgets.QWidget()
        self.tab_4.setObjectName("tab_4")
        self.widget_2 = QtWidgets.QWidget(self.tab_4)
        self.widget_2.setGeometry(QtCore.QRect(0, 0, 601, 391))
        self.widget_2.setObjectName("widget_2")
        self.tableWidget = QtWidgets.QTableWidget(self.widget_2)
        self.tableWidget.setGeometry(QtCore.QRect(70, 10, 460, 175))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.tableWidget.sizePolicy().hasHeightForWidth())
        self.tableWidget.setSizePolicy(sizePolicy)
        self.tableWidget.setRowCount(5)
        self.tableWidget.setColumnCount(4)
        self.tableWidget.setObjectName("tableWidget")
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(1, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(3, 2, item)
        self.widget1 = QtWidgets.QWidget(self.widget_2)
        self.widget1.setGeometry(QtCore.QRect(150, 240, 281, 91))
        self.widget1.setObjectName("widget1")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.widget1)
        self.gridLayout_3.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.label_29 = QtWidgets.QLabel(self.widget1)
        font = QtGui.QFont()
        font.setPointSize(10)
        #
        self.label_31 = QtWidgets.QLabel(self.widget1)
        self.label_31.setFont(font)
        self.label_31.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_31.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_31.setObjectName("label_31")
        self.gridLayout_3.addWidget(self.label_31, 0, 0, 1, 1)
        #
        self.label_32 = QtWidgets.QLabel(self.widget1)
        self.label_32.setFont(font)
        self.label_32.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_32.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_32.setObjectName("label_31")
        self.gridLayout_3.addWidget(self.label_32, 0, 2, 1, 1)
        #
        self.label_29.setFont(font)
        self.label_29.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_29.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_29.setObjectName("label_29")
        self.gridLayout_3.addWidget(self.label_29, 2, 0, 1, 1)
        self.label_28 = QtWidgets.QLabel(self.widget1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_28.setFont(font)
        self.label_28.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_28.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_28.setObjectName("label_28")
        self.gridLayout_3.addWidget(self.label_28, 1, 2, 1, 1)
        self.label_30 = QtWidgets.QLabel(self.widget1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_30.setFont(font)
        self.label_30.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_30.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_30.setObjectName("label_30")
        self.gridLayout_3.addWidget(self.label_30, 2, 2, 1, 1)
        self.label_27 = QtWidgets.QLabel(self.widget1)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_27.setFont(font)
        self.label_27.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_27.setObjectName("label_27")
        self.gridLayout_3.addWidget(self.label_27, 1, 0, 1, 1)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem, 0, 1, 1, 1)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem1, 1, 1, 1, 1)
        #
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.gridLayout_3.addItem(spacerItem1, 2, 1, 1, 1)
        #
        self.tabWidget_2.addTab(self.tab_4, "")
        self.lineEdit_6 = QtWidgets.QLineEdit(self.tab)
        self.lineEdit_6.setGeometry(QtCore.QRect(330, 470, 791, 22))
        self.lineEdit_6.setObjectName("lineEdit_6")
        self.label_14 = QtWidgets.QLabel(self.tab)
        self.label_14.setGeometry(QtCore.QRect(82, 110, 152, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_14.setFont(font)
        self.label_14.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_14.setObjectName("label_14")
        self.lineEdit_8 = QtWidgets.QLineEdit(self.tab)
        self.lineEdit_8.setGeometry(QtCore.QRect(240, 110, 251, 20))
        self.lineEdit_8.setText("")
        self.lineEdit_8.setObjectName("lineEdit_8")
        self.label_15 = QtWidgets.QLabel(self.tab)
        self.label_15.setGeometry(QtCore.QRect(80, 160, 51, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_15.setFont(font)
        self.label_15.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_15.setObjectName("label_15")
        self.lineEdit_9 = QtWidgets.QLineEdit(self.tab)
        self.lineEdit_9.setGeometry(QtCore.QRect(140, 160, 121, 20))
        self.lineEdit_9.setObjectName("lineEdit_9")
        #
        #self.lineEdit_9.setEnabled(False)
        #
        self.lineEdit_10 = QtWidgets.QLineEdit(self.tab)
        self.lineEdit_10.setGeometry(QtCore.QRect(380, 160, 111, 20))
        self.lineEdit_10.setObjectName("lineEdit_10")
        #
        #self.lineEdit_10.setEnabled(False)
        #
        self.label_16 = QtWidgets.QLabel(self.tab)
        self.label_16.setGeometry(QtCore.QRect(310, 160, 51, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_16.setFont(font)
        self.label_16.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_16.setObjectName("label_16")
        self.label_17 = QtWidgets.QLabel(self.tab)
        self.label_17.setGeometry(QtCore.QRect(80, 210, 191, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_17.setFont(font)
        self.label_17.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_17.setObjectName("label_17")
        self.label_18 = QtWidgets.QLabel(self.tab)
        self.label_18.setGeometry(QtCore.QRect(80, 250, 191, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_18.setFont(font)
        self.label_18.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_18.setObjectName("label_18")
        self.label_19 = QtWidgets.QLabel(self.tab)
        self.label_19.setGeometry(QtCore.QRect(410, 210, 81, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_19.setFont(font)
        self.label_19.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_19.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_19.setObjectName("label_19")
        self.label_20 = QtWidgets.QLabel(self.tab)
        self.label_20.setGeometry(QtCore.QRect(410, 250, 81, 20))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_20.setFont(font)
        self.label_20.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_20.setAlignment(QtCore.Qt.AlignRight|QtCore.Qt.AlignTrailing|QtCore.Qt.AlignVCenter)
        self.label_20.setObjectName("label_20")
        self.label_12 = QtWidgets.QLabel(self.tab)
        self.label_12.setGeometry(QtCore.QRect(82, 60, 152, 16))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_12.setFont(font)
        self.label_12.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_12.setObjectName("label_12")
        self.lineEdit_7 = QtWidgets.QLineEdit(self.tab)
        self.lineEdit_7.setGeometry(QtCore.QRect(240, 60, 251, 20))
        self.lineEdit_7.setText("")
        self.lineEdit_7.setObjectName("lineEdit_7")
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.layoutWidget = QtWidgets.QWidget(self.tab_2)
        self.layoutWidget.setGeometry(QtCore.QRect(740, 100, 371, 171))
        self.layoutWidget.setObjectName("layoutWidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.layoutWidget)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setHorizontalSpacing(18)
        self.gridLayout_2.setVerticalSpacing(20)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.label_24 = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_24.sizePolicy().hasHeightForWidth())
        self.label_24.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_24.setFont(font)
        self.label_24.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_24.setObjectName("label_24")
        self.gridLayout_2.addWidget(self.label_24, 2, 0, 1, 1)
        self.pushButton_4 = QtWidgets.QPushButton(self.layoutWidget)
        self.pushButton_4.setObjectName("pushButton_4")
        self.gridLayout_2.addWidget(self.pushButton_4, 0, 2, 1, 1)
        self.lineEdit_5 = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_5.setEnabled(False)
        self.lineEdit_5.setObjectName("lineEdit_5")
        self.gridLayout_2.addWidget(self.lineEdit_5, 2, 1, 1, 1)
        self.doubleSpinBox_3 = QtWidgets.QDoubleSpinBox(self.layoutWidget)
        self.doubleSpinBox_3.setEnabled(False)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.doubleSpinBox_3.sizePolicy().hasHeightForWidth())
        self.doubleSpinBox_3.setSizePolicy(sizePolicy)
        self.doubleSpinBox_3.setDecimals(0)
        self.doubleSpinBox_3.setStepType(QtWidgets.QAbstractSpinBox.DefaultStepType)
        self.doubleSpinBox_3.setProperty("value", 4.0)
        self.doubleSpinBox_3.setObjectName("doubleSpinBox_3")
        self.gridLayout_2.addWidget(self.doubleSpinBox_3, 1, 1, 1, 2)
        self.lineEdit_4 = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_4.setObjectName("lineEdit_4")
        self.gridLayout_2.addWidget(self.lineEdit_4, 0, 1, 1, 1)
        self.label_23 = QtWidgets.QLabel(self.layoutWidget)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_23.setFont(font)
        self.label_23.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_23.setObjectName("label_23")
        self.gridLayout_2.addWidget(self.label_23, 1, 0, 1, 1)
        self.label_4 = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_4.sizePolicy().hasHeightForWidth())
        self.label_4.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_4.setFont(font)
        self.label_4.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_4.setObjectName("label_4")
        self.gridLayout_2.addWidget(self.label_4, 0, 0, 1, 1)
        self.label_2 = QtWidgets.QLabel(self.tab_2)
        self.label_2.setGeometry(QtCore.QRect(220, 50, 201, 31))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(14)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_5 = QtWidgets.QLabel(self.tab_2)
        self.label_5.setGeometry(QtCore.QRect(810, 50, 221, 31))
        font = QtGui.QFont()
        font.setFamily("Calibri")
        font.setPointSize(14)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.widget2 = QtWidgets.QWidget(self.tab_2)
        self.widget2.setGeometry(QtCore.QRect(140, 100, 371, 351))
        self.widget2.setObjectName("widget2")
        self.gridLayout = QtWidgets.QGridLayout(self.widget2)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setHorizontalSpacing(18)
        self.gridLayout.setVerticalSpacing(20)
        self.gridLayout.setObjectName("gridLayout")
        self.label_25 = QtWidgets.QLabel(self.widget2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_25.sizePolicy().hasHeightForWidth())
        self.label_25.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_25.setFont(font)
        self.label_25.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_25.setObjectName("label_25")
        self.gridLayout.addWidget(self.label_25, 4, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.widget2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 0, 0, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.widget2)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 0, 1, 1, 1)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.widget2)
        self.lineEdit_3.setEnabled(True)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout.addWidget(self.lineEdit_3, 3, 1, 1, 1)
        self.pushButton_3 = QtWidgets.QPushButton(self.widget2)
        self.pushButton_3.setObjectName("pushButton_3")
        self.gridLayout.addWidget(self.pushButton_3, 0, 2, 1, 1)
        self.label_21 = QtWidgets.QLabel(self.widget2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_21.sizePolicy().hasHeightForWidth())
        self.label_21.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_21.setFont(font)
        self.label_21.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_21.setObjectName("label_21")
        self.gridLayout.addWidget(self.label_21, 3, 0, 1, 1)
        self.checkBox_4 = QtWidgets.QCheckBox(self.widget2)
        self.checkBox_4.setEnabled(False)
        self.checkBox_4.setText("")
        self.checkBox_4.setChecked(True)
        self.checkBox_4.setObjectName("checkBox_4")
        self.gridLayout.addWidget(self.checkBox_4, 1, 1, 1, 1)
        self.doubleSpinBox_2 = QtWidgets.QDoubleSpinBox(self.widget2)
        self.doubleSpinBox_2.setEnabled(False)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.doubleSpinBox_2.sizePolicy().hasHeightForWidth())
        self.doubleSpinBox_2.setSizePolicy(sizePolicy)
        self.doubleSpinBox_2.setDecimals(0)
        self.doubleSpinBox_2.setStepType(QtWidgets.QAbstractSpinBox.DefaultStepType)
        self.doubleSpinBox_2.setProperty("value", 10.0)
        self.doubleSpinBox_2.setObjectName("doubleSpinBox_2")
        self.gridLayout.addWidget(self.doubleSpinBox_2, 2, 1, 1, 2)
        self.lineEdit_12 = QtWidgets.QLineEdit(self.widget2)
        self.lineEdit_12.setEnabled(True)
        self.lineEdit_12.setObjectName("lineEdit_12")
        self.gridLayout.addWidget(self.lineEdit_12, 4, 1, 1, 1)
        self.label_26 = QtWidgets.QLabel(self.widget2)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_26.sizePolicy().hasHeightForWidth())
        self.label_26.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_26.setFont(font)
        self.label_26.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_26.setObjectName("label_26")
        self.gridLayout.addWidget(self.label_26, 5, 0, 1, 1)
        self.label_10 = QtWidgets.QLabel(self.widget2)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_10.setFont(font)
        self.label_10.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_10.setObjectName("label_10")
        self.gridLayout.addWidget(self.label_10, 1, 0, 1, 1)
        self.label_13 = QtWidgets.QLabel(self.widget2)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_13.setFont(font)
        self.label_13.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_13.setObjectName("label_13")
        self.gridLayout.addWidget(self.label_13, 2, 0, 1, 1)
        self.lineEdit_13 = QtWidgets.QLineEdit(self.widget2)
        self.lineEdit_13.setEnabled(True)
        self.lineEdit_13.setObjectName("lineEdit_13")
        self.gridLayout.addWidget(self.lineEdit_13, 5, 1, 1, 1)
        self.widget3 = QtWidgets.QWidget(self.tab_2)
        self.widget3.setGeometry(QtCore.QRect(40, 750, 1161, 22))
        self.widget3.setObjectName("widget3")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.widget3)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_8 = QtWidgets.QLabel(self.widget3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_8.setFont(font)
        self.label_8.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_8.setObjectName("label_8")
        self.horizontalLayout.addWidget(self.label_8)
        self.lineEdit_11 = QtWidgets.QLineEdit(self.widget3)
        self.lineEdit_11.setObjectName("lineEdit_11")
        self.horizontalLayout.addWidget(self.lineEdit_11)
        spacerItem2 = QtWidgets.QSpacerItem(70, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem2)
        self.label_22 = QtWidgets.QLabel(self.widget3)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.label_22.setFont(font)
        self.label_22.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.label_22.setObjectName("label_22")
        self.horizontalLayout.addWidget(self.label_22)
        spacerItem3 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem3)
        self.checkBox_5 = QtWidgets.QCheckBox(self.widget3)
        self.checkBox_5.setEnabled(True)
        self.checkBox_5.setText("")
        self.checkBox_5.setCheckable(True)
        self.checkBox_5.setChecked(False)
        self.checkBox_5.setObjectName("checkBox_5")
        self.horizontalLayout.addWidget(self.checkBox_5)
        spacerItem4 = QtWidgets.QSpacerItem(20, 20, QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem4)
        self.tabWidget.addTab(self.tab_2, "")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(350, 20, 651, 31))
        font = QtGui.QFont()
        font.setFamily("Arial")
        font.setPointSize(16)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_11 = QtWidgets.QLabel(self.centralwidget)
        self.label_11.setGeometry(QtCore.QRect(1060, 10, 211, 61))
        self.label_11.setObjectName("label_11")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1281, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        self.tabWidget_2.setCurrentIndex(1)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        now = datetime.now()
        MainWindow.setWindowTitle(_translate("MainWindow", "DS-20K QCC"))
        self.label_7.setText(_translate("MainWindow", "Output (.csv) File Directory:"))
        self.pushButton.setText(_translate("MainWindow", "Run"))
        self.tabWidget_2.setTabText(self.tabWidget_2.indexOf(self.tab_3), _translate("MainWindow", "Current Tile"))
        item = self.tableWidget.verticalHeaderItem(0)
        item.setText(_translate("MainWindow", "Min Cs:"))
        item = self.tableWidget.verticalHeaderItem(1)
        item.setText(_translate("MainWindow", "Max Cs:"))
        item = self.tableWidget.verticalHeaderItem(2)
        item.setText(_translate("MainWindow", "Range:"))
        item = self.tableWidget.verticalHeaderItem(3)
        item.setText(_translate("MainWindow", "Mean Cs:"))
        item = self.tableWidget.verticalHeaderItem(4)
        item.setText(_translate("MainWindow", "Mean Rs:"))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "Quad[1]"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "Quad[2]"))
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "Quad[3]"))
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "Quad[4]"))
        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)
        self.tableWidget.setSortingEnabled(__sortingEnabled)
        self.label_29.setText(_translate("MainWindow", "Test Conclusion:"))
        #
        self.label_31.setText(_translate("MainWindow", "Anomalies Found:"))
        self.label_32.setText(_translate("MainWindow", "UNKNOWN"))
        #
        self.label_28.setText(_translate("MainWindow", "UNKNOWN"))
        self.label_30.setText(_translate("MainWindow", "UNKNOWN"))
        self.label_27.setText(_translate("MainWindow", "Total Average Capacitance:"))
        self.tabWidget_2.setTabText(self.tabWidget_2.indexOf(self.tab_4), _translate("MainWindow", "Test Results"))
        self.lineEdit_6.setText(_translate("MainWindow", "E:\\"))
        self.label_14.setText(_translate("MainWindow", "Operator Name:"))
        self.label_15.setText(_translate("MainWindow", "Date:"))
        #
        self.date_string = now.strftime("%d/%m/%Y")
        self.date_time   = now.strftime("%H:%M:%S")
        self.timeStamp = self.date_time
        self.lineEdit_9.setText(_translate("MainWindow", self.date_string))
        self.lineEdit_10.setText(_translate("MainWindow", self.date_time))
        #
        self.label_16.setText(_translate("MainWindow", "Time:"))
        self.label_17.setText(_translate("MainWindow", "LCR Meter Connection Status:"))
        self.label_18.setText(_translate("MainWindow", "Matrix Connection Status:"))
        self.label_19.setText(_translate("MainWindow", "UNKNOWN"))
        self.label_20.setText(_translate("MainWindow", "UNKNOWN"))
        self.label_12.setText(_translate("MainWindow", "Device Serial Number:      "))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), _translate("MainWindow", "Test Interface"))
        self.label_24.setText(_translate("MainWindow", "Measurement Mode:"))
        self.pushButton_4.setText(_translate("MainWindow", "Connect"))
        self.lineEdit_5.setText(_translate("MainWindow", "4-Wire"))
        self.lineEdit_4.setText(_translate("MainWindow", "USB0::0x05E6::0x6510::04384100::0::INSTR"))
        self.label_23.setText(_translate("MainWindow", "Number Of Channels:"))
        self.label_4.setText(_translate("MainWindow", "Switch USB Address:"))
        self.label_2.setText(_translate("MainWindow", "LCR Meter Configuration:"))
        self.label_5.setText(_translate("MainWindow", "Switch Matrix Configuration:"))
        self.label_25.setText(_translate("MainWindow", "Voltage Level [mV]:"))
        self.label_3.setText(_translate("MainWindow", "E4980A GPIB Address:"))
        self.lineEdit_2.setText(_translate("MainWindow", "17"))
        self.lineEdit_3.setText(_translate("MainWindow", "CSRS"))
        self.pushButton_3.setText(_translate("MainWindow", "Connect"))
        self.label_21.setText(_translate("MainWindow", "Measurement Mode:"))
        self.lineEdit_12.setText(_translate("MainWindow", "100"))
        self.label_26.setText(_translate("MainWindow", "Frequency [KHz]:"))
        self.label_10.setText(_translate("MainWindow", "Auto Range:"))
        self.label_13.setText(_translate("MainWindow", "Number Of Averages:"))
        self.lineEdit_13.setText(_translate("MainWindow", "10"))
        self.label_8.setText(_translate("MainWindow", "Output (.csv) File Directory:"))
        self.lineEdit_11.setText(_translate("MainWindow", "D:\\"))
        self.label_22.setText(_translate("MainWindow", "Lock Write Dir:")) 
        #
        self.pushButton.clicked.connect(self.RunScript)
        self.checkBox_5.clicked.connect(self.lockWriteDir)
        self.pushButton_3.clicked.connect(self.retryConnectLCR)
        self.pushButton_4.clicked.connect(self.retryConnectSwitch)
        #        
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "Test Setup"))
        self.label.setText(_translate("MainWindow", "Darkside-20K Quadrant Capacitance Measurement Controller"))
        self.label_11.setText(_translate("MainWindow", "<html><head/><body><p><img src=\":/ukri_logo/ukri_logo_small.png\"/></p></body></html>"))

    def lockWriteDir(self):
        print("Write Directory locked/unlocked to/from Setup Page Default...")
        if self.checkBox_5.isChecked() == True:
            self.writeDirMode = "master"
            self.lineEdit_6.setText(self.lineEdit_11.text())
            self.lineEdit_6.setEnabled(False)
        elif self.checkBox_5.isChecked() == False:
            self.writeDirMode = "local"
            self.lineEdit_6.setEnabled(True)

    def retryConnectLCR(self):
        self.Address = self.lineEdit_2.text()
        print("Attempting to Connect to Instruement on Address: "+self.Address)
        # CONNECT TO LCR SIGNAL OUT HERE <<<    
        self.checkLCR = LCRMeter()
        self.L_okQ = self.checkLCR.instConnect(self.Address)
        if self.L_okQ == 1:
            self.printcmd("Connected to LCR Meter...")
        else:
            self.printcmd("Connection to LCR Meter Failed...")
        self.switchTabMain(0)

    def retryConnectSwitch(self):
        self.comPort = self.lineEdit_4.text()
        print("Attempting to Connect to Switch Controller on COM Port: "+self.comPort)
        # CONNECT TO SWITCH SIGNAL HERE <<<
        self.checkSwitch = SwitchSystem()
        self.S_okQ = self.checkSwitch.instConnect(self.comPort)
        if self.S_okQ == 1:
            self.printcmd("Connected to DAQ6510...")
        else:
            self.printcmd("Connection to DAQ6510 Failed...")
        self.switchTabMain(0)

    def printcmd(self, text):
        self.plainTextEdit.insertPlainText("\n["+self.timeStamp+"] >>> ")
        self.plainTextEdit.insertPlainText(text)
        self.plainTextEdit.update()
        self.plainTextEdit.verticalScrollBar().setValue(self.plainTextEdit.verticalScrollBar().maximum())

    def switchTabMini(self, state):
        self.tabWidget_2.setCurrentIndex(state)
    
    def switchTabMain(self, state):
        self.tabWidget.setCurrentIndex(state)

    #run button pressed function - generates worker thread 
    def RunScript(self):
        self.pushButton.setEnabled(False)
        self.thread = QThread()
        self.worker = Worker(
            serNO               = self.lineEdit_7.text(), 
            gpibAdr             = self.lineEdit_2.text(), 
            filePath            = self.lineEdit_6.text(), 
            User                = self.lineEdit_8.text(),
            switchTCPIP         = self.lineEdit_4.text(),
            nS                  = self.doubleSpinBox_2.value(),
            MODE                = self.lineEdit_3.text(),
            FREQ                = self.lineEdit_13.text(),
            LVL                 = self.lineEdit_12.text()
            )
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        # PYQT Signals from Worker thread to MAIN GUI:
        self.worker.passStatus.connect(self.changePassStatusLabel)
        self.worker.cmdPrint.connect(self.printcmd)
        self.worker.miniSwitch.connect(self.switchTabMini)
        self.worker.PBEnableDisable.connect(self.PBEnableDisable)
        self.worker.updateCol.connect(self.updateGUICOL)
        self.worker.LCRStatus.connect(self.changeGUILCRStatus)
        self.worker.SwitchStatus.connect(self.changeGUISwitchStatus)
        self.worker.triggerTable.connect(self.tableFill) # MAYBE?
        #
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.start()

    def PBEnableDisable(self, state):
        if state == 0:
            self.pushButton.setEnabled(False)
        elif state == 1:
            self.pushButton.setEnabled(True)
        else:
            self.pushButton.setEnabled(False)

    #function to start the timeKeeper background thread
    def startTimeKeeper(self):
        self.timeThread = QThread()
        self.timeKeeper = timeKeeper()
        self.timeKeeper.moveToThread(self.timeThread)
        self.timeThread.started.connect(self.timeKeeper.run)
        self.timeKeeper.timeDate.connect(self.updateTimeDate)
        #
        self.timeThread.start()

    #function called by/connected to timeKeeper background thread to update time and date in GUI ^
    def updateTimeDate(self, inputStr):
        _translate = QtCore.QCoreApplication.translate
        self.spltArr = inputStr.split(",")
        self.timeUpdate = self.spltArr[0]
        self.dateUpdate = self.spltArr[1]
        self.timeStamp = self.timeUpdate
        self.lineEdit_9.setText(_translate("MainWindow", self.dateUpdate))
        self.lineEdit_10.setText(_translate("MainWindow", self.timeUpdate))

    #function to fill out results table in second mini tab
    def tableFill(self, Quadrant, Cs_NPA, Rs_NPA):
        self.minCs          = np.amin(Cs_NPA)
        self.maxCs          = np.amax(Cs_NPA)
        self.rangeCs        = (self.maxCs - self.minCs)
        self.averageCs      = np.average(Cs_NPA)
        self.averageRs      = np.average(Rs_NPA)

        if Quadrant < 5:
            self.tableWidget.setItem(0, Quadrant, QtWidgets.QTableWidgetItem("{:.3e}".format(self.minCs)))
            self.tableWidget.setItem(1, Quadrant, QtWidgets.QTableWidgetItem("{:.3e}".format(self.maxCs)))
            self.tableWidget.setItem(2, Quadrant, QtWidgets.QTableWidgetItem("{:.3e}".format(self.rangeCs)))
            self.tableWidget.setItem(3, Quadrant, QtWidgets.QTableWidgetItem("{:.3e}".format(self.averageCs)))
            self.tableWidget.setItem(4, Quadrant, QtWidgets.QTableWidgetItem("{:.3e}".format(self.averageRs)))
        else:
            self.tableWidget.setRowCount(0)

    #following 
    def changeGUILCRStatus(self, status):
        _translate = QtCore.QCoreApplication.translate
        if status == 0:
            self.label_19.setStyleSheet("color: red")
            self.label_19.setText(_translate("MainWindow", "UNKNOWN"))
        elif status == 1:
            self.label_19.setStyleSheet("color: green")
            self.label_19.setText(_translate("MainWindow", "CONNECTED"))
        else:
            self.label_19.setStyleSheet("color: red")
            self.label_19.setText(_translate("MainWindow", "UNKNOWN"))
        
    def changeGUISwitchStatus(self, status):
        _translate = QtCore.QCoreApplication.translate
        if status == 0:
            self.label_20.setStyleSheet("color: red")
            self.label_20.setText(_translate("MainWindow", "UNKNOWN"))
        elif status == 1:
            self.label_20.setStyleSheet("color: green")
            self.label_20.setText(_translate("MainWindow", "CONNECTED"))
        else:
            self.label_20.setStyleSheet("color: red")
            self.label_20.setText(_translate("MainWindow", "UNKNOWN"))

    def changeAverageCapLabel(self, value, colour):
        _translate = QtCore.QCoreApplication.translate
        self.label_28.setStyleSheet("color: "+colour)
        self.label_28.setText(_translate("MainWindow", value))

    def changeAnomalyLabel(self, value, colour):
        _translate = QtCore.QCoreApplication.translate
        self.label_32.setStyleSheet("color: "+colour)
        self.label_32.setText(_translate("MainWindow", value))
    
    def changePassStatusLabel(self, state):
        _translate = QtCore.QCoreApplication.translate
        if state == 1:
            self.label_30.setText(_translate("MainWindow", "PASS"))
            self.label_30.setStyleSheet("color: green")
            self.printcmd("TEST PASSED")
        elif state == 0:
            self.label_30.setText(_translate("MainWindow", "FAIL"))
            self.label_30.setStyleSheet("color: red")
            self.printcmd("TEST FAILED")
        elif state == 2:
            self.label_30.setText(_translate("MainWindow", "UNKNOWN"))
            self.label_30.setStyleSheet("color: black")
        else:
            self.label_30.setText(_translate("MainWindow", "UNKNOWN"))
            self.label_30.setStyleSheet("color: black")

    def anomalyFlag(self):
        self.printcmd("CAUTION: Multiple Anomalies/Outliers found in samples..")

    def updateGUICOL(self, v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12):
        self.tab_3.inputCol(v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12)
        self.tab_3.update()

'''End Class'''

class Worker(QObject):
    finished            = pyqtSignal()
    progress            = pyqtSignal(int)
    passStatus          = pyqtSignal(int)
    LCRStatus           = pyqtSignal(int)
    SwitchStatus        = pyqtSignal(int)
    cmdPrint            = pyqtSignal(str)
    miniSwitch          = pyqtSignal(int)
    PBEnableDisable     = pyqtSignal(int)
    updateCol           = pyqtSignal(int, int, int, int, int, int, int, int, int, int, int, int)
    triggerTable        = pyqtSignal(int, object, object) # Quadrant, NPARRAY0, NPARRAY1

    def __init__(self, serNO, gpibAdr, filePath, User, switchTCPIP, nS, MODE, FREQ, LVL) -> None:
        super().__init__(parent=None)
        self.lcr        = LCRMeter()
        self.switch     = SwitchSystem()
        #
        ui.timeKeeper.timeDate_0.connect(self.updateTimeDate) # <<< ? - is this ok? 
        #
        self.serNO          = serNO
        self.GPIB           = gpibAdr
        self.filePath       = filePath
        self.User           = User
        self.switchTCPIP    = switchTCPIP
        self.nS             = nS
        self.ColArray       = [185, 185, 185, 185, 185, 185, 185, 185, 185, 185, 185, 185]
        self.MODE           = MODE
        self.FREQ           = FREQ
        self.LVL            = LVL
    
    # TEST THREAD
    #def run(self):
    #    print(">>>New Thread Created<<<")
    #    self.passStatus.emit(2)
    #    self.cmdPrint.emit("Test Initiated...")
    #    #
    #    self.connectLCROKQ = self.lcr.instConnect(self.GPIB) # Attempt to connect to LCR meter, 
    #    if self.connectLCROKQ == 0:
    #        self.cmdPrint.emit("LCR Connection Failed...") # Print string to command line
    #    elif self.connectLCROKQ == 1:
    #        self.cmdPrint.emit("LCR Connection Successful...")
    #    else:
    #        self.cmdPrint.emit("LCR Connection Failed...")
    #    #
    #    self.initialsofName = self.pullInitials(self.User) # Pull initials from "User" string and capitalise for CSV generation
    #    #
    #    self.miniSwitch.emit(0) # Switch to miniTab 1
    #    self.updateCol.emit(250, 250, 30, 185, 185, 185, 185, 185, 185, 185, 185, 185) # Update Colours of Quadrants on miniTab 1
    #    time.sleep(0.5) 
    #    self.updateCol.emit(0, 220, 0, 250, 250, 30, 185, 185, 185, 185, 185, 185)
    #    time.sleep(0.5)
    #    self.updateCol.emit(0, 220, 0, 0, 220, 0, 250, 250, 30, 185, 185, 185)
    #    time.sleep(0.5)
    #    self.updateCol.emit(0, 220, 0, 0, 220, 0, 0, 220, 0, 250, 250, 30)
    #    time.sleep(0.5)     
    #    self.updateCol.emit(0, 220, 0, 0, 220, 0, 0, 220, 0, 0, 220, 0)
    #    time.sleep(0.8)
    #    #
    #    self.miniSwitch.emit(1) # Switch to miniTab 2
    #    self.passStatus.emit(1) # Change Test Conclusion Status to "PASS"
    #    self.PBEnableDisable.emit(1) # Enable "Run" Button
    #    #
    #    self.finished.emit() # Close Worker

    def genCSV(self, Cs_NP_1, Rs_NP_1, Cs_NP_2, Rs_NP_2, Cs_NP_3, Rs_NP_3, Cs_NP_4, Rs_NP_4, SerNo, filePath, User):
        self.filePath               = filePath
        self.userInitials           = self.pullInitials(User)
        self.csvTimeStamp           = self.date.replace("/", "")+"_"+self.timeStamp.replace(":", "")
        #self.fileName               = SerNo + "_" + self.csvTimeStamp + ".csv"
        self.fileName               = self.csvTimeStamp + "_" + SerNo + ".csv"
        self.data                   = np.vstack((Cs_NP_1, Rs_NP_1, Cs_NP_2, Rs_NP_2, Cs_NP_3, Rs_NP_3, Cs_NP_4, Rs_NP_4))
        self.index                  = ["Quad 1 'Cs'", "Quad 1 'Rs'", "Quad 2 'Cs'", "Quad 2 'Rs'", "Quad 3 'Cs'", "Quad 3 'Rs'", "Quad 4 'Cs'", "Quad 4 'Rs'"]
        self.column                 = ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10"]
        self.newDF                  = pd.DataFrame(self.data, index=self.index, columns=self.column)
        # Add meta data and averages to dataframe:
        self.newDF.loc["Q1ACs:"]    = ["{:.3e}".format(np.average(Cs_NP_1)), "Q2ACs:", "{:.3e}".format(np.average(Cs_NP_2)), "Q3ACs:", "{:.3e}".format(np.average(Cs_NP_3)), "Q4ACs:", "{:.3e}".format(np.average(Cs_NP_4)), " ", " ", " "]
        self.newDF.loc["Q1ARs:"]    = ["{:.3e}".format(np.average(Rs_NP_1)), "Q2ARs:", "{:.3e}".format(np.average(Cs_NP_1)), "Q3ARs:", "{:.3e}".format(np.average(Cs_NP_1)), "Q4ARs:", "{:.3e}".format(np.average(Cs_NP_1)), " ", " ", " "]
        self.newDF.loc["Freq:"]     = [self.FREQ+"KHz", "Amp:", self.LVL+"mV", "MM:", self.MODE, "Operator:", self.userInitials, "S/N:", SerNo, " "]
        #
        if self.filePath[-1] != "\\":
            self.filePath           = self.filePath + "\\" # Check last character is a \
        #
        if self.checkFN(self.fileName, self.filePath) == 1:
            self.cmdPrint.emit("'.csv' file generated...")
            self.newDF.to_csv(self.filePath + self.fileName, index=True, header=True)
        elif self.checkFN(self.fileName, self.filePath) == 0:
            self.cmdPrint.emit("'.csv' file with this name and file path already exists...")
        else:
            self.cmdPrint.emit("Directory Not Found...")

    def checkFN(self, fileName, filePath):
        try:
            self.inDir          = os.listdir(filePath)
        except:
            return 2
        else:
            if fileName in self.inDir:
                return 0
            else:
                return 1

    def updateTimeDate(self, inputStr):
        self.spltArr        = inputStr.split(",")
        self.timeUpdate     = self.spltArr[0]
        self.dateUpdate     = self.spltArr[1]
        self.timeStamp      = self.timeUpdate
        self.date           = self.dateUpdate

    def pullInitials(self, fullname):
        self.xs = (fullname)
        self.name_list = self.xs.split()
        self.initials = ""
        for self.name in self.name_list:  # go through each name
            self.initials += self.name[0].upper()  # append the initial
        return self.initials

    def updateTile(self, tileQ, R, G, B):
        if tileQ == 1:
            self.ColArray[0] = R
            self.ColArray[1] = G
            self.ColArray[2] = B
        elif tileQ == 2:
            self.ColArray[3] = R
            self.ColArray[4] = G
            self.ColArray[5] = B
        elif tileQ == 3:
            self.ColArray[6] = R
            self.ColArray[7] = G
            self.ColArray[8] = B
        elif tileQ == 4:
            self.ColArray[9] = R
            self.ColArray[10] = G
            self.ColArray[11] = B
        self.updateCol.emit(self.ColArray[0], 
                            self.ColArray[1], 
                            self.ColArray[2], 
                            self.ColArray[3], 
                            self.ColArray[4], 
                            self.ColArray[5], 
                            self.ColArray[6], 
                            self.ColArray[7], 
                            self.ColArray[8], 
                            self.ColArray[9], 
                            self.ColArray[10], 
                            self.ColArray[11]
                            )

    # Assess results per quadrant and colour
    def QCU(self, avg_C, tile):
        if avg_C >= 28e-9 and avg_C <= 36e-9:
            self.updateTile(tile, 0, 220, 0)
        elif avg_C < 28e-9 or avg_C > 36e-9:
            self.updateTile(tile, 250, 250, 30)

    #def QCU(self, avg_C, tile):
    #    self.error          = abs(avg_C - 33e-9)
    #    mapVal              = interp1d([0, 10e-9], [0, 10])
    #    self.mapError       = mapVal(self.error)

    # RUN STATEMENT
    def run(self):
        print(">>>New 'Run' Thread Created<<<")
        self.passStatus.emit(2)
        self.cmdPrint.emit("Test Initiated...")
        #
        self.connect_LCR_OKQ = self.lcr.instConnect(self.GPIB) # Attempt to connect to LCR meter, 
        if self.connect_LCR_OKQ == 1:
            self.connect_Switch_OKQ = self.switch.instConnect(self.switchTCPIP) # Attempt to connect to Switch,
            self.cmdPrint.emit("Connected to LCR Meter OK...")
            self.LCRStatus.emit(1)
            if self.connect_Switch_OKQ == 1:
                self.cmdPrint.emit("Connected to DAQ6510 OK...")  
                self.SwitchStatus.emit(1)
                # Initialise Instruments HERE, and perform test > 
                #
                self.lcr.config(self.MODE, self.FREQ, self.LVL) # Config LCR measurement mode etc...
                self.switch.config() # Config Switch Card
                #
                #    >>> Begin Measurements <<<
                #
                self.updateTile(1, 185, 185, 185)   # set all tiles to grey
                self.updateTile(2, 185, 185, 185)
                self.updateTile(3, 185, 185, 185)
                self.updateTile(4, 185, 185, 185)
                #self.triggerTable.emit(5)
                #
                self.miniSwitch.emit(0)
                self.switch.setTile(1)
                time.sleep(0.5)
                self.avg_Cs_1, self.avg_Rs_1, self.Cs_np_1, self.Rs_np_1 = self.lcr.get_AVG_nSamp(self.nS)
                self.QCU(self.avg_Cs_1, 1)
                self.triggerTable.emit(0, self.Cs_np_1, self.Rs_np_1)
                #
                self.switch.setTile(2)
                time.sleep(0.5)
                self.avg_Cs_2, self.avg_Rs_2, self.Cs_np_2, self.Rs_np_2 = self.lcr.get_AVG_nSamp(self.nS)
                self.QCU(self.avg_Cs_2, 2)
                self.triggerTable.emit(1, self.Cs_np_2, self.Rs_np_2)
                #
                self.switch.setTile(3)
                time.sleep(0.5)
                self.avg_Cs_3, self.avg_Rs_3, self.Cs_np_3, self.Rs_np_3 = self.lcr.get_AVG_nSamp(self.nS)
                self.QCU(self.avg_Cs_3, 3)
                self.triggerTable.emit(2, self.Cs_np_3, self.Rs_np_3)
                #
                self.switch.setTile(4)
                time.sleep(0.5)
                self.avg_Cs_4, self.avg_Rs_4, self.Cs_np_4, self.Rs_np_4 = self.lcr.get_AVG_nSamp(self.nS)
                self.QCU(self.avg_Cs_4, 4)
                self.triggerTable.emit(3, self.Cs_np_4, self.Rs_np_4)
                #
                print(self.avg_Cs_1)
                print(self.avg_Cs_2)
                print(self.avg_Cs_3)
                print(self.avg_Cs_4)
                #
                self.miniSwitch.emit(1)
                #
                self.genCSV(self.Cs_np_1, 
                            self.Rs_np_1, 
                            self.Cs_np_2, 
                            self.Rs_np_2, 
                            self.Cs_np_3, 
                            self.Rs_np_3, 
                            self.Cs_np_4, 
                            self.Rs_np_4, 
                            self.serNO, 
                            self.filePath, 
                            self.User
                            )
                #
                self.cmdPrint.emit("Test Completed...")
                #
            else:
                self.cmdPrint.emit("Connection to DAQ6510 Failed...")
                self.SwitchStatus.emit(0)
        else:
            self.cmdPrint.emit("Connection to LCR Meter Failed...")
            self.LCRStatus.emit(0)
        #
        self.PBEnableDisable.emit(1) # Enable "Run" Button
        self.finished.emit()

'''End Class'''

class TileWidget(QtWidgets.QWidget):

    def __init__(self):
        super().__init__()
        self.hex1 = 185
        self.hex2 = 185
        self.hex3 = 185
        self.hex4 = 185
        self.hex5 = 185
        self.hex6 = 185
        self.hex7 = 185
        self.hex8 = 185
        self.hex9 = 185
        self.hex10 = 185
        self.hex11 = 185
        self.hex12 = 185

    def inputCol(self, val1, val2, val3, val4, val5, val6, val7, val8, val9, val10, val11, val12):
        self.hex1 = val1
        self.hex2 = val2
        self.hex3 = val3
        self.hex4 = val4
        self.hex5 = val5
        self.hex6 = val6
        self.hex7 = val7
        self.hex8 = val8
        self.hex9 = val9
        self.hex10 = val10
        self.hex11 = val11
        self.hex12 = val12

    def paintEvent(self, event):
        quad1 = QtGui.QPainter(self)
        #quad1.begin(self)
        self.drawTiles(quad1, self.hex1, self.hex2, self.hex3, self.hex4, self.hex5, self.hex6, self.hex7, self.hex8, self.hex9, self.hex10, self.hex11, self.hex12)
        #quad1.end()
    
    def drawTiles(self, quad1, hex1, hex2, hex3, hex4, hex5, hex6, hex7, hex8, hex9, hex10, hex11, hex12):
        quad1.setBrush(QtGui.QColor(hex1, hex2, hex3))
        quad1.drawRect(120, 20, 50, 70)
        quad1.drawRect(180, 20, 50, 70)
        quad1.drawRect(120, 110, 50, 70)
        quad1.drawRect(180, 110, 50, 70)
        quad1.drawRect(240, 20, 50, 70)
        quad1.drawRect(240, 110, 50, 70)

        quad1.setBrush(QtGui.QColor(hex4, hex5, hex6))
        quad1.drawRect(300, 20, 50, 70)
        quad1.drawRect(300, 110, 50, 70)
        quad1.drawRect(360, 20, 50, 70)
        quad1.drawRect(420, 20, 50, 70)
        quad1.drawRect(360, 110, 50, 70)
        quad1.drawRect(420, 110, 50, 70)
        
        quad1.setBrush(QtGui.QColor(hex7, hex8, hex9))
        quad1.drawRect(120, 200, 50, 70)
        quad1.drawRect(120, 290, 50, 70)
        quad1.drawRect(180, 200, 50, 70)
        quad1.drawRect(180, 290, 50, 70)
        quad1.drawRect(240, 200, 50, 70)
        quad1.drawRect(240, 290, 50, 70)

        quad1.setBrush(QtGui.QColor(hex10, hex11, hex12))
        quad1.drawRect(300, 200, 50, 70)
        quad1.drawRect(300, 290, 50, 70)
        quad1.drawRect(360, 200, 50, 70)
        quad1.drawRect(420, 200, 50, 70)
        quad1.drawRect(360, 290, 50, 70)
        quad1.drawRect(420, 290, 50, 70)

'''End Class'''

class LCRMeter():

    def __init__(self):
        super().__init__()

    def instConnect(self, Address):
        global rm
        rm = pyvisa.ResourceManager()
        try:
            self.E4980A              = rm.open_resource("GPIB0::"+Address+"::INSTR")
        except:
            self.connectState = False
            print("Connection Failed...")
            print("Error:", sys.exc_info()[0])
            return 0
        else:
            self.connectState = True 
            self.E4980A.timeout      = 8000 #set timeout limit [ms]
            return 1

    def instDisconnect(self):
        try:
            self.E4980A.close()
            self.connectState = False
        except:
            self.connectState = True

    def config(self, MODE, FREQ, LVL):
        self.condFreq = float(FREQ)*1000
        self.condLVL = float(LVL)/1000
        self.E4980A.write("*RST")
        self.E4980A.write("BIAS:STATe OFF")
        self.E4980A.write("FUNCtion:IMPedance:TYPE "+MODE)
        self.E4980A.write("FREQuency:CW "+str(self.condFreq))
        self.E4980A.write(":VOLTage:LEVel "+str(round(self.condLVL, 3)))
        self.E4980A.write(":APERture LONG")
        self.E4980A.write(":INITiate:CONTinuous ON")
        self.E4980A.write(":TRIG:SOUR BUS")

    def MEAS_conv(self, meas_comb_raw):
        meas_split      = meas_comb_raw.split(",")
        raw_Cs          = float(meas_split[0])
        raw_Rs          = float(meas_split[1])
        return raw_Cs, raw_Rs

    def get_AVG_nSamp(self, n_of_samples) :
        avg_arr_Cs = []
        avg_arr_Rs = []
        for i in range(int(n_of_samples)) :
            self.E4980A.write("*TRG")
            Cs_Rs_val           = self.E4980A.query(":FETCh?")
            raw_Cs, raw_Rs      = self.MEAS_conv(Cs_Rs_val)
            avg_arr_Cs.append(raw_Cs)
            avg_arr_Rs.append(raw_Rs)
            time.sleep(0.5)
        Cs_np = np.asarray(avg_arr_Cs)
        Rs_np = np.asarray(avg_arr_Rs)
        avg_Cs = np.average(Cs_np)
        avg_Rs = np.average(Rs_np)
        return avg_Cs, avg_Rs, Cs_np, Rs_np

'''End Class'''

class SwitchSystem():

    def __init__(self):
        super().__init__()

    def instConnect(self, Address):
        try:
            #self.DAQ6510              = rm.open_resource("GPIB0::"+Address+"::INSTR")
            self.DAQ6510              = rm.open_resource(Address)
        except:
            self.connectState = False
            print("Connection Failed...")
            print("Error:", sys.exc_info()[0])
            return 0
        else:
            self.connectState = True 
            self.DAQ6510.timeout      = 8000 #set timeout limit [ms]
            return 1

    def instDisconnect(self):
        try:
            self.DAQ6510.close()
        except:
            self.connectState = True
            return 1
        else:
            self.connectState = False
            return 0

    def config(self):
        # Configuration Function for Switch Mainframe
        self.DAQ6510.write("*RST") # Reset DMM 
        self.DAQ6510.write("*LANG SCPI") 
        # Isolate DMM 
        self.DAQ6510.write("ROUT:OPEN (@124, 125)") # Open Backplane Isolation Relays

    def setTile(self, Tile):
        if Tile == 1:
            self.DAQ6510.write("ROUT:OPEN:ALL")
            self.DAQ6510.write("ROUT:MULT:CLOS (@123)")
            self.DAQ6510.write("ROUT:MULT:CLOS (@110, 120)")
        elif Tile == 2:
            self.DAQ6510.write("ROUT:OPEN:ALL")
            self.DAQ6510.write("ROUT:MULT:CLOS (@123)")
            self.DAQ6510.write("ROUT:MULT:CLOS (@109, 119)")
        elif Tile == 3:
            self.DAQ6510.write("ROUT:OPEN:ALL")
            self.DAQ6510.write("ROUT:MULT:CLOS (@123)")
            self.DAQ6510.write("ROUT:MULT:CLOS (@108, 118)")
        elif Tile == 4:
            self.DAQ6510.write("ROUT:OPEN:ALL")
            self.DAQ6510.write("ROUT:MULT:CLOS (@123)")
            self.DAQ6510.write("ROUT:MULT:CLOS (@107, 117)")
        else:
            self.DAQ6510.write("ROUT:OPEN:ALL")

'''End Class'''

if __name__ == "__main__":
    import sys
    #
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    ui.printcmd("GUI Initialised...")
    #
    ui.switchTabMini(0)
    MainWindow.show()
    sys.exit(app.exec_())