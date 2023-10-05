import datetime
import os
import platform
import sys
import traceback
import subprocess
import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QLabel, QVBoxLayout, QWidget, \
    QMessageBox, QProgressDialog, QCheckBox, QHBoxLayout, QSpacerItem, QSizePolicy, QFrame, QTableWidget, \
    QTableWidgetItem, QLineEdit, QStatusBar
from PyQt5.QtCore import Qt, QDate, QSettings
from PyQt5 import QtWidgets, QtCore
from ReadOutput import *


class 主介面(QMainWindow):
    def __init__(self):
        super().__init__()
        self.文件選擇 = 子介面_文件選擇()
        self.核對表格 = 子介面_核對表格()
        # 判斷使用者是否有選擇功能執行，若有，就會在執行功能後輸出文件
        self.執行確認 = False
        self.initUI()

    def initUI(self):
        self.setStyleSheet("""
                            QLabel {
                                color: #FF0000;
                                font-size: 18px;
                            }
                            QMainWindow {
                                background-color: #FDFEAA;
                            }
                        """)

        # 介面標題與大小
        self.setWindowTitle("Output核對")
        self.setGeometry(300, 300, 700, 300)

        # 創建一個 主佈局(水平)用來容納左右兩側的布局
        Mainlayout = QtWidgets.QHBoxLayout()

        # 左右兩側的布局，垂直為主
        left_layout = QtWidgets.QVBoxLayout()
        right_layout = QtWidgets.QVBoxLayout()

        # 添加左右兩側的布局
        Mainlayout.addLayout(left_layout)
        Mainlayout.addLayout(right_layout)

        # 添加一個 QLabel 顯示標題
        title_label = QLabel("Output核對", self)
        title_label.setStyleSheet("font-size: 24px; font-weight: bold;")

        # 創建一個狀態列
        self.執行狀態列 = QStatusBar()
        self.執行狀態列.setStyleSheet("font-size: 15px; font-weight: bold;")
        self.執行狀態列.setFixedSize(150, 20)

        # 匯入每日Output和尾數
        self.匯入按鈕 = QPushButton("匯入Output")
        self.匯入按鈕.clicked.connect(self.Import)
        self.匯入按鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.匯入按鈕.setFixedSize(150, 30)

        # 匯出Output檔案
        self.匯出按鈕 = QPushButton("匯出Output", self)
        self.匯出按鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.匯出按鈕.setFixedSize(150, 30)
        self.匯出按鈕.clicked.connect(self.Export_Output)

        # 顯示匯入與核對資料
        self.顯示列表按鈕 = QPushButton("顯示列表")
        self.顯示列表按鈕.clicked.connect(self.ShowTable)
        self.顯示列表按鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.顯示列表按鈕.setFixedSize(150, 30)

        # 核對DIP檔案
        self.核對DIP按鈕 = QPushButton("DIP移轉核對", self)
        self.核對DIP按鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.核對DIP按鈕.setFixedSize(150, 30)
        self.核對DIP按鈕.clicked.connect(self.Output_check)

        # layout.addWidget(title_label, alignment=Qt.AlignCenter)
        left_layout.addWidget(title_label, alignment=Qt.AlignCenter)
        left_layout.addWidget(self.文件選擇)
        right_layout.addSpacerItem(QSpacerItem(40, 40, QSizePolicy.Minimum, QSizePolicy.Minimum))
        right_layout.addWidget(self.匯入按鈕, alignment=Qt.AlignTop)
        right_layout.addWidget(self.匯出按鈕, alignment=Qt.AlignTop)
        right_layout.addWidget(self.顯示列表按鈕, alignment=Qt.AlignTop)
        right_layout.addWidget(self.核對DIP按鈕, alignment=Qt.AlignTop)
        right_layout.addWidget(self.執行狀態列, alignment=Qt.AlignBottom)

        # 创建一个 QWidget 作为布局的容器
        container = QWidget(self)
        container.setLayout(Mainlayout)
        container.setGeometry(0, 0, 700, 300)

        self.loadSettings()

    def Import(self):
        try:
            if self.文件選擇.Output選擇確認:
                self.執行狀態列.showMessage('匯入資料中....', 0)
                Output = Read_Output(self.文件選擇.Output)
                self.核對表格.ConcatTable(匯入資料=Output)
                QMessageBox.information(self, "結果", "匯入資料新增成功")
                self.執行狀態列.showMessage('匯入完成', 2000)
            else:
                QMessageBox.warning(self, "警告", "尚未選擇Output文件!")
                self.執行狀態列.showMessage('匯入失敗!', 2000)
                return
        except:
            error_message = traceback.format_exc()
            QMessageBox.warning(self, "匯入功能錯誤", f"錯誤 : {error_message}")

    def Export_Output(self):
        # 獲取桌面路徑
        系統 = platform.system()
        目錄 = os.path.expanduser("~")

        if 系統 == "Windows":
            桌面路徑 = os.path.join(目錄, "Desktop").replace('\\', '/')
        elif 系統 == "Darwin":  # macOS
            桌面路徑 = os.path.join(目錄, "Desktop").replace('\\', '/')
        elif 系統 == "Linux":
            桌面路徑 = os.path.join(目錄, "Desktop").replace('\\', '/')
        else:
            # 默認返回用戶主目錄
            桌面路徑 = 目錄.replace('\\', '/')

        # 預設路徑 若此路徑不存在，則採用桌面路徑
        預設路徑 = '//file-server/生管部/五股廠/Jimmy/Output匯出檔'

        # 設定檔名
        當天日期_文字格式 = datetime.datetime.now().strftime('%y%m%d%H%M')
        輸出檔名 = 'Output尾數檔' + 當天日期_文字格式 + '.xlsx'

        if self.核對表格.儲存資料.empty:
            confirm = QMessageBox.question(self, "確認", "Output資料為空，確定進行輸出?",
                                           QMessageBox.Yes | QMessageBox.No)
            if confirm == QMessageBox.Yes:
                if os.path.exists(預設路徑):
                    self.核對表格.儲存資料.to_excel(f'{預設路徑}/{輸出檔名}', index=False, sheet_name='匯出檔')
                    self.執行狀態列.showMessage('檔案輸出完成!', 2000)
                    QMessageBox.information(self, '結果', 'Output尾數檔案輸出完成!')
                else:
                    self.核對表格.儲存資料.to_excel(f'{桌面路徑}/{輸出檔名}', index=False, sheet_name='匯出檔')
                    self.執行狀態列.showMessage('檔案輸出完成!', 2000)
                    QMessageBox.information(self, '結果', '預設路徑不存在，Output尾數檔案已輸出至桌面!')

        else:
            confirm = QMessageBox.question(self, "確認", "即將輸出Output資料，確定執行?",
                                           QMessageBox.Yes | QMessageBox.No)
            if confirm == QMessageBox.Yes:
                if os.path.exists(預設路徑):
                    self.核對表格.儲存資料.to_excel(f'{預設路徑}/{輸出檔名}', index=False, sheet_name='匯出檔')
                    self.執行狀態列.showMessage('檔案輸出完成!', 2000)
                    QMessageBox.information(self, '結果', 'Output尾數檔案輸出完成!')
                else:
                    self.核對表格.儲存資料.to_excel(f'{桌面路徑}/{輸出檔名}', index=False, sheet_name='匯出檔')
                    self.執行狀態列.showMessage('檔案輸出完成!', 2000)
                    QMessageBox.information(self, '結果', '預設路徑不存在，Output尾數檔案已輸出至桌面!')

    def Output_check(self):
        try:
            if self.文件選擇.DIP檔案選擇確認:
                if self.核對表格.儲存資料.empty:
                    if not self.文件選擇.Output選擇確認:
                        QMessageBox.warning(self, '警告', 'Output資料為空，請先選擇Outptut文件匯入!')
                        return
                    else:
                        QMessageBox.warning(self, '警告', 'Output資料為空，請確認是否已匯入資料!')
                        return
                else:
                    self.執行狀態列.showMessage('檔案核對中...', 0)
                    DIP = Read_DIP(self.文件選擇.DIP)
                    for 足標, 欄位 in self.核對表格.儲存資料.iterrows():
                        if 欄位['母工單單號'] in DIP.index:
                            self.核對表格.儲存資料.at[足標, '尾數'] = DIP.loc[欄位['母工單單號'], '尾數 ']
                            self.核對表格.儲存資料.at[足標, '移轉小記'] = DIP.loc[欄位['母工單單號'], '移轉小記']
                            self.核對表格.儲存資料.at[足標, '總計'] = DIP.loc[欄位['母工單單號'], '總計']
                            self.核對表格.儲存資料.at[足標, '餘數'] = DIP.loc[欄位['母工單單號'], '餘數']
                            self.核對表格.儲存資料.at[足標, '註記'] = DIP.loc[欄位['母工單單號'], 'Unnamed: 15']
                        elif 欄位['工號'] in DIP.index:
                            self.核對表格.儲存資料.at[足標, '尾數'] = DIP.loc[欄位['工號'], '尾數 ']
                            self.核對表格.儲存資料.at[足標, '移轉小記'] = DIP.loc[欄位['工號'], '移轉小記']
                            self.核對表格.儲存資料.at[足標, '總計'] = DIP.loc[欄位['工號'], '總計']
                            self.核對表格.儲存資料.at[足標, '餘數'] = DIP.loc[欄位['工號'], '餘數']
                            self.核對表格.儲存資料.at[足標, '註記'] = DIP.loc[欄位['工號'], 'Unnamed: 15']
                        else:
                            self.核對表格.儲存資料.at[足標, '尾數'] = '查無資料'

                    self.核對表格.UpdateTable()
                    self.執行狀態列.showMessage('核對完成!', 2000)
                    QMessageBox.information(self, '結果', 'DIP移轉紀錄已核對完成，並填入表格中!')
            else:
                QMessageBox.warning(self, '警告', '尚未選擇DIP文件!')
                return
        except:
            error_message = traceback.format_exc()
            QMessageBox.warning(self, "核對功能錯誤", f"錯誤 : {error_message}")

    def ShowTable(self):
        self.核對表格.show()

    def closeEvent(self, event):
        self.saveSettings()
        event.accept()

    def saveSettings(self):
        settings = QtCore.QSettings("PSI", "Output_check")
        settings.setValue("Output暫存資料", self.核對表格.儲存資料)
        pass

    def loadSettings(self):
        settings = QtCore.QSettings("PSI", "Output_check")
        self.核對表格.儲存資料 = settings.value("Output暫存資料", '初次載入')
        self.核對表格.Import_data_and_table()
        pass


class 子介面_文件選擇(QMainWindow):
    def __init__(self):
        super().__init__()
        self.文件夾選擇確認 = False
        self.Output選擇確認 = False
        self.DIP檔案選擇確認 = False
        self.輸出位置選擇確認 = False
        self.initUI()

    def initUI(self):
        # 創建一個 水平佈局
        layout = QVBoxLayout()
        # layout = QtWidgets.QGridLayout()

        # 添加一個

        self.button_Output檔案選擇 = QPushButton("Output檔案", self)
        self.button_Output檔案選擇.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.button_Output檔案選擇.setFixedSize(150, 30)
        self.button_Output檔案選擇.clicked.connect(self.selectOutputFile)

        self.button_DIP檔案選擇 = QPushButton("DIP檔案", self)
        self.button_DIP檔案選擇.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.button_DIP檔案選擇.setFixedSize(150, 30)
        self.button_DIP檔案選擇.clicked.connect(self.selectDIPFile)

        self.選擇文件顯示 = QLineEdit(self)
        self.選擇文件顯示.setStyleSheet(
            "font-size: 18px;border: 2px groove black; background-color: white; padding: 2px;")
        self.選擇文件顯示.setFixedSize(500, 25)
        self.選擇文件顯示.setReadOnly(True)

        self.DIP文件顯示 = QLineEdit(self)
        self.DIP文件顯示.setStyleSheet(
            "font-size: 18px;border: 2px groove black; background-color: white; padding: 2px;")
        self.DIP文件顯示.setFixedSize(500, 25)
        self.DIP文件顯示.setReadOnly(True)

        # 另外創建一個水平布局

        layout.addWidget(self.button_Output檔案選擇)
        layout.addWidget(self.選擇文件顯示)
        layout.addWidget(self.button_DIP檔案選擇)
        layout.addWidget(self.DIP文件顯示)

        layout.setAlignment(Qt.AlignVCenter | Qt.AlignLeft)

        # 創建一個 QWidget 作為佈局的容器
        container = QWidget(self)
        container.setLayout(layout)
        container.setGeometry(0, 0, 700, 250)

    def selectDirectory(self):
        # 創建一個 QFileDialog 的實例，用於顯示文件對話框。
        文件選擇視窗 = QFileDialog()

        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        文件夾路徑 = 文件選擇視窗.getExistingDirectory(self, "選擇檔案")
        self.文件夾選擇確認 = 文件夾路徑
        self.文件夾 = 文件夾路徑
        self.選擇資料夾顯示.setText(f"{文件夾路徑}")

    def selectOutputFile(self):
        預設文件夾路徑 = "//file-server/生管部/五股廠/Jimmy/Output輸出檔"
        文件選擇視窗 = QFileDialog()
        文件選擇視窗.setDirectory(預設文件夾路徑)

        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        檔案路徑, _ = 文件選擇視窗.getOpenFileName(self, "選擇檔案")
        self.Output選擇確認 = 檔案路徑
        self.Output = 檔案路徑
        self.選擇文件顯示.setText(f"{檔案路徑}")

    def selectDIPFile(self):
        預設文件夾路徑 = "//file-server/共用區/DIP/DIP-每日入帳紀錄"
        文件選擇視窗 = QFileDialog()
        文件選擇視窗.setDirectory(預設文件夾路徑)

        # 使用變數 file_path 來接收文件路徑，而 _ 變數表示我們不關心文件類型
        檔案路徑, _ = 文件選擇視窗.getOpenFileName(self, "選擇檔案")
        self.DIP檔案選擇確認 = 檔案路徑
        self.DIP = 檔案路徑
        self.DIP文件顯示.setText(f"{檔案路徑}")

    def OpenDocument(self):
        # 電腦讀取的文件路徑和程式所需的文件路徑不一樣
        # 所以要用指令要求電腦打開指定文件夾時，要先將正斜號(/)替換為反斜號(\)
        fixed_path = self.輸出位置.replace('/', '\\')
        subprocess.Popen(['explorer', fixed_path])


class 子介面_核對表格(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("Output核對內容")
        self.setGeometry(100, 100, 1500, 600)

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        # 主要佈局與按鈕佈局
        Mainlayout = QVBoxLayout(self.central_widget)
        Button_layout = QHBoxLayout(self.central_widget)

        # 以表格顯示Excel資料
        self.tableWidget = QTableWidget(self)
        self.tableWidget.setAlternatingRowColors(True)  # 每行顏色交錯顯示
        self.tableWidget.verticalHeader().setVisible(False)  # 垂直標題關閉
        self.tableWidget.setStyleSheet("font-size: 14px;font-weight: bold")

        # 資料筆數標籤
        self.資料比數標籤 = QLabel()
        self.資料比數標籤.setStyleSheet("font-size: 16px;font-weight: bold")
        self.資料比數標籤.setFixedSize(150, 30)

        # 添加刪除特定項目的按鈕
        self.刪除按鈕 = QPushButton("刪除選定項目", self)
        self.刪除按鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.刪除按鈕.setFixedSize(150, 30)
        self.刪除按鈕.clicked.connect(self.delete_selected_rows)

        # 添加清空所有項目的按鈕
        self.清空按鈕 = QPushButton("清空所有項目", self)
        self.清空按鈕.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.清空按鈕.setFixedSize(150, 30)
        self.清空按鈕.clicked.connect(self.deleteALL)

        self.搜索框 = QLineEdit(self)
        self.搜索框.setStyleSheet("font-size: 16px;font-family: 新細明體;font-weight: bold")
        self.搜索框.setFixedSize(400, 30)
        self.搜索框.setPlaceholderText('請輸入工單...')  # 文字框內的提示字，使用者輸入內容後會消失
        self.搜索框.textChanged.connect(self.filterTable)  # 文字框變化時的函數就會執行

        # 佈局管理
        Button_layout.addWidget(self.搜索框, alignment=Qt.AlignLeft)
        Button_layout.addWidget(self.刪除按鈕)
        Button_layout.addWidget(self.清空按鈕)
        Button_layout.addWidget(self.資料比數標籤)

        Mainlayout.addLayout(Button_layout)
        Mainlayout.addWidget(self.tableWidget)

    def Import_data_and_table(self):
        # 給完全是第一次使用者的設定，系統判斷後會將儲存資料設置為Dataframe
        if isinstance(self.儲存資料, str) and self.儲存資料 == '初次載入':
            self.儲存資料 = pd.DataFrame(columns=['母工單單號', '工號', '名稱規格', '工令量', 'SOURCE',
                                                  '尾數', '移轉小記', '總計', '餘數', '註記'])
            self.資料比數標籤.setText(f'資料筆數: {len(self.儲存資料)}/{self.tableWidget.rowCount()}')
            QMessageBox.information(self, '結果', '初次載入設定完成，可以開始匯入資料核對!')

        # 先前設定過但沒匯入資料，就不再次設定，直接略過
        elif isinstance(self.儲存資料, pd.DataFrame):
            if self.儲存資料.empty:
                self.資料比數標籤.setText(f'資料筆數: {len(self.儲存資料)}/{self.tableWidget.rowCount()}')
                pass
            # 若有資料則匯入表格內
            else:
                num_rows, num_cols = self.儲存資料.shape

                # 設置表格的行數和列數
                self.tableWidget.setRowCount(num_rows)
                self.tableWidget.setColumnCount(num_cols)

                # 設置表頭 (表格第一列的欄位名稱)
                self.tableWidget.setHorizontalHeaderLabels(self.儲存資料.columns)

                # 將數據插入表格中，先迭代行(橫的)，再迭代欄位(直的)
                # 也就是針對每行將各個欄位的資訊填入
                # self.儲存資料.iat[i, j] 是一個用於訪問 Pandas DataFrame 的特定位置的方法。
                # 它的使用方式是 iat[row_index, column_index]，其中 row_index 是行的索引，column_index 是列的索引
                # 取得值後再轉換成字串，填入表格中
                for i in range(num_rows):
                    for j in range(num_cols):
                        item = QTableWidgetItem(str(self.儲存資料.iat[i, j]))
                        self.tableWidget.setItem(i, j, item)

                self.資料比數標籤.setText(f'資料筆數: {len(self.儲存資料)}/{self.tableWidget.rowCount()}')

    def ConcatTable(self, 匯入資料):
        try:
            self.儲存資料 = pd.concat([self.儲存資料, 匯入資料], ignore_index=True)

            num_rows, num_cols = self.儲存資料.shape

            # 設置表格的行數和列數
            self.tableWidget.setRowCount(num_rows)
            self.tableWidget.setColumnCount(num_cols)

            # 設置表頭 (表格第一列的欄位名稱)
            self.tableWidget.setHorizontalHeaderLabels(self.儲存資料.columns)

            # 將數據插入表格中，先迭代行(橫的)，再迭代欄位(直的)
            # 也就是針對每行將各個欄位的資訊填入
            # self.儲存資料.iat[i, j] 是一個用於訪問 Pandas DataFrame 的特定位置的方法。
            # 它的使用方式是 iat[row_index, column_index]，其中 row_index 是行的索引，column_index 是列的索引
            # 取得值後再轉換成字串，填入表格中
            for i in range(num_rows):
                for j in range(num_cols):
                    item = QTableWidgetItem(str(self.儲存資料.iat[i, j]))
                    self.tableWidget.setItem(i, j, item)

            self.資料比數標籤.setText(f'資料筆數: {len(self.儲存資料)}/{self.tableWidget.rowCount()}')

        except:
            error_message = traceback.format_exc()
            QMessageBox.warning(self, "資料匯入錯誤", f"錯誤 : {error_message} \n儲存資料: {self.儲存資料}")

    def UpdateTable(self):
        # 請注意任何有關表格的操作，都必須注意是否連動到Data
        # 否則只清空表格，不會影響到self.儲存資料內的資料，可能造成操作上的問題
        try:
            num_rows, num_cols = self.儲存資料.shape

            # 設置表格的行數和列數
            self.tableWidget.setRowCount(num_rows)
            self.tableWidget.setColumnCount(num_cols)

            # 設置表頭 (表格第一列的欄位名稱)
            self.tableWidget.setHorizontalHeaderLabels(self.儲存資料.columns)

            # 將數據插入表格中，先迭代行(橫的)，再迭代欄位(直的)
            # 也就是針對每行將各個欄位的資訊填入
            # self.儲存資料.iat[i, j] 是一個用於訪問 Pandas DataFrame 的特定位置的方法。
            # 它的使用方式是 iat[row_index, column_index]，其中 row_index 是行的索引，column_index 是列的索引
            # 取得值後再轉換成字串，填入表格中
            for i in range(num_rows):
                for j in range(num_cols):
                    item = QTableWidgetItem(str(self.儲存資料.iat[i, j]))
                    self.tableWidget.setItem(i, j, item)
        except:
            error_message = traceback.format_exc()
            QMessageBox.warning(self, "資料匯入錯誤", f"錯誤 : {error_message} \n儲存資料: {self.儲存資料}")

    def filterTable(self):
        # .strip() 默認為去掉開頭和末尾的空格
        # 取得文字框內容
        搜索子串 = self.搜索框.text().strip()

        # 迭代所有表格內容
        for row in range(self.tableWidget.rowCount()):
            # (row, 0) 分別代表行索引(橫的)和列索引(直的)
            母工單單號 = self.tableWidget.item(row, 0).text()
            工號 = self.tableWidget.item(row, 1).text()

            # 如果搜索字串出現在表格內就不隱藏，否則隱藏
            if 搜索子串 in 母工單單號 or 搜索子串 in 工號:
                self.tableWidget.setRowHidden(row, False)
            else:
                self.tableWidget.setRowHidden(row, True)

    def deleteALL(self):
        try:
            confirm = QtWidgets.QMessageBox.question(self, "清空確認", "即將清除所有項目，確定要這麼做嗎?",
                                                     QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if confirm == QtWidgets.QMessageBox.Yes:
                # 清空DataFrame和表格資料
                self.tableWidget.setRowCount(0)
                self.儲存資料 = self.儲存資料.iloc[0:0]
                # 顯示剩餘筆數
                self.資料比數標籤.setText(f'資料筆數: {len(self.儲存資料)}/{self.tableWidget.rowCount()}')
        except:
            error_message = traceback.format_exc()
            QMessageBox.warning(self, "清除功能錯誤", f"錯誤 : {error_message}")

    def delete_selected_rows(self):
        try:
            選擇項目 = set()
            for item in self.tableWidget.selectedItems():
                選擇項目.add(item.row())

            # 刪除所選的行
            選擇項目 = list(選擇項目)
            self.儲存資料 = self.儲存資料.drop(選擇項目)
            self.儲存資料 = self.儲存資料.reset_index(drop=True)
            self.tableWidget.removeRow(選擇項目[0])

            # 顯示剩餘筆數
            self.資料比數標籤.setText(f'資料筆數: {len(self.儲存資料)}/{self.tableWidget.rowCount()}')

        except:
            error_message = traceback.format_exc()
            QMessageBox.warning(self, "刪除功能錯誤", f"錯誤 : {error_message}")


# 這是 Python 中的慣用語法，表示如果這個程式碼是直接被執行而不是被當作模組引入，則執行下面的程式碼塊
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = 主介面()
    window.show()
    sys.exit(app.exec_())
