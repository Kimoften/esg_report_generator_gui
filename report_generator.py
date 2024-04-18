from openai import OpenAI
import pdfplumber
import re
import fitz
import json
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, Listbox, Toplevel
import threading
import json
import time
import sys
from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt, Signal, QThread)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QLabel, QMainWindow, QMenuBar, QTextEdit,
    QPushButton, QSizePolicy, QStatusBar, QWidget, QInputDialog, QMessageBox, QFileDialog, QPlainTextEdit, QFrame, QProgressBar, QListWidgetItem, QScrollArea, QRadioButton, QCheckBox, QLineEdit, QTabWidget)
# import logo1_rc

class WorkerThread(QThread):
    finished = Signal(object)

    def __init__(self, func, args=None, parent=None):
        super().__init__(parent)
        self.func = func
        self.args = args if args is not None else ()

    def run(self):
        result = self.func(*self.args)
        self.finished.emit(result)
    
    # def __del__(self):
    #     self.wait()  # 쓰레드가 종료될 때까지 대기하도록 함

class GRIApp(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"GRI Draft Generator")
        MainWindow.resize(1027, 693)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.label = QLabel(self.centralwidget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(-10, 0, 1051, 631))
        self.label.setStyleSheet(u"background-color:rgb(255, 217, 102)")
        self.label_pic = QLabel(self.centralwidget)
        self.label_pic.setObjectName(u"label_pic")
        self.label_pic.setGeometry(QRect(420, 220, 241, 91))
        self.pixmap = QPixmap()
        self.pixmap.load("logo1.png")
        self.label_pic.setPixmap(self.pixmap)
        self.label_2 = QLabel(self.centralwidget)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(350, 320, 351, 61))
        font = QFont()
        font.setPointSize(20)
        font.setBold(True)
        self.label_2.setFont(font)
        self.label_2.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.pushButton = QPushButton(self.centralwidget)
        self.pushButton.clicked.connect(self.select_pdf)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(200, 400, 671, 91))
        font1 = QFont()
        font1.setPointSize(15)
        self.pushButton.setFont(font1)
        self.pushButton.setStyleSheet(u"")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 1027, 33))
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.pdf_path = None
        self.raw_data = None
        self.index_list = []
        self.key = None  # key 속성을 추가합니다.      
        self.request_key()  #
        self.retranslateUi(MainWindow)


        QMetaObject.connectSlotsByName(MainWindow)

    def show_loading(self):
        self.loading_window = QMainWindow()
        self.loading_window.setWindowTitle("Loading")
        self.loading_window.setGeometry(0, 0, 200, 100)

        self.loading_label = QLabel("Loading, please wait...", self.loading_window)
        self.loading_label.setGeometry(0, 0, 200, 50)

        self.progress_bar = QProgressBar(self.loading_window)
        self.progress_bar.setGeometry(0, 50, 200, 50)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress

        self.start_time = time.time()

        self.loading_window.show()

    def update_loading_label(self):
        if self.loading_window.isVisible():
            elapsed_time = int(time.time() - self.start_time)
            self.loading_label.setText(f"Loading, please wait... {elapsed_time} seconds")

    def hide_loading(self):
        if self.loading_window is not None:
            self.loading_window.close()
    
    # def run_async(self, func, *args, callback=None):
    #     def run():
    #         self.show_loading()
    #         result = func(*args)
    #         self.hide_loading()
    #         if callback:
    #             callback(result)
            
    #     thread = WorkerThread(func=run)
    #     thread.finished.connect(callback)
    #     thread.start()

    def request_key(self):
    # 사용자로부터 키를 입력받는 메서드입니다.
        key, ok_pressed = QInputDialog.getText(None, "Input", "Please enter the key:")
        if ok_pressed:
            self.key = key
            if not self.key:
                QMessageBox.warning(self, "Warning", "The key is required to proceed.")
                self.request_key()  # 유효한 키를 받을 때까지 재요청합니다

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"GRI Draft Generator", None))
        self.label.setText("")
        self.label_pic.setText("")
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"\uc804\uae30 \ubcf4\uace0\uc11c\ub97c \ucca8\ubd80\ud574\uc8fc\uc138\uc694", None))
        self.pushButton.setText(QCoreApplication.translate("MainWindow", u"PDF \ud30c\uc77c\uc744 \ucca8\ubd80\ud574\uc8fc\uc138\uc694", None))

    def select_pdf(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        self.pdf_path, _ = QFileDialog.getOpenFileName(self.centralwidget, "Select PDF", "", "PDF files (*.pdf)", options=options)
        if self.pdf_path:
            QMessageBox.information(self.centralwidget, "PDF Selected", f"Selected PDF: {self.pdf_path}")
            self.prompt_for_raw_data()  # 파일 선택 후 텍스트 입력창 호출
            
    def prompt_for_raw_data(self):
        MainWindow.close()

        self.input_window = QWidget()
        self.input_window.setWindowTitle("rawdata_input")
        self.input_window.resize(1059, 664)

        # 배경색 설정
        self.input_window.setStyleSheet("background-color: rgb(255, 217, 102);")

        # 라벨 생성
        self.label_2 = QLabel("GS 건설", self.input_window)
        self.label_2.setGeometry(130, 20, 91, 31)
        font = QFont()
        font.setPointSize(15)
        font.setBold(True)
        self.label_2.setFont(font)

        self.label_pic_2 = QLabel(self.input_window)
        self.label_pic_2.setObjectName(u"label_pic")
        self.label_pic_2.setGeometry(QRect(10, 10, 111, 41))
        self.label_pic_2.setPixmap(self.pixmap)
        self.label_pic_2.setScaledContents(True)

        # 선 생성
        self.line = QFrame(self.input_window)
        self.line.setObjectName(u"line")
        self.line.setGeometry(QRect(0, 56, 1059, 2))
        self.line.setStyleSheet(u"background-color: rgb(0, 0, 0);")
        self.line.setFrameShape(QFrame.Shape.HLine)
        self.line.setFrameShadow(QFrame.Shadow.Sunken)

        # 텍스트 입력 위젯 생성
        self.plainTextEdit = QPlainTextEdit(self.input_window)
        self.plainTextEdit.setObjectName(u"plainTextEdit")
        self.plainTextEdit.setGeometry(QRect(100, 140, 841, 471))
        self.plainTextEdit.setStyleSheet("background-color: white;")
        self.plainTextEdit.setOverwriteMode(True)
        self.plainTextEdit.setCenterOnScroll(True)
        self.plainTextEdit.setPlaceholderText(u"내용을 입력해주세요")
        font1 = QFont()
        font1.setPointSize(9)  # 폰트 크기 설정
        self.plainTextEdit.setFont(font1)  # 설정한 폰트 적용

        # 라벨 생성
        self.label_3 = QLabel("클라이언트 인터뷰, 고객사 성과자료 등 Raw Data를 입력해주세요!", self.input_window)
        self.label_3.setGeometry(QRect(100, 90, 641, 41))
        self.label_3.setFont(font)

        # 확인 버튼 생성
        ok_button = QPushButton("다음", self.input_window)
        ok_button.setGeometry(500, 630, 75, 23)
        ok_button.setStyleSheet(u"background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);")
        ok_button.clicked.connect(self.process_raw_data)


        self.input_window.show()
    
    def process_raw_data(self):
    # 입력된 텍스트 가져오기
        self.raw_data = self.plainTextEdit.toPlainText().strip()
        if self.raw_data:
            # 스레드 시작
            self.show_loading()
            self.thread = WorkerThread(func=self.get_index_and_titles)
            self.thread.finished.connect(self.show_items)
            self.thread.start()
        else:
            QMessageBox.warning(None, "Warning", "Please enter some text.")

    def get_index_and_titles(self):
        # 데이터 처리 작업
        self.index_list = Show_indexList(self.raw_data, self.key)
        titles = get_GRI_Title(self.index_list)
        self.combined_list = [f"({self.index_list[i]['disclosure_num']}): [{title}] - {self.index_list[i]['description']}" for i, title in enumerate(titles)]
        self.disclosure_num_list = [item["disclosure_num"] for item in self.index_list]
    

    
    def show_items(self):
        # 이전에 추가된 위젯을 모두 제거합니다.
        self.input_window.close()
        self.hide_loading()
        # self.input_window.close()
        
        self.list_window = QWidget()
        self.list_window.setWindowTitle("index_select")
        self.list_window.resize(1059, 664)

        # 배경색 설정
        self.list_window.setStyleSheet("background-color: rgb(255, 217, 102);")

        # 라벨 생성
        self.label_4 = QLabel("GS 건설", self.list_window)
        self.label_4.setGeometry(130, 20, 91, 31)
        font = QFont()
        font.setPointSize(15)
        font.setBold(True)
        self.label_4.setFont(font)

        self.label_5 = QLabel(self.list_window)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setGeometry(QRect(100, 140, 841, 441))
        self.label_5.setStyleSheet(u"background-color: rgb(255, 255, 255);")

        self.label_pic_2 = QLabel(self.list_window)
        self.label_pic_2.setObjectName(u"label_pic")
        self.label_pic_2.setGeometry(QRect(10, 10, 111, 41))
        self.label_pic_2.setPixmap(self.pixmap)
        self.label_pic_2.setScaledContents(True)

        # 선 생성
        self.line = QFrame(self.list_window)
        self.line.setObjectName(u"line")
        self.line.setGeometry(QRect(0, 56, 1059, 2))
        self.line.setStyleSheet(u"background-color: rgb(0, 0, 0);")
        self.line.setFrameShape(QFrame.Shape.HLine)
        self.line.setFrameShadow(QFrame.Shadow.Sunken)


        self.gri1 = QScrollArea(self.list_window)
        self.gri1.setObjectName(u"gri1")
        self.gri1.setGeometry(QRect(170, 170, 751, 61))
        self.gri1.setStyleSheet(u"border-color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);")
        self.gri1.setWidgetResizable(True)
        self.scrollAreaWidgetContents1 = QWidget()
        self.scrollAreaWidgetContents1.setObjectName(u"gri1_content")
        self.scrollAreaWidgetContents1.setGeometry(QRect(0, 0, 749, 59))
        self.gri1.setWidget(self.scrollAreaWidgetContents1)
        text_label = QLabel(self.combined_list[0], self.scrollAreaWidgetContents1)
        text_label.setGeometry(0, 0, 751, 61)
        text_label.setWordWrap(True)

        self.gri2 = QScrollArea(self.list_window)
        self.gri2.setObjectName(u"gri2")
        self.gri2.setGeometry(QRect(170, 250, 751, 61))
        self.gri2.setStyleSheet(u"border-color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);")
        self.gri2.setWidgetResizable(True)
        self.scrollAreaWidgetContents2 = QWidget()
        self.scrollAreaWidgetContents2.setObjectName(u"gri2_content")
        self.scrollAreaWidgetContents2.setGeometry(QRect(0, 0, 749, 59))
        self.gri2.setWidget(self.scrollAreaWidgetContents2)
        text_label2 = QLabel(self.combined_list[1], self.scrollAreaWidgetContents2)
        text_label2.setGeometry(0, 0, 751, 61)
        text_label2.setWordWrap(True)

        self.gri3 = QScrollArea(self.list_window)
        self.gri3.setObjectName(u"gri3")
        self.gri3.setGeometry(QRect(170, 330, 751, 61))
        self.gri3.setStyleSheet(u"border-color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);")
        self.gri3.setWidgetResizable(True)
        self.scrollAreaWidgetContents3 = QWidget()
        self.scrollAreaWidgetContents3.setObjectName(u"gri3_content")
        self.scrollAreaWidgetContents3.setGeometry(QRect(0, 0, 749, 59))
        self.gri3.setWidget(self.scrollAreaWidgetContents3)
        text_label3 = QLabel(self.combined_list[2], self.scrollAreaWidgetContents3)
        text_label3.setGeometry(0, 0, 751, 61)
        text_label3.setWordWrap(True)

        self.gri4 = QScrollArea(self.list_window)
        self.gri4.setObjectName(u"gri4")
        self.gri4.setGeometry(QRect(170, 410, 751, 61))
        self.gri4.setStyleSheet(u"border-color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);")
        self.gri4.setWidgetResizable(True)
        self.scrollAreaWidgetContents4 = QWidget()
        self.scrollAreaWidgetContents4.setObjectName(u"gri4_content")
        self.scrollAreaWidgetContents4.setGeometry(QRect(0, 0, 749, 59))
        self.gri4.setWidget(self.scrollAreaWidgetContents4)
        text_label4 = QLabel(self.combined_list[3], self.scrollAreaWidgetContents4)
        text_label4.setGeometry(0, 0, 751, 61)
        text_label4.setWordWrap(True)

        self.gri5 = QScrollArea(self.list_window)
        self.gri5.setObjectName(u"gri5")
        self.gri5.setGeometry(QRect(170, 490, 751, 61))
        self.gri5.setStyleSheet(u"border-color: rgb(0, 0, 0); background-color: rgb(255, 255, 255);")
        self.gri5.setWidgetResizable(True)
        self.scrollAreaWidgetContents5 = QWidget()
        self.scrollAreaWidgetContents5.setObjectName(u"gri5_content")
        self.scrollAreaWidgetContents5.setGeometry(QRect(0, 0, 749, 59))
        self.gri5.setWidget(self.scrollAreaWidgetContents5)
        text_label5 = QLabel(self.combined_list[4], self.scrollAreaWidgetContents5)
        text_label5.setGeometry(0, 0, 751, 61)
        text_label5.setWordWrap(True)

        self.gricheck1 = QCheckBox(self.list_window)
        self.gricheck1.setObjectName(u"gricheck1")
        self.gricheck1.setStyleSheet(u"background-color: rgb(255, 255, 255);")
        self.gricheck1.setGeometry(QRect(130, 190, 16, 20))
        
        self.gricheck2 = QCheckBox(self.list_window)
        self.gricheck2.setObjectName(u"gricheck2")
        self.gricheck2.setStyleSheet(u"background-color: rgb(255, 255, 255);")
        self.gricheck2.setGeometry(QRect(130, 270, 16, 20))
        
        self.gricheck3 = QCheckBox(self.list_window)
        self.gricheck3.setObjectName(u"gricheck3")
        self.gricheck3.setStyleSheet(u"background-color: rgb(255, 255, 255);")
        self.gricheck3.setGeometry(QRect(130, 350, 16, 20))

        self.gricheck4 = QCheckBox(self.list_window)
        self.gricheck4.setObjectName(u"gricheck4")
        self.gricheck4.setStyleSheet(u"background-color: rgb(255, 255, 255);")
        self.gricheck4.setGeometry(QRect(130, 430, 16, 20))

        self.gricheck5 = QCheckBox(self.list_window)
        self.gricheck5.setObjectName(u"gricheck5")
        self.gricheck5.setStyleSheet(u"background-color: rgb(255, 255, 255);")
        self.gricheck5.setGeometry(QRect(130, 510, 16, 20))


        # 라벨 생성
        self.label_6 = QLabel("맞춤 GRI Index를 3개 선택해주세요!", self.list_window)
        self.label_6.setGeometry(QRect(100, 90, 641, 41))
        self.label_6.setFont(font)

        # 확인 버튼 생성
        ok_button = QPushButton("다음", self.list_window)
        ok_button.setGeometry(500, 630, 75, 23)
        ok_button.setStyleSheet(u"background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);")
        ok_button.clicked.connect(self.get_checked_items)


        self.list_window.show()
    
    def get_checked_items(self):
        self.checked_items = []
        self.checked_disclosure_nums = []
        if self.gricheck1.isChecked():
            self.checked_items.append(0)
            self.checked_disclosure_nums.append(self.disclosure_num_list[0])
        if self.gricheck2.isChecked():
            self.checked_items.append(1)
            self.checked_disclosure_nums.append(self.disclosure_num_list[1])
        if self.gricheck3.isChecked():
            self.checked_items.append(2)
            self.checked_disclosure_nums.append(self.disclosure_num_list[2])
        if self.gricheck4.isChecked():
            self.checked_items.append(3)
            self.checked_disclosure_nums.append(self.disclosure_num_list[3])
        if self.gricheck5.isChecked():
            self.checked_items.append(4)
            self.checked_disclosure_nums.append(self.disclosure_num_list[4])
        
        if len(self.checked_items) == 3:
            self.show_loading()
            self.thread = WorkerThread(func=self.extract_text)
            self.thread.finished.connect(self.edit_text)
            self.thread.start()

        else:
            QMessageBox.critical(self.list_window, "Error", "정확하게 3개 골라주세요.")
        
        

    
    def extract_text(self):
        self.extracted_text=[]
        for number in self.checked_items:
            disclosure_num = self.index_list[number]['disclosure_num']
            pages = find_gri_pages(self.pdf_path, disclosure_num)
            if type(pages) == list:
                self.extracted_text += extract_text_from_pages(self.pdf_path, pages)
            else:
                self.extracted_text += ["no page in previous report"]
        
        # print(self.extracted_text)
        return self.extracted_text



    def edit_text(self):
        self.hide_loading()
        self.list_window.close()
        self.edit_window = QWidget()
        self.edit_window.setWindowTitle("edit_text")
        self.edit_window.resize(1059, 664)

        # 배경색 설정
        self.edit_window.setStyleSheet("background-color: rgb(255, 217, 102);")

        # 라벨 생성
        self.label_7 = QLabel("GS 건설", self.edit_window)
        self.label_7.setGeometry(130, 20, 91, 31)
        font = QFont()
        font.setPointSize(15)
        font.setBold(True)
        self.label_7.setFont(font)

        self.label_8 = QLabel(self.edit_window)
        self.label_8.setObjectName(u"label_5")
        self.label_8.setGeometry(QRect(100, 140, 841, 441))
        self.label_8.setStyleSheet(u"background-color: rgb(255, 255, 255);")

        self.label_pic_2 = QLabel(self.edit_window)
        self.label_pic_2.setObjectName(u"label_pic")
        self.label_pic_2.setGeometry(QRect(10, 10, 111, 41))
        self.label_pic_2.setPixmap(self.pixmap)
        self.label_pic_2.setScaledContents(True)

        # 선 생성
        self.line = QFrame(self.edit_window)
        self.line.setObjectName(u"line")
        self.line.setGeometry(QRect(0, 56, 1059, 2))
        self.line.setStyleSheet(u"background-color: rgb(0, 0, 0);")
        self.line.setFrameShape(QFrame.Shape.HLine)
        self.line.setFrameShadow(QFrame.Shadow.Sunken)

        self.label_9 = QLabel("GRI Index에 해당하는 전기 보고서 내용을 검수해주세요!", self.edit_window)
        self.label_9.setGeometry(QRect(100, 90, 641, 41))
        self.label_9.setFont(font)
        
        self.tabWidget = QTabWidget(self.edit_window)
        self.tabWidget.setObjectName(u"tabWidget")
        self.tabWidget.setGeometry(QRect(100, 140, 841, 471))
        
        self.tab_1 = QPlainTextEdit(self.extracted_text[0], self.edit_window)
        self.tab_1.setObjectName(u"tab_1")
        self.tab_1.setGeometry(QRect(100, 140, 841, 471))
        self.tab_1.setStyleSheet("background-color: white;")
        font1 = QFont()
        font1.setPointSize(9)  # 폰트 크기 설정
        self.tab_1.setFont(font1)
        self.tabWidget.addTab(self.tab_1, self.checked_disclosure_nums[0])

        self.tab_2 = QPlainTextEdit(self.extracted_text[1], self.edit_window)
        self.tab_2.setObjectName(u"tab_2")
        self.tab_2.setGeometry(QRect(100, 140, 841, 471))
        self.tab_2.setStyleSheet("background-color: white;")
        self.tab_2.setFont(font1)
        self.tabWidget.addTab(self.tab_2, self.checked_disclosure_nums[1])

        self.tab_3 = QPlainTextEdit(self.extracted_text[2], self.edit_window)
        self.tab_3.setObjectName(u"tab_3")
        self.tab_3.setGeometry(QRect(100, 140, 841, 471))
        self.tab_3.setStyleSheet("background-color: white;")
        self.tab_3.setFont(font1)
        self.tabWidget.addTab(self.tab_3, self.checked_disclosure_nums[2])


        submit_button = QPushButton("초안 요청", self.edit_window)
        submit_button.setGeometry(500, 630, 75, 23)
        submit_button.setStyleSheet(u"background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);")
        submit_button.clicked.connect(self.process_edit_text)

        self.edit_window.show()
    
    def process_edit_text(self):
    # 입력된 텍스트 가져오기
        self.text_data1 = self.tab_1.toPlainText().strip()
        self.text_data2 = self.tab_2.toPlainText().strip()
        self.text_data3 = self.tab_3.toPlainText().strip()

        if self.text_data1 and self.text_data2 and self.text_data3:
            # 스레드 시작
            self.show_loading()
            self.thread = WorkerThread(func=self.generate_draft)
            self.thread.finished.connect(self.show_draft)
            self.thread.start()
        else:
            QMessageBox.warning(None, "Warning", "Please enter some text.")
    
    def generate_draft(self):
        self.draft1 = get_draft(self.text_data1, self.checked_disclosure_nums[0], self.raw_data, self.key)
        self.draft2 = get_draft(self.text_data2, self.checked_disclosure_nums[1], self.raw_data, self.key)
        self.draft3 = get_draft(self.text_data3, self.checked_disclosure_nums[2], self.raw_data, self.key)



    def show_draft(self):
        self.hide_loading()
        self.edit_window.close()
        self.result_window = QWidget()
        self.result_window.setWindowTitle("edit_text")
        self.result_window.resize(1059, 664)

        # 배경색 설정
        self.result_window.setStyleSheet("background-color: rgb(255, 217, 102);")

        # 라벨 생성
        self.label_10 = QLabel("GS 건설", self.result_window)
        self.label_10.setGeometry(130, 20, 91, 31)
        font = QFont()
        font.setPointSize(15)
        font.setBold(True)
        self.label_10.setFont(font)

        self.label_11 = QLabel(self.result_window)
        self.label_11.setObjectName(u"label_11")
        self.label_11.setGeometry(QRect(100, 140, 841, 441))
        self.label_11.setStyleSheet(u"background-color: rgb(255, 255, 255);")

        self.label_pic_2 = QLabel(self.result_window)
        self.label_pic_2.setObjectName(u"label_pic")
        self.label_pic_2.setGeometry(QRect(10, 10, 111, 41))
        self.label_pic_2.setPixmap(self.pixmap)
        self.label_pic_2.setScaledContents(True)

        # 선 생성
        self.line = QFrame(self.result_window)
        self.line.setObjectName(u"line")
        self.line.setGeometry(QRect(0, 56, 1059, 2))
        self.line.setStyleSheet(u"background-color: rgb(0, 0, 0);")
        self.line.setFrameShape(QFrame.Shape.HLine)
        self.line.setFrameShadow(QFrame.Shadow.Sunken)

        self.label_12 = QLabel("ESG 보고서 초안", self.edit_window)
        self.label_12.setGeometry(QRect(100, 90, 641, 41))
        self.label_12.setFont(font)
        
        self.tabWidget2 = QTabWidget(self.result_window)
        self.tabWidget2.setObjectName(u"tabWidget")
        self.tabWidget2.setGeometry(QRect(100, 140, 841, 471))
        
        self.tab_6 = QTextEdit(str(self.draft1), self.result_window)
        self.tab_6.setObjectName(u"tab_6")
        self.tab_6.setGeometry(QRect(100, 140, 841, 471))
        self.tab_6.setStyleSheet("background-color: white;")
        font1 = QFont()
        font1.setPointSize(9)  # 폰트 크기 설정
        self.tab_6.setFont(font1)
        self.tab_6.setReadOnly(True)
        self.tabWidget2.addTab(self.tab_6, self.checked_disclosure_nums[0])

        self.tab_7 = QTextEdit(str(self.draft2), self.result_window)
        self.tab_7.setObjectName(u"tab_7")
        self.tab_7.setGeometry(QRect(100, 140, 841, 471))
        self.tab_7.setStyleSheet("background-color: white;")
        self.tab_7.setFont(font1)
        self.tab_7.setReadOnly(True)
        self.tabWidget2.addTab(self.tab_7, self.checked_disclosure_nums[1])

        self.tab_8 = QTextEdit(str(self.draft3), self.result_window)
        self.tab_8.setObjectName(u"tab_8")
        self.tab_8.setGeometry(QRect(100, 140, 841, 471))
        self.tab_8.setStyleSheet("background-color: white;")
        self.tab_8.setFont(font1)
        self.tab_8.setReadOnly(True)
        self.tabWidget2.addTab(self.tab_8, self.checked_disclosure_nums[2])

        # submit_button = QPushButton("다시하기", self.result_window)
        # submit_button.setGeometry(500, 630, 75, 23)
        # submit_button.setStyleSheet(u"background-color: rgb(0, 0, 0); color: rgb(255, 255, 255);")
        # submit_button.clicked.connect(self.restart_application)

        self.result_window.show()
    
#     def restart_application(self):
#         close_current_application()
        
#         # MainWindow = QMainWindow()
#         # ui = GRIApp()
#         # ui.setupUi(MainWindow)
#         # MainWindow.show()
#         # sys.exit(app.exec())

# def close_current_application():
#     QApplication.instance().quit()  # 현재 실행 중인 프로세스 종료
#         # 새로운 프로세스 시작
#     new_process = QApplication(sys.argv)
#     main_window = QMainWindow()
#     new_ui = GRIApp()
#     new_ui.setupUi(main_window)
#     main_window.show()
#     sys.exit(new_process.exec_())
    


def Show_indexList(raw_data, key):
    index_list = get_index(raw_data, key)
    return index_list

def Create_Draft(raw_data, index_list, selected_numbers, pdf_path, key):
    draft = []
    for number in selected_numbers:
        disclosure_num = index_list[number]['disclosure_num']
        pages = find_gri_pages(pdf_path, disclosure_num)
        if type(pages) == list:
            extracted_pages = extract_text_from_pages(pdf_path, pages)
        else:
            extracted_pages = ["no page in previous report"]
        for extracted_page in extracted_pages:
            small_draft = [pages, disclosure_num]
            small_draft.append(get_draft(extracted_page, disclosure_num, raw_data,key))
        draft.append(small_draft)
    return draft

def get_GRI_Title(index_list):
    Title_list = []
    for index in index_list:
        gri = index['disclosure_num']
        GRI_title = translate(gri)
        Title_list.append(GRI_title)
    return Title_list

def translate(target_value):
    a = {'GRI Index': '공시제목', 'GRI 2-1': '조직의 세부정보', 'GRI 2-2': '조직의 지속가능성 보고에 포함된 주체', 'GRI 2-3': '보고기간, 빈도 및 연락처', 'GRI 2-4': '정보의 재진술', 'GRI 2-5': '외부검증', 'GRI 2-6': '활동, 가치 사슬 및 기타 비즈니스 관계', 'GRI 2-7': '직원', 'GRI 2-8': '근로자가 아닌 근로자', 'GRI 2-9': '지배구조 및 구성', 'GRI 2-10': '최고 거버넌스 기구의 추천 및 선정', 'GRI 2-11': '최고 거버넌스 기구의 의장', 'GRI 2-12': '영향 관리를 감독하는 최고 거버넌스 기구의 역할', 'GRI 2-13': '영향관리 책임 위임', 'GRI 2-14': '지속가능성 보고에 있어 최고 거버넌스 기구의 역할', 'GRI 2-15': '이해 상충', 'GRI 2-16': '중요한 우려 사항에 대한 의사소통', 'GRI 2-17': '최고 거버넌스 기구의 집단적 지식', 'GRI 2-18': '최고 거버넌스 기구의 성과 평가', 'GRI 2-19': '보상 정책', 'GRI 2-20': '보수 결정 과정', 'GRI 2-21': '연간 총보상비율', 'GRI 2-22': '지속 가능한 개발 전략에 대한 선언문', 'GRI 2-23': '정책 공약', 'GRI 2-24': '정책 공약 내재화', 'GRI 2-25': '부정적인 영향을 해결하기 위한 프로세스', 'GRI 2-26': '조언을 구하고 문제를 제기하는 메커니즘', 'GRI 2-27': '법률 및 규정 준수', 'GRI 2-28': '회원 협회', 'GRI 2-29': '이해관계자 참여에 대한 접근 방식', 'GRI 2-30': '단체 교섭 협약', 'GRI 3-1': '중요 주제를 결정하는 프로세스', 'GRI 3-2': '중요 주제 목록', 'GRI 3-3': '중요 주제 관리', 'GRI 201-1': '직접적인 경제가치 창출 및 배분', 'GRI 201-2': '기후변화로 인한 재정적 영향과 기타 위험과 기회', 'GRI 201-3': '확정급여제도 채무 및 기타 퇴직제도', 'GRI 201-4': '정부로부터 재정 지원을 받았습니다.', 'GRI 202-1': '현지 최저임금 대비 성별 표준 신입사원 임금 비율', 'GRI 202-2': '지역사회에서 채용된 고위 경영진의 비율', 'GRI 203-1': '인프라 투자 및 서비스 지원', 'GRI 203-2': '간접적인 경제적 영향이 크다', 'GRI 204-1': '현지 공급업체에 대한 지출 비율', 'GRI 205-1': '부패와 관련된 위험에 대해 평가된 운영', 'GRI 205-2': '부패 방지 정책 및 절차에 대한 커뮤니케이션 및 교육', 'GRI 205-3': '확인된 비리사례 및 이에 대한 조치', 'GRI 206-1': '반경쟁 행위, 독점 금지, 독점 관행에 대한 법적 조치', 'GRI 207-1': '세금에 대한 접근', 'GRI 207-2': '세금 거버넌스, 통제 및 위험 관리', 'GRI 207-3': '이해관계자 참여 및 조세 관련 우려 사항 관리', 'GRI 207-4': '국가별 보고', 'GRI 301-1': '무게나 부피에 따라 사용되는 재료', 'GRI 301-2': '재활용 투입재 사용', 'GRI 301-3': '재생제품 및 그 포장재', 'GRI 302-1': '조직 내 에너지 소비', 'GRI 302-2': '조직 외부의 에너지 소비', 'GRI 302-3': '에너지 집약도', 'GRI 302-4': '에너지 소비 감소', 'GRI 302-5': '제품 및 서비스의 에너지 요구 사항 감소', 'GRI 303-1': '공유자원으로서 물과의 상호작용', 'GRI 303-2': '배수로 인한 영향 관리', 'GRI 303-3': '물 철수', 'GRI 303-4': '배수', 'GRI 303-5': '물 소비량', 'GRI 304-1': '보호지역 및 보호지역 밖의 생물다양성 가치가 높은 지역 내 또는 인근 지역에서 소유, 임대, 관리되는 운영 현장', 'GRI 304-2': '활동, 제품, 서비스가 생물다양성에 미치는 중대한 영향', 'GRI 304-3': '서식지 보호 또는 복원', 'GRI 304-4': '운영으로 영향을 받는 지역에 서식지가 있는 IUCN 적색 목록 종 및 국가 보존 목록 종', 'GRI 305-1': '직접(Scope 1) 온실가스 배출', 'GRI 305-2': '에너지 간접(Scope 2) 온실가스 배출', 'GRI 305-3': '기타 간접(Scope 3) GHG 배출', 'GRI 305-4': '온실가스 배출 집약도', 'GRI 305-5': '온실가스 배출 감소', 'GRI 305-6': '오존층 파괴 물질(ODS) 배출', 'GRI 305-7': '질소산화물(NOx), 황산화물(SOx) 및 기타 상당한 대기 배출물', 'GRI 306-1': '폐기물 발생 및 심각한 폐기물 관련 영향', 'GRI 306-2': '심각한 폐기물 관련 영향 관리', 'GRI 306-3': '폐기물 발생', 'GRI 306-4': '폐기물이 폐기 대상으로 전환됨', 'GRI 306-5': '폐기 대상 폐기물', 'GRI 308-1': '환경 기준에 따라 심사를 거친 신규 공급업체', 'GRI 308-2': '공급망의 부정적인 환경 영향과 취해진 조치', 'GRI 401-1': '신규 직원 채용 및 직원 이직률', 'GRI 401-2': '임시직, 시간제 직원에게는 제공되지 않고 정규직 직원에게는 제공되는 혜택', 'GRI 401-3': '육아휴직', 'GRI 402-1': '운영 변경에 대한 최소 통지 기간', 'GRI 403-1': '산업 보건 및 안전 관리 시스템', 'GRI 403-2': '위험 식별, 위험 평가 및 사고 조사', 'GRI 403-3': '산업 보건 서비스', 'GRI 403-4': '산업안전보건에 관한 근로자 참여, 상담, 소통', 'GRI 403-5': '산업 보건 및 안전에 관한 근로자 교육', 'GRI 403-6': '근로자 건강증진', 'GRI 403-7': '비즈니스 관계와 직접적으로 연결된 산업 보건 및 안전 영향의 예방 및 완화', 'GRI 403-8': '산업 보건 및 안전 관리 시스템의 적용을 받는 근로자', 'GRI 403-9': '업무 관련 부상', 'GRI 403-10': '업무상 질병', 'GRI 404-1': '직원 1인당 연간 평균 교육 시간', 'GRI 404-2': '직원 기술 향상을 위한 프로그램 및 전환 지원 프로그램', 'GRI 404-3': '정기적인 성과 및 경력 개발 검토를 받는 직원 비율', 'GRI 405-1': '거버넌스 기구와 직원의 다양성', 'GRI 405-2': '남성 대비 여성의 기본급 및 보수 비율', 'GRI 406-1': '차별 사건 및 이에 대한 시정조치', 'GRI 407-1': '결사 및 단체교섭의 자유가 침해될 수 있는 사업장 및 공급업체', 'GRI 408-1': '아동 노동 발생 위험이 높은 사업장 및 공급업체', 'GRI 409-1': '강제 노동 발생 위험이 높은 사업장 및 공급업체', 'GRI 410-1': '인권 정책이나 절차에 대한 교육을 받은 보안 인력', 'GRI 411-1': '원주민 권리 침해 사건', 'GRI 413-1': '지역사회 참여, 영향 평가, 개발 프로그램을 통한 운영', 'GRI 413-2': '지역사회에 실질적이고 잠재적으로 중대한 부정적 영향을 미치는 사업장', 'GRI 414-1': '사회적 기준을 통해 심사를 거친 신규 공급업체', 'GRI 414-2': '공급망 내 부정적인 사회적 영향과 이에 대한 조치', 'GRI 415-1': '정치적 기부', 'GRI 416-1': '제품 및 서비스 카테고리의 건강 및 안전 영향 평가', 'GRI 416-2': '제품 및 서비스의 건강 및 안전 영향에 관한 규정 위반 사건', 'GRI 417-1': '제품 및 서비스 정보와 라벨링에 대한 요구사항', 'GRI 417-2': '제품, 서비스 정보 및 라벨링과 관련된 규정 위반 사건', 'GRI 417-3': '마케팅 커뮤니케이션 관련 규정 위반 사건', 'GRI 418-1': '고객개인정보보호 위반 및 고객정보 분실 사실이 입증된 불만사항', None: None}
    if a.get(target_value):
        return a.get(target_value)
    # 찾지 못한 경우 None 반환
    return None

def get_index(raw_data,key):
    client = OpenAI(api_key=key)
    response = client.chat.completions.create(
    model="gpt-4-turbo-preview",
    messages=[
        {"role": "system", "content": """ You are a GRI Standards expert. Based on your best understanding of all indicators within the GRI Standards, please identify five GRI Standards indexes to which ***company interview content*** applies. 

[form of answer].
For each GRI Index, briefly state the key content of the GRI Index in **one sentence**, followed by the key reason for presenting the GRI Index in **Max three sentences**.

[Notes]
1) Do not include **GRI Indexes numbered in the 100s** in your answer as they are not applicable.
2) For more specific application, treat all Index numbers like example below and address the detailed guideline content.
3) Elaborate on delivering aspect of Changing degree and new creation of before & after.
@@example
GRI 403-3: 직업 건강 및 안전 관리 시스템
- 조직의 직업 건강 및 안전 관리 시스템에 대한 정보를 제공합니다.
- 안전 조회 번역 웹 구축은 직원들의 안전과 건강을 위한 정보 접근성을 높이는 조치로, 직업 건강 및 안전 관리 시스템의 일환으로 볼 수 있습니다.
@@

결과물은 다음 형식으로 반드시 나와야해:
{"disclosure_num": "GRI {Number-number}", "description": description }, {"disclosure_num": "GRI {Number-number}", "description": description },...]

Try your best and **Answer in Korean**.
    """},
        {"role": "user", "content": f"raw_data:{raw_data}"}
    ],
    temperature=0,
    top_p=0
    )
    return json.loads(response.choices[0].message.content)

def get_draft(extracted_text,index,raw_data,key):
    client = OpenAI(api_key=key)
    response = client.chat.completions.create(
    model="gpt-4",
    messages=[
        {"role": "system", "content": """
        @@As an ESG report consultant, you should prepare draft of ***Raw Data-GRI Index Matching*** updates to the ***Previous Report Content***.


@@Updating Task Description
Updating Task's Main Goal is delivering aspect of Changing Degree and New Creation in the comparison of before & after of Company's GRI Index related behavior.
Therefore, you'd better focus on this main goal while making draft of ***Raw Data-GRI Index Matching*** updates to the ***Previous Report Content***.


@@While drafting, do's and don'ts are listed below. 
[Do's]
- Properly Include Quantified Contents 
- Contents Based on verifiable, objective metrics
- What the company is doing specifically to address the issue in the relevant GRI Index-related activity, how much and how it plans to do so in the future
- elaborate more on actions and add detailed contents.
- Properly Add subtitle that summarize the key action of the paragraph

[Don'ts]
- Vague Contents
- Contents Based on qualitative indicators
- Rhetorical words such as hard, committed, strategically, company-wide, etc.
- Including specific GRI Index number in the answer


@@Also, please refer to the example of a good and bad update draft below for your drafting assignment.
(Good update draft)
"Hilton Hotels is committed to doubling our investment in social impact and halving our environmental footprint by 2030 through responsible hospitality across our value chain. Doubling our investment in social impact means contributing to the development of the communities in which our hotels are located, which we do by actively collaborating with local businesses, purchasing from local companies, actively hiring local residents and providing vocational training to local students and young adults. To make this happen, we have increased our budget by more than 10 per cent each year for the past five years since 2016, and will spend more than twice as much by 2030 as we did in 2016. (middle) With the goal of reducing our environmental footprint by more than half in 2030 compared to 2016, we have increased our share of renewable energy by more than 5 per cent each year for the past five years, and reduced water use by more than 5 per cent per hotel each year. At this rate, we will be on track to achieve our goal by 2030." 

(Bad update draft)
"OOOO is committed to implementing ESG management by pooling efforts across the organisation. We have established an ESG Committee under the Board of Directors, a dedicated ESG team, and are conducting ESG management training for all employees to internalise ESG management. (Medium) In the environmental area, we are working to reduce greenhouse gas emissions and minimise waste across our business value chain, and in the social area, we are working to strengthen human and labour rights, fair operations, and consumer protection."


@@Please only provide updated drafts in your response, as well as your suggestions for narrative and photographic ideas that convey the context and background of each draft.
[Answer Format]
**Draft**
**narrative and photographic ideas**



@@Please do your best to complete the draft ESG report. You should *Answer in Korean*.
        """},
        {"role": "user", "content": f''' 
         ***Raw Data-GRI Index Matching*** \n
{index}: {raw_data}\n\n
         ***Previous Report Content***\n
{extracted_text}
'''}
    ]
    )
    return response.choices[0].message.content

def extract_text_from_pages(pdf_path, page_numbers):
    """
    PDF 파일에서 지정된 여러 페이지들의 텍스트를 추출합니다.

    :param pdf_path: PDF 파일의 경로
    :param page_numbers: 텍스트를 추출할 페이지 번호들의 리스트 (0부터 시작)
    :return: 각 페이지의 텍스트를 담은 리스트
    """
    # PDF 파일 열기
    doc = fitz.open(pdf_path)
    
    # 추출된 텍스트를 저장할 리스트
    pages_text = []
    
    for page_number in page_numbers:
        # 각 페이지의 텍스트 추출하고 리스트에 추가
        page_text = doc[page_number-1].get_text()
        pages_text.append(page_text)
    
    # PDF 문서 닫기
    doc.close()
    
    return pages_text

def find_gri_pages(pdf_path, gri_value):
    
    pages_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            pages_text += page.extract_text() + "\n"
    
    pattern = rf"{gri_value}.*?((?:(?!1|2|3\b)\d+(?:(?:~|, )\d+)*)|해당사항 없음)"
    match = re.search(pattern, pages_text, re.DOTALL)

    if match:
        matched_text = match.group(1)
        if "해당사항 없음" in matched_text:
            return "해당사항 없음"
        else:
            parts = matched_text.split(',')
            numbers = []

            for part in parts:
                # Trim spaces and check if the part represents a range
                part = part.strip()
                if '~' in part:
                    # If part is a range, split it by '~' and convert both ends to integers
                    start, end = map(int, part.split('~'))
                    # Add all numbers in the range to the list, including both ends
                    numbers.extend(range(start, end + 1))
                else:
                    # If part is a single number, convert it to integer and add to the list
                    numbers.append(int(part))

            return numbers
    else:
        return "GRI 값에 해당하는 페이지를 찾을 수 없습니다."
        
    input_string = page_str(pdf_path, gri_value)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    MainWindow = QMainWindow()
    ui = GRIApp()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())