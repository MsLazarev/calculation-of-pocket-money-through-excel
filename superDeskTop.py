from openpyxl import *
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import * 
#from DeskTopMom import *
#from DeskTopDad import *
from time import *

wb = reader.excel.load_workbook(filename="Дань.xlsx", data_only=True)
app = QApplication([])
sheet = wb.active
loc = localtime()
year_now = loc.tm_year
month_now = loc.tm_mon
day_now = loc.tm_mday
alphavit = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
first_title = "Выбор подсчета"
win1_width, win1_height = 800, 600
class FirstWin(QWidget):
    def __init__(self):
        super().__init__()
        self.set_appear()
        self.initUI()
        self.connects()
        self.show()
    def excel(self):
        sheet[coordinats] = self.money1.text()
        wb.save("Дань.xlsx")
        #self.hide()
        #exit()
    def set_appear(self):
        self.setWindowTitle(first_title)
        self.resize(win1_width, win1_height)
    def initUI(self):
        self.choose_question = QLabel("Сколько заплатил плательщик в указанную дату?")
        self.data = QLabel(date[:10])
        self.mom = QLabel("Мама")
        self.dad = QLabel("Папа")
        self.money1 = QLineEdit("Введите сумму")
        self.accept = QPushButton("Сохранить")
        self.choose_question.setFont(QFont('Arial', 15))
        self.data.setFont(QFont('Arial', 15))
        self.mom.setFont(QFont('Arial', 15))
        self.dad.setFont(QFont('Arial', 15))
        self.money1.setFont(QFont('Arial', 15))
        self.accept.setFont(QFont('Arial', 15))
        self.v_line = QVBoxLayout()
        self.h_line = QHBoxLayout()
        self.v_line.addWidget(self.choose_question, alignment = Qt.AlignCenter)
        self.v_line.addWidget(self.data, alignment = Qt.AlignCenter)
        if payer == "мама":
            self.v_line.addWidget(self.mom, alignment = Qt.AlignCenter)
        else:
            self.v_line.addWidget(self.dad, alignment = Qt.AlignCenter)
        self.v_line.addWidget(self.money1, alignment = Qt.AlignCenter)
        self.v_line.addWidget(self.accept, alignment = Qt.AlignCenter)
        self.v_line.addLayout(self.h_line)
        self.setLayout(self.v_line)
    def connects(self):
        self.accept.clicked.connect(self.excel)
date = ""
letter = alphavit[0]
name_coordinats = []
data_find = "none"
for cellObj in sheet["A1":"J2"]:
    for i in cellObj:
        data_find = i.value
        if data_find == "Дата":
            name_coordinats.append(i.coordinate)
for i in range(len(name_coordinats)):
    coordinata = name_coordinats[i]
    index = alphavit.index(coordinata[0])
    for cellObj in sheet[coordinata[0] + str(int(coordinata[1]) + 2): alphavit[index+2] + str(48)]:
        for j in cellObj:
            if j.value == None:
                coordinats = j.coordinate
                #print(j.coordinate, 111, alphavit[index+1] + str(coordinata[1]))
                index_letter = alphavit.index(j.coordinate[0])
                if alphavit[index_letter-1] == alphavit[index]:
                    payer = "мама"
                    mw = FirstWin()
                    app.exec_()
                elif alphavit[index_letter-2] == alphavit[index]:
                    payer = "папа"
                    mw = FirstWin()
                    app.exec_()
            elif len(str(j.value)) > 10:
                date = str(j.value)
                date = date[8:10] + date[4:8] + date[0:4]
                if int(date[6:10]) > year_now:
                    exit()
                elif int(date[3:5]) > month_now and int(date[6:10]) == year_now:
                    exit()
                elif int(date[0:2]) > day_now and int(date[3:5]) == month_now and int(date[6:10]) == year_now:
                    exit()
                else:
                    pass