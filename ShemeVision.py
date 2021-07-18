from PyQt5.QtWidgets import (QApplication, QWidget, QMainWindow, QPushButton, QAction, QHeaderView, QLineEdit,
                             QTableWidget, QTableWidgetItem, QVBoxLayout, QHBoxLayout)
# from PyQt5.QtGui import QIcon, Qt
from PyQt5.QtGui import QPainter, QStandardItemModel, QIcon
from PyQt5.Qt import Qt
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import numpy as np
import smtplib as root
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tkinter import *
from email.mime.base import MIMEBase
from email import encoders


class DataEntryForm(QWidget):
    def __init__(self):
        super().__init__()
        #self.items = 0

        # left side
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(('Участок', 'Длина', 'Марка провода', 'P, кВт', 'Q, квар', 'U, кВ'))
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.layoutRight = QVBoxLayout()

        # right side
        self.table1 = QTableWidget()
        self.table1.setColumnCount(8)
        self.table1.setHorizontalHeaderLabels(('R, Ом', 'X, Ом', 'dP, кВт', 'dQ, квар', 'dS, кВА', 'dU, кВ', 'dU, %','dSсум., кВА'))
        self.table1.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # buttons
        self.lineEditDescription = QLineEdit()
        self.lineEditPrice = QLineEdit()
        self.butonGetExcel = QPushButton('Импорт')
        self.butonVerification = QPushButton('Верификация')
        self.buttonClear = QPushButton('Очистка')
        self.buttonExport = QPushButton('Экспорт')
        self.buttonSendEmail = QPushButton('Отправка')
        self.buttonCalculate = QPushButton('Рассчет')
        self.buttonQuit = QPushButton('Выход')

        # self.buttonAdd.setEnabled(False)

        # right buttons
        self.layoutRight.setSpacing(10)
        self.layoutRight.addWidget(self.butonGetExcel)
        self.layoutRight.addWidget(self.butonVerification)
        self.layoutRight.addWidget(self.buttonClear)
        self.layoutRight.addWidget(self.buttonExport)
        self.layoutRight.addWidget(self.buttonSendEmail)
        self.layoutRight.addWidget(self.buttonCalculate)
        self.layoutRight.addWidget(self.buttonQuit)

        self.layout = QHBoxLayout()
        self.layout.addWidget(self.table, 100)
        self.layout.addWidget(self.table1, 100)
        self.layout.addLayout(self.layoutRight, 10)

        self.setLayout(self.layout)

        self.buttonQuit.clicked.connect(lambda: app.quit())
        self.buttonClear.clicked.connect(self.reset_table)
        self.buttonClear.clicked.connect(self.reset_table1)
        self.butonGetExcel.clicked.connect(self.getExcel)
        self.butonVerification.clicked.connect(self.makeVerification)
        self.buttonSendEmail.clicked.connect(self.send_file)
        self.buttonCalculate.clicked.connect(self.calc)
        self.buttonExport.clicked.connect(self.export)
        self.lineEditDescription.textChanged[str].connect(self.check_disable)
        self.lineEditPrice.textChanged[str].connect(self.check_disable)


    def calc (self):
        # global scheme_array
        self.reset_table1()
        xls = pd.ExcelFile('Марки проводов.xlsx')
        input_cable = pd.read_excel(xls, 'Лист1')
        cable_array = np.array(input_cable)
        scheme_array = df_array

        r0_dict = {}
        x0_dict = {}
        for i in range(len(cable_array)):
            r0_dict[cable_array[i, 0]] = cable_array[i, 1]
            x0_dict[cable_array[i, 0]] = cable_array[i, 2]

        def calc_res_r(r0, l):
            r = r0 * l
            return round(r, 3)

        def calc_res_x(x0, l):
            x = x0 * l
            return round(x, 3)

        def calc_delta_p(p, q, u, r):
            dp = (p ** 2 + q ** 2) / u ** 2 * r
            return round(dp, 3)

        def calc_delta_q(p, q, u, x):
            dq = (p ** 2 + q ** 2) / u ** 2 * x
            return round(dq, 3)

        def calc_delta_s(p, q):
            ds = (p ** 2 + q ** 2) ** 0.5
            return round(ds, 3)

        def calc_delta_u(p, q, r, x, u):
            du = (p * r + q * x) / u / 10 ** 3
            return round(du, 3)

        def calc_delta_u_perc(u, du):
            du_perc = du / u * 100
            return round(du_perc, 3)

        ## расчёт сопротивления R
        result_array = np.zeros((scheme_array.shape[0], 8))
        #result_array = np.full([scheme_array.shape[0], 8], None)

        for k in range(scheme_array.shape[0]):
            result_array[k, 0] = calc_res_r(r0_dict[scheme_array[k, 2]], scheme_array[k, 1])


        ## расчёт сопротивления X
        for k in range(scheme_array.shape[0]):
            result_array[k, 1] = calc_res_x(x0_dict[scheme_array[k, 2]], scheme_array[k, 1])


        ## расчёт потерь P
        for k in range(scheme_array.shape[0]):
            result_array[k, 2] = calc_delta_p(scheme_array[k, 3], scheme_array[k, 4], scheme_array[k, 5],
                                              result_array[k, 0])


        ## расчёт потерь Q
        for k in range(scheme_array.shape[0]):
            result_array[k, 3] = calc_delta_q(scheme_array[k, 3], scheme_array[k, 4], scheme_array[k, 5],
                                              result_array[k, 1])


        ## расчёт потерь S
        for k in range(scheme_array.shape[0]):
            result_array[k, 4] = calc_delta_s(result_array[k, 2], result_array[k, 3])


        ## расчёт потерь U
        for k in range(scheme_array.shape[0]):
            result_array[k, 5] = calc_delta_u(result_array[k, 2], result_array[k, 3], result_array[k, 0],
                                              result_array[k, 1], scheme_array[k, 5])


        ## расчёт потерь U %
        for k in range(scheme_array.shape[0]):
            result_array[k, 6] = calc_delta_u_perc(scheme_array[k, 5], result_array[k, 5])
        ## расчёт потерь dSсум., кВА
        result_array[0, 7] = round(sum(result_array[:, 4]),3)
        print(result_array[0, 7])
        result = np.c_[scheme_array, result_array]
        print(result)

        data_result = {'участок': result[:, 0], 'длина': result[:, 1], 'Марка провода': result[:, 2], 'P, кВт': result[:, 3],
                'Q, квар': result[:, 4], 'U, кВ': result[:, 5], 'R, Ом': result[:, 6], 'X, Ом': result[:, 7],
                'dP, кВт': result[:, 8],
                'dQ, квар': result[:, 9], 'dS, кВА': result[:, 10], 'dU, кВ': result[:, 11], 'dU, %': result[:, 12],'dSсум., кВА': result[:, 13]}
        data = {'R, Ом': result_array[:, 0], 'X, Ом': result_array[:, 1], 'dP, кВт': result_array[:, 2],
                'dQ, квар': result_array[:, 3], 'dS, кВА': result_array[:, 4], 'dU, кВ': result_array[:, 5], 'dU, %': result_array[:, 6]}

        global df_right_side
        df_right_side = result_array
        print(df_right_side)

        global fulldata
        fulldata = pd.DataFrame(data=data_result)
        #fulldata.to_excel('otchet.xlsx', index=False)
        self.fill_table_right_side()


    def getExcel(self):
        global df
        global df_array
        import_file_path = filedialog.askopenfilename()
        df = pd.read_excel(import_file_path)
        df_array = np.array(df)
        self.reset_table()
        self.fill_table()


    def fill_table_right_side(self):
        data_right_side = df_right_side
        print(data_right_side)
        for i in range(data_right_side.shape[0]):
            self.table1.insertRow(self.items)
            for j in range(data_right_side.shape[1]):
                print(data_right_side[-i-1, j])
                table_element = QTableWidgetItem(str(data_right_side[-i-1, j]))
                table_element.setTextAlignment(Qt.AlignCenter)
                self.table1.setItem(self.items, j, table_element)

    def fill_table(self):
        global df_array
        data = df_array
        for i in range(data.shape[0]):
            self.table.insertRow(self.items)
            for j in range(data.shape[1]):
                print(data[-i-1, j])
                table_element = QTableWidgetItem(str(data[-i-1, j]))
                table_element.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(self.items, j, table_element)

    def makeVerification(self):
        scheme_array_checking = np.delete(df_array, [0, 2], axis=1)
        if np.any(scheme_array_checking == 0) or np.any(scheme_array_checking < 0):
            messagebox.showinfo("Выполнено", "Ошибка: нулевые или отрицательные значения в исходных данных")
        else:
            messagebox.showinfo("Выполнено", "Исходные данные корректны")

    def export(self):
        fulldata.to_excel('otchet.xlsx', index=False)
        messagebox.showinfo("Выполнено", "Файл {} экспортирован".format("otchet.xlsx"))
        print('Файл экспортирован.')

    def check_disable(self):
        if self.lineEditDescription.text() and self.lineEditPrice.text():
            self.buttonExport.setEnabled(True)
        else:
            self.buttonExport.setEnabled(False)

    def reset_table(self):
        self.table.setRowCount(0)
        self.items = 0

    def reset_table1(self):
        self.table1.setRowCount(0)
        self.items = 0

    def send_file(self):
        screen = Tk()
        screen.resizable(width=False, height=False)
        screen.geometry('320x120')
        screen.title('Отправка отчёта')

        # Fucntion
        def send_mail(event):
            L = 'gerasim666test@mail.ru'
            P = 'starthack1'
            U = 'smpt.mail.ru'
            To = toaddr.get()
            T = topic.get()
            M = mess.get()
            N = 1
            filename = "otchet.xlsx"


            with open(filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {filename}")
            msg = MIMEMultipart()
            msg['Subject'] = T
            msg['From'] = L
            body = M
            msg.attach(MIMEText(body, 'plain'))
            msg.attach(part)
            msg.add_header('Content-Disposition', 'attachment', filename=('iso-8859-1'))
            server = root.SMTP_SSL(U, 465)
            server.login(L, P)
            server.sendmail(L, To, msg.as_string())
            server.quit()
            screen.quit()

        Ttoaddr = Label(text='Кому:', font='Consolas')
        toaddr = Entry(screen, font='Consolas')

        Ttopic = Label(text='Заголовок:', font='Consolas')
        topic = Entry(screen, font='Consolas')

        Tmess = Label(text='Сообщение:', font='Consolas')
        mess = Entry(screen, font='Consolas')

        enter = Button(text='Отправить', font='Consolas', width=20)

        Ttoaddr.grid(row=3, column=0, sticky=W, padx=1, pady=1)
        toaddr.grid(row=3, column=1, padx=1, pady=1)

        Ttopic.grid(row=4, column=0, sticky=W, padx=1, pady=1)
        topic.grid(row=4, column=1, padx=1, pady=1)

        Tmess.grid(row=5, column=0, sticky=W, padx=1, pady=1)
        mess.grid(row=5, column=1, padx=1, pady=1)

        enter.grid(row=8, column=1, padx=1, pady=1)

        # Bind
        enter.bind('<Button-1>', send_mail)  #

        # The end
        screen.mainloop()


class MainWindow(QMainWindow):
    def __init__(self, w):
        super().__init__()
        self.setWindowTitle('Scheme Vision')
        self.setWindowIcon(QIcon("expense.png"))
        self.resize(1400, 600)
        # exit action
        exitAction = QAction('Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.triggered.connect(lambda: app.quit())
        self.setCentralWidget(w)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = DataEntryForm()
    demo = MainWindow(w)
    demo.show()
    # root = Tk()
    # root.withdraw()
    sys.exit(app.exec_())

