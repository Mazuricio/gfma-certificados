# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'interface.ui'
##
## Created by: Qt User Interface Compiler version 6.5.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################

from PySide6.QtCore import (QCoreApplication, QDate, QDateTime, QLocale,
    QMetaObject, QObject, QPoint, QRect,
    QSize, QTime, QUrl, Qt)
from PySide6.QtGui import (QBrush, QColor, QConicalGradient, QCursor,
    QFont, QFontDatabase, QGradient, QIcon,
    QImage, QKeySequence, QLinearGradient, QPainter,
    QPalette, QPixmap, QRadialGradient, QTransform)
from PySide6.QtWidgets import (QApplication, QDateEdit, QGridLayout, QLabel,
    QLineEdit, QMainWindow, QPushButton, QSizePolicy,
    QStatusBar, QWidget, QFileDialog, QDialog, QDialogButtonBox, QVBoxLayout)


from datetime import datetime

from gerador import *
class CustomDialog(QDialog):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Mensagem")

        QBtn = QDialogButtonBox.Ok #| QDialogButtonBox.Cancel

        self.buttonBox = QDialogButtonBox(QBtn)
        self.buttonBox.accepted.connect(self.accept)
        #self.buttonBox.rejected.connect(self.reject)

        self.layout = QVBoxLayout()
        message = QLabel("Certificados gerados com sucesso")
        self.layout.addWidget(message)
        self.layout.addWidget(self.buttonBox)
        self.setLayout(self.layout)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(571, 465)
        sizePolicy = QSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.gridLayoutWidget = QWidget(self.centralwidget)
        self.gridLayoutWidget.setObjectName(u"gridLayoutWidget")
        self.gridLayoutWidget.setGeometry(QRect(0, 0, 571, 421))
        self.gridLayout = QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setObjectName(u"gridLayout")
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.label_6 = QLabel(self.gridLayoutWidget)
        self.label_6.setObjectName(u"label_6")

        self.gridLayout.addWidget(self.label_6, 5, 0, 1, 1)

        self.Titulo = QLineEdit(self.gridLayoutWidget)
        self.Titulo.setObjectName(u"Titulo")
        self.Titulo.setAlignment(Qt.AlignCenter)

        self.gridLayout.addWidget(self.Titulo, 0, 1, 1, 1)

        self.bt_saveas = QPushButton(self.gridLayoutWidget)
        self.bt_saveas.setObjectName(u"bt_saveas")

        self.gridLayout.addWidget(self.bt_saveas, 5, 1, 1, 1)

        self.Horas = QLineEdit(self.gridLayoutWidget)
        self.Horas.setObjectName(u"Horas")
        self.Horas.setAlignment(Qt.AlignCenter)

        self.gridLayout.addWidget(self.Horas, 3, 1, 1, 1)

        self.ciclo = QLineEdit(self.gridLayoutWidget)
        self.ciclo.setObjectName(u"ciclo")
        self.ciclo.setAlignment(Qt.AlignCenter)

        self.gridLayout.addWidget(self.ciclo, 2, 1, 1, 1)

        self.label_4 = QLabel(self.gridLayoutWidget)
        self.label_4.setObjectName(u"label_4")

        self.gridLayout.addWidget(self.label_4, 2, 0, 1, 1)

        self.label_5 = QLabel(self.gridLayoutWidget)
        self.label_5.setObjectName(u"label_5")

        self.gridLayout.addWidget(self.label_5, 3, 0, 1, 1)

        self.label_2 = QLabel(self.gridLayoutWidget)
        self.label_2.setObjectName(u"label_2")

        self.gridLayout.addWidget(self.label_2, 0, 0, 1, 1)

        self.bt_gerar = QPushButton(self.gridLayoutWidget)
        self.bt_gerar.setObjectName(u"bt_gerar")

        self.gridLayout.addWidget(self.bt_gerar, 7, 1, 1, 1)

        self.label = QLabel(self.gridLayoutWidget)
        self.label.setObjectName(u"label")

        self.gridLayout.addWidget(self.label, 4, 0, 1, 1)

        self.bt_read_excel = QPushButton(self.gridLayoutWidget)
        self.bt_read_excel.setObjectName(u"bt_read_excel")

        self.gridLayout.addWidget(self.bt_read_excel, 4, 1, 1, 1)

        self.label_3 = QLabel(self.gridLayoutWidget)
        self.label_3.setObjectName(u"label_3")

        self.gridLayout.addWidget(self.label_3, 1, 0, 1, 1)

        self.data = QDateEdit(self.gridLayoutWidget)
        self.data.setObjectName(u"data")
        self.data.setAlignment(Qt.AlignCenter)

        self.gridLayout.addWidget(self.data, 1, 1, 1, 1)

        self.ln_info = QLineEdit(self.gridLayoutWidget)
        self.ln_info.setObjectName(u"ln_info")
        self.ln_info.setAlignment(Qt.AlignCenter)
        self.ln_info.setReadOnly(True)

        self.gridLayout.addWidget(self.ln_info, 6, 1, 1, 1)

        self.label_7 = QLabel(self.gridLayoutWidget)
        self.label_7.setObjectName(u"label_7")

        self.gridLayout.addWidget(self.label_7, 6, 0, 1, 1)

        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)

        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.label_6.setText(QCoreApplication.translate("MainWindow", u"Salvar", None))
        self.bt_saveas.setText(QCoreApplication.translate("MainWindow", u"Buscar", None))
        self.Horas.setText(QCoreApplication.translate("MainWindow", u"1", None))
        self.ciclo.setText(QCoreApplication.translate("MainWindow", u"I Ciclo de Semin\u00e1rios do GFAMa", None))
        self.label_4.setText(QCoreApplication.translate("MainWindow", u"Ciclo", None))
        self.label_5.setText(QCoreApplication.translate("MainWindow", u"Horas", None))
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"Titulo", None))
        self.bt_gerar.setText(QCoreApplication.translate("MainWindow", u"GERAR", None))
        self.label.setText(QCoreApplication.translate("MainWindow", u"Excel", None))
        self.bt_read_excel.setText(QCoreApplication.translate("MainWindow", u"Buscar", None))
        self.label_3.setText(QCoreApplication.translate("MainWindow", u"data", None))
        self.ln_info.setText("")
        self.label_7.setText(QCoreApplication.translate("MainWindow", u"Info", None))
    # retranslateUi

class Ui(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("Gfma Certificados")
        self.lista = []
        self.salvar = None
        #setando data pro dia atual
        date_str = datetime.today().strftime('%d-%m-%Y')
        qdate = QDate.fromString(date_str, "d-M-yyyy")
        self.ui.data.setDate(qdate)

        ## botoes #
        self.ui.bt_read_excel.clicked.connect(self.find_excel)
        self.ui.bt_saveas.clicked.connect(self.saveas)
        self.ui.bt_gerar.clicked.connect(self.ui_gerar)
        ## Tela 1 ## 
        #self.ui.bt_file.clicked.connect(self.selec_file)

    def find_excel(self):
        file_filter = "Dat file(*.xltx)"
        response = QFileDialog.getOpenFileName(
            parent=self,
            caption="Excel",
            filter=file_filter
        )
        local = response[0]
        self.ui.bt_read_excel.setText(local.split('/')[-1])
        self.lista = ler(local)
        self.ui.ln_info.setText(f'Há {len(self.lista)} participantes')

    def saveas(self):
        response = QFileDialog.getExistingDirectory(
            parent=self,
            caption="Salvar"
        )
        self.salvar = response
        
        self.ui.bt_saveas.setText(response)
    def ui_gerar(self):

        ti = self.ui.Titulo.text()
        data = datetime.strptime(self.ui.data.text(), "%d/%m/%Y")
        ciclo = self.ui.ciclo.text()
        horas = self.ui.Horas.text()
        gerar(self.lista, ti, data.strftime("%d de %B de %Y"), ciclo, horas, self.salvar)
        dlg = CustomDialog()
        if dlg.exec():
            self.ui.bt_gerar.setText("Gerar novamente")

app = QApplication()
windows = Ui()
windows.show()
app.exec()