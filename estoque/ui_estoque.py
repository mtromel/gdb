# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'estoque.ui'
##
## Created by: Qt User Interface Compiler version 6.11.0
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
from PySide6.QtWidgets import (QApplication, QMainWindow, QProgressBar, QPushButton,
    QSizePolicy, QStatusBar, QTabWidget, QTextBrowser,
    QWidget)

class Ui_Estoque(object):
    def setupUi(self, Estoque):
        if not Estoque.objectName():
            Estoque.setObjectName(u"Estoque")
        Estoque.resize(434, 258)
        Estoque.setAutoFillBackground(False)
        Estoque.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextOnly)
        Estoque.setTabShape(QTabWidget.TabShape.Rounded)
        self.centralwidget = QWidget(Estoque)
        self.centralwidget.setObjectName(u"centralwidget")
        self.pushButton = QPushButton(self.centralwidget)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(80, 50, 261, 91))
        font = QFont()
        font.setBold(True)
        self.pushButton.setFont(font)
        self.pushButton.setLocale(QLocale(QLocale.Portuguese, QLocale.Brazil))
        self.progressBar = QProgressBar(self.centralwidget)
        self.progressBar.setObjectName(u"progressBar")
        self.progressBar.setGeometry(QRect(120, 200, 201, 23))
        self.progressBar.setValue(0)
        self.textBrowser = QTextBrowser(self.centralwidget)
        self.textBrowser.setObjectName(u"textBrowser")
        self.textBrowser.setEnabled(False)
        self.textBrowser.setGeometry(QRect(80, 150, 261, 31))
        self.textBrowser.setAutoFillBackground(True)
        Estoque.setCentralWidget(self.centralwidget)
        self.statusbar = QStatusBar(Estoque)
        self.statusbar.setObjectName(u"statusbar")
        Estoque.setStatusBar(self.statusbar)

        self.retranslateUi(Estoque)

        QMetaObject.connectSlotsByName(Estoque)
    # setupUi

    def retranslateUi(self, Estoque):
        Estoque.setWindowTitle(QCoreApplication.translate("Estoque", u"Gerador de Relat\u00f3rio de Estoque", None))
#if QT_CONFIG(tooltip)
        self.pushButton.setToolTip(QCoreApplication.translate("Estoque", u"<html><head/><body><p><span style=\" font-weight:400;\">Clique para gerar o Relat\u00f3rio de Estoque</span></p></body></html>", None))
#endif // QT_CONFIG(tooltip)
        self.pushButton.setText(QCoreApplication.translate("Estoque", u"Gerar Relat\u00f3rio", None))
    # retranslateUi

