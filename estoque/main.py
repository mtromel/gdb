import sys
from PySide6.QtWidgets import QApplication, QMainWindow
from PySide6.QtCore import QObject, Signal, QThread
from ui_estoque import Ui_Estoque
from estoque import GerarRelatorioEstoque

class RelatorioWorker(QObject):
    finished = Signal()
    progress = Signal(int)

    def executar(self):
        from estoque import GerarRelatorioEstoque

        self.progress.emit(10)
        relatorio = GerarRelatorioEstoque()
        relatorio.executar()
        self.progress.emit(100)
        self.finished.emit()

class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()

        self.ui = Ui_Estoque()
        self.ui.setupUi(self)
        self.ui.pushButton.clicked.connect(self.gerar_relatorio)
    
    def gerar_relatorio(self):
        self.thread = QThread()
        self.worker = RelatorioWorker()
        self.worker.moveToThread(self.thread)
        self.thread.started.connect(self.worker.executar)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.progress.connect(self.ui.progressBar.setValue)
        self.worker.finished.connect(self.finalizar_ui)
        self.ui.textBrowser.setText("Gerando o relatório de estoque...")
        self.ui.pushButton.setEnabled(False)
        self.thread.start()

    def finalizar_ui(self):
        self.ui.textBrowser.setText("Relatório de estoque gerado com sucesso!")
        self.ui.pushButton.setEnabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())