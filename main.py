import sys
import time
import traceback

from PyQt5.QtCore import QBasicTimer

from common.logger import logger
from handle import Sku, Strain

from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from ui import Ui_Form
import threading

class MainUi(QMainWindow,QFileDialog,Ui_Form):
    def __init__(self,parent=None):
        super(MainUi,self).__init__(parent)
        self.setupUi(self)
        self.confirm_Button.clicked.connect(self.confirm)
        self.mubandakai_pushButton.clicked.connect(self.muBanOpenfile)
        self.pushButton_2.clicked.connect(self.zonShuJuOpenfile)
        self.progressBar.setMinimum(0)
        self.progressBar.setMaximum(100)
        self.flag = [0, True]
        self.action = None
        self.timer = QBasicTimer()
        self.msg_box = QMessageBox(QMessageBox.Critical, "ERROR", "文件处理错误，请查看vinda.log文件")

    def timerEvent(self, a0: 'QTimerEvent') -> None:
        if self.flag[0] > 100:
            self.progressBar.setValue(100)
            self.timer.stop()
        else:
            if self.flag[1]:
                self.progressBar.setValue(self.flag[0])
            else:
                self.msg_box.exec()
                self.timer.stop()
                self.progressBar.setValue(0)
    def muBanOpenfile(self):
        filedir = self.getOpenFileName(self,"open file","./","All Files (*)")[0]
        self.mubanwenjian_textEdit.setText(filedir)

    def zonShuJuOpenfile(self):
        filedir = self.getOpenFileName(self,"open file","./","All Files (*)")[0]
        self.zonshuju_textEdit.setText(filedir)

    def confirm(self):
        self.flag[0] = 0
        self.flag[1] = True
        self.progressBar.setValue(self.flag[0])
        self.timer.start(500, self)
        data_type = self.type_ComboBox.currentText()
        currentDate = self.textEdit.toPlainText()
        summaryDataFile = self.mubanwenjian_textEdit.toPlainText()
        summaryDataSheet = self.mubansheet_textEdit.toPlainText()
        totalDataFile = self.zonshuju_textEdit.toPlainText()
        totalDataSheet = self.zonshujusheet_textEdit.toPlainText()
        yDataFile = self.qunianfile_textEdit.toPlainText()
        yDataSheet = self.quniansheet_textEdit.toPlainText()
        is_do = self.checkBox.isChecked()
        if data_type == "sku":
            self.action = threading.Thread(target=sku_worker, args=(self.flag, currentDate, summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet))
            self.action.start()
        elif data_type == "品系":
            self.action = threading.Thread(target=strain_worker, args=(is_do, self.flag, currentDate, summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet, yDataFile, yDataSheet))
            self.action.start()


def sku_worker(flag, currentDate, summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet):
    try:
        sku = Sku(flag, currentDate, summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet)
        sku.start()
    except Exception:
        logger.debug(traceback.format_exc())
        flag[1] = False

def strain_worker(is_do, flag, currentDate, summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet, yDataFile, yDataSheet):
    try:
        starin = Strain(is_do, flag, currentDate, summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet, yDataFile, yDataSheet)
        starin.start()
    except Exception:
        logger.debug(traceback.format_exc())
        flag[1] = False

def main():
    app = QApplication(sys.argv)
    myUi = MainUi()
    myUi.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()