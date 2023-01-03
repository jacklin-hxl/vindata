import sys

from handle import Sku
from processData import run, runThree, runTow

from PyQt5.QtWidgets import QApplication,QMainWindow,QFileDialog
from ui import Ui_Form

class MainUi(QMainWindow,QFileDialog,Ui_Form):
    def __init__(self,parent=None):
        super(MainUi,self).__init__(parent)
        self.setupUi(self)
        self.confirm_Button.clicked.connect(self.confirm)
        self.mubandakai_pushButton.clicked.connect(self.muBanOpenfile)
        self.pushButton_2.clicked.connect(self.zonShuJuOpenfile)

    def muBanOpenfile(self):
        filedir = self.getOpenFileName(self,"open file","./","All Files (*)")[0]
        self.mubanwenjian_textEdit.setText(filedir)

    def zonShuJuOpenfile(self):
        filedir = self.getOpenFileName(self,"open file","./","All Files (*)")[0]
        self.zonshuju_textEdit.setText(filedir)

    def confirm(self):
        
        data_type = self.type_ComboBox.currentText()
        currentDate = self.textEdit.toPlainText()
        summaryDataFile = self.mubanwenjian_textEdit.toPlainText()
        summaryDataSheet = self.mubansheet_textEdit.toPlainText()
        totalDataFile = self.zonshuju_textEdit.toPlainText()
        totalDataSheet = self.zonshujusheet_textEdit.toPlainText()
        yDataFile = self.qunianfile_textEdit.toPlainText()
        yDataSheet = self.quniansheet_textEdit.toPlainText()
        if data_type == "sku":
            # run(summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet, currentDate)
            sku = Sku(currentDate, summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet)
            sku.start()
        elif data_type == "品系":
            runTow(summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet, yDataFile, yDataSheet, currentDate)
        elif data_type == "业态品类":
            runThree(summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet, yDataFile, yDataSheet)
        self.textBrowser.setText("数据已完成，记得亲下木木呦 ^-^")


def main():
    app = QApplication(sys.argv)
    myUi = MainUi()
    myUi.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()