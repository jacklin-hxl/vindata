import sys
from processData import run

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
        self.label.setText("数据处理中, 可以先摸下木木呦 ^-^")
        summaryDataFile = self.mubanwenjian_textEdit.toPlainText()
        summaryDataSheet = self.mubansheet_textEdit.toPlainText()
        totalDataFile = self.zonshuju_textEdit.toPlainText()
        totalDataSheet = self.zonshujusheet_textEdit.toPlainText()
        run(summaryDataFile, summaryDataSheet, totalDataFile, totalDataSheet)
        self.label.setText("数据已完成，记得亲下木木呦 ^-^")


def main():
    app = QApplication(sys.argv)
    myUi = MainUi()
    myUi.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()