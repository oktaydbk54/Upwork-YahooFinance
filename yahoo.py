import yfinance as yf
from datetime import date
import pandas as pd

from PyQt5.QtWidgets import *
import sys

class Main(QWidget):
    
    def __init__(self):
        super(Main,self).__init__()

        self.setWindowTitle("Yahoo Finance Downloader")
        self.setFixedSize(400, 150)

        self.companiesLabel = QLabel("Companies : ")
        self.companiesLineEdit = QLineEdit()

        self.applyButton = QPushButton("Apply")

        HBox = QHBoxLayout()
        HBox.addWidget(QLabel("Which stock would you like to pulling ?"))

        HBox0 = QHBoxLayout()
        HBox0.addWidget(self.companiesLabel)
        HBox0.addWidget(self.companiesLineEdit)

        HBox1 = QHBoxLayout()
        HBox1.addStretch()
        HBox1.addWidget(self.applyButton)

        VBox0 = QVBoxLayout()
        VBox0.addLayout(HBox)
        VBox0.addLayout(HBox0)
        VBox0.addLayout(HBox1)

        self.setLayout(VBox0)

        self.applyButton.clicked.connect(self.yahooFunc)

        self.show()

    def yahooFunc(self):
        from datetime import date

        try:
            item = self.companiesLineEdit.text()
            date = date.today()
            result = yf.download(item, end = date)
            result.sort_values(by = "Date", ascending = False, inplace = True)
            result.reset_index(inplace= True)
            result["Date"] = result["Date"].astype(str)
            result.to_excel(item + ".xlsx", index = False)

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setText("Download Completed")
            msg.setWindowTitle("Information")
            msg.exec_()

        except:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText("Error!!!\nPlease check the values.")
            msg.setWindowTitle("Error")
            msg.exec_()
     


if __name__ == '__main__':
    app = QApplication(sys.argv)
    MAIN = Main()
    sys.exit(app.exec_())

