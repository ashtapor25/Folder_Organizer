# main.py

import sys
from folder_organizer.core import ZipfileLongPaths
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QHBoxLayout, QVBoxLayout, QLineEdit, QLabel, QMessageBox
from PyQt5.QtCore import QCoreApplication

class DragDropQLE(QLineEdit):

    def __init__(self, parent):
      super().__init__(parent)
      self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [str(u.toLocalFile()) for u in event.mimeData().urls()]
        for f in files:
            self.setText(f)


class MyApp(QWidget):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):

        self.setAcceptDrops(True)

        okButton = QPushButton('OK')
        okButton.clicked.connect(self.organize_connect_zip)
        cancelButton = QPushButton('Cancel')
        cancelButton.clicked.connect(QCoreApplication.instance().quit)

        zip_label = QLabel('Zip File: ', self)
        xlsx_label = QLabel('Xlsx File: ', self)

        self.zip_input_qle = DragDropQLE(self)
        self.zip_input_qle.setAcceptDrops(True)
        self.zip_input_qle.setDragEnabled(True)
        self.xlsx_input_qle = DragDropQLE(self)
        self.xlsx_input_qle.setAcceptDrops(True)
        self.xlsx_input_qle.setDragEnabled(True)

        hbox1 = QHBoxLayout()
        hbox1.addStretch(1)
        hbox1.addWidget(zip_label)
        hbox1.addStretch(1)
        hbox1.addWidget(self.zip_input_qle)
        hbox1.addStretch(1)

        hbox2 = QHBoxLayout()
        hbox2.addStretch(1)
        hbox2.addWidget(xlsx_label)
        hbox2.addStretch(1)
        hbox2.addWidget(self.xlsx_input_qle)
        hbox2.addStretch(1)

        hbox3 = QHBoxLayout()
        hbox3.addStretch(1)
        hbox3.addWidget(okButton)
        hbox3.addWidget(cancelButton)
        hbox3.addStretch(1)

        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addLayout(hbox1)
        vbox.addLayout(hbox2)
        vbox.addStretch(1)
        vbox.addLayout(hbox3)
        vbox.addStretch(1)

        self.setLayout(vbox)

        self.setWindowTitle('My First Application')
        self.move(300, 300)
        self.resize(400, 200)
        self.setGeometry(300, 300, 300, 200)
        self.show()

    def organize_connect_zip(self):

        zip_url = self.zip_input_qle.text()
        xlsx_url = self.xlsx_input_qle.text()

        self.instance = ZipfileLongPaths(zip_url, 'r', xlsx_url)
        self.instance.extract_ET_name()

        reply = QMessageBox.question(self, 'Message', 'Connect Zip Organizing Finished', QMessageBox.Yes, QMessageBox.Yes)

        QCoreApplication.instance().quit()

def run_app():
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
