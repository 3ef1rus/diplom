from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QApplication,QMainWindow

import sys

def add_lable():
    print("tap")


def application():
    app=QApplication(sys.argv)
    window = QMainWindow()
    
    window.setWindowTitle("Джарвис")
    window.setGeometry(300,250,500,400)
    
    main_text = QtWidgets.QLabel(window)
    main_text.setText("Это надписть")
    main_text.move(100,100)
    main_text.adjustSize()
    btn=QtWidgets.QPushButton(window)
    btn.move(70,150)
    btn.setText("Нажми")
    btn.setFixedWidth(200)
    btn.clicked.connect(add_lable)
    
    window.show()
    sys.exit(app.exec_())
    
if __name__=="__main__":
    application()