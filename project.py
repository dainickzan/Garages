import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog
app = QApplication(sys.argv)
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.button = QPushButton()
        self.setCentralWidget(self.button)
        self.button.clicked.connect(self.click)
    def click(self):
        print(QFileDialog.getOpenFileName())
window = MainWindow()
window.show()
app.exec()
