# This Python file uses the following encoding: utf-8
import sys
from PySide6.QtWidgets import QApplication, QMainWindow, QTabWidget
from ui_components import ConversionTab, CombinedConversionTab
from theme import APP_QSS

class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("PDF Tool Suite")
        self.resize(600, 400)
        self.setStyleSheet(APP_QSS)
        
        tabs = QTabWidget()
        tabs.addTab(CombinedConversionTab("doc"), "文件转换")
        tabs.addTab(CombinedConversionTab("extract"), "报告提取")
        
        self.setCentralWidget(tabs)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
