import sys

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication
from pyqt5_plugins.examplebutton import QtWidgets

from fp import Ui_fp

if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    # 适应高DPI设备
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    # 解决图片在不同分辨率显示模糊问题
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = Ui_fp()
    MainWindow.show()

    sys.exit(app.exec_())
