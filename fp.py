import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QMessageBox, QApplication

from main import *
from ui.fp2 import Ui_MainWindow
from analyse import *
class Ui_fp(QtWidgets.QMainWindow,Ui_MainWindow):
    def __init__(self):
        super(Ui_fp,self).__init__()
        self.setupUi(self)

        self.flag = False
        self.flag_2 = False
        self.flag_3 = False
        self.flag_4 = False
        self.FileDir = str()

        self.CheckboxDict = {1: '密码区', 2: '发票代码', 3: '发票号码', 4: '开票日期', 5: '机器编号', 6: '校验码',
                             7: '购买方名称', 8: '购买方纳税人识别号', 9: '购买方地址、电话', 10: '购买方开户行及账号',
                             11: '销售方名称', 12: '销售方纳税人识别号', 13: '销售方地址、电话',
                             14: '销售方开户行及账号',
                             15: '收款人', 16: '复核', 17: '开票人'}

        self.CheckboxDictEmpty = {1: '', 2: '', 3: '', 4: '', 5: '', 6: '',
                                  7: '购买方名称', 8: '购买方纳税人识别号', 9: '购买方地址、电话',
                                  10: '购买方开户行及账号',
                                  11: '销售方名称', 12: '销售方纳税人识别号', 13: '销售方地址、电话',
                                  14: '销售方开户行及账号',
                                  15: '', 16: '', 17: ''}

        self.CheckboxDictFull = {1: '密码区', 2: '发票代码', 3: '发票号码', 4: '开票日期', 5: '机器编号', 6: '校验码',
                                 7: '购买方名称', 8: '购买方纳税人识别号', 9: '购买方地址、电话',
                                 10: '购买方开户行及账号',
                                 11: '销售方名称', 12: '销售方纳税人识别号', 13: '销售方地址、电话',
                                 14: '销售方开户行及账号',
                                 15: '收款人', 16: '复核', 17: '开票人'}

        self.init_solt()

    def init_solt(self):


        self.checkBox_6.clicked.connect(lambda: self.ch(self.checkBox_6))
        self.checkBox_7.clicked.connect(lambda: self.ch(self.checkBox_7))
        self.checkBox_9.clicked.connect(lambda: self.ch(self.checkBox_9))
        self.checkBox_10.clicked.connect(lambda: self.ch(self.checkBox_10))
        self.checkBox_12.clicked.connect(lambda: self.ch(self.checkBox_12))
        self.checkBox_13.clicked.connect(lambda: self.ch(self.checkBox_13))
        self.checkBox_15.clicked.connect(lambda: self.ch(self.checkBox_15))
        self.checkBox_16.clicked.connect(lambda: self.ch(self.checkBox_16))
        self.checkBox_19.clicked.connect(lambda: self.ch(self.checkBox_19))

        self.pushButton_5.clicked.connect(self.open_folder)  # 选择文件夹
        self.pushButton_6.clicked.connect(self.open_file)  # 选择文件
        self.pushButton_9.clicked.connect(self.start)  # 开始识别
        self.pushButton_7.clicked.connect(self.check_all)  # 全选
        self.pushButton_8.clicked.connect(self.not_check_all)  # 全部选
        self.pushButton_10.clicked.connect(self.open_jpd)  # 打开图片
        self.pushButton_11.clicked.connect(self.open_excel)  # 打开表格
        self.pushButton.clicked.connect(self.open_1)  # 跳转到第一页
        self.pushButton_12.clicked.connect(self.check)  # 选择文件
        self.pushButton_13.clicked.connect(self.tongji)  # 统计
        self.pushButton_14.clicked.connect(self.lock)  # 查看

        self.pushButton_2.clicked.connect(self.open_2)
        self.pushButton_3.clicked.connect(self.open_3)
        self.pushButton_4.clicked.connect(self.open_4)


    def open_1(self):
        self.stackedWidget.setCurrentIndex(0)
    def open_2(self):
        self.stackedWidget.setCurrentIndex(1)
    def open_3(self):
        self.stackedWidget.setCurrentIndex(2)
    def open_4(self):
        self.stackedWidget.setCurrentIndex(3)




    def ch(self, che):  # 复选框

        dic2 = {
            "密码区": 1,
            "发票代码": 2,
            "发票号码": 3,
            "机器编号": 5,
            "复核": 16,
            "开票日期": 4,
            "校验码": 6,
            "收款人": 15,
            "开票人": 17
        }
        print(che.isChecked())
        if che.isChecked():  # 如果复选框被选中
            temp = che.text()  # 获取文本
            numtemp = dic2[che.text()]  # 获取数字
            self.CheckboxDict[numtemp] = temp
            print(self.CheckboxDict)

        else:
            print()
            numtemp = dic2[che.text()]  # 获取数字
            self.CheckboxDict[numtemp] = ''
            print(self.CheckboxDict)

    def open_folder(self):  # 选择文件夹
        fileInfo = QtWidgets.QFileDialog.getExistingDirectory(None, "选取文件夹")
        self.FileDir = fileInfo
        self.lineEdit.setText(fileInfo)
        if (self.lineEdit.text() != ""):
            self.flag = True

    def open_file(self):  # 选择文件
        fileInfo = QtWidgets.QFileDialog.getOpenFileName(None, "选择文件", "", "*.pdf")  # 选择文件
        # 第一个参数 None 表示这个对话框没有父窗口。第二个参数 "选择文件" 是对话框的标题。第三个参数是一个默认路径，这里为空字符串，表示没有默认路径。第四个参数 "*.pdf" 是一个文件模式，表示只显示 PDF 文件。
        self.FileDir = fileInfo[0]
        self.lineEdit.setText(fileInfo[0])
        if (self.lineEdit.text() != ""):
            self.flag = True

    def start(self):  # 开始识别
        if self.flag == True:
            Main(self.FileDir, self.CheckboxDict)
            QMessageBox.information(self, "提示", "识别完成", QMessageBox.Ok)
            self.flag_2 = True
        else:

            QMessageBox.information(self, "提示", "请先选择相应的发票文件夹。", QMessageBox.Ok)


    def open_jpd(self):  # 打开图片
        if self.flag_2 == True:
            file()
        else:
            QMessageBox.information(self, "提示", "请先点击识别。", QMessageBox.Ok)

    def open_excel(self):
        if self.flag_2 == True:
            dir()
        else:
            QMessageBox.information(self, "提示", "请先点击识别。", QMessageBox.Ok)

    def check_all(self):  # 全选
        self.checkBox_6.setChecked(True)
        self.checkBox_7.setChecked(True)
        self.checkBox_9.setChecked(True)
        self.checkBox_10.setChecked(True)
        self.checkBox_12.setChecked(True)
        self.checkBox_13.setChecked(True)
        self.checkBox_15.setChecked(True)
        self.checkBox_16.setChecked(True)
        self.checkBox_19.setChecked(True)
        self.CheckboxDict = self.CheckboxDictFull

    def not_check_all(self):  # 全不选
        self.checkBox_6.setChecked(False)
        self.checkBox_7.setChecked(False)
        self.checkBox_9.setChecked(False)
        self.checkBox_10.setChecked(False)
        self.checkBox_12.setChecked(False)
        self.checkBox_13.setChecked(False)
        self.checkBox_15.setChecked(False)
        self.checkBox_16.setChecked(False)
        self.checkBox_19.setChecked(False)
        self.CheckboxDict = self.CheckboxDictEmpty

    def check(self):  # 选择文件夹
        fileInfo = QtWidgets.QFileDialog.getExistingDirectory(None, "选取文件夹")
        self.FileDir = fileInfo
        self.lineEdit_2.setText(fileInfo)
        if (self.lineEdit_2.text() != ""):
            self.flag_3 = True
    def tongji(self):  # 统计
        if self.flag_3 == True:
            text = self.lineEdit_2.text()
            t = classify(text)
            QMessageBox.information(self, "提示", "统计完成", QMessageBox.Ok)
            self.label_11.setPixmap(QtGui.QPixmap(t))  # 我的图片格式为png.与代码在同一目录下
            self.label_11.setScaledContents(True)  # 图片大小与label适应，否则图片可能显示不全
            self.flag_4 = True
        else:
            QMessageBox.information(self, "提示", "请先选择文件夹。", QMessageBox.Ok)
    def lock(self):  # 查看图片
        if self.flag_4 == True:
            text = self.lineEdit_2.text()
            file_1(text)
        else:
            QMessageBox.information(self, "提示", "请先统计。", QMessageBox.Ok)

if __name__ == '__main__':
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    # 适应高DPI设备
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    # 解决图片在不同分辨率显示模糊问题
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = Ui_fp()
    MainWindow.show()

    sys.exit(app.exec_())
