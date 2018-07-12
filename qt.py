import sys
from PyQt5.QtWidgets import QApplication, QWidget, QMainWindow, QPushButton,\
QMessageBox,QLineEdit, QHBoxLayout, QGroupBox, QVBoxLayout, QFileDialog, QLabel,\
QGridLayout
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot

import excelParser

class App(QWidget):
    def __init__(self):
        super().__init__();
        self.title = 'Excel 处理程序'
        self.left = 200
        self.top = 200
        self.width = 320
        self.height = 0
        self.input_fname = ''
        self.output_fname = ''


        self.initUI()


    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        #core = QWidget()
        #layout = QVBoxLayout()
        #core.setLayout(layout)
        #self.setCentralWidget(core)

        layout = QGridLayout()
        self.setLayout(layout)
        layout.setColumnStretch(0,2)
        layout.setColumnStretch(1,1)


        label_input = QLabel("选择要处理的文件：")
        label_finput = QLabel('')
        layout.addWidget(label_input, 0,0)
        layout.addWidget(label_finput,1,0)

        label_output = QLabel("保存为：")
        label_foutput = QLabel('')
        layout.addWidget(label_output,2,0)
        layout.addWidget(label_foutput,3,0)

        btn_open = QPushButton('打开...')
        btn_save = QPushButton('存储为...')
        btn_start = QPushButton('开始处理')
        btn_exit = QPushButton('退出')

        layout.addWidget(btn_open,0,1)
        layout.addWidget(btn_save,1,1)
        layout.addWidget(btn_start,2,1)
        layout.addWidget(btn_exit,3,1)

        btn_open.clicked.connect(self.btn_open_click)
        btn_save.clicked.connect(self.btn_save_click)
        btn_start.clicked.connect(self.btn_start_click)
        btn_exit.clicked.connect(self.btn_exit_click)

        self.label_finput = label_finput
        self.label_foutput = label_foutput
        self.btn_start = btn_start
        self.btn_start.setEnabled(False)
        self.show()

    def btn_open_click(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"打开",
                                                  "", "Excel (*.xls);;All Files(*)",
                                                  options=options)
        if fileName:
            print(fileName)
            self.input_fname = fileName
            t = fileName.split('/')
            self.label_finput.setText(t[-1])

    def btn_save_click(self):
        name = self.label_finput.text()
        if name=='':
            return

        index = name.rfind('.')
        name = name[0:index] + '整理' + name[index:]
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "保存为", name, "All Files (*)", options=options)
        if fileName:
            print(fileName)
            self.output_fname = fileName
            t = fileName.split('/')
            self.label_foutput.setText(t[-1])
            self.btn_start.setEnabled(True)

    def btn_start_click(self):
        if self.label_finput.text()=='':
            box = QMessageBox(QMessageBox.Warning,'提示','没有打开文件')
            box.addButton(self.tr('好'), QMessageBox.YesRole)
            box.exec_()
        elif self.label_foutput.text()=='':
            box = QMessageBox(QMessageBox.Warning, '提示', '没有指定保存文件')
            box.addButton(self.tr('好'), QMessageBox.YesRole)
            box.exec_()
        else:
            print ('start')
            self.parseExcel()

    def parseExcel(self):
        succ = excelParser.convert(self.input_fname, self.output_fname)
        if succ == 0:
            box = QMessageBox(QMessageBox.Warning, '提示', '处理完成')
            box.addButton(self.tr('好'), QMessageBox.YesRole)
            box.exec_()

    def btn_exit_click(self):
        self.close()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
