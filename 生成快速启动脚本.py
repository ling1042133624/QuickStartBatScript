#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
    【使用方法】
    普通文件，假如设置好默认打开方式点击加文件，
    之后点击生成脚本就可以再桌面的 “一键启动最近文件.bat”,
    打开所需要的的文件.

    特殊文件或者文件夹（需要指定打开软件的），点击 “加文件夹 或 特殊文件（需要指定软件打开的）”，
    会引导你点击的，只要按着指引，指定文件和打开方式就可以,
    之后点击生成脚本就可以再桌面的 “一键启动最近文件.bat” 打开所需要的的文件

"""
import os.path
import sys
from PyQt5.QtWidgets import QWidget, QPushButton, QVBoxLayout, QLabel, QApplication, QFileDialog, QMessageBox, QLineEdit
from packages.QCandyUi.CandyWindow import *
import winreg


def get_desktop():
    """
    获取桌面路径
    :return:
    """
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                         r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')  # 利用系统的链表
    return winreg.QueryValueEx(key, "Desktop")[0]  # 返回的是Unicode类型数据


@colorful('blueGreen', "cartoon2.ico")
class QuickStartBatScript(QWidget):
    def __init__(self):
        QWidget.__init__(self)
        self.setWindowTitle("生成快速启动脚本 V0.1.0")
        self.windowTitle()
        self.resize(400, 100)
        self.file_count = 0  # 文件计数
        self.file_list = []  # 文件名列表
        self.file_label_list = []  # 显示列表

        self.special_file_count = 0  # 特殊文件计数
        self.special_files_name_list = []  # 特殊文件路径
        self.special_open_modes_list = []  # 特殊文件打开方式路径

        self.layout = QVBoxLayout()  # 竖向布局

        self.bat_name_lineEdit = QLineEdit()
        self.bat_name_lineEdit.setPlaceholderText("生成脚本的名字，不填就默认")
        self.layout.addWidget(self.bat_name_lineEdit)

        self.generate_bat_btn = QPushButton("生成脚本")  # 按钮生成脚本
        self.generate_bat_btn.clicked.connect(self.create_bat)
        self.layout.addWidget(self.generate_bat_btn)
        # self.layout.addSpacing(30)

        self.append_bat_btn = QPushButton("追加脚本")  # 按钮生成脚本
        self.append_bat_btn.clicked.connect(self.append_bat)
        self.layout.addWidget(self.append_bat_btn)
        self.layout.addSpacing(30)

        self.add_file_btn = QPushButton("加文件")  # 按钮加文件按钮
        self.add_file_btn.clicked.connect(self.add_file_slot)
        self.layout.addWidget(self.add_file_btn)

        self.add_special_file_btn = QPushButton("加文件夹 或 特殊文件（需要指定软件打开的）")  # 按钮加文件按钮
        self.add_special_file_btn.clicked.connect(self.add_special_file_slot)
        self.layout.addWidget(self.add_special_file_btn)
        self.layout.addSpacing(30)
        self.setLayout(self.layout)

    def add_file_slot(self):
        """
        槽函数，自动生成 “文件N” 按钮，并且设置按钮的槽函数
        :return:
        """
        btn = QPushButton("文件 %s" % self.file_count)  # 生成按钮和显示label，并布局
        btn.clicked.connect(self.get_file_name)
        self.layout.addWidget(btn)
        self.file_count += 1
        contents = QLabel("")
        self.file_label_list.append(contents)
        self.layout.addWidget(contents)

    def add_special_file_slot(self):
        """
        槽函数，自动生成 “特殊文件或文件夹N” 按钮，并且设置按钮的槽函数
        :return:
        """
        btn = QPushButton("特殊文件或文件夹 %s" % self.special_file_count)  # 生成按钮和显示label，并布局
        btn.clicked.connect(self.get_special_file)
        self.layout.addWidget(btn)

        self.special_file_count += 1
        contents = QLabel("")  # 路径
        self.special_files_name_list.append(contents)
        self.layout.addWidget(contents)
        contents1 = QLabel("")  # 打开方式
        self.special_open_modes_list.append(contents1)
        self.layout.addWidget(contents1)

    def create_bat(self):
        """
        生成脚本，放在桌面上，主要用到cmd的 start 命令
        :return:
        """
        cmd = "@echo off\n"
        self.write_cmds(cmd, "w")

    def append_bat(self):
        """
        追加脚本
        :return:
        """
        cmd = self.get_cmd()

        self.write_cmds(cmd, "a")

    def write_cmds(self, cmd, mode):

        file_path = str(get_desktop()) + "\%s" % \
                    (self.bat_name_lineEdit.text() + ".bat" if self.bat_name_lineEdit.text() else "一键启动最近文件.bat")
        ok = self.if_file_exites(file_path)

        file_path = file_path if not ok else file_path[:-4] + "(1)" + file_path[-4:]

        with open(file_path, mode) as f:
            f.writelines(cmd)
        self.notify_box()

    def get_cmd(self, cmd=""):
        """

        :param cmd: cmd 前缀
        :return:
        """
        for i in self.file_list:
            cmd += 'start "" "%s"\n' % i
        for num, i in enumerate(self.special_files_name_list):
            if i.text():
                if self.special_open_modes_list[num].text() == "" or \
                        self.special_open_modes_list[num].text() == " ":
                    cmd += 'start " " "%s"\n' % (i.text())
                else:
                    cmd += 'start " " "%s" "%s"\n' % \
                           (self.special_open_modes_list[num].text(), i.text())
        return cmd

    def get_file_name(self):
        """
        获取普通文件，假如设置好默认打开方式用这个就可以
        :return:
        """
        sender = self.sender()

        fname, _ = QFileDialog.getOpenFileNames(self, 'Open file', str(get_desktop()), "*")
        self.file_label_list[int(str(sender.text()).split("文件")[-1].strip())].setText(str(fname))
        self.file_list.extend(fname)

    def get_special_file(self):
        """
        用户引导，选择文件夹或者特殊文件的路径，并且确定其打开方式
        :return:
        """
        fname = None
        sender = self.sender()
        messageBox = QMessageBox()
        # createWindow(messageBox, "blueGreen")
        messageBox.setWindowTitle('选择')
        messageBox.setText('选择是文件夹？还是特殊文件？')
        messageBox.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        btnY = messageBox.button(QMessageBox.Yes)
        btnY.setText('文件夹')
        btnN = messageBox.button(QMessageBox.No)
        btnN.setText('特殊文件')
        messageBox.exec_()
        if messageBox.clickedButton() == btnY:
            fname = QFileDialog.getExistingDirectory(self, "选取文件夹", str(get_desktop()))
        elif messageBox.clickedButton() == btnN:
            fname, _ = QFileDialog.getOpenFileName(self, '选取特殊文件', str(get_desktop()), "*")
        self.select_fun(fname, sender)

    def select_fun(self, fname, sender):
        """
        get_special_file的复用函数，主要作用引导用户获取路径和打开方式
        :param fname:
        :param sender:
        :return:
        """
        self.special_files_name_list[int(str(sender.text()).split("特殊文件或文件夹")[-1].strip())].setText(str(fname))

        messageBox1 = QMessageBox()
        # createWindow(messageBox1, "blueGreen")
        messageBox1.setWindowTitle('选择')
        messageBox1.setText('是否需要特殊打开？')
        messageBox1.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        btnY = messageBox1.button(QMessageBox.Yes)
        btnY.setText('是')
        btnN = messageBox1.button(QMessageBox.No)
        btnN.setText('否')
        messageBox1.exec_()
        if messageBox1.clickedButton() == btnY:
            print('需要特殊打开')
            open_f, ok1 = QFileDialog.getOpenFileName(self,
                                                      "多文件选择",
                                                      str(get_desktop()),
                                                      "All Files (*);;Text Files (*.txt)")
            self.special_open_modes_list[int(str(sender.text()).split("特殊文件或文件夹")[-1].strip())].setText(str(open_f))

        elif messageBox1.clickedButton() == btnN:
            self.special_open_modes_list.append("")

    def notify_box(self):
        messageBox = QMessageBox()
        # createWindow(messageBox, "blueGreen")
        messageBox.setWindowTitle('提示')
        messageBox.setText('已经生成好脚本了~\n'
                            '在 %s' % (get_desktop()) + "\%s" % (
                                self.bat_name_lineEdit.text() + ".bat" if self.bat_name_lineEdit.text() else "一键启动最近文件.bat")
                            )
        messageBox.setStandardButtons(QMessageBox.Ok)
        messageBox.exec_()

    def if_file_exites(self, file_path):
        return os.path.isfile(file_path)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ui = QuickStartBatScript()
    ui.show()
    sys.exit(app.exec_())
