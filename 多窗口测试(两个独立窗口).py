import sys
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *

import sys
import urllib.request


class ChildWidget(QWidget):

    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        print('ChildWidget init')
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setWindowTitle('ChildWidget 模板')

        self.top_layout = QHBoxLayout()
        self.setLayout(self.top_layout)
        self.resize(500, 300)
        self.left_widget = QWidget()
        self.mid_layout = QHBoxLayout()
        self.right_widget = QWidget()
        self.setObjectName('Top_ChildWidget')
        self.setStyleSheet("#Top_ChildWidget{border: 2px solid red;}")

        self.top_layout.addWidget(self.left_widget)
        self.top_layout.addLayout(self.mid_layout)
        self.top_layout.addWidget(self.right_widget)

        self.next_btn = QPushButton('下一页')
        self.back_btn = QPushButton('返回')
        self.mid_layout.addWidget(self.next_btn)
        self.mid_layout.addWidget(self.back_btn)



class MainWidget(QWidget):
    main_widget_signal = pyqtSignal(dict)

    def __init__(self, parent=None, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        print("MainWidget init")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setWindowTitle('MainWidget 模板')

        #print(type(self).__name__)

        self.top_layout = QVBoxLayout()
        self.setLayout(self.top_layout)
        self.resize(500, 300)

        self.dataView = QTableView()
        self.model = QStandardItemModel(3, 0, self)

        self.model.setHorizontalHeaderItem(0, QStandardItem("姓名"))
        self.model.setHorizontalHeaderItem(1, QStandardItem("性别"))
        self.model.setHorizontalHeaderItem(2, QStandardItem("体重(kg)"))

        self.dataView.setModel(self.model)

        self.child_layout = QHBoxLayout()
        self.next_btn = QPushButton('下一页')
        self.back_btn = QPushButton('返回')
        self.child_layout.addWidget(self.next_btn)
        self.child_layout.addWidget(self.back_btn)

        self.top_layout.addLayout(self.child_layout)
        self.top_layout.addWidget(self.dataView)

        self.next_btn.clicked.connect(self.slot_next_btn_func)
        self.back_btn.clicked.connect(self.slot_back_btn_func)

    def slot_next_btn_func(self):
        msg_dict = {}
        msg_dict['name'] = 'MainWidget'
        print('slot_next_btn_func')
        msg_dict['action'] = 'next'
        self.main_widget_signal.emit(msg_dict)
        #print('MainWidget slot_next_btn_func', msg_dict)


    def slot_back_btn_func(self):
        msg_dict = {}
        msg_dict['name'] = 'MainWidget'
        msg_dict['action'] = 'back'
        self.main_widget_signal.emit(msg_dict)
        print('MainWidget slot_back_btn_func', msg_dict)


def slot_call_back_func(my_dict):
    print("slot_call_back_func", my_dict)



if __name__ == '__main__':

    app = QApplication(sys.argv)

    main = MainWidget()

    main.main_widget_signal.connect(slot_call_back_func)

    child_win = ChildWidget()
    main.show()
    child_win.show()

    sys.exit(app.exec_())