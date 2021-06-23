import sys
from PyQt5.QtWidgets import QWidget, QCheckBox

class CCheckBox(QWidget):

    def __init__(self,name):
        super().__init__()
        self.checkBox = QCheckBox(name, self)
        self.checkBox.toggle()
        self.checkBox.stateChanged.connect(self.changed)

    def changed(self, state):
        self.checked = state




