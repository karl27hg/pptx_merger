from PyQt5.QtWidgets import QApplication, QTableWidget, QHeaderView, \
      QFileDialog, QTableWidgetItem
from PyQt5 import uic
import os

class PPTXMergerGUi(*uic.loadUiType("ui/merger_gui.ui")):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setupUi(self)
        self.__set_table_header()
        # button connect
        self.outputBtn.clicked.connect(self.__on_click_output_set_btn)
        self.inputAddBtn.clicked.connect(self.__on_click_input_add_btn)
        self.__input_files = []

    def __on_click_output_set_btn(self):
        # todo last file pos
        filename, _ = QFileDialog.getSaveFileName(self, "Save file names", "./", "Power Point (*.pptx)")
        if filename:
            self.outputLineEdit.setText(filename)

    def __on_click_input_add_btn(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Open PPTX files", "./", "Power Point (*.pptx)")
        if files:
            self.__input_files = files
            self.__update_inputTableWidget()

    def __update_inputTableWidget(self):
        assert isinstance(self.inputTableWidget, QTableWidget)
        self.inputTableWidget.setRowCount(len(self.__input_files))
        for r, file in enumerate(self.__input_files):
            filename = os.path.basename(file)
            self.inputTableWidget.setItem(r, 0, QTableWidgetItem(filename))

    def __set_table_header(self):
        assert isinstance(self.inputTableWidget, QTableWidget)
        headerLabels = ["file name"]
        self.inputTableWidget.clear()
        self.inputTableWidget.setColumnCount(len(headerLabels))
        self.inputTableWidget.setHorizontalHeaderLabels(headerLabels)
        header = self.inputTableWidget.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Stretch)
        # todo




if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    pptx_merger_gui = PPTXMergerGUi()
    pptx_merger_gui.show()
    app.exec_()