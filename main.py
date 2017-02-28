import os
import sys
import threading

from PyQt4 import QtGui, QtCore, Qt, uic
from PyQt4.QtGui import QTreeWidgetItem
from PyQt4.QtCore import pyqtSignal

from excel import ExcelBook
from doc_template import DocTemplate
from doc_writer import write_doc


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def build_tree_widget_item(parent, name, is_checkable=True):
    child = QTreeWidgetItem(parent)
    if is_checkable:
        child.setFlags(child.flags() | QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsTristate)
        child.setCheckState(0, QtCore.Qt.Unchecked)
    else:
        child.setFlags(QtCore.Qt.ItemIsEnabled)
    child.setText(0, name)

    return child


class MainWindow(QtGui.QWidget):
    update_progress_bar_signal = pyqtSignal('QString', int)
    generate_doc_finished_signal = pyqtSignal()

    def __init__(self):
        QtGui.QWidget.__init__(self)
        self.ui = uic.loadUi(resource_path("gui.ui"))

        self.picker = self.ui.Picker
        self.excel_button = self.ui.ExcelButton
        self.word_button = self.ui.WordButton
        self.generate_button = self.ui.GenerateDocButton
        self.output = self.ui.Output
        self.progress_bar = QtGui.QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.setMinimum(0)
        self.progress_label = QtGui.QLabel()
        self.progress_label.setFixedWidth(120)

        self.book = ExcelBook()
        self.doc = DocTemplate()

        self.constituency_picker = None
        self.dodgy_cell_picker = None
        self.empty_cell_picker = None
        self.missing_column_picker = None

        self.loading_counter = 0
        self.lock = threading.Lock()

        self.excel_loading_movie = QtGui.QMovie(resource_path("loading.gif"))
        self.word_loading_movie = QtGui.QMovie(resource_path("loading.gif"))

        self.excel_loading_movie.frameChanged.connect(self.set_excel_icon)
        self.word_loading_movie.frameChanged.connect(self.set_word_icon)

        self.doc.started.connect(self.word_loading_movie.start)
        self.doc.finished.connect(self.doc_loaded)

        self.update_progress_bar_signal.connect(self.update_progress_bar)
        self.generate_doc_finished_signal.connect(self.doc_generated)

        self.init_ui()

    def set_up_book(self):
        self.book.started.connect(self.excel_loading_movie.start)
        self.book.progress.connect(self.update_progress_bar)
        self.book.finished.connect(self.book_loaded)

    def init_ui(self):
        self.init_pickers()
        self.init_output()
        self.init_button_connections()
        self.init_status_bar()

        self.ui.show()

    def init_pickers(self):
        # self.year_picker = build_tree_widget_item(self.picker, "Years")
        self.constituency_picker = build_tree_widget_item(self.picker, "Constituencies")

    def init_status_bar(self):
        self.ui.statusBar().addPermanentWidget(self.progress_label, 0)
        self.ui.statusBar().addWidget(self.progress_bar, 1)

    def init_output(self):
        self.dodgy_cell_picker = build_tree_widget_item(self.output, "Cells with possible errors", False)
        self.dodgy_cell_picker.setDisabled(True)
        self.dodgy_cell_picker.setToolTip(0, "Cells that couldn't be decoded but aren't empty. Most likely a space")
        self.empty_cell_picker = build_tree_widget_item(self.output, "Empty cells", False)
        self.empty_cell_picker.setDisabled(True)
        self.missing_column_picker = build_tree_widget_item(self.output, "Table columns not in the Excel sheet", False)
        self.missing_column_picker.setDisabled(True)

    def init_button_connections(self):
        self.excel_button.clicked.connect(self.excel_button_clicked)
        self.word_button.clicked.connect(self.word_button_clicked)
        self.generate_button.clicked.connect(self.generate_doc)

    def load_picker_data(self):
        years = self.book.years

        # for year in years:
        #     build_tree_widget_item(self.year_picker, str(year))
        self.constituency_picker.takeChildren()

        constituencies = self.book.get_constituencies()[years[0]]
        for c in constituencies:
            build_tree_widget_item(self.constituency_picker, str(c))

    def get_selected_constituencies(self):
        cs = []

        if self.constituency_picker:
            num_children = self.constituency_picker.childCount()
            children = [self.constituency_picker.child(i) for i in range(num_children)]
            cs = [c.text(0) for c in children if c.checkState(0)]

        return cs

    def excel_button_clicked(self):
        filename = QtGui.QFileDialog.getOpenFileName(self, "Select Excel file", "", "Excel files (*.xlsx)")
        t = threading.Thread(target=self.load_excel, args=(filename,))
        t.start()

    def load_excel(self, filename):
        self.increment_counter(self.excel_button)
        self.book = ExcelBook()
        self.set_up_book()
        self.book.load(filename)
        if self.book.loaded:
            self.load_picker_data()
            self.excel_button.setText(self.book.name)
        self.decrement_counter(self.excel_button)

    def word_button_clicked(self):
        filename = QtGui.QFileDialog.getOpenFileName(self, "Select Word template", "", "Word files (*.docx)")
        t = threading.Thread(target=self.load_doc, args=(filename,))
        t.start()

    def load_doc(self, filename):
        self.increment_counter(self.word_button)
        self.doc.load(filename)
        if self.doc.loaded:
            self.word_button.setText(self.doc.name)
        self.decrement_counter(self.word_button)

    def do_generate_doc_work(self):
        constituencies_selected = self.get_selected_constituencies()

        if constituencies_selected:
            self.update_progress_bar_signal.emit("Writing docs...", 0)

            num_cs = len(constituencies_selected)
            for idx, c in enumerate(constituencies_selected):
                write_doc(c, self.book, self.doc)
                self.update_progress_bar_signal.emit(c, int((100.0 * (idx + 1)) / num_cs))

            self.update_progress_bar_signal.emit("Done!", 100)

        self.generate_doc_finished_signal.emit()

    def generate_doc(self):
        if not self.doc.loaded and not self.book.loaded:
            QtGui.QMessageBox.critical(self, "", "Choose an excel sheet and word template")
        elif not self.doc.loaded:
            QtGui.QMessageBox.critical(self, "", "Choose an excel sheet and word template")
        elif not self.book.loaded:
            QtGui.QMessageBox.critical(self, "", "Choose an excel sheet and word template")
        else:
            self.generate_button.setEnabled(False)
            self.excel_button.setEnabled(False)
            self.word_button.setEnabled(False)

            t = threading.Thread(target=self.do_generate_doc_work)
            t.start()

    def doc_generated(self):
        self.generate_button.setEnabled(True)
        self.excel_button.setEnabled(True)
        self.word_button.setEnabled(True)

    def increment_counter(self, src):
        self.lock.acquire()
        self.loading_counter += 1
        src.setEnabled(False)
        self.generate_button.setEnabled(False)
        self.lock.release()

    def decrement_counter(self, src):
        self.lock.acquire()
        self.loading_counter -= 1
        src.setEnabled(True)
        if self.loading_counter == 0:
            self.generate_button.setEnabled(True)
        self.lock.release()

    def doc_loaded(self):
        self.word_loading_movie.stop()
        self.word_button.setIcon(QtGui.QIcon(resource_path("word-icon.png")))
        self.update_excel_word_output()

    def book_loaded(self):
        self.excel_loading_movie.stop()
        self.excel_button.setIcon(QtGui.QIcon(resource_path("excel-icon.png")))
        if self.book.loaded:
            self.update_progress_bar("Excel loaded", 100)
        else:
            self.update_progress_bar("Excel load failed!", 0)
        dodgy_cells = self.book.get_dodgy_cells()
        empty_cells = self.book.get_empty_cells()

        if dodgy_cells:
            self.init_output_picker(self.dodgy_cell_picker, dodgy_cells)

        if empty_cells:
            self.init_output_picker(self.empty_cell_picker, empty_cells)

        self.update_excel_word_output()

    def init_output_picker(self, picker, data):
        picker.setDisabled(False)
        picker.takeChildren()
        for year in data:
            year_picker = build_tree_widget_item(picker, year, False)
            constituents = data[year]
            for constituent in constituents:
                constituent_picker = build_tree_widget_item(year_picker, constituent, False)
                columns = constituents[constituent]
                for col in columns:
                    build_tree_widget_item(constituent_picker, col.replace("\n", " "), False)

    def update_excel_word_output(self):
        doc_columns_not_in_excel = []

        if self.book.loaded and self.doc.loaded:
            doc_headers = self.doc.all_headers
            doc_columns_not_in_excel = [h for h in doc_headers if not self.book.column_exists(h)]

        if doc_columns_not_in_excel:
            self.missing_column_picker.setDisabled(False)
            self.missing_column_picker.takeChildren()
            for h in doc_columns_not_in_excel:
                build_tree_widget_item(self.missing_column_picker, h, False)

    def set_excel_icon(self, _):
        self.excel_button.setIcon(QtGui.QIcon(self.excel_loading_movie.currentPixmap()))

    def set_word_icon(self, _):
        self.word_button.setIcon(QtGui.QIcon(self.word_loading_movie.currentPixmap()))

    def update_progress_bar(self, key, val):
        self.progress_label.setText(key)
        if val > 100:
            val = 100
        self.progress_bar.setValue(val)


def quitting():
    pass


if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)

    window = MainWindow()
    app.connect(app, Qt.SIGNAL("aboutToQuit()"), quitting)
    sys.exit(app.exec_())
