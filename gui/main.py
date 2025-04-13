# ///////////////////////////////////////////////////////////////
#
# BY: WANDERSON M.PIMENTA
# PROJECT MADE WITH: Qt Designer and PySide6
# V: 1.0.0
#
# This project can be used freely for all uses, as long as they maintain the
# respective credits only in the Python scripts, any information in the visual
# interface (GUI) can be modified without any implication.
#
# There are limitations on Qt licenses if you want to use your products
# commercially, I recommend reading them on the official website:
# https://doc.qt.io/qtforpython/licenses.html
#
# ///////////////////////////////////////////////////////////////

import sys
import os
import platform
import time
from typing import Optional

import PySide6.QtCore

# Add parent directory to Python path
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from advanced_cracker import PasswordCracker, CrackingMode

# IMPORT / GUI AND MODULES AND WIDGETS
# ///////////////////////////////////////////////////////////////
from modules import *
from widgets import *
os.environ["QT_FONT_DPI"] = "96" # FIX Problem for High DPI and Scale above 100%

# SET AS GLOBAL WIDGETS
# ///////////////////////////////////////////////////////////////
widgets = None


from PySide6.QtWidgets import *
from PySide6.QtCore import Signal, Slot

class Worker(QThread):
    finished = Signal()
    intReady = Signal(int)
    @Slot()
    def procCounter(self):
        # box = QMessageBox()
        # box.setText("procCounter called")
        # box.exec()
        
        for i in range(1, 10):
            print("Worker started!")
            time.sleep(1)
            self.intReady.emit(i)

        self.finished.emit()

    # @Slot()  # QtCore.Slot
    # def run(self):
    #     '''
    #     Your code goes in this function
    #     '''
    #     print("Thread start")
    #     time.sleep(5)
    #     print("Thread complete")

class MainWindow(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)

        # SET AS GLOBAL WIDGETS
        # ///////////////////////////////////////////////////////////////
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        global widgets
        widgets = self.ui

        # USE CUSTOM TITLE BAR | USE AS "False" FOR MAC OR LINUX
        # ///////////////////////////////////////////////////////////////
        Settings.ENABLE_CUSTOM_TITLE_BAR = True

        # APP NAME
        # ///////////////////////////////////////////////////////////////
        title = "PyDracula - Modern GUI"
        description = "PyDracula APP - Theme with colors based on Dracula for Python."
        # APPLY TEXTS
        self.setWindowTitle(title)
        widgets.titleRightInfo.setText(description)

        # TOGGLE MENU
        # ///////////////////////////////////////////////////////////////
        widgets.toggleButton.clicked.connect(lambda: UIFunctions.toggleMenu(self, True))

        # SET UI DEFINITIONS
        # ///////////////////////////////////////////////////////////////
        UIFunctions.uiDefinitions(self)

        # QTableWidget PARAMETERS
        # ///////////////////////////////////////////////////////////////
        widgets.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # BUTTONS CLICK
        # ///////////////////////////////////////////////////////////////

        # LEFT MENUS
        widgets.btn_home.clicked.connect(self.buttonClick)
        widgets.btn_widgets.clicked.connect(self.buttonClick)
        
        # self.ui.pushButton.clicked.connect(self.pushButton_clicked)

        widgets.pushButton_openFile.clicked.connect(self.pushButton_openFile_clicked)
        widgets.pushButton_check.clicked.connect(self.pushButton_check_clicked)
        
        self.obj = Worker()
        self.thread = QThread()
        
        self.obj.moveToThread(self.thread)
        self.obj.intReady.connect(self.update_ready_ui)
        self.obj.finished.connect(self.update_finished_ui)
        self.thread.started.connect(self.obj.procCounter)
        

        # EXTRA LEFT BOX
        def openCloseLeftBox():
            UIFunctions.toggleLeftBox(self, True)
        # widgets.toggleLeftBox.clicked.connect(openCloseLeftBox)
        widgets.extraCloseColumnBtn.clicked.connect(openCloseLeftBox)

        # EXTRA RIGHT BOX
        def openCloseRightBox():
            UIFunctions.toggleRightBox(self, True)
        # widgets.settingsTopBtn.clicked.connect(openCloseRightBox)

        # SHOW APP
        # ///////////////////////////////////////////////////////////////
        self.show()

        # SET CUSTOM THEME
        # ///////////////////////////////////////////////////////////////
        useCustomTheme = False
        themeFile = "themes\py_dracula_light.qss"

        # SET THEME AND HACKS
        if useCustomTheme:
            # LOAD AND APPLY STYLE
            UIFunctions.theme(self, themeFile, True)

            # SET HACKS
            AppFunctions.setThemeHack(self)

        # SET HOME PAGE AND SELECT MENU
        # ///////////////////////////////////////////////////////////////
        widgets.stackedWidget.setCurrentWidget(widgets.home)
        widgets.btn_home.setStyleSheet(UIFunctions.selectMenu(widgets.btn_home.styleSheet()))

    def update_ready_ui(self, i):
        # box = QMessageBox()
        # box.setText("update_readd_ui called" + str(i))
        # box.exec()
        self.ui.plainTextEdit.appendPlainText(str(i) + "\n")
        # app.processEvents()
        
    @Slot()
    def update_finished_ui(self):
        self.ui.plainTextEdit.appendPlainText("Finished" + "\n")
        self.thread.quit()

    # BUTTONS CLICK
    # Post here your functions for clicked buttons
    # ///////////////////////////////////////////////////////////////
    def buttonClick(self):
        # GET BUTTON CLICKED
        btn = self.sender()
        btnName = btn.objectName()

        # SHOW HOME PAGE
        if btnName == "btn_home":
            widgets.stackedWidget.setCurrentWidget(widgets.home)
            UIFunctions.resetStyle(self, btnName)
            btn.setStyleSheet(UIFunctions.selectMenu(btn.styleSheet()))

        # SHOW WIDGETS PAGE
        if btnName == "btn_widgets":
            widgets.stackedWidget.setCurrentWidget(widgets.widgets)
            UIFunctions.resetStyle(self, btnName)
            btn.setStyleSheet(UIFunctions.selectMenu(btn.styleSheet()))

        # SHOW NEW PAGE
        if btnName == "btn_new":
            widgets.stackedWidget.setCurrentWidget(widgets.new_page) # SET PAGE
            UIFunctions.resetStyle(self, btnName) # RESET ANOTHERS BUTTONS SELECTED
            btn.setStyleSheet(UIFunctions.selectMenu(btn.styleSheet())) # SELECT MENU

        if btnName == "btn_save":
            print("Save BTN clicked!")

        # PRINT BTN NAME
        print(f'Button "{btnName}" pressed!')
        
    def pushButton_clicked(self):
        box = QMessageBox()
        self.thread.start()
        self.thread.stop()
        box.setText("线程已开启")
        box.exec()
        
    def pushButton_openFile_clicked(self):
        """
        Open file dialog to select a Word document
        """
        file_filter = "Word Document (*.docx *.doc);;All Files (*.*)"
        initial_dir = os.path.expanduser("~\Documents")
        
        file_path, _ = QFileDialog.getOpenFileName(
            parent=self,
            caption="Select Word Document",
            dir=initial_dir,
            filter=file_filter
        )
        
        if file_path:
            # Store the selected file path
            self.selected_file_path = file_path
            # Update UI to show selected file
            file_name = os.path.basename(file_path)
            self.ui.label.setText(f"Selected file: {file_name}")
            self.ui.plainTextEdit.appendPlainText(f"Loaded document: {file_path}\n")
            # Enable check button only after file is selected
            self.ui.pushButton_check.setEnabled(True)
            self.ui.lineEdit.setText(file_path)
        else:
            self.ui.plainTextEdit.appendPlainText("No file selected\n")
            self.ui.pushButton_check.setEnabled(False)

    def pushButton_check_clicked(self):
        file_to_crack = self.ui.lineEdit.text().strip()
        wordlist = "../wordlist.txt"
        mode = self.ui.comboBox.currentText().split(" ")[0].lower()
        threads = 4
        chunk_size = 1000
        timeout = 3600
        verify_hash = False

        try:
            cracker = PasswordCracker(
                file_to_crack,
                wordlist,
                mode=CrackingMode(mode),
                threads=threads,
                chunk_size=chunk_size,
                timeout=timeout,
                verify_hash=verify_hash
            )

            password, duration, stats = cracker.crack()

            if password:
                self.ui.plainTextEdit.appendPlainText(f"\nSuccess! Password found: {password}")
                self.ui.plainTextEdit.appendPlainText(f"Time taken: {duration:.2f} seconds")
                self.ui.plainTextEdit.appendPlainText(f"Attempts: {stats.attempts}")
            else:
                self.ui.plainTextEdit.appendPlainText("\nPassword not found")
                if stats.errors:
                    self.ui.plainTextEdit.appendPlainText("Errors encountered:")
                    for error in stats.errors:
                        self.ui.plainTextEdit.appendPlainText(f"- {error}")

        except Exception as e:
            self.ui.plainTextEdit.appendPlainText(f"Critical error: {str(e)}")
            # Don't exit the application on error, just show the error message
            # sys.exit(1)

    # RESIZE EVENTS
    # ///////////////////////////////////////////////////////////////
    def resizeEvent(self, event):
        # Update Size Grips
        UIFunctions.resize_grips(self)

    # MOUSE CLICK EVENTS
    # ///////////////////////////////////////////////////////////////
    def mousePressEvent(self, event):
        # SET DRAG POS WINDOW
        self.dragPos = event.globalPos()

        # PRINT MOUSE EVENTS
        if event.buttons() == Qt.LeftButton:
            print('Mouse click: LEFT CLICK')
        if event.buttons() == Qt.RightButton:
            print('Mouse click: RIGHT CLICK')

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("icon.ico"))
    window = MainWindow()
    sys.exit(app.exec_())
