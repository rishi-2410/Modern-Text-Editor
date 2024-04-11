import sys
from subprocess import call
from PyQt5 import QtWidgets
from PyQt5 import QtPrintSupport
from PyQt5 import QtGui, QtCore
from PyQt5.QtGui import QFontDatabase
from PyQt5.QtCore import Qt
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5 import uic
import time
import threading
import queue
import speech_recognition as sr
from datetime import date
from datetime import*
import pytz
from googletrans import Translator



class Main(QtWidgets.QMainWindow):

    def __init__(self,parent=None):
        QtWidgets.QMainWindow.__init__(self,parent)
        self.setStyleSheet(myStyleSheet(self))
        self.filename = ""

        self.changesSaved = True
        self.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\VCLOGO.png"))
        self.initUI()

    def initMenubar(self):

        menubar = self.menuBar()
        menubar.setFont(QFont("Arial Rounded MT Bold", 10))

        self.file = menubar.addMenu("File")
        self.edit = menubar.addMenu("Edit")
        self.edit_template_submenu = QMenu("Templates",self)
        self.edit_template_submenu.setIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\template.png"))
        self.insert = menubar.addMenu("Insert")
        self.view = menubar.addMenu("View")
        self.Help = menubar.addMenu("Help")

        # Add the most important actions to the menubar

        self.settingAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\setwinico.png"),"Settings",self)
        self.settingAction.setStatusTip("Settings")
        self.settingAction.triggered.connect(self.settingswin)

        self.file.addAction(self.newAction)
        self.file.addAction(self.openAction)
        self.file.addAction(self.saveAction)
        self.file.addSeparator()
        self.file.addAction(self.printAction)
        self.file.addAction(self.previewAction)
        self.file.addSeparator()
        self.file.addAction(self.settingAction)

        self.edit.addAction(self.undoAction)
        self.edit.addAction(self.redoAction)
        self.edit.addSeparator()
        self.edit.addAction(self.cutAction)
        self.edit.addAction(self.copyAction)
        self.edit.addAction(self.pasteAction)
        self.edit.addAction(self.selectallAction)
        self.edit.addSeparator()
        self.edit.addMenu(self.edit_template_submenu)
        self.Letter_temp_action = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\letter_template.png"),"Letter template")
        self.Letter_temp_action.triggered.connect(self.letter_temp_exec)
        self.job_app_form_temp_action = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\job_appln_form.png"),"Job application template")
        self.job_app_form_temp_action.triggered.connect(self.job_appn_form_temp_exec)
        self.edit_template_submenu.addAction(self.Letter_temp_action)
        self.edit_template_submenu.addAction(self.job_app_form_temp_action)

        self.insert.addAction(self.dateTimeAction)
        self.insert.addAction(self.tableAction)
        self.insert.addAction(self.imageAction)
        self.insert.addAction(self.symbolAction)

        # Toggling actions for the various bars
        toolbarAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\toolbar.png"),"Toggle Ribbon",self)
        toolbarAction.triggered.connect(self.toggleRibbon)

        formulabarAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\formulabar.png"),"Toggle Formulabar",self)
        formulabarAction.triggered.connect(self.toggleFormulabar)

        statusbarAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\statusbar.png"),"Toggle Statusbar",self)
        statusbarAction.triggered.connect(self.toggleStatusbar)

        self.view.addAction(toolbarAction)
        self.view.addAction(formulabarAction)
        self.view.addAction(statusbarAction)

        About = QtWidgets.QAction("About",self)
        About.triggered.connect(self.abtactionfunc)

        self.Help.addAction(About)

    def initToolbar(self):

        self.statusbar = self.statusBar()

        self.newAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\newdoc.png"),"New",self)
        self.newAction.setShortcut("Ctrl+N")
        self.newAction.setStatusTip("Create a new document")
        self.newAction.triggered.connect(self.new)

        self.openAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\opendoc.png"),"Open file",self)
        self.openAction.setStatusTip("Open existing document")
        self.openAction.setShortcut("Ctrl+O")
        self.openAction.triggered.connect(self.open)

        self.saveAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\savedoc.png"),"Save",self)
        self.saveAction.setStatusTip("Save document")
        self.saveAction.setShortcut("Ctrl+S")
        self.saveAction.triggered.connect(self.save)

        self.printAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\print.png"),"Print document",self)
        self.printAction.setStatusTip("Print document")
        self.printAction.setShortcut("Ctrl+P")
        self.printAction.triggered.connect(self.printHandler)

        self.previewAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\printpreview.png"),"Page view",self)
        self.previewAction.setStatusTip("Preview page before printing")
        self.previewAction.setShortcut("Ctrl+Shift+P")
        self.previewAction.triggered.connect(self.preview)

##        self.findAction = QtWidgets.QAction(QtGui.QIcon("icons/find.png"),"Find and replace",self)
##        self.findAction.setStatusTip("Find and replace words in your document")
##        self.findAction.setShortcut("Ctrl+F")

        self.cutAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\cut.png"),"Cut to clipboard",self)
        self.cutAction.setStatusTip("Delete and copy text to clipboard")
        self.cutAction.setShortcut("Ctrl+X")
        self.cutAction.triggered.connect(self.text.cut)

        self.copyAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\copy.png"),"Copy to clipboard",self)
        self.copyAction.setShortcut("Ctrl+C")
        self.copyAction.triggered.connect(self.text.copy)
        self.copyAction.triggered.connect(lambda:self.statusbar.showMessage("Text copied to clipboard",2000))

        self.pasteAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\paste.png"),"Paste from clipboard",self)
        self.pasteAction.setStatusTip("Paste text from clipboard")
        self.pasteAction.setShortcut("Ctrl+V")
        self.pasteAction.triggered.connect(self.text.paste)

        self.selectallAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\selectall.png"),"Select all the text",self)
        self.selectallAction.setStatusTip("Select all the text")
        self.selectallAction.setShortcut("Ctrl+A")
        self.selectallAction.triggered.connect(self.text.selectAll)

        self.undoAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\undo.png"),"Undo last action",self)
        self.undoAction.setStatusTip("Undo last action")
        self.undoAction.setShortcut("Ctrl+Z")
        self.undoAction.triggered.connect(self.text.undo)

        self.redoAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\redo.png"),"Redo last undone thing",self)
        self.redoAction.setStatusTip("Redo last undone thing")
        self.redoAction.setShortcut("Ctrl+Y")
        self.redoAction.triggered.connect(self.text.redo)

        self.templateAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\template.png"),"Templates",self)
        self.templateAction.setStatusTip("Explore Templates")
        self.templateAction.triggered.connect(self.template_Dialog)

        self.translateAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\translate.png"),"Translate",self)
        self.translateAction.setStatusTip("Explore Templates")
        self.translateAction.triggered.connect(self.translate_Dialog)
        

        self.toolbar = self.addToolBar("Options")
        self.toolbar.setFloatable(False)
        self.toolbar.setMovable(False)
        self.toolbar.setIconSize(QSize(20,20))

        self.toolbar.addAction(self.newAction)
        self.toolbar.addAction(self.openAction)
        self.toolbar.addAction(self.saveAction)

        self.toolbar.addSeparator()

        self.toolbar.addAction(self.printAction)
        self.toolbar.addAction(self.previewAction)

        self.toolbar.addSeparator()

        self.toolbar.addAction(self.cutAction)
        self.toolbar.addAction(self.copyAction)
        self.toolbar.addAction(self.pasteAction)
        self.toolbar.addAction(self.selectallAction)
        self.toolbar.addAction(self.undoAction)
        self.toolbar.addAction(self.redoAction)

        self.toolbar.addSeparator()

        self.toolbar.addAction(self.templateAction)

        self.toolbar.addSeparator()

        self.toolbar.addAction(self.translateAction)

        shadow_effect = QGraphicsDropShadowEffect()
        shadow_effect.setBlurRadius(10)
        shadow_effect.setColor(QtGui.QColor(136, 136, 136))
        shadow_effect.setXOffset(2)
        shadow_effect.setYOffset(2)
        self.toolbar.setGraphicsEffect(shadow_effect)

    
        self.toolbar.setStyleSheet("""background:white;
border-top-right-radius:7px;
border-top-left-radius:7px;
margin-top:1px;
margin-left:7px;
margin-right:7px;
padding:2px;""")

        self.addToolBarBreak()

    def translate_Dialog(self):
        translate_dialog = QDialog()
        translate_dialog.setWindowTitle("Language Translator")
        translate_dialog.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\translate.png"))
        translate_dialog.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint,False)
        layout = QVBoxLayout(translate_dialog)

        source_language = QComboBox(translate_dialog)
        target_language = QComboBox(translate_dialog)
        source_language.addItems(["Afrikaans", "Albanian", "Amharic", "Arabic", "Armenian", "Azerbaijani", "Basque", "Belarusian", "Bengali", "Bosnian", "Bulgarian", "Catalan", "Cebuano", "Chichewa", "Chinese (Simplified)", "Chinese (Traditional)", "Corsican", "Croatian", "Czech", "Danish", "Dutch", "English", "Esperanto", "Estonian", "Filipino", "Finnish", "French", "Frisian", "Galician", "Georgian", "German", "Greek", "Gujarati", "Haitian Creole", "Hausa", "Hawaiian", "Hebrew", "Hindi", "Hmong", "Hungarian", "Icelandic", "Igbo", "Indonesian", "Irish", "Italian", "Japanese", "Javanese", "Kannada", "Kazakh", "Khmer", "Korean", "Kurdish (Kurmanji)", "Kyrgyz", "Lao", "Latin", "Latvian", "Lithuanian", "Luxembourgish", "Macedonian", "Malagasy", "Malay", "Malayalam", "Maltese", "Maori", "Marathi", "Mongolian", "Myanmar (Burmese)", "Nepali", "Norwegian", "Pashto", "Persian", "Polish", "Portuguese", "Punjabi", "Romanian", "Russian", "Samoan", "Scots Gaelic", "Serbian", "Sesotho", "Shona", "Sindhi", "Sinhala", "Slovak", "Slovenian", "Somali", "Spanish", "Sundanese", "Swahili", "Swedish", "Tajik", "Tamil", "Telugu", "Thai", "Turkish", "Ukrainian", "Urdu", "Uzbek", "Vietnamese", "Welsh", "Xhosa", "Yiddish", "Yoruba", "Zulu"])
        source_language.setCurrentIndex(21)
        layout.addWidget(source_language)

        # Input text field
        input_text_box = QTextEdit(translate_dialog)
        layout.addWidget(input_text_box)

        # Target language combobox
        target_language.addItems(["Afrikaans", "Albanian", "Amharic", "Arabic", "Armenian", "Azerbaijani", "Basque", "Belarusian", "Bengali", "Bosnian", "Bulgarian", "Catalan", "Cebuano", "Chichewa", "Chinese (Simplified)", "Chinese (Traditional)", "Corsican", "Croatian", "Czech", "Danish", "Dutch", "English", "Esperanto", "Estonian", "Filipino", "Finnish", "French", "Frisian", "Galician", "Georgian", "German", "Greek", "Gujarati", "Haitian Creole", "Hausa", "Hawaiian", "Hebrew", "Hindi", "Hmong", "Hungarian", "Icelandic", "Igbo", "Indonesian", "Irish", "Italian", "Japanese", "Javanese", "Kannada", "Kazakh", "Khmer", "Korean", "Kurdish (Kurmanji)", "Kyrgyz", "Lao", "Latin", "Latvian", "Lithuanian", "Luxembourgish", "Macedonian", "Malagasy", "Malay", "Malayalam", "Maltese", "Maori", "Marathi", "Mongolian", "Myanmar (Burmese)", "Nepali", "Norwegian", "Pashto", "Persian", "Polish", "Portuguese", "Punjabi", "Romanian", "Russian", "Samoan", "Scots Gaelic", "Serbian", "Sesotho", "Shona", "Sindhi", "Sinhala", "Slovak", "Slovenian", "Somali", "Spanish", "Sundanese", "Swahili", "Swedish", "Tajik", "Tamil", "Telugu", "Thai", "Turkish", "Ukrainian", "Urdu", "Uzbek", "Vietnamese", "Welsh", "Xhosa", "Yiddish", "Yoruba", "Zulu"])
        target_language.setCurrentIndex(21)
        layout.addWidget(target_language)

        # Text area to show the translated text
        translated_text_box = QTextEdit(translate_dialog)
        translated_text_box.setReadOnly(True)
        layout.addWidget(translated_text_box)

        # Button to trigger translation
        translate_button = QPushButton("Translate", translate_dialog)
        translate_button.clicked.connect(lambda: self.translate_text(input_text_box, source_language, target_language, translated_text_box))
        layout.addWidget(translate_button)

        # Translate Selection Button
        translate_selection_button = QPushButton("Translate Selection", translate_dialog)
        translate_selection_button.clicked.connect(lambda: self.translate_selection(input_text_box, source_language, target_language, translated_text_box))
        layout.addWidget(translate_selection_button)

        translate_dialog.setLayout(layout)
        translate_dialog.exec_()

    def translate_text(self,input_text_box, source_language, target_language, translated_text_box):
        # Get the input text and source/target language codes    

        input_text = input_text_box.toPlainText()
        source_lang = source_language.currentText()
        target_lang = target_language.currentText()

        # Translate the text
        translator = Translator()
        translated_text = translator.translate(input_text, src=source_lang, dest=target_lang).text

        # Display the translated text
        translated_text_box.setText(translated_text)

    def translate_selection(self,input_text_box, source_language, target_language, translated_text_box):
        # Get the selected text
        selected_text = self.text.textCursor().selectedText()
        if not selected_text:
            return  # No text selected, do nothing
        else:
            input_text_box.setText(selected_text)

        # Get the source/target language codes
        source_lang = source_language.currentText()
        target_lang = target_language.currentText()

        # Translate the selected text
        translator = Translator()
        translated_text = translator.translate(selected_text, src=source_lang, dest=target_lang).text

        # Insert the translated text
        cursor = translated_text_box.textCursor()
        cursor.insertText(translated_text)
        cursor_sel = self.text.textCursor()
        cursor_sel.insertText(translated_text)


    def template_Dialog(self):
    # Create a QDialog instance
        self.Template_dialog = QDialog()
        self.Template_dialog.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint,False)
        self.Template_dialog.setWindowTitle("Script - Templates")
        self.Template_dialog.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\template.png"))
        
        # Create four buttons
        Letter_temp = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\letter_template.png"),"")
        Letter_temp.setIconSize(QSize(200,200))
        Letter_temp.clicked.connect(self.letter_temp_exec)
        job_app_form_temp = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\job_appln_form.png"),"")
        job_app_form_temp.setIconSize(QSize(200,200))
        job_app_form_temp.clicked.connect(self.job_appn_form_temp_exec)
        button3 = QPushButton('Button 3')
        button4 = QPushButton('Button 4')
        
        # Create a grid layout
        grid_layout = QGridLayout()
        
        # Add buttons to the grid layout
        grid_layout.addWidget(Letter_temp, 0, 0)
        grid_layout.addWidget(job_app_form_temp, 0, 1)
        grid_layout.addWidget(button3, 1, 0)
        grid_layout.addWidget(button4, 1, 1)
        
        # Set the layout of the QDialog to the grid layout
        self.Template_dialog.setLayout(grid_layout)
        
        # Show the QDialog
        self.Template_dialog.exec_()

    def letter_temp_exec(self):
        self.text.textCursor().insertText("""
[Date]

From: [Sender's Name]
[Title/Address/City,State,Zip]

To: [Recipient's Name]
[Title/Address/City,State,Zip]

Dear/Respected [Recipient's Name]

<introduction>

<body>

<conclusion>

Yours
Sincerely/Faithfull,

[Sender's Name]
[Title]
            """)

    def job_appn_form_temp_exec(self):
        self.text.textCursor().insertText("""
[Company Logo]

[Company Name]
[Address]
[City, State, Zip Code]
[Phone Number]
[Email Address]

Job Application Form

Personal Information:

Full Name:
Address:
City:
State:
Zip Code:
Phone Number:
Email Address:
Date of Birth:
Social Security Number (optional):

Position Information:

Position Applied For:
Date Available to Start:
Desired Salary:
Are you legally eligible to work in the United States? [Yes/No]
Are you over the age of 18? [Yes/No]

Education:

High School:
Dates Attended:
Degree:
GPA:

College/University:
Dates Attended:
Degree:
GPA:

Other Education (if applicable):
Dates Attended:
Degree:
GPA:

Employment History:

Most Recent Employer:
Dates of Employment:
Position:
Duties/Responsibilities:
Reason for Leaving:

Previous Employer:
Dates of Employment:
Position:
Duties/Responsibilities:
Reason for Leaving:

References:

Reference 1:
Name:
Position:
Company:
Phone Number:
Email Address:

Reference 2:
Name:
Position:
Company:
Phone Number:
Email Address:

Reference 3:
Name:
Position:
Company:
Phone Number:
Email Address:

Cover Letter:

Please attach your cover letter here.
Attach your resume here.
            """)


    def calendershow(self):

        self.calwin = QDialog(self)
        self.calwin.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\datetime.png"))
        self.calwin.setWindowTitle("Script - Set Date")
        self.calwin.setFixedSize(700, 280)
        self.calwin.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint,False)
        

        main_layout = QHBoxLayout(self.calwin)

        self.calendar = QCalendarWidget(self.calwin)
        main_layout.addWidget(self.calendar)

        self.scroll = QScrollArea(self.calwin)
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setWidgetResizable(True)
        main_layout.addWidget(self.scroll)

        self.widget = QWidget(self.calwin)
        self.scroll.setWidget(self.widget)

        layout = QVBoxLayout(self.widget)

        self.add_button(date.today().strftime("%A, %d %B %Y"), lambda: self.text.textCursor().insertText(date.today().strftime("%A, %d %B %Y")))
        self.add_button(date.today().strftime("%a, %d %B %Y %p"), lambda: self.text.textCursor().insertText(date.today().strftime("%a, %d %B %Y %p")))
        self.add_button(QDateTime.currentDateTime().toString("hh:mm:ss"), lambda: self.text.textCursor().insertText(QDateTime.currentDateTime().toString("hh:mm:ss")))
        self.add_button(QDateTime.currentDateTime().toString("hh:mm:ss AP"), lambda: self.text.textCursor().insertText(QDateTime.currentDateTime().toString("hh:mm:ss AP")))
        self.add_button(QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss"), lambda: self.text.textCursor().insertText(QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss")))
        self.add_button(QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss AP"), lambda: self.text.textCursor().insertText(QDateTime.currentDateTime().toString("yyyy-MM-dd hh:mm:ss AP")))
        self.add_button(f"{date.today().strftime('%A, %d %B %Y')}, {QDateTime.currentDateTime().toString('hh:mm:ss AP')}", lambda: self.text.textCursor().insertText(f"{date.today().strftime('%A, %d %B %Y')}, {QDateTime.currentDateTime().toString('hh:mm:ss AP')}"))
        self.add_button(f"{date.today().strftime('%a, %d %B %Y')}, {QDateTime.currentDateTime().toString('hh:mm:ss AP')}", lambda: self.text.textCursor().insertText(f"{date.today().strftime('%a, %d %B %Y')}, {QDateTime.currentDateTime().toString('hh:mm:ss AP')}"))

        self.calendar.clicked.connect(self.date_selected)


        self.calwin.show()

    def add_button(self, text, action):
        btn = QPushButton(text, self)
        btn.setFont(QFont("Arial Rounded MT Bold", 10))
        btn.clicked.connect(action)
        layout = self.scroll.widget().layout()
        layout.addWidget(btn)

    def date_selected(self, date):
        selected_date = date.toString(Qt.ISODate)
        self.text.textCursor().insertText(selected_date)


    def showdate(self, qDate):
        cursor = self.text.textCursor()
        cursor.insertText('{0}/{1}/{2}'.format(qDate.day(), qDate.month(), qDate.year()))

    def textcopynotify(self):
        pass

    def initInsertbar(self):

        self.insertbar = self.addToolBar("Insert")
        self.insertbar.setIconSize(QSize(20,20))
        self.insertbar.setFloatable(False)
        self.insertbar.setMovable(False)

        self.dateTimeAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\datetime.png"),"Insert current date/time",self)
        self.dateTimeAction.setStatusTip("Insert current date/time")
        self.dateTimeAction.setShortcut("Ctrl+D")
        self.dateTimeAction.triggered.connect(self.calendershow)

##        wordCountAction = QtWidgets.QAction(QtGui.QIcon("icons/count.png"),"See word/symbol count",self)
##        wordCountAction.setStatusTip("See word/symbol count")
##        wordCountAction.setShortcut("Ctrl+W")
##        wordCountAction.triggered.connect(self.wordCount)

        self.tableAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\table.png"),"Insert table",self)
        self.tableAction.setStatusTip("Insert table")
        self.tableAction.setShortcut("Ctrl+T")
        self.tableAction.triggered.connect(self.tableDialog)

        self.imageAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\add_image.png"),"Insert image",self)
        self.imageAction.setStatusTip("Insert image")
        self.imageAction.setShortcut("Ctrl+Shift+I")
        self.imageAction.triggered.connect(self.insertImage)

        self.symbolAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\symico.png"),"Symbols",self)
        self.symbolAction.triggered.connect(self.symbol_win)

        self.insertbar.addAction(self.dateTimeAction)
        self.insertbar.addAction(self.tableAction)
        self.insertbar.addAction(self.imageAction)

        self.insertbar.addSeparator()

        self.insertbar.addAction(self.symbolAction)

        shadow_effect = QGraphicsDropShadowEffect()
        shadow_effect.setBlurRadius(10)
        shadow_effect.setColor(QtGui.QColor(136, 136, 136))
        shadow_effect.setXOffset(2)
        shadow_effect.setYOffset(2)
        self.insertbar.setGraphicsEffect(shadow_effect)

    
        self.insertbar.setStyleSheet("""background:white;
border-bottom-right-radius:7px;
border-bottom-left-radius:7px;
margin-left:7px;
margin-right:7px;
border:;
padding:2px;""")

        self.addToolBarBreak()

    def tableDialog(self):
        self.dialog = QDialog()
        self.dialog.setWindowTitle('Insert Table')
        self.dialog.setWindowIcon(QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\table.png"))
        self.dialog.setWindowFlags(self.dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self.dialog.setStyleSheet('background-color: white; font-family: Arial Rounded MT Bold;')

        self.layout = QVBoxLayout(self.dialog)

        self.rows_layout = QHBoxLayout()
        self.rows_label = QLabel('Number of Rows:', self.dialog)
        self.rows_label.setFont(QFont("Arial Rounded MT Bold", 10))
        self.rows_edit = QLineEdit(self.dialog)
        self.rows_edit.setFont(QFont("Arial Rounded MT Bold", 10))
        self.rows_edit.setFixedWidth(150)
        self.rows_layout.addWidget(self.rows_label)
        self.rows_layout.addWidget(self.rows_edit)

        self.columns_layout = QHBoxLayout()
        self.columns_label = QLabel('Number of Columns:', self.dialog)
        self.columns_label.setFont(QFont("Arial Rounded MT Bold", 10))
        self.columns_edit = QLineEdit(self.dialog)
        self.columns_edit.setFont(QFont("Arial Rounded MT Bold", 10))
        self.columns_edit.setFixedWidth(150)
        self.columns_layout.addWidget(self.columns_label)
        self.columns_layout.addWidget(self.columns_edit)

        self.border_layout = QHBoxLayout()
        self.border_label = QLabel('Border Type:', self.dialog)
        self.border_label.setFont(QFont("Arial Rounded MT Bold", 10))
        self.border_combobox = QComboBox(self.dialog)
        self.border_combobox.setFont(QFont("Arial Rounded MT Bold", 10))
        self.border_combobox.setFixedWidth(150)
        self.border_combobox.addItems(["Thick", "Thin"])
        self.border_layout.addWidget(self.border_label)
        self.border_layout.addWidget(self.border_combobox)

        self.button_layout = QHBoxLayout()
        self.insert_button = QPushButton(QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\tableinsertico.png"),'', self.dialog)
        self.insert_button.setFixedWidth(25)
        self.insert_button.setFixedHeight(25)
        self.insert_button.setIconSize(QSize(25,25))
        self.insert_button.clicked.connect(self.insertTable)
        self.button_layout.addWidget(self.insert_button)

        self.layout.addLayout(self.rows_layout)
        self.layout.addLayout(self.columns_layout)
        self.layout.addLayout(self.border_layout)
        self.layout.addLayout(self.button_layout)

        self.dialog.setLayout(self.layout)
        self.dialog.exec_()

    def insertTable(self):
        rows_text = self.rows_edit.text().strip()
        columns_text = self.columns_edit.text().strip()
        border_type = self.border_combobox.currentText()

        if not rows_text or not columns_text:
            QMessageBox.warning(self.dialog, 'Warning', 'Please enter the number of rows and columns.')
            return

        if not rows_text.isdigit() or not columns_text.isdigit():
            QMessageBox.warning(self.dialog, 'Warning', 'Invalid input for rows or columns. Please enter numerical values.')
            return

        if border_type == 'Thick':
            self.insertTablethick()
        elif border_type == 'Thin':
            self.insertTablethin()

    def insertTablethick(self):
        rows_text = self.rows_edit.text().strip()
        columns_text = self.columns_edit.text().strip()

        if not rows_text or not columns_text:
            QMessageBox.warning(self.dialog, 'Warning', 'Please enter the number of rows and columns.')
            return

        if not rows_text.isdigit() or not columns_text.isdigit():
            QMessageBox.warning(self.dialog, 'Warning', 'Invalid input for rows or columns.')
            return

        cursor = self.text.textCursor()
        rows = int(rows_text)
        columns = int(columns_text)
        cursor.insertTable(rows, columns)
        self.dialog.close()

    def insertTablethin(self):
        rows_text = self.rows_edit.text().strip()
        columns_text = self.columns_edit.text().strip()

        if not rows_text or not columns_text:
            QMessageBox.warning(self.dialog, 'Warning', 'Please enter the number of rows and columns.')
            return

        if not rows_text.isdigit() or not columns_text.isdigit():
            QMessageBox.warning(self.dialog, 'Warning', 'Invalid input for rows or columns. Please enter numerical values.')
            return

        rows = int(rows_text)
        columns = int(columns_text)

        # Generate HTML string for the table with specified minimum widths for all cells
        html_table = "<table border='1' style='border-collapse: collapse;'>"
        cell_min_width = 50  # Adjust this value as needed
        for _ in range(rows):
            html_table += "<tr>"
            for _ in range(columns):
                html_table += f"<td width='{cell_min_width}'></td>"
            html_table += "</tr>"
        html_table += "</table>"

        # Insert HTML string into QTextEdit
        cursor = self.text.textCursor()
        cursor.insertHtml(html_table)

        self.dialog.close()

    def initFormatbar(self):

        self.formatbar = self.addToolBar("Format")
        self.formatbar.setIconSize(QSize(20,20))
        self.formatbar.setFloatable(False)
        self.formatbar.setMovable(False)

        fontBox = QtWidgets.QFontComboBox(self)
        fontBox.currentFontChanged.connect(lambda font: self.text.setCurrentFont(font))
        fontBox.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        fontBox.setMaximumSize(180, 28)
        fontBox.setStyleSheet("""font-size: 15px;
margin-left:10px;
            """)

        fontSize = QtWidgets.QComboBox(self)
        fontSize.setEditable(True)
        fontSize.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        fontSize.setMaximumSize(70, 28)
        fontSize.setStyleSheet("""font-size: 15px;
margin-left:3px;
            """)
        # Add items to the combo box
        for size in range(8, 100, 2):
            fontSize.addItem(f"{size} pt")

        # Connect the signal
        fontSize.currentIndexChanged.connect(self.setFontSize)

        # Set initial value
        fontSize.setCurrentIndex(1)

        fontColor = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\fontcolor.png"),"Change font color",self)
        fontColor.triggered.connect(self.fontColorChanged)

        bgAct = QtWidgets.QAction("change Background Color",self,  triggered=self.changeBGColor)
        bgAct.setStatusTip("change Background Color")
        bgAct.setIcon(QIcon('C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\pgbgcolor.png'))

        boldAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\bold.png"),"Bold",self)
        boldAction.triggered.connect(self.bold)

        italicAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\italic.png"),"Italic",self)
        italicAction.triggered.connect(self.italic)

        underlAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\underline.png"),"Underline",self)
        underlAction.triggered.connect(self.underline)

        strikeAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\strikeout.png"),"Strike-out",self)
        strikeAction.triggered.connect(self.strike)

        superAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\superscript.png"),"Superscript",self)
        superAction.triggered.connect(self.superScript)

        subAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\subscript.png"),"Subscript",self)
        subAction.triggered.connect(self.subScript)

        self.comboStyle = QComboBox(self.formatbar)
        index = 0
        self.comboStyle.setEditable(True)
        icon = QIcon('C:\\Users\\rishi\\OneDrive\\Desktop\\Vidwo Worsksuit SCRIPT\\ICON\\bulletlist.png')
        self.comboStyle.setItemIcon(0 , QIcon('C:\\Users\\rishi\\OneDrive\\Desktop\\Vidwo Worsksuit SCRIPT\\ICON\\bulletlist.png'))
        self.comboStyle.addItem("⚫")
        self.comboStyle.addItem("⚪")
        self.comboStyle.addItem("■")
        self.comboStyle.addItem("1.")
        self.comboStyle.addItem("a.")
        self.comboStyle.addItem("A.")
        self.comboStyle.addItem("i")
        self.comboStyle.addItem("I.")
        self.comboStyle.setStyleSheet("""
font-size:17px;
""")
        self.comboStyle.activated.connect(self.textStyle)

        alignLeft = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\left.png"),"Align left",self)
        alignLeft.triggered.connect(self.alignLeft)

        alignCenter = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\center.png"),"Align center",self)
        alignCenter.triggered.connect(self.alignCenter)

        alignRight = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\right.png"),"Align right",self)
        alignRight.triggered.connect(self.alignRight)

        alignJustify = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\justify.png"),"Align justify",self)
        alignJustify.triggered.connect(self.alignJustify)

        indentAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\indent.png"),"Indent Area",self)
        indentAction.setShortcut("Ctrl+Tab")
        indentAction.triggered.connect(self.indent)

        dedentAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\dedent.png"),"Dedent Area",self)
        dedentAction.setShortcut("Shift+Tab")
        dedentAction.triggered.connect(self.dedent)

        backColor = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\bgcolor.png"),"Change background color",self)
        backColor.triggered.connect(self.highlight)


        self.formatbar.addWidget(fontBox)
        self.formatbar.addWidget(fontSize)

        self.formatbar.addSeparator()

        self.formatbar.addAction(fontColor)
        self.formatbar.addAction(backColor)
        self.formatbar.addAction(bgAct)

        self.formatbar.addSeparator()

        self.formatbar.addAction(boldAction)
        self.formatbar.addAction(italicAction)
        self.formatbar.addAction(underlAction)
        self.formatbar.addAction(strikeAction)
        self.formatbar.addAction(superAction)
        self.formatbar.addAction(subAction)

        self.formatbar.addSeparator()

        self.formatbar.addWidget(self.comboStyle)

        self.formatbar.addSeparator()

        self.formatbar.addAction(alignLeft)
        self.formatbar.addAction(alignCenter)
        self.formatbar.addAction(alignRight)
        self.formatbar.addAction(alignJustify)

        shadow_effect = QGraphicsDropShadowEffect()
        shadow_effect.setBlurRadius(10)
        shadow_effect.setColor(QtGui.QColor(136, 136, 136))
        shadow_effect.setXOffset(2)
        shadow_effect.setYOffset(2)
        self.formatbar.setGraphicsEffect(shadow_effect)

        self.formatbar.setStyleSheet("""
background:white;
margin-left:7px;
margin-right:7px;
padding:2px;
border:0px;
QToolBar QToolButton:hover {
                background-color: lightgrey;
            }
""")

        self.formatbar.addSeparator()


        self.formatbar.addAction(indentAction)
        self.formatbar.addAction(dedentAction)

        self.addToolBarBreak()

    def editHTML(self):
        rtext = """<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0//EN" "http://www.w3.org/TR/REC-html40/strict.dtd">
<html><head><meta name="qrichtext" content="1" /><style type="text/css">
p, li { white-space: pre-wrap; }
</style></head><body>""" + '\n'
        btext = """<!--EndFragment--></p></body></html>"""
        all = self.editor.textCursor().selection().toHtml()
        #dlg = QInputDialog(self, Qt.Window)

        #dlg.setOption(QInputDialog.UsePlainTextEditForTextInput, True)
        #new, ok = dlg.getMultiLineText(self, 'change HTML', "edit HTML", all.replace(rtext, ""))
        #if ok:
        #    self.editor.textCursor().insertHtml(new)
        #    self.statusBar().showMessage("HTML changed")
        #else:
        #    self.statusBar().showMessage("HTML not changed")
        
        self.heditor = htmlEditor()
        self.heditor.ed.setPlainText(all.replace(rtext, "").replace(btext, ""))
        self.heditor.setGeometry(0, 0, 800, 600)
        self.heditor.show()

    def changeBGColor(self):
        all = self.text.document().toHtml()
        bgcolor = all.partition("<body style=")[2].partition(">")[0].partition('bgcolor="')[2].partition('"')[0]
        if not bgcolor == "":
            col = QColorDialog.getColor(QColor(bgcolor), self)
            if not col.isValid():
                return
            else:
                colorname = col.name()
                new = all.replace("bgcolor=" + '"' + bgcolor + '"', "bgcolor=" + '"' + colorname + '"')
                self.text.document().setHtml(new)
        else:
            col = QColorDialog.getColor(QColor("#FFFFFF"), self)
            if not col.isValid():
                return
            else:
                all = self.text.document().toHtml()
                body = all.partition("<body style=")[2].partition(">")[0]
                newbody = body + "bgcolor=" + '"' + col.name() + '"'
                new = all.replace(body, newbody)    
                self.text.document().setHtml(new)

    def textStyle(self, styleIndex):
        cursor = self.text.textCursor()
        if styleIndex:
            styleDict = {
                1: QTextListFormat.ListDisc,
                2: QTextListFormat.ListCircle,
                3: QTextListFormat.ListSquare,
                4: QTextListFormat.ListDecimal,
                5: QTextListFormat.ListLowerAlpha,
                6: QTextListFormat.ListUpperAlpha,
                7: QTextListFormat.ListLowerRoman,
                8: QTextListFormat.ListUpperRoman,
                        }

            style = styleDict.get(styleIndex, QTextListFormat.ListDisc)
            cursor.beginEditBlock()
            blockFmt = cursor.blockFormat()
            listFmt = QTextListFormat()

            if cursor.currentList():
                listFmt = cursor.currentList().format()
            else:
                listFmt.setIndent(1)
                blockFmt.setIndent(0)
                cursor.setBlockFormat(blockFmt)

            listFmt.setStyle(style)
            cursor.createList(listFmt)
            cursor.endEditBlock()
        else:
            bfmt = QTextBlockFormat()
            bfmt.setObjectIndex(-1)
            cursor.mergeBlockFormat(bfmt)

    def setFontSize(self, index):
        # Get the QComboBox object that emitted the signal
        combo_box = self.sender()
        # Get the selected item text from the QComboBox
        item_text = combo_box.currentText()
        # Extract the font size from the text (assuming it's always at the beginning followed by a space)
        font_size = int(item_text.split(' ')[0])
        # Set the font size
        self.text.setFontPointSize(font_size)

    def initFormulabar(self):
    # Add a new toolbar
        self.addToolBarBreak() 
        self.Formulabar = self.addToolBar("Formula bar")
        self.Formulabar.setStyleSheet("""background:white;
                                        border-radius:7px;
                                        margin-top:4px;
                                        margin-bottom:4px;
                                        margin-left:7px;
                                        margin-right:7px;
                                        padding:0px;""")

        shadow_effect = QGraphicsDropShadowEffect()
        shadow_effect.setBlurRadius(10)
        shadow_effect.setColor(QtGui.QColor(136, 136, 136))
        shadow_effect.setXOffset(2)
        shadow_effect.setYOffset(2)
        self.Formulabar.setGraphicsEffect(shadow_effect)

        # Create a QLineEdit widget
        self.formula_line = QtWidgets.QLineEdit()
        self.formula_line.setClearButtonEnabled(True)
        self.formula_line.setStyleSheet("""
font-family: 'Courier New', monospace;
background:lightgrey;
border-radius:5px;
font-size:17px;
padding:2px;
font
            """)

        code = ["RAND()","$C","TEMP.LETTER()","TEMP.FORM()"]
        completer = QCompleter(code)
        completer.setCaseSensitivity(Qt.CaseSensitive)
        self.formula_line.setCompleter(completer)
        self.Formulabar.setVisible(False)
        self.Formulabar.addWidget(self.formula_line)

        # Create a QPushButton widget
        self.run = QPushButton("")
        self.run.setShortcut("Return")
        self.run.setIconSize(QSize(25,25))
        self.run.setIcon(QIcon('C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\run.png'))
        self.run.clicked.connect(self.run_formula)
        self.run.setStyleSheet("""border-radius:4px;""")
        self.Formulabar.addWidget(self.run)

        self.addToolBarBreak()    

    def run_formula(self):
##        formula = self.Formulabar.text()
        if self.formula_line.text() == "RAND()":
            self.text.textCursor().insertText("""Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. At varius vel pharetra vel turpis nunc. Purus faucibus ornare suspendisse sed nisi lacus sed. Nunc sed id semper risus in. Id aliquet risus feugiat in ante. Tristique et egestas quis ipsum suspendisse ultrices gravida. Augue neque gravida in fermentum et sollicitudin ac orci phasellus. Suspendisse potenti nullam ac tortor vitae purus faucibus ornare suspendisse. Elementum pulvinar etiam non quam lacus suspendisse faucibus interdum. Aenean sed adipiscing diam donec adipiscing. Mi eget mauris pharetra et. A condimentum vitae sapien pellentesque habitant morbi. Ullamcorper a lacus vestibulum sed. Rhoncus dolor purus non enim praesent elementum. Dictumst vestibulum rhoncus est pellentesque elit ullamcorper dignissim. Arcu ac tortor dignissim convallis aenean et tortor at. Pellentesque sit amet porttitor eget dolor morbi non arcu.

Amet massa vitae tortor condimentum lacinia quis. Mattis ullamcorper velit sed ullamcorper morbi tincidunt ornare. Arcu vitae elementum curabitur vitae nunc sed velit dignissim. Vestibulum morbi blandit cursus risus at ultrices. Purus in massa tempor nec feugiat nisl. Dictumst quisque sagittis purus sit amet volutpat. Arcu cursus vitae congue mauris rhoncus aenean. Massa placerat duis ultricies lacus sed turpis tincidunt id aliquet. Maecenas accumsan lacus vel facilisis volutpat est velit egestas dui. Netus et malesuada fames ac turpis egestas. Amet facilisis magna etiam tempor orci. Iaculis urna id volutpat lacus laoreet non curabitur. Eu ultrices vitae auctor eu augue ut. Semper viverra nam libero justo laoreet sit amet cursus. Potenti nullam ac tortor vitae purus faucibus. Ridiculus mus mauris vitae ultricies. Ut morbi tincidunt augue interdum velit euismod in. Ipsum suspendisse ultrices gravida dictum. Vitae auctor eu augue ut. Morbi enim nunc faucibus a pellentesque sit amet porttitor eget.""")
        elif self.formula_line.text() == "$C":
            self.text.setText("")

        elif self.formula_line.text() == "TEMP.LETTER()":
            self.letter_temp_exec()

        elif self.formula_line.text() == "TEMP.FORM(JBAPPN)":
            self.job_appn_form_temp_exec()

    def initUI(self):

        self.text = QtWidgets.QTextEdit(self)
        self.text.setAutoFormatting(QtWidgets.QTextEdit.AutoAll)
        self.text.setCursorWidth(1)
        self.text.setFixedWidth(1930)
        self.text.setUndoRedoEnabled(True)
        self.text.setTabChangesFocus(True)
        self.text.setAcceptRichText(True)
        self.text.setOverwriteMode(True)
        self.text.setStyleSheet("""

border-radius:10px;
margin-top:10px;
margin-bottom:10px;
padding:5px;
background: white;
border:1px solid #c6c6c6;
color: black;
selection-background-color: #0033ff;
selection-color: #ffffff;

            """)

        self.cursorVisibility = QCheckBox("")
        self.cursorVisibility.setIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\textcur.png"))
        self.cursorVisibility.setFont(QFont("Arial Rounded MT Bold", 10))
        self.cursorVisibility.setChecked(False)
        self.cursorVisibility.stateChanged.connect(self.cursorVisibilityfunc)
        self.cursorVisibility.setStyleSheet("""
background-color:#f5f5f5;
margin-right:5px;
""")

        # Set the tab stop width to around 33 pixels which is
        # more or less 8 spaces
        self.text.setTabStopWidth(33)

        self.initToolbar()
        self.initFormatbar()
        self.initInsertbar()
        self.initFormulabar()
        self.initMenubar()

        self.setCentralWidget(self.text)

        # Initialize a statusbar for the window

        self.word_count_label = QLabel("Word Count: 0")
        
        

        self.abtaction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\abouticon.png"),"",self)
        self.abtaction.setIconSize(QSize(20,20))
        self.abtaction.pressed.connect(self.abtactionfunc)
        self.abtaction.setToolTip('About Vidwo Script') 
        self.abtaction.setStyleSheet("""
background:#f5f5f5;
margin-right:5px;
border:0px solid #f5f5f5;
""")
        self.abtmove_to_down_action = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\move_to_down.png"),"",self)
        self.abtmove_to_down_action.pressed.connect(self.move_to_end)
        self.abtmove_to_down_action.setToolTip('Move to end') 
        self.abtmove_to_down_action.setStyleSheet("""
background:#f5f5f5;
border:0px solid #f5f5f5;
margin-right:5px;
""")

        self.statusbar.setFont(QFont("Arial Rounded MT Bold", 9))

        self.statusbar.addPermanentWidget(self.cursorVisibility)

        self.statusbar.addPermanentWidget(self.abtmove_to_down_action)

        self.statusbar.addPermanentWidget(self.abtaction)

        self.statusbar.setStyleSheet("""
background:#f5f5f5;
font-size:15px;
padding:3px;
border:1px solid #bdbdbd;
border-radius:7px;
margin:4px;
""")
        
        # If the cursor position changes, call the function that displays
        # the line and column number
        self.text.cursorPositionChanged.connect(self.cursorPosition)

        # We need our own context menu for tables
        self.text.setContextMenuPolicy(Qt.CustomContextMenu)
        self.text.customContextMenuRequested.connect(self.context)

        self.text.textChanged.connect(self.changed)
        self.setWindowTitle("Vidwo - Script")

    def move_to_end(self):
        cursor = self.text.textCursor()
        cursor.movePosition(cursor.End)
        self.text.setTextCursor(cursor)

    def update_word_count(self,word_count_label):
        wrdcnt = self.text.toPlainText()
        word_count = len(wrdcnt.split())
        self.word_count_label.setText(f"Word Count: {word_count}")
        

    def changed(self):
        self.changesSaved = False

    def cursorVisibilityfunc(self):
        if self.cursorVisibility.isChecked() == True:
            self.text.setCursorWidth(0)
        else:
            self.text.setCursorWidth(1)

    def settingswin(self):
        self.settingwin = QDialog(self)
        self.settingwin.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\setwinico.png"))
        self.settingwin.setGeometry(300,300,430,250)
        self.settingwin.setWindowTitle("Script - Settings")
        self.settingwin.setFixedSize(630, 450)
        self.settingwin.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint,False)
        tab = QTabWidget(self.settingwin)
        tab.setFixedWidth(630)
        tab.setFixedHeight(450)
        tab.TabPosition(QTabWidget.West)
        tab.setStyleSheet("""

border-radius:6px;
""")

        # personal page
        Textcusror = QWidget(self)
##        Textcursor.setTabIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\textcur.png"))

        self.curwidchangeL = QLabel(self.settingwin)
        self.curwidchangeL.setText("Change cursor width")
        self.curwidchangeL.setFont(QFont("Arial Rounded MT Bold", 10))
        self.curwidchangeL.setStyleSheet("""
color:Black;
margin-top:7px;
background:lightgrey;
""")

        curwidrange = QSlider(Qt.Orientation.Horizontal, self)
        curwidrange.setRange(0, 100)
        curwidrange.setValue(1)
        curwidrange.setSingleStep(1)
        curwidrange.setPageStep(5)
        curwidrange.setTickPosition(QSlider.TickPosition.TicksBelow)
        curwidrange.setStyleSheet("""
""")

        curwidrange.valueChanged.connect(self.curwidchanger)

        
        
        layout = QFormLayout()
        Textcusror.setLayout(layout)
        layout.addRow(self.curwidchangeL)
        layout.addRow(curwidrange)

        # contact pane
        contact_page = QWidget(self)
        layout = QFormLayout()
        contact_page.setLayout(layout)

        # add pane to the tab widget
        tab.addTab(Textcusror, 'Text Cursor')
        tab.addTab(contact_page, 'Contact Info')

        self.settingwin.exec()

    def curwidchanger(self, value):
        self.text.setCursorWidth(int(value))

    def symbol_win(self):
        self.sywin = QDialog(self)
        self.sywin.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\symico.png"))
        self.sywin.setGeometry(300,300,430,250)
        self.sywin.setWindowTitle("Symbols")
        self.sywin.setFixedSize(330, 300)
        self.sywin.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint,False)
        self.sywin.setStyleSheet("""
background:white;
""")

        self.scroll = QScrollArea(self.sywin)

        self.widget = QWidget(self.sywin)

        tsymcursor = self.text.textCursor()

        ######################
        greek_letters = {
            "pi": "π",
    "infinity": "∞",
    "sigmaC": "Σ",
    "delta": "Δ",
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "omega": "ω",
    "lambda_": "λ",
    "theta": "θ",
    "phi": "φ",
    "rho": "ρ",
    "sigma": "σ",
    "tau": "τ",
    "upsilon": "υ",
    "xi": "ξ",
    "zeta": "ζ",
    "integral": "∫",
    "product": "∏",
    "delta": "Δ",
    "epsilon": "ε",
    "omicron": "ο",
    "eta": "η",
    "kappa": "κ",
    "mu": "μ",
    "nu": "ν",
    "pi_": "π",
    "upsilon_": "υ",
    "xi": "ξ",
    "rho_": "ρ",
    "sigma_": "σ",
    "tau_": "τ",
    "phi_": "φ",
    "chi": "χ",
    "psi": "ψ",
    "omega_": "ω",
    "alpha": "α",
    "beta": "β",
    "gamma": "γ",
    "delta": "δ",
    "epsilon": "ε",
    "zeta": "ζ",
    "eta": "η",
    "theta": "θ",
    "iota": "ι",
    "kappa": "κ",
    "lambda_": "λ",
    "mu": "μ",
    "nu": "ν",
    "xi": "ξ",
    "omicron": "ο",
    "pi": "π",
    "rho": "ρ",
    "sigma": "σ",
    "tau": "τ",
    "upsilon": "υ",
    "phi": "φ",
    "chi": "χ",
    "psi": "ψ",
    "omega": "ω",
    "Alpha": "Α",
    "Beta": "Β",
    "Gamma": "Γ",
    "Delta": "Δ",
    "Epsilon": "Ε",
    "Zeta": "Ζ",
    "Eta": "Η",
    "Theta": "Θ",
    "Iota": "Ι",
    "Kappa": "Κ",
    "Lambda": "Λ",
    "Mu": "Μ",
    "Nu": "Ν",
    "Xi": "Ξ",
    "Omicron": "Ο",
    "Pi": "Π",
    "Rho": "Ρ",
    "Sigma": "Σ",
    "Tau": "Τ",
    "Upsilon": "Υ",
    "Phi": "Φ",
    "Chi": "Χ",
    "Psi": "Ψ",
    "Omega": "Ω",
    "int_": "∫",
    "sum": "∑",
    "sqrt": "√",
    "times": "×",
    "divide": "÷",
    "leq": "≤",
    "geq": "≥",
    "neq": "≠",
    "approx": "≈",
    "infty": "∞",
    "degree": "°",
    "perp": "⊥",
    "parallel": "∥",
    "angle": "∠",
    "triangle": "△",
    "cong": "≅",
    "sim": "∼",
    "propto": "∝",
    "forall": "∀",
    "exist": "∃",
    "neg": "¬",
    "and_": "∧",
    "or_": "∨",
    "cup": "∪",
    "cap": "∩",
    "subset": "⊂",
    "supset": "⊃",
    "in_": "∈",
    "notin": "∉",
    "subseteq": "⊆",
    "supseteq": "⊇",
    "emptyset": "∅",
    "nabla": "∇",
    "partial": "∂",
    "triangleq": "≜",
    "circ": "∘",
    "varepsilon": "ε",
    "varphi": "φ",
    "vartheta": "θ",
    "varrho": "ρ",
    "varsigma": "ς",
    "varpi": "π",
    "imath": "ı",
    "jmath": "ȷ",
    "hbar": "ħ",
    "ell": "ℓ",
    "wp": "℘",
    "Re": "ℜ",
    "Im": "ℑ",
    "nabla": "∇",
    "square": "□",
    "diamond": "◇",
    "triangle": "△",
    "backslash": "∖",
    "uparrow": "↑",
    "downarrow": "↓",
    "updownarrow": "↕",
    "leftrightarrow": "↔",
    "rightarrow": "→",
    "leftarrow": "←",
    "Rightarrow": "⇒",
    "Leftarrow": "⇐",
    "Leftrightarrow": "⇔",
    "mapsto": "↦",
    "longrightarrow": "⟶",
    "longleftarrow": "⟵",
    "Longleftrightarrow": "⟷",
    "longmapsto": "⟼",
    "upharpoonright": "↾",
    "upharpoonleft": "↿",
    "downharpoonright": "⇂",
    "downharpoonleft": "⇃",
    "Uparrow": "⇑",
    "Downarrow": "⇓",
    "Updownarrow": "⇕",
    "circlearrowright": "↻",
    "circlearrowleft": "↺",
    "rightleftharpoons": "⇌",
    "leftrightharpoons": "⇋",
    "rightleftharpoonup": "⇀",
    "rightleftharpoondown": "⇁",
    "leftrightharpoonup": "↼",
    "leftrightharpoondown": "↽",
    "rightarrowtail": "↣",
    "leftarrowtail": "↢",
    "hookrightarrow": "↪",
    "hookleftarrow": "↩",
    "twoheadrightarrow": "↠",
    "twoheadleftarrow": "↞",
    "rightwavearrow": "↝",
    "leftwavearrow": "↜",
    "rightsquigarrow": "⇝",
    "leftsquigarrow": "⇜",
    "dashrightarrow": "⇢",
    "dashleftarrow": "⇠",
    "looparrowright": "↬",
    "looparrowleft": "↫",
    "Lsh": "↰",
    "Rsh": "↱",
    "rightsquigarrow": "⇝",
    "leftsquigarrow": "⇜",
    "multimap": "⊸",
    "longmapsto": "⟼",
    "longrightarrow": "⟶",
    "longleftrightarrow": "⟷",
    "longleftarrow": "⟵",
    "longleftrightarrow": "⟷",
    "hookrightarrow": "↪",
    "hookleftarrow": "↩",
    "leftrightsquigarrow": "↭",
    "curvearrowright": "↷",
    "curvearrowleft": "↶",
    "circlearrowright": "↻",
    "circlearrowleft": "↺",
    "Rightarrow": "⇒",
    "Leftarrow": "⇐",
    "Leftrightarrow": "⇔",
    "mapsto": "↦",
    "longrightarrow": "⟶",
    "longleftarrow": "⟵",
    "Longleftrightarrow": "⟷",
    "longmapsto": "⟼",
    "upharpoonright": "↾",
    "upharpoonleft": "↿",
    "downharpoonright": "⇂",
    "downharpoonleft": "⇃",
    "Uparrow": "⇑",
    "Downarrow": "⇓",
    "Updownarrow": "⇕",
    "circlearrowright": "↻",
    "circlearrowleft": "↺",
    "rightleftharpoons": "⇌",
    "leftrightharpoons": "⇋",
    "rightleftharpoonup": "⇀",
    "rightleftharpoondown": "⇁",
    "leftrightharpoonup": "↼",
    "leftrightharpoondown": "↽",
    "rightarrowtail": "↣",
    "leftarrowtail": "↢",
    "hookrightarrow": "↪",
    "hookleftarrow": "↩",
    "twoheadrightarrow": "↠",
    "twoheadleftarrow": "↞",
    "rightwavearrow": "↝",
    "leftwavearrow": "↜",
    "rightsquigarrow": "⇝",
    "leftsquigarrow": "⇜",
    "dashrightarrow": "⇢",
    "dashleftarrow": "⇠",
    "looparrowright": "↬",
    "looparrowleft": "↫",
    "Lsh": "↰",
    "Rsh": "↱",
    "rightsquigarrow": "⇝",
    "leftsquigarrow": "⇜",
    "multimap": "⊸",
    "longmapsto": "⟼",
    "longrightarrow": "⟶",
    "longleftrightarrow": "⟷",
    "longleftarrow": "⟵",
    "longleftrightarrow": "⟷",
    "hookrightarrow": "↪",
    "hookleftarrow": "↩",
    "leftrightsquigarrow": "↭",
    "curvearrowright": "↷",
    "curvearrowleft": "↶",
    "circlearrowright": "↻",
    "circlearrowleft": "↺",
    "Rightarrow": "⇒",
    "Leftarrow": "⇐",
    "Leftrightarrow": "⇔",
    "mapsto": "↦",
    "longrightarrow": "⟶",
    "longleftarrow": "⟵",
    "Longleftrightarrow": "⟷",
    "longmapsto": "⟼",
    "upharpoonright": "↾",
    "upharpoonleft": "↿",
    "downharpoonright": "⇂",
    "downharpoonleft": "⇃",
    "Uparrow": "⇑",
    "Downarrow": "⇓",
    "Updownarrow": "⇕",
    "circlearrowright": "↻",
    "circlearrowleft": "↺",
    "rightleftharpoons": "⇌",
    "leftrightharpoons": "⇋",
    "rightleftharpoonup": "⇀",
    "rightleftharpoondown": "⇁",
    "leftrightharpoonup": "↼",
    "leftrightharpoondown": "↽",
    "rightarrowtail": "↣",
    "leftarrowtail": "↢",
    "hookrightarrow": "↪",
    "hookleftarrow": "↩",
    "twoheadrightarrow": "↠",
    "twoheadleftarrow": "↞",
    "rightwavearrow": "↝",
    "leftwavearrow": "↜",
    "rightsquigarrow": "⇝",
    "leftsquigarrow": "⇜",
    "dashrightarrow": "⇢",
    "dashleftarrow": "⇠",
    "looparrowright": "↬",
    "looparrowleft": "↫",
    "Lsh": "↰",
    "Rsh": "↱",
    "rightsquigarrow": "⇝",
    "leftsquigarrow": "⇜",
    "multimap": "⊸",
    "longmapsto": "⟼",
    "longrightarrow": "⟶",
    "longleftrightarrow": "⟷",
    "longleftarrow": "⟵",
    "longleftrightarrow": "⟷",
    "hookrightarrow": "↪",
    "hookleftarrow": "↩",
    "leftrightsquigarrow": "↭",
    "curvearrowright": "↷",
    "curvearrowleft": "↶",
    "circlearrowright": "↻",
    "circlearrowleft": "↺",
    "Rightarrow": "⇒",
    "Leftarrow": "⇐",
    "Leftrightarrow": "⇔",
    "mapsto": "↦",
    "longrightarrow": "⟶",
    "longleftarrow": "⟵",
    "Longleftrightarrow": "⟷",
    "longmapsto": "⟼",
    "upharpoonright": "↾",
    "upharpoonleft": "↿",
    "downharpoonright": "⇂",
    "downharpoonleft": "⇃",
    "Uparrow": "⇑",
    "Downarrow": "⇓",
    "Updownarrow": "⇕",
        }

        def create_button(text, parent, slot):
            button = QPushButton(text, parent)
            button.setFixedSize(25, 25)
            button.setFont(QFont("Arial", 12, QFont.Bold))
            button.pressed.connect(slot)
            return button

        for name, letter in greek_letters.items():
            setattr(self, name, create_button(letter, self, lambda letter=letter: tsymcursor.insertText(letter)))
        
        self.grid = QGridLayout(self.sywin)

        self.grid.addWidget(self.pi, 0, 0)
        self.grid.addWidget(self.infinity, 0, 1)
        self.grid.addWidget(self.sigmaC, 0, 2)
        self.grid.addWidget(self.delta, 0, 3)
        self.grid.addWidget(self.alpha, 0, 4)
        self.grid.addWidget(self.beta, 0, 5)
        self.grid.addWidget(self.gamma, 0, 6)
        self.grid.addWidget(self.omega, 0, 7)
        self.grid.addWidget(self.lambda_, 0, 8)
        self.grid.addWidget(self.theta, 0, 9)

        self.grid.addWidget(self.phi, 1, 0)
        self.grid.addWidget(self.rho, 1, 1)
        self.grid.addWidget(self.sigma, 1, 2)
        self.grid.addWidget(self.tau, 1, 3)
        self.grid.addWidget(self.upsilon, 1, 4)
        self.grid.addWidget(self.xi, 1, 5)
        self.grid.addWidget(self.zeta, 1, 6)
        self.grid.addWidget(self.integral, 1, 7)
        self.grid.addWidget(self.product, 1, 8)
        self.grid.addWidget(self.epsilon, 1, 9)

        self.grid.addWidget(self.omicron, 2, 0)
        self.grid.addWidget(self.eta, 2, 1)
        self.grid.addWidget(self.kappa, 2, 2)
        self.grid.addWidget(self.mu, 2, 3)
        self.grid.addWidget(self.nu, 2, 4)
        self.grid.addWidget(self.pi_, 2, 5)
        self.grid.addWidget(self.rho_, 2, 6)
        self.grid.addWidget(self.sigma_, 2, 7)
        self.grid.addWidget(self.tau_, 2, 8)
        self.grid.addWidget(self.upsilon_, 2, 9)

        self.grid.addWidget(self.phi_, 3, 0)
        self.grid.addWidget(self.chi, 3, 1)
        self.grid.addWidget(self.psi, 3, 2)
        self.grid.addWidget(self.omega_, 3, 3)
        self.grid.addWidget(self.Alpha, 3, 4)
        self.grid.addWidget(self.Beta, 3, 5)
        self.grid.addWidget(self.Gamma, 3, 6)
        self.grid.addWidget(self.Delta, 3, 7)
        self.grid.addWidget(self.Epsilon, 3, 8)
        self.grid.addWidget(self.Zeta, 3, 9)

        self.grid.addWidget(self.Eta, 4, 0)
        self.grid.addWidget(self.Theta, 4, 1)
        self.grid.addWidget(self.Iota, 4, 2)
        self.grid.addWidget(self.Kappa, 4, 3)
        self.grid.addWidget(self.Lambda, 4, 4)
        self.grid.addWidget(self.Mu, 4, 5)
        self.grid.addWidget(self.Nu, 4, 6)
        self.grid.addWidget(self.Xi, 4, 7)
        self.grid.addWidget(self.Omicron, 4, 8)
        self.grid.addWidget(self.Pi, 4, 9)

        self.grid.addWidget(self.Rho, 5, 0)
        self.grid.addWidget(self.Sigma, 5, 1)
        self.grid.addWidget(self.Tau, 5, 2)
        self.grid.addWidget(self.Upsilon, 5, 3)
        self.grid.addWidget(self.Phi, 5, 4)
        self.grid.addWidget(self.Chi, 5, 5)
        self.grid.addWidget(self.Psi, 5, 6)
        self.grid.addWidget(self.Omega, 5, 7)
        self.grid.addWidget(self.int_, 5, 8)
        self.grid.addWidget(self.sum, 5, 9)

        self.grid.addWidget(self.sqrt, 6, 0)
        self.grid.addWidget(self.times, 6, 1)
        self.grid.addWidget(self.divide, 6, 2)
        self.grid.addWidget(self.leq, 6, 3)
        self.grid.addWidget(self.geq, 6, 4)
        self.grid.addWidget(self.neq, 6, 5)
        self.grid.addWidget(self.approx, 6, 6)
        self.grid.addWidget(self.infty, 6, 7)
        self.grid.addWidget(self.degree, 6, 8)
        self.grid.addWidget(self.perp, 6, 9)

        self.grid.addWidget(self.parallel, 7, 0)
        self.grid.addWidget(self.angle, 7, 1)
        self.grid.addWidget(self.triangle, 7, 2)
        self.grid.addWidget(self.cong, 7, 3)
        self.grid.addWidget(self.sim, 7, 4)
        self.grid.addWidget(self.propto, 7, 5)
        self.grid.addWidget(self.forall, 7, 6)
        self.grid.addWidget(self.exist, 7, 7)
        self.grid.addWidget(self.neg, 7, 8)
        self.grid.addWidget(self.and_, 7, 9)

        self.grid.addWidget(self.or_, 8, 0)
        self.grid.addWidget(self.cup, 8, 1)
        self.grid.addWidget(self.cap, 8, 2)
        self.grid.addWidget(self.subset, 8, 3)
        self.grid.addWidget(self.supset, 8, 4)
        self.grid.addWidget(self.in_, 8, 5)
        self.grid.addWidget(self.notin, 8, 6)
        self.grid.addWidget(self.subseteq, 8, 7)
        self.grid.addWidget(self.supseteq, 8, 8)
        self.grid.addWidget(self.emptyset, 8, 9)

        self.grid.addWidget(self.nabla, 9, 0)
        self.grid.addWidget(self.partial, 9, 1)
        self.grid.addWidget(self.triangleq, 9, 2)
        self.grid.addWidget(self.circ, 9, 3)
        self.grid.addWidget(self.varepsilon, 9, 4)
        self.grid.addWidget(self.varphi, 9, 5)
        self.grid.addWidget(self.vartheta, 9, 6)
        self.grid.addWidget(self.varrho, 9, 7)
        self.grid.addWidget(self.varsigma, 9, 8)
        self.grid.addWidget(self.varpi, 9, 9)

        self.grid.addWidget(self.imath, 10, 0)
        self.grid.addWidget(self.jmath, 10, 1)
        self.grid.addWidget(self.hbar, 10, 2)
        self.grid.addWidget(self.ell, 10, 3)
        self.grid.addWidget(self.wp, 10, 4)
        self.grid.addWidget(self.Re, 10, 5)
        self.grid.addWidget(self.Im, 10, 6)
        self.grid.addWidget(self.nabla, 10, 7)
        self.grid.addWidget(self.square, 10, 8)
        self.grid.addWidget(self.diamond, 10, 9)

        self.grid.addWidget(self.triangle, 11, 0)
        self.grid.addWidget(self.backslash, 11, 1)
        self.grid.addWidget(self.uparrow, 11, 2)
        self.grid.addWidget(self.downarrow, 11, 3)
        self.grid.addWidget(self.updownarrow, 11, 4)
        self.grid.addWidget(self.leftrightarrow, 11, 5)
        self.grid.addWidget(self.rightarrow, 11, 6)
        self.grid.addWidget(self.leftarrow, 11, 7)
        self.grid.addWidget(self.Rightarrow, 11, 8)
        self.grid.addWidget(self.Leftarrow, 11, 9)

        self.grid.addWidget(self.Leftrightarrow, 12, 0)
        self.grid.addWidget(self.mapsto, 12, 1)
        self.grid.addWidget(self.longrightarrow, 12, 2)
        self.grid.addWidget(self.longleftarrow, 12, 3)
        self.grid.addWidget(self.Longleftrightarrow, 12, 4)
        self.grid.addWidget(self.longmapsto, 12, 5)
        self.grid.addWidget(self.upharpoonright, 12, 6)
        self.grid.addWidget(self.upharpoonleft, 12, 7)
        self.grid.addWidget(self.downharpoonright, 12, 8)
        self.grid.addWidget(self.downharpoonleft, 12, 9)

        self.grid.addWidget(self.Uparrow, 13, 0)
        self.grid.addWidget(self.Downarrow, 13, 1)
        self.grid.addWidget(self.Updownarrow, 13, 2)
        self.grid.addWidget(self.circlearrowright, 13, 3)
        self.grid.addWidget(self.circlearrowleft, 13, 4)
        self.grid.addWidget(self.rightleftharpoons, 13, 5)
        self.grid.addWidget(self.leftrightharpoons, 13, 6)
        self.grid.addWidget(self.rightleftharpoonup, 13, 7)
        self.grid.addWidget(self.rightleftharpoondown, 13, 8)
        self.grid.addWidget(self.leftrightharpoonup, 13, 9)

        self.grid.addWidget(self.leftrightharpoondown, 14, 0)
        self.grid.addWidget(self.rightarrowtail, 14, 1)
        self.grid.addWidget(self.leftarrowtail, 14, 2)
        self.grid.addWidget(self.hookrightarrow, 14, 3)
        self.grid.addWidget(self.hookleftarrow, 14, 4)
        self.grid.addWidget(self.twoheadrightarrow, 14, 5)
        self.grid.addWidget(self.twoheadleftarrow, 14, 6)
        self.grid.addWidget(self.rightwavearrow, 14, 7)
        self.grid.addWidget(self.leftwavearrow, 14, 8)
        self.grid.addWidget(self.rightsquigarrow, 14, 9)

        self.grid.addWidget(self.leftsquigarrow, 15, 0)
        self.grid.addWidget(self.dashrightarrow, 15, 1)
        self.grid.addWidget(self.dashleftarrow, 15, 2)
        self.grid.addWidget(self.looparrowright, 15, 3)
        self.grid.addWidget(self.looparrowleft, 15, 4)
        self.grid.addWidget(self.Lsh, 15, 5)
        self.grid.addWidget(self.Rsh, 15, 6)
        self.grid.addWidget(self.rightsquigarrow, 15, 7)
        self.grid.addWidget(self.leftsquigarrow, 15, 8)
        self.grid.addWidget(self.multimap, 15, 9)

        self.widget.setLayout(self.grid)

        #Scroll Area Properties
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.scroll.setWidgetResizable(True)
        self.scroll.setWidget(self.widget)
        
        self.sywin.show()

    def plusf(self):
        cursor = self.text.textCursor()
        cursor.insertText("+")

    def subtf(self):
        cursor = self.text.textCursor()
        cursor.insertText("-")

    def multf(self):
        cursor = self.text.textCursor()
        cursor.insertText("×")

    def divif(self):
        cursor = self.text.textCursor()
        cursor.insertText("÷")

    def equalf(self):
        cursor = self.text.textCursor()
        cursor.insertText("=")

    def noteqf(self):
        cursor = self.text.textCursor()
        cursor.insertText("≠")

    def appeqf(self):
        cursor = self.text.textCursor()
        cursor.insertText("≈")

    def notappeqf(self):
        cursor = self.text.textCursor()
        cursor.insertText("≉")

    def congruf(self):
        cursor = self.text.textCursor()
        cursor.insertText("≅")

    def abtactionfunc(self):
        
        self.abtdi = QDialog(self)
        self.abtdi.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\abouticon.png"))
        self.abtdi.setGeometry(300,300,480,250)
        self.abtdi.setWindowTitle("Script - About")
        self.abtdi.setFixedSize(470, 250)
        self.abtdi.setWindowFlag(QtCore.Qt.WindowContextHelpButtonHint,False)
        self.abtdi.setStyleSheet("""
background:white;
""")

        self.topframe = QFrame(self.abtdi)
        self.topframe.resize(499,100)
        self.topframe.setStyleSheet("""background-color:royalblue;
border-radius:4px;
""")

        h1layout=QHBoxLayout(self.topframe)
        vinlayout=QVBoxLayout(self.topframe)

        self.vidLabel = QLabel(self.topframe)
        self.vidLabel.setText("Vidwo")
        self.vidLabel.setFont(QFont("Arial Rounded MT Bold", 30, QFont.Bold))
        self.vidLabel.setAlignment(Qt.AlignCenter)
        self.vidLabel.setStyleSheet("""
color:white;
margin-left:40px;
""")
        self.sLabel = QLabel(self.topframe)
        self.sLabel.setText("|")
        self.sLabel.setFont(QFont("Fira Mono Bold", 30))
        self.sLabel.setAlignment(Qt.AlignCenter)
        self.sLabel.setStyleSheet("""
color:white;
""")
        self.scLabel = QLabel(self.topframe)
        self.scLabel.setText("Script")
        self.scLabel.setFont(QFont("Fira Mono Bold", 20, QFont.Bold))
        self.scLabel.setAlignment(Qt.AlignCenter)
        self.scLabel.setStyleSheet("""
margin:0px;
color:white;
margin-right:40px;
""")
        self.pLabel = QLabel(self.topframe)
        self.pLabel.setText("PERSONAL")
        self.pLabel.setFont(QFont("Beware", 13))
        self.pLabel.setAlignment(Qt.AlignCenter)
        self.pLabel.setStyleSheet("""
color:lightgrey;
margin:0px;
margin-right:40px;
""")

        h1layout.addWidget(self.vidLabel)
        h1layout.addWidget(self.sLabel)

        
        vinlayout.addWidget(self.scLabel)
        vinlayout.addWidget(self.pLabel)
        self.setLayout(vinlayout)
        h1layout.addLayout(vinlayout)
        self.setLayout(h1layout)
        
        self.midframe = QFrame(self.abtdi)
        h1lay = QVBoxLayout(self.midframe)

        self.versLabel = QLabel(self.midframe)
        self.versLabel.setText("Version : 1.4.0")
        self.versLabel.setFont(QFont("Beware", 13))
        self.versLabel.setAlignment(Qt.AlignLeft)
        self.versLabel.setStyleSheet("""
color:Black;
margin:0px;
margin-right:40px;
""")

        self.statsLabel = QLabel(self.midframe)
        self.statsLabel.setText("Status : ACTIVATED ")
        self.statsLabel.setFont(QFont("Beware", 13))
        self.statsLabel.setAlignment(Qt.AlignLeft)
        self.statsLabel.setStyleSheet("""
color:Black;
margin:0px;
margin-right:40px;
background:lightgreen;
border:1px solid green;
""")

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.toggleFlash)
        self.timer.start(1000)  # Flash every 1 second

        self.flashOn = True

        self.licLabel = QLabel(self.midframe)
        self.licLabel.setText("Developed by Rishikesh")
        self.licLabel.setFont(QFont("Beware", 13))
        self.licLabel.setAlignment(Qt.AlignLeft)
        self.licLabel.setStyleSheet("""
color:Black;
margin:0px;
margin-right:40px;
""")

        self.updtLabel = QLabel(self.midframe)
        self.updtLabel.setText("Software Update : v 1.4.0 - Up to Date")
        self.updtLabel.setFont(QFont("Beware", 13))
        self.updtLabel.setAlignment(Qt.AlignLeft)
        self.updtLabel.setStyleSheet("""
color:Black;
margin:0px;
margin-right:40px;
""")
        
        h1lay.addWidget(self.versLabel)
        h1lay.addWidget(self.statsLabel)
        h1lay.addWidget(self.licLabel)
        h1lay.addWidget(self.updtLabel)
        
        self.midframe.resize(499,100)
        self.midframe.setStyleSheet("""background-color:lightgrey;
border-radius:4px;
""")
        vlayout=QVBoxLayout(self.abtdi)
        vlayout.addWidget(self.topframe)
        vlayout.addWidget(self.midframe)
        self.setLayout(vlayout)
        
        self.abtdi.exec()
##        call(["python", "C:\\Users\\Rishikesh\\Documents\\Vidwo Script Package\\VSUI.py"])

    def toggleFlash(self):
        if self.flashOn:
            self.statsLabel.setStyleSheet("""
            color:Black;
            margin:0px;
            margin-right:40px;
            background:lightgreen;
            border:1px solid green;
            """)
            self.licLabel.setStyleSheet("""
            color:Black;
            margin:0px;
            margin-right:40px;
            background: royalblue;
            border:1px solid royalblue;
            """)
            self.flashOn = False
        elif self.flashOn:
            self.statsLabel.setStyleSheet("""
            color:Black;
            margin:0px;
            margin-right:40px;
            background:lightgreen;
            border:1px solid green;
            """)
            self.licLabel.setStyleSheet("""
            color:Black;
            margin:0px;
            margin-right:40px;
            background: royalblue;
            border:1px solid royalblue;
            """)
            self.flashOn = False

    def closeEvent(self,event):

        if self.changesSaved:

            event.accept()

        else:
        
            popup = QtWidgets.QMessageBox(self)

            popup.setIcon(QtWidgets.QMessageBox.Warning)
            
            popup.setText("The document has been modified")
            
            popup.setInformativeText("Do you want to save your changes?")
            
            popup.setStandardButtons(QtWidgets.QMessageBox.Save   |
                                      QtWidgets.QMessageBox.Cancel |
                                      QtWidgets.QMessageBox.Discard)
            
            popup.setDefaultButton(QtWidgets.QMessageBox.Save)

            answer = popup.exec_()

            if answer == QtWidgets.QMessageBox.Save:
                self.save()

            elif answer == QtWidgets.QMessageBox.Discard:
                event.accept()

            else:
                event.ignore()

    def context(self,pos):
        pass

        # # Grab the cursor
        # cursor = self.text.textCursor()

        # # Grab the current table, if there is one
        # table = cursor.currentTable()

        # # Above will return 0 if there is no current table, in which case
        # # we call the normal context menu. If there is a table, we create
        # # our own context menu specific to table interaction
        # if table:

        #     menu = QtGui.QMenu(self)

        #     menu_corner_widget = QMenu()
        #     menu_corner_widget.setAttribute(Qt.WA_TranslucentBackground)
        #     menu.setCornerWidget(menu_corner_widget)


        #     menu.addAction(appendRowAction)
        #     menu.addAction(appendColAction)

        #     menu.addSeparator()

        #     menu.addAction(removeRowAction)
        #     menu.addAction(removeColAction)

        #     menu.addSeparator()

        #     menu.addAction(insertRowAction)
        #     menu.addAction(insertColAction)

        #     menu.addSeparator()

        #     menu.addAction(mergeAction)
        #     menu.addAction(splitAction)

        #     # Convert the widget coordinates into global coordinates
        #     pos = self.mapToGlobal(pos)

        #     # Add pixels for the tool and formatbars, which are not included
        #     # in mapToGlobal(), but only if the two are currently visible and
        #     # not toggled by the user

        #     if self.toolbar.isVisible():
        #         pos.setY(pos.y() + 45)

        #     if self.formatbar.isVisible():
        #         pos.setY(pos.y() + 45)
                
        #     # Move the menu to the new position
        #     menu.move(pos)

        #     menu.show()

        # else:

        #     event = QtGui.QContextMenuEvent(QtGui.QContextMenuEvent.Mouse,QtCore.QPoint())

        #     self.text.contextMenuEvent(event)

    


    def toggleRibbon(self):

        state = self.toolbar.isVisible()

        # Set the visibility to its inverse
        self.toolbar.setVisible(not state)

        state1 = self.insertbar.isVisible()

        # Set the visibility to its inverse
        self.insertbar.setVisible(not state1)

        state2 = self.formatbar.isVisible()

        # Set the visibility to its inverse
        self.formatbar.setVisible(not state2)

    def toggleFormulabar(self):

        state_t = self.Formulabar.isVisible()
        self.Formulabar.setVisible(not state_t)

    def toggleStatusbar(self):

        state = self.statusbar.isVisible()

        # Set the visibility to its inverse
        self.statusbar.setVisible(not state)

    def new(self):

        spawn = Main()

        spawn.show()

    def open(self):

        # Get filename and show only .writer files
        #PYQT5 Returns a tuple in PyQt5, we only need the filename
        self.filename = QtWidgets.QFileDialog.getOpenFileName(self, 'Open File',".","(*.writer)")[0]

        if self.filename:
            with open(self.filename,"rt") as file:
                self.text.setText(file.read())

    def save(self):

        # Only open dialog if there is no filename yet
        #PYQT5 Returns a tuple in PyQt5, we only need the filename
        if not self.filename:
          self.filename = QtWidgets.QFileDialog.getSaveFileName(self, 'Save File')[0]

        if self.filename:
            
            # Append extension if not there yet
            if not self.filename.endswith(".writer"):
              self.filename += ".writer"

            # We just store the contents of the text file along with the
            # format in html, which Qt does in a very nice way for us
            with open(self.filename,"wt") as file:
                file.write(self.text.toHtml())

            self.changesSaved = True

    def preview(self):

        # Open preview dialog
        preview = QtPrintSupport.QPrintPreviewDialog()

        # If a print is requested, open print dialog
        preview.paintRequested.connect(lambda p: self.text.print_(p))

        preview.exec_()

    def printHandler(self):

        # Open printing dialog
        dialog = QtPrintSupport.QPrintDialog()

        if dialog.exec_() == QtWidgets.QDialog.Accepted:
            self.text.document().print_(dialog.printer())

    def cursorPosition(self):

        cursor = self.text.textCursor()

        # Mortals like 1-indexed things
        line = cursor.blockNumber() + 1
        col = cursor.columnNumber()

        self.statusbar.showMessage("Ln: {} | Col: {}".format(line,col))

    def wordCount(self):

        wc = wordcount.WordCount(self)

        wc.getText()

        wc.show()

    def insertImage(self):

        # Get image file name
        #PYQT5 Returns a tuple in PyQt5
        filename = QtWidgets.QFileDialog.getOpenFileName(self, 'Insert image',".","Images (*.png *.xpm *.jpg *.bmp *.gif)")[0]

        if filename:
            
            # Create image object
            image = QtGui.QImage(filename)

            # Error if unloadable
            if image.isNull():

                popup = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Critical,
                                          "Image load error",
                                          "Could not load image file!",
                                          QtWidgets.QMessageBox.Ok,
                                          self)
                popup.show()

            else:

                cursor = self.text.textCursor()

                cursor.insertImage(image,filename)

    def fontColorChanged(self):

        # Get a color from the text dialog
        color = QtWidgets.QColorDialog.getColor()

        # Set it as the new text color
        self.text.setTextColor(color)

    def highlight(self):

        color = QtWidgets.QColorDialog.getColor()

        self.text.setTextBackgroundColor(color)

    def bold(self):

        if self.text.fontWeight() == QtGui.QFont.Bold:

            self.text.setFontWeight(QtGui.QFont.Normal)

        else:

            self.text.setFontWeight(QtGui.QFont.Bold)

    def italic(self):

        state = self.text.fontItalic()

        self.text.setFontItalic(not state)

    def underline(self):

        state = self.text.fontUnderline()

        self.text.setFontUnderline(not state)

    def strike(self):

        # Grab the text's format
        fmt = self.text.currentCharFormat()

        # Set the fontStrikeOut property to its opposite
        fmt.setFontStrikeOut(not fmt.fontStrikeOut())

        # And set the next char format
        self.text.setCurrentCharFormat(fmt)

    def superScript(self):

        # Grab the current format
        fmt = self.text.currentCharFormat()

        # And get the vertical alignment property
        align = fmt.verticalAlignment()

        # Toggle the state
        if align == QtGui.QTextCharFormat.AlignNormal:

            fmt.setVerticalAlignment(QtGui.QTextCharFormat.AlignSuperScript)

        else:

            fmt.setVerticalAlignment(QtGui.QTextCharFormat.AlignNormal)

        # Set the new format
        self.text.setCurrentCharFormat(fmt)

    def subScript(self):

        # Grab the current format
        fmt = self.text.currentCharFormat()

        # And get the vertical alignment property
        align = fmt.verticalAlignment()

        # Toggle the state
        if align == QtGui.QTextCharFormat.AlignNormal:

            fmt.setVerticalAlignment(QtGui.QTextCharFormat.AlignSubScript)

        else:

            fmt.setVerticalAlignment(QtGui.QTextCharFormat.AlignNormal)

        # Set the new format
        self.text.setCurrentCharFormat(fmt)

    def alignLeft(self):
        self.text.setAlignment(Qt.AlignLeft)

    def alignRight(self):
        self.text.setAlignment(Qt.AlignRight)

    def alignCenter(self):
        self.text.setAlignment(Qt.AlignCenter)

    def alignJustify(self):
        self.text.setAlignment(Qt.AlignJustify)

    def indent(self):

        # Grab the cursor
        cursor = self.text.textCursor()

        if cursor.hasSelection():

            # Store the current line/block number
            temp = cursor.blockNumber()

            # Move to the selection's end
            cursor.setPosition(cursor.anchor())

            # Calculate range of selection
            diff = cursor.blockNumber() - temp

            direction = QtGui.QTextCursor.Up if diff > 0 else QtGui.QTextCursor.Down

            # Iterate over lines (diff absolute value)
            for n in range(abs(diff) + 1):

                # Move to start of each line
                cursor.movePosition(QtGui.QTextCursor.StartOfLine)

                # Insert tabbing
                cursor.insertText("\t")

                # And move back up
                cursor.movePosition(direction)

        # If there is no selection, just insert a tab
        else:

            cursor.insertText("\t")

    def handleDedent(self,cursor):

        cursor.movePosition(QtGui.QTextCursor.StartOfLine)

        # Grab the current line
        line = cursor.block().text()

        # If the line starts with a tab character, delete it
        if line.startswith("\t"):

            # Delete next character
            cursor.deleteChar()

        # Otherwise, delete all spaces until a non-space character is met
        else:
            for char in line[:8]:

                if char != " ":
                    break

                cursor.deleteChar()

    def dedent(self):

        cursor = self.text.textCursor()

        if cursor.hasSelection():

            # Store the current line/block number
            temp = cursor.blockNumber()

            # Move to the selection's last line
            cursor.setPosition(cursor.anchor())

            # Calculate range of selection
            diff = cursor.blockNumber() - temp

            direction = QtGui.QTextCursor.Up if diff > 0 else QtGui.QTextCursor.Down

            # Iterate over lines
            for n in range(abs(diff) + 1):

                self.handleDedent(cursor)

                # Move up
                cursor.movePosition(direction)

        else:
            self.handleDedent(cursor)


    def bulletList(self):

        cursor = self.text.textCursor()

        # Insert bulleted list
        cursor.insertList(QtGui.QTextListFormat.ListDisc)

    def numberList(self):

        cursor = self.text.textCursor()

        # Insert list with numbers
        cursor.insertList(QtGui.QTextListFormat.ListDecimal)

    def textSize(self, pointSize):
        pointSize = float(self.comboSize.currentText())
        if pointSize > 0:
            fmt = QTextCharFormat()
            fmt.setFontPointSize(pointSize)
            self.mergeFormatOnWordOrSelection(fmt)

def myStyleSheet(self):
    return """
QToolBar::item{
background:#DCDCDC;
optacity:0.35;
}

QScrollBar:vertical {
            border: 0px solid #999999;
            background:transparent;
            width:10px;    
            padding: 0px 3px 0px 0px;
        }
        QScrollBar::handle:vertical {         
       
            min-height: 0px;
          	border: 0px solid red;
			border-radius: 3px;
			background-color: grey;
        }
        QScrollBar::add-line:vertical {       
            height: 0px;
            subcontrol-position: bottom;
            subcontrol-origin: margin;
            border-radius: 3px;
        }
        QScrollBar::sub-line:vertical {
            height: 0 px;
            subcontrol-position: top;
            subcontrol-origin: margin;
            border-radius: 3px;
        }
QScrollBar:vertical {
            border: 0px solid #999999;
            background:transparent;
            width:10px;    
            padding: 0px 3px 0px 0px;
        }
        QScrollBar::handle:vertical {         
       
            min-height: 0px;
          	border: 0px solid red;
			border-radius: 3px;
			background-color: grey;
        }
        QScrollBar::add-line:vertical {       
            height: 0px;
            subcontrol-position: bottom;
            subcontrol-origin: margin;
            border-radius: 3px;
        }
        QScrollBar::sub-line:vertical {
            height: 0 px;
            subcontrol-position: top;
            subcontrol-origin: margin;
            border-radius: 3px;
        }
QScrollBar:horizontal:hover {
            border: 0px solid #999999;
            background:transparent;
            width:5px;    
            padding: 0px 3px 0px 0px;
        }
        QScrollBar::handle:vertical {         
       
            min-height: 0px;
          	border: 0px solid red;
			border-radius: 3px;
			background-color: grey;
        }
        QScrollBar::add-line:vertical {       
            height: 0px;
            subcontrol-position: bottom;
            subcontrol-origin: margin;
        }
        QScrollBar::sub-line:vertical {
            height: 0 px;
            subcontrol-position: top;
            subcontrol-origin: margin;
        }
        
QLineEdit{
border-radius:7px;
background:darkgrey;
font-size:14px;
padding:4px;
}
QLineEdit:focus{
border-bottom: 2px solid #004F98;
}

QMainWindow
{
background:#e3e5e9;
border-radius:4px;
}

QTextEdit
{
border-radius:10px;
margin_top:5px;
margin_bottom:5px;
margin-left:300px;
margin-right:300px;
padding:5px;
width:200px;
color: black;
selection-background-color: #0033ff;
selection-color: #ffffff;
}

}
QMenuBar
{
margin-top:7px;
margin-left:7px;
margin-right:7px;
margin-bottom:3px;
border-radius:5px;
background: royalblue;
font-size:15px;
width:50px;
color:white;
padding:3px;
}
            QMenuBar::item {
border-radius:5px;
font-size:20px;

            }
            QMenuBar::item:selected {
            border-radius:5px;
            background-color:#1F4BE8;
            }
            QMenu {
background-color: white;
                margin:4px;
                padding:4px;
                        }
            QMenu::item {
            }
            QMenu::item:selected {
                background-color: lightgrey;
                color:black;
                border-radius:5px;
            }

QMenuBar:hover{
font-size:15px bold;
}
QToolBar:
{
background:#DCDCDC
font-size:15px;
}
QVBoxLayout{
background:black;
margin:0px;
border-radius:20px;
}
QStatusBar{
background:white;
}

    """

def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setStyle(QStyleFactory.create('Fusion'))


    main = Main()
    main.showMaximized()

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
