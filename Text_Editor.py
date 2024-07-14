import sys
import os
from subprocess import call
from PyQt5 import QtWidgets
from PyQt5 import QtPrintSupport
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog
from PyQt5 import QtGui, QtCore
from PyQt5.QtGui import QFontDatabase
from PyQt5.QtCore import Qt
from PyQt5 import QtGui
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
import ctypes
from googletrans import Translator
from PyQt5.QtGui import QSyntaxHighlighter, QTextCharFormat, QFont, QColor, QTextCursor
from PyQt5.QtCore import Qt, QRegExp
from ctypes import windll, c_int, c_uint, POINTER, Structure
from PyQt5.QtMultimedia import QSound


class Main(QtWidgets.QMainWindow):

    def __init__(self,parent=None):
        QtWidgets.QMainWindow.__init__(self,parent)
        self.setStyleSheet(myStyleSheet(self))
        self.filename = ""
        self.changesSaved = True
        self.setWindowIcon(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\VSLOGO.png"))
        self.initUI()

    def create_tabs(self):

        shadow_effect = QGraphicsDropShadowEffect()
        shadow_effect.setBlurRadius(10)
        shadow_effect.setColor(QtGui.QColor(136, 136, 136))
        shadow_effect.setXOffset(2)
        shadow_effect.setYOffset(2)

        self.tab_widget = QTabWidget()
        # self.tab_widget.setGraphicsEffect(shadow_effect)
        self.tab_widget.setMovable(True)
        tabbar = self.tab_widget.findChild(QtWidgets.QTabBar)
        if tabbar:
            tabbar.setFixedHeight(50)

        self.tab_widget.setStyleSheet("""

QTabWidget::pane { /* The tab widget frame */
            background: white;
            border-radius: 10px;
            height:140px;
        }
        QTabWidget::tab-bar {
            alignment: left;
        }
        
        QTabBar::tab {
            background: #e3e5e9;
            border: 0px solid #e3e5e9;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
            margin-bottom:7px;
                                      font-size:15px;
        }
                                      
        QTabBar::tab:hover {
            background: #e3e5e9;
            border-bottom: 0px solid gray;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
            margin-bottom:7px;
                                      font-weight:bold;
                                      font-size:15px;
        }
        
        QTabBar::tab:selected {
            background: #e3e5e9;
            border-bottom: 3px solid royalblue;
            color:royalblue;
            padding: 5px;
            border-radius:3px;
                                      font-weight:bold;
                                      margin-bottom:7px;
        }

""")
        # First tab
        tab1 = QWidget()
        self.newAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\newdoc.png"),"",self)
        self.newAction.setShortcut("Ctrl+N")
        self.newAction.setIconSize(QSize(60, 60))
        self.newAction.setStatusTip("Create a new document")
        self.newAction.setToolTip("Create a new document")
        self.newAction.clicked.connect(self.new)
        self.newActionlabel = QLabel("New")
        self.newActionlabel.setAlignment(QtCore.Qt.AlignCenter)
        self.newActionlabel.setStyleSheet("""
font-family:Verdana;
                                              font size:13px;
                                              border:none;
                                              color:gray;
""")

        self.openAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\opendoc.png"),"",self)
        self.openAction.setStatusTip("Open existing document")
        self.openAction.setToolTip("Open existing document")
        self.openAction.setShortcut("Ctrl+O")
        self.openAction.setIconSize(QSize(25, 25))
        self.openAction.clicked.connect(self.open)

        self.customopenAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\customopen.png"),"",self)
        self.customopenAction.setStatusTip("Open existing document via script file dialog")
        self.customopenAction.setToolTip("Open existing document via script file dialog")
        self.customopenAction.setShortcut("Ctrl+O")
        self.customopenAction.setIconSize(QSize(25, 25))
        self.customopenAction.clicked.connect(self.customopen)

        self.saveAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\savedoc.png"),"",self)
        self.saveAction.setStatusTip("Save document")
        self.saveAction.setToolTip("Save document")
        self.saveAction.setShortcut("Ctrl+S")
        self.saveAction.setIconSize(QSize(25, 25))
        self.saveAction.clicked.connect(self.save)

        self.saveasPDFAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\saveaspdf.png"),"",self)
        self.saveasPDFAction.setStatusTip("Save document as PDF")
        self.saveasPDFAction.setToolTip("Save document as PDF")
        self.saveasPDFAction.setShortcut("Ctrl+S")
        self.saveasPDFAction.setIconSize(QSize(25, 25))
        self.saveasPDFAction.clicked.connect(self.save_as_pdf)

        self.printAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\print.png"),"",self)
        self.printAction.setStatusTip("Print document")
        self.printAction.setToolTip("Print document")
        self.printAction.setShortcut("Ctrl+P")
        self.printAction.setIconSize(QSize(25, 25))
        self.printAction.clicked.connect(self.printHandler)

        self.previewAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\printpreview.png"),"",self)
        self.previewAction.setStatusTip("Preview page before printing")
        self.previewAction.setToolTip("Preview page before printing")
        self.previewAction.setShortcut("Ctrl+Shift+P")
        self.previewAction.setIconSize(QSize(25, 25))
        self.previewAction.clicked.connect(self.preview)

        tab1_layout = QVBoxLayout()
        
        # Create a frame with horizontal layout
        frame1 = QFrame()
        #frame1.setFrameShape(QFrame.StyledPanel)
        frame1.setFixedWidth(280)
        frame1_layout = QHBoxLayout(frame1)
        
        # Create two frames inside frame1 with vertical layout
        frame2 = QFrame()
        frame2.setFrameShape(QFrame.StyledPanel)
        frame2.setFixedWidth(100)
        frame2_layout = QVBoxLayout(frame2)
        frame2_layout.addWidget(self.newAction)
        frame2_layout.addWidget(self.newActionlabel)
        frame2.setStyleSheet("""
border: none;      
""")
        
        frame2a = QFrame()
        frame2a.setFrameShape(QFrame.StyledPanel)
        frame2a.setFixedWidth(50)
        frame2a_layout = QVBoxLayout(frame2a)
        frame2a_layout.addWidget(self.openAction)
        frame2a_layout.addWidget(self.customopenAction)
        frame2a.setStyleSheet("""
border: none;      
                              border-right: 1px solid lightgrey;
""")
        
        frame3 = QFrame()
        frame3.setFrameShape(QFrame.StyledPanel)
        frame3.setFixedWidth(50)
        frame3_layout = QVBoxLayout(frame3)
        frame3_layout.addWidget(self.saveAction)
        frame3_layout.addWidget(self.saveasPDFAction)
        frame3.setStyleSheet("""
QFrame {
                border: none;
                border-right: 1px solid lightgrey;
            }
""")

        frame4 = QFrame()
        frame4.setFrameShape(QFrame.StyledPanel)
        frame4.setFixedWidth(50)
        frame4_layout = QVBoxLayout(frame4)
        frame4_layout.addWidget(self.printAction)
        frame4_layout.addWidget(self.previewAction)
        frame4.setStyleSheet("""
border: none;     
""")
        
        # Add frame2 and frame3 to frame1's layout
        frame1_layout.addWidget(frame2)
        frame1_layout.addWidget(frame2a)
        frame1_layout.addWidget(frame3)
        frame1_layout.addWidget(frame4)
        
        tab1_layout.addWidget(frame1)
        tab1.setLayout(tab1_layout)
        
        # Second tab
        tab2 = QWidget()
        tab2_layout = QVBoxLayout()

        self.cutAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\cut.png"), "", self)
        self.cutAction.setShortcut("Ctrl+X")
        self.cutAction.setIconSize(QSize(25, 25))
        self.cutAction.setStatusTip("Delete and copy text to clipboard")
        self.cutAction.setToolTip("Cut to clipboard")
        self.cutAction.clicked.connect(self.text.cut)

        self.copyAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\copy.png"), "", self)
        self.copyAction.setShortcut("Ctrl+C")
        self.copyAction.setIconSize(QSize(25, 25))
        self.copyAction.setToolTip("Copy to clipboard")
        self.copyAction.clicked.connect(self.text.copy)
        self.copyAction.clicked.connect(lambda: self.statusbar.showMessage("Text copied to clipboard", 2000))

        self.pasteAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\paste.png"), "", self)
        self.pasteAction.setShortcut("Ctrl+V")
        self.pasteAction.setIconSize(QSize(25, 25))
        self.pasteAction.setStatusTip("Paste text from clipboard")
        self.pasteAction.setToolTip("Paste from clipboard")
        self.pasteAction.clicked.connect(self.text.paste)

        self.selectallAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\selectall.png"), "", self)
        self.selectallAction.setShortcut("Ctrl+A")
        self.selectallAction.setIconSize(QSize(25, 25))
        self.selectallAction.setStatusTip("Select all the text")
        self.selectallAction.setToolTip("Select all the text")
        self.selectallAction.clicked.connect(self.text.selectAll)

        self.undoAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\undo.png"), "", self)
        self.undoAction.setShortcut("Ctrl+Z")
        self.undoAction.setIconSize(QSize(25, 25))
        self.undoAction.setStatusTip("Undo last action")
        self.undoAction.setToolTip("Undo last action")
        self.undoAction.clicked.connect(self.text.undo)

        self.redoAction = QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\redo.png"), "", self)
        self.redoAction.setShortcut("Ctrl+Y")
        self.redoAction.setIconSize(QSize(25, 25))
        self.redoAction.setStatusTip("Redo last action")
        self.redoAction.setToolTip("Redo last action")
        self.redoAction.clicked.connect(self.text.redo)

        frame1 = QFrame()
        frame1.setFixedWidth(1140)
        frame1_layout = QHBoxLayout(frame1)
        
        frame2 = QFrame()
        frame2.setFrameShape(QFrame.StyledPanel)
        frame2.setFixedWidth(50)
        frame2_layout = QVBoxLayout(frame2)
        frame2_layout.addWidget(self.cutAction)
        frame2_layout.addWidget(self.selectallAction)
        frame2.setStyleSheet("""
border: none;  
""")
        
        frame3 = QFrame()
        frame3.setFrameShape(QFrame.StyledPanel)
        frame3.setFixedWidth(50)
        frame3_layout = QVBoxLayout(frame3)
        frame3_layout.addWidget(self.copyAction)
        frame3_layout.addWidget(self.pasteAction)
        frame3.setStyleSheet("""
QFrame {
                border: none;
                border-right: 1px solid lightgrey;
            }   
""")

        frame4 = QFrame()
        frame4.setFrameShape(QFrame.StyledPanel)
        frame4.setFixedWidth(50)
        frame4_layout = QVBoxLayout(frame4)
        frame4_layout.addWidget(self.undoAction)
        frame4_layout.addWidget(self.redoAction)
        frame4.setStyleSheet("""
QFrame {
                border: none;
                border-right: 1px solid lightgrey;
            }
""")
        
        frame5 = QtWidgets.QFrame()
        frame5.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame5.setFixedWidth(180)
        frame5_layout = QtWidgets.QVBoxLayout(frame5)
        frame5.setStyleSheet("""
QFrame {
                border: none;
                             border-right: 1px solid lightgrey;
                             margin-right:10px;
            }
""")

        fontBox = QtWidgets.QFontComboBox(self)
        fontBox.currentFontChanged.connect(lambda font: self.text.setCurrentFont(font))
        fontBox.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        fontBox.setMaximumSize(180, 28)
        fontBox.setStyleSheet("""
            QComboBox {
                font-size: 15px;
                background-color: #FFFFFF;
                selection-background-color: royalblue;
                selection-color: white;
                border: 1px solid #CCCCCC;
                padding: 2px;
                border-radius: 4px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: darkgray;
                border-left-style: solid;
                background: white;
                image: url(C:/Users/rishi/OneDrive/Documents/VS_Icons/dda.png);
                background-size: 5px;
                background-repeat: no-repeat;
                background-position: right 10px center;
            }
        """)


        fontSize = QtWidgets.QComboBox(self)
        fontSize.setEditable(True)
        fontSize.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        fontSize.setMaximumSize(70, 28)
        fontSize.setStyleSheet("""
            QComboBox {
                font-size: 15px;
                background-color: #FFFFFF;
                selection-background-color: royalblue;
                selection-color: white;
                border: 1px solid #CCCCCC;
                padding: 2px;
                border-radius: 4px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: darkgray;
                border-left-style: solid;
                background: white;
                image: url(C:/Users/rishi/OneDrive/Documents/VS_Icons/dda.png);
                background-size: 5px;
                background-repeat: no-repeat;
                background-position: right 10px center;
            }
        """)

        for size in range(8, 100, 2):
            fontSize.addItem(f"{size} pt")

        fontSize.currentIndexChanged.connect(self.setFontSize)

        fontSize.setCurrentIndex(1)
        frame5_layout.addWidget(fontBox)
        frame5_layout.addWidget(fontSize)

        self.underline_combo = QtWidgets.QComboBox()
        self.underline_combo.addItem("No Underline", QTextCharFormat.NoUnderline)
        self.underline_combo.addItem("Solid Line", QTextCharFormat.SingleUnderline)
        self.underline_combo.addItem("Dashed Line", QTextCharFormat.DashUnderline)
        self.underline_combo.addItem("Dotted Line", QTextCharFormat.DotLine)
        self.underline_combo.addItem("Dash Dot Line", QTextCharFormat.DashDotLine)
        self.underline_combo.addItem("Dash Dot Dot Line", QTextCharFormat.DashDotDotLine)
        self.underline_combo.addItem("Wave Line", QTextCharFormat.WaveUnderline)
        self.underline_combo.currentIndexChanged.connect(self.applyFormatting)
        self.underline_combo.setStyleSheet("""
QComboBox {
                font-size: 15px;
                background-color: #FFFFFF;
                selection-background-color: royalblue;
                selection-color: white;
                border: 1px solid #CCCCCC;
                padding: 2px;
                border-radius:4px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: darkgray;
                border-left-style: solid;
                background:white;
                image: url(C:/Users/rishi/OneDrive/Documents/VS_Icons/dda.png);
                background-size: 5px;
                background-repeat: no-repeat;
                background-position: right 10px center;
            }
""")
        self.comboStyle = QtWidgets.QComboBox()
        index = 0
        self.comboStyle.setEditable(True)
        icon = QIcon('C:\\Users\\rishi\\OneDrive\\Desktop\\Vidwo Worsksuit SCRIPT\\ICON\\bulletlist.png')
        self.comboStyle.setItemIcon(0 , QIcon('C:\\Users\\rishi\\OneDrive\\Desktop\\Vidwo Worsksuit SCRIPT\\ICON\\bulletlist.png'))
        self.comboStyle.addItem("No bullet", 0)
        self.comboStyle.addItem("●", 1)
        self.comboStyle.addItem("○", 2)
        self.comboStyle.addItem("■", 3)
        self.comboStyle.addItem("⒈", 4)
        self.comboStyle.addItem("a.", 5)
        self.comboStyle.addItem("A.", 6)
        self.comboStyle.addItem("ⅰ", 7)
        self.comboStyle.addItem("Ⅰ", 8)
        self.comboStyle.setStyleSheet("""
QComboBox {
                font-size: 15px;
                background-color: #FFFFFF;
                selection-background-color: royalblue;
                selection-color: white;
                border: 1px solid #CCCCCC;
                padding: 2px;
                padding-left:5px;
                border-radius:4px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left-width: 1px;
                border-left-color: darkgray;
                border-left-style: solid;
                background:white;
                image: url(C:/Users/rishi/OneDrive/Documents/VS_Icons/dda.png);
                background-size: 5px;
                background-repeat: no-repeat;
                background-position: right 10px center;
            }
""")
        self.comboStyle.activated.connect(self.textStyle)

        fontColorButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/fontcolor.png"), "", self)
        fontColorButton.clicked.connect(self.fontColorChanged)
        fontColorButton.setIconSize(QSize(25, 25))
        fontColorButton.setToolTip("Change font color")
        fontColorButton.setStatusTip("Change font color")

        bgActButton = QtWidgets.QPushButton(QtGui.QIcon('C:/Users/rishi/OneDrive/Documents/VS_Icons/pgbgcolor.png'),"", self)
        bgActButton.clicked.connect(self.changeBGColor)
        bgActButton.setIconSize(QSize(25, 25))
        bgActButton.setToolTip("Change Background Color")
        bgActButton.setStatusTip("Change Background Color")

        self.backColorButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/bgcolor.png"), "", self)
        self.backColorButton.clicked.connect(self.highlight)
        self.backColorButton.setIconSize(QtCore.QSize(25, 25))
        self.backColorButton.setToolTip("Change background color")
        self.backColorButton.setStatusTip("Change background color")

        remformattingbtn = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/remform.png"), "", self)
        remformattingbtn.clicked.connect(self.remove_formatting)
        remformattingbtn.setIconSize(QtCore.QSize(25, 25))
        remformattingbtn.setToolTip("Remove all formatting on text")
        remformattingbtn.setStatusTip("Remove all formatting on text")

        boldButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/bold.png"), "", self)
        boldButton.clicked.connect(self.bold)
        boldButton.setIconSize(QSize(25, 25))
        boldButton.setToolTip("Bold")
        boldButton.setStatusTip("Bold")

        italicButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/italic.png"), "", self)
        italicButton.clicked.connect(self.italic)
        italicButton.setIconSize(QSize(25, 25))
        italicButton.setToolTip("Italic")
        italicButton.setStatusTip("Italic")

        underlButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/underline.png"), "", self)
        underlButton.clicked.connect(self.underline)
        underlButton.setIconSize(QSize(25, 25))
        underlButton.setToolTip("Underline")
        underlButton.setStatusTip("Underline")

        strikeButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/strikeout.png"), "", self)
        strikeButton.clicked.connect(self.strike)
        strikeButton.setIconSize(QSize(25, 25))
        strikeButton.setToolTip("Strike-out")
        strikeButton.setStatusTip("Strike-out")

        superButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/superscript.png"), "", self)
        superButton.clicked.connect(self.superScript)
        superButton.setIconSize(QSize(25, 25))
        superButton.setToolTip("Superscript")
        superButton.setStatusTip("Superscript")

        subButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/subscript.png"), "", self)
        subButton.clicked.connect(self.subScript)
        subButton.setIconSize(QSize(25, 25))
        subButton.setToolTip("Subscript")
        subButton.setStatusTip("Subscript")

        capAllButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/Capall.png"), "", self)
        capAllButton.clicked.connect(self.capitalizeSelectedText)
        capAllButton.setIconSize(QSize(25, 25))
        capAllButton.setToolTip("Capitalize Selected Text")
        capAllButton.setStatusTip("Capitalize Selected Text")

        lowAllButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/Lowall.png"), "", self)
        lowAllButton.clicked.connect(self.lowercaseSelectedText)
        lowAllButton.setIconSize(QSize(25, 25))
        lowAllButton.setToolTip("Lowercase Selected Text")
        lowAllButton.setStatusTip("Lowercase Selected Text")

        swapAllButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/swapcase.png"), "", self)
        swapAllButton.clicked.connect(self.swapcaseSelectedText)
        swapAllButton.setIconSize(QSize(25, 25))
        swapAllButton.setToolTip("Swapcase Selected Text")
        swapAllButton.setStatusTip("Swapcase Selected Text")

        self.alignCenterButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\center.png"), "", self)
        self.alignCenterButton.setIconSize(QSize(25, 25))
        self.alignCenterButton.setStatusTip("Align center")
        self.alignCenterButton.setToolTip("Align center")
        self.alignCenterButton.clicked.connect(self.alignCenterf)

        self.alignLeftButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\left.png"), "", self)
        self.alignLeftButton.setIconSize(QSize(25, 25))
        self.alignLeftButton.setStatusTip("Align right")
        self.alignLeftButton.setToolTip("Align right")
        self.alignLeftButton.clicked.connect(self.alignLeftf)

        self.alignRightButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\right.png"), "", self)
        self.alignRightButton.setIconSize(QSize(25, 25))
        self.alignRightButton.setStatusTip("Align right")
        self.alignRightButton.setToolTip("Align right")
        self.alignRightButton.clicked.connect(self.alignRightf)

        self.alignJustifyButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\justify.png"), "", self)
        self.alignJustifyButton.setIconSize(QSize(25, 25))
        self.alignJustifyButton.setStatusTip("Align justify")
        self.alignJustifyButton.setToolTip("Align justify")
        self.alignJustifyButton.clicked.connect(self.alignJustifyf)

        indentButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\indent.png"), "", self)
        indentButton.setIconSize(QSize(25, 25))
        indentButton.setShortcut("Ctrl+Tab")
        indentButton.setStatusTip("Indent Area")
        indentButton.setToolTip("Indent Area")
        indentButton.clicked.connect(self.indent)

        dedentButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\dedent.png"), "", self)
        dedentButton.setIconSize(QSize(25, 25))
        dedentButton.setShortcut("Shift+Tab")
        dedentButton.setStatusTip("Dedent Area")
        dedentButton.setToolTip("Dedent Area")
        dedentButton.clicked.connect(self.dedent)

        templateButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\template.png"), "", self)
        templateButton.setIconSize(QSize(25, 25))
        templateButton.setStatusTip("Explore Templates")
        templateButton.setToolTip("Explore Templates")
        templateButton.clicked.connect(self.template_Dialog)

        translateButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\translate.png"), "", self)
        translateButton.setIconSize(QSize(25, 25))
        translateButton.setStatusTip("Translate")
        translateButton.setToolTip("Translate")
        translateButton.clicked.connect(self.translate_Dialog)

        editPageBodyButton = QtWidgets.QPushButton(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\edithtmlpagebody.png"), "", self)
        editPageBodyButton.setIconSize(QSize(60, 60))
        editPageBodyButton.setStatusTip("Edit page body using HTML")
        editPageBodyButton.setToolTip("Edit page body using HTML")
        editPageBodyButton.clicked.connect(self.editBody)
        editPageBodyButtonlabel = QLabel("Page HTML")
        editPageBodyButtonlabel.setAlignment(QtCore.Qt.AlignCenter)
        editPageBodyButtonlabel.setStyleSheet("""
font-family:Verdana;
                                              font size:13px;
                                              border:none;
                                              color:gray;
""")

        frame6 = QFrame()
        frame6.setFrameShape(QFrame.StyledPanel)
        frame6.setFixedWidth(70)
        frame6_layout = QVBoxLayout(frame6)
        frame6_layout.addWidget(fontColorButton)
        frame6_layout.addWidget(bgActButton)
        frame6.setStyleSheet("""
QFrame {
                             margin-left:25px;
                border: none;
            }    
""")
        
        frame6a = QFrame()
        frame6a.setFrameShape(QFrame.StyledPanel)
        frame6a.setFixedWidth(50)
        frame6a_layout = QVBoxLayout(frame6a)
        frame6a_layout.addWidget(self.backColorButton)
        frame6a_layout.addWidget(remformattingbtn)
        frame6a.setStyleSheet("""
QFrame {
                border: none;
            }    
""")
        
        frame6b = QFrame()
        frame6b.setFrameShape(QFrame.StyledPanel)
        frame6b.setFixedWidth(90)
        frame6b_layout = QVBoxLayout(frame6b)
        frame6b_layout.addWidget(editPageBodyButton)
        frame6b_layout.addWidget(editPageBodyButtonlabel)
        frame6b.setStyleSheet("""
QFrame {
                border: none;
                border-right: 1px solid lightgrey;
            }    
""")

        frame7 = QFrame()
        frame7.setFrameShape(QFrame.StyledPanel)
        frame7.setFixedWidth(130)
        frame7_layout = QVBoxLayout(frame7)
        frame7_layout.addWidget(self.comboStyle)
        frame7_layout.addWidget(self.underline_combo)
        frame7.setStyleSheet("""
border: none;      
""")
        
        frame7a = QFrame()
        frame7a.setFrameShape(QFrame.StyledPanel)
        frame7a.setFixedWidth(50)
        frame7a_layout = QVBoxLayout(frame7a)
        frame7a_layout.addWidget(strikeButton)
        frame7a_layout.addWidget(underlButton)
        frame7a.setStyleSheet("""
border: none;     
""")
        
        frame8 = QFrame()
        frame8.setFrameShape(QFrame.StyledPanel)
        frame8.setFixedWidth(50)
        frame8_layout = QVBoxLayout(frame8)
        frame8_layout.addWidget(boldButton)
        frame8_layout.addWidget(italicButton)
        frame8.setStyleSheet("""
border: none;      
""")
        
        frame9 = QFrame()
        frame9.setFrameShape(QFrame.StyledPanel)
        frame9.setFixedWidth(50)
        frame9_layout = QVBoxLayout(frame9)
        frame9_layout.addWidget(superButton)
        frame9_layout.addWidget(subButton)
        frame9.setStyleSheet("""
border: none;    
""")
        
        frame10 = QFrame()
        frame10.setFrameShape(QFrame.StyledPanel)
        frame10.setFixedWidth(50)
        frame10_layout = QVBoxLayout(frame10)
        frame10_layout.addWidget(capAllButton)
        frame10_layout.addWidget(lowAllButton)
        frame10_layout.addWidget(swapAllButton)
        frame10.setStyleSheet("""
QFrame {
                border: none;
                border-right: 1px solid lightgrey;
            } 
""")
        
        frame11 = QFrame()
        frame11.setFrameShape(QFrame.StyledPanel)
        frame11.setFixedWidth(50)
        frame11_layout = QVBoxLayout(frame11)
        frame11_layout.addWidget(self.alignLeftButton)
        frame11_layout.addWidget(self.alignRightButton)
        frame11.setStyleSheet("""
border: none;    
""")
        frame12 = QFrame()
        frame12.setFrameShape(QFrame.StyledPanel)
        frame12.setFixedWidth(50)
        frame12_layout = QVBoxLayout(frame12)
        frame12_layout.addWidget(self.alignCenterButton)
        frame12_layout.addWidget(self.alignJustifyButton)
        frame12.setStyleSheet("""
border: none;    
""")
        frame13 = QFrame()
        frame13.setFrameShape(QFrame.StyledPanel)
        frame13.setFixedWidth(50)
        frame13_layout = QVBoxLayout(frame13)
        frame13_layout.addWidget(indentButton)
        frame13_layout.addWidget(dedentButton)
        frame13.setStyleSheet("""
QFrame{
                              border:none;
                              border-right:1px solid lightgrey;
                              }   
""")
        
        frame14 = QFrame()
        frame14.setFrameShape(QFrame.StyledPanel)
        frame14.setFixedWidth(50)
        frame14_layout = QVBoxLayout(frame14)
        frame14_layout.addWidget(templateButton)
        frame14_layout.addWidget(translateButton)
        frame14.setStyleSheet("""
border: none;    
""")

        frame1_layout.addWidget(frame2)
        frame1_layout.addWidget(frame3)
        frame1_layout.addWidget(frame4)
        frame1_layout.addWidget(frame5)
        frame1_layout.addWidget(frame6)
        frame1_layout.addWidget(frame6a)
        frame1_layout.addWidget(frame6b)
        frame1_layout.addWidget(frame7)
        frame1_layout.addWidget(frame7a)
        frame1_layout.addWidget(frame8)
        frame1_layout.addWidget(frame9)
        frame1_layout.addWidget(frame10)
        frame1_layout.addWidget(frame11)
        frame1_layout.addWidget(frame12)
        frame1_layout.addWidget(frame13)
        frame1_layout.addWidget(frame14)
        
        tab2_layout.addWidget(frame1)
        tab2.setLayout(tab2_layout)

        # Third tab
        tab3 = QWidget()

        self.dateTimeButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/datetime.png"), "", self)
        self.dateTimeButton.clicked.connect(self.calendershow)
        self.dateTimeButton.setIconSize(QtCore.QSize(60, 60))
        self.dateTimeButton.setToolTip("Insert current date/time")
        self.dateTimeButton.setStatusTip("Insert current date/time")
        self.dateTimeButtonlabel = QLabel("Date & Time")
        self.dateTimeButtonlabel.setAlignment(QtCore.Qt.AlignCenter)
        self.dateTimeButtonlabel.setStyleSheet("""
font-family:Verdana;
                                              font size:13px;
                                              border:none;
                                              color:gray;
                                              alignment:center;
""")

        self.tableButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/table.png"), "", self)
        self.tableButton.clicked.connect(self.tableDialog)
        self.tableButton.setIconSize(QtCore.QSize(60, 60))
        self.tableButton.setToolTip("Insert table")
        self.tableButton.setStatusTip("Insert table")
        self.tableButtonlabel = QLabel("Table")
        self.tableButtonlabel.setAlignment(QtCore.Qt.AlignCenter)
        self.tableButtonlabel.setStyleSheet("""
font-family:Verdana;
                                              font size:13px;
                                              border:none;
                                              color:gray;
                                              alignment:center;
""")

        self.imageButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/add_image.png"), "", self)
        self.imageButton.clicked.connect(self.insertImage)
        self.imageButton.setIconSize(QtCore.QSize(60, 60))
        self.imageButton.setToolTip("Insert image")
        self.imageButton.setStatusTip("Insert image")
        self.imageButtonlabel = QLabel("Images")
        self.imageButtonlabel.setAlignment(QtCore.Qt.AlignCenter)
        self.imageButtonlabel.setStyleSheet("""
font-family:Verdana;
                                              font size:13px;
                                              border:none;
                                              color:gray;
                                              alignment:center;
""")

        self.symbolButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/symico.png"), "", self)
        self.symbolButton.clicked.connect(self.symbol_win)
        self.symbolButton.setIconSize(QtCore.QSize(25, 25))
        self.symbolButton.setToolTip("Symbols")

        self.equationButton = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/equation.png"), "", self)
        self.equationButton.clicked.connect(self.eq_win)
        self.equationButton.setIconSize(QtCore.QSize(25, 25))
        self.equationButton.setToolTip("Equation")

        self.insertlink = QtWidgets.QPushButton(QtGui.QIcon("C:/Users/rishi/OneDrive/Documents/VS_Icons/insertlink1.png"), "", self)
        self.insertlink.clicked.connect(self.open_link_dialog)
        self.insertlink.setIconSize(QtCore.QSize(25, 25))
        self.insertlink.setToolTip("Symbols")
        
        tab3_layout = QVBoxLayout()

        frame1 = QFrame()
        frame1.setFixedWidth(410)
        frame1_layout = QHBoxLayout(frame1)

        # Create two frames inside frame1 with vertical layout
        frame2 = QtWidgets.QFrame()
        frame2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        frame2.setFixedWidth(100)
        frame2_layout = QtWidgets.QVBoxLayout(frame2)
        frame2_layout.addWidget(self.tableButton)
        frame2_layout.addWidget(self.tableButtonlabel)
        frame2.setStyleSheet("""
        QFrame {
            border: none;
            border-right: 1px solid lightgrey;
        }
        """)

        frame3 = QFrame()
        frame3.setFrameShape(QFrame.StyledPanel)
        frame3.setFixedWidth(100)
        frame3_layout = QVBoxLayout(frame3)
        frame3_layout.addWidget(self.imageButton)
        frame3_layout.addWidget(self.imageButtonlabel)
        frame3.setStyleSheet("""
        QFrame {
            border: none;
            border-right: 1px solid lightgrey;
        }
        """)

        frame4 = QFrame()
        frame4.setFrameShape(QFrame.StyledPanel)
        frame4.setFixedWidth(100)
        frame4_layout = QVBoxLayout(frame4)
        frame4_layout.addWidget(self.dateTimeButton)
        frame4_layout.addWidget(self.dateTimeButtonlabel)
        frame4.setStyleSheet("""
        QFrame {
            border: none;
            border-right: 1px solid lightgrey;
        }
        """)

        frame5 = QFrame()
        frame5.setFrameShape(QFrame.StyledPanel)
        frame5.setFixedWidth(50)
        frame5_layout = QVBoxLayout(frame5)
        frame5_layout.addWidget(self.symbolButton)
        frame5_layout.addWidget(self.equationButton)
        frame5.setStyleSheet("""
        QFrame {
            border: none;
            border-right: 1px solid lightgrey;
                             margin-right:5px;
        }
        """)

        frame6 = QFrame()
        frame6.setFrameShape(QFrame.StyledPanel)
        frame6.setFixedWidth(50)
        frame6_layout = QVBoxLayout(frame6)
        frame6_layout.addWidget(self.insertlink)
        frame6.setStyleSheet("""
        QFrame {
            border: none;
        }
        """)

        frame1_layout.addWidget(frame2)
        frame1_layout.addWidget(frame3)
        frame1_layout.addWidget(frame4)
        frame1_layout.addWidget(frame5)
        frame1_layout.addWidget(frame6)

        tab3_layout.addWidget(frame1)
        tab3.setLayout(tab3_layout)

        tab4 = QWidget()
        tab4_layout = QVBoxLayout()
        toolbar4 = QToolBar()
        tab4_layout.addWidget(toolbar4)
        tab4.setLayout(tab4_layout)

        # Add tabs to the QTabWidget
        self.tab_widget.addTab(tab1, 'File')
        self.tab_widget.addTab(tab2, 'Home')
        self.tab_widget.addTab(tab3, 'Insert')
        self.tab_widget.addTab(tab4, 'Format')
        initialTabIndex = 1  # Tab 2 (index starts from 0)
        self.tab_widget.setCurrentIndex(initialTabIndex)

    def open_link_dialog(self):
        cursor = self.text.textCursor()
        selected_text = cursor.selectedText()

        dialog = QDialog(self)
        dialog.setWindowTitle('Insert Link')

        layout = QVBoxLayout(dialog)

        tab_widget = QTabWidget()
        web_link_tab = QWidget()
        tab_widget.addTab(web_link_tab, "Web Link")

        web_link_layout = QVBoxLayout(web_link_tab)

        text_label = QLabel("Selected Text:")
        web_link_layout.addWidget(text_label)

        text_edit = QLineEdit(selected_text)
        text_edit.setReadOnly(True)  # Make the selected text read-only
        web_link_layout.addWidget(text_edit)

        link_label = QLabel("URL:")
        web_link_layout.addWidget(link_label)

        link_edit = QLineEdit()
        web_link_layout.addWidget(link_edit)

        insert_button = QPushButton("Insert Link")
        web_link_layout.addWidget(insert_button)

        layout.addWidget(tab_widget)
        dialog.setLayout(layout)

        def insert_link():
            text = text_edit.text()  # Use the text_edit QLineEdit to get the selected text
            url = link_edit.text()
            if not url:
                QMessageBox.warning(dialog, "Error", "URL cannot be empty")
                return
            link = f'<a href="{url}">{text}</a>'
            cursor.insertHtml(link)
            dialog.accept()

        insert_button.clicked.connect(insert_link)

        dialog.exec_()

    def remove_formatting(self):
        cursor = self.text.textCursor()
        if cursor.hasSelection():
            # Clear formatting
            char_format = QtGui.QTextCharFormat()
            char_format.setFontWeight(QtGui.QFont.Normal)
            char_format.setFontItalic(False)
            char_format.setFontUnderline(False)
            char_format.setForeground(QtGui.QBrush(QtCore.Qt.black))
            char_format.setBackground(QtGui.QBrush(QtCore.Qt.transparent))
            cursor.mergeCharFormat(char_format)

    def customopen(self):
        customopendialog = QtWidgets.QDialog(self)
        customopendialog.setWindowTitle("Script - File Dialog")
        customopendialog.setStyleSheet("""
background:white;
""")

        # Main layout for the dialog
        main_layout = QtWidgets.QVBoxLayout(customopendialog)

        # Layout for folder path input
        path_input_layout = QtWidgets.QHBoxLayout()
        main_layout.addLayout(path_input_layout)

        # Line edit for folder path input
        self.folderLineEdit = QtWidgets.QLineEdit()
        self.folderLineEdit.setPlaceholderText("Enter folder path here")
        self.folderLineEdit.setStyleSheet("""
QLineEdit {
    border-radius: 5px;
    font-size: 15px;
    padding:3px;
    font-family: Verdana;
    background: white;
    border:1px solid lightgrey;
}
""")
        path_input_layout.addWidget(self.folderLineEdit)

        # Button to trigger folder scan
        self.scanButton = QtWidgets.QPushButton("Scan for Files")
        self.scanButton.clicked.connect(self.on_folder_entered)
        self.scanButton.setStyleSheet("""
padding:4px;
background:royalblue;
font-weight:bold;
color:white;
                                      font-size:15px;
                                      border-radius:5px;
""")
        path_input_layout.addWidget(self.scanButton)

        # Layout for extension input
        ext_input_layout = QtWidgets.QHBoxLayout()
        main_layout.addLayout(ext_input_layout)

        # Line edit for extension input
        self.extensionLineEdit = QtWidgets.QLineEdit()
        self.extensionLineEdit.setPlaceholderText("Enter file extension (e.g., .txt)")
        self.extensionLineEdit.setStyleSheet("""
QLineEdit {
    border-radius: 5px;
    font-size: 15px;
    padding:3px;
    font-family: Courier;
    background: white;
    border:1px solid lightgrey;
}
""")
        ext_input_layout.addWidget(self.extensionLineEdit)

        # Layout for tree view and list widget
        content_layout = QtWidgets.QHBoxLayout()
        main_layout.addLayout(content_layout)

        # Tree view for folder selection
        self.treeView = QtWidgets.QTreeView()
        content_layout.addWidget(self.treeView)

        # Set up the file system model
        self.model = QtWidgets.QFileSystemModel()
        self.model.setRootPath(QtCore.QDir.rootPath())
        self.treeView.setModel(self.model)
        self.treeView.setRootIndex(self.model.index(QtCore.QDir.rootPath()))
        self.treeView.setColumnWidth(0, 250)

        # Apply stylesheet to tree view for border radius
        self.treeView.setStyleSheet("""
            QTreeView {
                border-radius: 10px;
                border: 1px solid lightgrey;
            }
            QTreeView::item:selected {
                background-color: royalblue;
                color: white;
                border-radius:6px;

            }
            QTreeView::item {
            border-radius:6px;
            background:white;
                                    margin:2px;
                                    padding:2px;
            }
                                    QTreeView::item:hover {
            border-radius:6px;
            background:white;
                                    margin:2px;
                                    padding:2px;
                border: 2px solid royalblue;
                                      color:royalblue;
            }
        """)

        # Connect the selection signal
        self.treeView.selectionModel().selectionChanged.connect(self.on_folder_selected)

        # List widget to display files
        self.listWidget = QtWidgets.QListWidget()
        content_layout.addWidget(self.listWidget)

        # Apply stylesheet to list widget for text size, border radius, and selection color
        self.listWidget.setStyleSheet("""
            QListWidget {
                border-radius: 10px;
                border: 1px solid royalblue;
                padding: 5px;
            }
            QListWidget::item {
                border-radius: 5px;
                border: 1px solid lightgrey;
                margin: 2px;
                padding: 5px;
                font-size: 14pt;
            }
            QListWidget::item:hover {
                border-radius: 5px;
                border: 2px solid royalblue;
                                      color:royalblue;
                margin: 2px;
                padding: 5px;
                font-size: 14pt;
            }
            QListWidget::item:selected {
                background-color: royalblue;
                color: white;
            }
        """)

        # Connect item selection to update line edit with file path
        self.listWidget.itemSelectionChanged.connect(self.on_item_selected)

        # Label to show status
        self.statusLabel = QtWidgets.QLabel("")
        main_layout.addWidget(self.statusLabel)

        customopendialog.resize(800, 600)
        customopendialog.exec_()

    def on_folder_entered(self):
        # Get the directory from the line edit
        folder_path = self.folderLineEdit.text()
        if os.path.isdir(folder_path):
            self.statusLabel.setText(f"Selected folder: {folder_path}")
            self.scan_for_files(folder_path)
            # Expand the tree view to the entered directory
            index = self.model.index(folder_path)
            if index.isValid():
                self.treeView.setCurrentIndex(index)
                self.treeView.scrollTo(index, QtWidgets.QAbstractItemView.PositionAtCenter)
        else:
            self.statusLabel.setText("Invalid folder path. Please enter a valid directory.")

    def on_folder_selected(self, selected, deselected):
        indexes = self.treeView.selectionModel().selectedIndexes()
        if indexes:
            # Get the selected folder path
            folder_path = self.model.filePath(indexes[0])
            if os.path.isdir(folder_path):
                self.statusLabel.setText(f"Selected folder: {folder_path}")
                self.scan_for_files(folder_path)

    def scan_for_files(self, directory):
        self.statusLabel.setText("Scanning for files...")
        QtWidgets.QApplication.processEvents()  # Update the UI

        extension = self.extensionLineEdit.text().strip()
        files = self.find_files(directory, extension)
        self.listWidget.clear()
        for file_path in files:
            file_name = os.path.basename(file_path)
            list_item = QtWidgets.QListWidgetItem(file_name)
            list_item.setData(QtCore.Qt.UserRole, file_path)
            self.listWidget.addItem(list_item)
        
        self.statusLabel.setText(f"<span style='color:green'>Found {len(files)} files in {directory}</span>")

    def find_files(self, directory, extension):
        files = []
        for root, dirs, files_in_dir in os.walk(directory):
            for file in files_in_dir:
                if not extension or file.endswith(extension):
                    files.append(os.path.join(root, file))
        return files

    def on_item_selected(self):
        selected_items = self.listWidget.selectedItems()
        if selected_items:
            file_path = selected_items[0].data(QtCore.Qt.UserRole)
            self.folderLineEdit.setText(file_path)
            self.open_file_in_textedit(file_path)

    def open_file_in_textedit(self, file_path):
        try:
            with open(file_path, 'r') as file:
                content = file.read()
                self.text.setText(content)
        except Exception as e:
            error_message = f"<span style='color:red'>{e}</span>"
            self.statusLabel.setText(f"Failed to open file: {error_message}")

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

        right_widget = QWidget(self)
        right_layout = QHBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setAlignment(Qt.AlignCenter)

        # Add widgets to the right widget
     
#         self.FinderLine = QtWidgets.QLineEdit()
#         self.FinderLine.setStyleSheet("""
# font-family: arial;
# background:#89A4F3;
# border-radius:5px;
# font-size:17px;
# padding:2px;
# width:400px;
# margin-top:3px;
# margin-right:750px;
#             """)
        
#         right_layout.addWidget(self.FinderLine)

        # Set the right widget as the menu bar's right corner widget
        menubar.setCornerWidget(right_widget, Qt.TopRightCorner)

        # Add the most important actions to the menubar

        self.settingAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\setwinico.png"),"Settings",self)
        self.settingAction.setStatusTip("Settings")
        self.settingAction.triggered.connect(self.settingswin)

        self.editpagebodyashtml = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\edithtmlpagebody.png"),"Format page HTML",self)
        self.editpagebodyashtml.setStatusTip("Edit page body using HTML")
        self.editpagebodyashtml.triggered.connect(self.editBody)

        self.file.addAction(self.newAction)
        self.file.addAction(self.openAction)
        self.file.addAction(self.saveAction)
        self.file.addAction(self.saveasPDFAction)
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
        self.edit.addAction(self.editpagebodyashtml)
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

        ribbonAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\toolbar.png"),"Toggle Ribbon",self)
        ribbonAction.triggered.connect(self.toggleribbon)

        formulabarAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\formulabar.png"),"Toggle Formulabar",self)
        formulabarAction.triggered.connect(self.toggleFormulabar)

        statusbarAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\statusbar.png"),"Toggle Statusbar",self)
        statusbarAction.triggered.connect(self.toggleStatusbar)

        lefttabaction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\toolbar.png"),"Move Tabs to Left",self)
        lefttabaction.triggered.connect(self.toggletableft)

        centertabaction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\formulabar.png"),"Move Tabs to Center",self)
        centertabaction.triggered.connect(self.toggletabcenter)

        righttabaction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\statusbar.png"),"Move Tabs to Right",self)
        righttabaction.triggered.connect(self.toggletabright)

        self.view.addAction(ribbonAction)
        self.view.addAction(formulabarAction)
        self.view.addAction(statusbarAction)
        self.view.addSeparator()
        self.view.addAction(lefttabaction)
        self.view.addAction(centertabaction)
        self.view.addAction(righttabaction)

        About = QtWidgets.QAction("About",self)
        About.triggered.connect(self.abtactionfunc)

        self.Help.addAction(About)

    def toggletableft(self):
        self.set_tabbar_alignment("left")

    def toggletabcenter(self):
        self.set_tabbar_alignment("center")

    def toggletabright(self):
        self.set_tabbar_alignment("right")

    def set_tabbar_alignment(self, alignment):
        # Remove any existing stylesheet from the tab widget
        self.tab_widget.setStyleSheet("")

        # Define a new stylesheet with updated alignment
        if alignment == "left":
            stylesheet = """
            QTabWidget::pane { /* The tab widget frame */
            background: white;
            border-radius: 10px;
            height:140px;
        }
        QTabWidget::tab-bar {
            alignment: left;
        }
        
        QTabBar::tab {
            background: #e3e5e9;
            border: 0px solid #e3e5e9;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
                                      font-size:15px;
        }
                                      
        QTabBar::tab:hover {
            background: #e3e5e9;
            border-bottom: 0px solid gray;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
                                      font-weight:bold;
                                      font-size:15px;
        }
        
        QTabBar::tab:selected {
            background: #e3e5e9;
            border-bottom: 3px solid royalblue;
            color:royalblue;
            padding: 5px;
            border-radius:3px;
                                      font-weight:bold;
        }
            """
        elif alignment == "center":
            stylesheet = """
           QTabWidget::pane { /* The tab widget frame */
            background: white;
            border-radius: 10px;
            height:140px;
        }
        QTabWidget::tab-bar {
            alignment: center;
        }
        
        QTabBar::tab {
            background: #e3e5e9;
            border: 0px solid #e3e5e9;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
                                      font-size:15px;
        }
                                      
        QTabBar::tab:hover {
            background: #e3e5e9;
            border-bottom: 0px solid gray;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
                                      font-weight:bold;
                                      font-size:15px;
        }
        
        QTabBar::tab:selected {
            background: #e3e5e9;
            border-bottom: 3px solid royalblue;
            color:royalblue;
            padding: 5px;
            border-radius:3px;
                                      font-weight:bold;
        }
            """
        elif alignment == "right":
            stylesheet = """
            QTabWidget::pane { /* The tab widget frame */
            background: white;
            border-radius: 10px;
            height:140px;
        }
        QTabWidget::tab-bar {
            alignment: right;
        }
        
        QTabBar::tab {
            background: #e3e5e9;
            border: 0px solid #e3e5e9;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
                                      font-size:15px;
        }
                                      
        QTabBar::tab:hover {
            background: #e3e5e9;
            border-bottom: 0px solid gray;
            padding: 5px;
            width:70px;
            border-radius:3px;
            font-family: Arial;
                                      font-weight:bold;
                                      font-size:15px;
        }
        
        QTabBar::tab:selected {
            background: #e3e5e9;
            border-bottom: 3px solid royalblue;
            color:royalblue;
            padding: 5px;
            border-radius:3px;
                                      font-weight:bold;
        }
            """

        # Apply the updated stylesheet to the tab widget
        self.tab_widget.setStyleSheet(stylesheet)

    def editBody(self):
        all = self.text.document().toHtml()
        body = all.partition("<body style=")[2].partition(">")[0]
        dlg = QInputDialog()
        mybody, ok = dlg.getText(self, 'change body style', "", QLineEdit.Normal, body, Qt.Dialog)
        if ok:
            new = all.replace(body, mybody)    
            self.text.document().setHtml(new)
            self.statusBar().showMessage("body style changed")
        else:
            self.statusBar().showMessage("body style not changed")

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

        self.saveasPDFAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\saveaspdf.png"),"Save As PDF",self)
        self.saveasPDFAction.setStatusTip("Save document")
        self.saveasPDFAction.setShortcut("Ctrl+S")
        self.saveasPDFAction.triggered.connect(self.save_as_pdf)

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

    def save_as_pdf(self):
        # Get the file name and path to save the PDF
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getSaveFileName(self, "Save as PDF", "", "PDF files (*.pdf);;All files (*.*)")

        if file_path:
            # Create a QPrinter object
            printer = QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(file_path)
            self.text.document().print_(printer)

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

        copy_button = QPushButton("Copy", translate_dialog)
        copy_button.clicked.connect(lambda: QApplication.clipboard().setText(translated_text_box.toPlainText()))
        layout.addWidget(copy_button)

        # Button to trigger translation
        translate_button = QPushButton("Translate", translate_dialog)
        translate_button.clicked.connect(lambda: self.translate_text(input_text_box, source_language, target_language, translated_text_box))
        layout.addWidget(translate_button)

        # Translate Selection Button
        translate_selection_button = QPushButton("Translate and Replace Selection", translate_dialog)
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
        
        # Create a grid layout
        grid_layout = QGridLayout()
        
        # Add buttons to the grid layout
        grid_layout.addWidget(Letter_temp, 0, 0)
        grid_layout.addWidget(job_app_form_temp, 0, 1)
        
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

        self.dateTimeAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\datetime.png"),"Insert current date/time",self)
        self.dateTimeAction.setStatusTip("Insert current date/time")
        self.dateTimeAction.setShortcut("Ctrl+D")
        self.dateTimeAction.triggered.connect(self.calendershow)

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

        self.equationAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\equation.png"),"Equation",self)
        self.equationAction.triggered.connect(self.eq_win)

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
            self.show_warning_message_ent_rc()

        if not rows_text.isdigit() or not columns_text.isdigit():
            self.show_warning_message_inval_rc()
            return

        if border_type == 'Thick':
            self.insertTablethick()
        elif border_type == 'Thin':
            self.insertTablethin()

    def show_warning_message_ent_rc(self):
        popup = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning,
                                    'Warning',
                                    'Please enter the number of rows and columns.',
                                    parent=self.dialog)
        
        # Apply custom stylesheet to the QMessageBox
        popup.setStyleSheet("""
                                QLabel{
                                font-size:15px;
                                font-weight: bold;
                                }
            QPushButton {
                background-color: light grey; /* Button background color */
                color: black; /* Button text color */
                border: 1px solid #c6c6c6; /* Button border */
                padding: 5px 10px; /* Button padding */
                margin: 5px; /* Button margin */
                border-radius: 5px; /* Button border radius */
                font-family: Arial; /* Button font family */
                font-size: 15px; /* Button font size */
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e0e0e0; /* Hover background color */
                                color:royalblue;
                                border:1px solid royalblue;
            }
            QPushButton:pressed {
                background-color: #d0d0d0; /* Pressed background color */
            }
        """)

        popup.exec_()

    def show_warning_message_inval_rc(self):
        popup = QtWidgets.QMessageBox(QtWidgets.QMessageBox.Warning,
                                    'Warning',
                                    'Please enter the number of rows and columns.',
                                    parent=self.dialog)
        
        # Apply custom stylesheet to the QMessageBox
        popup.setStyleSheet("""
                                QLabel{
                                font-size:15px;
                                font-weight: bold;
                                }
            QPushButton {
                background-color: light grey; /* Button background color */
                color: black; /* Button text color */
                border: 1px solid #c6c6c6; /* Button border */
                padding: 5px 10px; /* Button padding */
                margin: 5px; /* Button margin */
                border-radius: 5px; /* Button border radius */
                font-family: Arial; /* Button font family */
                font-size: 15px; /* Button font size */
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e0e0e0; /* Hover background color */
                                color:royalblue;
                                border:1px solid royalblue;
            }
            QPushButton:pressed {
                background-color: #d0d0d0; /* Pressed background color */
            }
        """)

        popup.exec_()

    def insertTablethick(self):
        rows_text = self.rows_edit.text().strip()
        columns_text = self.columns_edit.text().strip()

        if not rows_text or not columns_text:
            self.show_warning_message_ent_rc()

        if not rows_text.isdigit() or not columns_text.isdigit():
            self.show_warning_message_inval_rc()

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

        self.boldAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\bold.png"),"Bold",self)
        self.boldAction.triggered.connect(self.bold)

        self.italicAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\italic.png"),"Italic",self)
        self.italicAction.triggered.connect(self.italic)

        self.underlAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\underline.png"),"Underline",self)
        self.underlAction.triggered.connect(self.underline)

        self.strikeAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\strikeout.png"),"Strike-out",self)
        self.strikeAction.triggered.connect(self.strike)

        self.superAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\superscript.png"),"Superscript",self)
        self.superAction.triggered.connect(self.superScript)

        self.subAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\subscript.png"),"Subscript",self)
        self.subAction.triggered.connect(self.subScript)

        Capall = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\Lowall.png"),"Capitalize Selected Text",self)
        Capall.triggered.connect(self.capitalizeSelectedText)
        Lowall = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\Capall.png"),"Lowercase Selected Text",self)
        Lowall.triggered.connect(self.lowercaseSelectedText)

        self.comboStyle.activated.connect(self.textStyle)

        self.alignLeft = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\left.png"),"Align left",self)
        self.alignLeft.triggered.connect(self.alignLeftf)

        self.alignCenter = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\center.png"),"Align center",self)
        self.alignCenter.triggered.connect(self.alignCenterf)

        self.alignRight = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\right.png"),"Align right",self)
        self.alignRight.triggered.connect(self.alignRightf)

        self.alignJustify = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\justify.png"),"Align justify",self)
        self.alignJustify.triggered.connect(self.alignJustifyf)

        indentAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\indent.png"),"Indent Area",self)
        indentAction.setShortcut("Ctrl+Tab")
        indentAction.triggered.connect(self.indent)

        dedentAction = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\dedent.png"),"Dedent Area",self)
        dedentAction.setShortcut("Shift+Tab")
        dedentAction.triggered.connect(self.dedent)

        self.backColor = QtWidgets.QAction(QtGui.QIcon("C:\\Users\\rishi\\OneDrive\\Documents\\VS_Icons\\bgcolor.png"),"Change background color",self)
        self.backColor.triggered.connect(self.highlight)

    def capitalizeSelectedText(self):
        cursor = self.text.textCursor()
        if not cursor.hasSelection():
            return
        
        selected_text = cursor.selectedText()
        cursor.beginEditBlock()
        cursor.insertText(selected_text.upper())  # Capitalize the selected text
        cursor.endEditBlock()
        
        self.text.setTextCursor(cursor)

    def lowercaseSelectedText(self):
        cursor = self.text.textCursor()
        if not cursor.hasSelection():
            return
        
        selected_text = cursor.selectedText()
        cursor.beginEditBlock()
        cursor.insertText(selected_text.lower())  # Lowercase the selected text
        cursor.endEditBlock()
        
        self.text.setTextCursor(cursor)

    def swapcaseSelectedText(self):
        cursor = self.text.textCursor()
        if not cursor.hasSelection():
            return
        
        selected_text = cursor.selectedText()
        cursor.beginEditBlock()
        cursor.insertText(selected_text.swapcase())  # Lowercase the selected text
        cursor.endEditBlock()
        
        self.text.setTextCursor(cursor)

    def applyFormatting(self):
        # Get the selected underline style from the combo box
        index = self.underline_combo.currentIndex()
        underline_style = self.underline_combo.itemData(index)
        
        # Get the current QTextCursor
        cursor = self.text.textCursor()
        
        # Create a QTextCharFormat for underline style
        underline_format = QTextCharFormat()
        underline_format.setUnderlineStyle(underline_style)
        
        # Apply the format to the selected text range
        cursor.mergeCharFormat(underline_format)
        self.text.setTextCursor(cursor)

    def editHTML(self):
        pass

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
        a4_width_mm = 210
        mm_per_inch = 25.4
        dpi = 96
        a4_width_pixels = int((a4_width_mm / mm_per_inch) * dpi)
        self.text.setFixedSize(a4_width_pixels, 700)  # Set height as needed
        self.text.setAutoFormatting(QtWidgets.QTextEdit.AutoAll)
        self.text.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.text.customContextMenuRequested.connect(self.context_menu)
        self.text.setCursorWidth(1)
        self.text.setFixedWidth(a4_width_pixels)
        self.text.setUndoRedoEnabled(True)
        self.text.setTabChangesFocus(True)
        self.text.setAcceptRichText(True)
        self.text.setOverwriteMode(False)
        self.text.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.text.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.text.setStyleSheet("""
QTextEdit {
                padding: 5px;
                background: white;
                border: 1px solid #c6c6c6;
                color: black;
                selection-background-color: #0033ff;
                selection-color: #ffffff;
                border-radius:10px;
            }
            QScrollBar:vertical {
                background: rgba(0, 0, 0, 0);
                width: 12px;
                margin: 0px 0px 0px 0px;
            }
            QScrollBar::handle:vertical {
                background: #b0b0b0;
                min-height: 20px;
                border-radius: 4px;
            }
            QScrollBar::add-line:vertical {
                background: #c6c6c6;
                height: 0px;
                subcontrol-position: bottom;
                subcontrol-origin: margin;
            }
            QScrollBar::sub-line:vertical {
                background: #c6c6c6;
                height: 0px;
                subcontrol-position: top;
                subcontrol-origin: margin;
            }
            QScrollBar:horizontal {
                background: rgba(0, 0, 0, 0);
                height: 9px;
            }
            QScrollBar::handle:horizontal {
                background: #b0b0b0;
                min-width: 20px;
                border-radius: 4px;
            }
            QScrollBar::add-line:horizontal {
                background: #c6c6c6;
                width: 0px;
                subcontrol-position: right;
                subcontrol-origin: margin;
            }
            QScrollBar::sub-line:horizontal {
                background: #c6c6c6;
                width: 0px;
                subcontrol-position: left;
                subcontrol-origin: margin;
            }
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

        self.tab_widget = QTabWidget()

        self.create_tabs()
        self.initToolbar()
        self.initFormatbar()
        self.initInsertbar()
        self.initMenubar()
        self.initFormulabar()

        container_widget = QWidget()
        self.setCentralWidget(container_widget)

        # Create a layout for the container widget
        container_layout = QVBoxLayout(container_widget)
        container_layout.addWidget(self.tab_widget)
        container_layout.addWidget(self.Formulabar)
        container_layout.addWidget(self.text, alignment=QtCore.Qt.AlignCenter)

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

    def context_menu(self, position):
        menu = QMenu()
        menu.addAction(self.undoAction)
        menu.addAction(self.redoAction)
        menu.addAction(self.cutAction)
        menu.addAction(self.copyAction)
        menu.addAction(self.pasteAction)
        menu.addAction(self.selectallAction)
        menu.addAction(self.editpagebodyashtml)
        menu.addSeparator()
        menu.addAction(self.boldAction)
        menu.addAction(self.italicAction)
        menu.addAction(self.underlAction)
        menu.addAction(self.alignLeft)
        menu.addAction(self.alignCenter)
        menu.addAction(self.alignRight)
        menu.addAction(self.alignJustify)
        menu.addAction(self.backColor)
        menu.addAction(self.translateAction)
        menu.setStyleSheet("""
            QMenu {
                background-color: white;
                border: 1px solid lightgrey;
                border-radius:5px;
            }
            QMenu::item {
                padding:5px;
                           margin:3px;
            }
            QMenu::item:hover {
                background: royalblue;
                           color:white;
                           font-weight:bold;
            }
        """)

        menu.exec_(self.text.mapToGlobal(position))

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
background:#C5C3C3;
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

        # add pane to the tab widget
        tab.addTab(Textcusror, 'Text Cursor')


        self.settingwin.exec()

    def toggletoRichText(self, state):
        if state == Qt.Checked:
            self.text.setAcceptRichText(False)
        else:
            self.text.setAcceptRichText(True)

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

    def eq_win(self):
        self.eqwin = QDialog(self)
        self.eqwin.setWindowTitle("Equations")
        self.eqwin.setStyleSheet("background:white;")

        tsymcursor = self.text.textCursor()

        trigo_fn = {
            "sin": "sin()",
            "cos": "cos()",
            "tan": "tan()",
            "cosec": "cosec()",
            "sec": "sec()",
            "cot": "cot()",
            "sinh": "sinh()",
            "cosh": "cosh()",
            "tanh": "tanh()",
            "cosech": "cosech()",
            "sech": "sech()",
            "coth": "coth()",
            "sin⁻¹": "sin⁻¹()",
            "cos⁻¹": "cos⁻¹()",
            "tan⁻¹": "tan⁻¹()",
            "cosec⁻¹": "cosec⁻¹()",
            "sec⁻¹": "sec⁻¹()",
            "cot⁻¹": "cot⁻¹()",
            "sinh⁻¹": "sinh⁻¹()",
            "cosh⁻¹": "cosh⁻¹()",
            "tanh⁻¹": "tanh⁻¹()",
            "cosech⁻¹": "cosech⁻¹()",
            "sech⁻¹": "sech⁻¹()",
            "coth⁻¹": "coth⁻¹()"
        }

        log_fn = {
            "log": "log()",
            "ln": "ln()",
            "log⁻¹": "log⁻¹()"
        }

        def create_button(text, parent, slot):
            button = QPushButton(text, parent)
            button.setFixedSize(80, 30)
            button.setFont(QFont("Arial", 10, QFont.Bold))
            button.pressed.connect(slot)
            return button

        # Create a tab widget
        tab = QTabWidget(self.eqwin)
        tab.setGeometry(10, 10, 370, 280)  # Adjust the geometry as needed
        

        # Create the tab for Trigonometric Functions
        trig_functions_tab = QWidget()
        trig_grid = QGridLayout(trig_functions_tab)

        # Add buttons for trigonometric functions
        row, column = 0, 0
        for name, letter in trigo_fn.items():
            button = create_button(letter, trig_functions_tab, lambda letter=letter: tsymcursor.insertText(letter))
            setattr(self, name.replace('⁻¹', '⁻'), button)  # Set the button as an attribute of the class
            trig_grid.addWidget(button, row, column)
            column += 1
            if column > 3:  # Adjust column limit as per your layout
                column = 0
                row += 1

        trig_functions_tab.setLayout(trig_grid)

        log_functions_tab = QWidget()
        log_grid = QGridLayout(log_functions_tab)

        # Add buttons for trigonometric functions
        row, column = 0, 0
        for name, letter in log_fn.items():
            button = create_button(letter, log_functions_tab, lambda letter=letter: tsymcursor.insertText(letter))
            setattr(self, name.replace('⁻¹', '⁻'), button)  # Set the button as an attribute of the class
            log_grid.addWidget(button, row, column)
            column += 1
            if column > 3:  # Adjust column limit as per your layout
                column = 0
                row += 1

        log_functions_tab.setLayout(trig_grid)

        # Add the Trigonometric Functions tab to the tab widget
        tab.addTab(trig_functions_tab, "Trigonometric Functions")
        tab.addTab(log_functions_tab, "Logarithmic Functions")

        self.eqwin.show()

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
        self.vidLabel.setAlignment(Qt.self.alignCenter)
        self.vidLabel.setStyleSheet("""
color:white;
margin-left:40px;
""")
        self.sLabel = QLabel(self.topframe)
        self.sLabel.setText("|")
        self.sLabel.setFont(QFont("Fira Mono Bold", 30))
        self.sLabel.setAlignment(Qt.self.alignCenter)
        self.sLabel.setStyleSheet("""
color:white;
""")
        self.scLabel = QLabel(self.topframe)
        self.scLabel.setText("Script")
        self.scLabel.setFont(QFont("Fira Mono Bold", 20, QFont.Bold))
        self.scLabel.setAlignment(Qt.self.alignCenter)
        self.scLabel.setStyleSheet("""
margin:0px;
color:white;
margin-right:40px;
""")
        self.pLabel = QLabel(self.topframe)
        self.pLabel.setText("PERSONAL")
        self.pLabel.setFont(QFont("Beware", 13))
        self.pLabel.setAlignment(Qt.self.alignCenter)
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
            popup.setIcon(QtWidgets.QMessageBox.Question)
            popup.setWindowTitle("Unsaved Changes")
            popup.setStyleSheet("""
                                QLabel{
                                font-size:15px;
                                font-weight: bold;
                                }
            QPushButton {
                background-color: light grey; /* Button background color */
                color: black; /* Button text color */
                border: 1px solid #c6c6c6; /* Button border */
                padding: 5px 10px; /* Button padding */
                margin: 5px; /* Button margin */
                border-radius: 5px; /* Button border radius */
                font-family: Arial; /* Button font family */
                font-size: 15px; /* Button font size */
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e0e0e0; /* Hover background color */
                                color:royalblue;
                                border:1px solid royalblue;
            }
            QPushButton:pressed {
                background-color: #d0d0d0; /* Pressed background color */
            }
        """)
            popup.setText("The document has been modified. \nDo you want to save your changes?")
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

    def toggleribbon(self):
        is_visible = self.tab_widget.isVisible()
        self.tab_widget.setVisible(not is_visible)


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
                popup.setStyleSheet("""
                                QLabel{
                                font-size:15px;
                                font-weight: bold;
                                }
            QPushButton {
                background-color: light grey; /* Button background color */
                color: black; /* Button text color */
                border: 1px solid #c6c6c6; /* Button border */
                padding: 5px 10px; /* Button padding */
                margin: 5px; /* Button margin */
                border-radius: 5px; /* Button border radius */
                font-family: Arial; /* Button font family */
                font-size: 15px; /* Button font size */
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e0e0e0; /* Hover background color */
                                color:royalblue;
                                border:1px solid royalblue;
            }
            QPushButton:pressed {
                background-color: #d0d0d0; /* Pressed background color */
            }
        """)
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

    def alignLeftf(self):
        self.text.setAlignment(Qt.AlignLeft)

    def alignRightf(self):
        self.text.setAlignment(Qt.AlignRight)

    def alignCenterf(self):
        self.text.setAlignment(Qt.AlignCenter)

    def alignJustifyf(self):
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
    QPushButton{
    padding:3px;
    margin:0px;
    border-radius:10px;
    }
QPushButton:hover{
background-color:lightgrey;
}
QPushButton:selected{
background-color:grey;
}
QTabWidget::pane { /* The tab widget frame */
            background: white;
            border-radius: 10px;
        }
        QTabWidget::tab-bar {
            alignment: center;
            border-radius:4px;
            background:white;
        }
        
        QTabBar::tab {
            background: white;
            border: 1px solid lightgrey;
            padding: 5px;
            width:150px;
            border-radius:3px;
            font-family: Arial;
            margin-bottom:7px;
        }
        
        QTabBar::tab:selected {
            background: white;
            border-bottom: 3px solid royalblue;
            color:royalblue;
            padding: 5px;
            border-radius:3px;
            margin-bottom:7px;
        }

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

QMainWindow
{
background:#e3e5e9;
border-radius:4px;
}
QMenuBar
{
margin-top:7px;
margin-left:7px;
margin-right:7px;
margin-bottom:3px;
border-radius:8px;
background: royalblue;
width:50px;
padding:3px;
color:white;
}
            QMenuBar::item {
font-size:20px;
alignment: center;
            }
            QMenuBar::item:selected {
            border-radius:4px;
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
