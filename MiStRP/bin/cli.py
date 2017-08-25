import datetime
import PyQt5
import time
import io
import traceback
import json
import re
import os
import console as cnsl
import ctypes.wintypes
import pyexcel as pe
import numpy as np
import pyqtgraph.pyqtgraph as pg
from pyexcel.cookbook import extract_a_sheet_from_a_book as pe_get_sheet
from collections import defaultdict, namedtuple
from PyQt5 import QtCore
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

# Constants
GO_PIC = r'C:\Users\mscales\Desktop\Development\MiStRP\bin\go.png'
ACCUM_PIC = r'C:\Users\mscales\Desktop\Development\MiStRP\bin\accum.png'
QUIT_PIC = r'C:\Users\mscales\Desktop\Development\MiStRP\bin\quit24.png'
SEARCH_PIC = r'C:\Users\mscales\Desktop\Development\MiStRP\bin\search.png'
SETTINGS_PIC = r'C:\Users\mscales\Desktop\Development\MiStRP\bin\settings.png'
CALL_SLA_ARG = "C:\\users\\mscales\\desktop\\development\\daily sla parser - automated version\\bin\\cli.py"
ACCUM_ARG = "C:\\users\\mscales\\desktop\\development\\MSSQLupdater\\bin\\client_counter.py"
SPREADSHEET_VIEWER_FILE_TEMPLATE = r'C:\users\\mscales\desktop\Development\Daily SLA Parser - Automated Version\Output\{0}v2_Incoming DID Summary.xlsx'
ACCUM_VIEWER_FILE = r'C:\\users\\mscales\\desktop\\development\\MSSQLupdater\\output\\temp_file.xlsx'
TEST_SPREADSHEET = r'C:\\users\\mscales\\desktop\\development\\MSSQLupdater\\output\\temp_file.xlsx'


# TODO: fix this shit
# Need to fix Scatterplotitem>Opts = False-> True ** changed to fix qpixmap error **


class MainFrame(QMainWindow):

    application_color = pyqtSignal(str, name='selected color')
    application_font = pyqtSignal(QFont, name='selected font')
    application_style = pyqtSignal(str, name='application style')

    def __init__(self, parent=None):
        super(MainFrame, self).__init__(parent)

        # Main Frame widget
        self.te = QTextEdit(self)
        self.main_widget = self.main_widget_factory()
        self.setCentralWidget(self.main_widget)

        (self.global_constants,
         self.local_constants) = self.constants_factory()

        self.settings_button = self.settings_button_factory()

        self.status_bar()

        # Buttons
        self.report_button = QAction(QIcon(self.global_constants.GO_PIC), "Run Program", self)
        self.report_button.setShortcut('Ctrl+R')
        self.report_button.setStatusTip('Run SLA Program')
        self.report_button.triggered.connect(self.call_sla)

        self.accumulator_button = QAction(QIcon(self.global_constants.ACCUM_PIC), "Accumulator", self)
        self.accumulator_button.setShortcut('Ctrl+A')
        self.accumulator_button.setStatusTip('Run accumulator')
        self.accumulator_button.triggered.connect(self.accumulate_spreadsheet_data)

        self.spreadsheet_viewer_button = QAction(QIcon(self.global_constants.SEARCH_PIC), "View spreadsheets", self)
        self.spreadsheet_viewer_button.setShortcut('Ctrl+V')
        self.spreadsheet_viewer_button.setStatusTip('View spreadsheets for a given date range.')
        self.spreadsheet_viewer_button.triggered.connect(self.spreadsheet_viewer)

        self.about_action = QAction("&About", self)
        self.about_action.setStatusTip('Information about this program')
        self.about_action.triggered.connect(self.about)

        self.about_qt_action = QAction("About &Qt", self)
        self.about_qt_action.setStatusTip('GUI application information')
        self.about_qt_action.triggered.connect(qApp.aboutQt)

        self.exit_action = QAction(QIcon(self.global_constants.QUIT_PIC), 'Exit', self)
        self.exit_action.setStatusTip('About Qt library')
        self.exit_action.triggered.connect(self.close)

        # Menus
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu('&File')
        file_menu.addAction(self.report_button)
        file_menu.addAction(self.spreadsheet_viewer_button)
        file_menu.addAction(self.accumulator_button)
        file_menu.addAction(self.exit_action)

        # edit_menu = menu_bar.addMenu('&Edit')
        # edit_menu.addAction(self.settings_button)

        about_menu = menu_bar.addMenu('&About')
        about_menu.addAction(self.about_action)
        about_menu.addAction(self.about_qt_action)

        # Toolbar
        toolbar = QToolBar()
        toolbar.setIconSize(QSize(50, 50))
        toolbar.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon | QtCore.Qt.AlignLeading)
        self.addToolBar(QtCore.Qt.LeftToolBarArea, toolbar)
        toolbar.addAction(self.report_button)
        toolbar.addAction(self.spreadsheet_viewer_button)
        toolbar.addAction(self.accumulator_button)
        toolbar.addWidget(self.settings_button)
        toolbar.addAction(self.exit_action)

        # Process Window Bindings
        self.sla_calendar = None
        self.sv_calendar = None
        self.acc_calendar = None

        self.reset_binding_dict = {'daily sla parser - automated version': 'sla_calendar',
                                   'MSSQLupdater': 'acc_calendar',
                                   'viewer': 'sv_calendar',
                                   'slicer': 'sla_slicer'}

        # Main frame layout
        self.resize(1000, 500)
        self.center_frame()
        self.setWindowTitle('Mike\'s Streamlined Reporting Program')

    def center_frame(self):
        screen_dimensions = self.frameGeometry()
        center_dimensions = QDesktopWidget().availableGeometry().center()
        screen_dimensions.moveCenter(center_dimensions)
        self.move(screen_dimensions.topLeft())

    def about(self):
        QMessageBox.about(self,
                          "About MiStRP",
                          "Version 1.0\n"
                          "Developed by Michael Scales\n"
                          "Application to run python modules")

    def status_bar(self):
        self.statusBar().showMessage("Ready")

    def closeEvent(self, event):
        QApplication.quit()
        # if QMessageBox.question(None,
        #                         'Quit MiStRP?',
        #                         "Are you sure to quit?",
        #                         QMessageBox.Yes | QMessageBox.No,
        #                         QMessageBox.No) == QMessageBox.Yes:
        #     QApplication.quit()
        # else:
        #     event.ignore()

    def call_sla(self):
        if self.sla_calendar is None:
            sla_program = ProcessMenu(parent=self, process='sla_program', constants=self.global_constants)
            save_widget = SaveWidget(self)
            save_widget.status_message.connect(self.statusBar().showMessage)
            self.sla_calendar = SplitterFrame(None, sla_program, save_widget)
            sla_program.save_items.connect(save_widget.set_item_list)
            save_widget.save_request_list.connect(sla_program.get_save_info)
            sla_program.process_menu_std_out.connect(self.write_to_central_widget)
            self.sla_calendar.exit_status.connect(self.reset_binding)
        else:
            self.sla_calendar.raise_()
            # child_frame = ChildFrame(self, 'sla_program')  # ProcessMenu
            # child_frame.show()
            # layout = QVBoxLayout()
            # layout.addWidget(child_frame)
            # self.sla_calendar.setLayout(layout)
            # self.sla_calendar.show()
            # print("ChildFrame Children - qsplitter")
            # print(self.sla_calendar.children()[0])
            # print("Qsplitter Children")
            # print(self.sla_calendar.children()[0].children())
            # print("last")
            # print(self.sla_calendar.children()[0].children()[1].children())
            # string_handle = self.sla_calendar.children()[0].children()[1]
            # string_handle.string_output.connect(self.write_to_central_widget)

            # self.sla_calendar = ProcessMenu(self, 'sla_program')
            # self.sla_calendar.string_output.connect(self.write_to_central_widget)

    def spreadsheet_viewer(self):
        if self.sv_calendar is None:
            sv_program = ProcessMenu(parent=self, process='viewer', constants=self.global_constants)
            save_widget = SaveWidget(self)
            save_widget.status_message.connect(self.statusBar().showMessage)
            self.sv_calendar = SplitterFrame(None, sv_program, save_widget)
            sv_program.save_items.connect(save_widget.set_item_list)
            save_widget.save_request_list.connect(sv_program.get_save_info)
            sv_program.process_menu_std_out.connect(self.write_to_central_widget)
            self.sv_calendar.exit_status.connect(self.reset_binding)
        else:
            self.sv_calendar.raise_()

    def accumulate_spreadsheet_data(self):
        if self.acc_calendar is None:
            acccum_program = ProcessMenu(parent=self, process='MSSQLupdater', constants=self.global_constants)
            save_widget = SaveWidget(self)
            save_widget.status_message.connect(self.statusBar().showMessage)
            self.acc_calendar = SplitterFrame(None, acccum_program, save_widget)
            acccum_program.save_items.connect(save_widget.set_item_list)
            save_widget.save_request_list.connect(acccum_program.get_save_info)
            acccum_program.process_menu_std_out.connect(self.write_to_central_widget)
            self.acc_calendar.exit_status.connect(self.reset_binding)
        else:
            self.acc_calendar.raise_()

    def open_tab(self, args):
        for index in range(self.main_widget.count()):
            if self.main_widget.widget(index).windowTitle() == args.title:
                self.main_widget.setCurrentIndex(index)
                return
        tab = TableWidget(file=args.file, page=args.page, window_title=args.title, special_case=args.special_case)
        self.main_widget.addTab(tab, args.title)

    def close_tab(self, index):
        self.main_widget.widget(index).close()
        self.main_widget.removeTab(index)

    def reset_binding(self, process):
        setattr(self, self.reset_binding_dict[process], None)

    def write_to_central_widget(self, string_text):
        cursor = self.te.textCursor()
        cursor.movePosition(cursor.End)
        cursor.insertText(string_text)
        self.te.ensureCursorVisible()

    def main_widget_factory(self):
        main_widget = QTabWidget()
        main_widget.setTabsClosable(True)
        main_widget.tabCloseRequested[int].connect(self.close_tab)
        main_widget.addTab(self.te, "MiStRP")
        main_widget.tabBar().tabButton(0, QTabBar.RightSide).resize(0, 0)
        return main_widget

    def settings_button_factory(self):
        sla_client_dict_settings_args = self.local_constants(
            file=r'C:\Users\mscales\Desktop\Development\Daily SLA Parser - Automated Version\bin\CONFIG.xlsx',
            page='CLIENT_DICT',
            title='SLA SETTINGS/CLIENT_DICT',
            special_case=True
        )
        sla_constants_settings_args = self.local_constants(
            file=r'C:\Users\mscales\Desktop\Development\Daily SLA Parser - Automated Version\bin\CONFIG.xlsx',
            page='CONSTANTS',
            title='SLA SETTINGS/CONSTANTS',
            special_case=True
        )
        sla_client_info_settings_args = self.local_constants(
            file=r'C:\Users\mscales\Desktop\Development\Daily SLA Parser - Automated Version\bin\CONFIG.xlsx',
            page='CLIENT LIST INFO',
            title='SLA SETTINGS/CLIENT LIST INFO',
            special_case=True
        )
        sla_settings = QMenu("sla settings", self)
        sla_constants_action = QAction(QIcon(self.global_constants.SETTINGS_PIC), "SLA SETTINGS/CLIENT_DICT", self)
        sla_client_dict_action = QAction(QIcon(self.global_constants.SETTINGS_PIC), "SLA SETTINGS/CONSTANTS", self)
        sla_client_info_action = QAction(QIcon(self.global_constants.SETTINGS_PIC), "SLA SETTINGS/CLIENT LIST INFO",
                                         self)
        sla_constants_action.triggered.connect(lambda: self.open_tab(sla_client_dict_settings_args))
        sla_client_dict_action.triggered.connect(lambda: self.open_tab(sla_constants_settings_args))
        sla_client_info_action.triggered.connect(lambda: self.open_tab(sla_client_info_settings_args))
        sla_settings.addAction(sla_constants_action)
        sla_settings.addAction(sla_client_dict_action)
        sla_settings.addAction(sla_client_info_action)

        accum_settings_args = self.local_constants(
            file=r'C:\Users\mscales\Desktop\Development\MSSQLupdater\bin\client_list_file.xlsx',
            title='Accumulator Settings/Sheet1',
            special_case=True
        )
        accum_settings = QMenu("accum settings", self)
        accum_action = QAction(QIcon(self.global_constants.SETTINGS_PIC), "Accumulator Settings", self)
        accum_action.triggered.connect(lambda: self.open_tab(accum_settings_args))
        accum_settings.addAction(accum_action)

        self_settings_args = self.local_constants(
            file=r'C:\Users\mscales\Desktop\Development\MiStRP\bin\config.xlsx',
            page='CONFIG',
            title='MiStRP Settings/CONFIG',
            special_case=True
        )
        self_settings = QMenu("MiStRP Settings", self)
        self_action = QAction(QIcon(self.global_constants.SETTINGS_PIC), "Self Settings", self)
        self_action.triggered.connect(lambda: self.open_tab(self_settings_args))
        self_settings.addAction(self_action)

        SettingMenu = QMenu()
        SettingMenu.addMenu(sla_settings)
        SettingMenu.addMenu(accum_settings)
        SettingMenu.addMenu(self_settings)

        SettingButton = QToolButton()
        SettingButton.setIcon(QIcon(self.global_constants.SETTINGS_PIC))
        tab = SettingsWidget(parent=self)
        tab.selected_color.connect(self.application_color.emit)
        tab.selected_font.connect(self.application_font.emit)
        tab.selected_style.connect(self.application_style.emit)
        SettingButton.clicked.connect(lambda: self.main_widget.addTab(tab, 'Display Settings'))
        SettingButton.setPopupMode(1)
        SettingButton.setMenu(SettingMenu)
        return SettingButton

    def constants_factory(self):
        Node = namedtuple('Node', 'title file page special_case')
        Node.__new__.__defaults__ = (None,) * len(Node._fields)
        constants = namedtuple('Node',
                               'GO_PIC ACCUM_PIC QUIT_PIC SEARCH_PIC SETTINGS_PIC CALL_SLA_ARG ACCUM_ARG '
                               'SPREADSHEET_VIEWER_FILE_TEMPLATE ACCUM_VIEWER_FILE TEST_SPREADSHEET CALL_SLICER_ARG')
        constants.__new__.__defaults__ = (None,) * len(constants._fields)
        SELF_PATH = os.path.dirname(path.dirname(path.abspath(__file__)))
        constants_sheet = pe.get_dict(file_name='%s/bin/config.xlsx' % SELF_PATH, name_columns_by_row=0)
        index_dict = {}
        for index, item in enumerate(constants_sheet['Constant']):
            index_dict[item] = index
        GO_PIC = constants_sheet['Argument'][index_dict['GO_PIC']]
        ACCUM_PIC = constants_sheet['Argument'][index_dict['ACCUM_PIC']]
        QUIT_PIC = constants_sheet['Argument'][index_dict['QUIT_PIC']]
        SEARCH_PIC = constants_sheet['Argument'][index_dict['SEARCH_PIC']]
        SETTINGS_PIC = constants_sheet['Argument'][index_dict['SETTINGS_PIC']]
        CALL_SLA_ARG = constants_sheet['Argument'][index_dict['CALL_SLA_ARG']]
        ACCUM_ARG = constants_sheet['Argument'][index_dict['ACCUM_ARG']]
        SPREADSHEET_VIEWER_FILE_TEMPLATE = constants_sheet['Argument'][index_dict['SPREADSHEET_VIEWER_FILE_TEMPLATE']]
        ACCUM_VIEWER_FILE = constants_sheet['Argument'][index_dict['ACCUM_VIEWER_FILE']]
        TEST_SPREADSHEET = constants_sheet['Argument'][index_dict['TEST_SPREADSHEET']]
        CALL_SLICER_ARG = constants_sheet['Argument'][index_dict['CALL_SLICER_ARG']]
        return constants(GO_PIC=GO_PIC, ACCUM_PIC=ACCUM_PIC,
                         QUIT_PIC=QUIT_PIC, SEARCH_PIC=SEARCH_PIC,
                         SETTINGS_PIC=SETTINGS_PIC,
                         CALL_SLA_ARG=CALL_SLA_ARG,
                         ACCUM_ARG=ACCUM_ARG,
                         SPREADSHEET_VIEWER_FILE_TEMPLATE=SPREADSHEET_VIEWER_FILE_TEMPLATE,
                         ACCUM_VIEWER_FILE=ACCUM_VIEWER_FILE,
                         TEST_SPREADSHEET=TEST_SPREADSHEET,
                         CALL_SLICER_ARG=CALL_SLICER_ARG), Node


class SettingsWidget(QWidget):
    # TODO: custom styles, etc
    selected_color = pyqtSignal(str, name='selected color')
    selected_font = pyqtSignal(QFont, name='selected color')
    selected_style = pyqtSignal(str, name='style choice')

    def __init__(self, parent=None):
        super(SettingsWidget, self).__init__(parent)
        grid_input = QGridLayout()
        fontChoice = QPushButton('Font', self)
        fontChoice.clicked.connect(self.font_choice)
        fontColor = QPushButton('Font bg Color', self)
        fontColor.clicked.connect(self.color_picker)
        checkBox = QCheckBox('Enlarge Window', self)
        checkBox.stateChanged.connect(self.enlarge_window)
        self.styleChoice = QLabel("Windows Vista", self)

        comboBox = QComboBox(self)
        comboBox.addItem("motif")
        comboBox.addItem("Windows")
        comboBox.addItem("cde")
        comboBox.addItem("Plastique")
        comboBox.addItem("Cleanlooks")
        comboBox.addItem("windowsvista")

        comboBox.activated[str].connect(self.style_choice)
        grid_input.addWidget(fontChoice, 0, 0)
        grid_input.addWidget(fontColor, 0, 1)
        grid_input.addWidget(checkBox, 1, 0)
        grid_input.addWidget(self.styleChoice, 1, 1)
        grid_input.addWidget(comboBox, 1, 2)
        self.setLayout(grid_input)

    def font_choice(self):
        font, valid = QFontDialog.getFont()
        if valid:
            self.selected_font.emit(font)

    def color_picker(self):
        color = QColorDialog.getColor(initial=QtCore.Qt.darkBlue)
        self.selected_color.emit("QMainWindow { background-color: %s}" % color.name())

    def style_choice(self, text):
        self.styleChoice.setText(text)
        QApplication.setStyle(QStyleFactory.create(text))
        self.selected_style.emit(text)

    def enlarge_window(self, state):
        # TODO: This needs to modify the QMainWindows
        if state == QtCore.Qt.Checked:
            self.setGeometry(50, 50, 1000, 600)
        else:
            self.setGeometry(50, 50, 500, 300)


class ProcessButtons(QListWidget):
    def __init__(self, parent=None, process=None):
        super(ProcessButtons, self).__init__(parent)
        print(process)
        self.process = None
        self.process_list = {'daily sla parser - automated version': ["Override Report", "stuff2", "stuff3"],
                             'viewer': ['none', 'none1', 'none2'],
                             'MSSQLupdater': ['none3', 'none4', 'none5'],
                             'sla_slicer.py': ['none6', 'none7', 'none8']}
        self.init_kwds(process)
        self.itemClicked.connect(self.change_transmit_arguments)
        self.argument_list = []
        self.show()

    def init_kwds(self, process):
        self.process = process
        # print(process)
        # print(self.process_list[process])
        for option in self.process_list[process]:
            self.addItem(option)

    def change_transmit_arguments(self, argument):
        list_argument = argument.text()
        if list_argument in self.argument_list:
            argument.setSelected(False)
            self.argument_list.remove(list_argument)
        else:
            self.argument_list.append(list_argument)

    def get_arguments(self):
        return self.argument_list


class ProcessMenu(QMainWindow):
    process_menu_std_out = pyqtSignal(str, name='menu_std_out')
    save_items = pyqtSignal(defaultdict, name='save_items')
    return_save_info = pyqtSignal(defaultdict, name='return_save_info')
    progress_update = pyqtSignal(float, name='progress_made')

    def __init__(self, parent=None, process=None, constants=None):
        super(ProcessMenu, self).__init__(parent)
        self.process = None
        self.program_name = None
        # Member Attributes
        self.constants = constants
        self.__callable = None
        self.progress = 0
        self.args = {}
        self.__kwds = {  # can be simplified into one dict
            'sla_program': False,
            'MSSQLupdater': False,
            'viewer': False,
            'slicer': False,
            'name': None
        }
        self.__process_args = {
            'sla_program': self.constants.CALL_SLA_ARG,
            'MSSQLupdater': self.constants.ACCUM_ARG,
            'viewer': 'viewer',
            'slicer': self.constants.CALL_SLICER_ARG
        }
        self.init_kwds(process)

        self.__excel_dock = ExcelDockWidget(self,
                                            window_name='Excel Dock',
                                            widget1=ExcelTabContainer(title='Excel Dock Widget', parent=self),
                                            hide_button_name='Show Excel')
        self.__buttons = ButtonsDockWidget(self,
                                           window_name='Button Dock',
                                           widget1=ProcessButtons(process=self.program_name),
                                           hide_button_name='Show Buttons')
        self.__graph_dock = GraphicsDockWidget(self,
                                               window_name='Graphics Dock',
                                               widget1=GraphData('Graph'),
                                               hide_button_name='Show Graph')
        self.central_wig_text = TextWindow(self, title='Error Log')
        self.central_wig_progress = QProgressBar(self)
        self.__calendar_dock = CalendarDockWidget(self,
                                                  window_name='Calendar Dock',
                                                  widget1=MyCalendarWidget('date1'),
                                                  widget2=MyCalendarWidget('date2'),
                                                  button=QPushButton('Run Report'),
                                                  hide_button_name='Show Calendars')
        self.__calendar_dock.visibilityChanged.connect(lambda: self.__graph_dock.show())
        self.progress_update.connect(self.central_wig_progress.setValue)

        self.central_wig_progress.setMaximumHeight(100)
        self.central_wig_progress.setRange(0, 12056)

        splitter = SplitterFrame(self,
                                 widget1=self.central_wig_progress,
                                 widget2=self.central_wig_text,
                                 orientation='Vertical',
                                 arrow_direction='left')
        bottom_box = QMainWindow()
        bottom_box.setCentralWidget(splitter)
        bottom_box.addDockWidget(Qt.RightDockWidgetArea, self.__buttons)
        bottom_box.addDockWidget(Qt.LeftDockWidgetArea, self.__graph_dock)
        bottom_box.addDockWidget(Qt.BottomDockWidgetArea, self.__excel_dock)
        bottom_box.addDockWidget(Qt.TopDockWidgetArea, self.__calendar_dock)

        self.__calendar_dock.calendar_date_range.connect(self.prepare_process)
        self.__excel_dock.tab_container_data.connect(self.__graph_dock.add_curve)
        self.setCentralWidget(bottom_box)
        self.center_frame()
        self.show()

    def prepare_process(self, start_date, end_date):
        # allow Qt's blocking of signals paradigm to control flow

        if self.signalsBlocked() is not True:
            self.exec_process([self.process,
                               start_date.strftime("%m%d%Y"),
                               end_date.strftime("%m%d%Y")])
            self.__buttons.hide()

    def read_std_out(self):
        qprocess = self.sender()
        characters = {"Converting:": '\nConverting:', 'Writing converted': '\nWriting converted'}
        insert_text_string = str(qprocess.readAllStandardOutput(), encoding='UTF-8')
        for ch in characters:
            if ch in insert_text_string:
                insert_text_string = insert_text_string.replace(ch, characters[ch])
        self.process_menu_std_out.emit(insert_text_string)

    def read_std_err(self):
        qprocess = self.sender()
        self.error_log = str(qprocess.readAllStandardError(), encoding='UTF-8')

    def exec_process(self, params):
        callable_program = self.get_callable_token(params[0], '\\')
        args, window_title = self.make_args(params[1], params[2], callable_program)
        process = RunProcess()
        self.__callable = process
        thread_worker = ProgressWorker(self, process)
        thread_worker.update_progress.connect(self.set_progress)
        self.process_menu_std_out.connect(thread_worker.accumulate_std_out)
        process.qprocess.started.connect(thread_worker.start)
        process.qprocess.readyReadStandardOutput.connect(self.read_std_out)
        process.qprocess.readyReadStandardError.connect(self.read_std_err)
        process.do_work(params)
        process.error_signal.connect(lambda: self.central_wig_text.append(self.error_log))
        process.finished_signal.connect(
            lambda: self.__excel_dock.make_tabs(args))
        process.finished_signal.connect(
            lambda: self.save_items.emit(args))
        process.finished_signal.connect(
            lambda: self.__graph_dock.make_graph(params[1], params[2]))

    def get_callable_token(self, line_to_split, delimiter, callable_string=None):
        arg_tokens = line_to_split.split(delimiter)
        for arg in arg_tokens:
            if arg == self.program_name:
                callable_string = arg
        return callable_string

    def make_args(self, start_date, end_date, callable_program, fmt='%m%d%Y'):
        args = defaultdict(list)
        if callable_program in ('MSSQLupdater'):
            window_title = 'Accumulater'
            sheet_title = r'Accumulater {0} to {1}'.format(start_date, end_date)
            args[sheet_title].append(ACCUM_VIEWER_FILE)
            args[sheet_title].append(start_date)
            args[sheet_title].append(end_date)
        else:
            window_title = 'Daily Reports'
            try:
                start_date = datetime.datetime.strptime(start_date, fmt)
                end_date = datetime.datetime.strptime(end_date, fmt)
            except TypeError:
                pass
            finally:
                first_report_date = start_date
            while start_date <= end_date:
                start_date_string = start_date.strftime(fmt)
                spreadsheet = SPREADSHEET_VIEWER_FILE_TEMPLATE.format(start_date_string)
                sheet_title = r'Report {0}'.format(start_date_string)
                args[sheet_title].append(spreadsheet)
                args[sheet_title].append(first_report_date)
                args[sheet_title].append(end_date)
                self.args[sheet_title] = spreadsheet
                start_date = start_date + datetime.timedelta(days=1)
        return args, window_title

    def init_kwds(self, process):
        try:
            if process in self.__kwds:
                self.__kwds[process] = True
                self.__kwds['name'] = process
                self.process = self.__process_args[process]
                # TODO: clean this up
                tokens = self.__process_args[process].split('\\')
                try:
                    self.program_name = str(tokens[-3])
                except IndexError:
                    self.program_name = str(tokens[0])
                    # print(self.program_name)
        except KeyError:
            print("failed to import process settings... check ProcessMenu->init_kwds")
            sys.exit()

    def get_save_info(self):
        # return_dict = {}
        # for (k, v) in self.args.items():
        #     # print('{0}\n'.format(item))
        #     if k in list:
        #         return_dict[k] = v
        # self.return_save_info.emit(return_dict)
        return self.args

    def set_progress(self, progress):
        self.central_wig_progress.setValue(progress)

    def get_text_size(self):
        return self.progress

    def get_callable(self):
        return self.__callable

    def center_frame(self):
        screen_dimensions = self.frameGeometry()
        center_dimensions = QDesktopWidget().availableGeometry().center()
        screen_dimensions.moveCenter(center_dimensions)
        self.move(screen_dimensions.topLeft())


class ProgressWorker(QThread):
    update_progress = pyqtSignal(float, name='update_progress')

    def __init__(self, parent, proc):
        super(ProgressWorker, self).__init__(parent)
        self.proc = proc
        self.proc.finished_signal.connect(
            lambda: self.update_progress.emit(self.total))
        self.progress = 0
        self.buffer = 0
        self.total = 12056

    def run(self):
        while (self.progress <= self.total) and self.proc.qprocess.pid() is not None:
            if self.progress <= self.buffer:
                self.progress += self.buffer * .01
            self.update_progress.emit(self.progress)
            time.sleep(0.2)

    def accumulate_std_out(self, std_out):
        self.buffer += len(std_out)


class SplitterFrame(QWidget):
    exit_status = pyqtSignal(str, name='exit_status')
    status_message = pyqtSignal(str, name='status_message')

    def __init__(self, parent, widget1, widget2, orientation='horizontal', arrow_direction='right'):
        super(SplitterFrame, self).__init__(parent)
        if orientation == 'horizontal':
            orientation = Qt.Horizontal
        else:
            orientation = Qt.Vertical
        if arrow_direction == 'right':
            arrow = QtCore.Qt.RightArrow
        else:
            arrow = QtCore.Qt.LeftArrow

        self.splitter = QSplitter(orientation)
        self.splitter.addWidget(self.center_widget(widget1))
        self.splitter.addWidget(widget2)
        self.splitter.setSizes([1, 0])

        layout = QVBoxLayout(self)
        layout.addWidget(self.splitter, alignment=QtCore.Qt.AlignBottom)
        handle = self.splitter.handle(1)
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        button = QToolButton(handle)
        button.setText("Save")
        button.setCheckable(True)
        button.setArrowType(arrow)
        button.clicked.connect(self.handle_splitter_button)
        layout.addWidget(button, alignment=QtCore.Qt.AlignCenter)
        handle.setLayout(layout)
        self.center_frame()
        self.show()

    def handle_splitter_button(self):
        if not all(self.splitter.sizes()):
            self.splitter.setSizes([1, 1])
        else:
            self.splitter.setSizes([1, 0])

    def center_widget(self, widget):
        centered_widget = QWidget()
        temp_layout = QHBoxLayout()
        temp_layout.addWidget(widget, alignment=QtCore.Qt.AlignCenter)
        centered_widget.setLayout(temp_layout)
        return centered_widget

    def closeEvent(self, event):
        self.exit_status.emit(self.children()[1].children()[1].children()[1].program_name)
        event.accept()

    def center_frame(self):
        screen_dimensions = self.geometry()
        center_dimensions = QDesktopWidget().availableGeometry().topLeft()
        screen_dimensions.moveCenter(center_dimensions)
        self.move(screen_dimensions.center())


class HorizontalWidget(QWidget):
    def __init__(self, parent=None, widget1=None, widget2=None):
        super(HorizontalWidget, self).__init__(parent)
        if widget1 and widget2:
            layout = QHBoxLayout()
            layout.addWidget(widget1, alignment=QtCore.Qt.AlignCenter)
            layout.addWidget(widget2, alignment=QtCore.Qt.AlignCenter)
            exp_layout = QVBoxLayout()
            exp_layout.addLayout(layout)
        else:
            layout = QHBoxLayout()
            if widget1:
                layout.addWidget(widget1, alignment=QtCore.Qt.AlignCenter)
            if widget2:
                layout.addWidget(widget2, alignment=QtCore.Qt.AlignCenter)
            exp_layout = QVBoxLayout()
            exp_layout.addLayout(layout)
        self.setLayout(exp_layout)

    def show_widgets(self, show_widget='Both'):
        if show_widget == 'Both':
            self.children()[1].show()
            self.children()[2].show()
        elif show_widget == 'left':
            self.children()[1].show()
        else:
            self.children()[2].show()


class DockWidget(QDockWidget):
    def __init__(self, parent=None, window_name=None, widget1=None, widget2=None, button=None, hide_button_name=None):
        super(DockWidget, self).__init__(window_name, parent)
        self.setFeatures(QDockWidget.DockWidgetFloatable |
                         QDockWidget.DockWidgetMovable |
                         QDockWidget.DockWidgetClosable)
        dock_widgets = HorizontalWidget(widget1=widget1, widget2=widget2)
        v_layout = QVBoxLayout()
        v_layout.addWidget(dock_widgets, alignment=QtCore.Qt.AlignCenter)
        if button:
            self.button = button
            v_layout.addWidget(self.button, alignment=QtCore.Qt.AlignBottom)
        if hide_button_name:
            self.hide_button = QPushButton(str(hide_button_name))
            self.splitter = QSplitter(Qt.Horizontal)
            button_layout = QVBoxLayout()
            button_layout.addWidget(self.hide_button, alignment=QtCore.Qt.AlignCenter)
            set_widget = QWidget()
            set_widget.setLayout(button_layout)
            self.splitter.addWidget(set_widget)
            set_widget = QWidget()
            set_widget.setLayout(v_layout)
            self.splitter.addWidget(set_widget)
            self.splitter.setSizes([1, 0])
            self.splitter.show()
            handle = self.splitter.handle(1)
            self.layout = QVBoxLayout()
            self.layout.setContentsMargins(0, 0, 0, 0)
            self.switch_frame = QAction(handle)
            self.switch_frame.triggered.connect(self.handle_splitter)
            self.hide_button.clicked.connect(self.switch_frame.trigger)
            self.setWidget(self.splitter)
        else:
            set_widget = QWidget()
            set_widget.setLayout(v_layout)
            self.setWidget(set_widget)
        self.show()

    def handle_splitter(self):
        if self.splitter.sizes()[0] > 0:
            self.splitter.setSizes([0, 1])
        else:
            self.splitter.setSizes([1, 0])

    def closeEvent(self, event):
        event.ignore()
        if self.isFloating():
            self.setFloating(False)
        if self.splitter.sizes()[1] > 0:
            self.splitter.setSizes([1, 0])


class GraphicsDockWidget(DockWidget):
    def __init__(self, parent, window_name=None,
                 widget1=None, widget2=None,
                 button=None, hide_button_name=None):
        super(GraphicsDockWidget, self).__init__(parent, window_name, widget1,
                                                 widget2, button, hide_button_name)

    def make_graph(self, date1, date2):
        graph = GraphData()
        graph.create_graph(date1, date2)
        graph.show()
        self.setWidget(graph)
        self.show()

    def add_curve(self, plot_data):
        self.children()[-1].set_y_axis(plot_data)
        # print(self.children())


class MyStringAxis(pg.AxisItem):
    def __init__(self, xdict, *args, **kwargs):
        super(MyStringAxis, self).__init__(*args, **kwargs)
        self.x_values = np.asarray(xdict.keys())
        self.x_strings = xdict.values()

    def get_x_axis_list(self):
        return list(self.x_values.tolist())

    def tickStrings(self, values, scale, spacing):
        strings = []
        for v in values:
            # vs is the original tick value
            vs = v * scale
            # if we have vs in our values, show the string
            # otherwise show nothing
            if vs in self.x_values:
                # Find the string with x_values closest to vs
                vstr = self.x_strings[np.abs(self.x_values - vs).argmin()]
            else:
                vstr = ""
            strings.append(vstr)
        return strings


class GraphData(QWidget):
    def __init__(self, title='Default', parent=None):
        super(GraphData, self).__init__(parent)
        self.plot = None
        self.current_x_axis = None
        self.curves = []
        self.win = pg.GraphicsWindow(title)
        layout = QVBoxLayout()
        layout.addWidget(self.win, alignment=QtCore.Qt.AlignVCenter)
        self.setLayout(layout)

    def make_xdict(self, date1, date2, xlist=None, fmt='%m%d%Y'):
        start_date = datetime.datetime.strptime(date1, fmt)
        end_date = datetime.datetime.strptime(date2, fmt)
        if xlist is None:
            xlist = []
        while start_date <= end_date:
            xlist.append(start_date.strftime('%A\n%m%d%y'))
            start_date += datetime.timedelta(days=1)
        return dict(enumerate(xlist))

    def create_graph(self, date1, date2):
        self.set_x_axis(date1, date2)
        self.plot = self.win.addPlot(axisItems={'bottom': self.current_x_axis},
                                     name='something', clickable=True)
        self.show()

    def set_x_axis(self, start_date, end_date):
        xdict = self.make_xdict(start_date, end_date)
        string_axis = MyStringAxis(xdict, orientation='bottom')
        string_axis.setTicks([xdict.items()])
        self.current_x_axis = string_axis

    def set_y_axis(self, plot_objects=None):
        list_of_keys = self.current_x_axis.get_x_axis_list()
        for plot_object in plot_objects:
            #     self.plot = self.win.addPlot(axisItems={'bottom': self.current_x_axis},
            #                                  name='something',
            #                                  clickable=True)
            for curve in plot_object.data.items():
                curve_name = str(curve[0])
                data_points = curve[1]
                y_values = []
                for data_pt in data_points:
                    if data_pt is not None:
                        y_values.append(int(data_pt.split('-')[1]))
                new_line = self.plot.plot(list_of_keys, y_values, pen=QColor(28, 78, 99))
                new_line.curve.setClickable(True)
                new_line.sigClicked.connect(self.clicked)
                self.curves.append(new_line)

    def clicked(self):
        print('clicked')


class PlotData:
    def __init__(self, headers):
        self.curves = None
        self.data = defaultdict(list)
        self.name = None
        self.create_curves(headers)

    def create_curves(self, headers):
        self.curves = headers

    def make_data(self, data):
        self.name = data.pop(0)
        init_keys = self.curves
        for key in init_keys:
            self.data[key].append(None)

    def add_data(self, data):
        pass

    def get_name(self):
        return self.name

    def __str__(self):
        # for curve in self.curves.items():
        print(self.name)
        print(self.curves)
        print(self.data)


class CalendarDockWidget(DockWidget):
    calendar_date_range = pyqtSignal(datetime.date, datetime.date, name='calendar_date_range')
    ready_to_send = pyqtSignal(bool, name='ready_to_send')

    def __init__(self, parent, window_name=None,
                 widget1=None, widget2=None,
                 button=None, hide_button_name=None):
        super(CalendarDockWidget, self).__init__(parent, window_name, widget1,
                                                 widget2, button, hide_button_name)
        self.__date1 = None
        self.__date2 = None
        widget1.updated_date.connect(self.update_dates)
        widget2.updated_date.connect(self.update_dates)
        button.setEnabled(False)
        button.clicked.connect(self.emit_dates)
        self.ready_to_send.connect(button.setEnabled)

    def update_dates(self, date, chc):
        if chc == 'date1':
            self.__date1 = date.toPyDate()
            self.ready_to_send.emit(True)
        if chc == 'date2':
            self.__date2 = date.toPyDate()

    def emit_dates(self):
        if self.__date2 is None:
            self.__date2 = self.__date1
        if self.__date1 <= self.__date2:
            self.calendar_date_range.emit(self.__date1, self.__date2)
            self.hide()


class MyCalendarWidget(QWidget):
    updated_date = pyqtSignal(QtCore.QDate, str, name='updated_date')

    def __init__(self, name, parent=None):
        super(MyCalendarWidget, self).__init__(parent)
        self.name = name
        self.cal = QCalendarWidget(self)
        initial_date = QtCore.QDate.currentDate().addDays(-1)
        self.cal.setSelectedDate(initial_date)
        self.cal.setGridVisible(True)
        self.cal.clicked[QtCore.QDate].connect(self.show_date)
        self.text_area = QLabel(self)
        date = self.cal.selectedDate()
        self.text_area.setText(date.toString())

        hbox_layout = QHBoxLayout()
        hbox_layout.addWidget(self.cal)
        vbox_layout = QVBoxLayout()
        vbox_layout.addLayout(hbox_layout)
        vbox_layout.addWidget(self.text_area)
        self.setLayout(vbox_layout)

    def show_date(self, date):
        self.updated_date.emit(date, self.name)
        self.text_area.setText(date.toString())


class SaveWidget(QWidget):
    save_request_list = pyqtSignal(list, name='save_request_list')
    status_message = pyqtSignal(str, name='save_status')

    def __init__(self, parent):
        super(SaveWidget, self).__init__(parent)
        self.setWindowTitle('MiStRP Explorer')

        save_button = QPushButton("save", self)
        save_button.clicked.connect(self.save_event)

        self.list_items = QListWidget(self)
        self.list_items.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # exit = QAction('Exit', self)
        # exit.setShortcut('Ctrl+Q')
        # exit.setStatusTip('Exit Exit')
        # exit.triggered.connect(self.close)
        # settings = QAction('&Settings', self)
        # settings.setStatusTip('Settings and settings and settings')
        #
        # self.statusBar()
        # menubar = self.menuBar()
        # file = menubar.addMenu('&Date')
        # file.addAction(exit)
        # options = menubar.addMenu('&Option')
        # options.addAction(settings)
        #
        # self.mainWidget = QWidget(self)
        # self.setCentralWidget(self.mainWidget)
        #
        # select_path_label = QLabel("File Name:")
        # select_path_edit = QLineEdit()
        #
        # select_path = QPushButton("Save", self)
        # select_path.clicked.connect(self.save_event)
        #
        # cancel_button = QPushButton("Cancel", self)
        # cancel_button.clicked.connect(self.close)
        #
        # self.dirmodel = QFileSystemModel()
        # # Don't show files, just folders
        # self.dirmodel.setFilter(QtCore.QDir.NoDotAndDotDot | QtCore.QDir.AllDirs)
        #
        # self.folder_view = QTreeView(parent=self)
        # self.folder_view.setModel(self.dirmodel)
        # self.folder_view.clicked[QtCore.QModelIndex].connect(self.clicked)
        # # Don't show columns for size, file type, and last modified
        # self.folder_view.setHeaderHidden(True)
        # self.folder_view.hideColumn(1)
        # self.folder_view.hideColumn(2)
        # self.folder_view.hideColumn(3)
        #
        # self.selectionModel = self.folder_view.selectionModel()
        #
        # self.filemodel = QFileSystemModel()
        # self.filemodel.setFilter(QtCore.QDir.NoDotAndDotDot | QtCore.QDir.Files)  # Don't show folders, just files
        #
        # self.file_view = QListView(parent=self)
        # self.file_view.clicked[QtCore.QModelIndex].connect(self.update_selected_file)
        # self.file_view.setModel(self.filemodel)
        #
        # self.fileBrowserWidget = QWidget(self)
        # splitter_filebrowser = QSplitter()  # Horizontal top box
        # splitter_filebrowser.addWidget(self.folder_view)
        # splitter_filebrowser.addWidget(self.file_view)
        # splitter_filebrowser.setStretchFactor(0, 2)
        # splitter_filebrowser.setStretchFactor(1, 4)
        # hbox = QHBoxLayout()
        # hbox.addWidget(splitter_filebrowser)
        # self.fileBrowserWidget.setLayout(hbox)
        #
        # self.optionsWidget = QWidget(self)
        # vbox_options = QVBoxLayout(self.optionsWidget)
        # grid_input = QGridLayout()
        # grid_input.addWidget(select_path_label, 0, 0)
        # grid_input.addWidget(select_path_edit, 0, 1)
        # grid_input.addWidget(select_path, 0, 3)
        # grid_input.addWidget(cancel_button, 1, 3)
        # group_input = QGroupBox()
        # group_input.setLayout(grid_input)
        # vbox_options.addWidget(group_input)
        # self.optionsWidget.setLayout(vbox_options)
        #
        # splitter_filelist = QSplitter()
        # splitter_filelist.setOrientation(QtCore.Qt.Vertical)
        # splitter_filelist.addWidget(self.fileBrowserWidget)
        # splitter_filelist.addWidget(self.optionsWidget)
        #
        # vbox_main = QVBoxLayout(self.mainWidget)
        # vbox_main.addWidget(splitter_filelist)
        # vbox_main.setContentsMargins(0, 0, 0, 0)
        # self.setLayout(vbox_main)

        # self.dialog = QFileDialog()
        v_box = QVBoxLayout()
        v_box.addWidget(self.list_items)
        v_box.addWidget(save_button)
        self.setLayout(v_box)

    # def set_path(self):
    #     self.dirmodel.setRootPath("")
    #
    # def clicked(self):
    #     # get selected path of folder_view
    #     index = self.selectionModel.currentIndex()
    #     dir_path = self.dirmodel.filePath(index)
    #     self.filemodel.setRootPath(dir_path)
    #     self.file_view.setRootIndex(self.filemodel.index(dir_path))

    def set_item_list(self, items):
        item_names = list(items.keys())
        for item in item_names:
            self.list_items.addItem(str(item))
        self.list_items.setFixedHeight(
            self.list_items.sizeHintForRow(0) * self.list_items.count() + 2 * self.list_items.frameWidth())
        self.list_items.show()

    def save_event(self):
        try:
            process_menu = self.parent().children()[1].children()[1]
        except AttributeError:
            print("wrong parent")
        else:
            file_list = process_menu.get_save_info()

            items_to_save = []
            for item in self.list_items.selectedItems():
                default_file_name = file_list[str(item.text())]
                items_to_save.append(default_file_name)

            num_items = len(self.list_items.selectedIndexes())
            if num_items > 0:
                if num_items is 1:
                    file_name = items_to_save[0]
                    filename = self.save_file(file_name)
                else:
                    filename = self.save_multiple_files(items_to_save)
                self.status_message.emit("Saved '{}".format(filename))

    # def update_selected_file(self, index):
    #     self.current_file = index.data()
    #
    # def selected_save_data(self, aDict):
    #     for (k, v) in aDict.items():
    #         print(k)
    #         print(v)

    def get_save_dir(self):
        CSIDL_PERSONAL = 5  # My Documents
        SHGFP_TYPE_CURRENT = 0  # Get current, not default value
        buf = ctypes.create_unicode_buffer(ctypes.wintypes.MAX_PATH)
        ctypes.windll.shell32.SHGetFolderPathW(None, CSIDL_PERSONAL, None, SHGFP_TYPE_CURRENT, buf)
        return buf.value

    def save_file(self, file_name):
        dialog = QFileDialog()
        download_directory = self.get_save_dir()
        default_file = file_name.split('\\')[-1]
        default_directory = r'{0}\{1}'.format(download_directory, default_file)
        filename, extension = dialog.getSaveFileName(self,
                                                     caption='Select save location',  # directory=items_to_save[0],
                                                     directory=default_directory,
                                                     filter="Text files (*.txt);;Excel (*.xlsx *.xls);;All Files (*.*)",
                                                     initialFilter="Excel (*.xlsx *.xls)")
        if not filename:
            return

        self.file_saver(file_name, filename)
        return filename

    def save_multiple_files(self, items_to_save):
        dialog = QFileDialog()
        filename = dialog.getExistingDirectory(self,
                                               caption='Select folder',
                                               directory=self.get_save_dir())
        if not filename:
            return
        for file in items_to_save:
            self.file_saver(file, filename)
        return filename

    def file_saver(self, original_file, destination):
        # TODO: implement implicit file conversion based on filter type
        unsaved_file = pe.get_sheet(file_name=original_file)
        try:
            unsaved_file.save_as(filename=destination)
        except OSError:
            unsaved_file.save_as(filename=destination[0])


class ExcelDockWidget(DockWidget):
    tab_container_data = pyqtSignal(list, name='tab_container_data')
    tab_data = pyqtSignal(defaultdict, name='tab_data')

    def __init__(self, parent, window_name=None,
                 widget1=None, widget2=None,
                 button=None, hide_button_name=None):
        super(ExcelDockWidget, self).__init__(parent, window_name, widget1,
                                              widget2, button, hide_button_name)
        widget1.emit_dict.connect(lambda stuff: self.tab_container_data.emit(stuff))
        self.tab_data.connect(widget1.init_tabs)

    def make_tabs(self, args):
        # ACCUM_VIEWER_FILE
        self.tab_data.emit(args)


class ExcelTabContainer(QWidget):
    emit_dict = pyqtSignal(list)
    graph_xaxis = pyqtSignal(list)

    def __init__(self, title='Text Window', parent=None):
        super(ExcelTabContainer, self).__init__(parent)
        self.setWindowTitle(title)
        self.tabs = QTabWidget(self)
        v_box_layout = QVBoxLayout()
        v_box_layout.addWidget(self.tabs, alignment=QtCore.Qt.AlignCenter)
        h_box = QHBoxLayout()
        h_box.addLayout(v_box_layout)
        self.setLayout(h_box)
        self.tabs.currentChanged.connect(self.tab_change_event)
        self.hide()

    def init_tabs(self, args):
        excel_pages = sorted(args.keys())
        for page in excel_pages:
            try:
                spreadsheet = args[page][0]
                sheet_title = page
                self.append_spreadsheet(spreadsheet, sheet_title)
            except FileNotFoundError:
                pass
        self.show()

    def tab_change_event(self):
        index = self.tabs.currentIndex()
        handle = self.tabs.widget(index).children()
        try:
            use_handle = handle[1]
            use_handle.selected_data.connect(self.get_page_data)
        except Exception as e:
            print("passed due to {}".format(e))

    def get_page_data(self, stuff):
        try:
            indexices = self.tabs.count()
            for thing in stuff:
                # thing.__str__()
                client = thing.get_name()
                for list_position, curve in enumerate(thing.curves):
                    for index in range(indexices):
                        value = self.tabs.widget(index).children()[1].return_cell_value(client, curve)
                        thing.data[curve].append('%s-%s' % (index, value))
                        # thing.__str__()
            self.emit_dict.emit(stuff)
        except Exception as e:
            print("passed due to {}".format(e))

    def append_spreadsheet(self, spreadsheet, sheet_title):
        tab = QWidget(self)
        vBoxlayout = QVBoxLayout()
        excel_window = ExcelPageWidget(spreadsheet, sheet_title)
        vBoxlayout.addWidget(excel_window, alignment=QtCore.Qt.AlignCenter)
        h_box = QHBoxLayout()
        h_box.addLayout(vBoxlayout)
        tab.setLayout(h_box)
        self.tabs.addTab(tab, sheet_title)


class TableWidget(QTableWidget):
    def __init__(self,
                 parent=None,
                 window_title='Default Title',
                 file=None,
                 page=None):
        super(TableWidget, self).__init__(parent)
        self.sheet = None
        self.clip = QApplication.clipboard()
        self.row_dict = {}
        self.column_dict = {}
        self.open_sheet(file, page)
        self.setColumnCount(self.sheet.number_of_columns() - 1)
        self.setRowCount(self.sheet.number_of_rows() - 1)
        self.data = self.sheet.to_array()
        self.setWindowTitle(window_title)
        self.set_headers()
        self.set_my_data()
        self.resizeColumnsToContents()
        self.show()

    def open_sheet(self, file, page):
        try:
            if page:
                book = pe.get_book(file_name=file)
                sheet = book[page]
            else:
                sheet = pe.get_sheet(file_name=file)
            # try:
            #     sheet = pe.get_sheet(file_name=file)
            #     print("got sheet")
            # except AttributeError:
            #     book = pe.get_book(file_name=file)
            #     sheet = book[page]
            #     print("got sheet from book")
        except Exception as e:
            print('Error opening sheets in -> TableWidget')
            print('error = {0}'.format(e))
            pass
        else:
            self.sheet = sheet

    def set_my_data(self):
        for row_index, row in enumerate(self.data):
            for column_index, item in enumerate(row):
                new_item = QTableWidgetItem(str(item))
                self.setItem(row_index, column_index, new_item)

    def set_headers(self):
        self.remove_redundant_date()
        try:
            self.setHorizontalHeaderLabels(self.data[0])
        except TypeError:
            pass
        for index, header in enumerate(self.data[0]):
            self.column_dict[header] = index
        self.data.remove(self.data[0])
        vertical_headers = []
        for index, row in enumerate(self.data):
            vertical_headers.append(row[0])
            self.row_dict[row[0]] = index
            row.remove(row[0])
        try:
            self.setVerticalHeaderLabels(vertical_headers)
        except TypeError:
            pass

    def remove_redundant_date(self):
        try:
            fmt = "%A %m/%d/%Y"
            date_to_remove = self.data[0][0]
            if not not datetime.datetime.strptime(date_to_remove, fmt):
                self.data[0].remove(date_to_remove)
        except (ValueError, TypeError):
            pass

    def sizeHint(self):
        """Reimplemented to define a better size hint for the width of the
        TableEditor."""
        print("*i'm using the sizeHint*")
        x = self.style().pixelMetric(QStyle.PM_ScrollBarExtent,
                                     QStyleOptionHeader(), self)
        for column in range(self.columnCount() + 1):
            x += self.sizeHintForColumn(column) * 2.5
        y = x * .8
        size_hint = QSize(x, y)
        return size_hint

    def keyPressEvent(self, e):
        if e.modifiers() & QtCore.Qt.ControlModifier:
            selected = self.selectedRanges()
            if e.key() == QtCore.Qt.Key_C:  # copy
                # Set column headers
                s = '\t' + "\t".join([str(self.horizontalHeaderItem(i).text()) for i in
                                      range(selected[0].leftColumn(), selected[0].rightColumn() + 1)])
                st = []
                st.append(([str(self.horizontalHeaderItem(i).text()) for i in
                            range(selected[0].leftColumn(), selected[0].rightColumn() + 1)]))
                print(s)
                print('st: %s' % st)
                print("left header")
                s = '{0}\n'.format(s)

                for r in range(selected[0].topRow(), selected[0].bottomRow() + 1):
                    # Set row headers
                    s += '{0}\t'.format(str(self.verticalHeaderItem(r).text()))
                    for c in range(selected[0].leftColumn(), selected[0].rightColumn() + 1):
                        try:
                            # Copy cell values
                            s += "{0}\t".format(str(self.item(r, c).text()))
                        except AttributeError:
                            s += "\t"
                    s = s[:-1] + "\n"  # eliminate last '\t'
                self.clip.setText(s)

            if e.key() == QtCore.Qt.Key_X:  # copy w/o labels
                s = ''
                for r in range(selected[0].topRow(), selected[0].bottomRow() + 1):
                    for c in range(selected[0].leftColumn(), selected[0].rightColumn() + 1):
                        try:
                            # Copy cell values
                            s += "{0}\t".format(str(self.item(r, c).text()))
                        except AttributeError:
                            s += "\t"
                    s = s[:-1] + "\n"  # eliminate last '\t'
                self.clip.setText(s)

                # if e.modifiers() & QtCore.Qt.ShiftModifier:
                #     selected = self.selectionModel().selectedIndexes()
                #     if e.key() == QtCore.Qt.Key_Right:
                #         # TODO: this still does not work
                #         print("working")
                #         print(selected.selectedRows())


class ExcelPageWidget(TableWidget):
    selected_data = pyqtSignal(list, name='selected_data')

    def __init__(self, spreadsheet=None, popup_title='Test', parent=None, tab_sheet=None):
        super(ExcelPageWidget, self).__init__(parent=parent,
                                              window_title=popup_title,
                                              file=spreadsheet,
                                              page=tab_sheet)
        # self.book = pe.get_sheet(
        #     file_name=r'C:\Users\mscales\Desktop\Development\Daily SLA Parser - Automated Version\bin\CONFIG.xlsx')
        # # self.sheet = self.book['CONSTANTS']
        #
        # # self.title = popup_title
        # # self.setWindowTitle(popup_title)

    def mouseReleaseEvent(self, event):
        QTableWidget.mouseReleaseEvent(self, event)
        if event.button() == QtCore.Qt.RightButton:  # Release event only if done with left button, you can remove if necessary
            selected = self.selectedRanges()

            headers = list(([str(self.horizontalHeaderItem(i).text()) for i in
                             range(selected[0].leftColumn(), selected[0].rightColumn() + 1)]))
            # headers.count()
            # try:
            #     count_list = headers.count()
            #     print("converted to count")
            # except:
            #     pass
            # else:
            #     try:
            #         print(count_list)
            #         print("printed count list")
            #     except:
            #         pass
            #     else:
            #         try:
            #             print(dict(count_list))
            #             print("dictified")
            #         except:
            #             pass
            plots = []
            for r in range(selected[0].topRow(), selected[0].bottomRow() + 1):
                # Set row headers
                data = []
                client_name = self.verticalHeaderItem(r).text()
                data.append(client_name)
                for c in range(selected[0].leftColumn(), selected[0].rightColumn() + 1):
                    # Copy cell values
                    cell_value = self.item(r, c).text()
                    data.append(self.remove_cell_format(cell_value))
                new_plot = PlotData(headers)
                new_plot.make_data(data)
                plots.append(new_plot)
            self.selected_data.emit(plots)

    def return_cell_value(self, row, column):
        row = self.row_dict[row]
        column = self.column_dict[column]
        cell = self.data[row][column]
        return self.remove_cell_format(cell)

    def remove_cell_format(self, value_to_convert):
        try:
            value_to_convert = value_to_convert.split('%')[0]
        except AttributeError:
            value_to_return = int(value_to_convert)
        else:
            try:
                value_to_return = int(float(value_to_convert))
            except ValueError:
                h, m, s = [int(float(i)) for i in value_to_convert.split(':')]
                value_to_return = (3600 * int(h)) + (60 * int(m)) + int(s)
        return value_to_return


class TextWindow(QWidget):
    status_message = pyqtSignal(str, name='text_status')

    def __init__(self, parent, title="Test Window"):
        super(TextWindow, self).__init__(parent)
        self.setWindowTitle(title)
        self.te = QTextEdit(self)
        h_box_layout = QHBoxLayout()
        h_box_layout.addWidget(self.te)
        v_box_layout = QVBoxLayout()
        v_box_layout.addLayout(h_box_layout)
        self.setLayout(v_box_layout)
        self.te.setMinimumSize(500, 300)
        self.show()

    def append(self, text_string):
        cursor = self.te.textCursor()
        cursor.movePosition(cursor.End)
        cursor.insertText(text_string)
        self.te.ensureCursorVisible()


class RunProcess(QObject):
    finished_signal = pyqtSignal(bool, name='finished signal')
    error_signal = pyqtSignal(bool, name='error_signal')

    def __init__(self, parent=None):
        super(RunProcess, self).__init__(parent)
        self.qprocess = QtCore.QProcess(self)
        self.qprocess.finished.connect(lambda:
                                       self.exit_process(self.qprocess.exitCode())
                                       )

    def do_work(self, *params):
        self.qprocess.start("python", *params)  # qprocess throws IOError here for viewer. override class method

    def __str__(self):
        print("I'm a process named: {}".format(self))

    def exit_process(self, process_code):
        if process_code is 1:
            self.error_signal.emit(True)
        else:
            self.finished_signal.emit(True)


class ButtonsDockWidget(DockWidget):
    # TODO: add args for programs
    # 1. override for sla program's check_report_completed
    def __init__(self, parent, window_name=None,
                 widget1=None, widget2=None,
                 button=None, hide_button_name=None):
        super(ButtonsDockWidget, self).__init__(parent, window_name, widget1,
                                                widget2, button, hide_button_name)

        # def make_buttons(self):
        #     list_items = QListWidget()
        #     for x in range(0, 5):
        #         list_items.addItem(str(x))
        #     list_items.show()
        #     self.setWidget(list_items)


class Window(QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.setGeometry(50, 50, 500, 300)
        self.setWindowTitle("PyQT tuts!")
        self.setWindowIcon(QIcon('pythonlogo.png'))

        extractAction = QAction("&GET TO THE CHOPPAH!!!", self)
        extractAction.setShortcut("Ctrl+Q")
        extractAction.setStatusTip('Leave The App')
        extractAction.triggered.connect(self.close_application)

        openEditor = QAction("&Editor", self)
        openEditor.setShortcut("Ctrl+E")
        openEditor.setStatusTip('Open Editor')
        openEditor.triggered.connect(self.editor)

        openFile = QAction("&Open File", self)
        openFile.setShortcut("Ctrl+O")
        openFile.setStatusTip('Open File')
        openFile.triggered.connect(self.file_open)

        saveFile = QAction("&Save File", self)
        saveFile.setShortcut("Ctrl+S")
        saveFile.setStatusTip('Save File')
        saveFile.triggered.connect(self.file_save)

        self.statusBar()

        mainMenu = self.menuBar()

        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(extractAction)
        fileMenu.addAction(openFile)
        fileMenu.addAction(saveFile)

        editorMenu = mainMenu.addMenu("&Editor")
        editorMenu.addAction(openEditor)

        self.home()

    def home(self):
        btn = QPushButton("Quit", self)
        btn.clicked.connect(self.close_application)
        btn.resize(btn.minimumSizeHint())
        btn.move(0, 100)

        extractAction = QAction(QIcon('todachoppa.png'), 'Flee the Scene', self)
        extractAction.triggered.connect(self.close_application)
        self.toolBar = self.addToolBar("Extraction")
        self.toolBar.addAction(extractAction)

        fontChoice = QAction('Font', self)
        fontChoice.triggered.connect(self.font_choice)
        # self.toolBar = self.addToolBar("Font")
        self.toolBar.addAction(fontChoice)

        color = QColor(0, 0, 0)

        fontColor = QAction('Font bg Color', self)
        fontColor.triggered.connect(self.color_picker)

        self.toolBar.addAction(fontColor)

        checkBox = QCheckBox('Enlarge Window', self)
        checkBox.move(300, 25)
        checkBox.stateChanged.connect(self.enlarge_window)

        self.progress = QProgressBar(self)
        self.progress.setGeometry(200, 80, 250, 20)

        self.btn = QPushButton("Download", self)
        self.btn.move(200, 120)
        self.btn.clicked.connect(self.download)

        # print(self.style().objectName())
        self.styleChoice = QLabel("Windows Vista", self)

        comboBox = QComboBox(self)
        comboBox.addItem("motif")
        comboBox.addItem("Windows")
        comboBox.addItem("cde")
        comboBox.addItem("Plastique")
        comboBox.addItem("Cleanlooks")
        comboBox.addItem("windowsvista")

        comboBox.move(50, 250)
        self.styleChoice.move(50, 150)
        comboBox.activated[str].connect(self.style_choice)

        cal = QCalendarWidget(self)
        cal.move(500, 200)
        cal.resize(200, 200)

        self.show()

    def file_open(self):
        name = QFileDialog.getOpenFileName(self, 'Open File')
        file = open(name, 'r')

        self.editor()

        with file:
            text = file.read()
            self.textEdit.setText(text)

    def file_save(self):
        name = QFileDialog.getSaveFileName(self, 'Save File')
        file = open(name, 'w')
        text = self.textEdit.toPlainText()
        file.write(text)
        file.close()

    def color_picker(self):
        color = QColorDialog.getColor()
        self.styleChoice.setStyleSheet("QWidget { background-color: %s}" % color.name())

    def editor(self):
        self.textEdit = QTextEdit()
        self.setCentralWidget(self.textEdit)

    def font_choice(self):
        font, valid = QFontDialog.getFont()
        if valid:
            self.styleChoice.setFont(font)

    def style_choice(self, text):
        self.styleChoice.setText(text)
        QApplication.setStyle(QStyleFactory.create(text))

    def download(self):
        self.completed = 0

        while self.completed < 100:
            self.completed += 0.0001
            self.progress.setValue(self.completed)

    def enlarge_window(self, state):
        if state == QtCore.Qt.Checked:
            self.setGeometry(50, 50, 1000, 600)
        else:
            self.setGeometry(50, 50, 500, 300)

    def close_application(self):
        choice = QMessageBox.question(self, 'Extract!',
                                      "Get into the chopper?",
                                      QMessageBox.Yes | QMessageBox.No)
        if choice == QMessageBox.Yes:
            print("Extracting Naaaaaaoooww!!!!")
            sys.exit()
        else:
            pass


class MyApplication(QApplication):
    def __init__(self, argv):
        super(MyApplication, self).__init__(argv)
        self.clip = QApplication.clipboard()


def main():
    app = MyApplication(sys.argv)
    ex = MainFrame()
    ex.application_color.connect(app.setStyleSheet)
    ex.application_font.connect(app.setFont)
    ex.application_style.connect(lambda text: app.setStyle(QStyleFactory.create(text)))
    ex.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    from os import sys, path

    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
    main()

    # if __name__ == '__main__':
    #     import sys
    #
    #     if sys.flags.interactive != 1 or not hasattr(QtCore, 'PYQT_VERSION'):
    #         pg.QtGui.QApplication.exec_()
