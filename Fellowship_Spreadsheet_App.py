import PyQt5
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QGridLayout, QWidgetItem
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QFileDialog, QPlainTextEdit, QTabWidget
from styles import dark_fusion, default
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPalette, QColor, QPainter, QPen
from pandas import read_excel

class Wrapper(QWidget):
    def __init__(self):
        super().__init__()
        self.setMinimumSize(1005, 460)

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "","All Files (*);;CSV Files (*.csv);;Excel Files (*.xlsx)", options=options)
        if fileName:
            message_box.messageUpdate("Attempting to open file {}........".format(fileName))
            if fileName[fileName.index('.'):] == ".csv":
                csv_file = open(fileName, "r")
                for i in csv_file:
                    i = i.split(',')
                    if type(i[0]) == 'str':
                        continue
                    else:
                        if cascade_sheets.returnCurrent() == 0:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(0)))
                            spreadsheet.appendData(i)
                        elif cascade_sheets.returnCurrent() == 1:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(1)))
                            spreadsheet2.appendData(i)
                        elif cascade_sheets.returnCurrent() == 2:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(2)))
                            spreadsheet3.appendData(i)
                        else:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(3)))
                            spreadsheet4.appendData(i)
            elif fileName[fileName.index('.'):] == ".xlsx":
                xlsx_file = read_excel(fileName)
                for index, row in xlsx_file.iterrows():
                    if type(list(row)[0]) == 'str':
                        continue
                    else:
                        if cascade_sheets.returnCurrent() == 0:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(0)))
                            spreadsheet.appendData(list(row))
                        elif cascade_sheets.returnCurrent() == 1:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(1)))
                            spreadsheet2.appendData(list(row))
                        elif cascade_sheets.returnCurrent() == 2:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(2)))
                            spreadsheet3.appendData(list(row))
                        else:
                            message_box.messageUpdate("Placing Data into {}......".format(cascade_sheets.tabText(3)))
                            spreadsheet4.appendData(list(row))
        if cascade_sheets.returnCurrent() == 0:
           spreadsheet.insertRow(spreadsheet.numrows)
        elif cascade_sheets.returnCurrent() == 1:
           spreadsheet2.insertRow(spreadsheet2.numrows)
        elif cascade_sheets.returnCurrent() == 2:
           spreadsheet3.insertRow(spreadsheet3.numrows)
        elif cascade_sheets.returnCurrent() == 3:
           spreadsheet4.insertRow(spreadsheet4.numrows)
        message_box.messageUpdate("Done!\n")

    def validate_data(self):
        id_col_data = list()
        rowcount = 0
        colcount = 0
        sp_index = cascade_sheets.returnCurrent()
        tab_header = cascade_sheets.tabText(sp_index)
        message_box.messageUpdate("Validating {}......".format(tab_header))
        if sp_index == 0:
           rowcount = spreadsheet.numrows
           colcount = len(spreadsheet.headerLabels)
        elif sp_index == 1:
           rowcount = spreadsheet2.numrows
           colcount = len(spreadsheet2.headerLabels)
        elif sp_index == 2:
           rowcount = spreadsheet3.numrows
           colcount = len(spreadsheet3.headerLabels)
        elif sp_index == 3:
           rowcount = spreadsheet4.numrows
           colcount = len(spreadsheet4.headerLabels)
        for j in range(rowcount):
            for i in range(colcount):
                if i == 0:
                    try:
                        if sp_index == 0:
                            id_col_data.append(str(int(spreadsheet.returnCellData(j, i))))
                        elif sp_index == 1:
                            id_col_data.append(str(int(spreadsheet2.returnCellData(j, i))))
                        elif sp_index == 2:
                            id_col_data.append(str(int(spreadsheet3.returnCellData(j, i))))
                        elif sp_index == 3:
                            id_col_data.append(str(int(spreadsheet4.returnCellData(j, i))))
                    except ValueError:
                        message_box.messageUpdate("Error in {}: (ValueError) Invalid Value at ({},{})".format(tab_header,j+1,i+1))
                        continue
                else:
                    try:
                        if sp_index == 0:
                            str(int(spreadsheet.returnCellData(j, i)))
                        elif sp_index == 1:
                            str(int(spreadsheet2.returnCellData(j, i)))
                        elif sp_index == 2:
                            str(int(spreadsheet3.returnCellData(j, i)))
                        elif sp_index == 3:
                            str(int(spreadsheet4.returnCellData(j, i)))
                    except ValueError:
                        message_box.messageUpdate("Error in {}: (ValueError) Invalid Value at ({},{})".format(tab_header,j+1,i+1))
                        continue
        if len(id_col_data) != len(set(id_col_data)):
            message_box.messageUpdate("Error in {}: (PKeyError) ID field is primary key!! Cannot have invalid and/or repeat values!".format(tab_header))
        message_box.messageUpdate("Done!\n")

    def save_sheet(self):
        rowcount = 0
        colcount = 0
        sp_index = cascade_sheets.returnCurrent()
        tab_header = cascade_sheets.tabText(sp_index)
        return_text = dict()
        if sp_index == 0:
           rowcount = spreadsheet.numrows
           colcount = len(spreadsheet.headerLabels)
        elif sp_index == 1:
           rowcount = spreadsheet2.numrows
           colcount = len(spreadsheet2.headerLabels)
        elif sp_index == 2:
           rowcount = spreadsheet3.numrows
           colcount = len(spreadsheet3.headerLabels)
        elif sp_index == 3:
           rowcount = spreadsheet4.numrows
           colcount = len(spreadsheet4.headerLabels)
        for j in range(rowcount):
            for i in range(colcount):
                if sp_index == 0:
                    return_text[str(spreadsheet.horizontalHeaderItem(i).text())] = str(int(spreadsheet.returnCellData(j, i)))
                elif sp_index == 1:
                    return_text[str(spreadsheet1.horizontalHeaderItem(i).text())] = str(int(spreadsheet2.returnCellData(j, i)))
                elif sp_index == 2:
                    return_text[str(spreadsheet2.horizontalHeaderItem(i).text())] = str(int(spreadsheet3.returnCellData(j, i)))
                elif sp_index == 3:
                    return_text[str(spreadsheet3.horizontalHeaderItem(i).text())] = str(int(spreadsheet4.returnCellData(j, i)))
            filename = "{}_{}.txt".format(tab_header,j+1)
            message_box.messageUpdate("Saving {}........".format(filename))
            file_write = open(filename, "w")
            file_write.write(str(return_text))
            file_write.close()
        message_box.messageUpdate("Done!\n")

class Spreadsheet(QTableWidget):
    def __init__(self, headerLabels):
        super().__init__()
        self.numlimit = 5
        self.setRowCount(self.numlimit)
        self.setColumnCount(len(headerLabels))
        self.numrows = 0
        self.headerLabels = headerLabels
        self.setHorizontalHeaderLabels(self.headerLabels)

    def appendData(self, data):
        if self.numrows >= 5:
            self.insertRow(self.numrows)
        for i in range(len(data)):
            newItem = QTableWidgetItem(str(data[i]))
            self.setItem(self.numrows,i,newItem)
        self.numrows += 1

    def returnCellData(self, row, col):
        return str(self.item(row, col).text())

class spreadsheet_tab_widget(QTabWidget):
    def __init__(self):
        super().__init__()
    def update_table_data(self):
        if self.currentIndex() == 0:
            spreadsheet.appendData([1,1,1,1,1,1,1])
        elif self.currentIndex() == 1:
            spreadsheet2.appendData([1,1,1,1,1])
        elif self.currentIndex() == 2:
            spreadsheet3.appendData([1,1,1,1,1,1,1,1])
        elif self.currentIndex() == 3:
            spreadsheet4.appendData([1,1,1,1,1,1,1])
    def returnCurrent(self):
        return self.currentIndex()

class MessageBox(QPlainTextEdit):
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)

    def messageUpdate(self, text):
        self.appendPlainText(text)

app = QApplication([])
dark_fusion(app)

window = Wrapper()

spreadsheet = Spreadsheet(['ID', 
                           'Connection Type', 
                           'Axial Load', 
                           'Shear Load', 
                           'Bolt Diameter', 
                           'Bolt Grade', 
                           'Plate Thickness'])
spreadsheet2 = Spreadsheet(['ID',
                            'Member Length',
                            'Tensile Load', 
                            'Support Condition at End 1',
                            'Support Condition at End 2'])
spreadsheet3 = Spreadsheet(['ID',
                            'End Plate Type',
                            'Shear Load',
                            'Axial Load',
                            'Moment Load',
                            'Bolt Diameter',
                            'Bolt Grade',
                            'Plate Thickness'])
spreadsheet4 = Spreadsheet(['ID',
                            'Angle Leg1',
                            'Angle Leg2',
                            'Angle Thickness',
                            'Shear Load',
                            'Bolt Diameter',
                            'Bolt Grade'])

spreadsheet_widget = Wrapper()
spreadsheet2_widget = Wrapper()
spreadsheet3_widget = Wrapper()
spreadsheet4_widget = Wrapper()

sp_layout1 = QVBoxLayout()
sp_layout2 = QVBoxLayout()
sp_layout3 = QVBoxLayout()
sp_layout4 = QVBoxLayout()

sp_layout1.addWidget(spreadsheet)
sp_layout2.addWidget(spreadsheet2)
sp_layout3.addWidget(spreadsheet3)
sp_layout4.addWidget(spreadsheet4)

spreadsheet_widget.setLayout(sp_layout1)
spreadsheet2_widget.setLayout(sp_layout2)
spreadsheet3_widget.setLayout(sp_layout3)
spreadsheet4_widget.setLayout(sp_layout4)

message_box = MessageBox()
cascade_sheets = spreadsheet_tab_widget()

major_layout = QGridLayout()
minor_interactive = QVBoxLayout()

load_inputs = QPushButton("Load Inputs")
validate = QPushButton("Validate")
submit = QPushButton("Submit")

cascade_sheets.addTab(spreadsheet_widget, "FinPlate")
cascade_sheets.addTab(spreadsheet2_widget, "TensionMember")
cascade_sheets.addTab(spreadsheet3_widget, "BCEndPlate")
cascade_sheets.addTab(spreadsheet4_widget, "CleatAngle")

load_inputs.clicked.connect(window.openFileNameDialog)
submit.clicked.connect(window.save_sheet)
validate.clicked.connect(window.validate_data)

minor_interactive.addWidget(load_inputs)
minor_interactive.addWidget(validate)
minor_interactive.addWidget(submit)

major_layout.addLayout(minor_interactive, 1, 1)
major_layout.addWidget(cascade_sheets, 1, 2)
major_layout.addWidget(message_box, 2, 2)

window.setLayout(major_layout)
window.show()
app.exec_()

