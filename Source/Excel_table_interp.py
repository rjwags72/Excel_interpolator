import sys
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QLineEdit, QMessageBox
from PyQt5.QtWidgets import QCheckBox, QLabel, QComboBox, QProgressBar
from Interpolation import int_table
from find_sheets import find_sheets
from PyQt5.QtCore import pyqtSlot, Qt

#from pyqtspinner.spinner import WaitingSpinner as spin
#import os.path

#globals ex,

#our widget
class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Interpolate a table from excel'
        self.left = 10
        self.top = 10
        self.width = 700
        self.height = 2000
        self.initUI()
    
    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        
        #create textinput for input filepath box & title
        self.file_path_label = QLabel(self)
        self.file_path_label.setText('Input Filepath')
        #self.file_path_label.resize(400,200)
        self.file_path_label.move(10,15)
        self.file_path = QLineEdit('Input Filepath', self)
        self.file_path.setDragEnabled(True)
        self.file_path.move(10, 40)
        self.file_path.resize(500,32)
        self.file_path.status = True
        self.file_path.setDisabled(True)
        self.file_path.textChanged.connect(self.set_out_to_in)
        
        #create textinput for input Sheet_name box & title
        self.sheet_name_label = QLabel(self)
        self.sheet_name_label.setText('Choose Sheet Name')
        #self.file_path_label.resize(400,200)
        self.sheet_name_label.move(10,600)
        self.sheet_name = QComboBox(self)
        self.sheet_name.move(10, 620)
        self.sheet_name.resize(500,32)
        self.sheet_name.activated[str].connect(self.set_out_to_in_sheet)
        
        #create check box to decide wether to drop or type
        self.input_type = QCheckBox("I would like to Drop in file",self)
        self.input_type.setChecked(True)
        self.input_type.stateChanged.connect(self.input_change)
        self.input_type.move(10, 75)
        
        #create check box to to output file to same place
        self.out_to_in = QCheckBox("Same as Input",self)
        self.out_to_in.setChecked(True)
        self.out_to_in.stateChanged.connect(self.same_out)
        self.out_to_in.move(400, 680)
        
        #create textinput for output filepath box & title
        self.out_file_path_label = QLabel(self)
        self.out_file_path_label.setText('Output Filepath')
        self.out_file_path_label.move(10,685)
        self.out_file_path = QLineEdit(self.file_path.text(), self)
        self.out_file_path.setDragEnabled(True)
        self.out_file_path.move(10, 705)
        self.out_file_path.resize(500,32)
        self.out_file_path.status = True
        self.out_file_path.setDisabled(True)
        
        #create textinput for output sheetname box & title ########
        self.out_sheet_name_label = QLabel(self)
        self.out_sheet_name_label.setText('Output Sheet Name')
        self.out_sheet_name_label.move(10,740)
        #self.out_sheet_name = QLineEdit(self.sheet_name.text(), self)
        self.out_sheet_name = QLineEdit( self)
        self.out_sheet_name.setDragEnabled(True)
        self.out_sheet_name.move(10, 765)
        self.out_sheet_name.resize(500,32)
        self.out_sheet_name.status = True
        self.out_sheet_name.setDisabled(True)
        
        #create check box  to output sheet to same sheetname
        self.out_sheet = QCheckBox("Same as Input",self)
        self.out_sheet.setChecked(True)
        self.out_sheet.stateChanged.connect(self.same_out_sheet)
        self.out_sheet.move(400, 740)
        
        #create step size input
        self.step_label = QLabel(self)
        self.step_label.setText('Step Size:')
        self.step_label.move(390,815)
        self.step_label.setStyleSheet("font: 23")
        self.step = QLineEdit('0.1', self)
        self.step.setDragEnabled(True)
        self.step.move(460, 810)
        self.step.resize(50,32)
        
        #create column to interpolate from input
        self.col_label = QLabel(self)
        self.col_label.setText('Column:')
        self.col_label.move(200,815)
        self.col_label.setStyleSheet("font: 23")
        self.col = QLineEdit(None, self)
        self.col.setDragEnabled(True)
        self.col.move(260, 810)
        self.col.resize(50,32)
        self.col.setDisabled(True)
        self.col.status = True
        
        #create check box to decide wether to drop or type
        self.col_box = QCheckBox("interpolate from first column with data",self)
        self.col_box.setChecked(True)
        self.col_box.stateChanged.connect(self.ret_none)
        self.col_box.move(250, 75)
        
        
        # create button
        self.updt_butt = QPushButton('Interpolate Table', self)
        self.updt_butt.move(10,810)
        self.updt_butt.clicked.connect(self.on_click)
        
        #initalize drop area
        self.drop_area = droparea('Drop here.', self)
        self.drop_area.resize(500, 500)
        self.drop_area.move(10,100)
        self.drop_area.status = False
        self.drop_area.setDisabled(False)
        
        #remove file from drop area
        
        
        self.show()
        
    @pyqtSlot()
    def on_click(self):

        file_path = self.file_path.text()
        sheet_name_read = str(self.sheet_name.currentText())
        steps = self.step.text()
        steps = float(steps)
        col_to_int = self.col.text()
        same_file = self.out_to_in.isChecked()
        same_sheet = self.out_sheet.isChecked()
        file_path_out = self.out_file_path.text()
        sheet_name_write = self.out_sheet_name.text()
        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(10, 850, 300, 25)
        self.pbar.show()
        self.saving_ind = QLabel(self)
        status = int_table(file_path, sheet_name_read, steps, col_to_int,
                             file_path_out, sheet_name_write, same_sheet, 
                             same_file, self.pbar, self.saving_ind)
        msg_complete = QMessageBox()
        msg_complete.setText(str(status))
        msg_complete.setStandardButtons(QMessageBox.Ok)
        msg_complete.show()
        msg_complete.exec_()
        self.saving_ind.hide()
        self.pbar.hide()
        
#int_table(fp,sn,steps,col_to_interp_from,fp2,sn2,same_sheet,same_file)
        
    #function to change the file input area
    def input_change(self):
        self.file_path.status = not(self.file_path.status)
        self.drop_area.status = not(self.drop_area.status)
        self.drop_area.setDisabled(self.drop_area.status)
        self.file_path.setDisabled(self.file_path.status)
    
    #function to set inout file to output file
    def same_out(self):
        self.out_file_path.status = not(self.out_file_path.status)
        self.out_file_path.setDisabled(self.out_file_path.status)
        if(self.out_file_path.status == True):
            self.out_file_path.setText(self.file_path.text())
        else:
            self.out_file_path.setText('Output Filepath')
    
    def set_out_to_in(self):
        if(self.out_to_in.isChecked() == True):
            self.out_file_path.setText(self.file_path.text())
    
    def set_out_to_in_sheet(self):
        if(self.out_sheet.isChecked() == True):
            self.out_sheet_name.setText(str(self.sheet_name.currentText()))
            
    def same_out_sheet(self):
        self.out_sheet_name.status = not(self.out_sheet_name.status)
        self.out_sheet_name.setDisabled(self.out_sheet_name.status)
        if(self.out_sheet_name.status == True):
            self.out_sheet_name.setText(str(self.sheet_name.currentText()))
        else:
            self.out_sheet_name.setText('Output Sheet Name')
            
    def ret_none(self):
        self.col.status = not(self.col.status)
        self.col.setDisabled(self.col.status)
        if(self.col.status == True):
            self.col.setText(None)
        
    #drop area class
class droparea(QLabel):
    
    def __init__(self, title, parent):
        super().__init__(title, parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("border: 1px solid black;") 
    def dragEnterEvent(self, e):
        e.accept()
        #global file_path_in
        #self.file_path.setText(e)
# =============================================================================
#         if e.mimeData().hasFormat(csvDATA):
#             e.accept()
#         else:
#             e.ignore()
# =============================================================================

    def dropEvent(self, e):
        self.setText(e.mimeData().text())
        global ex
        txt = str(e.mimeData().text())
        txt = txt.strip('file:///')
        ex.file_path.setText(txt)

        for i in range(0, ex.sheet_name.count()):
            ex.sheet_name.removeItem(i)
        ex.sheet_name.clear()
        
        sheets = find_sheets(txt)
        for i in sheets:
            ex.sheet_name.addItem(i)
        if (ex.out_sheet.isChecked() == True):
            ex.out_sheet_name.setText(str(ex.sheet_name.currentText()))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())

