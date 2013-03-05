import sys
import os
import xlwt, csv
from PySide.QtCore import *
from PySide.QtGui import *

csv_dir = ""
xls_dir = ""

class Form(QDialog):

	def __init__(self, parent=None):
		super(Form, self).__init__(parent)
		self.setWindowTitle("CSV to Excel")
		self.resize(450, 50)
		self.edit1 = QLineEdit("Select Input Folder..")
		self.edit1.setReadOnly(True)
		self.button1 = QPushButton("Input folder..")
		self.edit2 = QLineEdit("Select Output Folder..")
		self.edit2.setReadOnly(True)
		self.button2 = QPushButton("Output folder..")
		self.buttongo = QPushButton("Convert")

		self.progress = QProgressBar()
		self.progress.setMinimum(0)
		self.progress.setMaximum(100)
		self.progress.setValue(0)

		# Create layout and add widgets
		layout = QGridLayout()
		layout.addWidget(self.edit1, 1, 1)
		layout.addWidget(self.button1, 1, 2)
		layout.addWidget(self.edit2, 2, 1)
		layout.addWidget(self.button2, 2, 2)
		layout.addWidget(self.progress, 3, 1)
		layout.addWidget(self.buttongo, 3, 2)

		# Set dialog layout
		self.setLayout(layout)

		# Add button signal
		self.button1.clicked.connect(self.fInput)
		self.button2.clicked.connect(self.fOutput)
		self.buttongo.clicked.connect(self.convert)


	def greetings(self):
		print ("Hello %s" % self.edit.text())
	
	def fInput(self):
		#fileName = QFileDialog.getOpenFileName(self, "Open Image", "c:\Users\jgasimov\Pictures", "Image Files (*.png *.jpg *.bmp)")
		fileName = QFileDialog.getExistingDirectory(self, "Select Folder", os.getcwd(), options=QFileDialog.ShowDirsOnly)
		self.edit1.setText(str(fileName))
		global csv_dir
		csv_dir = str(fileName) 
	
	def fOutput(self):
		#fileName = QFileDialog.getOpenFileName(self, "Open Image", "c:\Users\jgasimov\Pictures", "Image Files (*.png *.jpg *.bmp)")
		fileName = QFileDialog.getExistingDirectory(self, "Select Folder", os.getcwd(), options=QFileDialog.ShowDirsOnly)
		self.edit2.setText(str(fileName)) 
		global xls_dir
		xls_dir = str(fileName)
	
	def messageOK(self, file_count):
		msgBox = QMessageBox()
		msgBox.setWindowTitle("Success")
		msgBox.setIcon(QMessageBox.Information)
		msgBox.setText("Finished!")
		msgBox.setInformativeText("Converted " + str(file_count) + " files.")
		msgBox.setStandardButtons(QMessageBox.Close)
		msgBox.exec_()

	def errWindow(self, errmsg):
		msgBox = QMessageBox()
		msgBox.setWindowTitle("Error")
		msgBox.setIcon(QMessageBox.Warning)
		msgBox.setText(errmsg)
		msgBox.setStandardButtons(QMessageBox.Ok)
		msgBox.exec_()	

	
	def convert(self):
		i = 0.0
		if csv_dir == "":
			self.errWindow("Please select Input Folder.")
		elif xls_dir == "":
			self.errWindow("Please select Output Folder.")
		else:
			os.chdir(csv_dir)
			file_count = len([name for name in os.listdir('.') if name.endswith(".csv")])
			if file_count == 0:
				self.errWindow("No CSV files found.")
			else:
				for files in os.listdir("."):
					if files.endswith(".csv"):
						wb = xlwt.Workbook(encoding='latin-1')
						ws = wb.add_sheet('Sheet1')
						sourceCSV = csv.reader(open(files, 'rb'), delimiter=";")
						for rowi, row in enumerate(sourceCSV):
							for coli, value in enumerate(row):
							 	ws.write(rowi, coli, value)
						xls_file = xls_dir + "\\" + files[:-4] + '.xls'
						wb.save(xls_file)
						i +=1
						self.progress.setValue(int(round(i/(file_count/100.0))))
				self.messageOK(file_count)

if __name__ == '__main__':
	# Create the Qt Application
	app = QApplication(sys.argv)
	app.setWindowIcon(QIcon("icon.ico"))
	# Create and show the form
	form = Form()
	form.show()
	# Run the main Qt loop
	sys.exit(app.exec_())