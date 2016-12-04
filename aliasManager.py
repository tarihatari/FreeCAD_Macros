# -*- coding: utf-8 -*-
''' Script for setting, moving and clearing aliases in the Spreadsheet.
    It allows to generate Part Families

    hatari 2016 v0.2

    GNU Lesser General Public License (LGPL)
'''
from PySide import QtGui, QtCore
import os

class MyButtons(QtGui.QDialog):
	""""""
	def __init__(self):
		super(MyButtons, self).__init__()
		self.initUI()
	def initUI(self):
		option1Button = QtGui.QPushButton("Set Aliases")
		option1Button.clicked.connect(self.onOption1)
		option2Button = QtGui.QPushButton("Clear Aliases")
		option2Button.clicked.connect(self.onOption2)
		option3Button = QtGui.QPushButton("Move Aliases")
		option3Button.clicked.connect(self.onOption3)
  		option4Button = QtGui.QPushButton("Generate Part Family")
		option4Button.clicked.connect(self.onOption4)
		#
		buttonBox = QtGui.QDialogButtonBox()
		buttonBox = QtGui.QDialogButtonBox(QtCore.Qt.Vertical)
		buttonBox.addButton(option1Button, QtGui.QDialogButtonBox.ActionRole)
		buttonBox.addButton(option2Button, QtGui.QDialogButtonBox.ActionRole)
		buttonBox.addButton(option3Button, QtGui.QDialogButtonBox.ActionRole)
  		buttonBox.addButton(option4Button, QtGui.QDialogButtonBox.ActionRole)
		#
		mainLayout = QtGui.QVBoxLayout()
		mainLayout.addWidget(buttonBox)
		self.setLayout(mainLayout)
		# define window		xLoc,yLoc,xDim,yDim
		self.setGeometry(400, 400, 300, 50)
		self.setWindowTitle("Alias Manager")
		self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
	def onOption1(self):
		self.retStatus = 1

		self.close()
	def onOption2(self):
		self.retStatus = 2
		self.close()
	def onOption3(self):
		self.retStatus = 3
		self.close()
  	def onOption4(self):
		self.retStatus = 4
		self.close()
  
  
# Define Aliases
def routine1():
    column = QtGui.QInputDialog.getText(None, "Column containing Values", "Enter Column Letter")
    if column[1]:
        col = str.capitalize(str(column[0])) # Always use capital characters for Spreadsheet
        startCell =  QtGui.QInputDialog.getInteger(None, "start Row number", "Input start Row number:")
        if startCell[1]:
            endCell =  QtGui.QInputDialog.getInteger(None, "end Row number", "Input end Row number:")
            if endCell[1]:
                for i in range(startCell[0],endCell[0]+1):
                    cellFrom = 'A' + str(i)
                    cellTo = str(col[0]) + str(i)
                    App.ActiveDocument.Spreadsheet.setAlias(cellTo, '')
                    App.ActiveDocument.recompute()
                    App.ActiveDocument.Spreadsheet.setAlias(cellTo, App.ActiveDocument.Spreadsheet.getContents(cellFrom))
# Clear Aliases
def routine2():
        column = QtGui.QInputDialog.getText(None, "Column containing Values", "Enter Column Letter")
        if column[1]:
            col = str.capitalize(str(column[0]))
            startCell =  QtGui.QInputDialog.getInteger(None, "start Row number", "Input start Row number:")
            if startCell[1]:
                endCell =  QtGui.QInputDialog.getInteger(None, "end Row number", "Input end Row number:")
                if endCell[1]:
                    for i in range(startCell[0],endCell[0]+1):
                        cellTo = str(col[0]) + str(i)
                        App.ActiveDocument.Spreadsheet.setAlias(cellTo, '')
                        App.ActiveDocument.recompute()
# Move Aliases
def routine3():
    columnFrom = QtGui.QInputDialog.getText(None, "Value Column", "Move From")
    if columnFrom[1]:
        columnTo = QtGui.QInputDialog.getText(None, "Value Column", "Move To")
        if columnTo[1]:
            colF = str.capitalize(str(columnFrom[0]))
            colT = str.capitalize(str(columnTo[0]))
            startCell =  QtGui.QInputDialog.getInteger(None, "start Row number", "Input start Row number:")
            if startCell[1]:
                endCell =  QtGui.QInputDialog.getInteger(None, "end Row number", "Input end Row number:")
                if endCell[1]:
                    for i in range(startCell[0],endCell[0]+1):
                        cellDef = 'A'+ str(i)                        
                        cellFrom = str(colF[0]) + str(i)
                        cellTo = str(colT[0]) + str(i)
                        App.ActiveDocument.Spreadsheet.setAlias(cellFrom, '')
                        App.ActiveDocument.recompute()
                        App.ActiveDocument.Spreadsheet.setAlias(cellTo, App.ActiveDocument.Spreadsheet.getContents(cellDef))


# Generate Part Family
def routine4():
    # Get Filename
    doc = FreeCAD.ActiveDocument    
    if not doc.FileName:
        FreeCAD.Console.PrintError('Must save project first\n')
        
    docDir, docFilename = os.path.split(doc.FileName)
    filePrefix = os.path.splitext(docFilename)[0]

    def char_range(c1, c2):
        """Generates the characters from `c1` to `c2`, inclusive."""
        for c in xrange(ord(c1), ord(c2)+1):
            yield str.capitalize(chr(c))
    columnFrom = QtGui.QInputDialog.getText(None, "Column", "Range From")
    if columnFrom[1]:
        columnTo = QtGui.QInputDialog.getText(None, "Column", "Range To")
        if columnTo[1]:
            startCell =  QtGui.QInputDialog.getInteger(None, "Start Cell Row", "Input Start Cell Row:")
            if startCell[1]:
                endCell =  QtGui.QInputDialog.getInteger(None, "End Cell Row", "Input End Cell Row:")
                if endCell[1]:    
                    fam_range = []
                    for c in char_range(str(columnFrom[0]), str(columnTo[0])):
                        fam_range.append(c)
                    for index in range(len(fam_range)-1):
                        for i in range(startCell[0],endCell[0]+1):
                            cellDef = 'A'+ str(i)                        
                            cellFrom = str(fam_range[index]) + str(i)
                            cellTo = str(fam_range[index+1]) + str(i)
                            App.ActiveDocument.Spreadsheet.setAlias(cellFrom, '')
                            App.ActiveDocument.recompute()
                            App.ActiveDocument.Spreadsheet.setAlias(cellTo, App.ActiveDocument.Spreadsheet.getContents(cellDef))
                            App.ActiveDocument.recompute()
                            sfx = str(fam_range[index+1]) + '1'
                        suffix = App.ActiveDocument.Spreadsheet.getContents(sfx)
                    
                        filename = filePrefix + '_' + suffix + '.fcstd'
                        filePath = os.path.join(docDir, filename)
                    
                        FreeCAD.Console.PrintMessage("Saving file to %s\n" % filePath)
                        App.getDocument(filePrefix).saveCopy(filePath)
                        



form = MyButtons()
form.exec_()
#try:
if form.retStatus==1:
    routine1()
elif form.retStatus==2:
    routine2()
elif form.retStatus==3:
    routine3()
elif form.retStatus==4:
    routine4()
#except:
#    pass


