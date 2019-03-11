#!python
from fdfgen import forge_fdf
import subprocess
import time
import uuid
import csv
import sys
import os

from PyQt5 import QtCore, QtWidgets, QtGui, uic
from PyQt5.QtCore import pyqtSignal

class fieldLoader(QtCore.QThread):
    loadFinished = pyqtSignal(list)
    updateProgress = pyqtSignal(int)
    noTemplateName = pyqtSignal()
    def __init__(self, parent=None):
        super(fieldLoader, self).__init__(parent)
        self.templateName = None

    def setTemplateName(self, template_name):
        self.templateName = template_name

    def run(self):
        print("starting process")
        
        if self.templateName is None:
            self.noTemplateName.emit()
        else:
            subprocess.call("pdftk " + self.templateName + " dump_data_fields output fields.txt")

            field_list = list()
            with open("fields.txt", "r") as fields_file:
                lines = fields_file.readlines()
            subprocess.Popen("del fields.txt", shell=True)
            totalLines = len(lines)
            linesList = list()
            myDict = None
            for lineNumber, eachLine in enumerate(lines):
                    self.updateProgress.emit(lineNumber/totalLines*100)
                    if "---" in eachLine:
                        continue
                    if "FieldType" in eachLine:
                        myDict = dict()
                        linesList.append(myDict)
                        start = eachLine.find(":")
                        field_type = eachLine[start + 2:-1]
                        myDict["FieldType"] = field_type
                    if "FieldName" in eachLine:
                        start = eachLine.find(":")
                        field_name = eachLine[start + 2: -1]
                        myDict["FieldName"] = field_name
            self.loadFinished.emit(linesList)

class generateFields(QtCore.QThread):
    finished = pyqtSignal(list)
    updateProgress = pyqtSignal(int)
    def __init__(self, parent=None):
        super(generateFields, self).__init__(parent)

    def setup(self, template, csvFile, csvColumns, outputDir, fieldData, fieldOrder, ignoreFirstRow):
        self.template = template
        self.csvFile = csvFile
        self.csvColumns = csvColumns
        self.outputDir = outputDir
        self.fieldData = fieldData
        self.fieldOrder = fieldOrder
        self.ignoreFirstRow = ignoreFirstRow

        self.csvData = None

    def run(self):
        print("started document generation")

        #build field list from CSV file
        self.csvData = list()
        with open(self.csvFile, "r") as f:
            reader = csv.reader(f)
            if self.ignoreFirstRow:
                next(reader)
            for row in reader:
                thisRow = dict()
                for column, each in enumerate(row):
                    if column in self.csvColumns:
                        thisRow[column] = each
                self.csvData.append(thisRow)
        fields = list()
        for eachRow in self.csvData:
            # fields = [('name', names_list[eachNum])]
            #build fields
            thisRowFields = list()
            for fieldName in self.fieldData.keys():
                if self.fieldData[fieldName] is not None:
                    if self.fieldData[fieldName][0] is "data":
                        thisColumn = (fieldName, self.fieldData[fieldName][1])
                    if self.fieldData[fieldName][0] is "link":
                        thisColumn = (fieldName, eachRow[self.fieldData[fieldName][1]])
                    thisRowFields.append(thisColumn)
            fields.append(thisRowFields)

        # fdf = forge_fdf("", fields, [], [], [])
        # out_filename = self.outputDir + "\\" + names_list[eachNum].replace(" ", "_") + ".pdf"
        # fdf_file_name = out_filename + ".fdf"
        # fdf_file = open(fdf_file_name, "wb")
        # fdf_file.write(fdf)
        # fdf_file.close()
        # subprocess.Popen("pdftk Certificate.pdf fill_form " + fdf_file_name + " output " +
        #              out_filename)

        self.finished.emit(fields)

class genDocument(QtCore.QThread):
    finished = pyqtSignal(str)
    def __init__(self, parent=None):
        super(genDocument, self).__init__(parent)

    def setup(self, thisuuid, template_file, fields, filename, directory):
        self.uuid = thisuuid
        self.fields = fields
        self.template_file = template_file
        self.filename = filename
        self.directory = directory
    
    def run(self):
        fdf = forge_fdf("", self.fields, [],[],[])
        self.directory = self.directory.replace("/", "\\")
        out_filename = self.directory + "\\" + self.filename + ".pdf"
        fdf_filename = out_filename + ".fdf"
        fdf_file = open(fdf_filename, "wb")
        fdf_file.write(fdf)
        fdf_file.close()
        subprocess.call("pdftk " + self.template_file + " fill_form " + fdf_filename + " output " + out_filename)
        print(fdf_filename)
        subprocess.call("del " + fdf_filename, shell=True)
        self.finished.emit(self.uuid)


class mainWindow(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(mainWindow, self).__init__(parent)

    def setup(self):
        self.tableWidget.setStyleSheet("""QTableWidget::item:selected{ background-color: rgb(0, 120, 215); color: white}""")
        self.csvTable.setStyleSheet("""QTableWidget::item:selected{ background-color: rgb(0, 120, 215); color: white}""")
        self.loadTemplateButton.clicked.connect(self.loadTemplateFields)
        self.loadColumnsButton.clicked.connect(self.loadCsvColumns)
        self.linkColumnButton.clicked.connect(self.linkAction)
        self.unlinkColumnButton.clicked.connect(self.unlinkAction)
        self.outputDirButton.clicked.connect(self.outputDirAction)
        self.helpButton.clicked.connect(self.helpAction)
        self.genButton.clicked.connect(self.generatorAction)
        self.tableWidget.itemChanged[QtWidgets.QTableWidgetItem].connect(self.tableItemChanged)
        self.fieldFillData = None
        self.fieldFillOrder = None
        self.outputDir = None
        self.csvFieldNumberList = list()

        self.maxThreads = 1
        self.runningThreads = 0
        self.threadList = dict()
        self.filenameColumn = None
        self.fields = None

    def generatorAction(self):
        self.myGen = generateFields(self)
        self.myGen.finished.connect(self.genFieldsFinished)
        #template, csvFile, csvColumns, outputDir, fieldData, fieldOrder
        if self.outputDir == None:
            QtWidgets.QMessageBox.information(
                self, 'No Output Directory Selected', "Please select an directory to save\ngenerated files to on the next screen.", QtWidgets.QMessageBox.Ok)
            self.outputDirAction()
        if len(self.csvFieldNumberList) == 0:
            QtWidgets.QMessageBox.information(
                self, 'No Link Columns', "Please link at least one column from the CSV table.", QtWidgets.QMessageBox.Ok)
        else:
            self.filenameColumn = self.csvFieldNumberList[0]
            if self.checkBox.isChecked():
                ignoreFirstRow = True
            else:
                ignoreFirstRow = False
            self.myGen.setup(self.templateLineEdit.text(), self.csvLineEdit.text(), self.csvFieldNumberList, self.outputDir, self.fieldFillData, self.fieldFillOrder, ignoreFirstRow)
            self.myGen.start()

    def genFieldsFinished(self, fields):
        print(fields)
        self.fields = fields
        #start the first thread
        self.startDocumentThread()

    def startDocumentThread(self):
        while len(self.fields) > 0 and self.runningThreads < self.maxThreads:
            thisuuid = str(uuid.uuid4().hex)
            fields = self.fields.pop()
            filename = fields[self.filenameColumn][1].replace(" ", "_")
            print("starting a new document ({}) creation thread: {}".format(filename, thisuuid))

            newThread = genDocument(self)
            newThread.setup(thisuuid, self.templateLineEdit.text(), fields, filename, self.outputDir)
            newThread.finished.connect(self.threadFinished)
            newThread.start()
            self.threadList[thisuuid] = newThread

    def threadFinished(self, thisuuid):
        print("Finished uuid: {}".format(thisuuid))
        thisThread = self.threadList.pop(thisuuid, None)
        self.startDocumentThread()
        if thisThread == None:
            print("error, thread id not found!!")
            
            



    def genDocumentsFinished(self):
        QtWidgets.QMessageBox.information(
                self, 'Completed', "All files Generated.", QtWidgets.QMessageBox.Ok)

    def unlinkAction(self):
        templateSelection = self.tableWidget.selectedIndexes()
        if len(templateSelection) != 3:
            QtWidgets.QMessageBox.information(
                self, 'Wrong Selection', "Please unlink one field at a time.", QtWidgets.QMessageBox.Ok)
        else:
            templateRow = templateSelection[0].row()
            currentLink = self.fieldFillData[self.fieldFillOrder[templateRow]]
            if currentLink is not None:
                if currentLink[0] == "link":
                    csvRow = currentLink[1]
                    self.fieldFillData[self.fieldFillOrder[templateRow]] = None
                    self.csvFieldNumberList.remove(csvRow)
                    print("csvRows: {}".format(self.csvFieldNumberList))
                    item = self.tableWidget.item(templateSelection[2].row(), templateSelection[2].column())
                    signalBocker = QtCore.QSignalBlocker(self.tableWidget)
                    item.setText("")
                else:
                    QtWidgets.QMessageBox.information(
                    self, 'No Link', "Nothing to unlink.", QtWidgets.QMessageBox.Ok)
            else:
                QtWidgets.QMessageBox.information(
                self, 'No Link', "Nothing to unlink.", QtWidgets.QMessageBox.Ok)


    def linkAction(self):
        templateSelection = self.tableWidget.selectedIndexes()
        columnSelection = self.csvTable.selectedIndexes()

        print(templateSelection)
        print(columnSelection)
        print(len(templateSelection))

        if len(templateSelection) != 3 or len(columnSelection) != 1:
            QtWidgets.QMessageBox.information(
                self, 'Wrong Selection', "You must select exactly 1 field and 1 column!", QtWidgets.QMessageBox.Ok)
        else:
            print("link Action!")
            templateRow = templateSelection[0].row()
            csvRow = columnSelection[0].row()
            self.fieldFillData[self.fieldFillOrder[templateRow]] = ("link", csvRow)
            self.csvFieldNumberList.append(csvRow)
            print("csvRows: {}".format(self.csvFieldNumberList))
            item = self.tableWidget.item(templateSelection[2].row(), templateSelection[2].column())
            signalBocker = QtCore.QSignalBlocker(self.tableWidget)
            item.setText("LINK: {}".format(csvRow+1))




    def loadCsvColumns(self):
        fileDialog = QtWidgets.QFileDialog()
        fileDialog.setFileMode(QtWidgets.QFileDialog.ExistingFile)
        fileDialog.setAcceptMode(0)  # open
        fileDialog.setNameFilters(["CSV (*.csv *.txt)"])
        fileDialog.selectNameFilter("CSV (*.csv *.txt)")
        if (fileDialog.exec_()):
            filename = fileDialog.selectedFiles()[0]
            self.csvLineEdit.setText(filename)

            with open(filename, "r") as f:
                reader = csv.reader(f)
                row = next(reader)
                for each in row:
                    rowNumber = self.csvTable.rowCount()
                    self.csvTable.insertRow(rowNumber)
                    self.csvTable.setItem(rowNumber, 0, QtWidgets.QTableWidgetItem(each))


    def loadTemplateFields(self):
        print("button clicked")
        fileDialog = QtWidgets.QFileDialog()
        fileDialog.setFileMode(QtWidgets.QFileDialog.ExistingFile)
        fileDialog.setAcceptMode(0)  # open
        fileDialog.setNameFilters(["PDF Files (*.pdf)"])
        fileDialog.selectNameFilter("PDF Files (*.pdf)")
        if (fileDialog.exec_()):
            fileList = fileDialog.selectedFiles()[0]
            self.templateLineEdit.setText(fileList)

            self.loadingWindow = uic.loadUi("loadingFields.ui")
            self.loadingWindow.show()
            self.myFieldLoader = fieldLoader(self)
            self.myFieldLoader.setTemplateName(fileList)
            self.myFieldLoader.updateProgress.connect(self.updateFieldLoadBar)
            self.myFieldLoader.loadFinished.connect(self.loaddingComplete)
            self.myFieldLoader.start()

    def updateFieldLoadBar(self, update):
        self.loadingWindow.progressBar.setValue(update)

    def loaddingComplete(self, myList):
        if len(myList) > 0:
            self.loadingWindow.progressBar.setValue(100)
            time.sleep(0.1)
            self.fieldFillData = dict()
            self.fieldFillOrder = dict()
            signalBocker = QtCore.QSignalBlocker(self.tableWidget)
            self.tableWidget.setRowCount(0)
            for eachRow in myList:
                row = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row)
                self.fieldFillOrder[row] = eachRow['FieldName']
                self.fieldFillData[eachRow['FieldName']] = None
                self.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(eachRow['FieldName']))
                self.tableWidget.setItem(row, 1, QtWidgets.QTableWidgetItem(eachRow['FieldType']))
                self.tableWidget.setItem(row, 2, QtWidgets.QTableWidgetItem(""))
            tableHeader = self.tableWidget.horizontalHeader()
            tableHeader.setSectionResizeMode(0,3)
            tableHeader.setSectionResizeMode(1,3)
            tableHeader.setSectionResizeMode(2,1)
            self.loadingWindow.close()
            print("Order: {}".format(self.fieldFillOrder))
            print("Data: {}".format(self.fieldFillData))
            signalBocker = None #just to prevent pylinting errors of unsed varable
            self.linkColumnButton.setEnabled(True)
            self.unlinkColumnButton.setEnabled(True)
        else:
            self.loadingWindow.close()
            self.templateLineEdit.setText("")
            QtWidgets.QMessageBox.information(
                self, 'No Fillable Fields', "PDF Template Filler only works on documents with fillable forms", QtWidgets.QMessageBox.Ok)




    def tableItemChanged(self, tableItem):
        print("item changed: {}".format(tableItem.text()))
        row = tableItem.row()
        if len(tableItem.text()) > 0:
            self.fieldFillData[self.fieldFillOrder[row]] = ("data", tableItem.text())
        else:
            self.fieldFillData[self.fieldFillOrder[row]] = None
        print("Order: {}".format(self.fieldFillOrder))
        print("Data: {}".format(self.fieldFillData))

    def outputDirAction(self):
        fileDialog = QtWidgets.QFileDialog()
        fileDialog.setFileMode(QtWidgets.QFileDialog.Directory)
        fileDialog.setAcceptMode(0)  # open
        # fileDialog.setNameFilters(["PDF Files (*.pdf)"])
        # fileDialog.selectNameFilter("PDF Files (*.pdf)")
        if (fileDialog.exec_()):
            fileList = fileDialog.selectedFiles()[0]
            self.outputLineEdit.setText(fileList)
            self.outputDir = fileList

    def helpAction(self):
        self.helpWindow = uic.loadUi("help.ui")
        self.helpWindow.show()




app = QtWidgets.QApplication(sys.argv)
theMainWindow = mainWindow()
theMainWindow = uic.loadUi("untitled.ui", baseinstance=theMainWindow)
theMainWindow.setup()
theMainWindow.show()
sys.exit(app.exec_())

