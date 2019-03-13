#!python
from fdfgen import forge_fdf
import subprocess
import time
import uuid
import csv
import sys
import os

from PyQt5 import QtCore, QtWidgets, uic
from PyQt5.QtCore import pyqtSignal

#defines for columns
dataColumn = 2
filenameColumn = 3

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

    def setup(self, csvFile, fieldData, CSVhasHeaders):
        self.csvFile = csvFile
        self.fieldData = fieldData
        self.CSVhasHeaders = hasHeaders

    def run(self):
        print("started document generation")

        #build field list from CSV file
        self.csvData = list()
        with open(self.csvFile, "r") as f:
            reader = csv.reader(f)
            if self.hasHeaders:
                headers = next(reader)
            else:
                headers = None

            for row in reader:
                thisRow = dict()
                for column, each in enumerate(row):
                    if column in self.csvColumns:
                        thisRow[column] = each
                csvData.append(thisRow)

        # rows = dict()
        # for eachRow in csvData:
        #     # fields = [('name', names_list[eachNum])]
        #     #build fields
        #     thisRowFields = list()
        #     for eachCol in range(len(self.fieldData)):
        #         if self.fieldData[eachCol] is not None:
        #             if self.fieldData[fieldName][0] is "data":
        #                 thisColumn = (fieldName, self.fieldData[fieldName][1])
        #             if self.fieldData[fieldName][0] is "link":
        #                 thisColumn = (fieldName, eachRow[self.fieldData[fieldName][1]])
        #             thisRowFields.append(thisColumn)
        #     rows.append(thisRowFields)
        self.finished.emit(["not working"])

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

class previewWindow(QtWidgets.QWidget):
    backSignal = pyqtSignal()
    nextSignal = pyqtSignal()
    def __init__(self):
        super(previewWindow, self).__init__(parent=None)

    def setup(self, fields, nameColumn):
        self.backButton.clicked.connect(self.backSignalFunction)
        self.nextButton.clicked.connect(self.nextSignalFunction)
        
        self.tableWidget.setColumnCount(len(fields[0])+1)
        headers = [i[0] for i in fields[0]]
        headers.append("Filename")
        print("headers: {}".format(headers))
        self.tableWidget.setHorizontalHeaderLabels(headers)

        for eachRow in fields:
            row = self.tableWidget.rowCount()
            self.tableWidget.insertRow(row)
            for col, eachColumn in enumerate(eachRow):
                print("doing col: {} with: {}".format(col, eachColumn))
                self.tableWidget.setItem(row, col, QtWidgets.QTableWidgetItem(eachColumn[1]))
            filename = eachRow[nameColumn][1].replace(" ", "_") + ".pdf"
            self.tableWidget.setItem(row, self.tableWidget.columnCount()-1, QtWidgets.QTableWidgetItem(filename))

    def backSignalFunction(self):
        self.backSignal.emit()

    def nextSignalFunction(self):
        self.nextSignal.emit()

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
        self.setNameButton.clicked.connect(self.setNameButtonAction)
        self.outputDirButton.clicked.connect(self.outputDirAction)
        self.helpButton.clicked.connect(self.helpAction)
        self.genButton.clicked.connect(self.generatorAction)
        self.tableWidget.itemChanged[QtWidgets.QTableWidgetItem].connect(self.tableItemChanged)

        self.debugPushButton.clicked.connect(self.debugAction)

        self.numberLinks = 0

        self.maxThreads = 1
        self.runningThreads = 0
        self.threadList = dict()

    def debugAction(self):
        templateSelection = self.tableWidget.selectedIndexes()
        if len(templateSelection) != 4:
            QtWidgets.QMessageBox.information(
                self, 'Wrong Selection', "Please unlink one field at a time.", QtWidgets.QMessageBox.Ok)
        else:
            templateRow = templateSelection[2].row()
            model = self.tableWidget.model()
            index = model.createIndex(templateRow, dataColumn)
            currentLink = model.data(index, QtCore.Qt.UserRole)
            if currentLink is not None:
                self.debugLineEdit.setText("{}, {}, {}".format(currentLink[0], currentLink[1], currentLink[2]))
            else:
                self.debugLineEdit.setText("None")

    def generatorAction(self):
        self.myGen = generateFields(self)
        self.myGen.finished.connect(self.genFieldsFinished)
        #template, csvFile, csvColumns, outputDir, fieldData, fieldOrder
        if self.outputLineEdit.text() == "":
            QtWidgets.QMessageBox.information(
                self, 'No Output Directory Selected', "Please select an directory to save\ngenerated files to on the next screen.", QtWidgets.QMessageBox.Ok)
            if not self.outputDirAction():
                return
        if self.numberLinks == 0:
            QtWidgets.QMessageBox.information(
                self, 'No Link Columns', "Please link at least one column from the CSV table.", QtWidgets.QMessageBox.Ok)
        else:

            self.filenameColumn = self.csvFieldNumberList[0]
            if self.checkBox.isChecked():
                hasHeaders = True
            else:
                hasHeaders = False
            self.myGen.setup(self.templateLineEdit.text(), self.csvLineEdit.text(), self.csvFieldNumberList, self.outputDir, self.fieldFillData, self.fieldFillOrder, hasHeaders)
            self.myGen.start()

    def genFieldsFinished(self, fields):
        print(fields)
        self.fields = fields
        thispreviewWindow = previewWindow()
        self.previewWindow = uic.loadUi("preview.ui", baseinstance=thispreviewWindow)
        self.previewWindow.setup(fields, self.filenameColumn)
        self.previewWindow.backSignal.connect(self.previewBack)
        self.previewWindow.nextSignal.connect(self.previewNext)
        self.hide()
        self.previewWindow.show()
    
    def previewNext(self):

        self.startDocumentThread()
    
    def previewBack(self):
        self.previewWindow.close()
        self.show()

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
        else:
            if len(self.fields) == 0:
                self.hide()
                self.genDocumentsFinished()

    def threadFinished(self, thisuuid):
        print("Finished uuid: {}".format(thisuuid))
        thisThread = self.threadList.pop(thisuuid, None)
        self.startDocumentThread()
        if thisThread == None:
            print("error, thread id not found!!")
     
    def genDocumentsFinished(self):
        reply = QtWidgets.QMessageBox.information(
                self, 'Completed', "All files Generated.", QtWidgets.QMessageBox.Ok)
        if reply:
            self.close()

    def setNameButtonAction(self):
        print("setbuttonaction")
        templateSelection = self.tableWidget.selectedIndexes()
        if len(templateSelection) != 4:
            QtWidgets.QMessageBox.information(
                self, 'Wrong Selection', "You can only set 1 field to the filename.", QtWidgets.QMessageBox.Ok)
        else:
            templateRow = templateSelection[2].row()
            model = self.tableWidget.model()
            index = model.createIndex(templateRow, dataColumn)
            currentLink = model.data(index, QtCore.Qt.UserRole)
            if currentLink is None:
                QtWidgets.QMessageBox.information(
                    self, 'Opps', "You can only set filename to linked fields.", QtWidgets.QMessageBox.Ok)
            elif currentLink[0] != "link":
                QtWidgets.QMessageBox.information(
                    self, 'Opps', "You can only set filename to linked fields.", QtWidgets.QMessageBox.Ok)
            else:
                if currentLink[2]:
                    QtWidgets.QMessageBox.information(
                        self, 'Opps', "This is already the filename......", QtWidgets.QMessageBox.Ok)
                else:
                    signalBocker = QtCore.QSignalBlocker(self.tableWidget)

                    #remove current filename
                    for eachRow in range(self.tableWidget.rowCount()):
                        eachIndex = model.createIndex(eachRow, dataColumn)
                        itemType = model.data(eachIndex, QtCore.Qt.UserRole)
                        if itemType is not None:
                            if itemType[2] == True:
                                self.tableWidget.item(eachRow, filenameColumn).setText("No")
                                newItemType = itemType[:-1] + (False,)
                                model.setData(eachIndex, newItemType, QtCore.Qt.UserRole)
                                break
                        else:
                            print("itemType is none")
                    #set new filename
                    newLink =  currentLink[:-1] + (True,)
                    model.setData(index, newLink, QtCore.Qt.UserRole)
                    self.tableWidget.item(templateRow, filenameColumn).setText("Yes")

    def unlinkAction(self):
        templateSelection = self.tableWidget.selectedIndexes()
        if len(templateSelection) != 4:
            QtWidgets.QMessageBox.information(
                self, 'Wrong Selection', "Please unlink one field at a time.", QtWidgets.QMessageBox.Ok)
        else:
            templateRow = templateSelection[2].row()
            model = self.tableWidget.model()
            index = model.createIndex(templateRow, dataColumn)
            currentLink = model.data(index, QtCore.Qt.UserRole)
            print("currentLink: {}".format(currentLink))
            if currentLink is not None:
                if currentLink[0] == "link":
                    signalBocker = QtCore.QSignalBlocker(self.tableWidget)
                    model.setData(index, None, QtCore.Qt.UserRole)     
                    item = self.tableWidget.item(templateSelection[2].row(), templateSelection[2].column())
                    item.setText("")
                    self.numberLinks = self.numberLinks - 1
                    if currentLink[2]:
                        self.tableWidget.item(templateSelection[2].row(), filenameColumn).setText("No")
                        if self.numberLinks > 0:
                            for eachRow in range(self.tableWidget.rowCount()):
                                eachIndex = model.createIndex(eachRow, dataColumn)
                                itemType = model.data(eachIndex, QtCore.Qt.UserRole)
                                if itemType is not None:
                                    if itemType[0] == "link":
                                        self.tableWidget.item(eachRow, filenameColumn).setText("Yes")
                                        newItemType = itemType[:-1] + (True,)
                                        model.setData(eachIndex, newItemType, QtCore.Qt.UserRole)
                                        break
                else:
                    QtWidgets.QMessageBox.information(
                    self, 'No Link', "Nothing to unlink.", QtWidgets.QMessageBox.Ok)
            else:
                QtWidgets.QMessageBox.information(
                self, 'No Link', "Nothing to unlink.", QtWidgets.QMessageBox.Ok)

    def linkAction(self):
        """Update the template table to display information about what will be in each document.
            We are storing the actual information about the column to link to in the CSV file in
            the userdata section of the model."""

        templateSelection = self.tableWidget.selectedIndexes()
        columnSelection = self.csvTable.selectedIndexes()

        print(templateSelection)
        print(columnSelection)
        print(len(templateSelection))

        if len(templateSelection) != 4 or len(columnSelection) != 1:
            QtWidgets.QMessageBox.information(
                self, 'Wrong Selection', "You must select exactly 1 field and 1 column!", QtWidgets.QMessageBox.Ok)
        else:
            print("link Action!")
            csvRow = columnSelection[0].row()
            item = self.tableWidget.item(templateSelection[2].row(), templateSelection[2].column())
            signalBocker = QtCore.QSignalBlocker(self.tableWidget)
            index = self.tableWidget.model().createIndex(templateSelection[2].row(), dataColumn)
            useAsSaveName = True if self.numberLinks == 0 else False
            if useAsSaveName:
                saveItem = self.tableWidget.item(templateSelection[2].row(), filenameColumn)
                saveItem.setText("Yes")
            self.tableWidget.model().setData(index, ("link", csvRow, useAsSaveName), QtCore.Qt.UserRole)
            self.numberLinks = self.numberLinks + 1

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
            self.myFieldLoader.loadFinished.connect(self.loadingComplete)
            self.myFieldLoader.start()

    def updateFieldLoadBar(self, update):
        self.loadingWindow.progressBar.setValue(update)

    def loadingComplete(self, myList):
        if len(myList) > 0:
            self.loadingWindow.progressBar.setValue(100)
            time.sleep(0.1)
            signalBocker = QtCore.QSignalBlocker(self.tableWidget)
            self.tableWidget.setRowCount(0)
            for eachRow in myList:
                row = self.tableWidget.rowCount()
                self.tableWidget.insertRow(row)
                self.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(eachRow['FieldName']))
                self.tableWidget.setItem(row, 1, QtWidgets.QTableWidgetItem(eachRow['FieldType']))
                self.tableWidget.setItem(row, 2, QtWidgets.QTableWidgetItem(""))
                self.tableWidget.setItem(row, 3, QtWidgets.QTableWidgetItem("No"))
            tableHeader = self.tableWidget.horizontalHeader()
            tableHeader.setSectionResizeMode(0,3)
            tableHeader.setSectionResizeMode(1,3)
            tableHeader.setSectionResizeMode(2,1)
            self.loadingWindow.close()
            signalBocker = None #just to prevent pylinting errors of unsed varable
            self.linkColumnButton.setEnabled(True)
            self.unlinkColumnButton.setEnabled(True)
            self.setNameButton.setEnabled(True)
        else:
            self.loadingWindow.close()
            self.templateLineEdit.setText("")
            QtWidgets.QMessageBox.information(
                self, 'No Fillable Fields', "PDF Template Filler only works on documents with fillable forms", QtWidgets.QMessageBox.Ok)

    def tableItemChanged(self, tableItem):
        print("item changed: {}".format(tableItem.text()))
        row = tableItem.row()
        col = tableItem.column()
        model = self.tableWidget.model()
        index = model.createIndex(row, col)
        if len(tableItem.text()) > 0:
            model.setData(index, ("data", tableItem.text()), QtCore.Qt.UserRole)
        else:
            model.setData(index, None, QtCore.Qt.UserRole)

    def outputDirAction(self):
        fileDialog = QtWidgets.QFileDialog()
        fileDialog.setFileMode(QtWidgets.QFileDialog.Directory)
        fileDialog.setAcceptMode(0)  # open
        # fileDialog.setNameFilters(["PDF Files (*.pdf)"])
        # fileDialog.selectNameFilter("PDF Files (*.pdf)")
        if (fileDialog.exec_()):
            fileList = fileDialog.selectedFiles()[0]
            self.outputLineEdit.setText(fileList)
            return True
        else:
            return False

    def helpAction(self):
        self.helpWindow = uic.loadUi("help.ui")
        self.helpWindow.show()




app = QtWidgets.QApplication(sys.argv)
theMainWindow = mainWindow()
theMainWindow = uic.loadUi("untitled.ui", baseinstance=theMainWindow)
theMainWindow.setup()
theMainWindow.show()
sys.exit(app.exec_())

