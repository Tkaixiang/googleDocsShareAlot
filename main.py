from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QLabel, QWidget, QPushButton, QApplication, QGridLayout, QLabel,\
                            QFileDialog, QTextEdit, QMessageBox, QLineEdit, QProgressBar
import sys
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import API
import os

class PyQtLayout(QWidget):
 
    def __init__(self):
        super().__init__()
        self.signInButton = ""
        self.signInStatus = ""
        self.selectDoc = ""
        self.selectDocButton = ""
        self.create = ""
        self.createButton = ""
        self.fileSelector = ""
        self.grid = QGridLayout()
        self.gauth = ""
        self.wordPath = ""
        self.CSVPath = ""
        self.ENames = ""
        self.FName = ""
        self.signedIn = False
        self.checkGauth()
        self.UI()
    
    def selectFile(self):
        path = QFileDialog.getOpenFileName(self,"Select a Word Document", "","Word Document (*.doc *.docx)")
        self.wordPath = path[0]

        fileName = path[0].split("/")
        fileName = fileName[len(fileName)-1]
        
        if (self.wordPath == ""):
            self.selectDocButton.setText("Select a Word Document")
        else:
            self.selectDocButton.setText(fileName)

        if (self.ENames != "") and (self.wordPath != "") and (self.FName != ""):
            self.createButton.setEnabled(True)
    
    def selectCSVMethod(self):
        path = QFileDialog.getOpenFileName(self,"Select a CSV List", "","CSV List (*.csv)")
        self.CSVPath = path[0]

        fileName = path[0].split("/")
        fileName = fileName[len(fileName)-1]
        if (self.CSVPath == ""):
            self.selectCSVButton.setText("Select a CSV List")
            self.CSVDetailsList.setPlainText("No student details loaded. Select a CSV to load student details")
        else:
            self.selectCSVButton.setText(fileName)
            studentList = API.processCSV(self.CSVPath)
            self.ENames = studentList
            studentString = "Loaded a total of <u><b>" + str(len(studentList)) + "</b></u> Student-Email pairs. <ul>"
            for x in studentList:
                studentString += ("<li> " + x + " - " + studentList[x] + "</li>") 
            
            studentString += "</ul>"
            self.CSVDetailsList.setText(studentString)

        if (self.ENames != "") and (self.wordPath != "") and (self.FName != ""):
            self.createButton.setEnabled(True)
        
    def checkGauth(self):
        if os.path.exists('credentials/credentials.json'):
            gauth = API.loadauth()
            if (gauth == False):
                os.remove('credentials/credentials.json')
                return
            self.gauth = gauth
            self.signedIn = True

    def signIn(self):
        messageBox = QMessageBox()
        messageBox.setWindowTitle("Redirecting you to Google authentication page")
        messageBox.setText("You are being redirected to the Google authentication page in your default browser. Kindly follow the instructions on that page to login 😊. <b>Press 'OK' to continue</b>")
        messageBox.setIcon(QMessageBox.Information)
        messageBox.exec_()

        if not os.path.exists("credentials"):
            os.mkdir("credentials")
        gauth = API.oauth()
        if (gauth == False):
            messageBox = QMessageBox()
            messageBox.setWindowTitle("Authentication Error")
            messageBox.setText("Oops, there was an error logging into your account. Please ensure that you allow the permissions requested in the authentication page. <b>Kindly click on the 'Sign In' button and try again.</b>")
            messageBox.setIcon(QMessageBox.Warning)
            messageBox.exec_()
            self.signInStatus.setText("Oops, authentication failure, please try again.")
        else:
            #Signed in successfully, load remaining widgets
            self.signedIn = True
            self.gauth = gauth
            self.UI()

    def inputFNameMethod(self):
        self.FName = self.inputFNameField.text()
        if (self.ENames != "") and (self.wordPath != "") and (self.FName != ""):
            self.createButton.setEnabled(True)        

    def uploadDocs(self, name):
        student = [name, self.ENames[name]]
        API.upload(self.gauth, student, self.fid, self.wordPath)   

    def updateProgress(self, j):
        self.progress.setValue(j)
        QApplication.processEvents()

    def createMethod(self):
        self.createButton.setText("Creating and Sharing...")
        self.createButton.setEnabled(False)

        self.fid = API.createFolder(self.gauth, self.FName)
        flink = "https://drive.google.com/drive/u/0/folders/" + self.fid

        messageBox = QMessageBox()
        messageBox.setWindowTitle("Create and Share Docs")
        messageBox.setText("Completed! Click <a href='%s'>here</a> to view the folder containing all newly created docs." % flink) 

        names = self.ENames.keys()
        self.progress.setRange(0, len(names))
        self.progress.setValue(0)
        for i, name in enumerate(names): 
            self.uploadDocs(name)
            self.updateProgress(i+1)  

        self.createButton.setText("Create and Share")
        self.createButton.setEnabled(True)
        messageBox.exec_()
 
    def UI(self):
        for i in reversed(range(self.grid.count())): 
            self.grid.takeAt(i).widget().deleteLater()
        #Sign In Button
        self.signInButton = QPushButton('Sign In')
        self.signInButton.clicked.connect(self.signIn)

        #Sign In Status
        self.signInStatus = QLabel("Not signed in. Click the 'Sign In' button on the left to sign in.")

        #Add widgets to layout
        self.grid.addWidget(self.signInButton, 0, 0)
        self.grid.addWidget(self.signInStatus, 0, 1)

        
        #If Signed In, add the file selection & submit widgets in
        if (self.signedIn):
            self.signInStatus.setText("<b><u>Signed in successfully</u></b>. Welcome to GoogleDocsShareAlot")
            self.signInButton.setText("Switch Accounts")

            self.selectDoc = QLabel("1. Select a Word Document (.docx or .doc) to upload: ")
            self.selectDocButton = QPushButton('Select a Word Document')
            self.selectDocButton.clicked.connect(self.selectFile)

            self.selectCSV = QLabel("2. Select a CSV Student-Email List (.csv) to upload: ")
            self.selectCSVButton = QPushButton('Select a CSV')
            self.selectCSVButton.clicked.connect(self.selectCSVMethod)
            self.CSVDetailsList = QTextEdit()
            self.CSVDetailsList.setReadOnly(True)
            self.CSVDetailsList.setPlainText("No student details loaded. Select a CSV to load student details")

            self.inputFName = QLabel("3. Name of Google Drive folder to save the documents into: <br> (<b>Press 'Enter' to confirm the name</b>): ")
            self.inputFNameField = QLineEdit()
            self.inputFNameField.returnPressed.connect(self.inputFNameMethod)

            self.create = QLabel("4. Create and share the selected Word Document: ")
            self.createButton = QPushButton('Create and Share')
            self.createButton.setEnabled(False)
            self.createButton.clicked.connect(self.createMethod)

            self.progress = QProgressBar()
            self.progress.setValue(0)

            self.grid.addWidget(self.selectDoc, 1, 0)
            self.grid.addWidget(self.selectDocButton, 1, 1)
            self.grid.addWidget(self.selectCSV, 2, 0)
            self.grid.addWidget(self.selectCSVButton, 2, 1)
            self.grid.addWidget(self.CSVDetailsList, 3, 0, 1, 2)
            self.grid.addWidget(self.inputFName, 4, 0)
            self.grid.addWidget(self.inputFNameField, 4, 1)
            self.grid.addWidget(self.create, 5, 0)
            self.grid.addWidget(self.createButton, 5, 1)
            self.grid.addWidget(self.progress, 6, 0, 1, 2)


        self.setLayout(self.grid) #Set the grid layout
        self.setGeometry(400,400,500,400) #Set Starting Location & Size
        self.setWindowTitle("GoogleDocsShareAlot")
        self.show()

def main():
        #Initialise
        app = QApplication(sys.argv)
        app.setStyle("Fusion")
        app.setWindowIcon(QtGui.QIcon("icon.png"))
        RunUI = PyQtLayout() #Create the UI object
        sys.exit(app.exec_()) #Prevent app from closing until user clicks exit

if __name__ == "__main__":
    main()

