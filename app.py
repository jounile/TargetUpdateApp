#!/usr/bin/env python3.6
import os
import sys, stat
from subprocess import PIPE, run
from shutil import copyfile, rmtree

from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog, QAction, QTextBrowser, QHBoxLayout, QVBoxLayout, QLabel, QPushButton, QStatusBar, QWidget, QGridLayout, QLineEdit
from PyQt5.QtCore import Qt, QSize, QThread, pyqtSignal
from datetime import datetime

from ftplib import FTP, error_perm, all_errors
import telnetlib

# Default variables
global_BuildVersion = ""
global_IsConnectionTest = False

# Default connection params
global_ftpHost = "192.168.2.110"
global_ftpPort = 22
global_ftpUsername = "pi"
global_ftpPassword = "raspberry"

global_destinationDir = "/files"

global_telnetHost = "192.168.2.110"
global_telnetPort = 23

global_7zipPath = "C:\\Program Files (x86)\\7-Zip\\7z.exe"

class TelnetThread(QThread):

    telnetStatus = pyqtSignal(object)

    def __init__(self):
        QThread.__init__(self)

    def run(self):
        self.openTelnetConnection()
        if not global_IsConnectionTest:
            self.sendSyncCommand()
            self.sendRsetCommand()
        else:
            self.telnetStatus.emit("Skipping Telnet commands. global_IsConnectionTest is set to True")
        self.closeTelnetConnection()

    def openTelnetConnection(self):
        self.telnet = telnetlib.Telnet(global_telnetHost, int(global_telnetPort))
        try:
            self.telnet.set_debuglevel(9)
            self.telnet.read_until(b"MRA2-IC_H> ")
            self.telnetStatus.emit("Opened Telnet Connection")
        except ConnectionRefusedError as e:
            self.telnetStatus.emit("Telnet connection error" + e)
            pass

    def sendSyncCommand(self):
        self.telnet.write(b"sync\n")
        self.telnetStatus.emit("Send sync command")

    def sendRsetCommand(self):
        self.telnet.write(b"rset\n")
        self.telnetStatus.emit("Send rset command")

    def closeTelnetConnection(self):
        self.telnet.close()
        self.telnetStatus.emit("Closed Telnet connection")

class ExtractThread(QThread):

    extractStatus = pyqtSignal(object)

    # Paths
    tempBuildDir = "C:\\temp"
    tempArchivePath = ""

    def __init__(self):
        QThread.__init__(self)

    def run(self):
        self.extractArchiveToTemp()

    def extractArchiveToTemp(self):
        if self.tempArchivePath:
            extractPath = self.getExtractPath(self.tempArchivePath)
            if not os.path.isdir(extractPath):
                self.extractArchive(self.tempArchivePath)
            else:
                self.extractStatus.emit("Skipping extract. Dir " + extractPath + " exists")

    def getExtractPath(self, tempArchivePath):
        extractPath, ext = os.path.splitext(tempArchivePath)
        return extractPath
  
    def extractArchive(self, archivePath):
        archiveName = os.path.basename(archivePath)
        self.extractStatus.emit("Started extracting " + archiveName)
        self.buildVersion, ext = os.path.splitext(archiveName)

        # Set build version to a global variable so it can be used across threads
        global global_BuildVersion 
        global_BuildVersion = self.buildVersion
        
        xFlag = self.tempBuildDir
        xFlag += "\\"
        xFlag += archiveName
        oFlag = "-o"
        oFlag += self.tempBuildDir
        oFlag += "\\"
        oFlag += self.buildVersion
        command = [global_7zipPath, 'x', xFlag, oFlag, '-r']
        result = run(command, stdout=PIPE, stderr=PIPE, universal_newlines=True)

        self.extractStatus.emit(result.stdout)
        self.extractStatus.emit("Finished extracting " + archiveName)

class FtpThread(QThread):

    ftpStatus = pyqtSignal(object)
    tempBuildDir = "C:\\temp"

    def __init__(self):
        QThread.__init__(self)

    def run(self):
        # Log params
        self.ftpStatus.emit("host: " + global_ftpHost)
        self.ftpStatus.emit("port: " + str(global_ftpPort))
        self.ftpStatus.emit("username: " + global_ftpUsername)
        self.ftpStatus.emit("password: " + global_ftpPassword)

        self.ftpStatus.emit('Connecting...')
        self.ftp = FTP(global_ftpHost)
        try:
            result = self.ftp.login(user=global_ftpUsername, passwd=global_ftpPassword)
            self.ftpStatus.emit(result)
        except:
             self.ftpStatus.emit("FTP login failed")

        self.ftp.cwd(global_destinationDir)
        ls = []
        self.ftp.retrlines('LIST', ls.append)
        for entry in ls:
            self.ftpStatus.emit(entry)

        self.ftpStatus.emit('Connected')

        if not global_IsConnectionTest:
            self.removeOldBackup()
            self.createBackup()
            self.createDestinationDirs()
            self.copyDeploymentTable()
            self.copyHmiToTarget()
            self.copyHudToTarget()
            self.copyIcBinary()
            self.copyHudBinary()
        else:
            self.ftpStatus.emit("Skipping FTP file transfer. global_IsConnectionTest is set to True")

        self.ftp.quit()
        self.ftpStatus.emit('Disconnected')

    def copyDeploymentTable(self):
        src = "If1DeploymentTable-DIHMI.idt"
        self.ftp.cwd(global_destinationDir + '/hmi')
        self.ftp.storbinary('STOR If1DeploymentTable-DIHMI.idt', open(src,'rb'))
        self.ftpStatus.emit("Copied deployment table")

    def FtpRmTree(self, path):
        wd = self.ftp.pwd()
        try:
            names = self.ftp.nlst(path)
        except all_errors:
            self.ftpStatus.emit("Could not remove " + path)
            return
        for name in names:
            if os.path.split(name)[1] in (".", ".."): continue
            try:
                self.ftp.cwd(name)
                self.ftp.cwd(wd)
                self.FtpRmTree(name)
            except all_errors:
                self.ftp.delete(name)
                self.ftpStatus.emit("DELETE " + path)
        try:
            self.ftp.rmd(path)
            self.ftpStatus.emit("RMD " + path)
        except all_errors:
            self.ftpStatus.emit("Could not remove " + path)

    def removeOldBackup(self):
        hmi_old = global_destinationDir + "/hmi_old"
        if "hmi_old" in self.ftp.nlst():
            self.FtpRmTree(hmi_old)
        else:
            self.ftpStatus.emit("No hmi_old backup was found on target")
        hmihud_old = global_destinationDir + "/hmihud_old"
        if "hmihud_old" in self.ftp.nlst():
            self.FtpRmTree(hmihud_old)
        else:
            self.ftpStatus.emit("No hmihud_old backup was found on target")

    def createBackup(self):
        self.ftp.rename("hmi", "hmi_old")
        self.ftp.rename("hmihud", "hmihud_old")

    def createDestinationDirs(self):
        if "hmi" in self.ftp.nlst():
            self.ftpStatus.emit("Remote dir hmi exists")
        else:
            hmi = global_destinationDir + "/hmi"
            self.ftp.mkd(hmi)
        if "hmihud" in self.ftp.nlst():
            self.ftpStatus.emit("Remote dir hmihud exists")
        else:
            hmihud = global_destinationDir + "/hmihud"
            self.ftp.mkd(hmihud)

    def copyHudBinary(self):
        src = os.path.join(self.tempBuildDir, global_BuildVersion)
        src = os.path.join(src, "dihmi_bin")
        src = os.path.join(src, "bin")
        src = os.path.join(src, "hud")
        src = os.path.join(src, "high")
        src = os.path.join(src, "HUDHMIMain_IC_H")
        self.ftp.cwd(global_destinationDir + '/hmi')
        self.ftp.storbinary('STOR HUDHMIMain_IC_H', open(src,'rb'))
        self.ftpStatus.emit("Copied HUDHMIMain_IC_H to target")

    def copyIcBinary(self):
        src = os.path.join(self.tempBuildDir, global_BuildVersion)
        src = os.path.join(src, "dihmi_bin")
        src = os.path.join(src, "bin")
        src = os.path.join(src, "ic")
        src = os.path.join(src, "high")
        src = os.path.join(src, "ICHMIMain_IC_H")
        self.ftp.cwd(global_destinationDir + '/hmi')
        self.ftp.storbinary('STOR ICHMIMain_IC_H', open(src,'rb'))
        self.ftpStatus.emit("Copied ICHMIMain_IC_H to target")

    def copyHmiToTarget(self):
        src = os.path.join(self.tempBuildDir, global_BuildVersion)
        src = os.path.join(src, "dihmi_bin")
        src = os.path.join(src, "hmi")
        self.ftpStatus.emit("src: " + src)
        self.ftp.cwd(global_destinationDir + '/hmi')
        self.placeFiles(src)
        self.ftpStatus.emit("Copied /hmi to target")

    def copyHudToTarget(self):
        src = os.path.join(self.tempBuildDir, global_BuildVersion)
        src = os.path.join(src, "dihmi_bin")
        src = os.path.join(src, "hmihud")
        self.ftp.cwd(global_destinationDir + '/hmihud')
        self.placeFiles(src)
        self.ftpStatus.emit("Copied /hmihud to target")
        
    def placeFiles(self, src):
        for name in os.listdir(src):
            localpath = os.path.join(src, name)
            if os.path.isfile(localpath):
                self.ftpStatus.emit("STOR: " + name + " " + localpath)
                self.ftp.storbinary('STOR ' + name, open(localpath,'rb'))
            elif os.path.isdir(localpath):
                self.ftpStatus.emit("MKD: " + name)
                try:
                    self.ftp.mkd(name)
                except error_perm as e:
                    if not e.args[0].startswith('550'): 
                        raise
                self.ftp.cwd(name)
                self.placeFiles(localpath)
                self.ftp.cwd("..")

class TargetUpdateApp(QMainWindow): 

    def __init__(self): 
        super().__init__()

        # Configure
        self.isConnected = True
        self.isCleanUp = True

        self.setWindowTitle("TargetUpdateApp")
        self.createMenu()

        self.extractThread = ExtractThread()
        self.ftpThread = FtpThread()
        self.telnetThread = TelnetThread()   

        self.labelHost = QLabel("Host:")
        self.inputHost = QLineEdit(self)

        self.labelPort = QLabel("Port:")
        self.inputPort = QLineEdit(self)
        
        self.labelUsername = QLabel("Username:")
        self.inputUsername = QLineEdit(self)

        self.labelPassword = QLabel("Password:")
        self.inputPassword = QLineEdit(self)

        self.testConnectionsBtn = QPushButton("Test connections")
        self.testConnectionsBtn.clicked.connect(lambda:self.connectionsTest())

        self.syncRsetBtn = QPushButton("Sync / Rset")
        self.syncRsetBtn.clicked.connect(lambda:self.syncRset())

        self.hBoxLayout = QHBoxLayout()
        self.hBoxLayout.addWidget(self.labelHost)
        self.hBoxLayout.addWidget(self.inputHost)
        self.hBoxLayout.addWidget(self.labelPort)
        self.hBoxLayout.addWidget(self.inputPort)
        self.hBoxLayout.addWidget(self.labelUsername)
        self.hBoxLayout.addWidget(self.inputUsername)
        self.hBoxLayout.addWidget(self.labelPassword)
        self.hBoxLayout.addWidget(self.inputPassword)
        self.hBoxLayout.addWidget(self.testConnectionsBtn)
        self.hBoxLayout.addWidget(self.syncRsetBtn)

        # Set default connection params
        self.inputHost.setText(global_ftpHost)
        self.inputPort.setText(str(global_ftpPort))
        self.inputUsername.setText(global_ftpUsername)
        self.inputPassword.setText(global_ftpPassword)

        # Log
        self.logOutput = QTextBrowser()
        self.logOutput.verticalScrollBar().setValue(self.logOutput.verticalScrollBar().maximum())
        
        # Status bar
        self.statusBar = QStatusBar()
        
        # Layout
        self.widget = QWidget(self)
        self.setCentralWidget(self.widget)
        self.vBoxlayout = QVBoxLayout()
        self.vBoxlayout.addLayout(self.hBoxLayout)
        self.vBoxlayout.addWidget(self.logOutput)
        self.vBoxlayout.addWidget(self.statusBar)
        self.widget.setLayout(self.vBoxlayout)

        self.resize(800, 500)
        self.show()

    def updateConnectionParams(self):
        # Set params
        global global_telnetHost
        global_telnetHost = self.inputHost.text()
        global global_ftpHost
        global_ftpHost = self.inputHost.text()
        global global_ftpPort
        global_ftpPort = self.inputPort.text()
        global global_ftpUsername
        global_ftpUsername = self.inputUsername.text()
        global global_ftpPassword
        global_ftpPassword = self.inputPassword.text()

    def connectionsTest(self):
        self.logOutput.clear()
        self.logOutput.append("Test connections")
        self.updateConnectionParams()
        self.startFtpThread(True)

    def syncRset(self):
        self.startTelnetThread(False)

    def createMenu(self):
        extractAction = QAction("&Select archive", self)
        extractAction.triggered.connect(self.copyArchiveToTemp)

        mainMenu = self.menuBar()
        fileMenu = mainMenu.addMenu('&File')
        fileMenu.addAction(extractAction)

    def onExtractStatus(self, status):
        self.logOutput.append("Extract status - " + status)

    def onExtractFinished(self):
        self.logOutput.append("Extract Thread finished")
        if self.isConnected:
            self.startFtpThread(False)
        else:
            self.logOutput.append("Skipping FTP connection. isConnected is set to False")
            self.logOutput.append("Skipping Telnet connection. isConnected is set to False")
            if self.isCleanUp:
                self.removeExtractedFromTemp()
                self.removeArchiveFromTemp()

    def onFtpFinished(self):
        self.logOutput.append("FTP Thread finished")
        if self.isCleanUp:
            self.removeExtractedFromTemp()
            self.removeArchiveFromTemp()
        self.logOutput.append("Done")
        
    def onTelnetStatus(self, status):
        self.logOutput.append("Telnet status - " + status)

    def onTelnetFinished(self):
        self.logOutput.append("Telnet Thread finished")
        self.logOutput.append("Done")

    def startExtractThread(self):
        try:
            self.extractThread.extractStatus.connect(self.onExtractStatus, Qt.UniqueConnection)
        except TypeError:
            # connected already
            pass
        try:
            self.extractThread.finished.connect(self.onExtractFinished, Qt.UniqueConnection)
        except TypeError:
            # connected already
            pass
        self.extractThread.start()

    def removeExtractedFromTemp(self):
        extractedDir = self.extractThread.getExtractPath(self.extractThread.tempArchivePath)
        if os.path.isdir(extractedDir):
            rmtree(extractedDir)
            self.logOutput.append("Removed " + extractedDir)

    def removeArchiveFromTemp(self):
        archivePath = self.extractThread.tempArchivePath
        if os.path.isfile(archivePath):
            os.remove(archivePath)
            self.logOutput.append("Removed " + archivePath)

    def openFileDialog(self):    
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self,"QFileDialog.getOpenFileName()", "", filter="*.7z", options=options)
        return filePath

    def copyArchiveToTemp(self):
        self.logOutput.clear()
        self.archivePath = self.openFileDialog()
        if self.archivePath:
            self.extractThread.tempArchivePath = os.path.join(self.extractThread.tempBuildDir, os.path.basename(self.archivePath))
            if not os.path.isfile(self.extractThread.tempArchivePath):
                copyfile(self.archivePath, self.extractThread.tempArchivePath)
                self.logOutput.append("Copied archive to " + self.extractThread.tempArchivePath)
                self.startExtractThread()
            else:
                self.logOutput.append("Skipping copying. Archive " + self.extractThread.tempArchivePath + " already exists")

    def onFtpStatus(self, status):
        self.logOutput.append("FTP status - " + status)

    def startFtpThread(self, isConnectionTest):
        global global_IsConnectionTest
        global_IsConnectionTest = isConnectionTest
        try:
            self.ftpThread.ftpStatus.connect(self.onFtpStatus, Qt.UniqueConnection)
        except TypeError:
            # connected already
            pass
        try:
            self.ftpThread.finished.connect(self.onFtpFinished, Qt.UniqueConnection)
        except TypeError:
            # connected already
            pass
        self.ftpThread.start()
        self.logOutput.append("FTP Thread started")

    def startTelnetThread(self, isConnectionTest):
        global global_IsConnectionTest
        global_IsConnectionTest = isConnectionTest
        try:
            self.telnetThread.telnetStatus.connect(self.onTelnetStatus, Qt.UniqueConnection)
        except TypeError:
            # connected already
            pass
        try:
            self.telnetThread.finished.connect(self.onTelnetFinished, Qt.UniqueConnection)
        except TypeError:
            # connected already
            pass
        self.telnetThread.start()
        self.logOutput.append("Telnet Thread started")

    def closeApplication(self):
        sys.exit()
        
def main():
    App = QApplication(sys.argv) 
    targetUpdateApp = TargetUpdateApp()
    targetUpdateApp.show()
    exit_code = App.exec()
    sys.exit(exit_code)
    
if __name__ == '__main__': 
    main()
