'''
Created on Sep 7, 2019

@author: Vicky Virus
'''

import subprocess
from ProjectPath import user_dir
from os import remove

class Report:
    
    fileForReport = user_dir+"\\Utilities\\Report.txt"
    
    def __init__(self):
        self.file = open(self.fileForReport, "a+")
    
    def start(self, stepName, desc):
        self.file.write("start" +"|"+stepName+"|"+desc+"|"+" "+"\n")
        
    def Pass(self, stepName, desc):
        self.file.write("pass" +"|"+stepName+"|"+desc+"|"+" "+"\n")
        
    def Fail(self, stepName, desc):
        self.file.write("fail" +"|"+stepName+"|"+desc+"|"+" "+"\n")
        
    def Warn(self, stepName, desc):
        self.file.write("warn" +"|"+stepName+"|"+desc+"|"+" "+"\n")
        
    def Info(self, stepName, desc):
        self.file.write("info" +"|"+stepName+"|"+desc+"|"+" "+"\n")
        
    def PassWithScreenShot(self, stepName, desc, imgPath):
        self.file.write("pass" +"|"+stepName+"|"+desc+"|"+imgPath+"\n")
        
    def FailWithScreenShot(self, stepName, desc, imgPath):
        self.file.write("fail" +"|"+stepName+"|"+desc+"|"+imgPath+"\n")    
        
    def stop(self):
        self.file.write("stop" +"|"+" "+"|"+" "+"|"+" "+"\n")
        
    def genExtentReoprt(self):
        self.file.close()
        print(self.fileForReport)
        subprocess.call(['java', '-jar', user_dir+'\\extentReport\\GenExtentReport.jar', self.fileForReport, user_dir+"\\extentReport"])
        remove(self.fileForReport)