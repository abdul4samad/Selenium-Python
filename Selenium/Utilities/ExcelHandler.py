'''
Created on Mar 15, 2019

@author: Vicky Virus
'''

import xlrd
import subprocess
from ProjectPath import user_dir
from os import remove

class excelHandler:
    
    fileForExcel = user_dir+"\\Utilities\\updateData.txt"

    def __init__(self, filePath, sheetName):
        self.wb = xlrd.open_workbook(filePath)
        self.filePath = filePath
        self.sheetName=sheetName
        self.sheet= self.wb.sheet_by_name(sheetName)
        #=======================================================================
        # for names in self.wb.sheet_names():
        #     print(names)
        #     print(self.sheetIndex)
        #     if(names==self.sheetName):
        #         print(self.sheetName)
        #         break
        #     else: self.sheetIndex +=1
        #=======================================================================


    def getTestData(self, sheetName):
        self.sheet= self.wb.sheet_by_name(sheetName)
        testData=[]
        for row in range(1, self.sheet.nrows):
            tempDictionary={}
            for col in range(self.sheet.ncols):
                tempDictionary[self.sheet.cell_value(0, col)] = self.sheet.cell_value(row, col)
            testData.append(tempDictionary)
            print(tempDictionary)
        return testData
    
    def getXpaths(self, sheetName):
        sheet= self.wb.sheet_by_name(sheetName)
        tempDictionary={}
        for row in range(0, sheet.nrows):
            if(sheet.cell_value(row, 1)!='' and sheet.cell_value(row, 0)!=''):
                tempDictionary[sheet.cell_value(row, 0)] = sheet.cell_value(row, 1)
        return tempDictionary
    

    def setExcelData(self, sheetName, rowx, colx, data):
        fileForExcel = open(self.fileForExcel, "a+")
        fileForExcel.write(sheetName +"|"+str(rowx)+"|"+str(colx)+"|"+str(data)+"\n")
        fileForExcel.close()

    def currentExcel(self, newExcelPath, sheetName):
        self.wb = xlrd.open_workbook(newExcelPath)
        self.filePath = newExcelPath
        self.sheetName=sheetName
        self.sheet= self.wb.sheet_by_name(sheetName)
        
        
    def updateChnages(self):
        print(self.filePath)
        subprocess.call(['java', '-jar', user_dir+'\\Utilities\\UpdateExcel.jar', self.filePath, self.fileForExcel])
        remove(self.fileForExcel)
        