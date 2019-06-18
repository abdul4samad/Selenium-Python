'''
Created on Mar 15, 2019

@author: Vicky Virus
'''
from Utilities.ExcelHandler import excelHandler
from Utilities.WebDriver import WebDriver
from ProjectPath import user_dir

loc = user_dir+"\\testData\\MLM.xlsx"
 
sheetName = "Sheet2"
 
newExcel = excelHandler(loc, sheetName)
 
testData = newExcel.getTestData(sheetName)
 
print(testData.__len__())
 
currentRow=1
 
for test in testData:
      
    if(test["RunMode"]=="Y"):
          
       # try:
              
            driver = WebDriver("chrome")
            xPaths = driver.loadXpaths("xpath.xlsx", "xPaths")
            for xpath in xPaths:
                print(xpath)
            driver.getUrl("http://www.google.com")
            driver.takeScreenShotFail()
            driver.getLocator("textbox_xpath").send_keys(test["Email"])
             
            driver.getLocator("search_xpath").click()
             
            #chromeDriver.find_element_by_xpath("//*[@onclick='ChkloginTOP()']").click()
            
            newExcel.setExcelData(sheetName, currentRow, 0, "Passed")
            driver.close()
        #except Exception:
            newExcel.setExcelData(sheetName, currentRow, 0, "Failed")
            print("Error Occurred")
            driver.close()
            
      
    else: newExcel.setExcelData(sheetName, currentRow, 0, "Skipped")
      
    currentRow +=1
    
newExcel.updateChnages()
#===============================================================================
# 
# newExcel.setExcelData(sheetName, 2, 3, "allOK")
# newExcel.setExcelData(sheetName, 3, 3, "allOK")
# newExcel.setExcelData(sheetName, 4, 3, "allOK")
# newExcel.setExcelData(sheetName, 5, 3, "allOK")
#   
# newExcel.updateChnages()
#===============================================================================
