'''
Created on Mar 15, 2019

@author: Vicky Virus
'''
from Utilities.ExcelHandler import excelHandler
from Utilities.WebDriver import WebDriver
from ProjectPath import user_dir
from Utilities.ExtentReportByJava import Report

scenarioWorkbook = user_dir+"\\testData\\MLM.xlsx"      #Workbook Path for test data
 
scenarioSheetName = "Sheet2"    #Sheet name for test data in above workbook
 
newExcel = excelHandler(scenarioWorkbook, scenarioSheetName)    #Object to handle above scenario workbook

newReport = Report()    #Instance of Report Class to generate extent report
 
testData = newExcel.getTestData(scenarioSheetName)      #To fetch and save data from excel as dictionary with row 0 as Key and below are value
 
print(len(testData))        #To Print To total row of test data
 
currentRow=1
 
for test in testData:       #To loop through each row of test data
      
    if(test["RunMode"]=="Y"):       #Only Run Mode = Y scenarios will Run
        newReport.start("new one "+str(currentRow), "nothing "+str(currentRow))     #Start Extent Report
        try:  
            driver = WebDriver("chrome")        #To Launch browser chrome, ie, or mozilla
            newReport.Pass("Launch Chrome", "Launched")
            
            xPaths = driver.loadXpaths("xpath.xlsx", "xPaths")      #To load xpath from xpath.xlsx workbook's xPaths sheet
                
            driver.getUrl("http://www.google.com")      #To Navigate to URL
            newReport.Pass("Navigate to URL"," ")

            newReport.PassWithScreenShot("Navigated", " ", driver.takeScreenShotPass())
            
            driver.getLocator("textbox_xpath").send_keys(test["District"])
              
            driver.getLocator("search_xpath").click()
              
            #chromeDriver.find_element_by_xpath("//*[@onclick='ChkloginTOP()']").click()
             
            newExcel.setExcelData(scenarioSheetName, currentRow, 0, "Passed")
            driver.close()
        except Exception:
            newExcel.setExcelData(scenarioSheetName, currentRow, 0, "Failed")
            print("exception occurred")
            driver.close()
             
        newReport.stop()   
    else: newExcel.setExcelData(scenarioSheetName, currentRow, 0, "Skipped")
       
    currentRow +=1
    
newExcel.updateChnages()
newReport.genExtentReoprt()