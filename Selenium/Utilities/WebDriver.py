'''
Created on Mar 23, 2019

@author: Vicky Virus
'''

from selenium import webdriver
from ProjectPath import user_dir
from _datetime import datetime
from Utilities import ExcelHandler
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotVisibleException, NoSuchElementException, StaleElementReferenceException
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

class WebDriver:
    '''
    classdocs
    '''
    driver=""
    testData=""

    def __init__(self, browserName):
        '''
        Constructor
        '''
        if(str(browserName).lower()=="chrome"):
            capabilities = DesiredCapabilities.CHROME.copy()
            capabilities['acceptInsecureCerts'] = True
            self.driver = webdriver.Chrome(user_dir+"\\resources\\drivers\\chromedriver.exe")
            self.driver.delete_all_cookies()
            self.driver.maximize_window()
        
        elif(str(browserName).lower()=="mozilla"):
            capabilities = DesiredCapabilities.FIREFOX.copy()
            capabilities['acceptInsecureCerts'] = True
            self.driver = webdriver.Firefox(user_dir+"\\resources\\drivers\\geckodriver.exe")
            self.driver.delete_all_cookies()
            self.driver.maximize_window()
        
        elif(str(browserName).lower()=="ie"):
            capabilities = DesiredCapabilities.INTERNETEXPLORER.copy()
            capabilities['acceptInsecureCerts'] = True
            self.driver = webdriver.Ie(user_dir+"\\resources\\drivers\\IEDriverServer.exe")
            self.driver.delete_all_cookies()
            self.driver.maximize_window()
        
    def loadXpaths(self, fileName, sheetName):
        xpathFromExcel = ExcelHandler.excelHandler(user_dir+"\\testData\\"+fileName, sheetName)   
        self.testData = xpathFromExcel.getXpaths(sheetName)
        return self.testData
                
    def getUrl(self, url):
        self.driver.get(url)
        
    def takeScreenShotPass(self):
        now = datetime.now()
        imgFile = user_dir+"\\screenshots\\pass\\pass_"+now.strftime("%Y%m%d%H%M%S")+".png"
        self.driver.get_screenshot_as_file(imgFile)
        return imgFile
        
    def takeScreenShotFail(self):
        now = datetime.now()
        imgFile = user_dir+"\\screenshots\\fail\\fail_"+now.strftime("%Y%m%d%H%M%S")+".png"
        self.driver.get_screenshot_as_file(imgFile)
        return imgFile
        
    def close(self):
        self.driver.quit()
    
    run=0    
    def getLocator(self, key):
        webElement=''
        try:
            if(self.run<2):
                if(str(key).endswith("_id")):
                    webElement = self.driver.find_element_by_id(self.testData[key])
                elif(str(key).endswith("_xpath")):
                    webElement = self.driver.find_element_by_xpath(self.testData[key])
                elif(str(key).endswith("_lnk")):
                    webElement = self.driver.find_element_by_partial_link_text(self.testData[key])
                elif(str(key).endswith("_tag")):
                    webElement = self.driver.find_element_by_tag_name(self.testData[key])
                elif(str(key).endswith("_class")):
                    webElement = self.driver.find_element_by_class_name(self.testData[key])
                elif(str(key).endswith("_css")):
                    webElement = self.driver.find_element_by_css_selector(self.testData[key])
                else:
                    webElement = self.driver.find_element_by_xpath(str(key))
        except NoSuchElementException or StaleElementReferenceException:
            print("Element not found")
            self.run+=1
            webElement = self.explicitWait(3000, self.getLocator(key))
            
        self.elementHighlight(webElement)   
        self.run=0    
        return webElement
    
    def explicitWait(self, seconds, webElement):
        wait = WebDriverWait(self.driver, seconds, poll_frequency=1, ignored_exceptions=[ElementNotVisibleException, NoSuchElementException])
        return wait.until(EC.visibility_of_element_located((webElement)))
    
    def elementHighlight(self, webElement):
        self.driver.execute_script("arguments[0].setAttribute('style', 'background:yellow; border:2px solid red;');", webElement)
        self.driver.implicitly_wait(0.05)
        self.driver.execute_script("arguments[0].setAttribute('style', 'border: solid 2px green ');", webElement)