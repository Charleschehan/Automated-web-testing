#! python3
#Automated testing program (PALMS, eConnect)
#by Charles Han

import os, glob, openpyxl, logging, datetime, time, random#, csv, webbrowser,
#import pandas as pd
#from PIL import Image
#from io import BytesIO
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

testFolder = os.getcwd()+ '\\Tests\\Run' #set the run folder path
screenshotFolder = os.getcwd()+'\\Screenshots\\' #set screenshot folder path
uploadFolder = os.getcwd()+'\\For upload\\' #set test upload file folder
logFolder = os.getcwd()+'\\Logs\\' #set results folder

today = str(datetime.date.today())
logging.basicConfig(filename = logFolder + 'log'+ today +'.txt', level=logging.DEBUG, format=' %(asctime)s - %(levelname)s - %(message)s')

#--- User selects browser to test in ---
print('Automated website testing program v0.1. by Charles Han')
print('Created using Python 3 with Selenium WebDriver')
print('Icon provided by icons8.com under Creative Commons Attribution-NoDerivs 3.0 Unported'+'\n')
print('Run tests using: 1.Chrome,  2.FireFox, 3.InternetExplorer ?')
browserSelect = input()
if browserSelect == "1": browser = webdriver.Chrome() 
if browserSelect == "2": browser = webdriver.Firefox() 
if browserSelect == "3": browser = webdriver.Ie() #IE not working currently

#-- default webdriver wait is 120sec/2min ---
#wait = WebDriverWait(browser,120)

stopFlag = False # boolean stop flag

#--- Recursive open all xlsx files in the run folder ---
for filename in glob.iglob(testFolder + '**/*.xlsx', recursive=True):
    print(filename)

#--- load xlsx workbook ---
    wb = openpyxl.load_workbook(filename,data_only=True)
    
#--- loop through the sheets ---
    logging.debug(wb.sheetnames)
    for s in wb.sheetnames:
        logging.debug(s)
        sheet = wb[s]
        
# TODO: check if formula cell values work as intended
#--- Iterate through the sheets and cells ---
        for cells in sheet.iter_rows():
            for c in cells:

                if c.value == 'start': # skips the rest of the sheet
                    stopFlag = False
                    print('start tests')
                
                if not(stopFlag):
                    #try:
                        #if c.value == 'setwait':
                        #    wait = WebDriverWait(browser,(c.offset(column=2).value))
            
                        # input: send keys to element
                        if c.value == 'input':       
                            htmlid = (c.offset(column=1).value) #get the value of the cell in the next column to the right
                            logging.debug(htmlid)
                            element = WebDriverWait(browser,120).until(EC.presence_of_element_located(eval(htmlid)))
                            print('Found <%s> element with HTML tag!' % (htmlid))
                            try:
                                element.clear()
                            except:
                                pass
                            value = str(c.offset(column=2).value) #get the value of the cell two columns to the right
                            element.send_keys(value)
                            print('Keys sent <%s>' % (value))

                        # click: simulate mouse click on element                          
                        if c.value == 'click':
                            htmlid =(c.offset(column=1).value) #get the value of the cell in the next column to the right
                            element = WebDriverWait(browser,120).until(EC.presence_of_element_located(eval(htmlid)))
                            print('Found <%s> element with HTML tag!' % (htmlid))
                            element.click()
                            print('element clicked')

                        # upload: upload file
                        if c.value == 'upload':       
                            htmlid = (c.offset(column=1).value) #get the value of the cell in the next column to the right
                            logging.debug(htmlid)
                            element = WebDriverWait(browser,120).until(EC.presence_of_element_located(eval(htmlid)))
                            print('Found <%s> element with HTML tag!' % (htmlid))
                            filename = str(c.offset(column=2).value) #get the value of the cell two columns to the right
                            element.send_keys(uploadFolder+filename)
                            print('File uploaded <%s>' % (filename))
                           
                        # screenshot: takes a screenshot and save it to screenshot folder
                        if c.value == 'screenshot':
                            fileName = str(c.offset(column=2).value) #get the value of the cell two columns to the right
                            browser.save_screenshot(screenshotFolder + fileName + '.png')
                            print('screenshot taken')
                            
                        # go: goes to the URL specified                   
                        if c.value == 'go':
                            url = str(c.offset(column=1).value) ##get the value of the cell to the right
                            browser.get(url)
                            print('going to <%s>' % (url))
                            
                        # TODO: submit form by id
                        if c.value == 'submit':
                            htmlid =(c.offset(column=1).value) #get the value of the cell in the next column to the right
                            element = WebDriverWait(browser,120).until(EC.presence_of_element_located(eval(htmlid)))
                            print('Found <%s> element with HTML tag!' % (htmlid))
                            element.submit()
                            print('element submitted')
                            
                        # wait: pauses the progam by number of seconds specified
                        if c.value == 'wait': 
                            waitTime = (c.offset(column=2).value)
                            if waitTime == 'user':
                                print('waiting for user input')
                                input()
                            else:
                                try:
                                    time.sleep(waitTime)
                                except:
                                    print ('wait time not valid')
                                    continue       

                        # stop: skips tests between stop/start
                        if c.value == 'stop': # skips the rest of the sheet
                            stopFlag = True 
                            print('stop tests')

                        # waitforuser: check title
                   #     if c.value == 'stop': # skips the rest of the sheet 
                            
                        # TODO: check title
                        if c.value == 'checktitle': # check screen title
                            print('checking webpage title')

                        # TODO: check URL
                        if c.value == 'checkURL': # check current URL
                            print('checking URL')

                        if c.value == 'dropdown': # dropdown selection. default is random select
                            
                            htmlid =(c.offset(column=1).value) #get the value of the cell in the next column to the right
                            element = WebDriverWait(browser,120).until(EC.presence_of_element_located(eval(htmlid)))
                            print('Found <%s> element with HTML tag!' % (htmlid))
                            #element.click()
                            index = (c.offset(column=2).value)
                            if index.isdigit() and index < len(element.options): # if index is specified
                                try:    #try to select dropdown item by index
                                    print('selecting dropdown by index')
                                    element.select_by_index(index)
                                except:
                                    continue
                            else:
                                try:    #try random select
                                    element.select_by_index(random.randint(0, len(element.options)))
                                    print('element clicked')
                                except:
                                    continue
                            # month = Select(driver.find_element_by_css_selector('#month'))
                            # month.select_by_index(randint(0, len(month.options)))
                            
                        # checkdisplayed: check element displayed
                        if c.value == 'checkdisplayed': # check if an element is no visible on the screen
                            htmlid =(c.offset(column=1).value) #get the value of the cell in the next column to the right
                            try:
                                element = WebDriverWait(browser,120).until(EC.presence_of_element_located(eval(htmlid)))
                            except NoSuchElementException:
                                print('<%s> element not found/displayed within the time limit' % (htmlid))
                            if element.is_displayed():
                                print('<%s> found and displayed' % (htmlid))
                            else:
                                print('<%s> element found but not displayed' % (htmlid))
                        
                    #except:
                    #    print("step %s skipped due to timeout!" % (c.value))
                    #    pass

## close: closes browser
##                if c.value == 'close': #Closes the web browser
##                    print('close browser')
##                    browser.close

# TODO: Catch exceptions/errors
