#Imports
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import xlsxwriter
import pyperclip
import time
import sys
import os

#Filepaths
file_path = 'vins.xlsx'
chrome_path = 'chromedriver.exe'
website_path = 'https://vpic.nhtsa.dot.gov/decoder'

#Get list of vins from clipboard
raw_list = str(pyperclip.paste()).split("\r\n")
#List compression to remove blanks
raw_list = [i for i in raw_list if i != '']
#Blank list for year, make, model
vin_list = []*len(raw_list)

#Get into website
options = webdriver.ChromeOptions()
#Change value to True to hide browser, False to show browser
options.headless = False
driver = webdriver.Chrome(chrome_path, options=options)
driver.get(website_path)

#loop through
def vin_func():
    for item in raw_list:
        #Enter vin number in search box
        vin_text = driver.find_element_by_id('VIN')
        vin_text.send_keys(item)
        #Enter button
        submit_button = driver.find_element_by_id('btnSubmit')
        submit_button.click()
        #Grab data
        year_grab_element = driver.find_element_by_id('decodedModelYear')
        make_grab_element = driver.find_element_by_id('decodedMake')
        model_grab_element = driver.find_element_by_id('decodedModel')
        #Convert to a string
        year_grab = year_grab_element.text
        make_grab = make_grab_element.text
        model_grab = model_grab_element.text
        #Function to rename bad vins
        def blank_func(x):
            if len(x) == 0:
                return 'Error'
            else:
                return x
        #Run results through blank function
        year_result = blank_func(year_grab)
        make_result = blank_func(make_grab)
        model_result = blank_func(model_grab)
        #Return full year/make/model
        vin_result = year_result+' '+make_result+' '+model_result
        #Add to list of results
        vin_list.append(vin_result)
        #Fix for StaleElementReferenceException
        vin_text = driver.find_element_by_id('VIN')
        #Clear search field
        vin_text.clear()
#Try/except to close script and close driver. Clean exit.
try:
    vin_func()
except Exception as e:
    driver.quit()
    print(e)
    sys.exit()


#Exit the driver
driver.quit()
#Format Excel book, add data
workbook = xlsxwriter.Workbook(file_path)
worksheet=workbook.add_worksheet()
worksheet.write(0,0,"VIN")
worksheet.write(0,1,"Vehicle")
def format_func(x,y):
    p=1
    for item in (x):
        col = y
        row = p
        worksheet.write(row,col,item)
        p=p+1
format_func(raw_list,0)
format_func(vin_list,1)
workbook.close()
#Open workbook
os.startfile(file_path)
