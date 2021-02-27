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
file_path = 'random_vins.xlsx'
chrome_path = 'chromedriver.exe'
website_path = 'https://vingenerator.org/'
#Blank lists
vin_list=[]
result_list=[]
#Get into website
options = webdriver.ChromeOptions()
#Change value to True to hide browser, False to show browser
options.headless = False
driver = webdriver.Chrome(chrome_path, options=options)
driver.get(website_path)

#loop through until list it reaches 20
def vin_func():
    while len(vin_list)<20:
        #Use get_attribute() to get value of class. Returns a string.
        vin_grab_element = driver.find_element_by_class_name("input").get_attribute('value')
        vehicle_grab_element = driver.find_element_by_class_name("description")
        #Convert to a string.
        vehicle_grab = vehicle_grab_element.text
        vehicle_grab = vehicle_grab.replace('VIN Description: ', '')
        #Add to lists.
        vin_list.append(vin_grab_element)
        result_list.append(vehicle_grab)
        #create a new VIN.
        generate_button = driver.find_element_by_class_name("button")
        generate_button.click()
#Try/except to close script and close driver. Clean exit.
try:
    vin_func()
    driver.quit()
except Exception as e:
    driver.quit()
    print(e)
    sys.exit()



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
format_func(vin_list,0)
format_func(result_list,1)
workbook.close()
#Open workbook
os.startfile(file_path)