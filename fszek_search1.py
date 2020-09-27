#*************************************************************************
#Settings
import time
import xlrd # pip install xlrd -> for xls
#import random

#************************************************************************
#Set selenium driver
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
chromedriver = "C:/IT Workshop/chromedriver"
driver = webdriver.Chrome(chromedriver)
#maximize window
driver.set_window_size(1024, 600)
driver.maximize_window()

#*************************************************************************
#read input data from xls
loc = ("C:/IT Workshop/Python/sciFiWriters.xls") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
print(sheet.nrows) #number of raws
countRows = sheet.nrows

#fill list with data from xls
dataList = []
count = 0
print("*** script fills dataList with data from xls ***")
while count < countRows:
    cella = str(sheet.cell_value(count, 0)) 
    print(count)
    print(cella)
    dataList.append(cella)
    count += 1

# printing the list using loop 
print("*** script loops through dataList and prints elements ***")
for i in range(len(dataList)): 
    print (dataList[i]) 

print (dataList[0])
time.sleep(5)   # 5 sec delay

#************************************************************************
#start test

talalatok = [] # create a talalatok list
#start browser
driver.get("http://saman.fszek.hu/WebPac/CorvinaWeb?action=simplesearchpage")

#automate website
# in for loop the script loops through dataList and searches for each element
for i in range(len(dataList)): 
    keyword = dataList[i]
    driver.find_element_by_id("text0").send_keys(keyword) # inputs the keyword
    driver.find_element_by_id("searchform").submit()
    talalat = driver.find_element_by_xpath("//td[contains(@class, 'thead_2of5')]").text
    print(talalat)
    talalatok.append(keyword + " : " + talalat)
    driver.find_element_by_id("text0").click()
    driver.find_element_by_id("text0").clear()

time.sleep(3)   # 3 sec delay

# printing the list using loop 
print("*** script loops through talalatok list and prints elements ***")
#wb = xlrd.open_workbook(loc) 

for i in range(len(talalatok)): 
#    sheet = wb.sheet_by_index(0)
#    sheet.write(i, 1, talalatok[i]) 
    print (talalatok[i]) 

f = open("C:/IT Workshop/Python/scifiwriters.txt", "w")

with open('C:/IT Workshop/Python/scifiwriters.txt', 'w') as filehandle:
    for listitem in talalatok:
        filehandle.write('%s\n' % listitem)



#close test
#print('Do you want to quit?:')
#answer = input()
#print: answer