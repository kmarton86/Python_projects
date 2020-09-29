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
loc = ("C:/IT Workshop/Python/sciFiWriters2.xls") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0)
#get the raws number
print("number of raws:" + str(sheet.nrows)) #number of raws
countRows = sheet.nrows

#*************************************************************************
#create lists
writerList = [] # 1 dimensional list for the writers
novelsList = [] # 1 dimensional list for the novels of one writer at time. Tha data will be changed at each iteration.
listOfNovelsList = [] #create a list for each novels. It will be a nested list, like this: [[1,2],[3,4]].

for i in range(countRows): 
    # In each iteration, add an empty list to the main list
    # the number of iteration equals the number of raws
    listOfNovelsList.append([])
#print('listOfNovelsList:')
print(listOfNovelsList)

#fill list with data from xls
count = 0
print("*** script fills dataList with data from xls ***")
while count < countRows:
    writer = str(sheet.cell_value(count, 0)) 
    novels = str(sheet.cell_value(count, 1)) 
    novelsList = novels.split(',')
    print(str(count) + ".:" + writer + " :")
    for i in range(len(novelsList)): 
        print (novelsList[i])
        listOfNovelsList[count].append(novelsList[i]) # listOfNovelsList's actual sublist (count) will be filled with the elements of novelsList
    writerList.append(writer) 
    print("novels number: " + str(len(novelsList)))  # number of elements of the novelsList 
    count += 1

# printing the writerList using loop 
print("*** script loops through dataList and prints elements ***")
for i in range(len(writerList)): 
    print (writerList[i]) 
#print (writerList[0]) #print the first element of writerList

# printing the listOfNovelsList using loop 
print(listOfNovelsList)
for i in range(len(listOfNovelsList)): 
    print (listOfNovelsList[i]) 

time.sleep(5)   # 5 sec delay

#************************************************************************
#automate browser

talalatok = [] # create a talalatok list
#start browser

driver.get("http://saman.fszek.hu/WebPac/CorvinaWeb?action=simplesearchpage")

#automate website

#start loop in writerList
# in for loop the script loops through writerList and searches for each element
for i in range(len(writerList)): 
    writer = writerList[i]
    driver.find_element_by_id("text0").send_keys(writer) # inputs the keyword
    driver.find_element_by_id("searchform").submit()
    # in try the script try to get tha talalatok text. If there is no talatok text the except will handle it
    try:
        talalat = driver.find_element_by_xpath("//td[contains(@class, 'thead_2of5')]").text #talált dokumentum
        print(talalat)
        talalatok.append(writer + " : " + talalat + ", novels: ")
        #if there is talaltok search for the novels
        driver.find_element_by_id("text0").click()
        driver.find_element_by_id("text0").clear()
        #start loop in listOfNovelsList
        for j in range(len(listOfNovelsList)): 
            novel = listOfNovelsList[i][j]
            print ("actual novel:" + listOfNovelsList[i][j]) 
            driver.find_element_by_id("text0").clear()
            driver.find_element_by_id("text0").send_keys(writer + " " + novel)
            driver.find_element_by_id("searchform").submit()
    except:
        print("nincs találat: " + writer)
        talalatok.append(writer + " : nincs találat ")
    driver.find_element_by_id("text0").click()
    driver.find_element_by_id("text0").clear()

time.sleep(3)   # 3 sec delay


#************************************************************************
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