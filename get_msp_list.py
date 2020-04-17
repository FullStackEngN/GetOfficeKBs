import urllib.request
from lxml import html
from lxml import etree
from mspfile import MspFile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import NoSuchElementException
from download_kb import DownloadFile
import datetime

currentDT = datetime.datetime.now()
print("Start: " + str(currentDT))

url = 'https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016#list-of-all-msp-files'

page = urllib.request.urlopen(url)
content = page.read().decode('utf-8')

'''
print(type(content))
f = open("content.txt","w")
f.write(content)
f.close()
'''

doc = html.document_fromstring(content)

'''
print(type(doc))

print(etree.tostring(doc))

f = open("doc.txt","w")
f.write(etree.tostring(doc).decode())
f.close()
'''

# //*[@id="main"]/div[2]/table
# /html/body/div[3]/div/section/div/div[1]/main/div[2]/table
msp_table = doc.xpath('//*[@id="main"]/table[2]')

'''
print (msp_table)

f = open("table.txt","w")
f.write(etree.tostring(msp_table[0]).decode())
f.close()


table_string = etree.tostring(msp_table[0]).decode()
print(type(table_string))

# print(table_string)
'''
th_list = doc.xpath('//*[@id="main"]/table[2]/thead/tr/th')
print(len(th_list))

td_list = doc.xpath('//*[@id="main"]/table[2]/tbody/tr')
print(len(td_list))

msp_file_list = []

for x in td_list:
    # print(x.tag)

    filename = x[0].text_content().strip()
    product = x[1].text_content().strip()
    non_security_release_date = x[2].text_content().strip()
    non_security_KB = x[3].text_content().strip()
    security_release_date = x[4].text_content().strip()
    security_KB = x[5].text_content().strip()
    security_greater_than_non_security = False

    mspFile = MspFile(filename, product, non_security_release_date,
                      non_security_KB, security_release_date, security_KB)

    # print(mspFile.tostring())

    msp_file_list.append(mspFile)

print(len(msp_file_list))

firefoxProfile = webdriver.FirefoxProfile()
firefoxProfile.set_preference("browser.download.folderList", 2)
firefoxProfile.set_preference(
    "browser.download.manager.showWhenStarting", False)
firefoxProfile.set_preference(
    "browser.download.dir", r"C:\Temp\Office2016_KBs\\")
firefoxProfile.set_preference(
    "browser.helperApps.neverAsk.openFile", "application/octet-stream")
firefoxProfile.set_preference(
    "browser.helperApps.neverAsk.saveToDisk", "application/octet-stream")

browser = webdriver.Firefox(firefox_profile=firefoxProfile)

for item in msp_file_list:
    if(item.security_greater_than_non_security):
        print("*** Only need security KB" + item.security_KB + "," + item.product)
        DownloadFile(browser, item.security_KB, item.product)
    else:
        if(item.non_security_KB != "Not applicable"):
            print("*** start KB" + item.non_security_KB + "," + item.product)
            DownloadFile(browser, item.non_security_KB, item.product)
            
        if(item.security_KB != "Not applicable"):
            print("*** start KB" + item.security_KB + "," + item.product)
            
    print("##############################")
# browser.implicitly_wait(10)
# browser.quit()

currentDT = datetime.datetime.now()
print("Done: " + str(currentDT))