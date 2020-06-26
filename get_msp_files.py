import datetime
import logging
import pathlib
import time
import urllib.request
from download_kb_file import download_file
from msp_file import MspFile
from lxml import html
from lxml import etree
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait


current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + "\\"

FORMAT = '%(asctime)s %(levelname)s %(message)s'
FILENAME = current_script_folder + "log.txt"

logging.basicConfig(filename=FILENAME, format=FORMAT)
logger = logging.getLogger('download_kb')

current_date_time = datetime.datetime.now()
logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

url = 'https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2013#list-of-all-msp-files'
logger.info("The download URL is " + url)

#target_download_folder = r"C:\Temp\Office2013_KBs\\"
target_download_folder = current_script_folder + "Office2013_KBs\\"

# url = 'https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016#list-of-all-msp-files'
#target_download_folder = r"C:\Temp\Office2016_KBs\\"
#target_download_folder = current_script_folder + "Office2016_KBs"

logger.info("The target download folder is " + target_download_folder)

page = urllib.request.urlopen(url)
content = page.read().decode('utf-8')

'''
# save the html page to local text file for test
logger.debug(type(content))
f = open("content.txt","w")
f.write(content)
f.close()
'''

doc = html.document_fromstring(content)

'''
# save the html content to local text file for test
logger.debug(type(doc))

# logger.debug(etree.tostring(doc))

f = open("doc.txt", "w")
f.write(etree.tostring(doc).decode())
f.close()
'''

# //*[@id="main"]/div[2]/table
# /html/body/div[3]/div/section/div/div[1]/main/div[2]/table
msp_table = doc.xpath('//*[@id="main"]/table[2]')

'''
#logger.debug (msp_table)

f = open("table.txt", "w")
f.write(etree.tostring(msp_table[0]).decode())
f.close()


table_string = etree.tostring(msp_table[0]).decode()
logger.debug(type(table_string))

# logger.debug(table_string)
'''

th_list = doc.xpath('//*[@id="main"]/table[2]/thead/tr/th')
#logger.debug("Column: " + str(len(th_list)))

td_list = doc.xpath('//*[@id="main"]/table[2]/tbody/tr')
logger.info("Rows: " + str(len(td_list)))

msp_file_list = []

for x in td_list:
    filename = x[0].text_content().strip()
    product = x[1].text_content().strip()
    non_security_release_date = x[2].text_content().strip()
    non_security_KB = x[3].text_content().strip()
    security_release_date = x[4].text_content().strip()
    security_KB = x[5].text_content().strip()
    security_greater_than_non_security = False

    mspFile = MspFile(filename, product, non_security_release_date,
                      non_security_KB, security_release_date, security_KB)

    # logger.info(mspFile.tostring())

    msp_file_list.append(mspFile)

msg = ""
new_line = '\n'

f = open(current_script_folder + "msp_file_list.txt", "w")

for element in msp_file_list:
    f.write(element.tostring())
    f.write('\n')
f.close()

logger.info("List length: " + str(len(msp_file_list)))

firefoxProfile = webdriver.FirefoxProfile()
firefoxProfile.set_preference("browser.download.folderList", 2)
firefoxProfile.set_preference(
    "browser.download.manager.showWhenStarting", False)
firefoxProfile.set_preference(
    "browser.download.dir", target_download_folder)
firefoxProfile.set_preference(
    "browser.helperApps.neverAsk.openFile", "application/octet-stream")
firefoxProfile.set_preference(
    "browser.helperApps.neverAsk.saveToDisk", "application/octet-stream")

browser = webdriver.Firefox(firefox_profile=firefoxProfile)

exclude_list = []
try:
    f = open(current_script_folder + "exclude_kb_list.txt", 'r')
    exclude_list = f.readlines()
except:
    logging.info("No exclude_kb_list file, so don't need exclude KBs.")
finally:
    f.close()

kb_list = []
ignore_kb_list = []

count = 0

for item in msp_file_list:

    count += 1

    if(item.security_greater_than_non_security):

        current_kb_number = "KB" + item.security_KB + "\n"

        if current_kb_number in exclude_list:
            logger.info(str.format(
                ">>[{0}]>> @@@Exclude the {1}", count, current_kb_number))
            continue

        kb_list.append("KB" + item.security_KB)
        ignore_kb_list.append("KB" + item.non_security_KB)

        msg = str.format(">>{0}>> Only need security KB{1}, {2}",
                         count, item.security_KB, item.product)
        logger.info(msg)

        msg = str.format(">>{0}>> *** start KB{1}, {2}", count,
                         item.security_KB, item.product)
        logger.info(msg)

        msg = str.format(">>{0}>> ... Ingore KB{1}, {2}",
                         count, item.non_security_KB, item.product)
        logger.info(msg)

        download_file(browser, item.security_KB,
                      item.product, target_download_folder)

        msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                         count, item.security_KB, item.product)
        logger.info(msg)

        time.sleep(3 * 60)
    else:
        if(item.non_security_KB != "Not applicable"):
            current_kb_number = "KB" + item.non_security_KB + "\n"

            if current_kb_number in exclude_list:
                logger.info(str.format(
                    ">>{0}>> @@@Exclude the {1}", count, current_kb_number))
                continue

            kb_list.append("KB" + item.non_security_KB)

            msg = str.format(">>{0}>> *** start KB{1}, {2}", count,
                             item.non_security_KB, item.product)
            logger.info(msg)

            download_file(browser, item.non_security_KB,
                          item.product, target_download_folder)

            msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                             count, item.non_security_KB, item.product)
            logger.info(msg)

            time.sleep(2 * 60)

        if(item.security_KB != "Not applicable"):

            current_kb_number = "KB" + item.security_KB + "\n"

            if current_kb_number in exclude_list:
                logger.info(str.format(">>{0}>> @@@Exclude the {1}",
                                       count, current_kb_number))
                continue

            kb_list.append("KB" + item.security_KB)

            msg = str.format(">>{0}>> *** start KB{1}, {2}", count,
                             item.security_KB, item.product)
            logger.info(msg)

            download_file(browser, item.security_KB,
                          item.product, target_download_folder)

            msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                             count, item.security_KB, item.product)
            logger.info(msg)

            time.sleep(2 * 60)

f = open(current_script_folder + "kb_list.txt", "w")
for element in kb_list:
    f.write(element)
    f.write('\n')
f.close()

f = open(current_script_folder + "ignore_kb_list.txt", "w")
for element in ignore_kb_list:
    f.write(element)
    f.write('\n')
f.close()

logger.info("##############################")
browser.quit()

current_date_time = datetime.datetime.now()
logger.info("Done: " + str(current_date_time))
