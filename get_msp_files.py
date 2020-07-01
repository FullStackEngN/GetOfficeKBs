import logging
import pathlib
import os
import urllib.request
import wget
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

logging.basicConfig(format=FORMAT, datefmt='%a, %d %b %Y %H:%M:%S',
                    filename=FILENAME, filemode='w', level=logging.DEBUG)

logger = logging.getLogger('download_kb')

console = logging.StreamHandler()
console.setLevel(logging.WARNING)
logger.addHandler(console)

logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

url = 'https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2013#list-of-all-msp-files'
target_download_folder = current_script_folder + "Office2013_KBs\\"


# url = 'https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016#list-of-all-msp-files'
#target_download_folder = current_script_folder + "Office2016_KBs"

logger.info("The download URL is " + url)
logger.info("The target download folder is " + target_download_folder)

if not os.path.exists(target_download_folder):
    os.makedirs(target_download_folder)
    logger.info("The target download folder doesn't exist, create it.")


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

    for line in f:
        exclude_list.append(line.strip().upper())

except:
    logging.info("No exclude_kb_list file, so don't need exclude KBs.")
finally:
    f.close()

expected_kb_list = []
try:
    f = open(current_script_folder + "expected_kb_list.txt", 'r')

    for line in f:
        expected_kb_list.append(line.strip().upper())

    logging.info("Read expected_kb_list file, length is :" +
                 len(expected_kb_list))
except:
    logging.info("No expected_kb_list file, so download all KBs.")
finally:
    f.close()

kb_list = []
ignore_kb_list = []

count = 0

download_links = []

if len(expected_kb_list) > 0:
    for item in msp_file_list:

        current_kb_number = "KB" + item.security_KB
        if current_kb_number in expected_kb_list:
            logger.info("Find the expected " + current_kb_number)

            links = download_file(browser, item.security_KB,
                                  item.product, target_download_folder)

            download_links += links

            msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                             count, item.security_KB, item.product)
            logger.info(msg)

        else:
            logger.info(current_kb_number +
                        " is not in the expected kb list, ignore it.")

        current_kb_number = "KB" + item.non_security_KB
        if current_kb_number in expected_kb_list:
            logger.info("Find the expected " + current_kb_number)

            links = download_file(browser, item.non_security_KB,
                                  item.product, target_download_folder)

            download_links += links

            msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                             count, item.security_KB, item.product)
            logger.info(msg)

        else:
            logger.info(current_kb_number +
                        " is not in the expected kb list, ignore it.")


else:
    for item in msp_file_list:

        count += 1

        current_kb_number = "KB" + item.security_KB

        if(item.security_greater_than_non_security):

            current_kb_number = "KB" + item.security_KB

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

            links = download_file(browser, item.security_KB,
                                  item.product, target_download_folder)
            download_links += links

            msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                             count, item.security_KB, item.product)
            logger.info(msg)

        else:
            if(item.non_security_KB != "Not applicable"):

                current_kb_number = "KB" + item.non_security_KB

                if current_kb_number in exclude_list:
                    logger.info(str.format(
                        ">>{0}>> @@@Exclude the {1}", count, current_kb_number))
                    continue

                kb_list.append("KB" + item.non_security_KB)

                msg = str.format(">>{0}>> *** start KB{1}, {2}", count,
                                 item.non_security_KB, item.product)
                logger.info(msg)

                links = download_file(browser, item.non_security_KB,
                                      item.product, target_download_folder)
                download_links += links

                msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                                 count, item.non_security_KB, item.product)
                logger.info(msg)

            if(item.security_KB != "Not applicable"):

                current_kb_number = "KB" + item.security_KB

                if current_kb_number in exclude_list:
                    logger.info(str.format(">>{0}>> @@@Exclude the {1}",
                                           count, current_kb_number))
                    continue

                kb_list.append("KB" + item.security_KB)

                msg = str.format(">>{0}>> *** start KB{1}, {2}", count,
                                 item.security_KB, item.product)
                logger.info(msg)

                links = download_file(browser, item.security_KB,
                                      item.product, target_download_folder)
                download_links += links

                msg = str.format(">>{0}>> ### finish KB{1}, {2}",
                                 count, item.security_KB, item.product)
                logger.info(msg)

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

f = open(current_script_folder + "links_kb_list.txt", "w")
for element in download_links:
    f.write(element[0] + ":KB" + element[1] + ":" + element[2])
    f.write('\n')
f.close()

browser.quit()

logger.info("Start download progress.")

target_download_folder_x64 = current_script_folder + "Office2013_KBs\\x64\\"
logger.info("The target download folder for 64bit KBs is " +
            target_download_folder_x64)

if not os.path.exists(target_download_folder_x64):
    os.makedirs(target_download_folder_x64)
    logger.info(
        "The target download folder for 64bit KBs doesn't exist, create it.")

target_download_folder_x86 = current_script_folder + "Office2013_KBs\\x86\\"
logger.info("The target download folder for 32bit KBs is " +
            target_download_folder_x86)

if not os.path.exists(target_download_folder_x86):
    os.makedirs(target_download_folder_x86)
    logger.info(
        "The target download folder for 32bit KBs doesn't exist, create it.")

for item in download_links:
    if item[0] == "x64":
        wget.download(url=item[2], out=target_download_folder_x64)
        
    if item[0] == "x86":
        wget.download(url=item[2], out=target_download_folder_x86)


logger.info("##############################")


logger.info("Done...")
