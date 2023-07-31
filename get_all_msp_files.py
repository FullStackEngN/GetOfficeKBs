import logging
import os
import pathlib
import urllib.request

import wget
from lxml import html
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

from extract_msp import extract_msp_from_cab
from get_msp_download_link import get_download_link
from msp_file import msp_file_item
from my_logger_object import create_logger_object


def check_kb_in_excluded_list(current_kb_number, kb_list_excluded):
    if current_kb_number in kb_list_excluded:
        logger = logging.getLogger("download_kb")
        logger.info(str.format(">>> @@@ ---{} is excluded by user", current_kb_number))
        return True
    else:
        return False


def get_kb_links_for_expected_kb(
    current_kb_number,
    kb_list_excluded,
    kb_list_expected,
    browser,
    target_download_folder,
):
    links = []

    if check_kb_in_excluded_list(current_kb_number, kb_list_excluded):
        return links

    if current_kb_number in kb_list_expected:
        links = get_download_link(browser, current_kb_number, target_download_folder)
    else:
        logger.info(current_kb_number + " is not in the expected kb list, ignore it.")

    return links


def get_kb_links(kb_number, kb_list_excluded, browser, target_download_folder):
    links = []

    if check_kb_in_excluded_list(kb_number, kb_list_excluded):
        return links

    links = get_download_link(browser, kb_number, target_download_folder)

    return links


current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + "\\"

FILENAME = current_script_folder + "log_" + os.path.basename(__file__) + ".log"
logger = create_logger_object(FILENAME)

logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

url = "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2013#list-of-all-msp-files"
target_download_folder = current_script_folder + "Folder_Office2013_KBs" + os.sep

url = "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016#list-of-all-msp-files"
target_download_folder = current_script_folder + "Folder_Office2016_KBs" + os.sep

logger.info("The download URL is " + url)
logger.info("The target download folder is " + target_download_folder)

if not os.path.exists(target_download_folder):
    os.makedirs(target_download_folder)
    logger.info("The target download folder doesn't exist, create it.")


page = urllib.request.urlopen(url)
content = page.read().decode("utf-8")

"""
# save the html page to local text file for test
logger.debug(type(content))
f = open("log_content.log","w")
f.write(content)
f.close()
"""

doc = html.document_fromstring(content)

"""
# save the html content to local text file for test
logger.debug(type(doc))

# logger.debug(etree.tostring(doc))

f = open("log_doc.log", "w")
f.write(etree.tostring(doc).decode())
f.close()
"""

# //*[@id="main"]/div[2]/table
# /html/body/div[3]/div/section/div/div[1]/main/div[2]/table
msp_table = doc.xpath('//*[@id="main"]/div[3]/table[2]')

"""
#logger.debug (msp_table)

f = open("log_table.log", "w")
f.write(etree.tostring(msp_table[0]).decode())
f.close()


table_string = etree.tostring(msp_table[0]).decode()
logger.debug(type(table_string))

# logger.debug(table_string)
"""

th_list = doc.xpath('//*[@id="main"]/div[3]/table[2]/thead/tr/th')
logger.debug("Column: " + str(len(th_list)))

td_list = doc.xpath('//*[@id="main"]/div[3]/table[2]/tbody/tr')
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

    mspFile = msp_file_item(
        filename,
        product,
        non_security_release_date,
        non_security_KB,
        security_release_date,
        security_KB,
    )

    msp_file_list.append(mspFile)

msg = ""
new_line = "\n"

f = open(current_script_folder + "output_msp_file_list.log", "w")

for element in msp_file_list:
    f.write(element.tostring())
    f.write("\n")
f.close()

logger.info("List length: " + str(len(msp_file_list)))

# options = webdriver.FirefoxOptions()
# browser = webdriver.Firefox(options=options)
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
browser = webdriver.Edge(options=options)

kb_list_excluded = []
try:
    f = open(current_script_folder + "kb_list_excluded.txt", "r")

    for line in f:
        kb_list_excluded.append(line.strip().upper())

    logging.info("Read excluded_kb_list file, length is: " + str(len(kb_list_excluded)))
except:
    logging.info("No excluded_kb_list file, so don't need exclude KBs.")
finally:
    f.close()

kb_list_expected = []
try:
    f = open(current_script_folder + "kb_list_expected.txt", "r")

    for line in f:
        kb_list_expected.append(line.strip().upper())

    logging.info("Read expected_kb_list file, length is: " + str(len(kb_list_expected)))
except:
    logging.info("No expected_kb_list file, so download all KBs.")
finally:
    f.close()

kb_list_download = []
kb_list_ignored = []
download_links = []

if len(kb_list_expected) > 0:
    for item in msp_file_list:
        logger.info(
            "******Start component: "
            + item.filename
            + ", non-security update: KB"
            + item.non_security_KB
            + "; security update: KB"
            + item.security_KB
        )

        if item.security_kb_greater_than_non_security_kb():
            current_kb_number = "KB" + item.security_KB
            kb_list_ignored.append("KB" + item.non_security_KB)
            logger.info("******###Ignore non security kb: " + item.tostring())
        else:
            current_kb_number = "KB" + item.non_security_KB
            kb_list_ignored.append("KB" + item.security_KB)
            logger.info("******###Ignore security kb: " + item.tostring())

        kb_list_download.append(current_kb_number + "," + item.filename)

        links = get_kb_links_for_expected_kb(
            current_kb_number,
            kb_list_excluded,
            kb_list_expected,
            browser,
            target_download_folder,
        )
        download_links += links

        logger.info(">>>Done get links for " + current_kb_number)

        logger.info(
            "######Finish component: "
            + item.filename
            + ", non-security update: KB"
            + item.non_security_KB
            + "; security update: KB"
            + item.security_KB
        )
else:
    for item in msp_file_list:
        logger.info(
            "***Start component: "
            + item.filename
            + ", non-security update: KB"
            + item.non_security_KB
            + "; security update: KB"
            + item.security_KB
        )

        if item.security_greater_than_non_security:
            logger.info(
                ">>>Only need download security KB"
                + item.security_KB
                + ", release date "
                + item.security_release_date
                + ", Ignore non-security KB"
                + item.non_security_KB
            )
            current_kb_number = "KB" + item.security_KB
            kb_list_ignored.append("KB" + item.non_security_KB)
        else:
            logger.info(
                ">>><<<Only need download non-security KB"
                + item.non_security_KB
                + ", release date "
                + item.non_security_release_date
                + ", Ignore security KB"
                + item.security_KB
            )
            current_kb_number = "KB" + item.non_security_KB
            kb_list_ignored.append("KB" + item.security_KB)

        kb_list_download.append(current_kb_number + "," + item.filename)

        links = get_kb_links(
            current_kb_number, kb_list_excluded, browser, target_download_folder
        )
        download_links += links

        logger.info(">>>Finish get links for " + current_kb_number)

        logger.info(
            "###Finish component: "
            + item.filename
            + ", non-security update: KB"
            + item.non_security_KB
            + "; security update: KB"
            + item.security_KB
        )

f = open(current_script_folder + "output_download_kb_list.log", "w")
for element in kb_list_download:
    f.write(element)
    f.write(new_line)
f.close()

f = open(current_script_folder + "output_ignored_kb_list.log", "w")
for element in kb_list_ignored:
    f.write(element)
    f.write(new_line)
f.close()

f = open(current_script_folder + "output_kb_links_list.log", "w")
for element in download_links:
    f.write(element[0] + ":" + element[1] + ":" + element[2])
    f.write(new_line)
f.close()

browser.quit()
logger.info("##############################")
logger.info("Start download progress.")

target_download_folder_x64 = target_download_folder + "x64"
logger.info("The target download folder for 64bit KBs is " + target_download_folder_x64)

if not os.path.exists(target_download_folder_x64):
    os.makedirs(target_download_folder_x64)
    logger.info("The target download folder for 64bit KBs doesn't exist, create it.")

target_download_folder_x86 = target_download_folder + "x86"
logger.info("The target download folder for 32bit KBs is " + target_download_folder_x86)

if not os.path.exists(target_download_folder_x86):
    os.makedirs(target_download_folder_x86)
    logger.info("The target download folder for 32bit KBs doesn't exist, create it.")

tmp_file_name = ""
for item in download_links:
    if item[0] == "x64":
        tmp_file_name = wget.detect_filename(url=item[2])
        wget.download(
            url=item[2],
            out=target_download_folder_x64 + os.sep + item[1] + "_" + tmp_file_name,
        )

    if item[0] == "x86":
        tmp_file_name = wget.detect_filename(url=item[2])
        wget.download(
            url=item[2],
            out=target_download_folder_x86 + os.sep + item[1] + "_" + tmp_file_name,
        )

logger.info("Finish download progress.")
logger.info("##############################")

source_folder = target_download_folder_x86
target_folder = target_download_folder_x86 + "_" + "msp"
temp_folder = target_download_folder_x86 + "_" + "temp"

extract_msp_from_cab(source_folder, target_folder, temp_folder)

source_folder = target_download_folder_x64
target_folder = target_download_folder_x64 + "_" + "msp"
temp_folder = target_download_folder_x64 + "_" + "temp"

extract_msp_from_cab(source_folder, target_folder, temp_folder)

logger.info("Done...")
