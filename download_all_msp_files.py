# This script automates the process of downloading all Office 2016 MSP update files.
# It scrapes the official Office update documentation for a list of MSP files and their KB numbers,
# determines which KBs to download or ignore, retrieves download links using Selenium,
# downloads the CAB files into architecture-specific folders, and extracts MSP files from the CABs.

import os
import pathlib
import urllib.request

from lxml import html
from selenium import webdriver

from lib_common_utils import ensure_folder_exists, process_download_links, read_kb_list
from lib_extract_msp import extract_msp_from_cab
from lib_get_msp_download_link import get_download_link, get_kb_links
from lib_logger_object import create_logger_object
from lib_msp_file import msp_file_item


def main():

    current_script_folder = str(pathlib.Path(__file__).parent.absolute())

    log_file_path = os.path.join(current_script_folder, f"log_{os.path.basename(__file__)}.log")

    logger = create_logger_object(log_file_path)

    logger.info("The script starts running.")
    logger.info(f"The script folder is {current_script_folder}")

    url = "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2013#list-of-all-msp-files"
    target_download_folder = os.path.join(current_script_folder, "Folder_Office2013_KBs")

    url = "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016#list-of-all-msp-files"
    target_download_folder = os.path.join(current_script_folder, "Folder_Office2016_KBs")

    logger.info(f"The download URL is {url}")
    logger.info(f"The target download folder is {target_download_folder}")

    ensure_folder_exists(target_download_folder, logger)

    page = urllib.request.urlopen(url)
    content = page.read().decode("utf-8")
    doc = html.document_fromstring(content)

    th_list = doc.xpath("//html/body//table[2]/thead/tr/th")
    logger.debug(f"Column: {len(th_list)}")
    td_list = doc.xpath("//html/body//table[2]/tbody/tr")
    logger.info(f"Rows: {len(td_list)}")

    msp_file_list = []
    for x in td_list:
        filename = x[0].text_content().strip()
        product = x[1].text_content().strip()
        non_security_release_date = x[2].text_content().strip()
        non_security_KB = x[3].text_content().strip()
        security_release_date = x[4].text_content().strip()
        security_KB = x[5].text_content().strip()

        mspFile = msp_file_item(
            filename,
            product,
            non_security_release_date,
            non_security_KB,
            security_release_date,
            security_KB,
        )
        msp_file_list.append(mspFile)

    with open(os.path.join(current_script_folder, "output_msp_file_list.log"), "w") as f:
        for element in msp_file_list:
            print(element.tostring(), file=f)

    logger.info("List length: " + str(len(msp_file_list)))

    options = webdriver.EdgeOptions()
    options.add_argument("--start-maximized")
    browser = webdriver.Edge(options=options)

    kb_list_excluded = read_kb_list(current_script_folder, "input_kb_list_excluded.txt", logger)

    kb_list_download = []
    kb_list_ignored = []
    download_links = []

    for msp_file_info in msp_file_list:
        logger.info(
            "***Start component: "
            + msp_file_info.filename
            + ", non-security update: KB"
            + msp_file_info.non_security_KB
            + "; security update: KB"
            + msp_file_info.security_KB
        )

        if msp_file_info.security_greater_than_non_security:
            logger.info("******###Ignore non security kb: " + msp_file_info.tostring())

            current_kb_number = "KB" + msp_file_info.security_KB
            kb_list_ignored.append("KB" + msp_file_info.non_security_KB)
        else:
            logger.info("******@@@Ignore security kb: " + msp_file_info.tostring())

            current_kb_number = "KB" + msp_file_info.non_security_KB
            kb_list_ignored.append("KB" + msp_file_info.security_KB)

        kb_list_download.append(current_kb_number + "," + msp_file_info.filename)

        links = get_kb_links(
            browser,
            current_kb_number,
            kb_list_excluded,
            target_download_folder,
            logger,
        )
        download_links += links

        logger.info(">>>Finish get links for " + current_kb_number)

        logger.info(
            "###Finish component: "
            + msp_file_info.filename
            + ", non-security update: KB"
            + msp_file_info.non_security_KB
            + "; security update: KB"
            + msp_file_info.security_KB
        )

    with open(current_script_folder + "tmp_download_kb_list.log", "w") as f:
        for element in kb_list_download:
            print(element, file=f)

    with open(current_script_folder + "tmp_ignored_kb_list.log", "w") as f:
        for element in kb_list_ignored:
            print(element, file=f)

    with open(current_script_folder + "tmp_kb_links_list.log", "w") as f:
        for element in download_links:
            print(element[0] + ":" + element[1] + ":" + element[2], file=f)

    browser.quit()
    logger.info("##############################")
    logger.info("Start download progress.")

    target_download_folder_x64 = os.path.join(target_download_folder, "x64")
    target_download_folder_x86 = os.path.join(target_download_folder, "x86")
    ensure_folder_exists(target_download_folder_x64, logger)
    ensure_folder_exists(target_download_folder_x86, logger)

    process_download_links(download_links, target_download_folder_x64, target_download_folder_x86, logger)

    logger.info("Finish download progress.")
    logger.info("##############################")

    # Extract MSP files from CABs for both architectures
    for arch_folder in [target_download_folder_x86, target_download_folder_x64]:
        source_folder = arch_folder
        target_folder = f"{arch_folder}_msp"
        temp_folder = f"{arch_folder}_temp"
        extract_msp_from_cab(source_folder, target_folder, temp_folder)

    logger.info("Done...")


if __name__ == "__main__":
    main()
