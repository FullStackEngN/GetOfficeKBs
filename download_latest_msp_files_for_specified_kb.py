# Check for duplicate KB numbers, find the corresponding component for each KB number, and get the latest update for the component.

import os
import pathlib
import re
import urllib.request

from lxml import html
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from lib_common_utils import ensure_folder_exists, process_download_links, read_kb_list, save_screenshot_with_log
from lib_extract_msp import extract_msp_from_cab
from lib_get_msp_download_link import get_download_link
from lib_logger_object import create_logger_object
from lib_msp_file import msp_file_item

# Get the current script folder
current_script_folder = str(pathlib.Path(__file__).parent.absolute()) + os.sep

log_file_path = os.path.join(current_script_folder, f"log_{os.path.basename(__file__)}.log")
logger = create_logger_object(log_file_path)

logger.info("The script starts running.")
logger.info("The script folder is " + current_script_folder)

# Set the target folder for output files and screenshots
target_download_folder = os.path.join(current_script_folder, "Folder_Latest_KB_Numbers")
logger.info("The target folder is " + target_download_folder)
ensure_folder_exists(target_download_folder, logger)

# Create the screenshots folder if it does not exist
screenshots_folder = os.path.join(target_download_folder, "screenshots")
ensure_folder_exists(screenshots_folder, logger)

kb_component_file = os.path.join(current_script_folder, "output_msp_file_name_for_specified_kb.txt")

# Check if the file does not exist or exists but is empty, then execute method 1
if not os.path.exists(kb_component_file) or os.path.getsize(kb_component_file) == 0:

    logger.warning(f"{kb_component_file} does not exist or is empty.")

    kb_list_specified = read_kb_list(current_script_folder, "input_kb_list_specified.txt", logger)

    kb_msp_file_name_list = []
    kb_msp_file_name_error_list = []

    try:
        browser = webdriver.Edge()
        for kb in kb_list_specified:
            kb_str = kb.strip().upper()
            logger.info(f"KB number: {kb_str}")

            kb_number = re.sub("KB", "", kb_str)
            logger.info(f"KB number only: {kb_number}")

            kb_desc_url = f"https://support.microsoft.com/kb/{kb_number}"
            logger.info(f"KB description url: {kb_desc_url}")

            try:
                browser.get(kb_desc_url)
            except Exception as ex:
                logger.error(f"Failed to load page for {kb_str}: {ex}")
                kb_msp_file_name_error_list.append(f"{kb_str}, ")
                continue

            package_name = None
            try:
                package_name = WebDriverWait(browser, 30).until(
                    EC.visibility_of_element_located((By.XPATH, "//p[contains(text(),'-glb.exe')]"))
                )
            except Exception:
                try:
                    package_name = WebDriverWait(browser, 10).until(
                        EC.visibility_of_element_located((By.XPATH, "//span[contains(text(),'-glb.exe')]"))
                    )
                except Exception as ex:
                    save_screenshot_with_log(
                        browser,
                        os.path.join(screenshots_folder, f"{kb_str}_error.png"),
                        logger,
                        f"Exception when waiting/loading element for {kb_str}: {ex}",
                        exception=True,
                    )

                    kb_msp_file_name_error_list.append(f"{kb_str}, ")
                    continue

            if package_name is not None:
                package_name_str = package_name.text
                logger.info(f"Package name: {package_name_str}")

                # Extract component name before the first dash, minus last 4 chars
                first_dash_index = package_name_str.find("-")
                if first_dash_index > 0:
                    first_substring = package_name_str[:first_dash_index]
                    component_name = first_substring[:-4]
                else:
                    component_name = package_name_str

                kb_msp_file_name_list.append(f"{kb_str},{package_name_str},{component_name}")
            else:
                save_screenshot_with_log(
                    browser,
                    os.path.join(screenshots_folder, f"{kb_str}_error.png"),
                    logger,
                    f"Can't find component name for {kb_str}",
                    exception=False,
                )

                kb_msp_file_name_error_list.append(f"{kb_str}, ")
    finally:
        browser.quit()

    logger.info("kb_msp_file_name_list count: " + str(len(kb_msp_file_name_list)))
    logger.info("kb_msp_file_name_list_error count: " + str(len(kb_msp_file_name_error_list)))

    # Write the results to output files
    kb_msp_file_name_list.sort()
    with open(kb_component_file, "w") as f:
        for item in kb_msp_file_name_list:
            f.write("%s\n" % item)

    kb_msp_file_name_error_list.sort()
    kb_component_error_file = os.path.join(current_script_folder, "output_msp_file_name_for_specified_kb_error.txt")
    with open(kb_component_error_file, "w") as f:
        for item in kb_msp_file_name_error_list:
            f.write("%s\n" % item)

    logger.info("Please check output file: " + kb_component_file)
    logger.info("Please check error output file: " + kb_component_error_file)

component_set = set()
try:
    with open(kb_component_file, "r") as f:
        for line in f:
            parts = line.strip().split(",")
            if len(parts) == 3:
                component_set.add(parts[2])
except Exception as ex:
    logger.error(f"Error reading component names: {ex}")

url = "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016#list-of-all-msp-files"
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
    if filename.split("-")[0] in component_set:
        logger.info(f"{filename} is in component_set")
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

logger.info("List length: " + str(len(msp_file_list)))

# Save msp_file_list to a txt file
msp_file_list_file = os.path.join(current_script_folder, "output_msp_file_list.log")
with open(msp_file_list_file, "w", encoding="utf-8") as f:
    for msp_file in msp_file_list:
        # Save as CSV: filename,product,non_security_release_date,non_security_KB,security_release_date,security_KB
        if msp_file.security_greater_than_non_security:
            print(
                msp_file.filename + "," + msp_file.security_release_date + ",KB" + msp_file.security_KB,
                file=f,
            )
        else:
            print(
                msp_file.filename + "," + msp_file.non_security_release_date + ",KB" + msp_file.non_security_KB,
                file=f,
            )

logger.info(f"Saved msp_file_list to {msp_file_list_file}")

options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
browser = webdriver.Edge(options=options)

kb_list_download = []
kb_list_ignored = []
download_links = []

logger.info(f"The target download folder is {target_download_folder}")

try:
    for msp_file_info in msp_file_list:
        logger.info(
            "***Start component: "
            + msp_file_info.filename
            + ", non-security update: KB"
            + msp_file_info.non_security_KB
            + "; security update: KB"
            + msp_file_info.security_KB
        )

        # Decide which KB to use for download
        if msp_file_info.security_greater_than_non_security:
            logger.info("******###Ignore non security kb: " + msp_file_info.tostring())
            current_kb_number = "KB" + msp_file_info.security_KB
            kb_list_ignored.append("KB" + msp_file_info.non_security_KB)
        else:
            logger.info("******@@@Ignore security kb: " + msp_file_info.tostring())
            current_kb_number = "KB" + msp_file_info.non_security_KB
            kb_list_ignored.append("KB" + msp_file_info.security_KB)

        kb_list_download.append(current_kb_number + "," + msp_file_info.filename)

        # Use the correct KB number for download links
        links = get_download_link(browser, current_kb_number, target_download_folder)
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
finally:
    browser.quit()

with open(os.path.join(current_script_folder, "tmp_download_kb_list.log"), "w") as f:
    for element in kb_list_download:
        print(element, file=f)

with open(os.path.join(current_script_folder, "tmp_ignored_kb_list.log"), "w") as f:
    for element in kb_list_ignored:
        print(element, file=f)

with open(os.path.join(current_script_folder, "tmp_kb_links_list.log"), "w") as f:
    for element in download_links:
        print(element[0] + ":" + element[1] + ":" + element[2], file=f)

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
