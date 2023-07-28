import logging
import os
import random
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def get_download_link(browser, kb_number, target_folder):
    logger = logging.getLogger("download_kb")

    if not os.path.exists(target_folder):
        os.makedirs(target_folder)

    if not os.path.exists(target_folder + "screenshots"):
        os.makedirs(target_folder + "screenshots")
        logging.info(
            "Screenshots folder doesn't exist, create it: "
            + target_folder
            + "screenshots"
        )

    url = str.format(
        "https://www.catalog.update.microsoft.com/Search.aspx?q={0}", kb_number
    )
    logger.info(str.format(">>>Start get download link for {0}: {1}", kb_number, url))
    browser.get(url)

    window_before = browser.window_handles[0]

    try:
        WebDriverWait(browser, 30).until(
            EC.visibility_of_element_located((By.ID, "tableContainer"))
        )
    except:
        browser.save_screenshot(
            target_folder + "screenshots\\" + kb_number + "_error.png"
        )
    finally:
        pass

    download_button_x64 = ""
    download_button_x86 = ""

    try:
        title_element = browser.find_element(
            "xpath", '//*[@id="tableContainer"]/table/tbody/tr[2]/td[2]/a'
        )
        if title_element.text.find("64") > -1:
            download_button_x64 = browser.find_element(
                "xpath", '//*[@id="tableContainer"]/table/tbody/tr[2]/td[8]/input'
            )
        elif title_element.text.find("32") > -1:
            download_button_x86 = browser.find_element(
                "xpath", '//*[@id="tableContainer"]/table/tbody/tr[2]/td[8]/input'
            )
        else:
            logger.error(
                "Can't find 64bit or 32bit KB download button for " + kb_number
            )
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_error.png"
            )
    except:
        if download_button_x64 == "" and download_button_x86 == "":
            logger.exception(
                "Can't find 32bit and 64bit KB download button for " + kb_number
            )
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_exception.png"
            )

        if download_button_x64 == "":
            logger.exception("Can't find 64bit KB download button for " + kb_number)
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_x64_exception.png"
            )

        if download_button_x86 == "":
            logger.exception("Can't find 32bit KB download button for " + kb_number)
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_x86_exception.png"
            )

    try:
        title_element = browser.find_element(
            "xpath", '//*[@id="tableContainer"]/table/tbody/tr[3]/td[2]/a'
        )
        if title_element.text.find("64") > -1:
            download_button_x64 = browser.find_element(
                "xpath", '//*[@id="tableContainer"]/table/tbody/tr[3]/td[8]/input'
            )
        elif title_element.text.find("32") > -1:
            download_button_x86 = browser.find_element(
                "xpath", '//*[@id="tableContainer"]/table/tbody/tr[3]/td[8]/input'
            )
        else:
            logger.error(
                "Can't find 64bit or 32bit KB download button for " + kb_number
            )
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_error.png"
            )
    except:
        if download_button_x64 == "" and download_button_x86 == "":
            logger.exception(
                "Can't find 32bit and 64bit KB download button for " + kb_number
            )
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_exception.png"
            )

        if download_button_x64 == "":
            logger.exception("Can't find 64bit KB download button for " + kb_number)
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_x64_exception.png"
            )

        if download_button_x86 == "":
            logger.exception("Can't find 32bit KB download button for " + kb_number)
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_x86_exception.png"
            )

    download_links = []

    if download_button_x64 != "":
        download_button_x64.click()
        download_link_x64 = get_download_link_from_pop_up_window(
            browser, window_before, target_folder, kb_number
        )

        if download_link_x64 == "":
            logger.error("Don't find download link for this " + kb_number)
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_x64_error_link.png"
            )
        else:
            download_links.append(["x64", kb_number, download_link_x64])

        logger.debug(
            str.format(
                "download link for 64bit {0} is {1}", kb_number, download_link_x64
            )
        )

    if download_button_x86 != "":
        download_button_x86.click()
        download_link_x86 = get_download_link_from_pop_up_window(
            browser, window_before, target_folder, kb_number
        )

        if download_link_x86 == "":
            logger.error("Don't find download link for this " + kb_number)
            browser.save_screenshot(
                target_folder + "screenshots\\" + kb_number + "_x86_error_link.png"
            )
        else:
            download_links.append(["x86", kb_number, download_link_x86])

        logger.debug(
            str.format(
                "download link for 32bit {0} is {1}", kb_number, download_link_x86
            )
        )

    return download_links


def get_download_link_from_pop_up_window(
    browser, window_handle, target_folder, kb_number
):
    logger = logging.getLogger("download_kb")

    # get the window handle after a new window has opened
    window_after = browser.window_handles[1]

    # switch on to new child window
    browser.switch_to.window(window_after)

    try:
        WebDriverWait(browser, 30).until(
            EC.visibility_of_element_located((By.ID, "downloadFiles"))
        )
    except:
        browser.save_screenshot(
            target_folder + "screenshots\\" + kb_number + "_exception.png"
        )
    finally:
        pass

    try:
        # http://download.windowsupdate.com/c/msdownload/update/software/crup/2019/02/access-x-none_619933a5aaeb898b29f41d0c5660531d958e0abb.cab
        link_elements = browser.find_elements("xpath", '//*[@id="downloadFiles"]//a')

        download_link = ""
        for element in link_elements:
            if element.text.find("-x-none") > -1:
                download_link = element.get_attribute("href")

        if download_link == "":
            logger.error("Don't find download link for this " + kb_number)
            browser.save_screenshot(
                target_folder
                + "screenshots\\"
                + kb_number
                + "_"
                + random.randrange(10)
                + "_error.png"
            )

        WebDriverWait(browser, 30).until(
            EC.element_to_be_clickable((By.ID, "downloadSettingsCloseButton"))
        )

        close_button = browser.find_element(
            "xpath", '//*[@id="downloadSettingsCloseButton"]'
        )
        close_button.click()
    except:
        browser.save_screenshot(
            target_folder + "screenshots\\" + kb_number + "_exception.png"
        )
        browser.close()
    finally:
        browser.switch_to.window(window_handle)

    return download_link
