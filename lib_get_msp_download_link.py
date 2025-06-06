import logging
import os
import random

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from lib_common_utils import ensure_folder_exists, save_screenshot_with_log


def get_kb_links(browser, kb_number, kb_list_excluded, target_download_folder, logger):
    """Get download links for a KB number unless excluded."""
    links = []

    if check_kb_in_excluded_list(kb_number, kb_list_excluded, logger):
        return links

    links = get_download_link(browser, kb_number, target_download_folder)

    return links


def check_kb_in_excluded_list(current_kb_number, kb_list_excluded, logger):
    """Check if KB is in the excluded list."""
    if current_kb_number in kb_list_excluded:
        logger.info(f">>> @@@ ---{current_kb_number} is excluded by user")
        return True
    return False


def get_download_link(browser, kb_number, target_folder):
    """Main logic to get download links for both x64 and x86 architectures."""
    logger = logging.getLogger("download_kb")

    ensure_folder_exists(target_folder, logger)

    screenshots_folder = os.path.join(target_folder, "screenshots")
    ensure_folder_exists(screenshots_folder, logger)

    url = f"https://www.catalog.update.microsoft.com/Search.aspx?q={kb_number}"
    logger.info(f">>>Start get download link for {kb_number}: {url}")
    browser.get(url)

    window_before = browser.window_handles[0]

    try:
        WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.ID, "tableContainer")))
    except Exception as e:
        save_screenshot_with_log(
            browser,
            os.path.join(screenshots_folder, f"{kb_number}_error.png"),
            logger,
            f"Timeout waiting for tableContainer for {kb_number}",
            exception=True,
        )

    download_button_x64 = ""
    download_button_x86 = ""

    # Try to find download buttons in the first two rows
    for row in [2, 3]:
        try:
            title_element = browser.find_element(By.XPATH, f'//*[@id="tableContainer"]/table/tbody/tr[{row}]/td[2]/a')
            if "64" in title_element.text:
                download_button_x64 = browser.find_element(
                    By.XPATH,
                    f'//*[@id="tableContainer"]/table/tbody/tr[{row}]/td[8]/input',
                )
            elif "32" in title_element.text:
                download_button_x86 = browser.find_element(
                    By.XPATH,
                    f'//*[@id="tableContainer"]/table/tbody/tr[{row}]/td[8]/input',
                )
            else:
                save_screenshot_with_log(
                    browser,
                    os.path.join(screenshots_folder, f"{kb_number}_error.png"),
                    logger,
                    f"Can't find 64bit or 32bit KB download button for {kb_number}",
                )
        except Exception as e:
            if download_button_x64 == "" and download_button_x86 == "":
                save_screenshot_with_log(
                    browser,
                    os.path.join(screenshots_folder, f"{kb_number}_exception.png"),
                    logger,
                    f"Can't find 32bit and 64bit KB download button for {kb_number}",
                    exception=True,
                )
            if download_button_x64 == "":
                save_screenshot_with_log(
                    browser,
                    os.path.join(screenshots_folder, f"{kb_number}_x64_exception.png"),
                    logger,
                    f"Can't find 64bit KB download button for {kb_number}",
                    exception=True,
                )
            if download_button_x86 == "":
                save_screenshot_with_log(
                    browser,
                    os.path.join(screenshots_folder, f"{kb_number}_x86_exception.png"),
                    logger,
                    f"Can't find 32bit KB download button for {kb_number}",
                    exception=True,
                )

    download_links = []
    for arch, button in [("x64", download_button_x64), ("x86", download_button_x86)]:

        link = handle_download_button(
            browser,
            window_before,
            arch,
            button,
            kb_number,
            target_folder,
            screenshots_folder,
            logger,
        )
        if link:
            download_links.append(link)
    return download_links


def handle_download_button(
    browser,
    window_before,
    arch,
    download_button,
    kb_number,
    target_folder,
    screenshots_folder,
    logger,
):
    """Click the download button and extract the download link from the popup."""
    if download_button != "":

        download_button.click()

        logger.debug(f"Clicked download button for {arch} {kb_number}")

        download_link = get_download_link_from_pop_up_window(browser, window_before, screenshots_folder, kb_number)
        if not download_link:
            save_screenshot_with_log(
                browser,
                os.path.join(screenshots_folder, f"{kb_number}_{arch}_error_link.png"),
                logger,
                f"Don't find download link for this {kb_number}",
            )
            return None

        logger.debug(f"download link for {arch} {kb_number} is {download_link}")
        return [arch, kb_number, download_link]
    else:
        logger.warning(f"No download button found for {arch} {kb_number}")
        return None


def get_download_link_from_pop_up_window(browser, window_handle, screenshots_folder, kb_number):
    """Extract the download link from the popup window."""
    logger = logging.getLogger("download_kb")
    download_link = ""
    try:
        window_after = browser.window_handles[1]
        browser.switch_to.window(window_after)
        try:
            WebDriverWait(browser, 30).until(EC.visibility_of_element_located((By.ID, "downloadFiles")))
        except Exception:
            save_screenshot_with_log(
                browser,
                os.path.join(screenshots_folder, f"{kb_number}_exception.png"),
                logger,
                f"Timeout waiting for downloadFiles for {kb_number}",
                exception=True,
            )
        finally:
            pass

        try:
            link_elements = browser.find_elements(By.XPATH, '//*[@id="downloadFiles"]//a')
            for element in link_elements:
                if "-x-none" in element.text:
                    download_link = element.get_attribute("href")
                    break

            if not download_link:
                save_screenshot_with_log(
                    browser,
                    os.path.join(
                        screenshots_folder,
                        f"{kb_number}_{random.randrange(10)}_error.png",
                    ),
                    logger,
                    f"Don't find download link for this {kb_number}",
                )

            WebDriverWait(browser, 30).until(EC.element_to_be_clickable((By.ID, "downloadSettingsCloseButton")))
            close_button = browser.find_element(By.XPATH, '//*[@id="downloadSettingsCloseButton"]')
            close_button.click()
        except Exception:
            save_screenshot_with_log(
                browser,
                os.path.join(screenshots_folder, f"{kb_number}_exception.png"),
                logger,
                f"Exception in extracting download link for {kb_number}",
                exception=True,
            )

            browser.close()
    except Exception:
        logger.exception(f"Error handling popup window for {kb_number}")
    finally:
        browser.switch_to.window(window_handle)

    return download_link
