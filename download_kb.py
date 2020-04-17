import time


def DownloadFile(browser, KB, product):

    browser.get(str.format('https://support.microsoft.com/help/{0}', KB))
    browser.implicitly_wait(5)
    
    link_tag_32bit = ""
    link_tag_64bit = ""
    get_download_link_32bit = "false"
    get_download_link_64bit = "false"

    link_tag_list_32bit = []
    link_tag_list_32bit.append(str.format(
        "Download update {0} for 32-bit version of {1}", KB, product))
    link_tag_list_32bit.append(str.format(
        "Download update KB{0} for 32-bit version of {1}", KB, product))
    link_tag_list_32bit.append(str.format(
        "Download security update {0} for the 32-bit version of {1}", KB, product))
    link_tag_list_32bit.append(str.format(
        "Download security update KB{0} for the 32-bit version of {1}", KB, product))

    link_tag_list_64bit = []
    link_tag_list_64bit.append(str.format(
        "Download update {0} for 64-bit version of {1}", KB, product))
    link_tag_list_64bit.append(str.format(
        "Download update KB{0} for 64-bit version of {1}", KB, product))
    link_tag_list_64bit.append(str.format(
        "Download security update {0} for the 64-bit version of {1}", KB, product))
    link_tag_list_64bit.append(str.format(
        "Download security update KB{0} for the 64-bit version of {1}", KB, product))

    for item in link_tag_list_32bit:
        if(get_download_link_32bit == "false"):
            try:
                link_tag_32bit = browser.find_element_by_link_text(item)
                get_download_link_32bit = "true"
            except:
                #print("NO TEXT: " + item)
                continue
        else:
            break

    for item in link_tag_list_64bit:
        if(get_download_link_64bit == "false"):
            try:
                link_tag_64bit = browser.find_element_by_link_text(item)
                get_download_link_64bit = "true"
            except:
                #print("NO TEXT: " + item)
                continue
        else:
            break

    if(get_download_link_32bit == "true"):
        try:
            href_link_tag_32bit = link_tag_32bit.get_attribute("href")
            print("32bit KB" + KB + ": " + href_link_tag_32bit)
        except:
            get_download_link_32bit = "false"
            browser.save_screenshot("KB" + KB + "_32bit_link.png")   
    else:
        print("Failed to get 32bit download link for this KB" + KB)


    if(get_download_link_64bit == "true"):
        try:
            href_link_tag_64bit = link_tag_64bit.get_attribute("href")
            print("64bit KB" + KB + ": " + href_link_tag_64bit)
        except:
            get_download_link_64bit = "false"
            browser.save_screenshot("KB" + KB + "_64bit_link.png")
    else:
        print("Failed to get 64bit download link for this KB" + KB)


    if(get_download_link_32bit == "true"):
        try:   
            print("start download 32bit KB file")
            # download 32bit file
            #print(href_link_tag_32bit)
            browser.get(href_link_tag_32bit)
            browser.implicitly_wait(5)

            download_link = browser.find_elements_by_class_name("download-button")
            # print(len(download_link))
            # print(download_link[0].get_attribute("href"))
            download_link_href = download_link[0].get_attribute("href")
            browser.get(download_link_href)
        
            time.sleep(15)
        except:
            print("FAILED TO GET 32bit KB" + KB + ": "  + href_link_tag_32bit)
            browser.save_screenshot("KB" + KB + "_32bit_error.png")

        try:
            if(get_download_link_64bit == "true"):
                print("start download 64bit KB file")
                # download 64bit file
                #print(href_link_tag_64bit)
                browser.get(href_link_tag_64bit)
                browser.implicitly_wait(5)

                download_link = browser.find_elements_by_class_name("download-button")
                # print(len(download_link))
                # print(download_link[0].get_attribute("href"))
                download_link_href = download_link[0].get_attribute("href")
                browser.get(download_link_href)

                time.sleep(15)
        except:
            print("FAILED TO GET 64bit KB" + KB + ": " + href_link_tag_64bit)
            browser.save_screenshot("KB" + KB + "_64bit_error.png")