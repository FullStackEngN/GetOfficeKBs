import time
import datetime

def DownloadFile(browser, KB, product):

    currentDirectory = r"C:\Temp\Office2016_KBs\\"

    browser.get(str.format('https://support.microsoft.com/help/{0}', KB))
    time.sleep(5)

    currentDT = datetime.datetime.now()
    print("******1: " + str(currentDT))

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
    link_tag_list_32bit.append(str.format(
        "Download the security update KB{0} for the 32-bit version of {1}", KB, product))

    link_tag_list_64bit = []
    link_tag_list_64bit.append(str.format(
        "Download update {0} for 64-bit version of {1}", KB, product))
    link_tag_list_64bit.append(str.format(
        "Download update KB{0} for 64-bit version of {1}", KB, product))
    link_tag_list_64bit.append(str.format(
        "Download security update {0} for the 64-bit version of {1}", KB, product))
    link_tag_list_64bit.append(str.format(
        "Download security update KB{0} for the 64-bit version of {1}", KB, product))
    link_tag_list_32bit.append(str.format(
        "Download the security update KB{0} for the 64-bit version of {1}", KB, product))

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


    currentDT = datetime.datetime.now()
    print("******2: " + str(currentDT))

    if(get_download_link_32bit == "true"):
        try:
            href_link_tag_32bit = link_tag_32bit.get_attribute("href")
            print("32bit KB" + KB + ": " + href_link_tag_32bit)
        except:
            print("Exception to get 32bit download link for this KB" + KB)
            get_download_link_32bit = "false"
            browser.save_screenshot(
                currentDirectory + "KB" + KB + "_32bit_link.png")
    else:
        print("No matched text for 32bit KB" + KB)

    if(get_download_link_64bit == "true"):
        try:
            href_link_tag_64bit = link_tag_64bit.get_attribute("href")
            print("64bit KB" + KB + ": " + href_link_tag_64bit)
        except:
            print("Exception to get 64bit download link for this KB" + KB)
            get_download_link_64bit = "false"
            browser.save_screenshot(
                currentDirectory + "KB" + KB + "_64bit_link.png")
    else:
        print("No matched text for 64bit KB" + KB)

    
    currentDT = datetime.datetime.now()
    print("******3: " + str(currentDT))

    if(get_download_link_32bit == "true"):
        try:
            print("******start download 32bit " + KB + "file")
            currentDT = datetime.datetime.now()
            print("******4: " + str(currentDT))
            # download 32bit file
            # print(href_link_tag_32bit)
            browser.get(href_link_tag_32bit)
            
            #browser.implicitly_wait(20)
            time.sleep(30)

            download_link = browser.find_elements_by_class_name("download-button")
            # print(len(download_link))
            # print(download_link[0].get_attribute("href"))
            download_link_href = download_link[0].get_attribute("href")
            browser.get(download_link_href)

            time.sleep(30)
        except:
            print("EXCEPTION TO DOWNLOAD 32bit KB" + KB + ": " + href_link_tag_32bit)
            browser.save_screenshot(
                currentDirectory + "KB" + KB + "_32bit_error.png")

        try:
            if(get_download_link_64bit == "true"):
                print("******start download 64bit "+ KB +" file")
                
                currentDT = datetime.datetime.now()
                print("******5: " + str(currentDT))
                # download 64bit file
                # print(href_link_tag_64bit)
                browser.get(href_link_tag_64bit)
                
                #browser.implicitly_wait(20)
                time.sleep(30)

                download_link = browser.find_elements_by_class_name("download-button")
                # print(len(download_link))
                # print(download_link[0].get_attribute("href"))
                download_link_href = download_link[0].get_attribute("href")
                browser.get(download_link_href)

                time.sleep(30)
        except:
            print("EXCEPTION TO DOWNLOAD 64bit KB" + KB + ": " + href_link_tag_64bit)
            browser.save_screenshot(
                currentDirectory + "KB" + KB + "_64bit_error.png")
