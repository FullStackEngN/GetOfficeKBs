# GetOfficeKBs

This repository automates the process of retrieving, downloading, and extracting Microsoft Office MSP (patch) files using official update catalogs and documentation.

## Features

- **Scrape Office Update Catalogs:** Automatically parses Microsoft documentation to get the latest MSP file lists and KB numbers.
- **Download CAB/MSP Files:** Uses Selenium and direct HTTP requests to fetch update files for Office 2013/2016.
- **Architecture Support:** Handles both x64 and x86 architectures.
- **Component Filtering:** Download only the MSPs you need by component.
- **Logging:** Detailed logging for all operations.
- **Automatic Extraction:** Extracts MSP files from downloaded CABs.
- **Error Handling:** Takes screenshots and logs errors for troubleshooting.

## Requirements

- Python 3.7+
- [Selenium](https://pypi.org/project/selenium/)
- [lxml](https://pypi.org/project/lxml/)
- [wget](https://pypi.org/project/wget/)
- Microsoft Edge WebDriver (or modify for your browser)
- [Black](https://pypi.org/project/black/) (for code formatting, optional)

## Setup

1. **Clone the repository:**
    ```sh
    git clone https://github.com/yourusername/GetOfficeKBs.git
    cd GetOfficeKBs
    ```

2. **Install dependencies:**
    ```sh
    pip install selenium
    pip install wget
    pip install lxml
    ```

3. **Download and place the appropriate WebDriver** (e.g., msedgedriver.exe) in your PATH.

4. **(Optional) Configure Black line length in VS Code:**
    ```jsonc
    "black-formatter.args": ["--line-length", "100"]
    ```

## Usage

- **Download all MSP files:**
    ```sh
    python download_all_msp_files.py
    ```

- **Download the latest MSP files for specified KBs:**

    Modify the `input_kb_list_specified.txt` file and add a few KBs.

    For example:

    KB5002362  
    KB5002198  
    KB5002348  

    ```sh
    python download_latest_msp_files_for_specified_kb.py
    ```

- **Download the MSP files for expected KBs:**

    Modify the `input_kb_list_expected.txt` file and add a few KBs.

    For example:

    KB5002362  
    KB5002198  
    KB5002348  

    ```sh
    python download_msp_files_for_expected_kb.py
    ```

## File Overview

- `download_all_msp_files.py` — Main script to download all MSP files for Office.
- `download_latest_msp_files_for_specified_kb.py` — Download the latest MSPs for specified KBs/components.
- `download_msp_files_for_specified_kb.py` — Download the MSPs for expected KBs/components only.
- `lib_get_msp_download_link.py` — Logic for scraping and retrieving download links.
- `lib_common_utils.py` — Utility functions (file I/O, logging, etc).
- `lib_extract_msp.py` — Extracts MSP files from CAB archives.
- `lib_logger_object.py` — Logger setup.
- `lib_msp_file.py` — MSP file data structure.

## Notes

- Make sure you have a stable internet connection.
- Some downloads may require retrying due to network or catalog issues.
- For large-scale downloads, ensure you have enough disk space.
