from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Test Firefox
# options = webdriver.FirefoxOptions()
# browser = webdriver.Firefox(options=options)

# Test Edge
options = webdriver.EdgeOptions()
options.add_argument("--start-maximized")
# options.add_argument("--headless")  # Uncomment to run Edge in headless mode
browser = webdriver.Edge(options=options)

try:
    browser.get("https://www.bing.com")
    # Wait until the page title contains 'Bing'
    WebDriverWait(browser, 10).until(EC.title_contains("Bing"))
    print("Page loaded successfully.")
except TimeoutException:
    print("Failed to load the page within the given time.")
except WebDriverException as e:
    print(f"WebDriver error: {e}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
finally:
    browser.quit()
