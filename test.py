import urllib.request

from lxml import etree, html

# Define the URL for the Office 2016 MSP files list
url = "https://docs.microsoft.com/en-us/officeupdates/msp-files-office-2016#list-of-all-msp-files"

# Download the web page content
with urllib.request.urlopen(url) as page:
    content = page.read().decode("utf-8")

# Save the raw HTML content to a log file for debugging
with open("log_content.log", "w", encoding="utf-8") as f:
    f.write(content)

# Parse the HTML content using lxml
doc = html.document_fromstring(content)

# Save the parsed HTML tree as a string for debugging
with open("log_doc.log", "w", encoding="utf-8") as f:
    f.write(etree.tostring(doc, pretty_print=True, encoding="unicode"))

# Extract all tables from the HTML
msp_tables = doc.xpath("//html/body//table")
if len(msp_tables) < 2:
    raise ValueError("Expected at least two tables in the HTML document.")

# Get the string representation of the second table (index 1)
table_string = etree.tostring(msp_tables[1], pretty_print=True, encoding="unicode")

# Extract column headers from the second table
th_list = doc.xpath("//html/body//table[2]/thead/tr/th")
print("Column count:", len(th_list))
for i, th in enumerate(th_list):
    print(f"th[{i}]: {th.text_content().strip()}")

# Extract all rows from the second table's body
td_list = doc.xpath("//html/body//table[2]/tbody/tr")
print("Row count:", len(td_list))
