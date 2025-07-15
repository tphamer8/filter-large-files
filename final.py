import gspread
from oauth2client.service_account import ServiceAccountCredentials
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import time

import os
import requests
import re
from urllib.parse import unquote, urlparse

# Authenticate with Google Sheets
def authenticate_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    return client

# Open the spreadsheet
def open_spreadsheet(client, spreadsheet_name):
    spreadsheet = client.open(spreadsheet_name)
    return spreadsheet

# Process rows and add data to "Old files" sheet
def process_rows(spreadsheet):
    page_content = spreadsheet.worksheet('Page Content')  # First sheet
    old_files = spreadsheet.worksheet('Old Files')  # Access "Old Files" sheet
    rows = page_content.get_all_records()  # Fetch all rows

    for row_idx, row in enumerate(rows, start=2):  # start=2 because Google Sheets is 1-indexed
        if row['Extract Old PDFs'] == 'TRUE':  # If checkbox is checked
            post_title = row['post_title']
            url = row['URL']
            processed = (row['Processed'] == 'Yes')

            if processed:
                print(f"{post_title} already processed.")
                continue

            print(f"Processing: {post_title}")
            
            # Update the "Processed" column (Column 8, which is column H)
            page_content.update_cell(row_idx, 8, 'Yes')

            # Add links to "Old files" sheet
            old_file_links = get_old(url)
            for old_link in old_file_links:
                file_type = 'PDF' if old_link.endswith('.pdf') else 'Image'  # Identify type

                # Append the row with clickable links
                old_files.append_row([post_title, url, old_link, file_type, 'Pending'])


# Find the file links with sites/default/files using Selenium
def get_old(url):
    # Set up Selenium options for headless browser
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    
    # Load the page using Selenium
    driver.get(url)
    
    # Wait for the page to load (you may want to adjust this for dynamic content)
    time.sleep(3)
    
    # Get the page source (HTML)
    html = driver.page_source
    driver.quit()  # Close the browser

    # Parse the page with BeautifulSoup
    soup = BeautifulSoup(html, 'html.parser')
    links = []

    # Find all the pdf links
    for link in soup.find_all('a', href=True):
        href = link['href']
        # Filter by sites/default/files
        if 'sites/default/files' in href:
            full_link = f"https://law.stanford.edu/wp-content/uploads{href}"
            links.append(full_link)
    
    # Find all resources with src (like images)
    for tag in soup.find_all(src=True):
        src = tag['src']
        # Filter by sites/default/files
        if 'sites/default/files' in src:
            full_link = f"{src}"
            links.append(full_link)
    
    return links

def downloadPendingPDFS(spreadsheet):
    # Set Download Directory
    download_dir = "/Users/tpham/Documents/Stanford Webmaster Files/File Reuploads/Automated Downloads"
    os.makedirs(download_dir, exist_ok=True)

    # Set up sheet tab
    old_files = spreadsheet.worksheet('Old Files')  # Access "Old Files" sheet
    rows = old_files.get_all_records()

    # Loop through the sheet
    for row_idx, row in enumerate(rows, start=2):
        if row['Type'] == 'PDF' and row['Status'] == 'Pending':
            url = row['URL']
            filename = os.path.join(download_dir, getFileName(url))
            try:
                response = requests.get(url, headers={
                    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/135.0.0.0 Safari/537.36"
                }, stream=True)

                if response.status_code == 200:
                    with open(filename, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)
                    print(f"Downloaded: {filename}")
                    old_files.update_cell(row_idx, 5, 'Downloaded')
                else:
                    print(f"Failed to download. Status code: {response.status_code}")
                    old_files.update_cell(row_idx, 5, 'Failed')

            except Exception as e:
                print(f"Error downloading: {e}")
            
def fix_links(spreadsheet):
    old_files = spreadsheet.worksheet('Old Files')
    rows = old_files.get_all_records()

    headers = old_files.row_values(1)
    url_col = headers.index('URL') + 1
    status_col = headers.index('Status') + 1

    for row_idx, row in enumerate(rows, start=2):
        if row['Type'] == 'PDF' and row['Status'] == 'Failed':
            url = row['URL']
            print("checking link " + url)

            # Look for a second URL scheme (http:// or https://) starting after position 10
            match = re.search(r'(https?://)', url[10:])
            if match:
                idx = match.start() + 10  # Adjust since we searched from index 10
                fixed_url = url[idx:]
                # print("fixing link " + url)

                old_files.update_cell(row_idx, url_col, fixed_url)
                old_files.update_cell(row_idx, status_col, 'Pending')
                print(f"Row {row_idx} updated:\n  Original URL: {url}\n  Fixed URL:    {fixed_url}")

def getFileName(url):
    path = urlparse(url).path
    filename = os.path.basename(path)
    return unquote(filename)

# Main execution
if __name__ == '__main__':
    client = authenticate_google_sheet()
    spreadsheet_name = 'Auto Old Link Updater'  # Replace with your spreadsheet name
    spreadsheet = open_spreadsheet(client, spreadsheet_name)
    # process_rows(spreadsheet)
    downloadPendingPDFS(spreadsheet)
    fix_links(spreadsheet)