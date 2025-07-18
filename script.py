import gspread
import requests
from oauth2client.service_account import ServiceAccountCredentials
import re
import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import unquote, urlparse
from PIL import Image
from io import BytesIO


def filter_links(spreadsheet):
    source_ws = spreadsheet.worksheet('500 largest files')
    all_values = source_ws.get_all_values()
    headers = all_values[0]
    data_rows = all_values[1:]

    # Indexes
    location_idx = headers.index("Location")
    size_idx = headers.index("Size")

    # Prepare image entries and transferred row indices
    image_rows = []
    transferred_row_nums = []

    for i, row in enumerate(data_rows):
        location = row[location_idx].strip().lower()
        if location.endswith((".jpg", ".png")):
            full_url = row[location_idx].replace("/sites/stanfordlaw", "https://law.stanford.edu", 1)
            size = row[size_idx]
            image_rows.append([size, full_url, ""])  # No "Found on site" content
            transferred_row_nums.append(i + 2)  # account for header offset

    # --- Create Images tab ---
    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Images"))
    except gspread.exceptions.WorksheetNotFound:
        pass
    image_ws = spreadsheet.add_worksheet(title="Images", rows=str(len(image_rows)+1), cols="3")
    image_ws.update("A1:D1", [["Size", "Location", "Found on site", "Notes"]])
    if image_rows:
        image_ws.update(f"A2:C{len(image_rows)+1}", image_rows)

    print(f"✅ {len(image_rows)} images copied to 'Images'. Checkboxes added in '500 largest files'. Notes tab ensured.")

    source_ws = spreadsheet.worksheet('500 largest files')
    all_values = source_ws.get_all_values()
    headers = all_values[0]
    data_rows = all_values[1:]

    # Indexes
    location_idx = headers.index("Location")
    size_idx = headers.index("Size")

    # Prepare rows
    image_rows = []
    transferred_rows = []

    for i, row in enumerate(data_rows):
        location = row[location_idx].strip().lower()
        if location.endswith(".jpg") or location.endswith(".png"):
            full_url = row[location_idx].replace("/sites/stanfordlaw", "https://law.stanford.edu", 1)
            size = row[size_idx]
            image_rows.append([size, full_url, ""])  # No "Found on site"
            transferred_rows.append(i + 2)  # Google Sheets row numbers start at 1

    # --- Create Images tab ---
    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Images"))
    except gspread.exceptions.WorksheetNotFound:
        pass
    image_ws = spreadsheet.add_worksheet(title="Images", rows=str(len(image_rows)+1), cols="3")
    image_ws.update("A1:C1", [["Size", "Location", "Found on site"]])
    if image_rows:
        image_ws.update(f"A2:C{len(image_rows)+1}", image_rows)

    # --- Add "Transferred" checkboxes in original sheet ---
    source_ws.batch_update([{
        "range": f"E{row}",
        "values": [["FALSE"]]
    } for row in transferred_rows])
    source_ws.format(f"E2:E{len(data_rows)+1}", {"checkbox": True})

    # --- Create Notes tab if not exists ---
    try:
        spreadsheet.add_worksheet(title="Notes", rows="100", cols="5")
    except gspread.exceptions.APIError:
        pass

    print(f"✅ {len(image_rows)} images copied to 'Images'. Checkboxes added. Notes tab ensured.")

    source_ws = spreadsheet.worksheet('500 largest files')
    all_values = source_ws.get_all_values()
    headers = all_values[0]
    data_rows = all_values[1:]

    # Identify column indexes
    location_idx = headers.index("Location")
    size_idx = headers.index("Size")
    found_idx = headers.index("Found on site")

    # Prepare rows for the Images tab
    image_rows = []
    transferred_rows = []

    for i, row in enumerate(data_rows):
        location = row[location_idx].strip().lower()
        if location.endswith(".jpg") or location.endswith(".png"):
            # Fix URL
            full_url = row[location_idx].replace("/sites/stanfordlaw", "https://law.stanford.edu", 1)
            size = row[size_idx]
            image_rows.append([size, full_url, ""])  # Leave "Found on site" blank
            transferred_rows.append(i + 2)  # Save the row number (1-indexed + 1 header)

    # --- Create the Images tab ---
    try:
        spreadsheet.del_worksheet(spreadsheet.worksheet("Images"))
    except gspread.exceptions.WorksheetNotFound:
        pass
    image_ws = spreadsheet.add_worksheet(title="Images", rows=str(len(image_rows)+1), cols="3")

    # Add headers and data
    image_ws.append_row(["Size", "Location", "Found on site"])
    for row in image_rows:
        image_ws.append_row(row)

    # --- Add checkboxes to the original sheet for those transferred ---
    for row_num in transferred_rows:
        source_ws.update_cell(row_num, 5, "FALSE")  # Column E = "Transferred"
    source_ws.format(f"E2:E{len(data_rows)+1}", {"checkbox": True})

    print(f"✅ {len(image_rows)} image rows copied to 'Images'. Checkboxes added in '500 largest files'.")

def write_image_titles(spreadsheet):
    source_ws = spreadsheet.worksheet('Images')
    all_values = source_ws.get_all_values()
    headers = all_values[0]

    try:
        location_index = headers.index("Location")
    except ValueError:
        print("No 'Location' column found.")
        return

    # Add "Title" column if not present
    if "Title" not in headers:
        headers.append("Title")
        source_ws.update('A1', [headers])  # update the first row with new headers
        title_index = len(headers) - 1
    else:
        title_index = headers.index("Title")

    # Collect values to update
    updated_values = []
    for row in all_values[1:]:
        url = row[location_index].strip() if len(row) > location_index else ""
        match = re.search(r'/([^/]+)\.(jpg|jpeg|png|webp)$', url, re.IGNORECASE)
        title = match.group(1) if match else ""

        # Pad the row if it's too short
        while len(row) <= title_index:
            row.append("")

        row[title_index] = title
        updated_values.append(row)

    # Write back only the data rows
    start_cell = f"A2"
    source_ws.update(start_cell, updated_values)
    print("Image titles written successfully.")

    source_ws = spreadsheet.worksheet('Images')
    all_values = source_ws.get_all_values()
    headers = all_values[0]
    data_rows = all_values[1:]

    try:
        location_index = headers.index("Location")
    except ValueError:
        print("No 'Location' column found.")
        return

    for row in data_rows:
        if len(row) > location_index:
            url = row[location_index].strip()
            # Match the filename without extension
            match = re.search(r'/([^/]+)\.(jpg|jpeg|png|webp)$', url, re.IGNORECASE)
            if match:
                print(match.group(1))  # Prints filename without extension


    """
    Downloads images from the 'Location' column for rows marked 'true' in the 'Download' column.
    Saves the images to the specified output folder.
    """
    # Define the output folder path
    home_dir = os.path.expanduser("~")  # Get the user's home directory
    output_folder = os.path.join(home_dir, "Documents", "Stanford Webmaster Files", "Images", "Downloaded Images")

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Fetch all rows from the spreadsheet
    data_rows = spreadsheet.get_all_values()
    headers = data_rows[0]

    # Find indices for 'Download' and 'Location' columns
    try:
        download_index = headers.index('Download')
        location_index = headers.index('Location')
    except ValueError as e:
        print(f"❌ Required column missing: {e}")
        return

    # Iterate through rows and download images
    for row_num, row in enumerate(data_rows[1:], start=2):  # Skip header, start row_num at 2
        if len(row) > download_index and row[download_index].strip().lower() == 'true':
            if len(row) > location_index:
                image_url = row[location_index].strip()
                if image_url:
                    try:
                        # Fetch the image
                        response = requests.get(image_url, stream=True)
                        response.raise_for_status()  # Raise an error for bad responses

                        # Extract the filename from the URL
                        filename = os.path.basename(image_url)
                        if not filename:  # Handle cases where the URL doesn't have a valid filename
                            filename = f"image_{row_num}.jpg"
                        output_path = os.path.join(output_folder, filename)

                        # Save the image
                        with open(output_path, "wb") as file:
                            file.write(response.content)
                        print(f"✅ Row {row_num}: Downloaded {filename} to {output_path}")
                    except Exception as e:
                        print(f"❌ Row {row_num}: Failed to download {image_url}: {e}")
                else:
                    print(f"⚠️ Row {row_num}: No URL found in 'Location' column.")
        else:
            print(f"⚠️ Row {row_num}: 'Download' not marked as true.")


    # Define the output folder path
    home_dir = os.path.expanduser("~")  # Get the user's home directory
    output_folder = os.path.join(home_dir, "Documents", "Stanford Webmaster Files", "Images", "Downloaded Images")

    # Authenticate and open the spreadsheet
    data_rows = spreadsheet.get_all_values()
    headers = data_rows[0]
    download_index = headers.index('Download')  # Find the column index for 'Download'

    # Find indices for 'Location' and 'Download'
    try:
        download_index = headers.index('Download')
        location_index = headers.index('Location')
    except ValueError:
        print("Required columns not found in the spreadsheet.")
        return

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    for row in data_rows[1:]: # skip header
        if len(row) > download_index and row[download_index].strip().lower() == 'true':
            image_url = row[headers.index('Location')].strip()
            if image_url:

                try:
                    # Fetch the image
                    response = requests.get(image_url, stream=True)
                    response.raise_for_status()  # Raise an error for bad responses

                    # Extract the filename from the URL
                    filename = os.path.basename(image_url)
                    output_path = os.path.join(output_folder, filename)

                    # Save the image
                    with open(output_path, "wb") as file:
                        file.write(response.content)
                    print(f"✅ Downloaded: {filename} to {output_path}")
                except Exception as e:
                    print(f"❌ Failed to download {image_url}: {e}")


    """
    Downloads images from the 'Location' column for rows marked 'true' in the 'Download' column.
    Saves the images to the specified output folder.
    """
    # Define the output folder path
    home_dir = os.path.expanduser("~")
    output_folder = os.path.join(home_dir, "Documents", "Stanford Webmaster Files", "Images", "Downloaded Images")

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Fetch all rows from the Images worksheet
    try:
        images_ws = spreadsheet.worksheet('Images')
        data_rows = images_ws.get_all_values()
        headers = data_rows[0]
    except gspread.exceptions.WorksheetNotFound:
        print("❌ 'Images' worksheet not found. Run filter_links() first.")
        return

    # Find indices for 'Download' and 'Location' columns
    try:
        download_index = headers.index('Download')
        location_index = headers.index('Location')
    except ValueError as e:
        print(f"❌ Required column missing: {e}")
        return

    # Iterate through rows and download images
    downloaded_count = 0
    for row_num, row in enumerate(data_rows[1:], start=2):
        if len(row) > download_index and row[download_index].strip().lower() == 'true':
            if len(row) > location_index:
                image_url = row[location_index].strip()
                if image_url:
                    try:
                        # Fetch the image
                        response = requests.get(image_url, stream=True, timeout=30)
                        response.raise_for_status()

                        # Extract the filename from the URL
                        filename = os.path.basename(image_url)
                        if not filename or '.' not in filename:
                            filename = f"image_{row_num}.jpg"
                        
                        output_path = os.path.join(output_folder, filename)

                        # Save the image
                        with open(output_path, "wb") as file:
                            file.write(response.content)
                        
                        print(f"✅ Row {row_num}: Downloaded {filename}")
                        downloaded_count += 1
                        
                    except Exception as e:
                        print(f"❌ Row {row_num}: Failed to download {image_url}: {e}")
                else:
                    print(f"⚠️ Row {row_num}: No URL found in 'Location' column.")

    print(f"📊 Total images downloaded: {downloaded_count}")


    """
    Downloads images from the 'Location' column for rows marked 'true' in the 'Download' column.
    Saves the images to the specified output folder.
    """
    # Define the output folder path
    home_dir = os.path.expanduser("~")
    output_folder = os.path.join(home_dir, "Documents", "Stanford Webmaster Files", "Images", "Downloaded Images")

    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Fetch all rows from the Images worksheet
    try:
        images_ws = spreadsheet.worksheet('Images')
        data_rows = images_ws.get_all_values()
        headers = data_rows[0]
    except gspread.exceptions.WorksheetNotFound:
        print("❌ 'Images' worksheet not found. Run filter_links() first.")
        return

    # Find indices for 'Download' and 'Location' columns
    try:
        download_index = headers.index('Download')
        location_index = headers.index('Location')
    except ValueError as e:
        print(f"❌ Required column missing: {e}")
        return

    # Iterate through rows and download images
    downloaded_count = 0
    for row_num, row in enumerate(data_rows[1:], start=2):
        if len(row) > download_index and row[download_index].strip().lower() == 'true':
            if len(row) > location_index:
                image_url = row[location_index].strip()
                if image_url:
                    try:
                        # Fetch the image
                        response = requests.get(image_url, stream=True, timeout=30)
                        response.raise_for_status()

                        # Extract the filename from the URL
                        filename = os.path.basename(image_url)
                        if not filename or '.' not in filename:
                            filename = f"image_{row_num}.jpg"
                        
                        output_path = os.path.join(output_folder, filename)

                        # Save the image
                        with open(output_path, "wb") as file:
                            file.write(response.content)
                        
                        print(f"✅ Row {row_num}: Downloaded {filename}")
                        downloaded_count += 1
                        
                    except Exception as e:
                        print(f"❌ Row {row_num}: Failed to download {image_url}: {e}")
                else:
                    print(f"⚠️ Row {row_num}: No URL found in 'Location' column.")

    print(f"📊 Total images downloaded: {downloaded_count}")


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
    """
    Downloads images using Selenium to bypass 403 restrictions.
    """
    # Define the output folder path
    home_dir = os.path.expanduser("~")
    output_folder = os.path.join(home_dir, "Documents", "Stanford Webmaster Files", "Images", "Downloaded Images")
    
    # Ensure the output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Fetch all rows from the Images worksheet
    try:
        images_ws = spreadsheet.worksheet('Images')
        data_rows = images_ws.get_all_values()
        headers = data_rows[0]
    except gspread.exceptions.WorksheetNotFound:
        print("❌ 'Images' worksheet not found. Run filter_links() first.")
        return

    # Find indices for 'Download' and 'Location' columns
    try:
        download_index = headers.index('Download')
        location_index = headers.index('Location')
    except ValueError as e:
        print(f"❌ Required column missing: {e}")
        return

    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument("--headless")  # Run in background
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    # Initialize the driver
    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        print("✅ Chrome driver initialized successfully")
    except Exception as e:
        print(f"❌ Failed to initialize Chrome driver: {e}")
        return

    downloaded_count = 0
    
    try:
        # Iterate through rows and download images
        for row_num, row in enumerate(data_rows[1:], start=2):
            if len(row) > download_index and row[download_index].strip().lower() == 'true':
                if len(row) > location_index:
                    image_url = row[location_index].strip()
                    if image_url:
                        try:
                            print(f"🔄 Row {row_num}: Attempting to download {image_url}")
                            
                            # Navigate to the image URL
                            driver.get(image_url)
                            time.sleep(2)  # Wait for page to load
                            
                            # Get cookies from the browser session
                            cookies = driver.get_cookies()
                            
                            # Create a cookie jar for requests
                            cookie_dict = {cookie['name']: cookie['value'] for cookie in cookies}
                            
                            # Extract filename from URL
                            filename = os.path.basename(image_url.split('?')[0])
                            if not filename or '.' not in filename:
                                filename = f"image_{row_num}.jpg"
                            
                            output_path = os.path.join(output_folder, filename)
                            
                            # Download using requests with cookies from Selenium
                            headers_for_request = {
                                'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
                                'Referer': image_url
                            }
                            
                            response = requests.get(
                                image_url,
                                headers=headers_for_request,
                                cookies=cookie_dict,
                                stream=True,
                                timeout=30
                            )
                            response.raise_for_status()
                            
                            # Save the image
                            with open(output_path, "wb") as file:
                                for chunk in response.iter_content(chunk_size=8192):
                                    file.write(chunk)
                            
                            print(f"✅ Row {row_num}: Downloaded {filename}")
                            downloaded_count += 1
                            
                            # Small delay between downloads
                            time.sleep(1)
                            
                        except Exception as e:
                            print(f"❌ Row {row_num}: Failed to download {image_url}: {e}")
                    else:
                        print(f"⚠️ Row {row_num}: No URL found in 'Location' column.")
    
    finally:
        # Always close the driver
        driver.quit()
        print(f"📊 Total images downloaded: {downloaded_count}")


    # Set Download Directory
    download_dir = "/Users/tpham/Documents/Stanford Webmaster Files/Images/Automated Downloads"
    os.makedirs(download_dir, exist_ok=True)

    # Set up sheet tab
    image_links = spreadsheet.worksheet('Images')  # Access "Old Files" sheet
    rows = image_links.get_all_records()

    # Loop through the sheet
    for row_idx, row in enumerate(rows, start=2):
        if row['Download'] == 'Pending':
            url = row['Location']
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
                    image_links.update_cell(row_idx, 6, 'Downloaded')
                else:
                    print(f"Failed to download. Status code: {response.status_code}")
                    image_links.update_cell(row_idx, 6, 'Failed')

            except Exception as e:
                print(f"Error downloading: {e}")

# def download_image(spreadsheet):

    # Set Download Directory
    download_dir = "/Users/tpham/Documents/Stanford Webmaster Files/Images/Automated Downloads"
    os.makedirs(download_dir, exist_ok=True)

    # Set up sheet tab
    image_links = spreadsheet.worksheet('Images')
    rows = image_links.get_all_records()

    # Resize settings
    max_size = (2000, 2000)  # Fit within 2000x2000 box

    for row_idx, row in enumerate(rows, start=2):
        if row.get('Download', '').strip().lower() == 'pending' or row.get('Download', '').strip().lower() == 'p':
            url = row.get('Location', '').strip()
            if not url:
                continue

            try:
                response = requests.get(url, headers={
                    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
                }, stream=True, timeout=30)

                if response.status_code == 200:
                    # Open image from stream
                    image = Image.open(BytesIO(response.content))
                    image.thumbnail(max_size)  # Resize while keeping aspect ratio

                    # Force save as JPEG
                    base_filename = os.path.splitext(getFileName(url))[0]
                    output_path = os.path.join(download_dir, f"{base_filename}.jpg")

                    # Convert to RGB if not JPEG-compatible
                    if image.mode in ("RGBA", "P"):  # Handle transparency or palettes
                        image = image.convert("RGB")

                    # Save with JPEG compression
                    image.save(output_path, format='JPEG', quality=95, optimize=True)
                    print(f"✅ Downloaded and resized: {output_path}")

                    # Mark as downloaded
                    image_links.update_cell(row_idx, 6, 'Downloaded')
                else:
                    print(f"❌ Failed to download (status {response.status_code}): {url}")
                    image_links.update_cell(row_idx, 6, 'Failed')

            except Exception as e:
                print(f"❌ Error downloading {url}: {e}")
                image_links.update_cell(row_idx, 6, 'Failed')

def download_image(spreadsheet):
    import os
    import requests
    from urllib.parse import urlparse, unquote
    from PIL import Image
    from io import BytesIO

    # Set Download Directory
    download_dir = "/Users/tpham/Documents/Stanford Webmaster Files/Images/Automated Downloads"
    os.makedirs(download_dir, exist_ok=True)

    # Set up sheet tab
    image_links = spreadsheet.worksheet('Images')
    rows = image_links.get_all_records()

    # Resize settings
    max_size = (2000, 2000)  # Fit within 2000x2000 box

    for row_idx, row in enumerate(rows, start=2):
        if row.get('Download', '').strip().lower() in ['pending', 'p', 'true']:
            url = row.get('Location', '').strip()
            if not url:
                continue

            try:
                response = requests.get(url, headers={
                    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7)"
                }, stream=True, timeout=30)

                if response.status_code == 200:
                    # Open image from stream
                    image = Image.open(BytesIO(response.content))
                    icc_profile = image.info.get("icc_profile")  # Try to retain ICC color profile
                    image.thumbnail(max_size)  # Resize while keeping aspect ratio

                    # Extract file name and extension
                    parsed_url = urlparse(url)
                    base_filename = os.path.splitext(os.path.basename(parsed_url.path))[0]
                    original_ext = os.path.splitext(parsed_url.path)[1].lower()
                    output_path = os.path.join(download_dir, base_filename + original_ext)

                    # Convert if saving as JPEG and incompatible mode
                    if original_ext in ['.jpg', '.jpeg'] and image.mode not in ("RGB", "L"):
                        image = image.convert("RGB")

                    # Save image with proper format and ICC profile if available
                    save_args = {"optimize": True}
                    if original_ext in ['.jpg', '.jpeg']:
                        save_args["format"] = 'JPEG'
                        save_args["quality"] = 95
                        if icc_profile:
                            save_args["icc_profile"] = icc_profile
                    else:
                        save_args["format"] = image.format or "PNG"  # fallback

                    image.save(output_path, **save_args)

                    print(f"✅ Downloaded and resized: {output_path}")
                    image_links.update_cell(row_idx, 6, 'Downloaded')
                else:
                    print(f"❌ Failed to download (status {response.status_code}): {url}")
                    image_links.update_cell(row_idx, 6, 'Failed')

            except Exception as e:
                print(f"❌ Error downloading {url}: {e}")
                image_links.update_cell(row_idx, 6, 'Failed')

def getFileName(url):
    path = urlparse(url).path
    filename = os.path.basename(path)
    return unquote(filename)

# Authenticate and run
def authenticate_google_sheet():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(credentials)
    return client

def open_spreadsheet(client, sheet_name):
    return client.open(sheet_name)

def main():
    client = authenticate_google_sheet()
    spreadsheet = open_spreadsheet(client, 'Media Library Audit - April 2025')
    
    # Uncomment the functions you want to run:
    # filter_links(spreadsheet)
    # write_image_titles(spreadsheet)
    download_image(spreadsheet)

if __name__ == "__main__":
    main()
