import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

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

# Authenticate and run
def authenticate_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    return client

def open_spreadsheet(client, spreadsheet_name):
    return client.open(spreadsheet_name)

def main():
    client = authenticate_google_sheet()
    spreadsheet = open_spreadsheet(client, 'Media Library Audit - April 2025')
    # filter_links(spreadsheet)
    write_image_titles(spreadsheet)

if __name__ == "__main__":
    main()
