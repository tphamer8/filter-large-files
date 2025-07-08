import gspread
from oauth2client.service_account import ServiceAccountCredentials
import re

def filter_links(spreadsheet):
    worksheet = spreadsheet.worksheet('500 largest files')
    records = worksheet.get_all_records()

    # Get headers and find jpg/png links
    headers = worksheet.row_values(1)
    image_rows = []
    for row in records:
        location = row.get("Location", "")
        if location.lower().endswith((".jpg", ".png")):
            # Replace the beginning of the URL
            updated_location = location.replace("/sites/stanfordlaw", "https://law.stanford.edu", 1)
            image_rows.append([row.get("Size", ""), updated_location, row.get("Found on site", "")])

    # Create "Images" worksheet or clear if exists
    try:
        image_ws = spreadsheet.worksheet("Images")
        spreadsheet.del_worksheet(image_ws)
    except:
        pass
    image_ws = spreadsheet.add_worksheet(title="Images", rows=str(len(image_rows)+1), cols="5")

    # Write headers and add "Transferred" column
    image_ws.append_row(headers[:3] + ["Notes", "Transferred"])

    # Write image rows and add unchecked checkboxes
    for row in image_rows:
        image_ws.append_row(row + ["", "FALSE"])

    # Add checkboxes to "Transferred" column (column E)
    cell_range = f"E2:E{len(image_rows)+1}"
    image_ws.update(cell_range, [["FALSE"]] * len(image_rows), value_input_option='USER_ENTERED')
    image_ws.format(cell_range, {"checkbox": True})

    print(f"Filtered {len(image_rows)} image links and added to 'Images' tab.")

# Authenticate and run
def authenticate_google_sheet():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    return client

def open_spreadsheet(client, spreadsheet_name):
    spreadsheet = client.open(spreadsheet_name)
    return spreadsheet

def main():
    client = authenticate_google_sheet()
    spreadsheet_name = 'Media Library Audit - April 2025'  # Replace with your spreadsheet name
    spreadsheet = open_spreadsheet(client, spreadsheet_name)
    filter_links(spreadsheet)

if __name__ == "__main__":
    main()
