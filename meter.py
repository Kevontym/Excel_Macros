import xlwings as xw

# Define the blue color (RGB tuple)
blue_color = (0, 176, 240)  # blue

# --- Start a hidden Excel instance ---
app = xw.App(visible=False)

# --- Open Workbooks ---
try:
    wb_source = app.books.open('Meter reading reasons.XLSX')  # Source file
    print("Source workbook opened successfully.")
except Exception as e:
    app.quit()
    raise Exception("Failed to open source workbook: " + str(e))

try:
    wb_dest = app.books.open('2024-12-02 Services with Customer Side Not Installed.xlsx')  # Destination file
    print("Destination workbook opened successfully.")
except Exception as e:
    wb_source.close()
    app.quit()
    raise Exception("Failed to open destination workbook: " + str(e))

# --- Access Sheets ---
ws_source = wb_source.sheets["Sheet3"]
ws_dest   = wb_dest.sheets["INACTIVE PREMISE REVIEW"]
print(f"Source sheet: {ws_source.name}")
print(f"Destination sheet: {ws_dest.name}")

# --- Helper: Find header column in row 1 ---
def find_column_by_header(sheet, header):
    for col in range(1, 51):  # Check first 50 columns
        cell_value = sheet.range((1, col)).value
        if cell_value and str(cell_value).strip().lower() == header.lower():
            return col
    return None

# --- Define Headers ---
# Source headers
source_match_header = "Installat."       # Used for matching (e.g., column A)
source_value_header = "RR"                # Column whose value is normally copied (now replaced by "X")
# Destination headers
dest_match_header = "Installation"        # Used for matching (e.g., column P)
dest_target_header = "Meter Reading Reason 22 (Kevon)"  # Target column (e.g., column Q)

# --- Locate Header Columns ---
source_match_col = find_column_by_header(ws_source, source_match_header)
if source_match_col is None:
    app.quit()
    raise ValueError(f"Header '{source_match_header}' not found in source sheet '{ws_source.name}'.")
print(f"Found header '{source_match_header}' in source at column {source_match_col}.")

source_value_col = find_column_by_header(ws_source, source_value_header)
if source_value_col is None:
    app.quit()
    raise ValueError(f"Header '{source_value_header}' not found in source sheet '{ws_source.name}'.")
print(f"Found header '{source_value_header}' in source at column {source_value_col}.")

dest_match_col = find_column_by_header(ws_dest, dest_match_header)
if dest_match_col is None:
    app.quit()
    raise ValueError(f"Header '{dest_match_header}' not found in destination sheet '{ws_dest.name}'.")
print(f"Found header '{dest_match_header}' in destination at column {dest_match_col}.")

dest_target_col = find_column_by_header(ws_dest, dest_target_header)
if dest_target_col is None:
    app.quit()
    raise ValueError(f"Header '{dest_target_header}' not found in destination sheet '{ws_dest.name}'.")
print(f"Found header '{dest_target_header}' in destination at column {dest_target_col}.")

# --- Determine Last Rows (assuming data starts at row 2) ---
last_row_source = ws_source.range((ws_source.cells.last_cell.row, source_match_col)).end('up').row
last_row_dest = ws_dest.range((ws_dest.cells.last_cell.row, dest_match_col)).end('up').row
print(f"Source sheet last row: {last_row_source}")
print(f"Destination sheet last row: {last_row_dest}")

# --- Build a Dictionary from the Source File ---
# Maps trimmed value from "Installat." to the corresponding "RR" value (but we won't use it since we are inserting "X")
source_dict = {}
for i in range(2, last_row_source + 1):
    key_val = ws_source.range((i, source_match_col)).value
    copy_val = ws_source.range((i, source_value_col)).value
    if key_val is not None:
        key = str(key_val).strip()
        source_dict[key] = str(copy_val).strip() if copy_val is not None else ""
print(f"Total keys in source dictionary: {len(source_dict)}")

# --- Process the Destination File ---
# For each row in the destination file (starting at row 2), if the value in the "Installation" column exists in source_dict,
# place an "X" in the cell of the "Meter Reading Reason 22 (Kevon)" column and highlight that cell blue.
match_count = 0
for i in range(2, last_row_dest + 1):
    dest_match_val = ws_dest.range((i, dest_match_col)).value
    if dest_match_val is not None:
        dest_match_str = str(dest_match_val).strip()
        if dest_match_str in source_dict:
            ws_dest.range((i, dest_target_col)).value = "X"  # Instead of copying the source value, place "X"
            ws_dest.range((i, dest_target_col)).color = blue_color
            match_count += 1

print(f"Total matching rows updated in destination: {match_count}")
if match_count == 0:
    print("Warning: No matching rows were found and updated.")

# --- Save the Updated Destination Workbook ---
try:
    wb_dest.save('2024-12-02 Services with Customer Side Not Installed_up.xlsx')
    print("Destination workbook saved as '2024-12-02 Services with Customer Side Not Installed_highlighted.xlsx'.")
except Exception as e:
    print("Failed to save the destination workbook: " + str(e))

# --- Make Excel Visible for Inspection ---
app.visible = True
print("Excel is now visible. Please inspect the destination workbook for the blue highlighted cells.")