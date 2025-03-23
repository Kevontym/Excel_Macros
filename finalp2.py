import xlwings as xw

# Define the green color (RGB tuple)
green_color = (144, 238, 144)  # light green

# --- Start a hidden Excel instance ---
app = xw.App(visible=False)

# --- Open Workbooks ---
try:
    wb_source = app.books.open('Greely Survey Review.xlsx')  # Source file with "PremiseID" & "Survey_Sta"
    print("Source workbook opened successfully.")
except Exception as e:
    app.quit()
    raise Exception("Failed to open source workbook: " + str(e))

try:
    wb_dest = app.books.open('2024-12-02 Services with Customer Side Not Installed_up.xlsx')  # Destination file with "Premise Number" & "Greely Survey Material (Kevon)"
    print("Destination workbook opened successfully.")
except Exception as e:
    wb_source.close()
    app.quit()
    raise Exception("Failed to open destination workbook: " + str(e))

# --- Access Sheets ---
# Adjust sheet names if necessary.
source_sheet_name = "LeadSurvey"
dest_sheet_name   = "INACTIVE PREMISE REVIEW"

try:
    ws_source = wb_source.sheets[source_sheet_name]
    print(f"Source sheet '{source_sheet_name}' found.")
except Exception as e:
    app.quit()
    raise Exception(f"Source sheet '{source_sheet_name}' not found: " + str(e))

try:
    ws_dest = wb_dest.sheets[dest_sheet_name]
    print(f"Destination sheet '{dest_sheet_name}' found.")
except Exception as e:
    app.quit()
    raise Exception(f"Destination sheet '{dest_sheet_name}' not found: " + str(e))

# --- Helper: Find header column in row 1 ---
def find_column_by_header(sheet, header):
    for col in range(1, 51):  # Search the first 50 columns
        cell_value = sheet.range((1, col)).value
        if cell_value and str(cell_value).strip().lower() == header.lower():
            return col
    return None

# --- Define Headers ---
# In the source workbook:
source_key_header    = "PremiseID"      # Matching key (e.g., column E)
source_value_header  = "Survey_Sta"     # Value to copy (e.g., column O)

# In the destination workbook:
dest_key_header      = "Premise Number"                      # Matching key (e.g., column P)
dest_target_header   = "Greely Survey Material (Kevon)"        # Target column to receive the copied value (e.g., column J)

# --- Locate Header Columns ---
source_key_col = find_column_by_header(ws_source, source_key_header)
if source_key_col is None:
    app.quit()
    raise ValueError(f"Header '{source_key_header}' not found in source sheet '{source_sheet_name}'.")
print(f"Found header '{source_key_header}' in source at column {source_key_col}.")

source_value_col = find_column_by_header(ws_source, source_value_header)
if source_value_col is None:
    app.quit()
    raise ValueError(f"Header '{source_value_header}' not found in source sheet '{source_sheet_name}'.")
print(f"Found header '{source_value_header}' in source at column {source_value_col}.")

dest_key_col = find_column_by_header(ws_dest, dest_key_header)
if dest_key_col is None:
    app.quit()
    raise ValueError(f"Header '{dest_key_header}' not found in destination sheet '{dest_sheet_name}'.")
print(f"Found header '{dest_key_header}' in destination at column {dest_key_col}.")

dest_target_col = find_column_by_header(ws_dest, dest_target_header)
if dest_target_col is None:
    app.quit()
    raise ValueError(f"Header '{dest_target_header}' not found in destination sheet '{dest_sheet_name}'.")
print(f"Found header '{dest_target_header}' in destination at column {dest_target_col}.")

# --- Determine Last Rows (assuming data starts at row 2) ---
last_row_source = ws_source.range((ws_source.cells.last_cell.row, source_key_col)).end('up').row
last_row_dest   = ws_dest.range((ws_dest.cells.last_cell.row, dest_key_col)).end('up').row
print(f"Source sheet last row: {last_row_source}")
print(f"Destination sheet last row: {last_row_dest}")

# --- Build a Dictionary from the Source File ---
# Map each trimmed key from "PremiseID" to its corresponding trimmed "Survey_Sta" value.
source_dict = {}
for i in range(2, last_row_source + 1):
    key_val = ws_source.range((i, source_key_col)).value
    copy_val = ws_source.range((i, source_value_col)).value
    if key_val is not None:
        key = str(key_val).strip()
        source_dict[key] = str(copy_val).strip() if copy_val is not None else ""
print(f"Total keys in source dictionary: {len(source_dict)}")

# --- Process the Destination File ---
# For each row in the destination file, if the value in the "Premise Number" column exists in source_dict,
# copy the corresponding Survey_Sta value into the "Greely Survey Material (Kevon)" column
# and highlight the entire row green.
match_count = 0
for i in range(2, last_row_dest + 1):
    dest_key_val = ws_dest.range((i, dest_key_col)).value
    if dest_key_val is not None:
        dest_key_str = str(dest_key_val).strip()
        if dest_key_str in source_dict:
            # Copy the Survey_Sta value into the target column.
            ws_dest.range((i, dest_target_col)).value = source_dict[dest_key_str]
            # Highlight the entire row green.
            ws_dest.range(f"{i}:{i}").color = green_color
            match_count += 1

print(f"Total matching rows updated in destination: {match_count}")
if match_count == 0:
    print("Warning: No matching rows were found and updated.")
else:
    print("At least one row was updated and highlighted with green in the destination workbook.")

# --- Save the Updated Destination Workbook ---
try:
    wb_dest.save('2024-12-02 Services with Customer Side Not Installed_highlighted.xlsx')
    print("Destination workbook saved as '2024-12-02 Services with Customer Side Not Installed_highlighted.xlsx'.")
except Exception as e:
    print("Failed to save the destination workbook: " + str(e))

# --- Make Excel Visible for Inspection ---
app.visible = True
print("Excel is now visible. Please inspect the destination workbook for the updated and green highlighted rows.")

# (Workbooks remain open for manual inspection.)