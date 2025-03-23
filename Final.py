import xlwings as xw

# Define the green color (RGB tuple)
green_color = (144, 238, 144)  # light green

# --- Start a hidden Excel instance ---
app = xw.App(visible=False)

# --- Open Workbooks ---
try:
    wb_source = app.books.open('Greely Survey Review.xlsx')  # Source workbook (contains "PremiseID" in column E)
    print("Source workbook opened successfully.")
except Exception as e:
    app.quit()
    raise Exception("Failed to open source workbook: " + str(e))

try:
    wb_dest = app.books.open('2024-12-02 Services with Customer Side Not Installed.xlsx')  # Destination workbook (contains "Premise Number")
    print("Destination workbook opened successfully.")
except Exception as e:
    wb_source.close()
    app.quit()
    raise Exception("Failed to open destination workbook: " + str(e))

# --- Access Sheets ---
source_sheet_name = "LeadSurvey"              # Adjust as needed (source)
dest_sheet_name   = "INACTIVE PREMISE REVIEW"  # Adjust as needed (destination)

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

# --- In Destination: Locate the header "Premise Number" and build a set of values ---
dest_header = "Premise Number"
dest_col = find_column_by_header(ws_dest, dest_header)
if dest_col is None:
    app.quit()
    raise ValueError(f"Header '{dest_header}' not found in destination sheet '{dest_sheet_name}'.")
print(f"Found header '{dest_header}' in destination at column {dest_col}.")

last_row_dest = ws_dest.range((ws_dest.cells.last_cell.row, dest_col)).end('up').row
dest_values_set = set()
for i in range(2, last_row_dest + 1):
    val = ws_dest.range((i, dest_col)).value
    if val is not None:
        dest_values_set.add(str(val).strip())
print(f"Total premise numbers found in destination: {len(dest_values_set)}")

# --- In Source: Locate the header "PremiseID" ---
source_header = "PremiseID"
source_col = find_column_by_header(ws_source, source_header)
if source_col is None:
    app.quit()
    raise ValueError(f"Header '{source_header}' not found in source sheet '{source_sheet_name}'.")
print(f"Found header '{source_header}' in source at column {source_col}.")

last_row_source = ws_source.range((ws_source.cells.last_cell.row, source_col)).end('up').row
print(f"Source sheet last row: {last_row_source}")

# --- Process each row in the source sheet ---
match_count = 0
for i in range(2, last_row_source + 1):
    cell_val = ws_source.range((i, source_col)).value
    cell_val_str = str(cell_val).strip() if cell_val is not None else ""
    if cell_val and cell_val_str in dest_values_set:
        # Highlight the entire row.
        # This highlights all columns in that row.
        ws_source.range(f"{i}:{i}").color = green_color
        match_count += 1

print(f"Total matching rows highlighted in source: {match_count}")
if match_count == 0:
    print("Warning: No matching rows were highlighted.")
else:
    print("At least one row was highlighted with the green color in the source workbook.")

# --- Save the Updated Source Workbook in place ---
try:
    wb_source.save()  # Save changes to the source workbook
    print("Source workbook saved successfully (in place).")
except Exception as e:
    app.quit()
    raise Exception("Failed to save the source workbook: " + str(e))

# --- Make the Excel app visible so you can inspect the highlights ---
app.visible = True
print("Excel is now visible. Please inspect the source workbook for the highlighted rows.")

# (Workbooks remain open so you can inspect them.)