import win32com.client
import pythoncom
import pywintypes
import os
import tkinter as tk
from tkinter import filedialog
import sys
import psutil
import ctypes

def is_excel_running():
    try:
        for proc in psutil.process_iter(['pid', 'name']):
            if 'EXCEL.EXE' in proc.info['name'].upper():
                return True
    except Exception as e:
        print(f"Error checking Excel processes: {e}")
        return False
    return False

def filed_documents_report_with_phrase_hit_athena():
    MB_OK = 0x0
    MB_ICONERROR = 0x10
    message = "Please close all instances of Excel and try again."
    title = "Error"
    if is_excel_running():
        ctypes.windll.user32.MessageBoxW(None, message, title, MB_OK | MB_ICONERROR)
        sys.exit()

    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes = [("Excel files", "*.xlsx;*.xls")],
        initialdir = os.path.expanduser("~")
        )

    if not file_path:
        print("No file selected. Exiting.")
        return

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = True  # For debugging, can be set to False later

    workbook_path = file_path
    wb = excel.Workbooks.Open(workbook_path)

    try:
        excel.ScreenUpdating = False

        modify_documents_filed_report(wb, excel)

        create_phrase_maintenance_sheet(wb, excel)

        create_phrase_building_sheet(wb, excel)

        create_filter_updates_sheet(wb, excel)

        post_analysis_formatting(wb, excel)

        excel.ScreenUpdating = True

        # Save the workbook
        wb.Save()

    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()

def modify_documents_filed_report(wb, excel):
    ws = wb.Worksheets("Documents Filed Report")

    headers = [
        "Member Match",
        "Summary Match",
        "DOS Match",
        "Signature Match",
        "Patient Match",
        "Provider Match"
    ]

    for head in headers:
        ws.Columns("A:A").Insert()
    
    for col, header in enumerate(headers, start=1):
        ws.Cells(1, col).Value = header

    # Formulas for matches
    last_row = ws.Cells(ws.Rows.Count, "G").End(-4162).Row  # -4162 is xlUp
    ws.Range(f"A2:A{last_row}").Formula = "=IF(L2=N2, \"EXACTMATCH\", \"NEEDSREVIEW\")"
    ws.Range(f"B2:B{last_row}").Formula = "=IF(M2=O2, \"EXACTMATCH\", \"NEEDSREVIEW\")"
    ws.Range(f"C2:C{last_row}").Formula = "=IF(P2=Q2, \"EXACTMATCH\", \"NEEDSREVIEW\")"
    ws.Range(f"D2:D{last_row}").Formula = "=IF(V2=W2, \"EXACTMATCH\", \"NEEDSREVIEW\")"
    ws.Range(f"E2:E{last_row}").Formula = "=IF(COUNTIF(Y2,\"*oun*\"), \"NOTFOUND\",IF(H2=I2,\"EXACTMATCH\",\"NEEDSREVIEW\"))"
    ws.Range(f"F2:F{last_row}").Formula = "=IF(AND(E2=\"NOTFOUND\",R2=S2),\"PTNFEXACTMATCH\",IF(AND(E2=\"NOTFOUND\",R2<>S2),\"PTNFNEEDSREVIEW\",IF(AND(E2=\"EXACTMATCH\",R2=S2),\"EXACTMATCH\",IF(AND(E2=\"NEEDSREVIEW\",R2=S2),\"EXACTMATCH\",\"NEEDSREVIEW\"))))"

    # Clear all formulas and convert to values
    for col in "ABCDEF":
        ws.Columns(col).Copy()
        ws.Columns(col).PasteSpecial(Paste=-4163)  # -4163 is xlPasteValues

    # Add Indexer Review column
    ws.Range("AE1").Value = "Indexer Review"
    ws.Range(f"AE2:AE{last_row}").Formula = "=VLOOKUP(K2, 'Phrase Hit Report'!A$2:I$30000, 5, FALSE)"

    # Add # Documents indexed column
    ws.Range("AF1").Value = "Documents Manually Indexed with No Phrase by HL7 Document Type and HL7 Summary Line"
    ws.Range(f"AF2:AF{last_row}").Formula = "=COUNTIFS(N:N,N:N, O:O,O:O, X:X, \"Manually Indexed\", K:K,\"=0\")"

    # Add # where Patient Not Found column
    ws.Range("AG1").Value = "Indexed Documents with Flag containing No Patient Found and No Phrase Hit by HL7 Document Type and HL7 Summary Line"
    ws.Range(f"AG2:AG{last_row}").Formula = "=COUNTIFS(N:N,N:N,O:O,O:O,K:K,\"=0\",X:X,\"Manually Indexed\",Y:Y,\"*oun*\")"

    # Format dates
    ws.Range("P:P,Q:Q,AA:AA,AB:AB,AC:AC").NumberFormat = "mm/dd/yyyy"

    # Set all rows to no fill
    ws.Range(f"A2:AG{last_row}").Interior.ColorIndex = -4142  # -4142 is xlNone

    # Freeze pane on header row
    ws.Rows("2:2").Select()
    excel.ActiveWindow.FreezePanes = True

    # Auto Fit Rows and set special color for added columns
    ws.Columns("A:AG").AutoFit()
    ws.Range("A1,B1,C1,D1,E1,F1,G1,AE1,AF1,AG1").Interior.ColorIndex = 10

    # Insert DocDetails column
    ws.Columns("G:G").Insert()
    ws.Range("G1").Value = "DocDetails"
    ws.Range(f"G2:G{last_row}").Formula = "=HYPERLINK(\"https://core.indxlogic.com/Document/ViewDetail/\"&H2)"

    # Apply filters
    ws.Range("A1").AutoFilter()
    ws.Range("A1").AutoFilter(Field=25, Criteria1="Manually Indexed")
    ws.Range("A1").AutoFilter(Field=12, Criteria1="<>0")
    ws.Range("A1").AutoFilter(Field=32, Criteria1="Yes")

    

def create_phrase_maintenance_sheet(wb, excel):
    wb.Sheets.Add().Name = "Phrase Maintenance"
    ws_pm = wb.Worksheets("Phrase Maintenance")
    ws_phr = wb.Worksheets("Phrase Hit Report")

    # Copy Phrase Hit Report to Phrase Maintenance
    ws_phr.Range("A:J").Copy(ws_pm.Range("A1"))

    # Rename columns
    new_headers = [
        "Total Hits in Reporting Period (Indexed)",
        "Count DT Changed",
        "Count Summary Changed",
        "Count DOS Changed",
        "Count Signature Changed",
        "Count Patient Found and Changed",
        "Count Patient Found where DOS changed",
        "Count Provider Changed where Patient Found"
    ]
    for col, header in enumerate(new_headers, start=10):
        ws_pm.Cells(1, col+1).Value = header

    # Set formulas for new columns
    last_row = ws_pm.Cells(ws_pm.Rows.Count, "A").End(-4162).Row
    formulas = [
        "=COUNTIFS('Documents Filed Report'!L:L, A2, 'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\")",
        "=COUNTIFS('Documents Filed Report'!L:L, A2,  'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\", 'Documents Filed Report'!A:A,\"NEEDSREVIEW\")",
        "=COUNTIFS('Documents Filed Report'!L:L, A2,'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\", 'Documents Filed Report'!B:B,\"NEEDSREVIEW\")",
        "=COUNTIFS('Documents Filed Report'!L:L, A2, 'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\", 'Documents Filed Report'!C:C,\"NEEDSREVIEW\")",
        "=COUNTIFS('Documents Filed Report'!L:L, A2, 'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\", 'Documents Filed Report'!D:D,\"NEEDSREVIEW\")",
        "=COUNTIFS('Documents Filed Report'!L:L, A2, 'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\", 'Documents Filed Report'!E:E,\"NEEDSREVIEW\")",
        "=COUNTIFS('Documents Filed Report'!L:L, A2, 'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\",  'Documents Filed Report'!E:E, \"EXACTMATCH\",  'Documents Filed Report'!C:C,\"NEEDSREVIEW\")",
        "=COUNTIFS('Documents Filed Report'!L:L, A2, 'Documents Filed Report'!Y:Y,\"Manually Indexed\", 'Documents Filed Report'!AF:AF,\"Yes\",  'Documents Filed Report'!E:E, \"<>NOTFOUND\",  'Documents Filed Report'!F:F,\"NEEDSREVIEW\")"
    ]
    for col, formula in enumerate(formulas, start=10):
        ws_pm.Range(f"{chr(65+col)}2:{chr(65+col)}{last_row}").Formula = formula

    # Format worksheet
    ws_pm.Columns("A:R").AutoFit()
    ws_pm.Range("I:I").NumberFormat = "mm/dd/yyyy"

    # Sort by Total Hits in Reporting Period
    ws_pm.Range(f"A1:P{last_row}").Sort(Key1=ws_pm.Range("J2"), Order1=2, Header=1)  # 2 is xlDescending

    # Apply filter
    ws_pm.Range("A1").AutoFilter()
    ws_pm.Range("A1").AutoFilter(Field=5, Criteria1="Yes")

    # Add Phrase Filter Button
    shape = ws_pm.Shapes.AddShape(5, 0, 0, 100, 20)  # 5 is msoShapeRoundedRectangle
    shape.TextFrame.Characters().Text = "Phrase Filter Button"
    shape.TextFrame.Characters().Font.Size = 10
    shape.Left = 0
    shape.Top = 0
    shape.OnAction = "PERSONAL.XLSB!PhraseMaintenancePhraseFilter"

def create_phrase_building_sheet(wb, excel):
     wb.Sheets.Add().Name = "Phrase Building"
     ws_pb = wb.Worksheets("Phrase Building")
     ws_dfr = wb.Worksheets("Documents Filed Report")

    # Copy relevant columns from Documents Filed Report
     ws_dfr.Range("O:P,AG:AH").Copy(ws_pb.Range("A1"))

    # Sort and remove duplicates
     last_row = ws_pb.Cells(ws_pb.Rows.Count, "A").End(-4162).Row
     ws_pb.Range(f"A1:D{last_row}").Sort(Key1=ws_pb.Range("C1"), Order1=2, Key2=ws_pb.Range("D1"), Order2=2, Header=1)
     ws_pb.Range(f"A1:D{last_row}").RemoveDuplicates(Columns=(1, 2, 3, 4), Header=1)

def create_filter_updates_sheet(wb, excel):
    wb.Sheets.Add().Name = "Filter Updates"
    ws = wb.Worksheets("Filter Updates")

    ws.Range("A1").Value = "Version 1.0"
    ws.Range("A9").Value = "Phrase Maintenance criteria: 1)Phrase is not 0. 2)Status is Manually Indexed. 3)Phrase Indexer Review = Yes."
    ws.Range("A10").Value = "Phrase Building criteria: 1)Phrase is 0. 2) Status is Manually Indexed."

    # Add shapes for buttons
    shape1 = ws.Shapes.AddShape(5, 0, 40, 200, 40)  # 5 is msoShapeRoundedRectangle
    shape1.TextFrame.Characters().Text = "Select this box to automatically apply the criteria used for Phrase Maintenance in the Documents Filed Report tab."
    shape1.TextFrame.Characters().Font.Size = 10
    shape1.OnAction = "PERSONAL.XLSB!PhraseMaintenanceFilters"

    shape2 = ws.Shapes.AddShape(5, 250, 40, 200, 40)
    shape2.TextFrame.Characters().Text = "Select this box to automatically apply the criteria used for Phrase Building to the Documents Filed Report tab."
    shape2.TextFrame.Characters().Font.Size = 10
    shape2.OnAction = "PERSONAL.XLSB!PhraseBuildingCriteria"

def post_analysis_formatting(wb, excel):
    # Phrase Maintenance sheet
    sheet = wb.Worksheets("Phrase Maintenance")
    print(f"Selected 'Phrase Maintenance' sheet. It has {sheet.UsedRange.Rows.Count} rows and {sheet.UsedRange.Columns.Count} columns.")

    # Store the data before making any changes
    data_range = sheet.UsedRange
    data_values = data_range.Value
    print(f"Stored existing data: {len(data_values)} rows")

    # Clear all filters
    if sheet.AutoFilterMode:
        sheet.AutoFilterMode = False

    # Apply formatting without changing data
    last_row = data_range.Rows.Count
    last_col = data_range.Columns.Count
    format_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, last_col))
    
    format_range.Interior.Pattern = -4142  # xlNone
    format_range.HorizontalAlignment = -4131  # xlGeneral
    format_range.VerticalAlignment = -4107  # xlBottom

    # Set date format for columns G and I
    sheet.Range("G:G,I:I").NumberFormat = "m/d/yyyy"

    # Set column widths
    sheet.Columns("D:D").ColumnWidth = 93.53
    sheet.Columns("B:C").ColumnWidth = 21.05
    sheet.Columns("E:F").ColumnWidth = 7.79
    sheet.Columns("J:M").ColumnWidth = 18.58

    # AutoFit rows and specific columns
    format_range.EntireRow.AutoFit()
    for col in ['B', 'C', 'E', 'F', 'G', 'I', 'J', 'K', 'L', 'M']:
        sheet.Columns(col).EntireColumn.AutoFit()

    format_range.VerticalAlignment = -4160  # xlTop


    # Restore the data
    sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, last_col)).Value = data_values
    print("Restored original data")

    # Apply filters
    if last_row > 1 and last_col >= 13:  # Check to ensure data exists and enough cols
        format_range.AutoFilter()
    else:
        print("Not enough data to apply filters")

    # Set Row Height
    sheet.Rows(1).RowHeight = 24

    format_range.WrapText = True
    format_range.Font.Color = 0
    sheet.Range("K1:R1").Interior.ColorIndex = 10

    print(f"Finished processing 'Phrase Maintenance' sheet. Final row count: {sheet.UsedRange.Rows.Count}")

    # Switch to Phrase Building sheet
    sheet = wb.Worksheets("Phrase Building")
    sheet.Activate()

    # Format cells without clearing content
    last_row = sheet.Cells(sheet.Rows.Count, "A").End(-4162).Row  # xlUp
    last_col = sheet.Cells(1, sheet.Columns.Count).End(-4159).Column  # xlToLeft
    
    format_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(last_row, last_col))
    format_range.Interior.Pattern = -4142  # xlNone

    # Set column widths
    sheet.Columns("A").ColumnWidth = 21.57
    sheet.Columns("B").ColumnWidth = 20.43
    sheet.Columns("C").ColumnWidth = 49.71
    sheet.Columns("D").ColumnWidth = 48.43
    sheet.Columns("E").ColumnWidth = 17.71

    # Add formula to column E if it doesn't exist
    if sheet.Cells(1, 5).Value != "%Patient Match":
        sheet.Cells(1, 5).Value = "%Patient Match"
        sheet.Range("E2").FormulaR1C1 = "=(RC[-2]-RC[-1])/RC[-2]"
        sheet.Range("E2").NumberFormat = "0.00%"

    last_data_row = sheet.Cells(sheet.Rows.Count, "C").End(-4162).Row  # xlUp
    if last_data_row > 2:
        fill_range = sheet.Range(f"E2:E{last_data_row}")
        fill_range.FillDown()
    # Apply filter
    format_range.AutoFilter()
    format_range.AutoFilter(Field=3, Criteria1="<>0")

    sheet.Range("A1:E1").Interior.ColorIndex = 10
    format_range.WrapText = True
    format_range.RowHeight = 15
    format_range.Font.Color = 0

filed_documents_report_with_phrase_hit_athena()

# Apply 'filter' to all table heads

# Combine column HL7 Document Type and Summary Line into one value and stacked appropriately
# IDL and HL7 Signature should both say "Unsigned" not "Not Signed"


# Improvements:

# Add tab to Phrase maintenance tab: Count where pt found and DOS changed 
