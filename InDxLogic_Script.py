import openpyxl
from openpyxl.styles import PatternFill

def copy_formula_to_range(ws, formula, start_cell, end_cell):
    for row in ws[start_cell:end_cell]:
        for cell in row:
            cell.value = formula

# Load workbook and select the necessary sheets
wb = openpyxl.load_workbook('test_sheet.xlsx')
doc_report_ws = wb['Documents Filed Report']
phrase_hit_ws = wb['Phrase Hit Report']

# Insert columns at the beginning
for _ in range(6):
    doc_report_ws.insert_cols(1)

# Rename columns
header_titles = ["Member Match", "Summary Match", "DOS Match", "Signature Match", "Patient Match", "Provider Match"]
for col_num, title in enumerate(header_titles, 1):
    doc_report_ws.cell(row=1, column=col_num).value = title

# Set formulas and fill down
row_count = doc_report_ws.max_row

formulas = [
    ('A2', 'IF(L2=N2, "EXACTMATCH", "NEEDSREVIEW")'),
    ('B2', 'IF(M2=O2, "EXACTMATCH", "NEEDSREVIEW")'),
    ('C2', 'IF(P2=Q2, "EXACTMATCH", "NEEDSREVIEW")'),
    ('D2', 'IF(V2=W2, "EXACTMATCH", "NEEDSREVIEW")'),
    ('E2', 'IF(COUNTIF(Y2,"*oun*"), "NOTFOUND", IF(H2=I2,"EXACTMATCH","NEEDSREVIEW"))'),
    ('F2', 'IF(AND(E2="NOTFOUND",R2=S2),"PTNFEXACTMATCH",IF(AND(E2="NOTFOUND",R2<>S2),"PTNFNEEDSREVIEW",IF(AND(E2="EXACTMATCH",R2=S2),"EXACTMATCH",IF(AND(E2="NEEDSREVIEW",R2=S2),"EXACTMATCH","NEEDSREVIEW"))))')
]

for start_cell, formula in formulas:
    end_cell = f'{start_cell[0]}{row_count}'
    copy_formula_to_range(doc_report_ws, f'={formula}', start_cell, end_cell)

# Clear formulas and copy values
for col_num in range(1, 7):
    for row in range(2, row_count + 1):
        cell = doc_report_ws.cell(row=row, column=col_num)
        cell.value = cell.value

# Add Indexer Review to end of document and paste with no values
doc_report_ws['AE1'].value = "Indexer Review"
for row in range(2, row_count + 1):
    doc_report_ws[f'AE{row}'] = f'=VLOOKUP(K{row}, \'Phrase Hit Report\'!A$2:I$30000, 5, FALSE)'
for row in range(1, row_count + 1):
    cell = doc_report_ws[f'AE{row}']
    cell.value = cell.value

# Add additional columns with specific formulas and copy values
additional_columns = [
    ('AF', "Documents Manually Indexed with No Phrase by HL7 Document Type and HL7 Summary Line", 
     'COUNTIFS(N:N,N:N, O:O,O:O, X:X, "Manually Indexed", K:K,"=0")'),
    ('AG', "Indexed Documents with Flag containing No Patient Found and No Phrase Hit by HL7 Document Type and HL7 Summary Line", 
     'COUNTIFS(N:N,N:N,O:O,O:O,K:K,"=0",X:X,"Manually Indexed",Y:Y,"*oun*")')
]

for col, header, formula in additional_columns:
    doc_report_ws[f'{col}1'].value = header
    for row in range(2, row_count + 1):
        doc_report_ws[f'{col}{row}'] = f'={formula}'
    for row in range(1, row_count + 1):
        cell = doc_report_ws[f'{col}{row}']
        cell.value = cell.value

# Format all dates as mm/dd/yyyy
date_columns = ['P', 'Q', 'AA', 'AB', 'AC']
for col in date_columns:
    for cell in doc_report_ws[col]:
        cell.number_format = 'mm/dd/yyyy'

# Set all rows to no fill
for row in doc_report_ws.iter_rows(min_row=2, max_row=row_count, min_col=1, max_col=33):
    for cell in row:
        cell.fill = PatternFill(fill_type=None)

# Freeze pane on header row
doc_report_ws.freeze_panes = doc_report_ws['A2']

# Auto Fit Rows and set special color for added columns
for col in doc_report_ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        if cell.coordinate in doc_report_ws.merged_cells: # not check merge_cells
            continue
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    doc_report_ws.column_dimensions[column].width = adjusted_width

# Set special color for added columns
fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
for col in range(1, 8):
    doc_report_ws.cell(row=1, column=col).fill = fill
doc_report_ws['AE1'].fill = fill
doc_report_ws['AF1'].fill = fill
doc_report_ws['AG1'].fill = fill

# Insert, rename, create hyperlinks, and match formatting, then remove formulas
doc_report_ws.insert_cols(7)
doc_report_ws['G1'].value = "DocDetails"
for row in range(2, row_count + 1):
    doc_report_ws[f'G{row}'] = f'=HYPERLINK("https://core.indxlogic.com/Document/ViewDetail/"&H{row})'
for row in range(1, row_count + 1):
    cell = doc_report_ws[f'G{row}']
    cell.value = cell.value

# Filter for Manually Indexed Only and Indexer Review is NO
doc_report_ws.auto_filter.ref = 'A1:AG1'
doc_report_ws.auto_filter.add_filter_column(24, ['Manually Indexed'])
doc_report_ws.auto_filter.add_filter_column(11, ['<>0'])
doc_report_ws.auto_filter.add_filter_column(31, ['Yes'])

# Add Phrase Maintenance Sheet
wb.create_sheet('Phrase Maintenance')
phrase_maintenance_ws = wb['Phrase Maintenance']

# Copy Phrase Hit Report to Phrase Maintenance
for col in range(1, 11):
    for row in range(1, phrase_hit_ws.max_row + 1):
        phrase_maintenance_ws.cell(row=row, column=col).value = phrase_hit_ws.cell(row=row, column=col).value

# Rename columns and set formulas for Phrase Maintenance
maintenance_headers = [
    "Total Hits in Reporting Period (Indexed)", "Count DT Changed", "Count Summary Changed",
    "Count DOS Changed", "Count Signature Changed", "Count Patient Found and Changed",
    "Count Provider Changed where Patient Found"
]
for col_num, title in enumerate(maintenance_headers, 10):
    phrase_maintenance_ws.cell(row=1, column=col_num).value = title

# Add formulas for Phrase Maintenance and copy values
maintenance_formulas = [
    ('J', 'COUNTIFS(\'Documents Filed Report\'!L:L, A2, \'Documents Filed Report\'!Y:Y,"Manually Indexed", \'Documents Filed Report\'!AF:AF,"Yes")'),
    ('K', 'COUNTIFS(\'Documents Filed Report\'!L:L, A2, \'Documents Filed Report\'!Y:Y,"Manually Indexed", \'Documents Filed Report\'!AF:AF,"Yes", \'Documents Filed Report\'!A:A,"NEEDSREVIEW")'),
    ('L', 'COUNTIFS(\'Documents Filed Report\'!L:L, A2,\'Documents Filed Report\'!Y:Y,"Manually Indexed", \'Documents Filed Report\'!AF:AF,"Yes", \'Documents Filed Report\'!B:B,"NEEDSREVIEW")'),
    ('M', 'COUNTIFS(\'Documents Filed Report\'!L:L, A2, \'Documents Filed Report\'!Y:Y,"Manually Indexed", \'Documents Filed Report\'!AF:AF,"Yes", \'Documents Filed Report\'!C:C,"NEEDSREVIEW")'),
    ('N', 'COUNTIFS(\'Documents Filed Report\'!L:L, A2, \'Documents Filed Report\'!Y:Y,"Manually Indexed", \'Documents Filed Report\'!AF:AF,"Yes", \'Documents Filed Report\'!D:D,"NEEDSREVIEW")'),
    ('O', 'COUNTIFS(\'Documents Filed Report\'!L:L, A2, \'Documents Filed Report\'!Y:Y,"Manually Indexed", \'Documents Filed Report\'!AF:AF,"Yes", \'Documents Filed Report\'!E:E,"NEEDSREVIEW")'),
    ('P', 'COUNTIFS(\'Documents Filed Report\'!L:L, A2, \'Documents Filed Report\'!Y:Y,"Manually Indexed", \'Documents Filed Report\'!AF:AF,"Yes",  \'Documents Filed Report\'!E:E, "<>NOTFOUND",  \'Documents Filed Report\'!F:F,"NEEDSREVIEW")')
]

for col, formula in maintenance_formulas:
    for row in range(2, row_count + 1):
        phrase_maintenance_ws[f'{col}{row}'] = f'={formula.replace("2", str(row))}'
    for row in range(1, row_count + 1):
        cell = phrase_maintenance_ws[f'{col}{row}']
        cell.value = cell.value

# Filter for Indexer Review = YES
phrase_maintenance_ws.auto_filter.ref = 'A1:P1'
phrase_maintenance_ws.auto_filter.add_filter_column(4, ['Yes'])

# Add and name new sheets
wb.create_sheet('Phrase Building')
phrase_building_ws = wb['Phrase Building']

# Copy data to new sheet and remove duplicates
cols_to_copy = [('O', 'A'), ('P', 'B'), ('AG', 'C'), ('AH', 'D')]
for src_col, dest_col in cols_to_copy:
    for row in range(1, row_count + 1):
        phrase_building_ws[f'{dest_col}{row}'].value = doc_report_ws[f'{src_col}{row}'].value

# Remove duplicates
data = list(phrase_building_ws.iter_rows(min_row=1, max_col=4, values_only=True))
unique_data = [list(t) for t in set(tuple(element) for element in data)]
for row, row_data in enumerate(unique_data, start=1):
    for col, cell_value in enumerate(row_data, start=1):
        phrase_building_ws.cell(row=row, column=col).value = cell_value

# Sorts table by Summary Count, Summary Line, and Phrase ID
phrase_building_ws.auto_filter.ref = 'A1:D1'
phrase_building_ws.auto_filter.add_sort_condition(f'C2:C{row_count}', descending=True)
phrase_building_ws.auto_filter.add_sort_condition(f'D2:D{row_count}', descending=True)

# Auto Fit Rows and set special color for added columns
for col in phrase_building_ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    phrase_building_ws.column_dimensions[column].width = adjusted_width

fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
phrase_building_ws['C1'].fill = fill
phrase_building_ws['D1'].fill = fill

# Save workbook
wb.save('your_updated_workbook.xlsx')