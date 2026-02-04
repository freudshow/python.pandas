from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy

def split_multiline_rows_in_sheet(file_path: str, sheet_name: str, output_path: str = None):
    """
    Split rows with multi-line cells into multiple rows.
    
    :param file_path: Path to the Excel file
    :param sheet_name: Name of the sheet to process
    :param output_path: Path to save the output file (defaults to file_path)
    """
    if output_path is None:
        output_path = file_path
    
    wb = load_workbook(file_path)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {wb.sheetnames}")
    
    ws = wb[sheet_name]
    
    # Find all rows with multi-line cells
    rows_to_process = []
    
    for row_idx in range(1, ws.max_row + 1):
        has_multiline = False
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value and isinstance(cell.value, str) and '\n' in cell.value:
                has_multiline = True
                break
        
        if has_multiline:
            rows_to_process.append(row_idx)
    
    print(f"Found {len(rows_to_process)} rows with multi-line cells: {rows_to_process}")
    
    # Process rows in reverse order (from bottom to top) to avoid row number shifts
    for original_row in reversed(rows_to_process):
        # Get all cell values and split by newlines
        max_lines = 1
        split_values = {}
        
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=original_row, column=col_idx)
            if cell.value and isinstance(cell.value, str) and '\n' in cell.value:
                lines = cell.value.split('\n')
                split_values[col_idx] = lines
                max_lines = max(max_lines, len(lines))
            else:
                split_values[col_idx] = [cell.value] * max_lines
        
        # Fill non-split cells with repetitions
        for col_idx in range(1, ws.max_column + 1):
            if col_idx not in split_values:
                cell = ws.cell(row=original_row, column=col_idx)
                split_values[col_idx] = [cell.value] * max_lines
        
        # Insert new rows below the original row
        if max_lines > 1:
            ws.insert_rows(original_row + 1, max_lines - 1)
            
            # Fill the split data back
            for line_idx in range(max_lines):
                for col_idx in range(1, ws.max_column + 1):
                    target_row = original_row + line_idx
                    new_cell = ws.cell(row=target_row, column=col_idx)
                    
                    # Get source cell for copying style
                    source_cell = ws.cell(row=original_row, column=col_idx)
                    
                    # Copy style
                    if source_cell.has_style:
                        new_cell.font = copy(source_cell.font)
                        new_cell.border = copy(source_cell.border)
                        new_cell.fill = copy(source_cell.fill)
                        new_cell.number_format = copy(source_cell.number_format)
                        new_cell.protection = copy(source_cell.protection)
                        new_cell.alignment = copy(source_cell.alignment)
                    
                    # Set value
                    value = split_values[col_idx][line_idx] if line_idx < len(split_values[col_idx]) else None
                    new_cell.value = value
    
    wb.save(output_path)
    print(f"File saved to: {output_path}")

if __name__ == '__main__':
    # Example usage
    file_path = 'data.xlsx'
    sheet_name = 'Sheet1'
    
    split_multiline_rows_in_sheet(file_path, sheet_name)
    
    print("\n--- After splitting ---")
    # Verify the result
    from openpyxl import load_workbook
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    
    # Show a sample of rows that had multi-line cells
    print(f"Total rows in {sheet_name}: {ws.max_row}")
    print("\nSample of processed rows (showing row 100 area):")
    for row in range(98, min(110, ws.max_row + 1)):
        values = [ws.cell(row=row, column=col).value for col in range(1, 11)]
        print(f"Row {row}: {values[:5]}... (showing first 5 cols)")
