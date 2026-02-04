"""replicate_app.py

Read an Excel file and replicate rows 1..773 into 11 variants each.
For each source row, produce 11 rows with column D set to hex 01..0B and
column J replacing the first numeric occurrence with 1..11 respectively.

Usage:
    python replicate_app.py input.xlsx output.xlsx [sheet]

Defaults: input.xlsx=data.xlsx, output.xlsx=data_replicated.xlsx, sheet=Sheet1
"""
from openpyxl import load_workbook, Workbook
import re
import sys


def replicate_block(input_path: str, output_path: str, sheet_name: str = 'Sheet1',
                    start_row: int = 1, end_row: int = 773, copies: int = 11,
                    seq_col: int = 4, desc_col: int = 10):
    """Keep rows 1..block_rows as the source block, then append copies-1 new blocks.

    The original block (k=1) is preserved. For k=2..copies, append the whole block
    with column `seq_col` set to hex values 02..(copies in hex) and the first numeric
    occurrence in `desc_col` replaced with the decimal k.
    """
    wb = load_workbook(input_path, data_only=False)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Sheet '{sheet_name}' not found in {input_path}")
    ws = wb[sheet_name]

    max_col = ws.max_column

    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = ws.title

    # read source block into memory (inclusive)
    if start_row < 1:
        raise ValueError("start_row must be >= 1")
    if end_row < start_row:
        raise ValueError("end_row must be >= start_row")

    src_block = []
    for r in range(start_row, end_row + 1):
        src_block.append([ws.cell(row=r, column=c).value for c in range(1, max_col + 1)])

    total = 0
    # append original block as-is (k=1)
    for row_vals in src_block:
        ws2.append(list(row_vals))
        total += 1

    # append copies k=2..copies
    for k in range(2, copies + 1):
        hex_val = format(k, '02X')
        for row_vals in src_block:
            new_row = list(row_vals)
            # set sequence column to hex
            new_row[seq_col - 1] = hex_val
            # replace first number occurrence in desc_col with decimal k
            desc = new_row[desc_col - 1]
            if isinstance(desc, str):
                m = re.search(r"\d+", desc)
                if m:
                    s, e = m.span()
                    desc = desc[:s] + str(k) + desc[e:]
                else:
                    # if no number present, try to insert '上 k 结算日' pattern when appropriate
                    desc = desc + ' ' + str(k)
                new_row[desc_col - 1] = desc
            ws2.append(new_row)
            total += 1

    wb2.save(output_path)
    return total


if __name__ == '__main__':
    inp = sys.argv[1] if len(sys.argv) > 1 else 'data.xlsx'
    out = sys.argv[2] if len(sys.argv) > 2 else 'data_replicated.xlsx'
    sheet = sys.argv[3] if len(sys.argv) > 3 else 'Sheet1'
    # optional: start_row end_row copies
    try:
        start_row = int(sys.argv[4]) if len(sys.argv) > 4 else 1
    except ValueError:
        print('start_row must be an integer')
        raise
    try:
        end_row = int(sys.argv[5]) if len(sys.argv) > 5 else 773
    except ValueError:
        print('end_row must be an integer')
        raise
    try:
        copies = int(sys.argv[6]) if len(sys.argv) > 6 else 11
    except ValueError:
        print('copies must be an integer')
        raise

    total = replicate_block(inp, out, sheet_name=sheet, start_row=start_row, end_row=end_row, copies=copies)
    expected = (end_row - start_row + 1) * copies
    print(f"Wrote {total} rows to {out} (expected {expected})")
