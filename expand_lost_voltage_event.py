import pandas as pd

def expand_data(input_file='data.xlsx', output_file='expanded_data.xlsx'):
    # Read the Excel file without headers
    df = pd.read_excel(input_file, header=None)
    
    # First row is headers
    headers = df.iloc[0]
    
    # Second row is the data template
    data_row = df.iloc[1]
    
    # Phases: A, B, C with codes 01, 02, 03
    phases = ['A', 'B', 'C']
    phase_codes = {'A': '01', 'B': '02', 'C': '03'}
    
    # Record indices: 1 to 10, hex formatted 01 to 0A
    records = range(1, 11)
    hex_records = [f'{i:02X}' for i in records]
    
    new_data = []
    
    for phase in phases:
        for i, hex_rec in enumerate(hex_records, 1):
            # Copy the data row
            new_row = data_row.copy()
            
            # Column C (index 2): phase code
            new_row.iloc[2] = phase_codes[phase]
            
            # Column D (index 3): record index in hex
            new_row.iloc[3] = hex_rec
            
            # Column J (index 9): modify the text
            original_j = str(new_row.iloc[9])
            new_j = original_j.replace('（上1次）A相失压记录：', f'（上{i}次）{phase}相失压记录：')
            new_row.iloc[9] = new_j
            
            new_data.append(new_row)
    
    # Create DataFrame for new data
    new_df = pd.DataFrame(new_data)
    
    # Combine headers and new data
    full_df = pd.concat([headers.to_frame().T, new_df], ignore_index=True)
    
    # Write to new Excel file
    full_df.to_excel(output_file, index=False, header=False)
    
    print(f"Expanded data written to {output_file}")

if __name__ == "__main__":
    expand_data()