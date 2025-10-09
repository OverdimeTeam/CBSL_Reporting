import os
from pathlib import Path
import openpyxl
import sys

def get_month_year_from_filename(filename):
    # Example: "NBD-MF-20-C1 to C6 July 2025.xlsx"
    parts = filename.split()
    for i, part in enumerate(parts):
        if part in {"Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec","January","February","March","April","May","June","July","August","September","October","November","December"}:
            month = part
            year = parts[i+1].replace(".xlsx","")
            return month, year
    return None, None

def find_files_safely(base_dir, pattern, description):
    """Safely find files and provide detailed error messages"""
    print(f"Looking for {description} in: {base_dir}")
    print(f"Search pattern: {pattern}")
    
    files = list(base_dir.glob(pattern))
    print(f"Found {len(files)} files matching pattern:")
    
    if not files:
        print(f"No files found matching pattern: {pattern}")
        print("Available files in directory:")
        try:
            all_files = [f.name for f in base_dir.iterdir() if f.is_file() and f.suffix == '.xlsx']
            if all_files:
                for file in sorted(all_files):
                    print(f"  - {file}")
            else:
                print("  No .xlsx files found in directory")
        except Exception as e:
            print(f"  Error listing files: {e}")
        return None
    
    for file in files:
        print(f"  {file.name}")
    
    return files[0]  # Return the first match

def main():
    # Get the base directory (parent of report_automations)
    base_dir = Path(__file__).resolve().parent.parent
    
    # Based on your file structure, input files are in working/monthly/09-19-2025(1)/
    working_dir = base_dir / "working"
    monthly_working_dir = working_dir / "monthly"
    outputs_monthly_dir = base_dir / "outputs" / "monthly"
    
    print(f"Base directory: {base_dir}")
    print(f"Working directory: {working_dir}")
    print(f"Monthly working directory: {monthly_working_dir}")
    print(f"Current working directory: {Path.cwd()}")
    
    # Create outputs monthly directory if it doesn't exist
    outputs_monthly_dir.mkdir(parents=True, exist_ok=True)
    
    # Search in the working/monthly directory first
    search_dirs = [
        monthly_working_dir,
        working_dir,
        base_dir
    ]
    
    file_c1_c6 = None
    
    # Search for C1-C6 file in multiple locations
    for search_dir in search_dirs:
        if search_dir.exists():
            print(f"\nSearching in: {search_dir}")
            
            # Look for the specific file you have
            c1_c6_pattern = "**/NBD-MF-20-C1 to C6*.xlsx"
            files = list(search_dir.glob(c1_c6_pattern))
            
            if files:
                file_c1_c6 = files[0]
                print(f"Found C1-C6 file: {file_c1_c6}")
                break
    
    if file_c1_c6 is None:
        print("\nCould not find C1 to C6 file. Looking for any similar files...")
        for search_dir in search_dirs:
            if search_dir.exists():
                xlsx_files = list(search_dir.rglob("*.xlsx"))
                if xlsx_files:
                    print(f"Excel files in {search_dir}:")
                    for f in xlsx_files:
                        print(f"  - {f.relative_to(base_dir)}")
        return
    
    # Extract month and year from filename
    month, year = get_month_year_from_filename(file_c1_c6.name)
    if not month or not year:
        print(f"Could not parse month/year from filename: {file_c1_c6.name}")
        print("Expected format: 'NBD-MF-20-C1 to C6 [Month] [Year].xlsx'")
        return
    
    print(f"\nParsed month: {month}, year: {year}")
    
    # Find the exact working folder structure to replicate in outputs
    working_folder_name = None
    for search_dir in search_dirs:
        if search_dir.exists():
            # Look for the folder containing our files
            for subfolder in search_dir.iterdir():
                if subfolder.is_dir():
                    # Check if this folder contains our target files
                    c1_c6_files = list(subfolder.glob("**/NBD-MF-20-C1 to C6*.xlsx"))
                    if c1_c6_files:
                        working_folder_name = subfolder.name
                        print(f"Found working folder: {working_folder_name}")
                        break
            if working_folder_name:
                break
    
    # Create output folder structure matching working folder
    if working_folder_name:
        out_folder = outputs_monthly_dir / working_folder_name
    else:
        # Fallback to date-based folder if working folder not found
        out_folder = outputs_monthly_dir / f"{year}_{month}"
    
    out_folder.mkdir(parents=True, exist_ok=True)
    print(f"Output folder: {out_folder}")
    
    try:
        # Load workbook
        print(f"\nLoading workbook...")
        wb_c1_c6 = openpyxl.load_workbook(file_c1_c6, data_only=True)
        print(f"Workbook loaded successfully")
        
        # Check if required sheets exist
        print(f"\nAvailable sheets in C1-C6 file: {wb_c1_c6.sheetnames}")
        
        # Check for C5 sheet in C1-C6 file
        if "NBD_NBL-MF-20-C5" not in wb_c1_c6.sheetnames:
            print(f"Required sheet 'NBD_NBL-MF-20-C5' not found in C1-C6 file")
            print(f"Available sheets: {wb_c1_c6.sheetnames}")
            return
        
        print(f"Required sheet found")
        
        # Get worksheet
        tgt_ws = wb_c1_c6["NBD_NBL-MF-20-C5"]
        
        print(f"\nProcessing data...")
        
        # Helper function to safely get numeric values
        def get_numeric_value(worksheet, cell_ref, description=""):
            """Safely extract numeric value from a cell, handling text/None values"""
            try:
                value = worksheet[cell_ref].value
                print(f"{description} ({cell_ref}): {value}")
                
                if value is None:
                    return 0
                elif isinstance(value, (int, float)):
                    return value
                elif isinstance(value, str):
                    # Try to extract number from string if possible
                    try:
                        # Remove common text and try to parse number
                        clean_value = value.replace(',', '').strip()
                        return float(clean_value)
                    except ValueError:
                        print(f"  Warning: Cell {cell_ref} contains text '{value}' - using 0")
                        return 0
                else:
                    print(f"  Warning: Cell {cell_ref} contains unexpected type {type(value)} - using 0")
                    return 0
            except Exception as e:
                print(f"  Error reading {cell_ref}: {e} - using 0")
                return 0
        
        # Hardcoded value for "Guarantees" (dummy data from work hub integration)
        print(f"\nSetting hardcoded value for 'Guarantees'...")
        guarantees_hardcoded_value = 2500000  # Hardcoded dummy value
        print(f"DUMMY DATA: Using hardcoded value {guarantees_hardcoded_value} for 'Guarantees'")
        print(f"NOTE: This value comes from work hub integration as master data (dummy purpose)")
        
        # Find "Guarantees" row in C5 sheet
        print(f"\nSearching for 'Guarantees' in C5 sheet...")
        
        guarantees_row = None
        for row in range(1, tgt_ws.max_row + 1):
            for col in range(1, tgt_ws.max_column + 1):
                cell_value = tgt_ws.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str) and "Guarantees" in cell_value:
                    print(f"Found 'Guarantees' at row {row}, column {col}")
                    guarantees_row = row
                    break
            if guarantees_row:
                break
        
        if guarantees_row is None:
            print("Could not find 'Guarantees' row in C5 sheet")
            print("Available cell values in C5 sheet:")
            for row in range(1, min(20, tgt_ws.max_row + 1)):  # Show first 20 rows
                for col in range(1, min(10, tgt_ws.max_column + 1)):  # Show first 10 columns
                    cell_value = tgt_ws.cell(row=row, column=col).value
                    if cell_value:
                        print(f"  Row {row}, Col {col}: {cell_value}")
            return
        
        # Update the Guarantees value
        # Assuming the value goes in column C (3rd column)
        guarantees_cell = f"C{guarantees_row}"
        print(f"Updating cell {guarantees_cell} with hardcoded value: {guarantees_hardcoded_value}")
        tgt_ws[guarantees_cell] = guarantees_hardcoded_value
        
        print(f"Data processing completed")
        print(f"Updated C5 sheet:")
        print(f"  - Guarantees: {guarantees_hardcoded_value} (HARDCODED DUMMY DATA)")
        
        # Save updated workbook
        out_path = out_folder / f"NBD_NBL_MF_20_C5_{month}_{year}_report.xlsx"
        wb_c1_c6.save(out_path)
        print(f"\nReport saved successfully to: {out_path}")
        
    except FileNotFoundError as e:
        print(f"File not found: {e}")
    except PermissionError as e:
        print(f"Permission error (file might be open in Excel): {e}")
    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='NBD-MF-20-C5 Report Automation')
    parser.add_argument('--working-dir', type=str, help='Working directory (not used, for compatibility)')
    parser.add_argument('--month', type=str, help='Report month (e.g., Jul)')
    parser.add_argument('--year', type=str, help='Report year (e.g., 2025)')
    args = parser.parse_args()

    main()
