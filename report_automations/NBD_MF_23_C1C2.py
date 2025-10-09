import os
import pandas as pd
import xlwings as xw
import re
import time
import subprocess
from datetime import datetime
import sys

# Fix Unicode encoding for Windows console
if sys.platform.startswith('win'):
    import codecs
    sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'replace')
    sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'replace')

# ==================== UTILITY FUNCTIONS ====================

def kill_excel_processes():
    """Kill any hanging Excel processes"""
    try:
        subprocess.call(['taskkill', '/F', '/IM', 'EXCEL.EXE'], 
                       stdout=subprocess.DEVNULL, 
                       stderr=subprocess.DEVNULL)
        time.sleep(2)
    except:
        pass

def log_step(step_name, start_time):
    """Log step completion with time taken"""
    elapsed = time.time() - start_time
    print(f"  ✓ Completed in {elapsed:.2f} seconds")
    return time.time()

def get_column_letter(col_num):
    """Convert column number to Excel column letter"""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result

def column_letter_to_index(col_letter):
    """Convert Excel column letter to 0-based index"""
    col_letter = col_letter.upper()
    result = 0
    for i, char in enumerate(reversed(col_letter)):
        result += (ord(char) - ord('A') + 1) * (26 ** i)
    return result - 1

def get_first_sheet_name(file_path):
    """Get the name of the first sheet in an Excel file"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(file_path)
            sheet_name = wb.sheets[0].name
            wb.close()
            app.quit()
            return sheet_name
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def is_valid_contract_number(value):
    """Check if value matches valid contract number patterns"""
    if not value or not isinstance(value, str):
        return False
    
    value_clean = str(value).strip().upper()
    
    if len(value_clean) < 5:
        return False
    
    if re.match(r'^LR\d{8}$', value_clean):
        return True
    
    if re.match(r'^[A-Z]{4}\d{9}$', value_clean):
        return True
    
    if re.match(r'^[A-Z]{2,3}\d{8,10}$', value_clean):
        return True
    
    if len(value_clean) >= 8 and any(c.isalpha() for c in value_clean) and any(c.isdigit() for c in value_clean):
        header_keywords = ['CONTRACT', 'NUMBER', 'NO', 'LOAN', 'YARD', 'PROPERTY', 'MORTGAGE', 'LIST', 'TOTAL', 'DATE']
        if not any(keyword in value_clean for keyword in header_keywords):
            return True
    
    return False

def is_valid_value(value):
    """Check if value is valid (not N/A, 0, -, or empty)"""
    if pd.isna(value) or value is None or value == '':
        return False
    
    value_str = str(value).strip().upper()
    
    if value_str in ['N/A', 'NA', '#N/A', 'NULL', 'NONE', 'NAN']:
        return False
    
    try:
        num_value = float(str(value).replace(',', ''))
        return num_value != 0
    except:
        return False

# ==================== FILE OPERATIONS ====================

def find_dynamic_folder(base_path):
    """Find the dynamic date folder inside monthly directory"""
    monthly_path = os.path.join(base_path, 'working', 'monthly')
    
    if not os.path.exists(monthly_path):
        print(f"ERROR: Monthly folder not found at {monthly_path}")
        return None
    
    folders = [f for f in os.listdir(monthly_path) if os.path.isdir(os.path.join(monthly_path, f))]
    
    if not folders:
        print("ERROR: No folders found in monthly directory")
        return None
    
    if len(folders) > 1:
        print(f"WARNING: Multiple folders found. Using: {folders[0]}")
    
    target_folder = os.path.join(monthly_path, folders[0], 'NBD_MF_23_IA')
    
    if os.path.exists(target_folder):
        print(f"Found folder: {folders[0]}/NBD_MF_23_IA")
        return target_folder
    
    print(f"ERROR: NBD_MF_23_IA subfolder not found")
    return None

def find_files(folder_path):
    """Find all required files using keywords"""
    files = {
        'summary': None, 'net': None, 'prod': None, 'loan_base': None,
        'reschedule': None, 'yard_stock': None, 'cbsl_provision': None, 
        'cadre': None, 'unutilized': None, 'po_listing': None, 'portfolio_recovery': None
    }
    keywords = {
        'summary': 'Summary', 'net': 'Net Portfolio', 'prod': 'Prod. wise Class. of Loans',
        'loan_base': 'Loan Base', 'reschedule': 'Reschedule Contract Details',
        'yard_stock': 'YARD STOCK', 'cbsl_provision': 'CBSL Provision Comparison',
        'cadre': 'Cadre', 'unutilized': 'Unutilized Amount',
        'po_listing': 'Po Listing - Internal', 'portfolio_recovery': 'Portfolio Report Recovery - Internal'
    }
    
    for root, _, file_list in os.walk(folder_path):
        for f in file_list:
            if f.startswith("~$") or not f.endswith((".xlsb", ".xlsx")):
                continue
            for key, keyword in keywords.items():
                if keyword in f and files[key] is None:
                    files[key] = os.path.join(root, f)
    
    return (files['summary'], files['net'], files['prod'], files['loan_base'],
            files['reschedule'], files['yard_stock'], files['cbsl_provision'], 
            files['cadre'], files['unutilized'], files['po_listing'], files['portfolio_recovery'])

# ==================== EXCEPTION FILE ====================

class ExceptionTracker:
    """Track exceptions and write to a single file"""
    
    def __init__(self, folder_path):
        self.folder_path = folder_path
        self.exception_file = os.path.join(folder_path, 'EXCEPTIONS.xlsx')
        self.exceptions = {}
    
    def add_exception(self, sheet_name, data_dict):
        """Add exception data for a specific sheet"""
        if sheet_name not in self.exceptions:
            self.exceptions[sheet_name] = []
        self.exceptions[sheet_name].append(data_dict)
    
    def write_exceptions(self):
        """Write all exceptions to a single Excel file"""
        if not self.exceptions:
            return
        
        with pd.ExcelWriter(self.exception_file, engine='openpyxl') as writer:
            for sheet_name, exception_list in self.exceptions.items():
                df = pd.DataFrame(exception_list)
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        print(f"  Exception file: {os.path.basename(self.exception_file)}")

# ==================== DATA READING ====================

def read_contracts_fast(file_path, sheet_name, column_letter):
    """Fast contract reading for large files"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            if file_path.endswith('.xlsb'):
                app = xw.App(visible=False)
                app.display_alerts = False
                wb = app.books.open(file_path)
                sheet = wb.sheets[sheet_name]
                last_row = sheet.used_range.last_cell.row
                
                chunk_size = 50000
                all_data = []
                
                for start in range(1, last_row + 1, chunk_size):
                    end = min(start + chunk_size - 1, last_row)
                    chunk = sheet.range(f'{column_letter}{start}:{column_letter}{end}').value
                    all_data.extend(chunk if isinstance(chunk, list) else [chunk])
                
                wb.close()
                app.quit()
                contracts = all_data
            else:
                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                col_idx = ord(column_letter.upper()) - ord('A')
                contracts = df.iloc[:, col_idx].tolist()
            
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise
    
    header_values = {'CONTRACT_NO', 'CONTRACT NO', 'CONTRACT NUMBER'}
    contracts_cleaned = [
        str(c).strip() for c in contracts 
        if c and str(c).strip() and str(c).strip().upper() not in header_values
    ]
    
    return contracts_cleaned

def classify_contracts(contracts, pattern):
    """Classify contracts by regex pattern"""
    compiled_pattern = re.compile(pattern, re.IGNORECASE)
    return [c for c in contracts if compiled_pattern.match(c)]

def read_net_portfolio_data(net_file, matching_contracts):
    """Read and format Net Portfolio data for matched contracts"""
    usecols = [2, 4, 5, 7, 20, 22, 23, 27, 33, 34, 38]
    df = pd.read_excel(net_file, sheet_name=0, header=None, usecols=usecols)
    
    df.columns = ['CLIENT_CODE', 'CONTRACT_NO', 'CONTRACT_PERIOD', 'CONTRACT_AMOUNT',
                  'GROSS_PORTFOLIO', 'DP_PROVISION', 'PROVISION', 'EQT_DESC', 
                  'CON_RNTFREQ', 'CON_INTRATE', 'PURPOSE']
    
    df['CONTRACT_NO'] = df['CONTRACT_NO'].astype(str).str.strip().str.upper()
    df = df[df['CONTRACT_NO'].isin(set(matching_contracts))].copy()
    
    for col in ['GROSS_PORTFOLIO', 'PROVISION', 'DP_PROVISION', 'CONTRACT_AMOUNT']:
        df[col] = df[col].apply(lambda x: f"{float(x):,.0f}" if pd.notna(x) and x != '' else '')
    
    df['CON_INTRATE'] = df['CON_INTRATE'].apply(
        lambda x: f"{float(x):.2f}%" if pd.notna(x) and x != '' else ''
    )
    
    for col in ['CLIENT_CODE', 'EQT_DESC', 'PURPOSE', 'CON_RNTFREQ', 'CONTRACT_PERIOD']:
        df[col] = df[col].fillna('').astype(str).str.strip()
    
    return df

def read_multiple_columns_from_file(file_path, sheet_name, lookup_col, data_cols, contracts):
    """Read multiple columns at once from a file for given contracts"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(file_path)
            
            if isinstance(sheet_name, int):
                sheet = wb.sheets[sheet_name]
            else:
                sheet = None
                for s in wb.sheets:
                    if sheet_name.upper() in s.name.upper():
                        sheet = s
                        break
                if not sheet:
                    sheet = wb.sheets[0]
            
            last_row = sheet.used_range.last_cell.row
            
            lookup_data = sheet.range(f'{lookup_col}1:{lookup_col}{last_row}').value
            lookup_list = lookup_data if isinstance(lookup_data, list) else [lookup_data]
            
            result_data = {}
            for col in data_cols:
                col_data = sheet.range(f'{col}1:{col}{last_row}').value
                col_list = col_data if isinstance(col_data, list) else [col_data]
                result_data[col] = col_list
            
            wb.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise
    
    contract_data_map = {}
    for i, contract in enumerate(lookup_list):
        if contract:
            contract_clean = str(contract).strip().upper()
            contract_data_map[contract_clean] = {col: result_data[col][i] for col in data_cols}
    
    result = {col: [] for col in data_cols}
    for contract in contracts:
        contract_clean = str(contract).strip().upper()
        if contract_clean in contract_data_map:
            for col in data_cols:
                result[col].append(contract_data_map[contract_clean][col])
        else:
            for col in data_cols:
                result[col].append('')
    
    return result

def read_columns_from_large_file(file_path, lookup_col_letter, data_col_letters, contracts):
    """Read columns from large Excel files using pandas"""
    if not os.path.exists(file_path):
        raise Exception(f"File not found: {file_path}")
    
    try:
        lookup_col_idx = column_letter_to_index(lookup_col_letter)
        data_col_indices = [column_letter_to_index(col) for col in data_col_letters]
        
        usecols = [lookup_col_idx] + data_col_indices
        df = pd.read_excel(file_path, sheet_name=0, header=None, usecols=usecols)
        
        df.columns = ['LOOKUP'] + [f'DATA_{i}' for i in range(len(data_col_letters))]
        df['LOOKUP'] = df['LOOKUP'].astype(str).str.strip().str.upper()
        
        lookup_dict = {}
        for idx, row in df.iterrows():
            contract = row['LOOKUP']
            if contract and contract != 'NAN':
                lookup_dict[contract] = {data_col_letters[i]: row[f'DATA_{i}'] 
                                        for i in range(len(data_col_letters))}
        
        result = {col: [] for col in data_col_letters}
        for contract in contracts:
            contract_clean = str(contract).strip().upper()
            if contract_clean in lookup_dict:
                for col in data_col_letters:
                    result[col].append(lookup_dict[contract_clean][col])
            else:
                for col in data_col_letters:
                    result[col].append(None)
        
        return result
        
    except Exception as e:
        raise Exception(f"Failed to read file with pandas: {str(e)}")

def calculate_contract_periods(loan_base_file, loan_base_sheet, bulk1_contracts):
    """Calculate contract periods using DATEDIF between Loan Date and Mat Date"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(loan_base_file)
            sheet = wb.sheets[loan_base_sheet]
            
            last_row = sheet.used_range.last_cell.row
            
            contracts_col = sheet.range(f'B1:B{last_row}').value
            loan_dates_col = sheet.range(f'E1:E{last_row}').value
            mat_dates_col = sheet.range(f'F1:F{last_row}').value
            
            wb.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise
    
    contracts_list = contracts_col if isinstance(contracts_col, list) else [contracts_col]
    loan_dates_list = loan_dates_col if isinstance(loan_dates_col, list) else [loan_dates_col]
    mat_dates_list = mat_dates_col if isinstance(mat_dates_col, list) else [mat_dates_col]
    
    contract_period_map = {}
    for contract, loan_date, mat_date in zip(contracts_list, loan_dates_list, mat_dates_list):
        if contract and loan_date and mat_date:
            contract_clean = str(contract).strip().upper()
            try:
                if isinstance(loan_date, str):
                    loan_date = pd.to_datetime(loan_date)
                if isinstance(mat_date, str):
                    mat_date = pd.to_datetime(mat_date)
                
                months_diff = (mat_date.year - loan_date.year) * 12 + (mat_date.month - loan_date.month)
                contract_period_map[contract_clean] = months_diff
            except:
                contract_period_map[contract_clean] = ''
    
    periods = []
    for contract in bulk1_contracts:
        contract_clean = str(contract).strip().upper()
        period = contract_period_map.get(contract_clean, '')
        periods.append(period if period != '' else '')
    
    return periods

def calculate_initial_valuation_bulk2(contracts, po_listing_file, portfolio_recovery_file, prod_file, exception_tracker):
    """Calculate Initial Valuation for Bulk 2 with 3-tier fallback logic"""
    initial_valuations = [None] * len(contracts)
    
    if po_listing_file and os.path.exists(po_listing_file):
        try:
            po_data = read_columns_from_large_file(po_listing_file, 'H', ['AH'], contracts)
            for i, val in enumerate(po_data['AH']):
                if is_valid_value(val):
                    initial_valuations[i] = val
        except Exception as e:
            pass
    
    remaining_contracts = [contracts[i] for i, v in enumerate(initial_valuations) if v is None]
    remaining_indices = [i for i, v in enumerate(initial_valuations) if v is None]
    
    if remaining_contracts and portfolio_recovery_file and os.path.exists(portfolio_recovery_file):
        try:
            recovery_data = read_columns_from_large_file(portfolio_recovery_file, 'A', ['B'], remaining_contracts)
            for idx, val in zip(remaining_indices, recovery_data['B']):
                if is_valid_value(val):
                    initial_valuations[idx] = val
        except Exception as e:
            pass
    
    remaining_contracts = [contracts[i] for i, v in enumerate(initial_valuations) if v is None]
    remaining_indices = [i for i, v in enumerate(initial_valuations) if v is None]
    
    if remaining_contracts:
        try:
            ia_data = read_multiple_columns_from_file(prod_file, 'IA Working', 'A', ['V'], remaining_contracts)
            for idx, val in zip(remaining_indices, ia_data['V']):
                if is_valid_value(val):
                    initial_valuations[idx] = val
        except Exception as e:
            pass
    
    missing_contracts = []
    for i, val in enumerate(initial_valuations):
        if val is None:
            missing_contracts.append({
                'Contract_No': contracts[i],
                'Issue': 'Initial Valuation Missing',
                'Description': 'Not found in Po Listing, Portfolio Recovery, or IA Working',
                'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            })
            initial_valuations[i] = ''
    
    if missing_contracts:
        for exception in missing_contracts:
            exception_tracker.add_exception('Bulk2_Initial_Valuation', exception)
    
    formatted_values = []
    for val in initial_valuations:
        if val and val != '':
            try:
                formatted_values.append(f"{float(val):,.0f}")
            except:
                formatted_values.append(str(val))
        else:
            formatted_values.append('')
    
    return formatted_values

def build_bulk1_dataframe(net_file, loan_base_file, loan_base_sheet, bulk1_contracts):
    """Build complete Bulk 1 DataFrame with all required data"""
    df_bulk1 = pd.DataFrame()
    df_bulk1['CONTRACT_NO'] = bulk1_contracts
    
    net_data = read_multiple_columns_from_file(net_file, 0, 'E', ['Z'], bulk1_contracts)
    df_bulk1['GROSS_PORTFOLIO'] = net_data['Z']
    
    df_bulk1['GROSS_PORTFOLIO'] = df_bulk1['GROSS_PORTFOLIO'].apply(
        lambda x: f"{float(x):,.0f}" if pd.notna(x) and x != '' and x is not None else ''
    )
    
    periods = calculate_contract_periods(loan_base_file, loan_base_sheet, bulk1_contracts)
    df_bulk1['CONTRACT_PERIOD'] = periods
    
    loan_base_data = read_multiple_columns_from_file(loan_base_file, loan_base_sheet, 'B', ['H', 'M', 'J'], bulk1_contracts)
    
    df_bulk1['CON_INTRATE'] = [
        f"{float(val):.2f}%" if pd.notna(val) and val != '' and val is not None else ''
        for val in loan_base_data['H']
    ]
    
    df_bulk1['CONTRACT_AMOUNT'] = [
        f"{float(val):,.0f}" if pd.notna(val) and val != '' and val is not None else ''
        for val in loan_base_data['M']
    ]
    
    df_bulk1['AF_VALUE'] = loan_base_data['J']
    
    return df_bulk1

def build_bulk2_dataframe(summary_file, net_file, po_listing_file, portfolio_recovery_file, 
                          prod_file, matching_contracts, exception_tracker):
    """Build complete Bulk 2 DataFrame with all required data including Initial Valuation"""
    df_bulk2 = read_net_portfolio_data(net_file, matching_contracts)
    
    summary_data = read_multiple_columns_from_file(summary_file, 'SUMMARY', 'A', ['AG', 'E', 'R'], df_bulk2['CONTRACT_NO'].tolist())
    
    df_bulk2['R_VALUE'] = summary_data['AG']
    df_bulk2['AB_VALUE'] = summary_data['E']
    df_bulk2['AD_VALUE'] = summary_data['R']
    
    initial_valuations = calculate_initial_valuation_bulk2(
        df_bulk2['CONTRACT_NO'].tolist(),
        po_listing_file,
        portfolio_recovery_file,
        prod_file,
        exception_tracker
    )
    
    df_bulk2['AF_VALUE'] = initial_valuations
    
    return df_bulk2

def build_bulk3_dataframe(net_file, unutilized_file, exception_tracker):
    """Build complete Bulk 3 DataFrame with Margin Trading data"""
    df_net = pd.read_excel(net_file, sheet_name=0, header=None)
    df_net.columns = ['Type'] + [f'Col_{i}' for i in range(1, len(df_net.columns))]
    
    margin_mask = df_net['Type'].astype(str).str.strip().str.upper() == 'MARGIN TRADING'
    
    if margin_mask.any():
        net_portfolio_value = df_net.loc[margin_mask, 'Col_25'].values[0]
        net_portfolio_value = float(net_portfolio_value) if pd.notna(net_portfolio_value) else 0
    else:
        net_portfolio_value = 0
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(unutilized_file)
            
            target_sheet = None
            for sheet in wb.sheets:
                if 'JULY' in sheet.name.upper() or 'JUL' in sheet.name.upper():
                    target_sheet = sheet
                    break
            
            if not target_sheet:
                target_sheet = wb.sheets[0]
            
            last_row = target_sheet.used_range.last_cell.row
            
            icam_codes = target_sheet.range(f'C1:C{last_row}').value
            closing_balances = target_sheet.range(f'G1:G{last_row}').value
            interest_rates = target_sheet.range(f'H1:H{last_row}').value
            
            wb.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise
    
    icam_codes = icam_codes if isinstance(icam_codes, list) else [icam_codes]
    closing_balances = closing_balances if isinstance(closing_balances, list) else [closing_balances]
    interest_rates = interest_rates if isinstance(interest_rates, list) else [interest_rates]
    
    df_unutilized = pd.DataFrame({
        'ICAM_CODE': icam_codes,
        'CLOSING_BALANCE': closing_balances,
        'INTEREST_RATE': interest_rates
    })
    
    df_unutilized['CLOSING_BALANCE_NUM'] = pd.to_numeric(df_unutilized['CLOSING_BALANCE'], errors='coerce')
    df_positive = df_unutilized[df_unutilized['CLOSING_BALANCE_NUM'] > 0].copy()
    
    df_positive = df_positive[~df_positive['ICAM_CODE'].astype(str).str.upper().str.contains('ICAM|CODE|TOTAL', na=False)]
    
    unutilized_sum = df_positive['CLOSING_BALANCE_NUM'].sum()
    
    difference = abs(net_portfolio_value - unutilized_sum)
    tolerance = 0.01
    values_tally = difference < tolerance
    
    if not values_tally:
        exception_tracker.add_exception('Bulk3_Margin_Trading', {
            'Check': 'Margin Trading Balance Comparison',
            'Net_Portfolio_Value': f"{net_portfolio_value:,.2f}",
            'Unutilized_Sum': f"{unutilized_sum:,.2f}",
            'Difference': f"{difference:,.2f}",
            'Status': 'MISMATCH',
            'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        })
    
    df_bulk3 = df_positive[['ICAM_CODE', 'CLOSING_BALANCE_NUM', 'INTEREST_RATE']].copy()
    df_bulk3.columns = ['CLIENT_CODE', 'GROSS_PORTFOLIO', 'CON_INTRATE']
    
    df_bulk3['GROSS_PORTFOLIO'] = df_bulk3['GROSS_PORTFOLIO'].apply(
        lambda x: f"{float(x):,.0f}" if pd.notna(x) else ''
    )
    
    df_bulk3['CON_INTRATE'] = df_bulk3['CON_INTRATE'].apply(
        lambda x: f"{float(x):.2f}%" if pd.notna(x) and x != '' else ''
    )
    
    return df_bulk3, values_tally

# ==================== SHEET OPERATIONS ====================

def process_moi_list(prod_file, reschedule_file):
    """Process MOI List sheet - add missing reschedule contracts"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            moi_sheet = wb.sheets['MOI List -31 Jan']
            
            last_row = moi_sheet.used_range.last_cell.row
            
            col_a = moi_sheet.range(f'A1:A{last_row}').value
            col_b = moi_sheet.range(f'B1:B{last_row}').value
            col_c = moi_sheet.range(f'C1:C{last_row}').value
            
            wb.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise
    
    col_a_list = col_a if isinstance(col_a, list) else [col_a]
    col_b_list = col_b if isinstance(col_b, list) else [col_b]
    col_c_list = col_c if isinstance(col_c, list) else [col_c]
    
    df_moi = pd.DataFrame({'A': col_a_list, 'B': col_b_list, 'C': col_c_list})
    
    reschedule_mask = df_moi['C'].astype(str).str.strip() == 'Reschedule Total List'
    reschedule_indices = df_moi[reschedule_mask].index.tolist()
    
    if not reschedule_indices:
        return
    
    moi_reschedule_contracts = set(
        str(val).strip().upper() for val in df_moi.loc[reschedule_mask, 'A']
        if val and str(val).strip() and is_valid_contract_number(str(val))
    )
    
    df_reschedule = pd.read_excel(reschedule_file, sheet_name=0)
    
    if 'CLO_NEWCONNO' in df_reschedule.columns:
        reschedule_contracts = set(
            str(val).strip().upper() for val in df_reschedule['CLO_NEWCONNO']
            if pd.notna(val) and is_valid_contract_number(str(val))
        )
    else:
        reschedule_contracts = set(
            str(val).strip().upper() for val in df_reschedule.iloc[:, 1]
            if pd.notna(val) and is_valid_contract_number(str(val))
        )
    
    missing_contracts = reschedule_contracts - moi_reschedule_contracts
    
    if not missing_contracts:
        return
    
    last_reschedule_idx = reschedule_indices[-1]
    template_b = df_moi.loc[last_reschedule_idx, 'B']
    template_c = df_moi.loc[last_reschedule_idx, 'C']
    
    insert_position = last_reschedule_idx + 1
    
    new_rows = [{'A': contract, 'B': template_b, 'C': template_c} 
                for contract in sorted(missing_contracts)]
    
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            moi_sheet = wb.sheets['MOI List -31 Jan']
            
            if new_rows:
                moi_sheet.range(f'{insert_position + 1}:{insert_position + len(new_rows)}').insert(shift='down')
                
                for i, row_data in enumerate(new_rows):
                    row_num = insert_position + i + 1
                    moi_sheet.range(f'A{row_num}').value = row_data['A']
                    moi_sheet.range(f'B{row_num}').value = row_data['B']
                    moi_sheet.range(f'C{row_num}').value = row_data['C']
            
            wb.save()
            wb.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def process_yard_and_property(prod_file, yard_stock_file, cbsl_provision_file):
    """Process YardandProperty List sheet"""
    df_yard = pd.read_excel(yard_stock_file, sheet_name=0)
    yard_all_values = df_yard.iloc[:, 0].tolist()
    
    yard_contracts = [
        str(val).strip() for val in yard_all_values 
        if is_valid_contract_number(str(val))
    ]
    
    df_property = pd.read_excel(cbsl_provision_file, sheet_name='PropertyMortgage')
    property_all_values = df_property.iloc[:, 0].tolist()
    
    property_contracts = [
        str(val).strip() for val in property_all_values 
        if is_valid_contract_number(str(val))
    ]
    
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            yard_sheet = wb.sheets['YardandProperty List']
            
            last_row = yard_sheet.used_range.last_cell.row
            if last_row > 1:
                yard_sheet.range(f'A2:A{last_row}').clear_contents()
                yard_sheet.range(f'D2:D{last_row}').clear_contents()
            
            current_header = yard_sheet.range('D1').value
            if current_header and 'June' in str(current_header):
                new_header = str(current_header).replace('June', 'July')
                yard_sheet.range('D1').value = new_header
            
            if yard_contracts:
                yard_array = [[c] for c in yard_contracts]
                yard_sheet.range(f'A2:A{1 + len(yard_contracts)}').value = yard_array
            
            if property_contracts:
                property_array = [[c] for c in property_contracts]
                yard_sheet.range(f'D2:D{1 + len(property_contracts)}').value = property_array
            
            wb.save()
            wb.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def process_staff_loan(prod_file, cadre_file):
    """Process Staff_Loan sheet"""
    max_retries = 3
    
    df_cadre = pd.read_excel(cadre_file, sheet_name=0, header=0)
    df_cadre.columns = df_cadre.columns.str.strip()
    
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb_prod = app.books.open(prod_file)
            
            staff_sheet = wb_prod.sheets['Staff_Loan']
            
            header_row = None
            for row_num in range(1, 11):
                row_values = []
                for col_num in range(1, 20):
                    col_letter = get_column_letter(col_num)
                    val = staff_sheet.range(f'{col_letter}{row_num}').value
                    if val:
                        row_values.append(val)
                
                if len(row_values) >= len(df_cadre.columns) * 0.5:
                    header_row = row_num
                    break
            
            if not header_row:
                header_row = 1
            
            last_col = staff_sheet.used_range.last_cell.column
            existing_headers = []
            for col_num in range(1, last_col + 1):
                col_letter = get_column_letter(col_num)
                header_val = staff_sheet.range(f'{col_letter}{header_row}').value
                if header_val:
                    existing_headers.append((col_num, col_letter, str(header_val).strip()))
            
            last_row = staff_sheet.used_range.last_cell.row
            data_start_row = header_row + 1
            if last_row >= data_start_row:
                staff_sheet.range(f'A{data_start_row}:{get_column_letter(last_col)}{last_row}').clear_contents()
            
            for col_num, col_letter, header in existing_headers:
                matching_cadre_col = None
                
                if header in df_cadre.columns:
                    matching_cadre_col = header
                else:
                    for cadre_col in df_cadre.columns:
                        if cadre_col.upper() == header.upper():
                            matching_cadre_col = cadre_col
                            break
                
                if matching_cadre_col:
                    data = df_cadre[matching_cadre_col].tolist()
                    data_array = [[d] for d in data]
                    staff_sheet.range(f'{col_letter}{data_start_row}:{col_letter}{data_start_row + len(data) - 1}').value = data_array
            
            wb_prod.save()
            wb_prod.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def process_mgt_acc(summary_file, prod_file):
    """Process MGT ACC values from Summary to FORM1 Working and NBD-MF-23-C1"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            
            # Read from Summary
            wb_summary = app.books.open(summary_file)
            summary_sheet = wb_summary.sheets['SUMMARY']
            last_row = summary_sheet.used_range.last_cell.row
            
            col_w = summary_sheet.range(f'W1:W{last_row}').value
            col_x = summary_sheet.range(f'X1:X{last_row}').value
            col_y = summary_sheet.range(f'Y1:Y{last_row}').value
            
            col_w_list = col_w if isinstance(col_w, list) else [col_w]
            col_x_list = col_x if isinstance(col_x, list) else [col_x]
            col_y_list = col_y if isinstance(col_y, list) else [col_y]
            
            net_portfolio_x = None
            net_portfolio_y = None
            total_adjustments_x = None
            
            for i, val in enumerate(col_w_list):
                if val and 'NET PORTFOLIO MGT ACC' in str(val).upper():
                    net_portfolio_x = col_x_list[i]
                    net_portfolio_y = col_y_list[i]
                elif val and 'TOTAL ADJUSTMENTS' in str(val).upper():
                    total_adjustments_x = col_x_list[i]
            
            wb_summary.close()
            
            # Write to FORM1 Working and NBD-MF-23-C1
            wb_prod = app.books.open(prod_file)
            
            # FORM1 Working sheet
            form1_sheet = wb_prod.sheets['FORM1 Working']
            form1_last_row = form1_sheet.used_range.last_cell.row
            
            col_f = form1_sheet.range(f'F1:F{form1_last_row}').value
            col_f_list = col_f if isinstance(col_f, list) else [col_f]
            
            as_per_ma_row = None
            acc_diff_row = None
            
            for i, val in enumerate(col_f_list, start=1):
                if val and 'AS PER MA' in str(val).upper():
                    as_per_ma_row = i
                elif val and 'ACC. DIFFERENCE' in str(val).upper():
                    acc_diff_row = i
            
            if as_per_ma_row and net_portfolio_x is not None:
                form1_sheet.range(f'G{as_per_ma_row}').value = net_portfolio_x
                if net_portfolio_y is not None:
                    form1_sheet.range(f'J{as_per_ma_row}').value = net_portfolio_y
            
            if acc_diff_row and total_adjustments_x is not None:
                form1_sheet.range(f'G{acc_diff_row}').value = total_adjustments_x
            
            # NBD-MF-23-C1 sheet
            c1_sheet = wb_prod.sheets['NBD-MF-23-C1']
            c1_last_row = c1_sheet.used_range.last_cell.row
            
            col_b = c1_sheet.range(f'B1:B{c1_last_row}').value
            col_b_list = col_b if isinstance(col_b, list) else [col_b]
            
            adjustments_row = None
            for i, val in enumerate(col_b_list, start=1):
                if val and 'ADJUSTMENTS' in str(val).upper():
                    adjustments_row = i
                    break
            
            if adjustments_row:
                c1_sheet.range(f'C{adjustments_row}').formula = "='FORM1 Working'!G11/1000"
            
            wb_prod.save()
            wb_prod.close()
            app.quit()
            break
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

# ==================== BULK OPERATIONS ====================

def find_bulk_rows_in_prod(prod_file, pattern):
    """Find bulk section rows by pattern"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']
            
            last_row = sheet.used_range.last_cell.row
            compiled_pattern = re.compile(pattern, re.IGNORECASE)
            bulk_rows = []
            
            chunk_size = 10000
            for start in range(1, last_row + 1, chunk_size):
                end = min(start + chunk_size - 1, last_row)
                data = sheet.range(f'A{start}:A{end}').value
                data_list = data if isinstance(data, list) else [data]
                
                for i, cell in enumerate(data_list):
                    if cell and compiled_pattern.match(str(cell).strip()):
                        bulk_rows.append(start + i)
            
            wb.close()
            app.quit()
            return bulk_rows
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def adjust_rows_in_prod(sheet, bulk_rows, required_rows, data_columns):
    """Adjust rows and copy formulas/formats from row above"""
    if not bulk_rows:
        return []
    
    current_rows = len(bulk_rows)
    
    if current_rows == required_rows:
        return bulk_rows
    
    start_row, end_row = bulk_rows[0], bulk_rows[-1]
    
    if required_rows > current_rows:
        rows_to_add = required_rows - current_rows
        
        sheet.range(f'{end_row + 1}:{end_row + rows_to_add}').insert(shift='down')
        
        template_row = sheet.range(f'{end_row}:{end_row}')
        template_row.copy()
        
        new_rows_range = sheet.range(f'{end_row + 1}:{end_row + rows_to_add}')
        new_rows_range.paste()
        
        try:
            sheet.api.Application.CutCopyMode = False
        except:
            pass
        
        for col in data_columns:
            sheet.range(f'{col}{end_row + 1}:{col}{end_row + rows_to_add}').clear_contents()
        
        return list(range(start_row, end_row + rows_to_add + 1))
    
    else:
        rows_to_remove = current_rows - required_rows
        delete_start = end_row - rows_to_remove + 1
        sheet.range(f'{delete_start}:{end_row}').delete(shift='up')
        
        return list(range(start_row, delete_start))

def paste_bulk1_dataframe_to_prod(prod_file, bulk1_rows, df_bulk1):
    """Paste Bulk 1 DataFrame to prod"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            app.calculation = 'manual'
            
            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']
            
            data_columns = ['A', 'C', 'K', 'T', 'V', 'W', 'AF']
            
            required_rows = len(df_bulk1)
            adjusted_rows = adjust_rows_in_prod(sheet, bulk1_rows, required_rows, data_columns)
            
            if not adjusted_rows:
                wb.close()
                app.quit()
                return adjusted_rows
            
            start_row = adjusted_rows[0]
            end_row = start_row + required_rows - 1
            
            data_mapping = {
                'A': df_bulk1['CONTRACT_NO'].tolist(),
                'C': df_bulk1['CONTRACT_NO'].tolist(),
                'K': df_bulk1['GROSS_PORTFOLIO'].tolist(),
                'T': df_bulk1['CONTRACT_PERIOD'].tolist(),
                'V': df_bulk1['CON_INTRATE'].tolist(),
                'W': df_bulk1['CONTRACT_AMOUNT'].tolist(),
                'AF': df_bulk1['AF_VALUE'].tolist()
            }
            
            for col_letter, values in data_mapping.items():
                values_array = [[v] for v in values]
                sheet.range(f'{col_letter}{start_row}:{col_letter}{end_row}').value = values_array
            
            app.calculation = 'automatic'
            wb.save()
            wb.close()
            app.quit()
            
            return adjusted_rows
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def paste_bulk2_dataframe_to_prod(prod_file, bulk2_rows, df_bulk2):
    """Paste Bulk 2 DataFrame to prod"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            app.calculation = 'manual'
            
            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']
            
            data_columns = ['A', 'C', 'D', 'E', 'K', 'L', 'M', 'S', 'T', 'V', 'W', 'R', 'AB', 'AD', 'AF']
            
            required_rows = len(df_bulk2)
            adjusted_rows = adjust_rows_in_prod(sheet, bulk2_rows, required_rows, data_columns)
            
            if not adjusted_rows:
                wb.close()
                app.quit()
                return adjusted_rows
            
            start_row = adjusted_rows[0]
            end_row = start_row + required_rows - 1
            
            data_mapping = {
                'A': df_bulk2['CONTRACT_NO'].tolist(),
                'C': df_bulk2['CLIENT_CODE'].tolist(),
                'D': df_bulk2['EQT_DESC'].tolist(),
                'E': df_bulk2['PURPOSE'].tolist(),
                'K': df_bulk2['GROSS_PORTFOLIO'].tolist(),
                'L': df_bulk2['PROVISION'].tolist(),
                'M': df_bulk2['DP_PROVISION'].tolist(),
                'S': df_bulk2['CON_RNTFREQ'].tolist(),
                'T': df_bulk2['CONTRACT_PERIOD'].tolist(),
                'V': df_bulk2['CON_INTRATE'].tolist(),
                'W': df_bulk2['CONTRACT_AMOUNT'].tolist(),
                'R': df_bulk2['R_VALUE'].tolist(),
                'AB': df_bulk2['AB_VALUE'].tolist(),
                'AD': df_bulk2['AD_VALUE'].tolist(),
                'AF': df_bulk2['AF_VALUE'].tolist()
            }
            
            for col_letter, values in data_mapping.items():
                values_array = [[v] for v in values]
                sheet.range(f'{col_letter}{start_row}:{col_letter}{end_row}').value = values_array
            
            app.calculation = 'automatic'
            wb.save()
            wb.close()
            app.quit()
            
            return adjusted_rows
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

def paste_bulk3_dataframe_to_prod(prod_file, bulk3_rows, df_bulk3):
    """Paste Bulk 3 DataFrame to prod"""
    max_retries = 3
    for attempt in range(max_retries):
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            app.screen_updating = False
            app.calculation = 'manual'
            
            wb = app.books.open(prod_file)
            sheet = wb.sheets['C1 & C2 Working']
            
            data_columns = ['C', 'K', 'V']
            
            required_rows = len(df_bulk3)
            adjusted_rows = adjust_rows_in_prod(sheet, bulk3_rows, required_rows, data_columns)
            
            if not adjusted_rows:
                wb.close()
                app.quit()
                return adjusted_rows
            
            start_row = adjusted_rows[0]
            end_row = start_row + required_rows - 1
            
            data_mapping = {
                'C': df_bulk3['CLIENT_CODE'].tolist(),
                'K': df_bulk3['GROSS_PORTFOLIO'].tolist(),
                'V': df_bulk3['CON_INTRATE'].tolist()
            }
            
            for col_letter, values in data_mapping.items():
                values_array = [[v] for v in values]
                sheet.range(f'{col_letter}{start_row}:{col_letter}{end_row}').value = values_array
            
            app.calculation = 'automatic'
            wb.save()
            wb.close()
            app.quit()
            
            return adjusted_rows
            
        except Exception as e:
            if attempt < max_retries - 1:
                kill_excel_processes()
                time.sleep(3)
            else:
                raise

# ==================== MAIN EXECUTION ====================

def main():
    script_start = time.time()
    base_path = r"C:\CBSL\Script"
    
    print("Checking for existing Excel processes...")
    kill_excel_processes()
    
    print("="*60)
    print("EXCEL AUTOMATION - COMPLETE")
    print(f"Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)
    
    folder = find_dynamic_folder(base_path)
    if not folder:
        return
    
    exception_tracker = ExceptionTracker(folder)
    
    (summary_file, net_file, prod_file, loan_base_file, reschedule_file, 
     yard_stock_file, cbsl_provision_file, cadre_file, unutilized_file,
     po_listing_file, portfolio_recovery_file) = find_files(folder)
    
    missing = []
    if not summary_file: missing.append("Summary")
    if not net_file: missing.append("Net Portfolio")
    if not prod_file: missing.append("Prod. wise Class. of Loans")
    if not loan_base_file: missing.append("Loan Base")
    if not reschedule_file: missing.append("Reschedule Contract Details")
    if not yard_stock_file: missing.append("YARD STOCK")
    if not cbsl_provision_file: missing.append("CBSL Provision Comparison")
    if not cadre_file: missing.append("Cadre")
    if not unutilized_file: missing.append("Unutilized Amount")
    if not po_listing_file: missing.append("Po Listing - Internal")
    if not portfolio_recovery_file: missing.append("Portfolio Report Recovery - Internal")
    
    if missing:
        print(f"\nERROR: Missing files: {', '.join(missing)}")
        return
    
    print(f"\nAll files found")
    
    # Step 1: MOI List
    print("\n" + "="*60)
    print("STEP 1: MOI LIST")
    print("="*60)
    step = time.time()
    process_moi_list(prod_file, reschedule_file)
    log_step("Step 1", step)
    
    # Step 2: Yard & Property
    print("\n" + "="*60)
    print("STEP 2: YARD AND PROPERTY")
    print("="*60)
    step = time.time()
    process_yard_and_property(prod_file, yard_stock_file, cbsl_provision_file)
    log_step("Step 2", step)
    
    # Step 3: Staff Loan
    print("\n" + "="*60)
    print("STEP 3: STAFF LOAN")
    print("="*60)
    step = time.time()
    process_staff_loan(prod_file, cadre_file)
    log_step("Step 3", step)
    
    # Step 4-7: Bulk 1
    print("\n" + "="*60)
    print("STEP 4-7: BULK 1 (LR Pattern)")
    print("="*60)
    step = time.time()
    net_contracts = read_contracts_fast(net_file, 0, 'E')
    net_bulk1 = classify_contracts(net_contracts, r'^LR\d{8}$')
    print(f"  Bulk 1: {len(net_bulk1)} contracts")
    log_step("Step 4-5", step)
    
    step = time.time()
    loan_base_sheet = get_first_sheet_name(loan_base_file)
    df_bulk1 = build_bulk1_dataframe(net_file, loan_base_file, loan_base_sheet, net_bulk1)
    bulk1_rows = find_bulk_rows_in_prod(prod_file, r'^LR\d{8}$')
    adjusted_bulk1_rows = paste_bulk1_dataframe_to_prod(prod_file, bulk1_rows, df_bulk1)
    log_step("Step 6-7", step)
    
    # Step 8-12: Bulk 2
    print("\n" + "="*60)
    print("STEP 8-12: BULK 2 (4-Letter + 9-Digit Pattern)")
    print("="*60)
    step = time.time()
    summary_contracts = read_contracts_fast(summary_file, 'SUMMARY', 'A')
    summary_bulk2 = classify_contracts(summary_contracts, r'^[A-Z]{4}\d{9}$')
    net_bulk2 = classify_contracts(net_contracts, r'^[A-Z]{4}\d{9}$')
    matches = set(summary_bulk2) & set(net_bulk2)
    print(f"  Bulk 2: {len(matches)} contracts")
    log_step("Step 8-9", step)
    
    step = time.time()
    df_bulk2 = build_bulk2_dataframe(summary_file, net_file, po_listing_file, 
                                      portfolio_recovery_file, prod_file, 
                                      list(matches), exception_tracker)
    bulk2_rows = find_bulk_rows_in_prod(prod_file, r'^[A-Z]{4}\d{9}$')
    adjusted_bulk2_rows = paste_bulk2_dataframe_to_prod(prod_file, bulk2_rows, df_bulk2)
    log_step("Step 10-12", step)
    
    # Step 13-15: Bulk 3
    print("\n" + "="*60)
    print("STEP 13-15: BULK 3 (MARGIN TRADING)")
    print("="*60)
    step = time.time()
    df_bulk3, values_tally = build_bulk3_dataframe(net_file, unutilized_file, exception_tracker)
    print(f"  Bulk 3: {len(df_bulk3)} entries")
    bulk3_rows = find_bulk_rows_in_prod(prod_file, r'^MARGIN\s+TRADING$')
    adjusted_bulk3_rows = paste_bulk3_dataframe_to_prod(prod_file, bulk3_rows, df_bulk3)
    log_step("Step 13-15", step)
    
    # Step 16: MGT ACC
    print("\n" + "="*60)
    print("STEP 16: MGT ACC PROCESSING")
    print("="*60)
    step = time.time()
    process_mgt_acc(summary_file, prod_file)
    log_step("Step 16", step)
    
    # Write exceptions
    print("\n" + "="*60)
    print("WRITING EXCEPTIONS")
    print("="*60)
    exception_tracker.write_exceptions()
    
    # Summary
    total_time = time.time() - script_start
    print("\n" + "="*60)
    print("COMPLETED SUCCESSFULLY")
    print("="*60)
    print(f"\nBulk 1: {len(net_bulk1)} | Bulk 2: {len(matches)} | Bulk 3: {len(df_bulk3)}")
    print(f"Total time: {total_time:.2f}s ({total_time/60:.2f} min)")
    print(f"Finished: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("="*60)

if __name__ == "__main__":
    main()