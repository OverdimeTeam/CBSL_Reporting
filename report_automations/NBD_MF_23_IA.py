import os
import sys
import json
import logging
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import requests
from contextlib import contextmanager
import win32com.client
import time

# Add bots directory to path for imports
bots_dir = Path(__file__).parent.parent / "bots"
sys.path.append(str(bots_dir))

# Configure logging first
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Import bot modules
try:
    import na_contract_numbers_search_bot_api as na_contract_bot
    logger.info("Successfully imported na_contract_numbers_search_bot_api module")
except ImportError as e:
    logger.warning(f"Could not import na_contract_numbers_search_bot_api: {e}")
    na_contract_bot = None

try:
    import IA_Working_Initial_valuation_bot as valuation_bot
    logger.info("Successfully imported IA_Working_Initial_valuation_bot module")
except ImportError as e:
    logger.warning(f"Could not import IA_Working_Initial_valuation_bot: {e}")
    valuation_bot = None

# Scienter bot not implemented yet - will be added later
logger.info("Scienter bot not yet implemented - will be added later")

# Context manager for error handling
@contextmanager
def error_handler(automation_instance, step_name, contract_number=None):
    """Context manager for consistent error handling across all steps"""
    try:
        yield
    except Exception as e:
        automation_instance.log_exception(step_name, str(e), contract_number)
        raise

# Function to run the contract search bot
def run_contract_search_bot(contract_numbers):
    """Run the contract search bot with the given contract numbers"""
    if na_contract_bot:
        try:
            logger.info("Running na_contract_numbers_search_bot_api...")
            # Call the main function of the bot
            na_contract_bot.main()
            logger.info("Contract search bot completed successfully")
            # Note: The bot saves results to files, so we need to read them back
            return _read_contract_search_results(contract_numbers)
        except Exception as e:
            logger.error(f"Contract search bot failed: {e}")
            return _generate_mock_data(contract_numbers)
    else:
        logger.warning("Contract search bot not available - using mock data")
        return _generate_mock_data(contract_numbers)

# Function to run the valuation bot
def run_valuation_bot_wrapper(contract_numbers):
    """Run the valuation bot with the given contract numbers"""
    if valuation_bot:
        try:
            logger.info("Running IA_Working_Initial_valuation_bot...")
            # Call the run_valuation_bot function of the bot
            results = valuation_bot.run_valuation_bot(contract_numbers)
            logger.info("Valuation bot completed successfully")
            return results
        except Exception as e:
            logger.error(f"Valuation bot failed: {e}")
            return {}
    else:
        logger.warning("Valuation bot not available")
        return {}

def _read_contract_search_results(contract_numbers):
    """Read the results from the contract search bot output files"""
    try:
        # This would read the actual output files from the bot
        # For now, return mock data until we implement file reading
        logger.warning("Reading bot results not yet implemented - using mock data")
        return _generate_mock_data(contract_numbers)
    except Exception as e:
        logger.error(f"Failed to read bot results: {e}")
        return _generate_mock_data(contract_numbers)

def _generate_mock_data(contract_numbers):
    """Generate mock data for contracts"""
    mock_data = {}
    for contract in contract_numbers:
        mock_data[contract] = {
            'client_code': f'MOCK_{contract[-4:]}',
            'equipment': 'Mock Equipment',
            'contract_period': 12,
            'frequency': 'M',
            'interest_rate': 15.0,
            'contract_amount': 100000,
            'AT_limit': 100000
        }
    return mock_data

class NBDMF23IAAutomation:
    def __init__(self, working_dir: str, month: str, year: str = "2025"):
        self.working_dir = Path(working_dir)
        self.month = month
        self.year = year
        self.month_year = f"{month}-{year}"
        self.exceptions = []
        
        # File paths
        self.main_file = self._find_file_by_prefix("Prod. wise Class. of Loans")
        self.disbursement_file = self._find_file_by_prefix("Disbursement with Budget")
        self.net_portfolio_file = self._find_file_by_prefix("Net Portfolio")
        self.po_listing_file = self._find_file_by_prefix("Po Listing - Internal")
        self.info_request_file = self._find_file_by_prefix("Copy of Information Request from Credit")
        self.portfolio_recovery_file = self._find_file_by_prefix("Portfolio Report Recovery", required=False)
        
        # Excel COM objects
        self.excel_app = None
        self.workbook = None
        self.workbook_path = None
        
        # Bot directories
        self.bots_dir = Path(__file__).parent.parent / "bots"
        
    def initialize_excel(self):
        """Initialize Excel COM application and open the workbook"""
        try:
            logger.info("Initializing Excel COM application...")
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = False
            self.excel_app.DisplayAlerts = False
            
            # Open the workbook
            self.workbook_path = str(self.main_file.absolute())
            self.workbook = self.excel_app.Workbooks.Open(self.workbook_path)
            logger.info(f"Successfully opened workbook: {self.main_file.name}")
            
        except Exception as e:
            logger.error(f"Failed to initialize Excel: {e}")
            raise
    
    def close_excel(self):
        """Close Excel application and clean up"""
        try:
            if self.workbook:
                self.workbook.Save()
                self.workbook.Close()
                logger.info("Workbook saved and closed")
            
            if self.excel_app:
                self.excel_app.Quit()
                logger.info("Excel application closed")
                
        except Exception as e:
            logger.error(f"Error closing Excel: {e}")
        finally:
            self.excel_app = None
            self.workbook = None
    
    def get_worksheet(self, sheet_name: str):
        """Get a worksheet by name"""
        try:
            worksheet = self.workbook.Worksheets(sheet_name)
            return worksheet
        except Exception as e:
            logger.error(f"Failed to get worksheet '{sheet_name}': {e}")
            raise
    
    def write_cell_value(self, sheet_name: str, row: int, col: int, value):
        """Write a value to a specific cell"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            worksheet.Cells(row, col).Value = value
        except Exception as e:
            logger.error(f"Failed to write to cell {chr(64 + col)}{row}: {e}")
            raise
    
    def write_cell_formula(self, sheet_name: str, row: int, col: int, formula: str):
        """Write a formula to a specific cell"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            worksheet.Cells(row, col).Formula = formula
        except Exception as e:
            logger.error(f"Failed to write formula to cell {chr(64 + col)}{row}: {e}")
            raise
    
    def fill_down_formula(self, sheet_name: str, start_row: int, end_row: int, col: int, formula: str):
        """Fill down a formula from start_row to end_row in the specified column"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            
            # Write the formula to the first row
            worksheet.Cells(start_row, col).Formula = formula
            
            # Select the range and fill down (Ctrl+D equivalent)
            range_obj = worksheet.Range(f"{chr(64 + col)}{start_row}:{chr(64 + col)}{end_row}")
            range_obj.FillDown()
            
            logger.info(f"Filled down formula from row {start_row} to {end_row} in column {chr(64 + col)}")
            
        except Exception as e:
            logger.error(f"Failed to fill down formula: {e}")
            raise
    
    def copy_range_values(self, sheet_name: str, source_range: str, target_range: str):
        """Copy values from source range to target range"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            source = worksheet.Range(source_range)
            target = worksheet.Range(target_range)
            
            source.Copy(target)
            logger.info(f"Copied values from {source_range} to {target_range}")
            
        except Exception as e:
            logger.error(f"Failed to copy range: {e}")
            raise
    
    def find_last_row(self, sheet_name: str, col: int = 1):
        """Find the last row with data in the specified column"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            last_row = worksheet.Cells(worksheet.Rows.Count, col).End(-4162).Row  # xlUp = -4162
            return last_row
        except Exception as e:
            logger.error(f"Failed to find last row: {e}")
            raise
    
    def read_cell_value(self, sheet_name: str, row: int, col: int):
        """Read a value from a specific cell"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            return worksheet.Cells(row, col).Value
        except Exception as e:
            logger.error(f"Failed to read cell {chr(64 + col)}{row}: {e}")
            raise
    
    def read_range_values(self, sheet_name: str, start_row: int, end_row: int, start_col: int, end_col: int):
        """Read values from a range of cells"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            range_obj = worksheet.Range(f"{chr(64 + start_col)}{start_row}:{chr(64 + end_col)}{end_row}")
            return range_obj.Value
        except Exception as e:
            logger.error(f"Failed to read range: {e}")
            raise
    
    def clear_range(self, sheet_name: str, start_row: int, end_row: int, start_col: int, end_col: int):
        """Clear values from a range of cells"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            range_obj = worksheet.Range(f"{chr(64 + start_col)}{start_row}:{chr(64 + end_col)}{end_row}")
            range_obj.ClearContents()
            logger.info(f"Cleared range {chr(64 + start_col)}{start_row}:{chr(64 + end_col)}{end_row}")
        except Exception as e:
            logger.error(f"Failed to clear range: {e}")
            raise
    
    def apply_vlookup_formula(self, sheet_name: str, target_col: int, lookup_value_col: int, 
                             table_array: str, col_index: int, start_row: int, end_row: int):
        """Apply VLOOKUP formula to a range of cells with Fill Down (Ctrl+D)"""
        try:
            # Create VLOOKUP formula
            formula = f'=VLOOKUP({chr(64 + lookup_value_col)}{start_row},{table_array},{col_index},FALSE)'
            
            # Apply formula to first row
            self.write_cell_formula(sheet_name, start_row, target_col, formula)
            
            # Fill down the formula to all rows (Ctrl+D equivalent)
            self.fill_down_formula(sheet_name, start_row, end_row, target_col, formula)
            
            logger.info(f"Applied VLOOKUP formula to column {chr(64 + target_col)} from row {start_row} to {end_row}")
            
        except Exception as e:
            logger.error(f"Failed to apply VLOOKUP formula: {e}")
            raise
    
    def apply_calculation_formula(self, sheet_name: str, target_col: int, formula: str, 
                                start_row: int, end_row: int):
        """Apply a calculation formula to a range of cells with Fill Down (Ctrl+D)"""
        try:
            # Apply formula to first row
            self.write_cell_formula(sheet_name, start_row, target_col, formula)
            
            # Fill down the formula to all rows (Ctrl+D equivalent)
            self.fill_down_formula(sheet_name, start_row, end_row, target_col, formula)
            
            logger.info(f"Applied calculation formula to column {chr(64 + target_col)} from row {start_row} to {end_row}")
            
        except Exception as e:
            logger.error(f"Failed to apply calculation formula: {e}")
            raise
    
    def demonstrate_pywin32_features(self):
        """Demonstrate pywin32 features including formulas and Fill Down"""
        try:
            logger.info("Demonstrating pywin32 features...")
            
            # Example 1: Apply VLOOKUP formula with Fill Down
            # This would look up values from another sheet/range
            # self.apply_vlookup_formula("IA Working", 3, 1, "NetPortfolio!A:C", 2, 3, 100)
            
            # Example 2: Apply calculation formula with Fill Down
            # This would calculate a value based on other cells
            # self.apply_calculation_formula("IA Working", 22, "=P3/V3", 3, 100)
            
            # Example 3: Write individual cell values
            self.write_cell_value("IA Working", 1, 1, "Contract Number")
            self.write_cell_value("IA Working", 1, 2, "Status")
            
            # Example 4: Clear a range of cells
            # self.clear_range("IA Working", 3, 10, 5, 8)
            
            logger.info("pywin32 features demonstration completed")
            
        except Exception as e:
            logger.error(f"Failed to demonstrate pywin32 features: {e}")
            raise
    
    def _find_file_by_prefix(self, prefix: str, required: bool = True) -> Path:
        """Find a file that starts with the given prefix in the working directory"""
        for file_path in self.working_dir.iterdir():
            if file_path.is_file() and file_path.name.startswith(prefix):
                logger.info(f"Found file: {file_path.name}")
                return file_path
        
        # If no file found and it's required, raise an error
        if required:
            available_files = [f.name for f in self.working_dir.iterdir() if f.is_file()]
            raise FileNotFoundError(f"No file found starting with '{prefix}' in {self.working_dir}. Available files: {available_files}")
        else:
            # Return None for optional files
            logger.info(f"Optional file '{prefix}' not found - continuing without it")
            return None
    
    def toggle_module_requirement(self, module_name: str, required: bool = True):
        """Easily toggle whether a module is required or optional"""
        logger.info(f"Setting {module_name} requirement to: {required}")
        # This method can be extended to dynamically change module requirements
        # For now, it just logs the change for debugging purposes
    
    def check_module_status(self):
        """Check the status of all required modules and bots"""
        status = {
            'na_contract_bot': 'Available' if na_contract_bot else 'Not Available',
            'valuation_bot': 'Available' if valuation_bot else 'Not Available',
            'scienter_bot': 'Not Implemented',
            'workhub24_api': 'Not Implemented'
        }
        
        logger.info("Module Status Check:")
        for module, status_text in status.items():
            logger.info(f"  {module}: {status_text}")
        
        return status
    
    def log_exception(self, step: str, message: str, contract_number: str = None):
        """Log exceptions for reporting"""
        exception_info = {
            'step': step,
            'message': message,
            'contract_number': contract_number,
            'timestamp': datetime.now().isoformat()
        }
        self.exceptions.append(exception_info)
        logger.error(f"Step {step}: {message} - Contract: {contract_number}")
    
    def get_exception_summary(self):
        """Get a summary of all exceptions by step"""
        if not self.exceptions:
            return "No exceptions recorded"
        
        step_counts = {}
        for exception in self.exceptions:
            if isinstance(exception, dict) and 'step' in exception:
                step = exception['step']
                step_counts[step] = step_counts.get(step, 0) + 1
        
        summary = "Exception Summary:\n"
        for step, count in sorted(step_counts.items()):
            summary += f"  Step {step}: {count} exceptions\n"
        
        return summary
    
    def step_1_copy_disbursement_data(self):
        """Step 1: Copy data from Disbursement workbook to IA Working sheet"""
        try:
            logger.info("Step 1: Copying disbursement data...")
            
            # Initialize Excel if not already done
            if not self.workbook:
                self.initialize_excel()
            
            # Read disbursement data using bulk method for better performance
            disbursement_data = self.read_bulk_data_from_excel(self.disbursement_file, sheet_name="month")
            
            # Extract required columns (A, H, AC) - Contract No, Net Amount, Base Rate
            contract_numbers = [row[0] if len(row) > 0 else None for row in disbursement_data]  # Column A
            net_amounts = [row[7] if len(row) > 7 else None for row in disbursement_data]      # Column H
            base_rates = [row[28] if len(row) > 28 else None for row in disbursement_data]     # Column AC
            
            # Prepare data for bulk writing - organize by columns
            column_a_data = [[contract] for contract in contract_numbers if contract is not None]
            column_o_data = [[amount] for amount in net_amounts if amount is not None]
            column_n_data = [[rate] for rate in base_rates if rate is not None]
            
            # Write data in bulk starting from row 3
            if column_a_data:
                self.write_bulk_data("IA Working", 3, 1, column_a_data)      # Column A
            if column_o_data:
                self.write_bulk_data("IA Working", 3, 15, column_o_data)     # Column O
            if column_n_data:
                self.write_bulk_data("IA Working", 3, 14, column_n_data)     # Column N
            
            logger.info(f"Step 1 completed: Bulk copied {len(contract_numbers)} records to IA Working sheet")
            
        except Exception as e:
            self.log_exception("1", f"Failed to copy disbursement data: {str(e)}")
            raise
    
    def step_2_vlookup_net_portfolio(self):
        """Step 2: VLOOKUP data from Net Portfolio file to IA Working Sheet"""
        try:
            logger.info("Step 2: Performing VLOOKUP from Net Portfolio file...")
            
            # Find the Net Portfolio file dynamically
            net_portfolio_file = self._find_file_by_prefix("Net Portfolio")
            if not net_portfolio_file:
                raise FileNotFoundError("Net Portfolio file not found")
            
            logger.info(f"Using Net Portfolio file: {net_portfolio_file}")
            
            # Read net portfolio data using bulk method
            net_portfolio_data = self.read_bulk_data_from_excel(net_portfolio_file)
            
            if not net_portfolio_data:
                logger.warning("No data found in Net Portfolio file")
                return
            
            # Get contract numbers from IA Working sheet (Column A, starting from row 3)
            contracts = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                contracts.append((row, contract))
                row += 1
            
            if not contracts:
                logger.warning("No contracts found in IA Working sheet")
                return
            
            logger.info(f"Found {len(contracts)} contracts in IA Working sheet")
            
            # Create lookup dictionary for faster matching
            # Map: Contract Number (A) -> {column mappings}
            lookup_dict = {}
            for row_data in net_portfolio_data:
                if len(row_data) > 38:  # Ensure we have enough columns for AM (index 38)
                    contract_key = row_data[0]  # Column A (Contract Number)
                    if contract_key:  # Only add if contract key exists
                        lookup_dict[contract_key] = {
                            'client_code': row_data[2] if len(row_data) > 2 else None,      # Column C (CLIENT_CODE)
                            'equipment': row_data[25] if len(row_data) > 25 else None,      # Column AB (EQT_DESC)
                            'purpose': row_data[38] if len(row_data) > 38 else None,        # Column AM (PURPOSE)
                            'frequency': row_data[33] if len(row_data) > 33 else None,      # Column AH (CON_RNTFREQ)
                            'contract_period': row_data[5] if len(row_data) > 5 else None,  # Column F (CONTRACT_PERIOD)
                            'interest_rate': row_data[34] if len(row_data) > 34 else None, # Column AI (CON_INTRATE)
                            'contract_amount': row_data[7] if len(row_data) > 7 else None   # Column H (CONTRACT_AMOUNT)
                        }
            
            logger.info(f"Created lookup dictionary with {len(lookup_dict)} contracts from Net Portfolio")
            
            # Process each contract and write data to the correct columns
            processed_count = 0
            for row, contract in contracts:
                if contract in lookup_dict:
                    data = lookup_dict[contract]
                    
                    # Write data to the specific columns (all starting from 3rd row)
                    # Column C (Client Code) - from CLIENT_CODE (C column of Net Portfolio)
                    if data['client_code'] is not None:
                        self.write_cell_value("IA Working", row, 3, data['client_code'])
                    
                    # Column D (Equipment) - from EQT_DESC (AB column of Net Portfolio)
                    if data['equipment'] is not None:
                        self.write_cell_value("IA Working", row, 4, data['equipment'])
                    
                    # Column E (Purpose) - from PURPOSE (AM column of Net Portfolio)
                    if data['purpose'] is not None:
                        self.write_cell_value("IA Working", row, 5, data['purpose'])
                    
                    # Column I (Frequency) - from CON_RNTFREQ (AH column of Net Portfolio)
                    if data['frequency'] is not None:
                        self.write_cell_value("IA Working", row, 9, data['frequency'])
                    
                    # Column J (Contract Period) - from CONTRACT_PERIOD (F column of Net Portfolio)
                    if data['contract_period'] is not None:
                        self.write_cell_value("IA Working", row, 10, data['contract_period'])
                    
                    # Column M (Contractual Interest Rate) - from CON_INTRATE (AI column of Net Portfolio)
                    if data['interest_rate'] is not None:
                        self.write_cell_value("IA Working", row, 13, data['interest_rate'])
                    
                    # Column P (Contract Amount) - from CONTRACT_AMOUNT (H column of Net Portfolio)
                    if data['contract_amount'] is not None:
                        self.write_cell_value("IA Working", row, 16, data['contract_amount'])
                    
                    processed_count += 1
                    
                else:
                    # Contract not found in Net Portfolio - log for debugging
                    logger.debug(f"Contract {contract} not found in Net Portfolio - skipping VLOOKUP")
            
            logger.info(f"Step 2 completed: VLOOKUP completed for {processed_count} out of {len(contracts)} contracts")
            
        except Exception as e:
            self.log_exception("2", f"Failed to perform Net Portfolio VLOOKUP: {str(e)}")
            raise
    
    def step_3_po_listing_vlookup(self):
        """Step 3: VLOOKUP from Po Listing for Vehicles and Machinery"""
        try:
            logger.info("Step 3: Bulk VLOOKUP from Po Listing for Vehicles and Machinery...")
            
            # Read Po Listing data using bulk method
            po_listing_data = self.read_bulk_data_from_excel(self.po_listing_file)
            
            # Get all rows from IA Working sheet where Column U = "Vehicles and Machinery"
            vehicles_machinery_rows = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                u_value = self.read_cell_value("IA Working", row, 21)  # Column U
                if u_value == "Vehicles and Machinery":
                    vehicles_machinery_rows.append((row, contract))
                row += 1
            
            if not vehicles_machinery_rows:
                logger.info("No rows found with 'Vehicles and Machinery' in Column U")
                return
            
            # Create lookup dictionary for Po Listing data
            po_lookup_dict = {}
            for row_data in po_listing_data:
                if len(row_data) > 33:  # Ensure we have enough columns
                    contract_key = row_data[7] if len(row_data) > 7 else None  # Column H
                    sell_price = row_data[33] if len(row_data) > 33 else None  # Column AH
                    if contract_key:
                        po_lookup_dict[contract_key] = sell_price
            
            # Prepare bulk data for Column V
            sell_price_data = []
            for row, contract in vehicles_machinery_rows:
                sell_price = po_lookup_dict.get(contract, None)
                sell_price_data.append([sell_price])
            
            # Write sell price data in bulk to Column V
            if sell_price_data:
                self.write_bulk_data("IA Working", 3, 22, sell_price_data)  # Column V (index 21)
            
            logger.info(f"Step 3 completed: Bulk Po Listing data imported for {len(vehicles_machinery_rows)} Vehicles and Machinery rows")
            
        except Exception as e:
            self.log_exception("3", f"Failed to perform Po Listing VLOOKUP: {str(e)}")
            raise
    
    def step_4_info_request_vlookup(self):
        """Step 4: VLOOKUP from Information Request file"""
        try:
            logger.info("Step 4: Bulk VLOOKUP from Information Request file...")
            
            # Read Information Request data using bulk method
            info_request_data = self.read_bulk_data_from_excel(self.info_request_file, sheet_name="Disbursements")
            
            # Get all contract numbers from IA Working sheet
            contracts = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                contracts.append((row, contract))
                row += 1
            
            if not contracts:
                logger.warning("No contracts found in IA Working sheet")
                return
            
            # Create lookup dictionary for Information Request data
            info_lookup_dict = {}
            for row_data in info_request_data:
                if len(row_data) > 6:  # Ensure we have enough columns
                    contract_key = row_data[2] if len(row_data) > 2 else None  # Column C
                    enterprise_type = row_data[6] if len(row_data) > 6 else None  # Column G
                    if contract_key:
                        # Standardize enterprise type values
                        if enterprise_type:
                            enterprise_type = str(enterprise_type).strip().upper()
                            if enterprise_type in ["SMALL", "MICRO", "MEDIUM", "COOPERATE", "OTHER", "ENTERPRISES"]:
                                if enterprise_type == "COOPERATE" or enterprise_type == "OTHER":
                                    enterprise_type = "Cooperate/Other Enterprises"
                                else:
                                    enterprise_type = enterprise_type.title()
                        info_lookup_dict[contract_key] = enterprise_type
            
            # Prepare bulk data for Column Y
            enterprise_type_data = []
            for row, contract in contracts:
                enterprise_type = info_lookup_dict.get(contract, None)
                enterprise_type_data.append([enterprise_type])
            
            # Write enterprise type data in bulk to Column Y
            if enterprise_type_data:
                self.write_bulk_data("IA Working", 3, 25, enterprise_type_data)  # Column Y (index 24)
            
            logger.info(f"Step 4 completed: Bulk Information Request data imported for {len(contracts)} contracts")
            
        except Exception as e:
            self.log_exception("4", f"Failed to perform Information Request VLOOKUP: {str(e)}")
            raise
    
    def step_5_reorganize_special_values(self):
        """Step 5: Reorganize rows with special values"""
        try:
            logger.info("Step 5: Reorganizing rows with special values...")
            
            # Get all rows from IA Working sheet
            all_rows = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                # Read all column values for this row
                row_data = []
                for col in range(1, 26):  # Read columns A-Z (1-26)
                    value = self.read_cell_value("IA Working", row, col)
                    row_data.append(value)
                
                all_rows.append((row, row_data))
                row += 1
            
            if not all_rows:
                logger.info("No rows found in IA Working sheet")
                return
            
            # Find rows with "00" or "rg" in column B
            rows_to_move = []
            for row_num, row_data in all_rows:
                b_value = row_data[1] if len(row_data) > 1 else None  # Column B (index 1)
                if b_value in ["00", "rg"]:
                    rows_to_move.append((row_num, row_data))
            
            if rows_to_move:
                # Get the last row number to append new rows
                last_row = all_rows[-1][0] + 1
                
                # Create new rows with updated values
                new_rows_data = []
                for row_num, row_data in rows_to_move:
                    new_row_data = row_data.copy()
                    # Update values
                    if new_row_data[1] == "00":  # Column B
                        new_row_data[1] = "FDL"
                    elif new_row_data[1] == "rg":  # Column B
                        new_row_data[1] = "Margin Trading"
                    new_rows_data.append(new_row_data)
                
                # Write new rows in bulk
                for i, new_row_data in enumerate(new_rows_data):
                    target_row = last_row + i
                    # Write each column value
                    for col_idx, value in enumerate(new_row_data):
                        if col_idx < 26:  # Ensure we don't exceed column Z
                            self.write_cell_value("IA Working", target_row, col_idx + 1, value)
                
                # Clear original rows with special values
                for row_num, _ in rows_to_move:
                    # Clear the entire row
                    self.clear_range("IA Working", row_num, row_num, 1, 26)
                
                logger.info(f"Step 5 completed: {len(rows_to_move)} rows reorganized")
            else:
                logger.info("Step 5 completed: No rows to reorganize")
            
        except Exception as e:
            self.log_exception("5", f"Failed to reorganize special values: {str(e)}")
            raise
    
    def step_6_handle_na_contracts(self):
        """Step 6: Handle #N/A contracts using bot API"""
        try:
            logger.info("Step 6: Handling #N/A contracts using bot API...")
            
            # Find contracts with #N/A in column C
            na_contracts = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                c_value = self.read_cell_value("IA Working", row, 3)  # Column C
                if c_value == "#N/A":
                    na_contracts.append((row, contract))
                row += 1
            
            if na_contracts:
                # Call bot API to get contract data
                contract_numbers = [contract for _, contract in na_contracts]
                try:
                    contract_data = run_contract_search_bot(contract_numbers)
                    
                    # Update worksheet with returned data
                    for row, contract in na_contracts:
                        if contract in contract_data:
                            data = contract_data[contract]
                            
                            # Update cells with bot data
                            self.write_cell_value("IA Working", row, 3, data.get('client_code', ''))      # Column C
                            self.write_cell_value("IA Working", row, 4, data.get('equipment', ''))        # Column D
                            self.write_cell_value("IA Working", row, 10, data.get('contract_period', '')) # Column J
                            self.write_cell_value("IA Working", row, 9, data.get('frequency', ''))        # Column I
                            
                            contract_amount = data.get('contract_amount', 0)
                            at_limit = data.get('AT_limit', 0)
                            
                            if contract_amount == 0 and at_limit:
                                self.write_cell_value("IA Working", row, 13, at_limit)  # Column M
                            else:
                                self.write_cell_value("IA Working", row, 16, contract_amount)  # Column P
                    
                    logger.info(f"Step 6 completed: Updated {len(na_contracts)} contracts with bot data")
                    
                except Exception as bot_error:
                    self.log_exception("6", f"Bot data generation failed: {str(bot_error)}")
                    logger.warning("Bot data generation failed, continuing without contract data update")
            else:
                logger.info("Step 6 completed: No #N/A contracts found")
            
        except Exception as e:
            self.log_exception("6", f"Failed to handle #N/A contracts: {str(e)}")
            raise
    
    def step_7_check_blank_cells(self):
        """Step 7: Check for blank cells and report exceptions"""
        try:
            logger.info("Step 7: Checking for blank cells...")
            
            blank_cells = []
            row = 3
            
            # Check all columns except E (column 5)
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                # Check columns A-Z (1-26) except column E (5)
                for col in range(1, 27):
                    if col != 5:  # Skip column E
                        cell_value = self.read_cell_value("IA Working", row, col)
                        if cell_value is None or str(cell_value).strip() == "":
                            blank_cells.append({
                                'row': row,
                                'column': col,
                                'contract': contract
                            })
                
                row += 1
            
            if blank_cells:
                for blank in blank_cells:
                    self.log_exception("7", 
                                     f"Blank cell found at {chr(64 + blank['column'])}{blank['row']}", 
                                     blank['contract'])
            
            logger.info(f"Step 7 completed: Found {len(blank_cells)} blank cells")
            
        except Exception as e:
            self.log_exception("7", f"Failed to check blank cells: {str(e)}")
            raise
    
    def step_8_handle_special_columns(self):
        """Step 8: Handle special column logic for FDL and Margin Trading"""
        try:
            logger.info("Step 8: Handling special column logic...")
            
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                b_value = self.read_cell_value("IA Working", row, 2)  # Column B
                
                if b_value == "FDL":
                    # Clear C,D,E,I,J,K cells
                    self.clear_range("IA Working", row, row, 3, 5)   # C,D,E
                    self.clear_range("IA Working", row, row, 9, 11)  # I,J,K
                    
                    # Check if L,M cells are empty or #N/A or 0
                    l_value = self.read_cell_value("IA Working", row, 12)  # Column L
                    m_value = self.read_cell_value("IA Working", row, 13)  # Column M
                    
                    if l_value is None or str(l_value) in ['#N/A', '0'] or l_value == 0:
                        n_value = self.read_cell_value("IA Working", row, 14)  # Column N
                        o_value = self.read_cell_value("IA Working", row, 15)  # Column O
                        
                        if n_value is not None:
                            self.write_cell_value("IA Working", row, 12, n_value)  # L
                            self.write_cell_value("IA Working", row, 17, o_value)  # Q
                            self.write_cell_value("IA Working", row, 16, o_value)  # P
                            self.write_cell_value("IA Working", row, 18, n_value)  # R
                
                elif b_value == "Margin Trading":
                    # Clear C,D,E,F,I,J,K cells
                    self.clear_range("IA Working", row, row, 3, 6)   # C,D,E,F
                    self.clear_range("IA Working", row, row, 9, 11)  # I,J,K
                    
                    # Check if L,M cells are empty or #N/A or 0
                    l_value = self.read_cell_value("IA Working", row, 12)  # Column L
                    m_value = self.read_cell_value("IA Working", row, 13)  # Column M
                    
                    if l_value is None or str(l_value) in ['#N/A', '0'] or l_value == 0:
                        n_value = self.read_cell_value("IA Working", row, 14)  # Column N
                        if n_value is not None:
                            self.write_cell_value("IA Working", row, 12, n_value)  # L
                            self.write_cell_value("IA Working", row, 25, "Small")  # Y
                
                row += 1
            
            logger.info("Step 8 completed: Special column logic applied")
            
        except Exception as e:
            self.log_exception("8", f"Failed to handle special columns: {str(e)}")
            raise
    
    def step_9_handle_purpose_column(self):
        """Step 9: Handle purpose column and Product_Cat sheet"""
        try:
            logger.info("Step 9: Handling purpose column and Product_Cat sheet...")
            
            # Clear 0 or #N/A values in column E
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                e_value = self.read_cell_value("IA Working", row, 5)  # Column E
                if e_value in [0, '#N/A'] or str(e_value) in ['0', '#N/A']:
                    self.write_cell_value("IA Working", row, 5, None)
                
                row += 1
            
            # Filter column G for #N/A and copy data
            data_to_copy = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                g_value = self.read_cell_value("IA Working", row, 7)  # Column G
                if g_value == '#N/A':
                    b_value = self.read_cell_value("IA Working", row, 2)  # Column B
                    d_value = self.read_cell_value("IA Working", row, 4)  # Column D
                    e_value = self.read_cell_value("IA Working", row, 5)  # Column E
                    f_value = self.read_cell_value("IA Working", row, 6)  # Column F
                    
                    data_to_copy.append({
                        'b': b_value,
                        'd': d_value,
                        'e': e_value,
                        'f': f_value
                    })
                
                row += 1
            
            # Remove duplicates
            unique_data = []
            seen = set()
            for item in data_to_copy:
                key = (item['b'], item['d'], item['e'], item['f'])
                if key not in seen:
                    seen.add(key)
                    unique_data.append(item)
            
            # Create or update Product_Cat sheet
            try:
                # Try to get the Product_Cat worksheet
                product_cat_ws = self.get_worksheet("Product_Cat")
                logger.info("Product_Cat sheet found, updating existing data")
            except:
                # Create new Product_Cat sheet if it doesn't exist
                logger.info("Product_Cat sheet not found, creating new sheet")
                product_cat_ws = self.workbook.Worksheets.Add()
                product_cat_ws.Name = "Product_Cat"
            
            # Clear existing data and add new data
            if unique_data:
                # Clear the sheet first
                self.clear_range("Product_Cat", 1, len(unique_data) + 1, 1, 5)
                
                # Add headers
                self.write_cell_value("Product_Cat", 1, 1, "B_Value")
                self.write_cell_value("Product_Cat", 1, 2, "D_Value")
                self.write_cell_value("Product_Cat", 1, 3, "E_Value")
                self.write_cell_value("Product_Cat", 1, 4, "F_Value")
                self.write_cell_value("Product_Cat", 1, 5, "Classification")
                
                # Add data
                for i, item in enumerate(unique_data):
                    row = i + 2
                    self.write_cell_value("Product_Cat", row, 1, item['b'])  # B -> A
                    self.write_cell_value("Product_Cat", row, 2, item['d'])  # D -> B
                    self.write_cell_value("Product_Cat", row, 3, item['e'])  # E -> C
                    self.write_cell_value("Product_Cat", row, 4, item['f'])  # F -> D
                    # Classification will be filled later via WorkHub24 API
            
            logger.info(f"Step 9 completed: {len(unique_data)} unique records copied to Product_Cat")
            
        except Exception as e:
            self.log_exception("9", f"Failed to handle purpose column: {str(e)}")
            raise
    
    def step_10_filter_vehicles_machinery(self):
        """Step 10: Filter Vehicles and Machinery rows"""
        try:
            logger.info("Step 10: Filtering Vehicles and Machinery rows...")
            
            vehicles_machinery_rows = []
            row = 3
            
            # Read through IA Working sheet to find rows with "Vehicles and Machinery" in Column U
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                u_value = self.read_cell_value("IA Working", row, 21)  # Column U
                if u_value == "Vehicles and Machinery":
                    vehicles_machinery_rows.append(row)
                
                row += 1
            
            logger.info(f"Step 10 completed: Found {len(vehicles_machinery_rows)} Vehicles and Machinery rows")
            return vehicles_machinery_rows
            
        except Exception as e:
            self.log_exception("10", f"Failed to filter Vehicles and Machinery: {str(e)}")
            raise
    
    def step_11_po_listing_mapping(self):
        """Step 11: Po Listing mapping and Portfolio Recovery fallback"""
        try:
            logger.info("Step 11: Po Listing mapping and Portfolio Recovery fallback...")
            
            # Read Po Listing data using bulk method
            po_listing_data = self.read_bulk_data_from_excel(self.po_listing_file)
            
            # Read Portfolio Recovery data (if file exists)
            portfolio_recovery_data = None
            if self.portfolio_recovery_file and self.portfolio_recovery_file.exists():
                portfolio_recovery_data = self.read_bulk_data_from_excel(self.portfolio_recovery_file)
            else:
                logger.warning("Portfolio Report Recovery file not found - will skip fallback mapping")
            
            # Get Vehicles and Machinery rows
            vehicles_rows = self.step_10_filter_vehicles_machinery()
            
            # Create lookup dictionaries for faster matching
            po_lookup_dict = {}
            for row_data in po_listing_data:
                if len(row_data) > 33:  # Ensure we have enough columns for AH (index 33)
                    contract_a = row_data[0] if len(row_data) > 0 else None  # Column A
                    contract_b = row_data[1] if len(row_data) > 1 else None  # Column B
                    sell_price = row_data[33] if len(row_data) > 33 else None  # Column AH
                    if contract_a and contract_b and contract_a == contract_b:
                        po_lookup_dict[contract_a] = sell_price
            
            portfolio_lookup_dict = {}
            if portfolio_recovery_data:
                for row_data in portfolio_recovery_data:
                    if len(row_data) > 33:  # Ensure we have enough columns
                        contract_a = row_data[0] if len(row_data) > 0 else None  # Column A
                        contract_b = row_data[1] if len(row_data) > 1 else None  # Column B
                        sell_price = row_data[33] if len(row_data) > 33 else None  # Column AH
                        if contract_a and contract_b and contract_a == contract_b:
                            portfolio_lookup_dict[contract_a] = sell_price
            
            # Process each Vehicles and Machinery row
            for row in vehicles_rows:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                
                # First try Po Listing
                if contract in po_lookup_dict:
                    sell_price = po_lookup_dict[contract]
                    self.write_cell_value("IA Working", row, 22, sell_price)  # Column V (index 21)
                else:
                    # Try Portfolio Recovery as fallback
                    if contract in portfolio_lookup_dict:
                        sell_price = portfolio_lookup_dict[contract]
                        self.write_cell_value("IA Working", row, 22, sell_price)  # Column V (index 21)
                    else:
                        logger.warning(f"No Po Listing or Portfolio Recovery data found for contract {contract}")
            
            # Convert V column to numbers from row 3
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                v_value = self.read_cell_value("IA Working", row, 22)  # Column V (index 21)
                if v_value and str(v_value).replace('.', '').replace('-', '').isdigit():
                    try:
                        numeric_value = float(v_value)
                        self.write_cell_value("IA Working", row, 22, numeric_value)  # Column V
                    except ValueError:
                        pass
                
                row += 1
            
            logger.info("Step 11 completed: Po Listing mapping and Portfolio Recovery fallback")
            
        except Exception as e:
            self.log_exception("11", f"Failed to perform Po Listing mapping: {str(e)}")
            raise
    
    def step_12_valuation_bot_integration(self):
        """Step 12: Integration with IA_Working_Initial_valuation_bot"""
        try:
            logger.info("Step 12: Running valuation bot for remaining #N/A values...")
            
            # Find remaining #N/A and "Not Valued" in column V
            remaining_contracts = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                v_value = self.read_cell_value("IA Working", row, 22)  # Column V (index 21)
                if v_value in ['#N/A', 'Not Valued']:
                    remaining_contracts.append((row, contract))
                
                row += 1
            
            if remaining_contracts:
                if valuation_bot:
                    try:
                        # Import and run the valuation bot
                        contract_numbers = [contract for _, contract in remaining_contracts]
                        
                        # Run the bot (assuming it returns a dictionary with contract numbers as keys)
                        bot_results = run_valuation_bot_wrapper(contract_numbers)
                        
                        # Update worksheet with bot results
                        for row, contract in remaining_contracts:
                            if contract in bot_results:
                                self.write_cell_value("IA Working", row, 22, bot_results[contract])  # Column V
                        
                        logger.info(f"Step 12 completed: Bot updated {len(remaining_contracts)} contracts")
                        
                    except Exception as bot_error:
                        self.log_exception("12", f"Valuation bot failed: {str(bot_error)}")
                        logger.warning("Valuation bot unavailable, continuing without updates")
                else:
                    logger.warning("Valuation bot unavailable, skipping valuation bot integration.")
            
            # Check for remaining #N/A values
            final_na_contracts = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                if self.read_cell_value("IA Working", row, 22) == '#N/A':  # Column V (index 21)
                    final_na_contracts.append(contract)
                
                row += 1
            
            if final_na_contracts:
                for contract in final_na_contracts:
                    self.log_exception("12", f"Contract still has #N/A after bot run", contract)
            
            logger.info("Step 12 completed: Valuation bot integration completed")
            
        except Exception as e:
            self.log_exception("12", f"Failed to integrate valuation bot: {str(e)}")
            raise
    
    def step_13_scienter_bot_integration(self):
        """Step 13: Integration with Scienter bot and calculations"""
        try:
            logger.info("Step 13: Running Scienter bot and performing calculations...")
            
            # Find contracts with LR00000049-like values in column A
            lr_contracts = []
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                if contract and str(contract).startswith('LR') and str(contract).replace('LR', '').isdigit():
                    lr_contracts.append(row)
                
                row += 1
            
            if lr_contracts:
                # Log missing scienter bot as exception since it's not implemented
                contract_numbers = [self.read_cell_value("IA Working", row, 1) for row in lr_contracts]
                self.log_exception("13", f"Scienter bot not implemented yet - LR contracts {contract_numbers} need manual processing")
                logger.warning("Scienter bot not implemented - LR contracts will need manual processing")
                
                # Perform calculations for W and X columns (placeholder for when bot is implemented)
                for row in lr_contracts:
                    # W column: P/V calculation
                    p_value = self.read_cell_value("IA Working", row, 16)  # Column P (index 15)
                    v_value = self.read_cell_value("IA Working", row, 22)  # Column V (index 21)
                    if p_value and v_value and v_value != 0:
                        try:
                            w_value = p_value / v_value  # Column W (index 22)
                            self.write_cell_value("IA Working", row, 23, w_value)
                        except Exception as e:
                            logger.error(f"Failed to calculate W column for row {row}: {e}")
                            self.log_exception("13", f"W column calculation failed for row {row}: {e}")
                    
                    # X column: S/SUM(S:S+1)*W calculation
                    s_value = self.read_cell_value("IA Working", row, 19)  # Column S (index 18)
                    s_next_value = self.read_cell_value("IA Working", row + 1, 19) if row + 1 <= self.find_last_row("IA Working", 1) else 0
                    w_value = self.read_cell_value("IA Working", row, 23)  # Column W (index 22)
                    if s_value is not None and w_value is not None:
                        try:
                            sum_s = s_value + (s_next_value or 0)
                            if sum_s != 0:
                                x_value = (s_value / sum_s) * w_value  # Column X (index 23)
                                self.write_cell_value("IA Working", row, 24, x_value)
                        except Exception as e:
                            logger.error(f"Failed to calculate X column for row {row}: {e}")
                            self.log_exception("13", f"X column calculation failed for row {row}: {e}")
                
                logger.info(f"Step 13 completed: LR contracts logged for manual processing")
            
            logger.info("Step 13 completed: Scienter bot integration completed")
            
        except Exception as e:
            self.log_exception("13", f"Failed to integrate Scienter bot: {str(e)}")
            raise
    
    def step_14_minimum_rate_api(self):
        """Step 14: Import minimum rate from WorkHub24 API"""
        try:
            logger.info("Step 14: Importing minimum rate from WorkHub24 API...")
            
            # TODO: Implement WorkHub24 API call
            # For now, log this as an exception since the API is not implemented
            self.log_exception("14", "WorkHub24 API for minimum rate not implemented yet - needs manual input")
            logger.warning("WorkHub24 API not implemented - minimum rate needs manual input")
            
            # Placeholder for minimum rate
            self.minimum_rate = 15.00  # Default value, should come from API
            logger.info(f"Using default minimum rate: {self.minimum_rate}%")
            
        except Exception as e:
            self.log_exception("14", f"Failed to import minimum rate: {str(e)}")
            self.log_exception("14", f"Step 14 failed: {str(e)}")
            # Set a default minimum rate to continue processing
            self.minimum_rate = 15.00
    
    def step_15_check_interest_rates(self):
        """Step 15: Check interest rates against minimum rate"""
        try:
            logger.info("Step 15: Checking interest rates against minimum rate...")
            
            # TODO: Get minimum rate from step 14
            minimum_rate = 15.00  # Placeholder - should come from step 14
            
            # Check L column values from row 3
            row = 3
            while True:
                contract = self.read_cell_value("IA Working", row, 1)  # Column A
                if not contract:
                    break
                
                l_value = self.read_cell_value("IA Working", row, 12)  # Column L (index 11)
                if l_value and l_value != '#N/A':
                    try:
                        rate = float(l_value)
                        if rate > minimum_rate:
                            self.log_exception("15", f"Interest rate {rate}% exceeds minimum {minimum_rate}%", contract)
                    except (ValueError, TypeError):
                        pass
                
                row += 1
            
            logger.info("Step 15 completed: Interest rate validation completed")
            
        except Exception as e:
            self.log_exception("15", f"Failed to check interest rates: {str(e)}")
            raise
    
    def step_16_final_validation(self):
        """Step 16: Final validation of C39 cell"""
        try:
            logger.info("Step 16: Final validation of C39 cell...")
            
            # Check C39 cell value directly from the worksheet
            try:
                c39_value = self.read_cell_value("NBD-MF-23-IA", 39, 3)  # Row 39, Column C (index 3)
                
                if c39_value != 0:
                    self.log_exception("16", f"C39 cell value is {c39_value}, expected 0")
                    logger.error(f"Final validation failed: C39 = {c39_value}")
                else:
                    logger.info("Step 16 completed: C39 cell validation passed")
                    
            except Exception as sheet_error:
                self.log_exception("16", f"Could not read C39 cell: {str(sheet_error)}")
                logger.error("NBD-MF-23-IA sheet structure insufficient for C39 validation")
            
        except Exception as e:
            self.log_exception("16", f"Failed to validate C39 cell: {str(e)}")
            raise
    
    def generate_exception_report(self):
        """Generate Excel report of all exceptions"""
        try:
            if not self.exceptions:
                logger.info("No exceptions to report")
                return
            
            # Validate that all exceptions are properly formatted dictionaries
            valid_exceptions = []
            for i, exception in enumerate(self.exceptions):
                if isinstance(exception, dict) and 'step' in exception and 'message' in exception:
                    valid_exceptions.append(exception)
                else:
                    logger.warning(f"Invalid exception format at index {i}: {exception}")
                    # Convert invalid exceptions to proper format
                    if isinstance(exception, str):
                        valid_exceptions.append({
                            'step': 'Unknown',
                            'message': exception,
                            'contract_number': None,
                            'timestamp': datetime.now().isoformat()
                        })
                    else:
                        logger.error(f"Cannot convert exception at index {i} to proper format: {exception}")
            
            if not valid_exceptions:
                logger.info("No valid exceptions to report after validation")
                return
            
            # Create exception report
            report_file = self.working_dir / f"NBD_MF_23_IA_Exceptions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            
            df_exceptions = pd.DataFrame(valid_exceptions)
            
            with pd.ExcelWriter(report_file, engine='openpyxl') as writer:
                df_exceptions.to_excel(writer, sheet_name='Exceptions', index=False)
            
            logger.info(f"Exception report generated: {report_file}")
            
        except Exception as e:
            logger.error(f"Failed to generate exception report: {str(e)}")
            # Log the current state of exceptions for debugging
            logger.error(f"Current exceptions state: {self.exceptions}")
            logger.error(f"Exceptions type: {type(self.exceptions)}")
            if self.exceptions:
                logger.error(f"First exception: {self.exceptions[0]} (type: {type(self.exceptions[0])})")
    
    def run_automation(self):
        """Run the complete automation workflow"""
        try:
            logger.info("Starting NBD-MF-23-IA automation...")
            
            # Initialize Excel COM application
            self.initialize_excel()
            
            # Optimize Excel performance for bulk operations
            self.optimize_excel_performance()
            
            # Execute all steps
            self.step_1_copy_disbursement_data()
            self.step_2_vlookup_net_portfolio()
            self.step_3_po_listing_vlookup()
            self.step_4_info_request_vlookup()
            self.step_5_reorganize_special_values()
            self.step_6_handle_na_contracts()
            self.step_7_check_blank_cells()
            self.step_8_handle_special_columns()
            self.step_9_handle_purpose_column()
            self.step_10_filter_vehicles_machinery()
            self.step_11_po_listing_mapping()
            self.step_12_valuation_bot_integration()
            self.step_13_scienter_bot_integration()
            self.step_14_minimum_rate_api()
            self.step_15_check_interest_rates()
            self.step_16_final_validation()
            
            # Generate exception report
            self.generate_exception_report()
            
            logger.info("NBD-MF-23-IA automation completed successfully!")
            
        except Exception as e:
            logger.error(f"Automation failed: {str(e)}")
            self.generate_exception_report()
            raise
        finally:
            # Restore Excel performance settings
            self.restore_excel_performance()
            # Always close Excel properly
            self.close_excel()

    def write_bulk_data(self, sheet_name: str, start_row: int, start_col: int, data):
        """Write bulk data to Excel starting from specified position"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            
            # Convert data to list if it's not already
            if not isinstance(data, list):
                data = [data]
            
            # If data is 2D (list of lists), write as range
            if data and isinstance(data[0], list):
                # Calculate range
                end_row = start_row + len(data) - 1
                end_col = start_col + len(data[0]) - 1
                range_str = f"{chr(64 + start_col)}{start_row}:{chr(64 + end_col)}{end_row}"
                
                # Write entire range at once
                worksheet.Range(range_str).Value = data
                logger.info(f"Bulk wrote {len(data)}x{len(data[0])} data to {range_str}")
            else:
                # Single column data
                end_row = start_row + len(data) - 1
                range_str = f"{chr(64 + start_col)}{start_row}:{chr(64 + start_col)}{end_row}"
                
                # Convert to 2D format for Excel
                data_2d = [[item] for item in data]
                worksheet.Range(range_str).Value = data_2d
                logger.info(f"Bulk wrote {len(data)} values to column {chr(64 + start_col)} rows {start_row}-{end_row}")
                
        except Exception as e:
            logger.error(f"Failed to write bulk data: {e}")
            raise
    
    def read_bulk_data_from_excel(self, file_path: str, sheet_name: str = None, usecols: str = None):
        """Read bulk data from Excel file using pandas for better performance"""
        try:
            logger.info(f"Starting to read Excel file: {file_path}")
            
            # Use optimized reading parameters for large files
            read_params = {
                'sheet_name': sheet_name,
                'usecols': usecols,
                'nrows': None,  # Read all rows
                'skiprows': None,
                'header': None,  # No header to avoid issues
                'engine': 'openpyxl'  # Explicitly use openpyxl for better performance
            }
            
            # Remove None values
            read_params = {k: v for k, v in read_params.items() if v is not None}
            
            logger.info(f"Reading with parameters: {read_params}")
            df = pd.read_excel(file_path, **read_params)
            
            logger.info(f"Successfully read DataFrame with shape: {df.shape}")
            
            # Convert to list of lists for bulk writing
            # Use a more robust approach that handles different pandas versions
            try:
                # Try the most reliable method first
                data = df.values.tolist()
                logger.info(f"Converted to list using df.values.tolist()")
            except Exception as e1:
                logger.warning(f"df.values.tolist() failed: {e1}, trying alternative method")
                try:
                    # Alternative method for newer pandas versions
                    data = df.to_numpy().tolist()
                    logger.info(f"Converted to list using df.to_numpy().tolist()")
                except Exception as e2:
                    logger.warning(f"df.to_numpy().tolist() failed: {e2}, using row-by-row conversion")
                    # Last resort: convert row by row
                    data = []
                    for index, row in df.iterrows():
                        data.append(row.tolist())
                    logger.info(f"Converted to list using row-by-row iteration")
            
            logger.info(f"Successfully read {len(data)} rows from {file_path}")
            return data
            
        except Exception as e:
            logger.error(f"Failed to read Excel file {file_path}: {e}")
            raise
    
    def bulk_vlookup_operation(self, sheet_name: str, target_col: int, lookup_col: int, 
                              source_data, source_lookup_col: int, source_value_col: int, 
                              start_row: int, end_row: int):
        """Perform bulk VLOOKUP operation for better performance"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            
            # Create lookup dictionary for faster matching
            lookup_dict = {}
            for row_data in source_data:
                if len(row_data) > max(source_lookup_col, source_value_col):
                    lookup_key = row_data[source_lookup_col]
                    lookup_value = row_data[source_value_col]
                    lookup_dict[lookup_key] = lookup_value
            
            # Prepare data for bulk writing
            result_data = []
            for row in range(start_row, end_row + 1):
                lookup_value = worksheet.Cells(row, lookup_col).Value
                if lookup_value in lookup_dict:
                    result_data.append([lookup_dict[lookup_value]])
                else:
                    result_data.append([None])
            
            # Write results in bulk
            self.write_bulk_data(sheet_name, start_row, target_col, result_data)
            
            logger.info(f"Bulk VLOOKUP completed for {len(result_data)} rows")
            
        except Exception as e:
            logger.error(f"Failed to perform bulk VLOOKUP: {e}")
            raise
    
    def bulk_copy_range_with_filter(self, source_sheet: str, target_sheet: str, 
                                   source_range: str, target_start_row: int, target_start_col: int,
                                   filter_col: int = None, filter_value: str = None):
        """Copy a range of data with optional filtering in bulk"""
        try:
            source_worksheet = self.get_worksheet(source_sheet)
            target_worksheet = self.get_worksheet(target_sheet)
            
            # Read source range
            source_range_obj = source_worksheet.Range(source_range)
            source_data = source_range_obj.Value
            
            if not source_data:
                logger.warning(f"No data found in source range {source_range}")
                return
            
            # Filter data if filter is specified
            if filter_col is not None and filter_value is not None:
                filtered_data = []
                for row_data in source_data:
                    if isinstance(row_data, list) and len(row_data) > filter_col:
                        if row_data[filter_col] == filter_value:
                            filtered_data.append(row_data)
                source_data = filtered_data
            
            if not source_data:
                logger.info(f"No data matches filter criteria: {filter_col}={filter_value}")
                return
            
            # Calculate target range
            rows = len(source_data)
            cols = len(source_data[0]) if source_data else 0
            target_end_row = target_start_row + rows - 1
            target_end_col = target_start_col + cols - 1
            target_range = f"{chr(64 + target_start_col)}{target_start_row}:{chr(64 + target_end_col)}{target_end_row}"
            
            # Write filtered data in bulk
            target_worksheet.Range(target_range).Value = source_data
            
            logger.info(f"Bulk copied {rows}x{cols} filtered data to {target_range}")
            
        except Exception as e:
            logger.error(f"Failed to bulk copy range: {e}")
            raise
    
    def bulk_clear_and_fill(self, sheet_name: str, clear_ranges: list, fill_data: dict):
        """Clear multiple ranges and fill with data in bulk operations"""
        try:
            worksheet = self.get_worksheet(sheet_name)
            
            # Clear all specified ranges
            for range_str in clear_ranges:
                worksheet.Range(range_str).ClearContents()
                logger.info(f"Cleared range: {range_str}")
            
            # Fill data in bulk
            for range_str, data in fill_data.items():
                if isinstance(data, list):
                    worksheet.Range(range_str).Value = data
                    logger.info(f"Filled range {range_str} with {len(data)} rows")
                else:
                    worksheet.Range(range_str).Value = data
                    logger.info(f"Filled range {range_str} with single value")
            
        except Exception as e:
            logger.error(f"Failed to bulk clear and fill: {e}")
            raise
    
    def optimize_excel_performance(self):
        """Optimize Excel performance settings for bulk operations"""
        try:
            if self.excel_app:
                # Disable screen updating for faster operations
                self.excel_app.ScreenUpdating = False
                
                # Disable automatic calculations
                self.excel_app.Calculation = -4105  # xlCalculationManual
                
                # Disable events
                self.excel_app.EnableEvents = False
                
                logger.info("Excel performance optimized for bulk operations")
                
        except Exception as e:
            logger.warning(f"Could not optimize Excel performance: {e}")
    
    def restore_excel_performance(self):
        """Restore Excel performance settings after bulk operations"""
        try:
            if self.excel_app:
                # Re-enable screen updating
                self.excel_app.ScreenUpdating = True
                
                # Re-enable automatic calculations
                self.excel_app.Calculation = -4105  # xlCalculationAutomatic
                
                # Re-enable events
                self.excel_app.EnableEvents = True
                
                logger.info("Excel performance settings restored")
                
        except Exception as e:
            logger.warning(f"Could not restore Excel performance: {e}")


def main():
    """Main function to run the automation"""
    # Default values so you don't have to pass arguments every time
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    month = "July"
    year = "2025"

    # If arguments are provided, they override the defaults
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--working-dir", default=working_dir, help="Working directory")
    parser.add_argument("--month", default=month, help="Month (e.g., Jan, Feb)")
    parser.add_argument("--year", default=year, help="Year")
    args = parser.parse_args()

    # Run automation
    automation = NBDMF23IAAutomation(args.working_dir, args.month, args.year)
    
    # Check module status before starting
    print("\n" + "="*50)
    print("MODULE STATUS CHECK")
    print("="*50)
    automation.check_module_status()
    
    try:
        # Run all steps
        automation.run_automation()
        logger.info("All steps completed successfully!")
        
        # Show exception summary
        print("\n" + "="*50)
        print(automation.get_exception_summary())
        print("="*50)
        
    except Exception as e:
        logger.error(f"Automation failed: {e}")
        # Generate exception report even if automation fails
        automation.generate_exception_report()
        
        # Show exception summary even on failure
        print("\n" + "="*50)
        print(automation.get_exception_summary())
        print("="*50)
    finally:
        # Ensure Excel is properly closed
        try:
            automation.close_excel()
        except:
            pass


if __name__ == "__main__":
    main()
