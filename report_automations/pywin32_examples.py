#!/usr/bin/env python3
"""
Comprehensive examples of using pywin32 for Excel automation
This file demonstrates various techniques for manipulating .xlsb files directly
"""

import sys
import logging
from pathlib import Path

# Add the current directory to path to import the main automation class
sys.path.append(str(Path(__file__).parent))

from NBD_MF_23_IA import NBDMF23IAAutomation

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def example_1_basic_cell_operations():
    """Example 1: Basic cell reading and writing"""
    print("\n" + "="*60)
    print("EXAMPLE 1: Basic Cell Operations")
    print("="*60)
    
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    automation = NBDMF23IAAutomation("July", "2025", working_dir)
    
    try:
        automation.initialize_excel()
        
        # Write values to cells
        automation.write_cell_value("IA Working", 1, 1, "Contract Number")
        automation.write_cell_value("IA Working", 1, 2, "Status")
        automation.write_cell_value("IA Working", 1, 3, "Amount")
        
        # Read values back
        contract_header = automation.read_cell_value("IA Working", 1, 1)
        status_header = automation.read_cell_value("IA Working", 1, 2)
        amount_header = automation.read_cell_value("IA Working", 1, 3)
        
        print(f"Headers written: {contract_header}, {status_header}, {amount_header}")
        
    finally:
        automation.close_excel()

def example_2_formula_operations():
    """Example 2: Working with formulas"""
    print("\n" + "="*60)
    print("EXAMPLE 2: Formula Operations")
    print("="*60)
    
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    automation = NBDMF23IAAutomation("July", "2025", working_dir)
    
    try:
        automation.initialize_excel()
        
        # Write some data first
        automation.write_cell_value("IA Working", 2, 1, "CON001")
        automation.write_cell_value("IA Working", 2, 2, 1000)
        automation.write_cell_value("IA Working", 2, 3, 500)
        
        # Write formulas
        automation.write_cell_formula("IA Working", 2, 4, "=B2+C2")  # Sum
        automation.write_cell_formula("IA Working", 2, 5, "=B2*0.1")  # 10% of B2
        
        print("Formulas written: =B2+C2 (sum) and =B2*0.1 (10%)")
        
    finally:
        automation.close_excel()

def example_3_fill_down_formulas():
    """Example 3: Fill Down (Ctrl+D) functionality"""
    print("\n" + "="*60)
    print("EXAMPLE 3: Fill Down (Ctrl+D) Formulas")
    print("="*60)
    
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    automation = NBDMF23IAAutomation("July", "2025", working_dir)
    
    try:
        automation.initialize_excel()
        
        # Write data in multiple rows
        for row in range(2, 7):
            automation.write_cell_value("IA Working", row, 1, f"CON{row:03d}")
            automation.write_cell_value("IA Working", row, 2, row * 1000)
            automation.write_cell_value("IA Working", row, 3, row * 500)
        
        # Write formula to first row
        automation.write_cell_formula("IA Working", 2, 4, "=B2+C2")
        
        # Fill down the formula to all rows (Ctrl+D equivalent)
        automation.fill_down_formula("IA Working", 2, 6, 4, "=B2+C2")
        
        print("Formula =B2+C2 filled down from row 2 to row 6")
        
    finally:
        automation.close_excel()

def example_4_vlookup_formulas():
    """Example 4: VLOOKUP formulas with Fill Down"""
    print("\n" + "="*60)
    print("EXAMPLE 4: VLOOKUP Formulas with Fill Down")
    print("="*60)
    
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    automation = NBDMF23IAAutomation("July", "2025", working_dir)
    
    try:
        automation.initialize_excel()
        
        # Create lookup table in another sheet (simulating Net Portfolio)
        automation.write_cell_value("NetPortfolio", 1, 1, "Contract")
        automation.write_cell_value("NetPortfolio", 1, 2, "Client")
        automation.write_cell_value("NetPortfolio", 1, 3, "Equipment")
        
        # Add some lookup data
        automation.write_cell_value("NetPortfolio", 2, 1, "CON001")
        automation.write_cell_value("NetPortfolio", 2, 2, "Client A")
        automation.write_cell_value("NetPortfolio", 2, 3, "Truck")
        
        automation.write_cell_value("NetPortfolio", 3, 1, "CON002")
        automation.write_cell_value("NetPortfolio", 3, 2, "Client B")
        automation.write_cell_value("NetPortfolio", 3, 3, "Car")
        
        # Write VLOOKUP formula to first row
        automation.write_cell_formula("IA Working", 2, 5, '=VLOOKUP(A2,NetPortfolio!A:C,2,FALSE)')
        
        # Fill down the VLOOKUP formula
        automation.fill_down_formula("IA Working", 2, 6, 5, '=VLOOKUP(A2,NetPortfolio!A:C,2,FALSE)')
        
        print("VLOOKUP formulas applied to look up client names")
        
    finally:
        automation.close_excel()

def example_5_calculation_formulas():
    """Example 5: Complex calculation formulas"""
    print("\n" + "="*60)
    print("EXAMPLE 5: Complex Calculation Formulas")
    print("="*60)
    
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    automation = NBDMF23IAAutomation("July", "2025", working_dir)
    
    try:
        automation.initialize_excel()
        
        # Write some financial data
        for row in range(2, 7):
            automation.write_cell_value("IA Working", row, 1, f"CON{row:03d}")
            automation.write_cell_value("IA Working", row, 2, row * 10000)  # Principal
            automation.write_cell_value("IA Working", row, 3, 0.15)         # Interest rate
        
        # Write calculation formulas
        # Monthly payment: =PMT(rate/12, term, principal)
        automation.write_cell_formula("IA Working", 2, 4, "=PMT(C2/12, 36, B2)")
        
        # Total interest: =principal * rate * term
        automation.write_cell_formula("IA Working", 2, 5, "=B2*C2*3")
        
        # Fill down the formulas
        automation.fill_down_formula("IA Working", 2, 6, 4, "=PMT(C2/12, 36, B2)")
        automation.fill_down_formula("IA Working", 2, 6, 5, "=B2*C2*3")
        
        print("Financial calculation formulas applied and filled down")
        
    finally:
        automation.close_excel()

def example_6_range_operations():
    """Example 6: Range operations and bulk updates"""
    print("\n" + "="*60)
    print("EXAMPLE 6: Range Operations")
    print("="*60)
    
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    automation = NBDMF23IAAutomation("July", "2025", working_dir)
    
    try:
        automation.initialize_excel()
        
        # Clear a range of cells
        automation.clear_range("IA Working", 10, 15, 1, 5)
        print("Cleared range A10:E15")
        
        # Find the last row with data
        last_row = automation.find_last_row("IA Working", 1)
        print(f"Last row with data in column A: {last_row}")
        
        # Read a range of values
        if last_row >= 5:
            values = automation.read_range_values("IA Working", 1, 5, 1, 3)
            print(f"Read range A1:C5: {values}")
        
    finally:
        automation.close_excel()

def run_all_examples():
    """Run all the pywin32 examples"""
    print("PYWIN32 EXCEL AUTOMATION EXAMPLES")
    print("="*60)
    print("This script demonstrates various pywin32 features:")
    print("1. Basic cell operations (read/write)")
    print("2. Formula operations")
    print("3. Fill Down (Ctrl+D) functionality")
    print("4. VLOOKUP formulas with Fill Down")
    print("5. Complex calculation formulas")
    print("6. Range operations")
    print("="*60)
    
    try:
        example_1_basic_cell_operations()
        example_2_formula_operations()
        example_3_fill_down_formulas()
        example_4_vlookup_formulas()
        example_5_calculation_formulas()
        example_6_range_operations()
        
        print("\n" + "="*60)
        print("ALL EXAMPLES COMPLETED SUCCESSFULLY!")
        print("="*60)
        print("\nKey pywin32 features demonstrated:")
        print("✅ Direct .xlsb file manipulation")
        print("✅ Cell reading and writing")
        print("✅ Formula application")
        print("✅ Fill Down (Ctrl+D) functionality")
        print("✅ Range operations")
        print("✅ VLOOKUP and calculation formulas")
        
    except Exception as e:
        print(f"\nExamples failed: {e}")
        logger.error(f"Examples failed: {e}")
        raise

if __name__ == "__main__":
    run_all_examples()
