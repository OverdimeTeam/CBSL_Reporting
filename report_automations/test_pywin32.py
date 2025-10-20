#!/usr/bin/env python3
"""
Test script to demonstrate pywin32 Excel automation features
This script shows how to use pywin32 to manipulate .xlsb files directly
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

def test_pywin32_features():
    """Test the pywin32 Excel automation features"""
    
    # Initialize automation with test directory
    working_dir = r"..\working\monthly\08-01-2025(2)\NBD_MF_23_IA"
    month = "July"
    year = "2025"
    
    automation = NBDMF23IAAutomation(working_dir, month, year)
    
    try:
        # Initialize Excel
        automation.initialize_excel()
        logger.info("Excel initialized successfully")
        
        # Test basic cell operations
        logger.info("Testing basic cell operations...")
        
        # Write some test data
        automation.write_cell_value("IA Working", 1, 1, "Test Contract")
        automation.write_cell_value("IA Working", 1, 2, "Test Status")
        
        # Read back the data
        value1 = automation.read_cell_value("IA Working", 1, 1)
        value2 = automation.read_cell_value("IA Working", 1, 2)
        
        logger.info(f"Read values: {value1}, {value2}")
        
        # Test formula writing
        logger.info("Testing formula operations...")
        automation.write_cell_formula("IA Working", 2, 1, "=A1")
        automation.write_cell_formula("IA Working", 2, 2, "=B1")
        
        # Test Fill Down (Ctrl+D equivalent)
        logger.info("Testing Fill Down functionality...")
        automation.fill_down_formula("IA Working", 2, 5, 1, "=A1")
        automation.fill_down_formula("IA Working", 2, 5, 2, "=B1")
        
        # Test range operations
        logger.info("Testing range operations...")
        automation.clear_range("IA Working", 10, 15, 1, 3)
        
        # Find last row
        last_row = automation.find_last_row("IA Working", 1)
        logger.info(f"Last row with data in column A: {last_row}")
        
        # Demonstrate VLOOKUP formula application
        logger.info("Testing VLOOKUP formula application...")
        # Note: This is commented out as it requires actual data ranges
        # automation.apply_vlookup_formula("IA Working", 3, 1, "NetPortfolio!A:C", 2, 3, 10)
        
        # Demonstrate calculation formula application
        logger.info("Testing calculation formula application...")
        # Note: This is commented out as it requires actual data
        # automation.apply_calculation_formula("IA Working", 22, "=P3/V3", 3, 10)
        
        logger.info("All pywin32 tests completed successfully!")
        
    except Exception as e:
        logger.error(f"Test failed: {e}")
        raise
    finally:
        # Always close Excel
        automation.close_excel()
        logger.info("Excel closed")

if __name__ == "__main__":
    try:
        test_pywin32_features()
        print("\n" + "="*50)
        print("PYWIN32 TEST COMPLETED SUCCESSFULLY!")
        print("="*50)
    except Exception as e:
        print(f"\nTest failed: {e}")
        sys.exit(1)
