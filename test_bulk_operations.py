#!/usr/bin/env python3
"""
Test script for bulk operations in NBD_MF_23_IA.py
This script demonstrates the performance improvements from bulk operations.
"""

import time
import logging
from pathlib import Path
import sys

# Add the report_automations directory to path
sys.path.append(str(Path(__file__).parent / "report_automations"))

from NBD_MF_23_IA import NBDMF23IAAutomation

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def test_bulk_operations():
    """Test the bulk operations functionality"""
    
    # Create automation instance
    working_dir = Path("working/monthly/08-01-2025(1)/NBD_MF_23_IA")
    automation = NBDMF23IAAutomation(working_dir, "Jul", 2025)
    
    try:
        logger.info("Testing bulk operations...")
        
        # Test 1: Bulk data writing
        logger.info("Test 1: Bulk data writing")
        start_time = time.time()
        
        # Create sample data
        sample_data = [[f"Contract{i}", f"Value{i}", f"Rate{i}"] for i in range(1, 1001)]
        
        # Initialize Excel
        automation.initialize_excel()
        
        # Test bulk write
        automation.write_bulk_data("IA Working", 3, 1, sample_data)
        
        end_time = time.time()
        logger.info(f"Bulk write completed in {end_time - start_time:.2f} seconds")
        
        # Test 2: Bulk data reading
        logger.info("Test 2: Bulk data reading")
        start_time = time.time()
        
        # Read data back
        read_data = automation.read_range_values("IA Working", 3, 1002, 1, 3)
        
        end_time = time.time()
        logger.info(f"Bulk read completed in {end_time - start_time:.2f} seconds")
        logger.info(f"Read {len(read_data)} rows of data")
        
        # Test 3: Bulk VLOOKUP operation
        logger.info("Test 3: Bulk VLOOKUP operation")
        start_time = time.time()
        
        # Create lookup data
        lookup_data = [[f"Contract{i}", f"Client{i}", f"Equipment{i}"] for i in range(1, 1001)]
        
        # Perform bulk VLOOKUP
        automation.bulk_vlookup_operation(
            "IA Working", 4, 1,  # target_col=4 (D), lookup_col=1 (A)
            lookup_data, 0, 2,   # source_data, source_lookup_col=0, source_value_col=2
            3, 1002              # start_row=3, end_row=1002
        )
        
        end_time = time.time()
        logger.info(f"Bulk VLOOKUP completed in {end_time - start_time:.2f} seconds")
        
        # Test 4: Bulk range operations
        logger.info("Test 4: Bulk range operations")
        start_time = time.time()
        
        # Clear a range
        automation.clear_range("IA Working", 1000, 1002, 1, 3)
        
        # Fill data in bulk
        fill_data = {"A1000:C1002": [["New1", "New2", "New3"], ["New4", "New5", "New6"], ["New7", "New8", "New9"]]}
        automation.bulk_clear_and_fill("IA Working", [], fill_data)
        
        end_time = time.time()
        logger.info(f"Bulk range operations completed in {end_time - start_time:.2f} seconds")
        
        logger.info("All bulk operation tests completed successfully!")
        
    except Exception as e:
        logger.error(f"Test failed: {e}")
        raise
    finally:
        # Close Excel
        automation.close_excel()

def test_performance_comparison():
    """Compare performance between old and new methods"""
    
    working_dir = Path("working/monthly/08-01-2025(1)/NBD_MF_23_IA")
    automation = NBDMF23IAAutomation(working_dir, "Jul", 2025)
    
    try:
        automation.initialize_excel()
        
        # Test old method (cell by cell)
        logger.info("Testing old method (cell by cell)...")
        start_time = time.time()
        
        for i in range(1, 101):
            automation.write_cell_value("IA Working", i, 1, f"Old{i}")
        
        end_time = time.time()
        old_method_time = end_time - start_time
        logger.info(f"Old method took {old_method_time:.2f} seconds")
        
        # Test new method (bulk)
        logger.info("Testing new method (bulk)...")
        start_time = time.time()
        
        bulk_data = [[f"New{i}"] for i in range(1, 101)]
        automation.write_bulk_data("IA Working", 1, 2, bulk_data)
        
        end_time = time.time()
        new_method_time = end_time - start_time
        logger.info(f"New method took {new_method_time:.2f} seconds")
        
        # Calculate improvement
        improvement = ((old_method_time - new_method_time) / old_method_time) * 100
        logger.info(f"Performance improvement: {improvement:.1f}%")
        
    except Exception as e:
        logger.error(f"Performance test failed: {e}")
        raise
    finally:
        automation.close_excel()

def main():
    """Main test function"""
    logger.info("Starting bulk operations tests...")
    
    try:
        # Test basic bulk operations
        test_bulk_operations()
        
        # Test performance comparison
        test_performance_comparison()
        
        logger.info("All tests completed successfully!")
        
    except Exception as e:
        logger.error(f"Test suite failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
