"""
File Manager for handling report storage with date-based folder structure
"""
import os
import logging
from typing import Dict, Any, List
from datetime import datetime
import pandas as pd

logger = logging.getLogger(__name__)

class FileManager:
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.base_path = config.get('base_path', './data')
        self.rupee_folder = config.get('rupee_folder', 'rupee_value')
        self.thousands_folder = config.get('thousands_folder', 'thousands_value')
        
    def _get_date_folder(self) -> str:
        """Get today's date folder path (YYYY-MM-DD)"""
        today = datetime.now().strftime('%Y-%m-%d')
        return os.path.join(self.base_path, today)
    
    def _create_date_structure(self) -> str:
        """Create date folder structure with subfolders"""
        date_folder = self._get_date_folder()
        
        # Create main date folder
        os.makedirs(date_folder, exist_ok=True)
        
        # Create rupee_value subfolder
        rupee_path = os.path.join(date_folder, self.rupee_folder)
        os.makedirs(rupee_path, exist_ok=True)
        
        # Create thousands_value subfolder
        thousands_path = os.path.join(date_folder, self.thousands_folder)
        os.makedirs(thousands_path, exist_ok=True)
        
        logger.info(f"Created date structure: {date_folder}")
        return date_folder
    
    def save_report(self, report_name: str, data: pd.DataFrame) -> Dict[str, str]:
        """Save report in both rupee_value and thousands_value folders"""
        # TODO: implement report saving logic
        try:
            # Create date structure
            date_folder = self._create_date_structure()
            
            # Generate timestamp for filename
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"{report_name}_{timestamp}.xlsx"
            
            # Save in rupee_value folder
            rupee_path = os.path.join(date_folder, self.rupee_folder, filename)
            data.to_excel(rupee_path, index=False, engine='openpyxl')
            
            # Create thousands version (divide numeric columns by 1000)
            thousands_data = data.copy()
            numeric_columns = thousands_data.select_dtypes(include=['number']).columns
            thousands_data[numeric_columns] = thousands_data[numeric_columns] / 1000
            
            # Save in thousands_value folder
            thousands_path = os.path.join(date_folder, self.thousands_folder, filename)
            thousands_data.to_excel(thousands_path, index=False, engine='openpyxl')
            
            result_paths = {
                'rupee_value': rupee_path,
                'thousands_value': thousands_path
            }
            
            logger.info(f"Report {report_name} saved successfully")
            logger.info(f"Rupee version: {rupee_path}")
            logger.info(f"Thousands version: {thousands_path}")
            
            return result_paths
            
        except Exception as e:
            logger.error(f"Failed to save report {report_name}: {e}")
            raise
    
    def get_report_files(self, date_str: str = None) -> List[Dict[str, str]]:
        """Get list of report files for a specific date"""
        # TODO: implement file listing
        if not date_str:
            date_str = datetime.now().strftime('%Y-%m-%d')
        
        date_folder = os.path.join(self.base_path, date_str)
        if not os.path.exists(date_folder):
            return []
        
        files = []
        
        # Check rupee_value folder
        rupee_path = os.path.join(date_folder, self.rupee_folder)
        if os.path.exists(rupee_path):
            for filename in os.listdir(rupee_path):
                if filename.endswith('.xlsx'):
                    files.append({
                        'filename': filename,
                        'type': 'rupee_value',
                        'path': os.path.join(rupee_path, filename),
                        'date': date_str
                    })
        
        # Check thousands_value folder
        thousands_path = os.path.join(date_folder, self.thousands_folder)
        if os.path.exists(thousands_path):
            for filename in os.listdir(thousands_path):
                if filename.endswith('.xlsx'):
                    files.append({
                        'filename': filename,
                        'type': 'thousands_value',
                        'path': os.path.join(thousands_path, filename),
                        'date': date_str
                    })
        
        return files
    
    def delete_old_files(self, days_to_keep: int = 30):
        """Delete files older than specified days"""
        # TODO: implement file cleanup
        try:
            cutoff_date = datetime.now() - pd.Timedelta(days=days_to_keep)
            
            if not os.path.exists(self.base_path):
                return
            
            for folder_name in os.listdir(self.base_path):
                folder_path = os.path.join(self.base_path, folder_name)
                
                if os.path.isdir(folder_path):
                    try:
                        folder_date = datetime.strptime(folder_name, '%Y-%m-%d')
                        if folder_date < cutoff_date:
                            import shutil
                            shutil.rmtree(folder_path)
                            logger.info(f"Deleted old folder: {folder_path}")
                    except ValueError:
                        # Skip folders that don't match date format
                        continue
                        
        except Exception as e:
            logger.error(f"Error during file cleanup: {e}")
    
    def create_dummy_report(self, report_name: str) -> pd.DataFrame:
        """Create dummy report for testing"""
        # TODO: remove this method in production
        dummy_data = {
            'id': [1, 2, 3, 4, 5],
            'name': ['Item A', 'Item B', 'Item C', 'Item D', 'Item E'],
            'value': [1000, 2000, 1500, 3000, 2500],
            'category': ['Type 1', 'Type 2', 'Type 1', 'Type 3', 'Type 2'],
            'date': pd.Timestamp.now()
        }
        
        df = pd.DataFrame(dummy_data)
        logger.info(f"Created dummy report {report_name} with {len(df)} rows")
        return df