#!/usr/bin/env python3
"""
IA Working Initial Valuation Bot
Extracts valuation amounts from Equipment details API response for contract numbers
"""

import os
import time
import logging
import json
import base64
import requests
from datetime import datetime
from dotenv import load_dotenv
from cryptography.fernet import Fernet
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Load environment variables from parent directory
load_dotenv('../.env')

# ---------------- Config ----------------
BASE_URL = "https://erp.assetline.lk/"
RECOVERY_URL = "https://erp.assetline.lk/Application/Home/RECOVERY"
LOGS_DIR = "logs"

# Create logs directory
os.makedirs(LOGS_DIR, exist_ok=True)

def load_contract_numbers_from_excel():
    """Load contract numbers from the Excel file used by NBD_MF_23_IA.py"""
    try:
        import pandas as pd
        from pathlib import Path
        
        # Get the parent directory (Script directory)
        script_dir = Path(__file__).parent.parent
        
        # Path to the Excel file (same as used in NBD_MF_23_IA.py)
        excel_file = script_dir / "Prod. wise Class. of Loans - Jul 2025.xlsb"
        
        if not excel_file.exists():
            raise FileNotFoundError(f"Excel file not found: {excel_file}")
        
        # Read the Excel file
        logger.info(f"Reading contract numbers from: {excel_file}")
        
        # Try to read as .xlsb first, if that fails, try .xlsx
        try:
            df = pd.read_excel(excel_file, sheet_name="IA Working")
        except Exception as e:
            logger.warning(f"Failed to read .xlsb file: {e}")
            # Try alternative file names
            alternative_files = [
                script_dir / "Prod. wise Class. of Loans - Jul 2025.xlsx",
                script_dir / "Prod. wise Class. of Loans - Jul 2025.xls"
            ]
            
            for alt_file in alternative_files:
                if alt_file.exists():
                    try:
                        df = pd.read_excel(alt_file, sheet_name="IA Working")
                        logger.info(f"Successfully read from alternative file: {alt_file}")
                        break
                    except Exception as alt_e:
                        logger.warning(f"Failed to read {alt_file}: {alt_e}")
                        continue
            else:
                raise Exception("Could not read any Excel file")
        
        # Extract contract numbers from column A starting from row 3 (index 2)
        contract_numbers = []
        for idx, row in df.iterrows():
            if idx >= 2:  # Skip first two rows (0 and 1), start from row 3 (index 2)
                contract = row.iloc[0]  # Column A
                if contract and pd.notna(contract) and str(contract).strip():
                    contract_str = str(contract).strip()
                    if contract_str != "nan" and contract_str != "":
                        contract_numbers.append(contract_str)
        
        logger.info(f"Successfully loaded {len(contract_numbers)} contract numbers from Excel file")
        return contract_numbers
        
    except Exception as e:
        logger.error(f"Failed to load contract numbers from Excel: {e}")
        raise

# ---------------- Logging Setup ----------------
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(LOGS_DIR, 'ia_working_valuation_bot.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ---------------- Credential Decryption ----------------
def load_credentials():
    """Load and decrypt credentials from .env file"""
    try:
        # Get encrypted credentials
        enc_username = os.getenv('ENC_USERNAME')
        enc_password = os.getenv('ENC_PASSWORD')
        
        if not enc_username or not enc_password:
            raise ValueError("ENC_USERNAME and ENC_PASSWORD must be set in environment variables")
        
        logger.info("Found encrypted credentials, attempting to decrypt...")
        
        # Get the root directory (parent of bots directory)
        root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        key_file = os.path.join(root_dir, ".key")
        
        if not os.path.exists(key_file):
            raise ValueError("Encryption key file (.key) not found in parent directory")
        
        # Read the encryption key
        with open(key_file, 'r') as f:
            key_str = f.read().strip()
        
        key = key_str.encode()
        fernet = Fernet(key)
        
        # Decrypt credentials
        username = fernet.decrypt(enc_username.encode()).decode()
        password = fernet.decrypt(enc_password.encode()).decode()
        
        logger.info("Successfully decrypted credentials")
        return username, password
        
    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        raise

# ---------------- Selenium Functions ----------------
def start_driver():
    """Start Chrome driver with optimized settings"""
    try:
        options = Options()
        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-plugins")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-images")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        # Check if Chrome is installed
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(os.getenv('USERNAME', '')),
            r"C:\Users\{}\AppData\Local\Google\Chrome\Application\chrome.exe".format(os.getenv('USERPROFILE', '').split('\\')[-1])
        ]
        
        chrome_found = False
        for path in chrome_paths:
            if os.path.exists(path):
                options.binary_location = path
                chrome_found = True
                logger.info(f"Found Chrome at: {path}")
                break
        
        if not chrome_found:
            raise Exception("Chrome browser not found. Please install Google Chrome.")
        
        # Start driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        logger.info("Chrome driver started successfully")
        return driver
        
    except Exception as e:
        logger.error(f"Failed to start Chrome driver: {e}")
        raise

def login_to_system(driver):
    """Login to the ERP system"""
    try:
        logger.info("Logging into ERP system...")
        driver.get(BASE_URL)
        time.sleep(3)
        
        # Load credentials
        username, password = load_credentials()
        
        # Find and fill username
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "strUserName"))
        )
        username_input.clear()
        username_input.send_keys(username)
        logger.info("Username entered")
        
        # Find and fill password
        password_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "strPassword"))
        )
        password_input.clear()
        password_input.send_keys(password)
        logger.info("Password entered")
        
        # Click sign in button
        signin_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "idSignIn"))
        )
        signin_button.click()
        logger.info("Sign in button clicked")
        
        # Wait for page to load and click SELECT button
        time.sleep(3)
        try:
            select_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//button[@id='idSignIn' and @type='submit']"))
            )
            select_button.click()
            logger.info("SELECT button clicked")
        except:
            logger.info("No additional SELECT button found, proceeding with login")
        
        time.sleep(3)
        logger.info("Login successful")
        return True
        
    except Exception as e:
        logger.error(f"Login failed: {e}")
        raise

def navigate_to_contract_inquiry(driver):
    """Navigate to Contract Inquiry page"""
    try:
        logger.info("Navigating to Contract Inquiry...")
        driver.get(RECOVERY_URL)
        time.sleep(3)
        
        # Click on Information Center tile
        info_center_tile = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, 
                "//div[contains(@class, 'metro-tile-body') and contains(@class, 'tile-type-menu')]//i[contains(@class, 'fa-file')]/.."))
        )
        info_center_tile.click()
        logger.info("Information Center tile clicked")
        time.sleep(3)
        
        # Click on Contract Inquiry tile
        contract_inquiry_tile = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH,
                "//div[contains(@class, 'metro-tile-body')]//i[contains(@class, 'fa-user')]/.."))
        )
        contract_inquiry_tile.click()
        logger.info("Contract Inquiry tile clicked")
        time.sleep(3)
        
        # Click on Add button
        add_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "btn_02461_Add"))
        )
        add_button.click()
        logger.info("Add button clicked")
        time.sleep(3)
        
        logger.info("Successfully navigated to Contract Inquiry")
        return True
        
    except Exception as e:
        logger.error(f"Failed to navigate to Contract Inquiry: {e}")
        raise

def process_contract(driver, contract_no):
    """Process a single contract to extract valuation amount using direct API call"""
    try:
        logger.info(f"Processing contract: {contract_no}")
        
        # Step 1: Find the contract number input field and enter contract number
        contract_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "strContractNo_RECContractInquiry"))
        )
        
        # Clear the field and enter contract number
        contract_input.clear()
        contract_input.send_keys(contract_no)
        logger.info(f"Entered contract number: {contract_no}")
        
        # Step 2: Hit Enter to trigger the search
        contract_input.send_keys(Keys.RETURN)
        logger.info("Pressed Enter to trigger search")
        
        # Step 3: Wait for page to load
        time.sleep(3)
        logger.info("Waiting for page to load...")
        
        # Step 4: Make direct API call to get Equipment Details
        try:
            # Get verification token from the page
            verification_token = ""
            try:
                token_element = driver.find_element(By.NAME, "__RequestVerificationToken")
                verification_token = token_element.get_attribute("value")
                logger.info("Found verification token")
            except:
                logger.warning("Could not find verification token")
            
            # Get cookies from the current session
            cookies = driver.get_cookies()
            logger.info(f"Retrieved {len(cookies)} cookies")
            
            # Create a requests session with the cookies
            import requests
            session = requests.Session()
            for cookie in cookies:
                session.cookies.set(cookie['name'], cookie['value'], domain=cookie['domain'])
            
            # Set headers matching the actual browser request
            headers = {
                "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
                "Accept": "*/*",
                "Accept-Language": "en-US,en;q=0.9",
                "Origin": "https://erp.assetline.lk",
                "Referer": "https://erp.assetline.lk/Application/Home/RECOVERY",
                "X-Requested-With": "XMLHttpRequest",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36 Edg/139.0.0.0",
                "Priority": "u=1, i",
                "Sec-Ch-Ua": '"Not;A=Brand";v="99", "Microsoft Edge";v="139", "Chromium";v="139"',
                "Sec-Ch-Ua-Mobile": "?0",
                "Sec-Ch-Ua-Platform": '"Windows"',
                "Sec-Fetch-Dest": "empty",
                "Sec-Fetch-Mode": "cors",
                "Sec-Fetch-Site": "same-origin"
            }
            
            # Prepare payload for Equipment Details API call
            payload = {
                "strContractNo_RECContractInquiry": contract_no,
                "strInquireType_RECContractInquiry": "MAININQ_EQUIPMENTDETAILS",
                "STR_FORM_ID": "02461",
                "STR_FUNCTION_ID": "CR",
                "STR_PREMIS": "HOD",
                "STR_INSTANT": "ALCL",
                "STR_APP_ID": "00023"
            }
            
            if verification_token:
                payload["__RequestVerificationToken"] = verification_token
            
            # Make the API request
            api_url = "https://erp.assetline.lk/Recovery/RECContractInquiry/Inquire"
            response = session.post(api_url, data=payload, headers=headers, timeout=30)
            
            logger.info(f"API call response status: {response.status_code}")
            
            if response.status_code == 200:
                try:
                    equipment_api_response = response.json()
                    logger.info("Successfully received Equipment Details API response")
                    
                    # Extract EQD_VALAMOUNT from the API response
                    valuation_amount = None
                    
                    if isinstance(equipment_api_response, dict):
                        logger.info(f"API Response structure: {list(equipment_api_response.keys())}")
                        
                        # Check if we have DATA field with dsResult
                        if "DATA" in equipment_api_response and "dsResult" in equipment_api_response["DATA"]:
                            ds_result = equipment_api_response["DATA"]["dsResult"]
                            logger.info(f"dsResult structure: {list(ds_result.keys())}")
                            
                            # Look for ContractInquire_Four which contains equipment details
                            if "ContractInquire_Four" in ds_result:
                                contract_four = ds_result["ContractInquire_Four"]
                                if isinstance(contract_four, list) and len(contract_four) > 0:
                                    # Get the first equipment item
                                    equipment_item = contract_four[0]
                                    if isinstance(equipment_item, dict) and "EQD_VALAMOUNT" in equipment_item:
                                        valuation_amount = str(equipment_item["EQD_VALAMOUNT"]).strip()
                                        logger.info(f"Found EQD_VALAMOUNT in ContractInquire_Four: {valuation_amount}")
                                    else:
                                        logger.warning("ContractInquire_Four item doesn't contain EQD_VALAMOUNT")
                                        logger.info(f"Available fields: {list(equipment_item.keys()) if isinstance(equipment_item, dict) else 'Not a dict'}")
                                else:
                                    logger.warning("ContractInquire_Four is empty or not a list")
                            else:
                                logger.warning("ContractInquire_Four not found in dsResult")
                                # Fallback: search through all dsResult items
                                for key, value in ds_result.items():
                                    if isinstance(value, list) and len(value) > 0:
                                        for item in value:
                                            if isinstance(item, dict) and "EQD_VALAMOUNT" in item:
                                                valuation_amount = str(item["EQD_VALAMOUNT"]).strip()
                                                logger.info(f"Found EQD_VALAMOUNT in {key}: {valuation_amount}")
                                                break
                                            if valuation_amount:
                                                break
                                    if valuation_amount:
                                        break
                        else:
                            # Fallback: search through the entire response
                            logger.info("No DATA/dsResult structure found, searching entire response...")
                            for key, value in equipment_api_response.items():
                                if isinstance(value, list) and len(value) > 0:
                                    for item in value:
                                        if isinstance(item, dict) and "EQD_VALAMOUNT" in item:
                                            valuation_amount = str(item["EQD_VALAMOUNT"]).strip()
                                            logger.info(f"Found EQD_VALAMOUNT in {key}: {valuation_amount}")
                                            break
                                        if valuation_amount:
                                            break
                                elif isinstance(value, dict) and "EQD_VALAMOUNT" in value:
                                    valuation_amount = str(value["EQD_VALAMOUNT"]).strip()
                                    logger.info(f"Found EQD_VALAMOUNT directly in {key}: {valuation_amount}")
                                    break
                                if valuation_amount:
                                    break
                    
                    if valuation_amount:
                        logger.info(f"✓ Successfully extracted valuation amount: {valuation_amount}")
                        return {
                            "contract_number": contract_no,
                            "valuation_amount": valuation_amount,
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                    else:
                        logger.warning("No EQD_VALAMOUNT field found in API response")
                        logger.info(f"API Response keys: {list(equipment_api_response.keys()) if isinstance(equipment_api_response, dict) else 'Not a dict'}")
                        logger.info(f"API Response structure: {json.dumps(equipment_api_response, indent=2)[:500]}...")
                        return {
                            "contract_number": contract_no,
                            "valuation_amount": "NO_EQD_VALAMOUNT_FOUND",
                            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        }
                        
                except json.JSONDecodeError as e:
                    logger.error(f"Failed to parse JSON response: {e}")
                    return {
                        "contract_number": contract_no,
                        "valuation_amount": f"JSON_PARSE_ERROR: {str(e)}",
                        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    }
            else:
                logger.error(f"API call failed with status: {response.status_code}")
                return {
                    "contract_number": contract_no,
                    "valuation_amount": f"API_CALL_FAILED: {response.status_code}",
                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                
        except Exception as api_error:
            logger.error(f"Error making API call: {api_error}")
            return {
                "contract_number": contract_no,
                "valuation_amount": f"API_CALL_ERROR: {str(api_error)}",
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
    except Exception as e:
        logger.error(f"Error processing contract {contract_no}: {e}")
        return {
            "contract_number": contract_no,
            "valuation_amount": f"ERROR: {str(e)}",
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

def display_results(contracts_data):
    """Display results in terminal"""
    try:
        if contracts_data:
            print("\n" + "="*80)
            print("VALUATION AMOUNT EXTRACTION SUMMARY")
            print("="*80)
            print(f"Total Contracts Processed: {len(contracts_data)}")
            print(f"Generated At: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("="*80)
            
            # Display each contract's valuation data
            print("\nEXTRACTED VALUATION DATA:")
            print("-" * 80)
            for contract_data in contracts_data:
                if contract_data:
                    contract_no = contract_data.get('contract_number', 'N/A')
                    valuation_amount = contract_data.get('valuation_amount', 'N/A')
                    print(f"\n{contract_no}:")
                    print(f"  Valuation Amount: {valuation_amount}")
                else:
                    print(f"\nNo data extracted")
            
            print("\n" + "="*80)
            return True
        else:
            print("\nNo contract data to display")
            return False
            
    except Exception as e:
        logger.error(f"Failed to display results: {e}")
        return False

def run_valuation_bot(contract_numbers):
    """
    Function to be called from NBD_MF_23_IA.py
    Returns a dictionary with contract numbers as keys and valuation amounts as values
    """
    try:
        logger.info(f"Running valuation bot for {len(contract_numbers)} contracts...")
        
        # Run the main function with provided contract numbers
        results = main(contract_numbers)
        
        if results:
            # Convert results to the expected format: {contract_number: valuation_amount}
            valuation_dict = {}
            for contract_data in results:
                if contract_data and contract_data.get('contract_number'):
                    contract_no = contract_data.get('contract_number')
                    valuation_amount = contract_data.get('valuation_amount', 'N/A')
                    valuation_dict[contract_no] = valuation_amount
            
            logger.info(f"Valuation bot completed. Processed {len(valuation_dict)} contracts.")
            return valuation_dict
        else:
            logger.warning("Valuation bot returned no results.")
            return {}
            
    except Exception as e:
        logger.error(f"Valuation bot execution failed: {e}")
        raise

# ---------------- Main flow ----------------
def main(contract_numbers=None):
    """Main execution function"""
    driver = None
    try:
        logger.info("Starting IA Working Initial Valuation Bot...")
        
        # Use provided contract numbers or load from Excel file
        if contract_numbers is None:
            contract_numbers = load_contract_numbers_from_excel()
        
        if not contract_numbers:
            logger.error("No contract numbers provided or loaded. Exiting.")
            return
        
        logger.info(f"Processing {len(contract_numbers)} contract numbers")
        
        # Start Chrome driver
        driver = start_driver()
        
        # Login to system
        if not login_to_system(driver):
            logger.error("Login failed. Exiting.")
            return
        
        # Navigate to Contract Inquiry
        if not navigate_to_contract_inquiry(driver):
            logger.error("Navigation to Contract Inquiry failed. Exiting.")
            return
        
        # Process each contract number
        contracts_data = []
        total_contracts = len(contract_numbers)
        
        for i, contract_no in enumerate(contract_numbers, 1):
            logger.info(f"Processing contract {i}/{total_contracts}: {contract_no}")
            
            # Process the contract
            result = process_contract(driver, contract_no)
            contracts_data.append(result)
            
            # Small delay between contracts
            time.sleep(2)
            
            # Clear the input field for the next contract (except for the last one)
            if i < total_contracts:
                try:
                    contract_input = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.ID, "strContractNo_RECContractInquiry"))
                    )
                    contract_input.clear()
                    logger.info("Cleared contract input field for next contract")
                except Exception as e:
                    logger.warning(f"Failed to clear contract input field: {e}")
        
        # Display results in terminal
        display_results(contracts_data)
        
        # Final summary
        successful_contracts = sum(1 for data in contracts_data if data and data.get("valuation_amount") and not data.get("valuation_amount").startswith(("NO_", "API_", "ERROR")))
        logger.info("="*80)
        logger.info("FINAL PROCESSING SUMMARY")
        logger.info("="*80)
        logger.info(f"Total Contracts: {total_contracts}")
        logger.info(f"Successfully Processed: {successful_contracts}")
        logger.info(f"Failed: {total_contracts - successful_contracts}")
        logger.info(f"Success Rate: {(successful_contracts/total_contracts)*100:.1f}%")
        logger.info("="*80)
        
        logger.info(f"Bot completed! Successfully processed {successful_contracts}/{total_contracts} contracts")
        
        return contracts_data
        
    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        raise
    finally:
        if driver:
            driver.quit()
            logger.info("Chrome driver closed")

if __name__ == "__main__":
    # When run directly, load contract numbers from Excel file
    try:
        main()
    except Exception as e:
        logger.error(f"Bot execution failed: {e}")
        exit(1)
