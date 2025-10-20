import os
import glob
import pandas as pd
from datetime import datetime, date
import numpy as np
import subprocess
import json
import logging
import sys
import traceback
from typing import Dict, Optional, Tuple, Any

# Create logs folder if it doesn't exist
logs_dir = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(logs_dir, exist_ok=True)

# Create timestamp-based log file name
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
log_file = os.path.join(logs_dir, f"NBD_MF23_IA_Report_{timestamp}.log")

# Enhanced logging configuration
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s",
    handlers=[
        logging.FileHandler(log_file, mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)
logger.info("="*80)
logger.info("STARTING NEW SESSION - NBD_MF23_IA_Report")
logger.info("Logging initialized. Logs are being written to %s", log_file)
logger.info("="*80)

# Add the parent "bots" folder relative to this file
current_dir = os.path.dirname(__file__)
bot_dir = os.path.join(current_dir, "..", "bots")
sys.path.insert(0, os.path.abspath(bot_dir))

# Import bot modules with error handling
try:
    import na_contract_numbers_search_bot_api as na_contract_bot
    logger.info("Successfully imported na_contract_numbers_search_bot_api module")
except ImportError as e:
    logger.warning(f"Could not import na_contract_numbers_search_bot_api: {e}")
    na_contract_bot = None

try:
    import IA_Working_Initial_valuation_bot as initial_valuation_bot
    logger.info("Successfully imported IA_Working_Initial_valuation_bot module")
except ImportError as e:
    logger.warning(f"Could not import IA_Working_Initial_valuation_bot: {e}")
    initial_valuation_bot = None


class NBD_MF23_IA_Report:
    def __init__(self, base_dir=r"working", working_dir=None):
        """Initialize the report generator with error handling.

        Args:
            base_dir: Base directory for default folder searching (used in standalone mode)
            working_dir: Direct path to working directory (used when called from app.py)
        """
        try:
            self.base_dir = base_dir
            self.working_dir = working_dir
            self.ia_folder = self._find_ia_folder()
            logger.info(f"Initialized NBD_MF23_IA_Report with base_dir: {base_dir}")
            if working_dir:
                logger.info(f"Using working_dir from app.py: {working_dir}")
            logger.info(f"IA folder found at: {self.ia_folder}")
        except Exception as e:
            logger.error(f"Failed to initialize NBD_MF23_IA_Report: {e}")
            raise

    @staticmethod
    def parse_date(date_input):
        """Parse date from various formats - for use by app.py or command line."""
        if isinstance(date_input, date):
            return date_input
        elif isinstance(date_input, str):
            try:
                # Try MM/DD/YYYY format first (command line)
                return datetime.strptime(date_input, "%m/%d/%Y").date()
            except ValueError:
                try:
                    # Try YYYY-MM-DD format (app.py)
                    return datetime.strptime(date_input, "%Y-%m-%d").date()
                except ValueError:
                    raise ValueError(f"Invalid date format: {date_input}. Expected MM/DD/YYYY or YYYY-MM-DD")
        else:
            raise ValueError(f"Invalid date type: {type(date_input)}. Expected string or date object")

    def _find_ia_folder(self):
        """Find the latest NBD_MF_23_IA date folder inside the working directory."""
        try:
            # Step 1: Locate project root (one level above report_automations)
            root_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

            # Step 2: Path to working/NBD_MF_23_IA
            ia_root = os.path.join(root_dir, "working", "NBD_MF_23_IA")
            if not os.path.exists(ia_root):
                error_msg = f"'NBD_MF_23_IA' folder not found at '{ia_root}'."
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)

            # Step 3: Get all subfolders (date-based)
            date_folders = [f for f in glob.glob(os.path.join(ia_root, "*")) if os.path.isdir(f)]
            if not date_folders:
                error_msg = f"No subfolders found under '{ia_root}'."
                logger.error(error_msg)
                raise FileNotFoundError(error_msg)

            # Step 4: Pick latest modified folder
            date_folders.sort(key=os.path.getmtime, reverse=True)
            latest_folder = date_folders[0]

            logger.info(f"Found latest IA working folder: {latest_folder}")
            return latest_folder

        except Exception as e:
            logger.error(f"Error finding IA folder: {e}")
            raise


    def _find_file(self, keyword: str) -> str:
        """Find a file containing the keyword in the IA folder with detailed logging."""
        try:
            logger.debug(f"Searching for file with keyword: '{keyword}' in {self.ia_folder}")

            files = os.listdir(self.ia_folder)
            # Filter out temporary Excel files that start with ~$
            matching_files = [f for f in files if keyword in f and not f.startswith('~$')]

            if not matching_files:
                error_msg = f"No file found with keyword: '{keyword}' in {self.ia_folder}"
                logger.error(error_msg)
                logger.error(f"Available files (excluding ~$ temp files): {[f for f in files if not f.startswith('~$')]}")
                raise FileNotFoundError(error_msg)

            if len(matching_files) > 1:
                logger.warning(f"Multiple files found with keyword '{keyword}': {matching_files}. Using first match.")

            selected_file = os.path.join(self.ia_folder, matching_files[0])
            logger.debug(f"Found file: {selected_file}")
            return selected_file

        except Exception as e:
            logger.error(f"Error finding file with keyword '{keyword}': {e}")
            raise

    # === File Loaders with Enhanced Error Handling ===
    def load_disbursement(self):
        """Load disbursement data with error handling."""
        try:
            logger.info("Loading Disbursement with Budget file...")
            file_disbursement = self._find_file("Disbursement with Budget")
            
            df = pd.read_excel(
                file_disbursement,
                sheet_name="month",
                usecols="A,H,AC"
            )
            
            # Validate required columns
            required_cols = ["CONTRACT NO", "Net Amount", "Base Rate"]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Missing required columns in Disbursement file: {missing_cols}")
            
            df_filtered = df[df["CONTRACT NO"].notna()]
            logger.info(f"Successfully loaded Disbursement data: {len(df_filtered)} rows")
            return df_filtered
            
        except Exception as e:
            logger.error(f"Failed to load Disbursement data: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_information_request_from_credit(self):
        """Load information request from credit with error handling."""
        try:
            logger.info("Loading Information Request from Credit file...")
            file_path = self._find_file("Information Request from Credit")
            
            df = pd.read_excel(
                file_path,
                sheet_name="Disbursements",
                usecols="C,G",
                skiprows=4
            )
            
            # Rename columns if needed
            if len(df.columns) == 2:
                df.columns = ["Contract No", "Micro/Small/Medium"]
            
            df_filtered = df[df["Contract No"].notna()]
            logger.info(f"Successfully loaded Information Request data: {len(df_filtered)} rows")
            return df_filtered
            
        except Exception as e:
            logger.error(f"Failed to load Information Request from Credit: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_net_portfolio(self):
        """Load net portfolio with error handling."""
        try:
            logger.info("Loading Net Portfolio file...")
            file_path = self._find_file("Net Portfolio")
            
            df = pd.read_excel(
                file_path,
                usecols="E,C,F,H,AB,AH,AI,AM"
            )
            
            # Validate required columns
            required_cols = ["CLIENT_CODE", "CONTRACT_NO"]
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                logger.warning(f"Missing expected columns in Net Portfolio: {missing_cols}")
            
            df_filtered = df[df["CLIENT_CODE"].notna()]
            logger.info(f"Successfully loaded Net Portfolio data: {len(df_filtered)} rows")
            return df_filtered
            
        except Exception as e:
            logger.error(f"Failed to load Net Portfolio: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_yard_and_Property_List(self):
        """Load yard and property list with comprehensive error handling."""
        try:
            logger.info("Loading CBSL Provision Comparison file...")
            file_path = self._find_file("CBSL Provision Comparison")
            
            # Load Property Mortgage
            try:
                df_propertyMortgage = pd.read_excel(
                    file_path,
                    sheet_name="PropertyMortgage",
                    usecols="A",
                    skiprows=5,
                    header=None
                )
                df_propertyMortgage.columns = ["Property Mortgage List"]
                logger.info(f"Loaded PropertyMortgage: {len(df_propertyMortgage)} rows")
            except Exception as e:
                logger.warning(f"Could not load PropertyMortgage sheet: {e}. Creating empty dataframe.")
                df_propertyMortgage = pd.DataFrame(columns=["Property Mortgage List"])

            # Load Repossessed List
            try:
                df_repossessedList = pd.read_excel(
                    file_path,
                    sheet_name="RepossessedList",
                    usecols="A",
                    skiprows=5,
                    header=None
                )
                df_repossessedList.columns = ["Yard Contract"]
                logger.info(f"Loaded RepossessedList: {len(df_repossessedList)} rows")
            except Exception as e:
                logger.warning(f"Could not load RepossessedList sheet: {e}. Creating empty dataframe.")
                df_repossessedList = pd.DataFrame(columns=["Yard Contract"])

            # Build combined dataframe
            df_yard_and_property_list = pd.DataFrame()
            df_yard_and_property_list["Yard Contract"] = df_repossessedList["Yard Contract"]
            df_yard_and_property_list[""] = [""] + ["Yard"] * (max(0, len(df_repossessedList) - 1))
            df_yard_and_property_list["Empty"] = ""
            
            max_len = max(len(df_repossessedList), len(df_propertyMortgage))
            df_yard_and_property_list = df_yard_and_property_list.reindex(range(max_len))
            df_yard_and_property_list["Property Mortgage List"] = df_propertyMortgage.reindex(range(max_len))["Property Mortgage List"].fillna("")
            
            logger.info(f"Successfully created yard and property list: {len(df_yard_and_property_list)} rows")
            return df_yard_and_property_list
            
        except Exception as e:
            logger.error(f"Failed to load Yard and Property List: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_product_cat(self):
        """Load product category with error handling."""
        try:
            logger.info("Loading Prod. wise Class. of Loans file...")
            file_path = self._find_file("Prod. wise Class. of Loans")
            
            df = pd.read_excel(
                file_path,
                sheet_name="Product_Cat",
                usecols="E,F,L,M,P,R,S"
            )
            
            # Rename columns
            df = df.rename(columns={
                "SPECIAL CATEGORIES (CONTRACT WISE) *Add here to pick contract wise cat": "SPECIAL CATEGORIES",
                df.columns[-1]: "M"
            })
            
            logger.info(f"Successfully loaded Product_Cat data: {len(df)} rows, columns: {df.columns.tolist()}")
            return df
            
        except Exception as e:
            logger.error(f"Failed to load Product_Cat: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_C1_C2_Working(self):
        """Load C1 & C2 Working with error handling."""
        try:
            logger.info("Loading C1 & C2 Working sheet...")
            file_path = self._find_file("Prod. wise Class. of Loans")
            
            df = pd.read_excel(
                file_path,
                sheet_name="C1 & C2 Working",
                header=1,
                usecols="A, AB, AD"
            )
            
            # Rename columns
            df = df.rename(columns={
                df.columns[0]: "Contract No",
                df.columns[1]: "Gross Exposure",
                df.columns[2]: "PD Category" if len(df.columns) > 2 else None
            })
            
            # Keep all contract rows (removed filtering of LR/Margin Trading to ensure proper mapping)
            logger.info(f"Keeping all contracts including LR and Margin Trading for proper PD Category and Gross Exposure mapping")
            
            logger.info(f"Successfully loaded C1 & C2 Working data: {len(df)} rows")
            return df
            
        except Exception as e:
            logger.error(f"Failed to load C1 & C2 Working: {e}")
            logger.error(traceback.format_exc())
            raise

    def load_po_listing(self):
        """Load PO listing with error handling."""
        try:
            logger.info("Loading Po Listing - Internal file...")
            file_path = self._find_file("Po Listing - Internal")
            df = pd.read_excel(file_path, usecols="H, AH")
            # Assign proper column names: H = Contract No, AH = SELL_PRICE
            df.columns = ["CONTRACT_NO", "SELL_PRICE"]
            df_filtered = df[df["CONTRACT_NO"].notna() & df["SELL_PRICE"].notna()]
            logger.info(f"Successfully loaded PO Listing data: {len(df_filtered)} rows with valid Contract No and SELL_PRICE")
            return df_filtered
        except Exception as e:
            logger.error(f"Failed to load PO Listing: {e}")
            logger.warning("Returning empty DataFrame for PO Listing")
            return pd.DataFrame(columns=["CONTRACT_NO", "SELL_PRICE"])

    def load_portfolio_report_recovery(self):
        """Load portfolio report recovery with error handling."""
        try:
            logger.info("Loading Portfolio Report Recovery - Internal file...")
            file_path = self._find_file("Portfolio Report Recovery - Internal")
            df = pd.read_excel(file_path, usecols="A,B")
            df.columns = ["CONTRACT_NO", "SELL_PRICE"]
            df_filtered = df[df["CONTRACT_NO"].notna()]
            logger.info(f"Successfully loaded Portfolio Report Recovery: {len(df_filtered)} rows")
            return df_filtered
        except FileNotFoundError:
            logger.warning("Portfolio Report Recovery file not found. Returning empty DataFrame.")
            return pd.DataFrame(columns=["CONTRACT_NO", "SELL_PRICE"])
        except Exception as e:
            logger.error(f"Failed to load Portfolio Report Recovery: {e}")
            return pd.DataFrame(columns=["CONTRACT_NO", "SELL_PRICE"])

    def safe_apply(self, df: pd.DataFrame, func, axis=1, default_value=None) -> pd.Series:
        """Safely apply a function to a dataframe with error handling."""
        try:
            return df.apply(func, axis=axis)
        except Exception as e:
            logger.error(f"Error in apply function: {e}")
            logger.error(f"Function: {func.__name__ if hasattr(func, '__name__') else 'lambda'}")
            logger.error(traceback.format_exc())
            if default_value is not None:
                return pd.Series([default_value] * len(df), index=df.index)
            raise

    def classify_collateral_excel_formula(self, row, property_mortgage_list, vehicles_machinery_list):
        """
        Classify collateral following Excel formula EXACTLY:
        =IF(ISNUMBER(MATCH(A3,'YardandProperty List'!D:D,0)),"Immovable Properties",
            IF(ISNUMBER(MATCH(T3,Product_Cat!$P$2:$P$9,0)),"Vehicles and Machinery",
                IF(B3="Margin Trading","Shares and Debt Securities-Listed",
                    IF(B3="FDL","Deposits (Cash-Backed)","Personal and Corporate Guarantees"))))

        Where: A3=Contract No, T3=PD Category, B3=Product
        """
        try:
            # Extract values exactly as in Excel formula
            contract_no = str(row["Contract No"]).strip() if pd.notna(row["Contract No"]) else ""  # A3
            pd_category = str(row["PD Category"]).strip() if pd.notna(row["PD Category"]) else ""  # T3
            product = str(row["Product"]).strip() if pd.notna(row["Product"]) else ""  # B3

            # Log the values being checked for this row
            logger.debug(f"Classifying: Contract={contract_no}, PD_Category={pd_category}, Product={product}")

            # Step 1: IF(ISNUMBER(MATCH(A3,'YardandProperty List'!D:D,0)),"Immovable Properties"
            if contract_no and contract_no in property_mortgage_list:
                logger.debug(f"  → Immovable Properties (Contract {contract_no} in property list)")
                return "Immovable Properties"

            # Step 2: IF(ISNUMBER(MATCH(T3,Product_Cat!$P$2:$P$9,0)),"Vehicles and Machinery"
            if pd_category and pd_category in vehicles_machinery_list:
                logger.debug(f"  → Vehicles and Machinery (PD Category '{pd_category}' in vehicles list)")
                return "Vehicles and Machinery"

            # Step 3: IF(B3="Margin Trading","Shares and Debt Securities-Listed"
            if product == "Margin Trading":
                logger.debug(f"  → Shares and Debt Securities-Listed (Product = 'Margin Trading')")
                return "Shares and Debt Securities-Listed"

            # Step 4: IF(B3="FDL","Deposits (Cash-Backed)"
            if product == "FDL":
                logger.debug(f"  → Deposits (Cash-Backed) (Product = 'FDL')")
                return "Deposits (Cash-Backed)"

            # Step 5: Default case
            logger.debug(f"  → Personal and Corporate Guarantees (default case)")
            return "Personal and Corporate Guarantees"

        except Exception as e:
            logger.error(f"Error classifying collateral for row: {e}")
            logger.error(f"Row data: Contract={row.get('Contract No')}, PD={row.get('PD Category')}, Product={row.get('Product')}")
            return "Personal and Corporate Guarantees"

    def build_prod_loans(self, df_disbursement, df_net_portfolio, df_information_request_from_credit,
                        df_product_cat, df_C1_C2_Working, df_yard_and_property_list,
                        df_po_listing, df_portfolio_report_recovery):
        """Build production loans dataframe with comprehensive error handling."""
        
        logger.info("="*80)
        logger.info("STARTING build_prod_loans METHOD")
        logger.info("="*80)
        
        try:
            # Initialize dataframe
            df_prod_loans = pd.DataFrame()
            df_prod_loans["Contract No"] = df_disbursement["CONTRACT NO"]

            logger.info(f"INITIAL CONTRACT COUNT: {len(df_prod_loans)} contracts from disbursement")
            
            # Process each column with error handling
            try:
                df_prod_loans["Disbursed Amount (Net of DC)"] = (
                    pd.to_numeric(df_disbursement["Net Amount"], errors="coerce")
                    .round(0)
                    .astype("Int64")
                )
                logger.info("✓ Disbursed Amount column created")
            except Exception as e:
                logger.error(f"Error creating Disbursed Amount column: {e}")
                df_prod_loans["Disbursed Amount (Net of DC)"] = None
            
            df_prod_loans["Base Rate"] = df_disbursement["Base Rate"]
            
            # Product column with replacement
            try:
                # Extract product codes and show samples for debugging
                df_prod_loans["Product"] = df_prod_loans["Contract No"].astype(str).str[2:4]

                # Show sample contract numbers and their extracted codes
                sample_contracts = df_prod_loans[["Contract No", "Product"]].head(10)
                logger.info(f"Sample contract numbers and extracted product codes:\n{sample_contracts.to_string()}")

                # Show unique product codes before replacement
                unique_products_before = df_prod_loans["Product"].unique()
                logger.info(f"Unique product codes before replacement: {unique_products_before}")

                df_prod_loans["Product"] = df_prod_loans["Product"].replace({
                    "00": "FDL",
                    "rg": "Margin Trading"
                })

                # Debug product classifications
                product_counts = df_prod_loans["Product"].value_counts()
                logger.info("✓ Product column created with replacements")
                logger.info(f"Product distribution: {product_counts.to_dict()}")

                fdl_count = len(df_prod_loans[df_prod_loans["Product"] == "FDL"])
                margin_trading_count = len(df_prod_loans[df_prod_loans["Product"] == "Margin Trading"])
                logger.info(f"FDL contracts: {fdl_count}, Margin Trading contracts: {margin_trading_count}")
            except Exception as e:
                logger.error(f"Error creating Product column: {e}")
                df_prod_loans["Product"] = ""
            
            # Build mapping dictionaries
            try:
                mapping_dicts = {
                    "Client Code": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CLIENT_CODE"])),
                    "Contract Period": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CONTRACT_PERIOD"])),
                    "Contract Amount": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CONTRACT_AMOUNT"])),
                    "Equipment": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["EQT_DESC"])),
                    "Frequency": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CON_RNTFREQ"])),
                    "Contractual Interest Rate": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["CON_INTRATE"])),
                    "Purpose": dict(zip(df_net_portfolio["CONTRACT_NO"], df_net_portfolio["PURPOSE"]))
                }
                
                for col, mapping in mapping_dicts.items():
                    df_prod_loans[col] = df_prod_loans["Contract No"].map(mapping).fillna("")
                    logger.info(f"✓ Mapped {col} column ({len([v for v in mapping.values() if v])} mappings)")

                logger.info("✓ All net portfolio mappings completed")
            except Exception as e:
                logger.error(f"Error in net portfolio mappings: {e}")
            
            # Clean Client Code
            df_prod_loans["Client Code"] = df_prod_loans["Client Code"].astype(str).str.replace(r"\.0$", "", regex=True)
            
            # Fill incomplete contracts
            df_prod_loans = self.fill_incomplete_contracts(df_prod_loans, na_contract_bot)
            
            # Corporate Clients classification
            try:
                df_prod_loans["Corporate Clients"] = df_prod_loans["Client Code"].str[0].apply(
                    lambda x: "Corporate Client" if x == "2" else "Non-Corporate"
                )
                logger.info("✓ Corporate Clients classification completed")
            except Exception as e:
                logger.error(f"Error in Corporate Clients classification: {e}")
                df_prod_loans["Corporate Clients"] = "Non-Corporate"
            
            # Calculate rates and tenure
            try:
                df_prod_loans["Contractual Interest Rate"] = pd.to_numeric(df_prod_loans["Contractual Interest Rate"], errors="coerce")
                df_prod_loans["Base Rate"] = pd.to_numeric(df_prod_loans["Base Rate"], errors="coerce")
                df_prod_loans["Minimum Rate (Final)"] = df_prod_loans[["Contractual Interest Rate", "Base Rate"]].min(axis=1)
                
                df_prod_loans["Contract Period"] = pd.to_numeric(df_prod_loans["Contract Period"], errors="coerce")
                conditions = [
                    df_prod_loans["Frequency"] == "D",
                    df_prod_loans["Frequency"] == "Q",
                    df_prod_loans["Frequency"] == "M",
                    df_prod_loans["Frequency"] == "W"
                ]
                choices = [
                    df_prod_loans["Contract Period"] / 30,
                    df_prod_loans["Contract Period"] * 3,
                    df_prod_loans["Contract Period"],
                    df_prod_loans["Contract Period"] / 4
                ]
                df_prod_loans["Tenure (Months)"] = np.select(conditions, choices, default=0)
                logger.info("✓ Rate and tenure calculations completed")
            except Exception as e:
                logger.error(f"Error in rate/tenure calculations: {e}")
            
            # MSME Classification
            try:
                contract_to_msme = dict(
                    zip(df_information_request_from_credit["Contract No"], 
                        df_information_request_from_credit["Micro/Small/Medium"])
                )
                df_prod_loans["MSME Classification"] = df_prod_loans["Contract No"].map(contract_to_msme).fillna("0")
                logger.info("✓ MSME Classification completed")
            except Exception as e:
                logger.error(f"Error in MSME Classification: {e}")
                df_prod_loans["MSME Classification"] = "0"
            
            # EIR Calculation with error handling
            try:
                df_prod_loans["Minimum Rate (Final)"] = pd.to_numeric(df_prod_loans["Minimum Rate (Final)"], errors="coerce")
                df_prod_loans["Tenure (Months)"] = pd.to_numeric(df_prod_loans["Tenure (Months)"], errors="coerce")
                
                # Avoid division by zero
                mask = df_prod_loans["Tenure (Months)"] > 0
                df_prod_loans.loc[mask, "EIR (%)"] = (
                    (1 + df_prod_loans.loc[mask, "Minimum Rate (Final)"] / 100 / 
                    (12 / df_prod_loans.loc[mask, "Tenure (Months)"])) ** 
                    (12 / df_prod_loans.loc[mask, "Tenure (Months)"]) - 1
                ) * 100

                # Convert to numeric & round
                df_prod_loans["EIR (%)"] = pd.to_numeric(df_prod_loans["EIR (%)"], errors="coerce").round(2)

                # 👉 Divide by 100 when saving (decimal form instead of percentage)
                df_prod_loans["EIR (Decimal)"] = df_prod_loans["EIR (%)"] / 10000

                logger.info("✓ EIR calculation completed")

            except Exception as e:
                logger.error(f"Error in EIR calculation: {e}")
                df_prod_loans["EIR (%)"] = ""
                df_prod_loans["EIR (Decimal)"] = ""

            
            # Type of Loans classification
            # Excel formula: =XLOOKUP(A3,Product_Cat!L:L,Product_Cat!M:M,
            #                 IF(OR(B3="MT",B3="Margin Trading"),"Margin Trading Loans",
            #                 IF(OR(B3="FDL",B3="FD Loan"),"Loans against Cash/Deposits",
            #                 XLOOKUP((B3&D3&E3&F3),Product_Cat!F:F,Product_Cat!E:E))))
            try:
                # Create lookup dictionaries from Product_Cat sheet
                # L:L = SPECIAL CATEGORIES (Contract No), M:M = result
                # Clean NaN values before creating dict
                special_cat_df = df_product_cat[["SPECIAL CATEGORIES", "M"]].dropna(subset=["SPECIAL CATEGORIES"])
                special_cat_map = dict(zip(special_cat_df["SPECIAL CATEGORIES"], special_cat_df["M"]))

                # F:F = LOOKUP (Product&Equipment&Purpose&Corporate), E:E = Classification
                lookup_df = df_product_cat[["LOOKUP", "Classification"]].dropna(subset=["LOOKUP"])
                lookup_map = dict(zip(lookup_df["LOOKUP"], lookup_df["Classification"]))

                logger.info(f"Special categories map has {len(special_cat_map)} entries")
                logger.info(f"Lookup map has {len(lookup_map)} entries")

                def classify_loan_type(row):
                    """Classify loan type matching Excel XLOOKUP formula."""
                    contract_no = str(row["Contract No"]) if pd.notna(row["Contract No"]) else ""
                    product = str(row["Product"]) if pd.notna(row["Product"]) else ""
                    equipment = str(row["Equipment"]) if pd.notna(row["Equipment"]) else ""
                    purpose = str(row["Purpose"]) if pd.notna(row["Purpose"]) else ""
                    corporate_clients = str(row["Corporate Clients"]) if pd.notna(row["Corporate Clients"]) else ""

                    # First XLOOKUP: Check special categories by Contract No (A3 in L:L column)
                    if contract_no and contract_no in special_cat_map:
                        return special_cat_map[contract_no]

                    # IF condition 1: Check for Margin Trading (B3="MT" or "Margin Trading")
                    if product in ["MT", "Margin Trading"]:
                        return "Margin Trading Loans"

                    # IF condition 2: Check for FD Loans (B3="FDL" or "FD Loan")
                    if product in ["FDL", "FD Loan"]:
                        return "Loans against Cash/Deposits"

                    # Second XLOOKUP: Use concatenated key (B3&D3&E3&F3) in F:F column
                    lookup_key = f"{product}{equipment}{purpose}{corporate_clients}"
                    result = lookup_map.get(lookup_key)

                    if result is not None and pd.notna(result):
                        return result

                    return None

                df_prod_loans["Type of Loans (1.3.1.0.0.0)"] = df_prod_loans.apply(classify_loan_type, axis=1)

                # Log statistics and sample results
                null_count = df_prod_loans["Type of Loans (1.3.1.0.0.0)"].isna().sum()
                logger.info(f"Type of Loans classification: {len(df_prod_loans) - null_count} classified, {null_count} null")

                # Show value counts for debugging
                value_counts = df_prod_loans["Type of Loans (1.3.1.0.0.0)"].value_counts()
                logger.info(f"Type of Loans value distribution:\n{value_counts}")

                logger.info("✓ Type of Loans classification completed")
            except Exception as e:
                logger.error(f"Error in Type of Loans classification: {e}")
                logger.error(traceback.format_exc())

            # PD Category mapping (MUST happen before Collateral classification)
            try:
                logger.info("Mapping PD Category (required for Collateral classification)...")
                contract_to_pd_category = dict(
                    zip(df_C1_C2_Working["Contract No"], df_C1_C2_Working["PD Category"])
                )
                equipment_to_category = dict(
                    zip(df_product_cat["EQT_DESC"], df_product_cat["Classification"])
                )
                df_prod_loans["PD Category"] = df_prod_loans["Contract No"].map(contract_to_pd_category)
                df_prod_loans["PD Category"] = df_prod_loans["PD Category"].fillna(
                    df_prod_loans["Equipment"].map(equipment_to_category)
                )
                logger.info(f"✓ PD Category mapping completed. Non-null values: {df_prod_loans['PD Category'].notna().sum()}")
            except Exception as e:
                logger.error(f"Error in PD Category mapping: {e}")
                df_prod_loans["PD Category"] = ""

            # Collateral/Security Type classification
            try:
                logger.info("="*60)
                logger.info("STARTING COLLATERAL/SECURITY TYPE CLASSIFICATION")
                logger.info("="*60)
                logger.info(f"Current df_prod_loans columns: {df_prod_loans.columns.tolist()}")
                logger.info(f"df_prod_loans shape: {df_prod_loans.shape}")

                # Get lists for classification
                property_mortgage_list = df_yard_and_property_list["Property Mortgage List"].dropna().astype(str).str.strip().tolist()
                logger.info(f"Property Mortgage List loaded: {len(property_mortgage_list)} items")
                if property_mortgage_list:
                    logger.info(f"Sample Property Mortgage contracts: {property_mortgage_list[:5]}")

                # Get Column P values (df_product_cat columns debugging)
                logger.info(f"df_product_cat columns: {df_product_cat.columns.tolist()}")
                logger.info(f"df_product_cat shape: {df_product_cat.shape}")

                if len(df_product_cat.columns) > 4:
                    p_column = df_product_cat.iloc[:, 4]
                    p_column_name = df_product_cat.columns[4]
                    logger.info(f"Using column at index 4: '{p_column_name}'")

                    # Show all values in P column
                    logger.info(f"All values in P column:\n{p_column.to_string()}")

                    vehicles_machinery_list = p_column.iloc[0:9].dropna().astype(str).str.strip().tolist()
                    vehicles_machinery_list = [x for x in vehicles_machinery_list if x and x != 'nan']
                    logger.info(f"Vehicles & Machinery List (P1:P9): {vehicles_machinery_list}")
                else:
                    vehicles_machinery_list = []
                    logger.warning("df_product_cat has insufficient columns for P column lookup")

                # Log exact matches that will be checked
                logger.info("EXACT MATCHES TO BE CHECKED:")
                logger.info(f"  Property Mortgage matches: Contract No must be IN {property_mortgage_list[:10]}...")
                logger.info(f"  Vehicles & Machinery matches: PD Category must be IN {vehicles_machinery_list}")
                logger.info(f"  Margin Trading matches: Product must EQUAL 'Margin Trading'")
                logger.info(f"  FDL matches: Product must EQUAL 'FDL'")

                # Check if PD Category column exists (required for classification)
                if "PD Category" not in df_prod_loans.columns:
                    logger.warning("PD Category column not found! Creating empty column...")
                    df_prod_loans["PD Category"] = ""

                # Show sample PD Category values for debugging
                logger.info("Sample PD Category values in df_prod_loans:")
                pd_sample = df_prod_loans["PD Category"].dropna().head(10).tolist()
                logger.info(f"PD Categories: {pd_sample}")
                logger.info(f"Total non-null PD Categories: {df_prod_loans['PD Category'].notna().sum()}")

                # Apply classification using EXACT Excel formula logic
                logger.info("Applying Collateral/Security Type classification (Excel formula)...")
                df_prod_loans["Collateral/Security Type"] = self.safe_apply(
                    df_prod_loans,
                    lambda row: self.classify_collateral_excel_formula(
                        row, property_mortgage_list, vehicles_machinery_list
                    ),
                    default_value="Personal and Corporate Guarantees"
                )

                # Log sample classifications for debugging
                logger.info("SAMPLE CLASSIFICATIONS (first 10 rows):")
                for idx in range(min(10, len(df_prod_loans))):
                    row = df_prod_loans.iloc[idx]
                    classification = row["Collateral/Security Type"]
                    logger.info(f"  Row {idx+1}: Contract={row['Contract No']}, PD={row.get('PD Category', 'N/A')}, Product={row.get('Product', 'N/A')} → {classification}")

                # Log classification results
                classification_counts = df_prod_loans["Collateral/Security Type"].value_counts()
                logger.info("="*40)
                logger.info("COLLATERAL CLASSIFICATION RESULTS:")
                logger.info("="*40)
                for ctype, count in classification_counts.items():
                    logger.info(f"  {ctype}: {count}")
                logger.info("="*40)

                # Special check for Vehicles and Machinery
                vehicles_count = classification_counts.get("Vehicles and Machinery", 0)
                logger.info(f"Found {vehicles_count} Vehicles and Machinery contracts (Expected: 6426)")

                if vehicles_count < 6426:
                    # Find PD Categories that are NOT in vehicles_machinery_list but should be
                    non_vm_contracts = df_prod_loans[df_prod_loans["Collateral/Security Type"] != "Vehicles and Machinery"]
                    unique_pd_categories = non_vm_contracts["PD Category"].dropna().unique()

                    logger.warning(f"MISSING {6426 - vehicles_count} Vehicles and Machinery contracts!")
                    logger.warning(f"PD Categories in non-Vehicles contracts: {list(unique_pd_categories)}")
                    logger.warning(f"Current vehicles_machinery_list: {vehicles_machinery_list}")

                    # Show specific contracts that might be missing
                    potential_missing = non_vm_contracts[non_vm_contracts["PD Category"].notna()]
                    if len(potential_missing) > 0:
                        logger.warning("Sample contracts with PD Categories that aren't classified as Vehicles and Machinery:")
                        for idx in range(min(5, len(potential_missing))):
                            row = potential_missing.iloc[idx]
                            logger.warning(f"  Contract={row['Contract No']}, PD={row.get('PD Category', 'N/A')}, Current Classification={row['Collateral/Security Type']}")
                else:
                    logger.info(f"✓ Found expected {vehicles_count} Vehicles and Machinery contracts")

            except Exception as e:
                logger.error(f"Error in Collateral/Security Type classification: {e}")
                df_prod_loans["Collateral/Security Type"] = "Personal and Corporate Guarantees"
            
            # Initial Valuation =====================================================
            try:
                logger.info("Starting Initial Valuation processing...")
                df_prod_loans["Initial Valuation"] = None
                mask = df_prod_loans["Collateral/Security Type"] == "Vehicles and Machinery"
                
                # Step 1: PO Listing (SELL_PRICE)
                if not df_po_listing.empty and "CONTRACT_NO" in df_po_listing.columns and "SELL_PRICE" in df_po_listing.columns:
                    contract_to_sell_price = dict(
                        zip(df_po_listing["CONTRACT_NO"].astype(str),
                            pd.to_numeric(df_po_listing["SELL_PRICE"], errors='coerce'))
                    )
                    # Remove any NaN or invalid values from the mapping
                    contract_to_sell_price = {k: v for k, v in contract_to_sell_price.items() if pd.notna(v) and v > 0}

                    df_prod_loans.loc[mask, "Initial Valuation"] = (
                        df_prod_loans.loc[mask, "Contract No"].astype(str).map(contract_to_sell_price)
                    )

                    po_applied = df_prod_loans.loc[mask, "Initial Valuation"].notna().sum()
                    logger.info(f"✓ Applied Initial Valuation from PO Listing (SELL_PRICE): {po_applied} contracts")
                else:
                    logger.warning("PO Listing data not available or missing required columns")
                
                # Step 2: Portfolio Report Recovery (SELL_PRICE)
                if not df_portfolio_report_recovery.empty and "CONTRACT_NO" in df_portfolio_report_recovery.columns and "SELL_PRICE" in df_portfolio_report_recovery.columns:
                    recovery_lookup = dict(
                        zip(df_portfolio_report_recovery["CONTRACT_NO"].astype(str),
                            pd.to_numeric(df_portfolio_report_recovery["SELL_PRICE"], errors='coerce'))
                    )
                    # Remove any NaN or invalid values from the mapping
                    recovery_lookup = {k: v for k, v in recovery_lookup.items() if pd.notna(v) and v > 0}

                    missing_mask = mask & (
                        df_prod_loans["Initial Valuation"].isna() |
                        (df_prod_loans["Initial Valuation"] == 0) |
                        (df_prod_loans["Initial Valuation"] == "") |
                        (df_prod_loans["Initial Valuation"] == "Not Valued") |
                        (df_prod_loans["Initial Valuation"] == "#N/A") |
                        (df_prod_loans["Initial Valuation"].astype(str).str.strip() == "") |
                        (df_prod_loans["Initial Valuation"].astype(str).str.upper() == "NOT VALUED") |
                        (df_prod_loans["Initial Valuation"].astype(str).str.upper() == "#N/A")
                    )

                    df_prod_loans.loc[missing_mask, "Initial Valuation"] = (
                        df_prod_loans.loc[missing_mask, "Contract No"].astype(str).map(recovery_lookup)
                    )

                    recovery_applied = df_prod_loans.loc[missing_mask, "Initial Valuation"].notna().sum()
                    logger.info(f"✓ Applied Initial Valuation from Portfolio Report Recovery (SELL_PRICE): {recovery_applied} contracts")
                else:
                    logger.warning("Portfolio Report Recovery data not available or missing required columns")
                
                # Step 3: Scienter Bot for ALL LR contracts with missing Initial Valuation
                still_missing_mask = mask & (
                    df_prod_loans["Initial Valuation"].isna() |
                    (df_prod_loans["Initial Valuation"] == 0) |
                    (df_prod_loans["Initial Valuation"] == "") |
                    (df_prod_loans["Initial Valuation"] == "Not Valued") |
                    (df_prod_loans["Initial Valuation"] == "#N/A") |
                    (df_prod_loans["Initial Valuation"].astype(str).str.strip() == "") |
                    (df_prod_loans["Initial Valuation"].astype(str).str.upper() == "NOT VALUED") |
                    (df_prod_loans["Initial Valuation"].astype(str).str.upper() == "#N/A")
                )

                # Filter for LR-type contracts only
                lr_mask = still_missing_mask & df_prod_loans["Contract No"].astype(str).str.startswith("LR")

                if lr_mask.any():
                    lr_contracts = df_prod_loans.loc[lr_mask, "Contract No"].astype(str).tolist()
                    logger.info(f"Found {len(lr_contracts)} LR-type contracts needing valuation")
                    logger.info(f"LR contracts to process: {lr_contracts}")

                    # Load Scienter bot
                    scienter_bot = self.load_scienter_valuation()

                    if scienter_bot:
                        try:
                            # Fetch valuations for ALL LR contracts from Scienter bot
                            logger.info(f"Sending {len(lr_contracts)} LR contracts to Scienter bot...")
                            lr_valuations = self.fetch_lr_valuations(lr_contracts, scienter_bot)

                            # Paste the returned data for each LR contract to Initial Valuation column
                            applied_count = 0
                            for contract, value in lr_valuations.items():
                                df_prod_loans.loc[df_prod_loans["Contract No"] == contract, "Initial Valuation"] = value
                                logger.info(f"Scienter bot: {contract} = {value}")
                                applied_count += 1

                            logger.info(f"✓ Applied {applied_count} valuations from Scienter bot for LR contracts")
                        except Exception as e:
                            logger.error(f"Error running Scienter bot: {e}")
                            logger.error(traceback.format_exc())
                    else:
                        logger.warning("Scienter bot not loaded, skipping LR contract valuation")
                else:
                    logger.info("No LR-type contracts require Scienter bot processing")
                
                # Step 4: initial_valuation_bot for ALL non-LR contracts with missing Initial Valuation
                # Update the mask to find non-LR contracts with missing values
                non_lr_missing_mask = mask & (
                    df_prod_loans["Initial Valuation"].isna() |
                    (df_prod_loans["Initial Valuation"] == 0) |
                    (df_prod_loans["Initial Valuation"] == "") |
                    (df_prod_loans["Initial Valuation"] == "Not Valued") |
                    (df_prod_loans["Initial Valuation"] == "#N/A") |
                    (df_prod_loans["Initial Valuation"].astype(str).str.strip() == "") |
                    (df_prod_loans["Initial Valuation"].astype(str).str.upper() == "NOT VALUED") |
                    (df_prod_loans["Initial Valuation"].astype(str).str.upper() == "#N/A")
                ) & ~df_prod_loans["Contract No"].astype(str).str.startswith("LR")  # Only non-LR contracts

                if non_lr_missing_mask.any() and initial_valuation_bot:
                    # Get ALL non-LR contracts that need valuation
                    contracts_to_fetch = df_prod_loans.loc[non_lr_missing_mask, "Contract No"].astype(str).tolist()

                    try:
                        logger.info(f"Fetching Initial Valuation for {len(contracts_to_fetch)} non-LR contracts via initial_valuation_bot")
                        logger.info(f"Non-LR contracts to process: {contracts_to_fetch[:20]}...")  # Show first 20
                        response = initial_valuation_bot.run_valuation_bot(contracts_to_fetch)

                        if isinstance(response, dict):
                            # Paste the returned data for each non-LR contract to Initial Valuation column
                            applied_count = 0
                            for contract, value in response.items():
                                # Only update if bot returned a valid value
                                if value and value not in ["", "Not Valued", "#N/A", "N/A"]:
                                    try:
                                        numeric_value = pd.to_numeric(value, errors='coerce')
                                        if pd.notna(numeric_value) and numeric_value > 0:
                                            df_prod_loans.loc[df_prod_loans["Contract No"] == contract, "Initial Valuation"] = numeric_value
                                            logger.info(f"initial_valuation_bot: {contract} = {numeric_value}")
                                            applied_count += 1
                                        else:
                                            logger.warning(f"initial_valuation_bot returned invalid/zero value for {contract}: {value}")
                                    except:
                                        logger.warning(f"initial_valuation_bot returned non-numeric value for {contract}: {value}")
                                else:
                                    logger.warning(f"initial_valuation_bot returned empty/N/A for {contract}: {value}")

                            logger.info(f"✓ Applied {applied_count} valuations from initial_valuation_bot for non-LR contracts")
                        else:
                            logger.warning(f"initial_valuation_bot returned unexpected response type: {type(response)}")
                    except Exception as e:
                        logger.error(f"Error running initial_valuation_bot: {e}")
                        logger.error(traceback.format_exc())
                
                # Log final statistics
                total_vehicles = mask.sum()
                valued_count = df_prod_loans.loc[mask, "Initial Valuation"].notna().sum()
                missing_count = total_vehicles - valued_count
                
                # Breakdown by contract type
                lr_total = df_prod_loans.loc[mask & df_prod_loans["Contract No"].astype(str).str.startswith("LR")].shape[0]
                lr_valued = df_prod_loans.loc[mask & df_prod_loans["Contract No"].astype(str).str.startswith("LR"), "Initial Valuation"].notna().sum()
                non_lr_total = df_prod_loans.loc[mask & ~df_prod_loans["Contract No"].astype(str).str.startswith("LR")].shape[0]
                non_lr_valued = df_prod_loans.loc[mask & ~df_prod_loans["Contract No"].astype(str).str.startswith("LR"), "Initial Valuation"].notna().sum()
                
                logger.info("="*60)
                logger.info("INITIAL VALUATION SUMMARY:")
                logger.info(f"  Total Vehicles & Machinery: {total_vehicles}")
                logger.info(f"  Total Valued: {valued_count} ({valued_count/total_vehicles*100:.1f}%)")
                logger.info(f"  Total Missing: {missing_count} ({missing_count/total_vehicles*100:.1f}%)")
                logger.info(f"  LR Contracts: {lr_valued}/{lr_total} valued")
                logger.info(f"  Non-LR Contracts: {non_lr_valued}/{non_lr_total} valued")
                logger.info("="*60)
                
                logger.info("✓ Initial Valuation processing completed")
            except Exception as e:
                logger.error(f"Error in Initial Valuation processing: {e}")
                logger.error(traceback.format_exc())
                df_prod_loans["Initial Valuation"] = None            

            # Other calculations with error handling
            try:
                # Annual Interest Cost ========================================================
                df_prod_loans["Annual Interest Cost"] = (
                    pd.to_numeric(df_prod_loans["Contract Amount"], errors="coerce") *
                    pd.to_numeric(df_prod_loans["EIR (%)"], errors="coerce") / 100
                )
                df_prod_loans["Annual Interest Cost"] = df_prod_loans["Annual Interest Cost"].round(0).astype("Int64")
                logger.info("✓ Annual Interest Cost calculated")
                
                # Gross Exposure ===========================================================
                contract_to_gross_exposure = dict(
                    zip(df_C1_C2_Working["Contract No"], df_C1_C2_Working["Gross Exposure"])
                )
                df_prod_loans["Gross Exposure"] = df_prod_loans["Contract No"].map(contract_to_gross_exposure)
                df_prod_loans["Gross Exposure"] = df_prod_loans["Gross Exposure"].fillna(df_prod_loans["Contract Amount"])
                df_prod_loans["Gross Exposure"] = (
                    pd.to_numeric(df_prod_loans["Gross Exposure"], errors="coerce")
                    .round(0)
                    .astype("Int64")
                )
                logger.info("✓ Gross Exposure calculated")
                
                # Normal/Concessionary classification ===============================================================
                prod_file_path = self._find_file("Prod. wise Class. of Loans")
                df_nbd_sheet = pd.read_excel(prod_file_path, sheet_name="NBD-MF-23-IA", header=None)
                l20_value = df_nbd_sheet.iloc[19, 11]
                
                df_staff_loan = pd.read_excel(prod_file_path, sheet_name="Staff_Loan", usecols="F", skiprows=3)
                icam_list = df_staff_loan["ICAM"].dropna().tolist()
                
                def calc_normal_concessionary(row):
                    try:
                        if row["Minimum Rate (Final)"] > l20_value:
                            return "Normal"
                        elif row["Client Code"] in icam_list:
                            return "Concessionary"
                        else:
                            return "Normal"
                    except:
                        return "Normal"
                
                df_prod_loans["Normal/Concessionary"] = self.safe_apply(
                    df_prod_loans, calc_normal_concessionary, default_value="Normal"
                )
                logger.info("✓ Normal/Concessionary classification completed")

                # PD Category already mapped earlier (before Collateral classification)
                logger.info("✓ PD Category mapping already completed earlier")
                
            except Exception as e:
                logger.error(f"Error in final calculations: {e}")
                logger.error(traceback.format_exc())
            
            logger.info("="*80)
            logger.info(f"build_prod_loans COMPLETED - Generated {len(df_prod_loans)} rows")
            logger.info(f"FINAL COLUMNS: {df_prod_loans.columns.tolist()}")
            logger.info("="*80)

            # Add LTV % column ================================================
            def calculate_ltv(row):
                try:
                    if row["Collateral/Security Type"] == "Vehicles and Machinery":
                        contract_amount = pd.to_numeric(row["Contract Amount"], errors='coerce')
                        initial_valuation = pd.to_numeric(row["Initial Valuation"], errors='coerce')

                        if (pd.notna(contract_amount) and pd.notna(initial_valuation) and
                            initial_valuation != 0 and initial_valuation is not None):
                            ltv_ratio = (contract_amount / initial_valuation) * 100  # Convert to percentage
                            return round(ltv_ratio, 2)  # Round to 2 decimal places
                    return 0.00
                except Exception as e:
                    logger.debug(f"Error calculating LTV for contract {row.get('Contract No', 'Unknown')}: {e}")
                    return 0.00

            df_prod_loans["LTV %"] = df_prod_loans.apply(calculate_ltv, axis=1)

            # WALTV % ==========================================================
            try:
                # Ensure Gross Exposure is numeric
                df_prod_loans["Gross Exposure"] = pd.to_numeric(df_prod_loans["Gross Exposure"], errors='coerce').fillna(0)

                # Compute total gross exposure grouped by Collateral/Security Type
                total_gross_by_type = df_prod_loans.groupby("Collateral/Security Type")["Gross Exposure"].transform("sum")

                # Formula directly
                df_prod_loans["WALTV %"] = 0.00
                mask = df_prod_loans["Collateral/Security Type"] == "Vehicles and Machinery"

                # Only calculate if we have valid data
                if mask.any() and total_gross_by_type[mask].sum() > 0:
                    waltv_calculation = (
                        df_prod_loans.loc[mask, "Gross Exposure"] / total_gross_by_type[mask]
                    ) * (df_prod_loans.loc[mask, "LTV %"] / 100)  # Convert LTV % back to ratio for calculation
                    df_prod_loans.loc[mask, "WALTV %"] = (waltv_calculation * 100).round(2)  # Convert to percentage with 2 decimals

                logger.info("✓ WALTV % calculation completed")
            except Exception as e:
                logger.error(f"Error in WALTV % calculation: {e}")
                df_prod_loans["WALTV %"] = 0


            # Check specifically for the missing columns
            required_columns = ["Collateral/Security Type", "Initial Valuation"]
            missing_columns = [col for col in required_columns if col not in df_prod_loans.columns]
            if missing_columns:
                logger.error(f"MISSING REQUIRED COLUMNS: {missing_columns}")
            else:
                logger.info("✓ All required columns present")

            # Get minimum rate threshold from master data and validate rates
            try:
                # Try to get from master data first, then use command line override if provided
                minimum_rate_final = self.get_minimum_rate_threshold(df_product_cat)

                # Use command line override if master data extraction failed
                if hasattr(self, 'minimum_rate_override') and self.minimum_rate_override is not None:
                    logger.info(f"Using command line minimum rate override: {self.minimum_rate_override}%")
                    minimum_rate_final = self.minimum_rate_override / 100

                exception_report_path = os.path.join(self.ia_folder, "Minimum_Rate_Exceptions.xlsx")

                logger.info(f"Validating Minimum Rate (Final) against threshold: {minimum_rate_final * 100}%")

                # Ensure required columns exist
                required_cols = ["Contract No", "Minimum Rate (Final)"]
                missing_cols = [col for col in required_cols if col not in df_prod_loans.columns]
                if missing_cols:
                    logger.warning(f"Missing columns for rate validation: {missing_cols}")
                else:
                    # Convert Minimum Rate (Final) to numeric for comparison
                    df_prod_loans["Minimum Rate (Final)"] = pd.to_numeric(df_prod_loans["Minimum Rate (Final)"], errors='coerce')

                    # Find exceptions where rate > threshold
                    exceptions_mask = df_prod_loans["Minimum Rate (Final)"] > minimum_rate_final
                    exceptions = df_prod_loans[exceptions_mask].copy()

                    if not exceptions.empty:
                        # Add reference column for clarity
                        exceptions["Expected Max Rate"] = minimum_rate_final
                        exceptions["Rate Difference"] = exceptions["Minimum Rate (Final)"] - minimum_rate_final

                        # Reorder columns for better readability
                        exception_cols = ["Contract No", "Product", "Client Code", "Minimum Rate (Final)",
                                        "Expected Max Rate", "Rate Difference", "Equipment", "Purpose"]
                        available_cols = [col for col in exception_cols if col in exceptions.columns]
                        remaining_cols = [col for col in exceptions.columns if col not in available_cols]
                        exceptions = exceptions[available_cols + remaining_cols]

                        # Export to Excel in same folder
                        exceptions.to_excel(exception_report_path, index=False)
                        logger.warning(f"Exception report generated at {exception_report_path} with {len(exceptions)} records exceeding {minimum_rate_final * 100}%")
                        logger.warning(f"Contracts with high rates: {exceptions['Contract No'].tolist()[:10]}...")  # Show first 10
                    else:
                        logger.info(f"✓ All {len(df_prod_loans)} contracts have rates within threshold ({minimum_rate_final * 100}%)")

            except Exception as e:
                logger.error(f"Error in minimum rate validation: {e}")
                logger.error(traceback.format_exc())
            
            # List of columns to divide by 100
            cols_to_scale = ["Minimum Rate (Final)", "Contractual Interest Rate", "Base Rate", "LTV %", "WALTV %"]

            # Divide numeric values by 100
            for col in cols_to_scale:
                if col in df_prod_loans.columns:
                    df_prod_loans[col] = pd.to_numeric(df_prod_loans[col], errors="coerce") / 100

            # Mask for FDL and Margin Trading rows (case-insensitive)
            mask_special = df_prod_loans["Product"].str.strip().str.lower().isin(["fdl", "margin trading"])

            # ---- Step 1: Apply transformations to special rows ----

            # Columns to clear
            cols_to_clear = ["Client Code", "Equipment", "Purpose", "Frequency", "Contract Period", "Tenure (Months)"]
            df_prod_loans.loc[mask_special, cols_to_clear] = ""

            # Assign Base Rate to interest columns for the special rows
            interest_cols = ["Minimum Rate (Final)", "Contractual Interest Rate", "EIR (%)"]
            df_prod_loans.loc[mask_special, interest_cols] = df_prod_loans.loc[mask_special, "Base Rate"]

            # Assign Contract Amount and Annual Interest Cost = Disbursed Amount (Net of DC)
            df_prod_loans.loc[mask_special, ["Contract Amount", "Annual Interest Cost"]] = df_prod_loans.loc[mask_special, "Disbursed Amount (Net of DC)"]

            # Additional adjustments for Margin Trading only
            mask_margin = df_prod_loans["Product"].str.strip().str.lower() == "margin trading"
            df_prod_loans.loc[mask_margin, "Corporate Clients"] = ""
            df_prod_loans.loc[mask_margin, "MSME Classification"] = "Small"

            # ---- Step 2: Move special rows to the bottom ----

            # Split into normal and special rows
            df_normal = df_prod_loans[~mask_special]
            df_special = df_prod_loans[mask_special]

            # Create an empty row (all blank)
            empty_row = pd.DataFrame([[""] * len(df_prod_loans.columns)], columns=df_prod_loans.columns)

            # Concatenate: normal rows + empty row + special rows
            df_prod_loans = pd.concat([df_normal, empty_row, df_special], ignore_index=True)


            # Columns to copy
            cols_to_copy = ["Product", "Equipment", "Purpose", "Corporate Clients"]

            # Filter rows where 'Type of Loans (1.3.1.0.0.0)' is '#N/A'
            mask_na_type = df_prod_loans["Type of Loans (1.3.1.0.0.0)"] == "#N/A"

            # Create new dataframe with the selected columns
            df_add_product_type = df_prod_loans.loc[mask_na_type, cols_to_copy].copy()

            # Remove duplicate rows (all columns duplicate)
            df_add_product_type.drop_duplicates(inplace=True)

            # Optional: reset index for clean sequential index
            df_add_product_type.reset_index(drop=True, inplace=True)


            # Rearrange columns in the specified order
            desired_column_order = [
                "Contract No",
                "Product",
                "Client Code",
                "Equipment",
                "Purpose",
                "Corporate Clients",
                "Type of Loans (1.3.1.0.0.0)",
                "Normal/Concessionary",
                "Frequency",
                "Contract Period",
                "Tenure (Months)",
                "Minimum Rate (Final)",
                "Contractual Interest Rate",
                "Base Rate",
                "Disbursed Amount (Net of DC)",
                "Contract Amount",
                "Annual Interest Cost",
                "EIR (%)",
                "Gross Exposure",
                "PD Category",
                "Collateral/Security Type",
                "Initial Valuation",
                "LTV %",
                "WALTV %",
                "MSME Classification"
            ]

            # Only include columns that exist in the dataframe
            existing_columns = [col for col in desired_column_order if col in df_prod_loans.columns]

            # Add any remaining columns that weren't specified
            remaining_columns = [col for col in df_prod_loans.columns if col not in existing_columns]
            final_column_order = existing_columns + remaining_columns

            df_prod_loans = df_prod_loans[final_column_order]
            logger.info(f"✓ Columns rearranged. Final order: {df_prod_loans.columns.tolist()}")

            # Final debugging - check what contracts we ended up with
            final_product_counts = df_prod_loans["Product"].value_counts()
            logger.info(f"FINAL CONTRACT COUNT: {len(df_prod_loans)} contracts")
            logger.info(f"FINAL Product distribution: {final_product_counts.to_dict()}")

            final_fdl_count = len(df_prod_loans[df_prod_loans["Product"] == "FDL"])
            final_margin_trading_count = len(df_prod_loans[df_prod_loans["Product"] == "Margin Trading"])
            logger.info(f"FINAL FDL contracts: {final_fdl_count}, FINAL Margin Trading contracts: {final_margin_trading_count}")

            print(df_prod_loans)
            return df_prod_loans
            
        except Exception as e:
            logger.error(f"CRITICAL ERROR in build_prod_loans: {e}")
            logger.error(traceback.format_exc())
            raise
    
    
    
    def save_report(self, df_prod_loans, picked_date):
        """Paste data to IA Working sheet using fast pandas Excel writing,
        check Product_Cat classification for missing values,
        and verify C39 of NBD-MF-23-IA sheet after pasting data."""
        try:
            # Find the Prod. wise Class. of Loans file
            prod_file_path = self._find_file("Prod. wise Class. of Loans")

            # Validate dataframe before saving
            if df_prod_loans.empty:
                logger.warning("DataFrame is empty. Nothing to paste.")
                return prod_file_path

            logger.info(f"Fast pasting {len(df_prod_loans)} rows x {len(df_prod_loans.columns)} columns to IA Working sheet")

            # Clean dataframe values
            df_clean = df_prod_loans.copy()
            for col in df_clean.columns:
                if pd.api.types.is_numeric_dtype(df_clean[col]):
                    df_clean[col] = df_clean[col].fillna(0)
                else:
                    df_clean[col] = df_clean[col].fillna("")

            # Extract unique combinations and paste to Product_Cat sheet
            logger.info("Extracting unique Product, Equipment, Purpose, Corporate Clients combinations...")
            cols_for_product_cat = ["Product", "Equipment", "Purpose", "Corporate Clients"]
            df_unique_combinations = df_clean[cols_for_product_cat].drop_duplicates().reset_index(drop=True)
            logger.info(f"Found {len(df_unique_combinations)} unique combinations")

            # Paste to Product_Cat sheet
            try:
                import xlwings as xw

                # Use xlwings to open the workbook
                app = xw.App(visible=False, add_book=False)
                app.display_alerts = False
                app.screen_updating = False

                try:
                    workbook = app.books.open(os.path.abspath(prod_file_path))
                    product_cat_sheet = workbook.sheets["Product_Cat"]

                    # Find the last row with data in Product_Cat sheet
                    last_row = product_cat_sheet.range('A1').expand('down').last_cell.row
                    logger.info(f"Product_Cat sheet currently has {last_row} rows")

                    # Determine starting row for new data (append below existing data)
                    start_row = last_row + 1

                    # Paste DataFrame directly
                    if not df_unique_combinations.empty:
                        product_cat_sheet.range(f'A{start_row}').value = df_unique_combinations.values
                        num_rows = len(df_unique_combinations)
                        logger.info(f"✓ Pasted {num_rows} unique combinations to Product_Cat sheet (rows {start_row} to {start_row + num_rows - 1})")

                    workbook.save()
                    workbook.close()
                finally:
                    app.quit()

            except Exception as product_cat_error:
                logger.error(f"Failed to paste data to Product_Cat sheet: {product_cat_error}")
                logger.error(traceback.format_exc())

            # Convert to absolute path
            abs_file_path = os.path.abspath(prod_file_path)
            logger.info(f"Working with file: {abs_file_path}")

            if not os.path.exists(abs_file_path):
                raise FileNotFoundError(f"File not found: {abs_file_path}")

            try:
                import xlwings as xw

                # Use xlwings to open the workbook
                app = xw.App(visible=False, add_book=False)
                app.display_alerts = False
                app.screen_updating = False

                try:
                    workbook = app.books.open(abs_file_path)

                    # === Step 1: Update IA Working sheet ===
                    worksheet = workbook.sheets["IA Working"]

                    # Clear existing data from row 3 onwards (keep headers in rows 1-2)
                    last_row = worksheet.range('A1').expand('down').last_cell.row
                    if last_row >= 3:
                        worksheet.range(f'A3:Z{last_row}').clear_contents()

                    # Write new data starting from row 3
                    if not df_clean.empty:
                        num_cols = min(len(df_clean.columns), 26)
                        worksheet.range('A3').value = df_clean.iloc[:, :num_cols].values

                    # === Step 2: Verify C39 in NBD-MF-23-IA sheet ===
                    try:
                        ia_sheet = workbook.sheets["NBD-MF-23-IA"]
                        c39_value = ia_sheet.range("C39").value
                        logger.info(f"C39 value in NBD-MF-23-IA: {c39_value}")

                        if c39_value not in (None, 0, "0"):
                            exception_data = pd.DataFrame(
                                [{"Sheet": "NBD-MF-23-IA", "Cell": "C39", "Value": c39_value}]
                            )
                            exception_filepath = os.path.join(
                                os.path.dirname(abs_file_path), "NBD-MF-23-IA_exceptions.xlsx"
                            )
                            exception_data.to_excel(exception_filepath, index=False)
                            logger.warning(f"⚠ Exception found (C39 != 0). Saved file: {exception_filepath}")
                        else:
                            logger.info("✓ C39 is 0. No exceptions.")

                    except Exception as check_error:
                        logger.error(f"Failed to check C39 value: {check_error}")

                    # Save and close workbook
                    workbook.save()
                    workbook.close()
                finally:
                    app.quit()

            except Exception as xlwings_error:
                logger.error(f"xlwings method failed: {xlwings_error}")
                logger.error(traceback.format_exc())
                raise

            # === Step 3: Rename the file with date format ===
            try:
                month_year = picked_date.strftime("%b %Y")
                file_dir = os.path.dirname(abs_file_path)
                new_filename = f"Prod. wise Class. of Loans - {month_year}.xlsb"
                new_file_path = os.path.join(file_dir, new_filename)

                os.rename(abs_file_path, new_file_path)
                logger.info(f"✓ File renamed to: {new_filename}")
                return new_file_path

            except Exception as e:
                logger.error(f"Failed to rename file: {e}")
                return abs_file_path

        except Exception as e:
            logger.error(f"Failed to paste data to IA Working sheet: {e}")
            logger.error(traceback.format_exc())
            raise


    
    def load_all(self, picked_date):
        """Orchestrator method with comprehensive error handling and recovery."""
        # Parse the date if it's not already a date object
        try:
            picked_date = self.parse_date(picked_date)
        except ValueError as e:
            logger.error(f"Date parsing error: {e}")
            raise

        logger.info("="*80)
        logger.info(f"Starting load_all process for date: {picked_date}")
        logger.info("="*80)
        
        loaded_data = {}
        errors = []
        
        # Define loading steps with fallback options
        loading_steps = [
            ("df_disbursement", self.load_disbursement, True),  # Required
            ("df_net_portfolio", self.load_net_portfolio, True),  # Required
            ("df_po_listing", self.load_po_listing, False),  # Optional
            ("df_information_request_from_credit", self.load_information_request_from_credit, True),  # Required
            ("df_product_cat", self.load_product_cat, True),  # Required
            ("df_C1_C2_Working", self.load_C1_C2_Working, True),  # Required
            ("df_yard_and_property_list", self.load_yard_and_Property_List, True),  # Required
            ("df_portfolio_report_recovery", self.load_portfolio_report_recovery, False),  # Optional
        ]
        
        # Load each dataset with error handling
        for name, loader_func, is_required in loading_steps:
            try:
                logger.info(f"Loading {name}...")
                loaded_data[name] = loader_func()
                logger.info(f"✓ Successfully loaded {name}")
            except Exception as e:
                error_msg = f"Failed to load {name}: {e}"
                logger.error(error_msg)
                errors.append((name, str(e)))
                
                if is_required:
                    logger.error(f"CRITICAL: Required dataset {name} could not be loaded")
                    raise RuntimeError(f"Cannot proceed without {name}. Error: {e}")
                else:
                    logger.warning(f"Optional dataset {name} failed to load, using empty DataFrame")
                    loaded_data[name] = pd.DataFrame()
        
        # Check if we have minimum required data
        required_datasets = ["df_disbursement", "df_net_portfolio", "df_product_cat"]
        missing_required = [ds for ds in required_datasets if ds not in loaded_data or loaded_data[ds].empty]
        
        if missing_required:
            error_msg = f"Missing required datasets: {missing_required}"
            logger.error(error_msg)
            raise RuntimeError(error_msg)
        
        # Build production loans
        try:
            logger.info("Starting build_prod_loans process...")
            df_prod_loans = self.build_prod_loans(
                loaded_data.get("df_disbursement"),
                loaded_data.get("df_net_portfolio"),
                loaded_data.get("df_information_request_from_credit", pd.DataFrame()),
                loaded_data.get("df_product_cat"),
                loaded_data.get("df_C1_C2_Working", pd.DataFrame()),
                loaded_data.get("df_yard_and_property_list", pd.DataFrame()),
                loaded_data.get("df_po_listing", pd.DataFrame()),
                loaded_data.get("df_portfolio_report_recovery", pd.DataFrame())
            )
            logger.info("✓ build_prod_loans completed successfully")
        except Exception as e:
            logger.error(f"Failed to build production loans: {e}")
            logger.error(traceback.format_exc())
            raise
        
        # Save report
        try:
            report_path = self.save_report(df_prod_loans, picked_date)
        except Exception as e:
            logger.error(f"Failed to save report: {e}")
            raise
        
        # Prepare return data
        month_year = picked_date.strftime("%b %Y")

        # Log summary
        logger.info("="*80)
        logger.info("PROCESS SUMMARY")
        logger.info(f"Date processed: {picked_date}")
        logger.info(f"Total rows generated: {len(df_prod_loans)}")
        logger.info(f"Data pasted to IA Working sheet in: {report_path}")
        if errors:
            logger.warning(f"Errors encountered during processing:")
            for name, error in errors:
                logger.warning(f"  - {name}: {error}")
        logger.info("="*80)

        return {
            "df_Prod_wise_Class_of_Loans": df_prod_loans,
            "df_Disbursement": loaded_data.get("df_disbursement"),
            "df_Net_Portfolio": loaded_data.get("df_net_portfolio"),
            "df_PO_Listing": loaded_data.get("df_po_listing"),
            "df_information_request_from_credit": loaded_data.get("df_information_request_from_credit"),
            "final_output_path": report_path,
            "report_path": report_path,
            "processing_errors": errors
        }
    
    def get_minimum_rate_threshold(self, df_product_cat):
        """
        Extract minimum rate threshold from master data file.

        Args:
            df_product_cat: Product category dataframe from master file

        Returns:
            float: Minimum rate threshold value
        """
        try:
            # Get the file path for Prod. wise Class. of Loans
            prod_file_path = self._find_file("Prod. wise Class. of Loans")

            # Read the NBD-MF-23-IA sheet to get L20 value (minimum rate threshold)
            df_nbd_sheet = pd.read_excel(prod_file_path, sheet_name="NBD-MF-23-IA", header=None)

            # Extract L20 value (row 19, column 11 in 0-indexed)
            minimum_rate_threshold = df_nbd_sheet.iloc[19, 11]

            # Convert to numeric and validate
            minimum_rate_threshold = pd.to_numeric(minimum_rate_threshold, errors='coerce')

            if pd.isna(minimum_rate_threshold):
                logger.warning("Could not extract minimum rate threshold from L20")
                # Check if command line override is available
                if hasattr(self, 'minimum_rate_override') and self.minimum_rate_override is not None:
                    logger.info(f"Using command line minimum rate override: {self.minimum_rate_override}%")
                    return self.minimum_rate_override / 100
                else:
                    logger.warning("No command line override provided, using default value of 0.06 (6.00%)")
                    return 0.06

            # Divide by 100 to convert percentage to decimal (e.g., 42 -> 0.42)
            minimum_rate_threshold = minimum_rate_threshold / 100

            logger.info(f"✓ Extracted minimum rate threshold from master data: {minimum_rate_threshold} (converted from {minimum_rate_threshold * 100}%)")
            return minimum_rate_threshold

        except Exception as e:
            logger.error(f"Error extracting minimum rate threshold: {e}")
            # Check if command line override is available
            if hasattr(self, 'minimum_rate_override') and self.minimum_rate_override is not None:
                logger.info(f"Using command line minimum rate override: {self.minimum_rate_override}%")
                return self.minimum_rate_override / 100
            else:
                logger.warning("No command line override provided, using default minimum rate threshold of 0.06 (6.00%)")
                return 0.06

    def load_scienter_valuation(self):
        """Load Scienter valuation bot with error handling."""
        try:
            logger.info("Loading IA_Working_Initial_valuation_Scienter_bot module...")
            import sys
            import os
            
            # Add the parent "bots" folder relative to this file
            current_dir = os.path.dirname(__file__)
            bot_dir = os.path.join(current_dir, "..", "bots")
            sys.path.insert(0, os.path.abspath(bot_dir))
            
            import IA_Working_Initial_valuation_Scienter_bot as scienter_bot
            logger.info("Successfully imported IA_Working_Initial_valuation_Scienter_bot module")
            return scienter_bot
        except ImportError as e:
            logger.warning(f"Could not import IA_Working_Initial_valuation_Scienter_bot: {e}")
            return None

    def fetch_lr_valuations(self, contracts_list, scienter_bot):
        """
        Fetch initial valuations for LR-type contracts using Scienter bot.
        
        Args:
            contracts_list: List of LR contract numbers
            scienter_bot: The imported Scienter bot module
            
        Returns:
            dict: Dictionary mapping contract numbers to their valuations
        """
        valuations = {}
        
        if not scienter_bot:
            logger.warning("Scienter bot not available, cannot fetch LR valuations")
            return valuations
        
        if not contracts_list:
            logger.info("No LR contracts to process")
            return valuations
        
        try:
            logger.info(f"Fetching valuations for {len(contracts_list)} LR contracts via Scienter bot")
            
            # Call the Scienter bot's valuation function
            # Adjust the function name based on the actual bot's API
            response = scienter_bot.run_valuation_bot(contracts_list)
            
            if isinstance(response, dict):
                for contract, value in response.items():
                    # Validate the value from bot
                    if value and value not in ["", "Not Valued", "#N/A", "N/A"]:
                        try:
                            numeric_value = pd.to_numeric(value, errors='coerce')
                            if pd.notna(numeric_value) and numeric_value > 0:
                                valuations[contract] = numeric_value
                                logger.debug(f"✓ Scienter bot: {contract} = {numeric_value}")
                            else:
                                logger.debug(f"Scienter bot returned invalid value for {contract}: {value}")
                        except:
                            logger.debug(f"Scienter bot returned non-numeric value for {contract}: {value}")
                    else:
                        logger.debug(f"Scienter bot returned empty/N/A for {contract}")
            else:
                logger.warning(f"Scienter bot returned unexpected response type: {type(response)}")
                
            logger.info(f"✓ Scienter bot returned {len(valuations)} valid valuations")
            
        except Exception as e:
            logger.error(f"Error running Scienter bot for LR contracts: {e}")
            logger.error(traceback.format_exc())
        
        return valuations

    def fill_incomplete_contracts(self, df_prod_loans, na_contract_bot):
        """Fill incomplete contracts with comprehensive error handling."""
        if na_contract_bot is None:
            logger.warning("Contract bot not available, skipping incomplete contract processing")
            return df_prod_loans
        
        try:
            critical_cols = [
                "Client Code", "Equipment", "Frequency",
                "Contract Period", "Contractual Interest Rate", "Contract Amount"
            ]
            
            # Find incomplete contracts
            mask = df_prod_loans[critical_cols].isnull().any(axis=1) | (df_prod_loans[critical_cols] == "").any(axis=1)
            incomplete_contracts = df_prod_loans.loc[mask, "Contract No"].dropna().unique()
            
            # Keep all incomplete contracts (removed LR/Margin Trading exclusion)
            incomplete_contracts = list(incomplete_contracts)
            
            if len(incomplete_contracts) == 0:
                logger.info("No incomplete contracts found")
                return df_prod_loans
            
            logger.info(f"Found {len(incomplete_contracts)} incomplete contracts to process")
            
            # Process contracts with bot
            try:
                session, verification_token = na_contract_bot.get_authenticated_session()
                updated_count = 0
                failed_contracts = []
                
                for i, cn in enumerate(incomplete_contracts, 1):
                    cn_str = str(cn).strip()
                    if not cn_str:
                        continue
                    
                    try:
                        logger.debug(f"Processing contract {i}/{len(incomplete_contracts)}: {cn_str}")
                        data, success = na_contract_bot.process_contract_with_retry(session, cn_str, verification_token)
                        
                        if success and isinstance(data, dict):
                            mapping = {
                                "Client Code": data.get("client_code", ""),
                                "Equipment": data.get("equipment", ""),
                                "Frequency": data.get("frequency", ""),
                                "Contract Period": data.get("contract_period", ""),
                                "Contractual Interest Rate": data.get("interest_rate", ""),
                                "Contract Amount": data.get("contract_amount", "")
                            }
                            
                            for k, v in mapping.items():
                                if v:  # Only update if we have a value
                                    df_prod_loans.loc[df_prod_loans["Contract No"] == cn_str, k] = v
                            
                            updated_count += 1
                            logger.debug(f"✓ Updated contract {cn_str}")
                        else:
                            failed_contracts.append(cn_str)
                            logger.warning(f"✗ Failed to get data for contract {cn_str}")
                            
                    except Exception as e:
                        logger.error(f"Error processing contract {cn_str}: {e}")
                        failed_contracts.append(cn_str)
                
                logger.info(f"Contract update summary: {updated_count} successful, {len(failed_contracts)} failed")
                if failed_contracts:
                    logger.warning(f"Failed contracts: {failed_contracts[:10]}...")  # Show first 10
                    
            except Exception as e:
                logger.error(f"Error in bot processing: {e}")
                logger.error(traceback.format_exc())
            
            return df_prod_loans
            
        except Exception as e:
            logger.error(f"Error in fill_incomplete_contracts: {e}")
            logger.error(traceback.format_exc())
            return df_prod_loans


def run_report(picked_date):
    """
    Function to be called by app.py or other modules.

    Args:
        picked_date: Can be a datetime.date object, string in "YYYY-MM-DD" format, or "MM/DD/YYYY" format

    Returns:
        Dictionary containing the processed data and file paths
    """
    try:
        logger.info("Starting NBD_MF23_IA_Report via run_report function")

        # Initialize report generator
        loader = NBD_MF23_IA_Report()

        # Load and process all data
        dfs = loader.load_all(picked_date)

        # Log results
        logger.info("✓ Processing completed successfully")
        logger.info(f"Final output location: {dfs['final_output_path']}")

        # Check for processing errors
        if dfs.get("processing_errors"):
            logger.warning("Some non-critical errors occurred during processing:")
            for name, error in dfs["processing_errors"]:
                logger.warning(f"  - {name}: {error}")

        return dfs

    except Exception as e:
        logger.error(f"FATAL ERROR in run_report: {e}")
        logger.error(traceback.format_exc())
        raise


if __name__ == "__main__":
    try:
        import argparse

        # Set up command line argument parsing
        parser = argparse.ArgumentParser(description='NBD MF23 IA Report Generator')

        # Support both standalone execution and app.py integration
        parser.add_argument('date', nargs='?', help='Processing date in MM/DD/YYYY or YYYY-MM-DD format')
        parser.add_argument('--base-dir', default=r"working\monthly",
                          help='Base directory for data files (default: working\\monthly)')
        parser.add_argument('--working-dir', help='Working directory (used by app.py)')
        parser.add_argument('--month', help='Report month (e.g., Jan) - used by app.py')
        parser.add_argument('--year', help='Report year (e.g., 2025) - used by app.py')
        parser.add_argument('--minimum-rate', type=float, metavar='RATE',
                          help='Minimum rate threshold (percentage) to use if master data is missing (e.g., 6.5)')
        parser.add_argument('--verbose', '-v', action='store_true',
                          help='Enable verbose logging')

        # Parse arguments
        args = parser.parse_args()

        if args.verbose:
            logging.getLogger().setLevel(logging.DEBUG)
            logger.info("Verbose logging enabled")

        logger.info("Starting NBD_MF23_IA_Report script")

        # Determine date: either from 'date' argument or from --month/--year
        picked_date = None
        if args.date:
            picked_date = args.date
            logger.info(f"Processing date: {args.date}")
        elif args.month and args.year:
            # Convert month name and year to date (first day of next month for app.py compatibility)
            month_num = datetime.strptime(args.month, "%b").month
            next_month = month_num + 1 if month_num < 12 else 1
            year = int(args.year) if month_num < 12 else int(args.year) + 1
            picked_date = f"{next_month:02d}/01/{year}"
            logger.info(f"Processing date derived from month/year: {picked_date} (for report month {args.month} {args.year})")
        else:
            raise ValueError("Either 'date' or '--month' and '--year' must be provided")

        # Determine directories
        if args.working_dir:
            # app.py provides full working directory path
            logger.info(f"Using working directory from app.py: {args.working_dir}")
            loader = NBD_MF23_IA_Report(working_dir=args.working_dir)
        else:
            # Standalone execution uses base_dir
            logger.info(f"Using base directory: {args.base_dir}")
            loader = NBD_MF23_IA_Report(base_dir=args.base_dir)

        if args.minimum_rate:
            logger.info(f"Minimum rate override: {args.minimum_rate}%")
            loader.minimum_rate_override = args.minimum_rate

        # Load and process all data
        dfs = loader.load_all(picked_date)

        # Log results
        logger.info("✓ Processing completed successfully")
        logger.info(f"Final output location: {dfs['final_output_path']}")

        # Check for processing errors
        if dfs.get("processing_errors"):
            logger.warning("Some non-critical errors occurred during processing:")
            for name, error in dfs["processing_errors"]:
                logger.warning(f"  - {name}: {error}")

        # Check if exception report was generated
        exception_file = os.path.join(loader.ia_folder, "Minimum_Rate_Exceptions.xlsx")
        if os.path.exists(exception_file):
            logger.info(f"Exception report generated: {exception_file}")

        logger.info("NBD MF23 IA Report generation completed successfully!")

    except SystemExit:
        # Allow argparse to handle help and error exits
        pass
    except Exception as e:
        logger.error(f"FATAL ERROR: {e}")
        logger.error(traceback.format_exc())
        sys.exit(1)