"""
Main entry point for the Automated Reporting System
"""
import logging
import yaml
from dotenv import load_dotenv
from logs.logger import setup_logging
from orchestrator.scheduler import ReportScheduler
from interface.api import create_app

def load_config():
    """Load configuration from config.yaml"""
    with open('config.yaml', 'r') as file:
        return yaml.safe_load(file)

def main():
    # Load environment variables
    load_dotenv()
    
    # Setup logging
    setup_logging()
    logger = logging.getLogger(__name__)
    
    # Load configuration
    config = load_config()
    logger.info("Configuration loaded successfully")
    
    # Initialize scheduler
    scheduler = ReportScheduler(config)
    
    # Start FastAPI application
    app = create_app(config, scheduler)
    
    logger.info("Automated Reporting System started")
    return app

if __name__ == "__main__":
    import uvicorn
    app = main()
    uvicorn.run(app, host="0.0.0.0", port=8000)