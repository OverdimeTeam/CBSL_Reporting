"""
Configuration management using Pydantic for validation.
"""

from typing import Dict, Any, Optional
from pydantic import BaseSettings, BaseModel
import yaml
from pathlib import Path


class DatabaseConfig(BaseModel):
    """Database configuration."""
    host: str = "localhost"
    port: int = 5432
    database: str = "reporting_db"
    username: str = "postgres"
    password: str = ""


class SchedulerConfig(BaseModel):
    """Scheduler configuration."""
    timezone: str = "UTC"
    max_workers: int = 4
    job_defaults: Dict[str, Any] = {
        'coalesce': False,
        'max_instances': 1
    }


class BotConfig(BaseModel):
    """Bot configuration."""
    timeout: int = 30
    max_retries: int = 3
    retry_delay: int = 5


class ExcelConfig(BaseModel):
    """Excel processing configuration."""
    template_path: str = "excel/templates"
    output_path: str = "storage/reports"
    retention_months: int = 12


class SecurityConfig(BaseModel):
    """Security configuration."""
    jwt_secret: str = "your-secret-key"
    jwt_algorithm: str = "HS256"
    token_expire_hours: int = 24


class Config(BaseSettings):
    """Main application configuration."""
    
    database: DatabaseConfig = DatabaseConfig()
    scheduler: SchedulerConfig = SchedulerConfig()
    bots: BotConfig = BotConfig()
    excel: ExcelConfig = ExcelConfig()
    security: SecurityConfig = SecurityConfig()
    
    class Config:
        env_file = ".env"
        env_nested_delimiter = "__"
    
    def __init__(self):
        super().__init__()
        self.load_yaml_config()
    
    def load_yaml_config(self):
        """Load configuration from YAML file."""
        config_path = Path("config.yaml")
        if config_path.exists():
            with open(config_path, 'r') as f:
                yaml_config = yaml.safe_load(f)
                # TODO: Merge YAML config with Pydantic config
                pass