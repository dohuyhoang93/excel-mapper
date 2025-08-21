"""
Manages loading and saving of application and job-specific configurations.
"""
import json
import os
from pathlib import Path
from typing import Dict, Any, Optional
import logging
from datetime import datetime

logger = logging.getLogger(__name__)

class ConfigurationManager:
    """Handles reading and writing of configuration files."""

    def __init__(self, config_dir: str = "configs"):
        self.config_dir = Path(config_dir)
        self.config_dir.mkdir(exist_ok=True)
        self.app_settings_path = self.config_dir / "app_settings.json"

    def get_default_app_settings(self) -> Dict[str, Any]:
        """Returns a dictionary with default application settings."""
        return {
            "theme": "flatly",
            "detection_keywords": "total,sum,cộng,tổng,thành tiền",
            "last_source_file": "",
            "last_dest_file": ""
        }

    def load_app_settings(self) -> Dict[str, Any]:
        """
        Loads global application settings.
        Returns default settings if the file doesn't exist or is invalid.
        """
        if not self.app_settings_path.exists():
            return self.get_default_app_settings()
        try:
            with open(self.app_settings_path, 'r', encoding='utf-8') as f:
                settings = json.load(f)
                # Ensure all keys are present, add defaults for missing ones
                defaults = self.get_default_app_settings()
                for key, value in defaults.items():
                    if key not in settings:
                        settings[key] = value
                return settings
        except (json.JSONDecodeError, IOError) as e:
            logger.warning(f"Could not load app settings from {self.app_settings_path}: {e}. Returning defaults.")
            return self.get_default_app_settings()

    def save_app_settings(self, settings: Dict[str, Any]):
        """Saves global application settings."""
        try:
            with open(self.app_settings_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
        except IOError as e:
            logger.error(f"Could not save app settings to {self.app_settings_path}: {e}")

    def load_job_config(self, file_path: str) -> Dict[str, Any]:
        """Loads a job-specific configuration from a given path."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError, IOError) as e:
            logger.error(f"Failed to load job configuration from {file_path}: {e}")
            raise e # Re-raise to be caught by the UI layer

    def save_job_config(self, settings: Dict[str, Any], file_path: str):
        """Saves a job-specific configuration to a given path."""
        try:
            settings["created_date"] = datetime.now().isoformat()
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)
        except IOError as e:
            logger.error(f"Failed to save job configuration to {file_path}: {e}")
            raise e # Re-raise to be caught by the UI layer