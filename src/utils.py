# src/utils.py

import os
import logging
import logging.handlers  # Add this import
import yaml
from datetime import datetime

def create_unique_filename(markdown_path: str, output_dir: str) -> str:
    """Generate unique output filename with timestamp"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    markdown_name = os.path.splitext(os.path.basename(markdown_path))[0]
    new_filename = f"{timestamp}_{markdown_name}.docx"
    return os.path.join(output_dir, new_filename)

def setup_logging():
    """Setup logging configuration"""
    config_path = os.path.join(os.path.dirname(__file__), '..', 'config', 'logging_config.yaml')
    
    if os.path.exists(config_path):
        try:
            with open(config_path, 'r') as f:
                config = yaml.safe_load(f)
                
            # Configure logging manually
            formatter = logging.Formatter(config['formatters']['detailed']['format'])
            
            # Console handler
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.DEBUG)
            console_handler.setFormatter(formatter)
            
            # File handler
            file_handler = logging.FileHandler("markdown_to_word_debug.log")
            file_handler.setLevel(logging.DEBUG)
            file_handler.setFormatter(formatter)
            
            # Root logger
            root_logger = logging.getLogger()
            root_logger.setLevel(logging.DEBUG)
            root_logger.addHandler(console_handler)
            root_logger.addHandler(file_handler)
            
        except Exception as e:
            _setup_default_logging()
            logging.warning(f"Failed to load logging config, using defaults: {str(e)}")
    else:
        _setup_default_logging()

def _setup_default_logging():
    """Setup default logging if config file is unavailable"""
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(levelname)s - %(message)s",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler("markdown_to_word_debug.log"),
        ],
    )