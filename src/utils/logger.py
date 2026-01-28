import logging
import os
from datetime import datetime

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler("app.log", encoding='utf-8')
    ]
)

def log_action(message: str):
    """Logs an action with timestamp."""
    logging.info(message)
