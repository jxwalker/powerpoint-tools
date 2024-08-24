import os
import json
import logging
from dataclasses import dataclass
import aiofiles

@dataclass
class Config:
    watson_api_key: str
    watson_service_url: str
    openai_api_key: str
    openai_model: str
    openai_max_tokens: int
    openai_temperature: float
    default_summarization_level: int
    anthropic_api_key: str
    anthropic_model: str
    anthropic_max_tokens: int
    min_characters: int
    rate_limit: float
    max_retries: int
    perplexity_api_key: str

async def load_config(config_path: str) -> Config:
    if not os.path.exists(config_path):
        logging.error(f"Configuration file not found: {config_path}")
        raise FileNotFoundError(f"Configuration file not found: {config_path}")

    if not os.access(config_path, os.R_OK):
        logging.error(f"No read permissions for configuration file: {config_path}")
        raise PermissionError(f"No read permissions for configuration file: {config_path}")

    try:
        async with aiofiles.open(config_path, 'r') as file:
            config_data = await file.read()
    except IOError as e:
        logging.error(f"Error reading configuration file {config_path}: {str(e)}")
        raise IOError(f"Failed to read configuration file: {str(e)}") from e

    try:
        config_dict = json.loads(config_data)
    except json.JSONDecodeError as e:
        logging.error(f"Error parsing configuration file {config_path}: {str(e)}")
        raise ValueError(f"Invalid JSON in configuration file: {str(e)}") from e

    expected_keys = set(Config.__annotations__.keys())
    missing_keys = expected_keys - set(config_dict.keys())
    if missing_keys:
        logging.error(f"Missing required configuration keys: {', '.join(missing_keys)}")
        raise ValueError(f"Missing required configuration keys: {', '.join(missing_keys)}")

    filtered_config = {k: v for k, v in config_dict.items() if k in expected_keys}
    
    try:
        return Config(**filtered_config)
    except TypeError as e:
        logging.error(f"Error creating Config object: {str(e)}")
        raise ValueError(f"Invalid configuration data: {str(e)}") from e

async def configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.WARNING
    logging.basicConfig(level=level, format="%(asctime)s - %(levelname)s - %(message)s")