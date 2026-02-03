"""Configuration loading utilities."""

from pathlib import Path
from typing import Any

import yaml

from app.constants import CONFIG_DIR


def load_base_config(config_path: Path | None = None) -> dict[str, Any]:
    """Load the base configuration from config.yaml.
    
    Args:
        config_path: Optional path to config file. Defaults to CONFIG_DIR / "config.yaml"
        
    Returns:
        Parsed configuration dictionary
        
    Raises:
        FileNotFoundError: If config file doesn't exist
    """
    if config_path is None:
        config_path = CONFIG_DIR / "config.yaml"
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def load_style_overrides(style_path: Path | None = None) -> dict[str, Any] | None:
    """Load style overrides from YAML file.
    
    Args:
        style_path: Optional path to style overrides. Defaults to CONFIG_DIR / "style-overrides.yaml"
        
    Returns:
        Parsed style overrides dictionary, or None if file not found
    """
    if style_path is None:
        style_path = CONFIG_DIR / "style-overrides.yaml"
    
    try:
        with open(style_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        return None
