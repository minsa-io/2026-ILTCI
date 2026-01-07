"""Configuration management for ILTCI presentation generator."""

import yaml
import logging
from pathlib import Path
from typing import Dict, Any


def load_yaml_file(file_path: Path) -> Dict[str, Any]:
    """Load a YAML file and return its contents."""
    if not file_path.exists():
        raise FileNotFoundError(f"Configuration file not found: {file_path}")
    
    with open(file_path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def merge_dicts(base: Dict[str, Any], overlay: Dict[str, Any]) -> Dict[str, Any]:
    """Recursively merge two dictionaries, overlay taking precedence."""
    result = base.copy()
    for key, value in overlay.items():
        if key in result and isinstance(result[key], dict) and isinstance(value, dict):
            result[key] = merge_dicts(result[key], value)
        else:
            result[key] = value
    return result


class Config:
    """Configuration manager that loads and merges main and template configs."""
    
    def __init__(self, config_path: str = 'config.yaml'):
        """Initialize configuration by loading main config and template config.
        
        Args:
            config_path: Path to main configuration file
        """
        self.config_path = Path(config_path)
        self._config = self._load_configuration()
        self._setup_logging()
    
    def _load_configuration(self) -> Dict[str, Any]:
        """Load and merge both configuration files."""
        # Load main config
        main_config = load_yaml_file(self.config_path)
        
        # Load template config
        template_config_path = Path(main_config['paths']['template_config'])
        template_config = load_yaml_file(template_config_path)
        
        # Merge configurations (template config is the base, main config overlays)
        merged = merge_dicts(template_config, main_config)
        
        logging.debug(f"Loaded main config from: {self.config_path}")
        logging.debug(f"Loaded template config from: {template_config_path}")
        
        return merged
    
    def _setup_logging(self):
        """Setup logging based on configuration."""
        log_level = self.get('settings.logging.level', 'INFO')
        numeric_level = getattr(logging, log_level.upper(), logging.INFO)
        logging.basicConfig(
            level=numeric_level,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
    
    def get(self, key_path: str, default: Any = None) -> Any:
        """Get configuration value using dot notation.
        
        Args:
            key_path: Dot-separated path to config value (e.g., 'fonts.title_slide.title')
            default: Default value if key not found
            
        Returns:
            Configuration value or default
        """
        keys = key_path.split('.')
        value = self._config
        
        for key in keys:
            if isinstance(value, dict) and key in value:
                value = value[key]
            else:
                return default
        
        return value
    
    def get_path(self, key: str) -> Path:
        """Get a path from configuration, ensuring it's a Path object.
        
        Args:
            key: Path key in config (e.g., 'template', 'content')
            
        Returns:
            Path object
        """
        path_str = self.get(f'paths.{key}')
        if path_str is None:
            raise ValueError(f"Path '{key}' not found in configuration")
        return Path(path_str)
    
    def validate_paths(self):
        """Validate that required paths exist."""
        required_paths = ['template', 'content']
        missing = []
        
        for path_key in required_paths:
            path = self.get_path(path_key)
            if not path.exists():
                missing.append(f"{path_key}: {path}")
        
        if missing:
            raise FileNotFoundError(
                f"Required files not found:\n" + "\n".join(f"  - {p}" for p in missing)
            )
    
    @property
    def template_path(self) -> Path:
        """Get template PowerPoint file path."""
        return self.get_path('template')
    
    @property
    def content_path(self) -> Path:
        """Get content markdown file path."""
        return self.get_path('content')
    
    @property
    def output_path(self) -> Path:
        """Get output PowerPoint file path."""
        return self.get_path('output')
    
    @property
    def assets_dir(self) -> Path:
        """Get assets directory path."""
        return self.get_path('assets_dir')
