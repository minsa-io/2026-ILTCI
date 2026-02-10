"""Constants and configuration paths for the Streamlit app."""

from pathlib import Path

# === Directory Paths ===
APP_DIR = Path(__file__).parent
PROJECT_ROOT = APP_DIR.parent
CONFIG_DIR = PROJECT_ROOT / "configs"
SRC_DIR = PROJECT_ROOT / "src"

# === Allowed File Extensions ===
ALLOWED_ASSET_EXTENSIONS: frozenset[str] = frozenset({
    # Images
    '.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.tiff',
    # Data/config
    '.yaml', '.yml', '.json',
    # Text/style
    '.css', '.txt', '.md',
})

# File types for Streamlit uploaders (without dots)
ASSET_FILE_TYPES: list[str] = [
    'png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff',
    'yaml', 'yml', 'json', 'css', 'txt', 'md'
]

# === Session State Keys ===
class SessionKeys:
    """Session state key constants to avoid magic strings."""
    SESSION_ID = 'session_id'
    SESSION_ASSETS_DIR = 'session_assets_dir'
    CUSTOM_ASSETS_DIR = 'custom_assets_dir'
    BASE_CONFIG = 'base_config'
    PPTX_BYTES = 'pptx_bytes'
    OUTPUT_FILENAME = 'output_filename'
    TEMPLATE_PATH = 'template_path'
    SAVED_FILES = 'saved_files'
    SAVED_ZIP_FILES = 'saved_zip_files'
    STYLE_OVERRIDES = 'style_overrides'
    LOG_LEVEL = 'log_level'
    CONTENT_SOURCE = 'content_source'
    ASSETS_SOURCE = 'assets_source'
    TEMPLATE_SOURCE = 'template_source'
    STYLE_MODE = 'style_mode'


# === UI Configuration Defaults ===
DEFAULT_PAGE_TITLE = 'PowerPoint Generator'
DEFAULT_PAGE_LAYOUT = 'wide'
DEFAULT_OUTPUT_FILENAME = 'presentation.pptx'
DEFAULT_LOG_LEVEL = 'INFO'
LOG_LEVELS = ['DEBUG', 'INFO', 'WARNING', 'ERROR']
