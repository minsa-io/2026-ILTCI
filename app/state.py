"""Session state management with typed dataclasses."""

from dataclasses import dataclass, field
from typing import Any
from pathlib import Path

import streamlit as st

from app.constants import SessionKeys


@dataclass
class Choices:
    """User choices from UI selections."""
    content_source: str = "Default"  # "Default" or "Upload custom content"
    assets_source: str = "Default"   # "Default" or "Upload custom assets"
    template_source: str = "Default"  # "Default" or "Upload custom template"
    style_mode: str = "default"  # "none", "default", or "custom overrides"
    output_filename: str = "presentation.pptx"
    use_temp_output: bool = True
    overwrite: bool = True
    log_level: str = "INFO"


@dataclass
class GenerationRequest:
    """Parameters for a presentation generation request."""
    content_path: str | None = None
    template_path: str | None = None
    assets_dir: str | None = None
    output_path: str | None = None
    style_overrides: dict[str, Any] | None = None
    overwrite: bool = True
    log_level: str = "INFO"


@dataclass
class AppState:
    """Application state container."""
    base_config: dict[str, Any] = field(default_factory=dict)
    pptx_bytes: bytes | None = None
    output_filename: str | None = None
    template_path: str | None = None
    custom_assets_dir: str | None = None
    style_overrides: dict[str, Any] | None = None
    saved_files: set[str] = field(default_factory=set)
    saved_zip_files: set[str] = field(default_factory=set)


# === Session State Wrapper Functions ===

def get_state_value(key: str, default: Any = None) -> Any:
    """Get a value from session state with default.
    
    Args:
        key: Session state key
        default: Default value if key not found
        
    Returns:
        Value from session state or default
    """
    return st.session_state.get(key, default)


def set_state_value(key: str, value: Any) -> None:
    """Set a value in session state.
    
    Args:
        key: Session state key
        value: Value to set
    """
    st.session_state[key] = value


def has_state_key(key: str) -> bool:
    """Check if a key exists in session state.
    
    Args:
        key: Session state key
        
    Returns:
        True if key exists
    """
    return key in st.session_state


def delete_state_key(key: str) -> None:
    """Delete a key from session state if it exists.
    
    Args:
        key: Session state key to delete
    """
    if key in st.session_state:
        del st.session_state[key]


def get_base_config() -> dict[str, Any]:
    """Get the base configuration from session state."""
    return get_state_value(SessionKeys.BASE_CONFIG, {})


def get_ui_config() -> dict[str, Any]:
    """Get the UI configuration section."""
    return get_base_config().get('ui', {})


def get_paths_config() -> dict[str, Any]:
    """Get the paths configuration section."""
    return get_base_config().get('paths', {})


def get_settings_config() -> dict[str, Any]:
    """Get the settings configuration section."""
    return get_base_config().get('settings', {})


def get_pptx_bytes() -> bytes | None:
    """Get the generated PPTX bytes from session state."""
    return get_state_value(SessionKeys.PPTX_BYTES)


def set_pptx_bytes(data: bytes | None) -> None:
    """Set the generated PPTX bytes in session state."""
    set_state_value(SessionKeys.PPTX_BYTES, data)


def get_output_filename() -> str | None:
    """Get the output filename from session state."""
    return get_state_value(SessionKeys.OUTPUT_FILENAME)


def set_output_filename(filename: str | None) -> None:
    """Set the output filename in session state."""
    set_state_value(SessionKeys.OUTPUT_FILENAME, filename)


def get_custom_assets_dir() -> str | None:
    """Get the custom assets directory from session state."""
    return get_state_value(SessionKeys.CUSTOM_ASSETS_DIR)


def set_custom_assets_dir(path: str | None) -> None:
    """Set the custom assets directory in session state."""
    set_state_value(SessionKeys.CUSTOM_ASSETS_DIR, path)


def get_style_overrides() -> dict[str, Any] | None:
    """Get the style overrides from session state."""
    return get_state_value(SessionKeys.STYLE_OVERRIDES)


def set_style_overrides(overrides: dict[str, Any] | None) -> None:
    """Set the style overrides in session state."""
    set_state_value(SessionKeys.STYLE_OVERRIDES, overrides)


def get_saved_files() -> set[str]:
    """Get the set of saved files from session state."""
    return get_state_value(SessionKeys.SAVED_FILES, set())


def set_saved_files(files: set[str]) -> None:
    """Set the saved files set in session state."""
    set_state_value(SessionKeys.SAVED_FILES, files)


def get_saved_zip_files() -> set[str]:
    """Get the set of saved zip files from session state."""
    return get_state_value(SessionKeys.SAVED_ZIP_FILES, set())


def set_saved_zip_files(files: set[str]) -> None:
    """Set the saved zip files set in session state."""
    set_state_value(SessionKeys.SAVED_ZIP_FILES, files)


def get_log_level() -> str:
    """Get the current log level from session state."""
    return get_state_value(SessionKeys.LOG_LEVEL, 'INFO')
