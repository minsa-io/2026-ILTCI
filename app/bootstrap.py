"""Application bootstrap and initialization."""

import sys
from typing import Any

import streamlit as st

from app.constants import (
    SRC_DIR,
    SessionKeys,
    DEFAULT_PAGE_TITLE,
    DEFAULT_PAGE_LAYOUT,
)
from app.config_loader import load_base_config
from app.state import set_state_value, has_state_key


def setup_python_path() -> None:
    """Add src directory to Python path for imports.
    
    This must be called before importing from iltci_pptx.
    """
    if str(SRC_DIR) not in sys.path:
        sys.path.insert(0, str(SRC_DIR))


def init_session_state() -> None:
    """Initialize session state with base configuration.
    
    Sets up all required session state variables with defaults.
    """
    if not has_state_key(SessionKeys.BASE_CONFIG):
        set_state_value(SessionKeys.BASE_CONFIG, load_base_config())
    
    if not has_state_key(SessionKeys.PPTX_BYTES):
        set_state_value(SessionKeys.PPTX_BYTES, None)
    
    if not has_state_key(SessionKeys.OUTPUT_FILENAME):
        set_state_value(SessionKeys.OUTPUT_FILENAME, None)
    
    if not has_state_key(SessionKeys.TEMPLATE_PATH):
        set_state_value(SessionKeys.TEMPLATE_PATH, None)
    
    # Track saved files to prevent re-saving on rerun
    if not has_state_key(SessionKeys.SAVED_FILES):
        set_state_value(SessionKeys.SAVED_FILES, set())
    
    if not has_state_key(SessionKeys.SAVED_ZIP_FILES):
        set_state_value(SessionKeys.SAVED_ZIP_FILES, set())


def configure_page(base_config: dict[str, Any]) -> None:
    """Configure Streamlit page settings.
    
    Args:
        base_config: Base configuration dictionary
    """
    ui_config = base_config.get('ui', {})
    page_config = ui_config.get('page', {})
    
    st.set_page_config(
        page_title=page_config.get('title', DEFAULT_PAGE_TITLE),
        layout=page_config.get('layout', DEFAULT_PAGE_LAYOUT)
    )


def render_header() -> None:
    """Render the application header."""
    st.title("🎯 PowerPoint Generator")
    st.markdown("Generate professional PowerPoint presentations from Markdown content.")
    st.divider()


def bootstrap_app() -> dict[str, Any]:
    """Bootstrap the application.
    
    Sets up Python path, initializes session state, and configures the page.
    
    Returns:
        The base configuration dictionary
    """
    setup_python_path()
    init_session_state()
    
    base_config = st.session_state[SessionKeys.BASE_CONFIG]
    configure_page(base_config)
    render_header()
    
    return base_config
