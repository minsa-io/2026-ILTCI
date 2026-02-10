"""Output configuration component."""

from typing import Any

import streamlit as st

from app.constants import DEFAULT_OUTPUT_FILENAME
from app.state import get_ui_config, get_settings_config


def render_output_config_section(base_config: dict[str, Any]) -> tuple[str, bool, bool]:
    """Render the output configuration section.
    
    Args:
        base_config: Base configuration dictionary
        
    Returns:
        Tuple of (output_filename, use_temp_output, overwrite)
    """
    st.subheader("📤 Output Configuration")
    
    ui_config = get_ui_config()
    defaults = ui_config.get('defaults', {})
    settings_config = get_settings_config()
    
    col3, col4 = st.columns(2)
    
    with col3:
        output_filename = st.text_input(
            "Output filename",
            value=defaults.get('output_filename', DEFAULT_OUTPUT_FILENAME),
            help="Name for the generated PPTX file"
        )
        
        if not output_filename.endswith('.pptx'):
            output_filename = output_filename + '.pptx'
            st.caption("ℹ️ .pptx extension will be added automatically")
    
    with col4:
        use_temp_output = st.checkbox(
            "Use temporary directory (recommended)",
            value=defaults.get('use_temp_output', True),
            help="Generate file in temp directory for clean download"
        )
        
        overwrite = st.checkbox(
            "Overwrite existing output",
            value=settings_config.get('overwrite_output', True),
            help="Allow overwriting if output file already exists"
        )
    
    st.divider()
    
    return output_filename, use_temp_output, overwrite
