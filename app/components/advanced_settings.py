"""Advanced settings component."""

from typing import Any

import streamlit as st

from app.constants import LOG_LEVELS, DEFAULT_LOG_LEVEL, SessionKeys
from app.state import get_settings_config, get_ui_config, get_paths_config


def render_advanced_settings(base_config: dict[str, Any]) -> None:
    """Render the advanced settings expander.
    
    Args:
        base_config: Base configuration dictionary
    """
    st.divider()
    
    with st.expander("🔧 Advanced Settings", expanded=False):
        ui_config = get_ui_config()
        advanced_config = ui_config.get('advanced', {})
        settings_config = get_settings_config()
        
        # Logging level (demoted to advanced)
        st.markdown("##### Logging")
        current_level = settings_config.get('logging', {}).get('level', DEFAULT_LOG_LEVEL)
        default_index = LOG_LEVELS.index(current_level) if current_level in LOG_LEVELS else 1
        
        st.selectbox(
            "Log level",
            options=LOG_LEVELS,
            index=default_index,
            key=SessionKeys.LOG_LEVEL,
            help="Verbosity of logging output"
        )
        
        # Template paths (if enabled in config)
        if advanced_config.get('show_template_paths', False):
            st.markdown("##### Paths")
            st.caption("Template and configuration paths (relative to project root)")
            
            paths_config = get_paths_config()
            
            st.text_input(
                "Template path",
                value=paths_config.get('template', 'templates/template.pptx'),
                disabled=True,
                help="PowerPoint template file"
            )
            
            st.text_input(
                "Template config path",
                value=paths_config.get('template_config', 'assets/template-config.yaml'),
                disabled=True,
                help="Template styling configuration"
            )
            
            st.text_input(
                "Notes path (optional)",
                value=paths_config.get('notes', ''),
                disabled=True,
                help="Speaker notes file"
            )
