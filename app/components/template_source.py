"""Template source selection component."""

from typing import Any

import streamlit as st

from app.constants import SessionKeys, TEMPLATE_FILE_TYPES
from app.state import get_paths_config


def render_template_source_section(base_config: dict[str, Any]) -> tuple[str, Any]:
    """Render the template source selection section.
    
    Args:
        base_config: Base configuration dictionary
        
    Returns:
        Tuple of (template_source, uploaded_template)
    """
    st.subheader("📑 Template Source")
    
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        template_source = st.radio(
            "Select template source:",
            ["None", "Default", "Upload custom template"],
            index=1,
            horizontal=True,
            key=SessionKeys.TEMPLATE_SOURCE
        )
    
    uploaded_template = None
    if template_source == "Upload custom template":
        with col_t2:
            uploaded_template = st.file_uploader(
                "Upload template (.pptx / .potx)",
                type=TEMPLATE_FILE_TYPES,
                help="Upload a custom PowerPoint template file (.pptx or .potx)"
            )
            if uploaded_template:
                st.success(f"✓ Uploaded: {uploaded_template.name}")
    elif template_source == "Default":
        with col_t2:
            default_template = get_paths_config().get('template', 'templates/template.pptx')
            st.info(f"Using default: `{default_template}`")
    else:
        # "None" - use blank presentation with built-in layouts only
        with col_t2:
            st.info("Using blank presentation (built-in layouts only)")
    
    st.divider()
    
    return template_source, uploaded_template
