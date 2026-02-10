"""Template source selection component."""

from typing import Any

import streamlit as st

from app.constants import SessionKeys
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
            ["Default", "Upload custom template"],
            horizontal=True,
            key=SessionKeys.TEMPLATE_SOURCE
        )
    
    uploaded_template = None
    if template_source == "Upload custom template":
        with col_t2:
            uploaded_template = st.file_uploader(
                "Upload template.pptx",
                type=['pptx'],
                help="Upload a custom PowerPoint template file"
            )
            if uploaded_template:
                st.success(f"✓ Uploaded: {uploaded_template.name}")
    else:
        with col_t2:
            default_template = get_paths_config().get('template', 'templates/template.pptx')
            st.info(f"Using default: `{default_template}`")
    
    st.divider()
    
    return template_source, uploaded_template
