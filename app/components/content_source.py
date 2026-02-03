"""Content source selection component."""

from typing import Any

import streamlit as st

from app.constants import SessionKeys
from app.state import get_paths_config


def render_content_source_section(base_config: dict[str, Any]) -> tuple[str, Any]:
    """Render the content source selection section.
    
    Args:
        base_config: Base configuration dictionary
        
    Returns:
        Tuple of (content_source, uploaded_file)
    """
    st.subheader("📄 Content Source")
    
    col1, col2 = st.columns(2)
    
    with col1:
        content_source = st.radio(
            "Select content source:",
            ["Default", "Upload custom content"],
            horizontal=True,
            key=SessionKeys.CONTENT_SOURCE
        )
    
    uploaded_file = None
    if content_source == "Upload custom content":
        with col2:
            uploaded_file = st.file_uploader(
                "Upload Markdown file",
                type=['md', 'txt'],
                help="Upload a .md or .txt file containing slide content"
            )
            if uploaded_file:
                st.success(f"✓ Uploaded: {uploaded_file.name}")
    else:
        with col2:
            default_content = get_paths_config().get('content', 'content/slides.md')
            st.info(f"Using default: `{default_content}`")
    
    st.divider()
    
    return content_source, uploaded_file
