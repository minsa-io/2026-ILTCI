"""Style overrides selection component."""

from typing import Any

import yaml
import streamlit as st

from app.constants import SessionKeys
from app.state import get_ui_config, set_style_overrides
from app.config_loader import load_style_overrides


def render_style_overrides_section(base_config: dict[str, Any]) -> str:
    """Render the style overrides selection section.
    
    Args:
        base_config: Base configuration dictionary
        
    Returns:
        The selected style mode ("none", "default", or "custom overrides")
    """
    st.subheader("🎨 Style Overrides")
    
    ui_config = get_ui_config()
    
    # Get default mode from config
    style_mode_options = ["none", "default", "custom overrides"]
    default_style_mode = ui_config.get("style_overrides_mode", "default")
    style_mode_index = style_mode_options.index(default_style_mode) if default_style_mode in style_mode_options else 0
    
    col_s1, col_s2 = st.columns(2)
    
    with col_s1:
        style_mode = st.radio(
            "Style overrides:",
            style_mode_options,
            key=SessionKeys.STYLE_MODE,
            index=style_mode_index,
            horizontal=True
        )
    
    style_overrides = None
    if style_mode == "default":
        with col_s2:
            st.info("Using default: `configs/style-overrides.yaml`")
        style_overrides = load_style_overrides()
        if style_overrides is None:
            st.warning("⚠️ Default style-overrides.yaml not found")
    elif style_mode == "custom overrides":
        with col_s2:
            uploaded_styles = st.file_uploader(
                "Upload custom styles YAML",
                type=["yaml", "yml"],
                key="styles_uploader"
            )
            if uploaded_styles:
                style_overrides = yaml.safe_load(uploaded_styles)
                st.success(f"✓ Loaded: {uploaded_styles.name}")
    else:
        with col_s2:
            st.info("No style overrides applied")
    
    # Store in session state for use during generation
    set_style_overrides(style_overrides)
    
    st.divider()
    
    return style_mode
