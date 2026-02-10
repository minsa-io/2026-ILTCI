"""Generate button and generation logic component."""

from typing import Any

import streamlit as st

from app.services.generation_service import (
    generate_presentation,
    write_temp_file,
    cleanup_temp_file,
)


def render_generate_section(
    content_source: str,
    template_source: str,
    assets_source: str,
    style_mode: str,
    output_filename: str,
    use_temp_output: bool,
    overwrite: bool,
    uploaded_file: Any,
    uploaded_template: Any,
) -> bool:
    """Render the generate button and handle generation.
    
    Args:
        content_source: Selected content source
        template_source: Selected template source
        assets_source: Selected assets source
        style_mode: Selected style mode
        output_filename: Output filename
        use_temp_output: Whether to use temp output directory
        overwrite: Whether to allow overwriting
        uploaded_file: Uploaded content file (if any)
        uploaded_template: Uploaded template file (if any)
        
    Returns:
        True if generation was successful, False otherwise
    """
    generate_clicked = st.button(
        "🚀 Generate PPTX",
        type="primary",
        use_container_width=True
    )
    
    if not generate_clicked:
        return False
    
    # Validate inputs
    if content_source == "Upload custom content" and uploaded_file is None:
        st.error("❌ Please upload a Markdown file first.")
        return False
    
    temp_content_path: str | None = None
    temp_template_path: str | None = None
    
    try:
        # Handle uploaded content file
        if content_source == "Upload custom content" and uploaded_file:
            temp_content_path = write_temp_file(uploaded_file.read(), '.md')
        
        # Handle uploaded template file
        if template_source == "Upload custom template" and uploaded_template:
            temp_template_path = write_temp_file(uploaded_template.read(), '.pptx')
        
        # Generate presentation
        with st.spinner('🔄 Generating presentation...'):
            result = generate_presentation(
                content_source=content_source,
                template_source=template_source,
                assets_source=assets_source,
                style_mode=style_mode,
                output_filename=output_filename,
                use_temp_output=use_temp_output,
                overwrite=overwrite,
                uploaded_content_path=temp_content_path,
                uploaded_template_path=temp_template_path,
            )
        
        if result.success:
            st.success("✅ PowerPoint generated successfully!")
            return True
        else:
            st.error(f"❌ {result.error_message}")
            if result.exception:
                st.exception(result.exception)
            return False
            
    finally:
        # Cleanup temp files
        cleanup_temp_file(temp_content_path)
        cleanup_temp_file(temp_template_path)
