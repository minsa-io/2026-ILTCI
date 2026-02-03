"""Download section component."""

import streamlit as st

from app.state import get_pptx_bytes, get_output_filename


def render_download_section() -> None:
    """Render the download section if a file is available."""
    pptx_bytes = get_pptx_bytes()
    output_filename = get_output_filename()
    
    if pptx_bytes is not None:
        st.divider()
        st.subheader("📥 Download")
        st.download_button(
            label=f"📥 Download {output_filename}",
            data=pptx_bytes,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key="download_button"
        )
