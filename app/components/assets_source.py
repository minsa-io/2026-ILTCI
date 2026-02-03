"""Assets source selection component."""

from pathlib import Path
from typing import Any

import streamlit as st

from app.constants import ASSET_FILE_TYPES, SessionKeys
from app.state import (
    get_paths_config,
    get_state_value,
    set_state_value,
    get_saved_files,
    set_saved_files,
    get_saved_zip_files,
    set_saved_zip_files,
)
from app.services.assets_service import (
    get_or_create_session_assets_dir,
    save_uploaded_files,
    save_uploaded_directory,
    extract_zip,
    clear_session_assets,
    get_session_asset_files,
    sync_session_files_with_uploaders,
    clear_upload_widget_state,
)


def _handle_asset_uploads(session_dir: Path) -> bool:
    """Handle all asset upload types and save files immediately.
    
    Provides two upload options via tabs:
    - Files/Folder: drag & drop individual files or an entire folder
    - Zip Archive: upload a zip file to extract
    
    Args:
        session_dir: Destination session directory
        
    Returns:
        True if any files were saved, False otherwise
    """
    files_saved = False
    
    # Use tabs to organize upload options
    tab_files, tab_zip = st.tabs(["📁 Files / Folder", "📦 Zip Archive"])
    
    with tab_files:
        st.caption("Drag & drop files or an entire folder. Supports images, CSS, YAML, JSON, and text files.")
        
        # Single unified uploader for files and folders
        uploaded_assets = st.file_uploader(
            "Drop files or folder here",
            accept_multiple_files=True,
            type=ASSET_FILE_TYPES,
            help="Upload individual files or drag & drop a folder. "
                 "If folder upload doesn't work in your browser, use the Zip Archive tab.",
            key="file_folder_uploader",
            label_visibility="collapsed"
        )
        
        # Process uploads immediately
        if uploaded_assets:
            saved_files = get_saved_files()
            new_files = [f for f in uploaded_assets if f.name not in saved_files]
            
            if new_files:
                try:
                    # Check if any files have folder paths (indicating folder upload)
                    has_folder_structure = any('/' in f.name or '\\' in f.name for f in new_files)
                    
                    if has_folder_structure:
                        saved, skipped = save_uploaded_directory(new_files, session_dir)
                        saved_files.update(saved)
                        set_saved_files(saved_files)
                        if saved:
                            st.success(f"✓ Saved {len(saved)} file(s) from folder")
                            files_saved = True
                        if skipped:
                            st.warning(f"Skipped unsupported files: {', '.join(skipped)}")
                    else:
                        saved = save_uploaded_files(new_files, session_dir)
                        saved_files.update(saved)
                        set_saved_files(saved_files)
                        if saved:
                            st.success(f"✓ Saved {len(saved)} file(s): {', '.join(saved)}")
                            files_saved = True
                except ValueError as e:
                    st.error(f"❌ Error saving files: {e}")
    
    with tab_zip:
        st.caption("Upload a zip archive containing your assets. Directory structure will be preserved.")
        
        # Zip uploader
        uploaded_zip = st.file_uploader(
            "Drop zip file here",
            type=['zip'],
            help="Upload a zip file containing assets. The folder structure inside will be preserved.",
            key="zip_uploader",
            label_visibility="collapsed"
        )
        
        # Process zip uploads immediately
        if uploaded_zip:
            saved_zip_files = get_saved_zip_files()
            
            if uploaded_zip.name not in saved_zip_files:
                try:
                    extracted = extract_zip(uploaded_zip, session_dir)
                    saved_zip_files.add(uploaded_zip.name)
                    set_saved_zip_files(saved_zip_files)
                    if extracted:
                        st.success(f"✓ Extracted {len(extracted)} file(s) from {uploaded_zip.name}")
                        files_saved = True
                except ValueError as e:
                    st.error(f"❌ Error extracting zip: {e}")
    
    # Sync disk state with uploader state - remove files that were removed from uploaders
    files_deleted = sync_session_files_with_uploaders(session_dir, uploaded_assets)
    
    return files_saved or files_deleted


def _display_session_assets(session_dir: Path) -> None:
    """Display the current session assets in an expander.
    
    Args:
        session_dir: Path to the session directory
    """
    files = get_session_asset_files(session_dir)
    
    if files:
        with st.expander(f"📁 Session Assets ({len(files)} files)", expanded=True):
            # Group files by type
            images = [f for f in files if Path(f).suffix.lower() in {'.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.tiff'}]
            configs = [f for f in files if Path(f).suffix.lower() in {'.yaml', '.yml', '.json'}]
            others = [f for f in files if f not in images and f not in configs]
            
            if images:
                st.markdown("**Images:**")
                for img in images:
                    st.text(f"  📷 {img}")
            
            if configs:
                st.markdown("**Config files:**")
                for cfg in configs:
                    st.text(f"  ⚙️ {cfg}")
            
            if others:
                st.markdown("**Other files:**")
                for other in others:
                    st.text(f"  📄 {other}")
    else:
        st.info("No custom assets uploaded yet. Upload files above to use custom assets.")


def render_assets_source_section(base_config: dict[str, Any]) -> str:
    """Render the assets source selection section.
    
    Args:
        base_config: Base configuration dictionary
        
    Returns:
        The selected assets source ("Default" or "Upload custom assets")
    """
    st.subheader("🖼️ Assets Source")
    
    col_a1, col_a2 = st.columns(2)
    
    with col_a1:
        assets_source = st.radio(
            "Select assets source:",
            ["Default", "Upload custom assets"],
            horizontal=True,
            key=SessionKeys.ASSETS_SOURCE
        )
    
    # Store assets_dir in session state for use during generation
    if assets_source == "Upload custom assets":
        # Get or create session directory
        session_assets_dir = get_or_create_session_assets_dir()
        set_state_value(SessionKeys.CUSTOM_ASSETS_DIR, str(session_assets_dir))
        
        with col_a2:
            st.info("🖼️ Using custom assets")
        
        # Handle uploads immediately (consolidated handler)
        _handle_asset_uploads(session_assets_dir)
        
        # Display current session assets
        _display_session_assets(session_assets_dir)
        
        # Clear session assets button
        if st.button("🗑️ Clear custom assets", key="clear_assets"):
            clear_session_assets(session_assets_dir)
            # Clear the tracking sets so files can be re-uploaded
            set_saved_files(set())
            set_saved_zip_files(set())
            # Clear widget states to reset uploaders
            clear_upload_widget_state()
            st.success("✅ Assets cleared successfully!")
            st.rerun()
    else:
        set_state_value(SessionKeys.CUSTOM_ASSETS_DIR, None)
        with col_a2:
            default_assets = get_paths_config().get('assets_dir', 'assets/')
            st.info(f"Using default: `{default_assets}`")
    
    st.divider()
    
    return assets_source
