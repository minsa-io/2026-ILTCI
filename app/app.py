"""Streamlit UI for PowerPoint Generator."""

import streamlit as st
import yaml
import sys
import tempfile
import shutil
import zipfile
import uuid
import logging
from pathlib import Path

# Add src directory to path for imports
app_dir = Path(__file__).parent
project_root = app_dir.parent
config_dir = project_root / "configs"
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from iltci_pptx.config import Config
from iltci_pptx.generator import PresentationGenerator


# Allowed file extensions for asset uploads
ALLOWED_ASSET_EXTENSIONS = {
    # Images
    '.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp', '.tiff',
    # Data/config
    '.yaml', '.yml', '.json',
    # Text/style
    '.css', '.txt', '.md',
}


def get_or_create_session_assets_dir() -> Path:
    """Get or create a session-scoped temporary assets directory.
    
    Uses the system temp directory to ensure session data is temporary
    and doesn't persist in the repository.
    
    Returns:
        Path to the session assets directory
    """
    if 'session_id' not in st.session_state:
        st.session_state.session_id = str(uuid.uuid4())
    
    # Check if we already have a temp directory for this session
    if 'session_assets_dir' not in st.session_state:
        # Create a new temp directory for this session
        temp_dir = tempfile.mkdtemp(prefix='iltci_assets_')
        st.session_state.session_assets_dir = temp_dir
        logging.info(f"Created temporary session assets directory: {temp_dir}")
    
    session_dir = Path(st.session_state.session_assets_dir)
    
    # Ensure directory exists (in case it was manually deleted)
    session_dir.mkdir(parents=True, exist_ok=True)
    return session_dir


def is_safe_filename(filename: str) -> bool:
    """Check if a filename is safe (no path traversal).
    
    Args:
        filename: The filename to check
        
    Returns:
        True if safe, False otherwise
    """
    # Reject path traversal
    if '..' in filename or filename.startswith('/'):
        return False
    
    # Reject Windows drive prefixes
    if ':' in filename and len(filename) > 2 and filename[1] == ':':
        return False
    
    return True


def save_uploaded_files(files, dest_dir: Path) -> list[str]:
    """Save uploaded files to destination directory, overwriting existing files.
    
    Args:
        files: List of Streamlit UploadedFile objects
        dest_dir: Destination directory
        
    Returns:
        List of saved filenames
        
    Raises:
        ValueError: If file has unsafe filename
    """
    saved_files = []
    for uploaded_file in files:
        filename = uploaded_file.name
        
        # Validate filename
        if not is_safe_filename(filename):
            raise ValueError(f"Unsafe filename: {filename}")
        
        # Check file extension
        ext = Path(filename).suffix.lower()
        if ext not in ALLOWED_ASSET_EXTENSIONS:
            st.warning(f"Skipping {filename}: unsupported file type")
            continue
        
        dest_path = dest_dir / filename
        
        # Write file (overwrites if exists)
        dest_path.write_bytes(uploaded_file.read())
        logging.info(f"Saved uploaded file: {dest_path}")
        saved_files.append(filename)
    
    return saved_files


def _strip_assets_prefix(rel_path: str) -> str:
    """Strip leading 'assets/' prefix from a path if present.
    
    This normalizes uploaded folder paths to match how resolve_asset_ref
    looks up assets (it also strips the 'assets/' prefix).
    
    Args:
        rel_path: Relative path that may have 'assets/' prefix
        
    Returns:
        Path with 'assets/' prefix stripped
    """
    # Normalize slashes
    normalized = rel_path.replace('\\', '/')
    
    # Strip leading './' if present
    if normalized.startswith('./'):
        normalized = normalized[2:]
    
    # Strip 'assets/' prefix if present
    if normalized.startswith('assets/'):
        normalized = normalized[7:]  # len('assets/') = 7
        logging.debug(f"Stripped 'assets/' prefix: {rel_path!r} -> {normalized!r}")
    
    return normalized


def save_uploaded_directory(files, dest_dir: Path) -> tuple[list[str], list[str]]:
    """Save uploaded directory files to a flat directory.
    
    When using folder upload, files have names like "assets/file.png".
    This function strips the 'assets/' prefix to save files flat in dest_dir,
    matching how resolve_asset_ref looks up assets.
    
    Args:
        files: List of Streamlit UploadedFile objects (from folder upload)
        dest_dir: Destination root directory
        
    Returns:
        Tuple of (saved_files, skipped_files)
        
    Raises:
        ValueError: If file has unsafe filename
    """
    skipped_files = []
    saved_files = []
    
    for uploaded_file in files:
        # Get the relative path (may include subdirectories like "assets/file.png")
        original_path = uploaded_file.name
        
        # Validate filename/path
        if not is_safe_filename(original_path):
            raise ValueError(f"Unsafe filename: {original_path}")
        
        # Strip 'assets/' prefix so files save flat (matching resolve_asset_ref behavior)
        rel_path = _strip_assets_prefix(original_path)
        
        # Check file extension
        ext = Path(rel_path).suffix.lower()
        if ext not in ALLOWED_ASSET_EXTENSIONS:
            skipped_files.append(original_path)
            continue
        
        # Build destination path (flat, no 'assets/' subdirectory)
        dest_path = dest_dir / rel_path
        
        # Create parent directories if needed (for any remaining subdirectories)
        dest_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Write file (overwrites if exists)
        dest_path.write_bytes(uploaded_file.read())
        logging.info(f"Saved directory file: {dest_path} (from {original_path})")
        saved_files.append(rel_path)
    
    return saved_files, skipped_files


def extract_zip(uploaded_zip, dest_dir: Path) -> list[str]:
    """Extract uploaded zip file to destination directory, stripping 'assets/' prefix.
    
    Args:
        uploaded_zip: Streamlit UploadedFile object for the zip
        dest_dir: Destination directory
        
    Returns:
        List of extracted file paths (with 'assets/' prefix stripped)
        
    Raises:
        ValueError: If zip contains unsafe paths
    """
    extracted_files = []
    with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
        for member in zip_ref.namelist():
            # Skip directories (they will be created when files are extracted)
            if member.endswith('/'):
                continue
            
            # Strip 'assets/' prefix to save flat (matching resolve_asset_ref behavior)
            stripped_member = _strip_assets_prefix(member)
            
            # Check for zip slip (path traversal) on the stripped path
            member_path = (dest_dir / stripped_member).resolve()
            if not str(member_path).startswith(str(dest_dir.resolve())):
                raise ValueError(f"Unsafe zip path: {member}")
            
            # Ensure parent directory exists
            member_path.parent.mkdir(parents=True, exist_ok=True)
            
            # Extract to stripped path (read from zip and write to new location)
            with zip_ref.open(member) as src, open(member_path, 'wb') as dst:
                dst.write(src.read())
            
            logging.info(f"Extracted zip member: {member} -> {stripped_member}")
            extracted_files.append(stripped_member)
    
    return extracted_files


def clear_session_assets(session_dir: Path) -> None:
    """Clear all assets in the session directory.
    
    Args:
        session_dir: Path to the session directory
    """
    if session_dir.exists():
        shutil.rmtree(session_dir)
        session_dir.mkdir(parents=True, exist_ok=True)
        logging.info(f"Cleared session assets: {session_dir}")


def get_session_asset_files(session_dir: Path) -> list[str]:
    """Get list of files in session assets directory.
    
    Args:
        session_dir: Path to the session directory
        
    Returns:
        List of relative file paths
    """
    if not session_dir.exists():
        return []
    
    files = []
    for item in session_dir.rglob('*'):
        if item.is_file():
            files.append(str(item.relative_to(session_dir)))
    return sorted(files)


def clear_upload_widget_state():
    """Clear the file uploader widget states from session_state."""
    keys_to_clear = ['file_folder_uploader', 'zip_uploader']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]


def sync_session_files_with_uploaders(session_dir: Path,
                                       uploaded_files: list | None) -> bool:
    """Sync session directory files with current uploader contents.
    
    When files are removed from the uploaders (via X button), this function
    detects the removal and deletes the corresponding files from disk.
    This ensures the disk state matches the uploader state.
    
    Args:
        session_dir: Path to the session assets directory
        uploaded_files: Current list of files in the file uploader (or None)
        
    Returns:
        True if any files were deleted, False otherwise
    """
    files_deleted = False
    
    # Get current file names from uploader, normalizing folder paths
    current_files = set()
    for f in (uploaded_files or []):
        # Check if it's a folder upload (has path separators)
        if '/' in f.name or '\\' in f.name:
            normalized = _strip_assets_prefix(f.name)
        else:
            normalized = f.name
        current_files.add(normalized)
    
    # Get tracked saved files
    saved_files = st.session_state.get('saved_files', set())
    
    # Find files that were removed from the uploader
    removed_files = saved_files - current_files
    for filename in removed_files:
        file_path = session_dir / filename
        if file_path.exists():
            try:
                file_path.unlink()
                logging.info(f"Removed file (no longer in uploader): {file_path}")
                files_deleted = True
            except Exception as e:
                logging.warning(f"Could not delete {file_path}: {e}")
    
    # Update tracking set to reflect current state
    if removed_files:
        st.session_state.saved_files = saved_files - removed_files
    
    # Clean up empty subdirectories
    if files_deleted:
        _cleanup_empty_dirs(session_dir)
    
    return files_deleted


def _cleanup_empty_dirs(session_dir: Path) -> None:
    """Remove empty subdirectories from session directory.
    
    Args:
        session_dir: Path to the session assets directory
    """
    if not session_dir.exists():
        return
    
    # Walk bottom-up to remove empty dirs
    for dirpath in sorted(session_dir.rglob('*'), reverse=True):
        if dirpath.is_dir() and not any(dirpath.iterdir()):
            try:
                dirpath.rmdir()
                logging.debug(f"Removed empty directory: {dirpath}")
            except Exception:
                pass


def handle_asset_uploads(session_dir: Path) -> bool:
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
            type=['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'tiff', 'yaml', 'yml', 'json', 'css', 'txt', 'md'],
            help="Upload individual files or drag & drop a folder. "
                 "If folder upload doesn't work in your browser, use the Zip Archive tab.",
            key="file_folder_uploader",
            label_visibility="collapsed"
        )
        
        # Process uploads immediately
        if uploaded_assets:
            # Track which files we've already saved this session to avoid re-saving
            saved_key = 'saved_files'
            if saved_key not in st.session_state:
                st.session_state[saved_key] = set()
            
            new_files = [f for f in uploaded_assets if f.name not in st.session_state[saved_key]]
            if new_files:
                try:
                    # Check if any files have folder paths (indicating folder upload)
                    has_folder_structure = any('/' in f.name or '\\' in f.name for f in new_files)
                    
                    if has_folder_structure:
                        saved, skipped = save_uploaded_directory(new_files, session_dir)
                        st.session_state[saved_key].update(saved)
                        if saved:
                            st.success(f"✓ Saved {len(saved)} file(s) from folder")
                            files_saved = True
                        if skipped:
                            st.warning(f"Skipped unsupported files: {', '.join(skipped)}")
                    else:
                        saved = save_uploaded_files(new_files, session_dir)
                        st.session_state[saved_key].update(saved)
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
            saved_key = 'saved_zip_files'
            if saved_key not in st.session_state:
                st.session_state[saved_key] = set()
            
            if uploaded_zip.name not in st.session_state[saved_key]:
                try:
                    extracted = extract_zip(uploaded_zip, session_dir)
                    st.session_state[saved_key].add(uploaded_zip.name)
                    if extracted:
                        st.success(f"✓ Extracted {len(extracted)} file(s) from {uploaded_zip.name}")
                        files_saved = True
                except ValueError as e:
                    st.error(f"❌ Error extracting zip: {e}")
    
    # Sync disk state with uploader state - remove files that were removed from uploaders
    # This ensures that when users click X to remove a file, it's also deleted from disk
    files_deleted = sync_session_files_with_uploaders(session_dir, uploaded_assets)
    
    return files_saved or files_deleted


def display_session_assets(session_dir: Path):
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


def load_base_config() -> dict:
    """Load the base configuration from config.yaml."""
    config_path = config_dir / "config.yaml"
    with open(config_path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def init_session_state():
    """Initialize session state with base configuration."""
    if 'base_config' not in st.session_state:
        st.session_state.base_config = load_base_config()
    if 'pptx_bytes' not in st.session_state:
        st.session_state.pptx_bytes = None
    if 'output_filename' not in st.session_state:
        st.session_state.output_filename = None
    if 'template_path' not in st.session_state:
        st.session_state.template_path = None
    # Track saved files to prevent re-saving on rerun
    if 'saved_files' not in st.session_state:
        st.session_state.saved_files = set()
    if 'saved_zip_files' not in st.session_state:
        st.session_state.saved_zip_files = set()


def main():
    """Main Streamlit application."""
    # Initialize session state
    init_session_state()
    
    base_config = st.session_state.base_config
    ui_config = base_config.get('ui', {})
    page_config = ui_config.get('page', {})
    defaults = ui_config.get('defaults', {})
    advanced_config = ui_config.get('advanced', {})
    
    # Page configuration
    st.set_page_config(
        page_title=page_config.get('title', 'PowerPoint Generator'),
        layout=page_config.get('layout', 'wide')
    )
    
    st.title("🎯 PowerPoint Generator")
    st.markdown("Generate professional PowerPoint presentations from Markdown content.")
    
    st.divider()
    
    # === Content Source Section ===
    st.subheader("📄 Content Source")
    
    col1, col2 = st.columns(2)
    
    with col1:
        content_source = st.radio(
            "Select content source:",
            ["Default", "Upload custom content"],
            horizontal=True,
            key="content_source"
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
            default_content = base_config.get('paths', {}).get('content', 'content/slides.md')
            st.info(f"Using default: `{default_content}`")
    
    st.divider()
    
    # === Assets Source Section ===
    st.subheader("🖼️ Assets Source")
    
    col_a1, col_a2 = st.columns(2)
    
    with col_a1:
        assets_source = st.radio(
            "Select assets source:",
            ["Default", "Upload custom assets"],
            horizontal=True,
            key="assets_source"
        )
    
    # Store assets_dir in session state for use during generation
    if assets_source == "Upload custom assets":
        # Get or create session directory
        session_assets_dir = get_or_create_session_assets_dir()
        st.session_state.custom_assets_dir = str(session_assets_dir)
        
        with col_a2:
            st.info("🖼️ Using custom assets")
        
        # Handle uploads immediately (consolidated handler)
        handle_asset_uploads(session_assets_dir)
        
        # Display current session assets
        display_session_assets(session_assets_dir)
        
        # Clear session assets button
        if st.button("🗑️ Clear custom assets", key="clear_assets"):
            clear_session_assets(session_assets_dir)
            # Clear the tracking sets so files can be re-uploaded
            st.session_state.saved_files = set()
            st.session_state.saved_zip_files = set()
            # Clear widget states to reset uploaders
            clear_upload_widget_state()
            st.success("✅ Assets cleared successfully!")
            st.rerun()
    else:
        st.session_state.custom_assets_dir = None
        with col_a2:
            default_assets = base_config.get('paths', {}).get('assets_dir', 'assets/')
            st.info(f"Using default: `{default_assets}`")
    
    st.divider()
    
    # === Template Source Section ===
    st.subheader("📑 Template Source")
    
    col_t1, col_t2 = st.columns(2)
    
    with col_t1:
        template_source = st.radio(
            "Select template source:",
            ["Default", "Upload custom template"],
            horizontal=True,
            key="template_source"
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
            default_template = base_config.get('paths', {}).get('template', 'templates/template.pptx')
            st.info(f"Using default: `{default_template}`")
    
    st.divider()
    
    # === Style Overrides Section ===
    st.subheader("🎨 Style Overrides")
    
    # Get default mode from config
    style_mode_options = ["none", "default", "custom overrides"]
    default_style_mode = ui_config.get("style_overrides_mode", "default")
    style_mode_index = style_mode_options.index(default_style_mode) if default_style_mode in style_mode_options else 0
    
    col_s1, col_s2 = st.columns(2)
    
    with col_s1:
        style_mode = st.radio(
            "Style overrides:",
            style_mode_options,
            key="style_mode",
            index=style_mode_index,
            horizontal=True
        )
    
    style_overrides = None
    if style_mode == "default":
        with col_s2:
            st.info("Using default: `configs/style-overrides.yaml`")
        try:
            with open(config_dir / "style-overrides.yaml") as f:
                style_overrides = yaml.safe_load(f)
        except FileNotFoundError:
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
    st.session_state.style_overrides = style_overrides
    
    st.divider()
    
    # === Output Configuration Section ===
    st.subheader("📤 Output Configuration")
    
    col3, col4 = st.columns(2)
    
    with col3:
        output_filename = st.text_input(
            "Output filename",
            value=defaults.get('output_filename', 'presentation.pptx'),
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
            value=base_config.get('settings', {}).get('overwrite_output', True),
            help="Allow overwriting if output file already exists"
        )
    
    # === Generate Button ===
    st.divider()
    
    generate_clicked = st.button("🚀 Generate PPTX", type="primary", use_container_width=True)
    if generate_clicked:
        # Validate inputs
        if content_source == "Upload custom content" and uploaded_file is None:
            st.error("❌ Please upload a Markdown file first.")
            return
        
        temp_dir = None
        temp_content_file = None
        temp_template_file = None
        
        try:
            # Build merged configuration
            import copy
            merged_config = copy.deepcopy(st.session_state.base_config)
            
            # Handle content path
            if content_source == "Upload custom content" and uploaded_file:
                # Write uploaded content to temp file
                temp_content_file = tempfile.NamedTemporaryFile(
                    mode='wb',
                    suffix='.md',
                    delete=False
                )
                temp_content_file.write(uploaded_file.read())
                temp_content_file.close()
                merged_config['paths']['content'] = temp_content_file.name
            # else: use default content path from config
            
            # Handle template path
            if template_source == "Upload custom template" and uploaded_template:
                # Write uploaded template to temp file
                temp_template_file = tempfile.NamedTemporaryFile(
                    mode='wb',
                    suffix='.pptx',
                    delete=False
                )
                temp_template_file.write(uploaded_template.read())
                temp_template_file.close()
                template_path = temp_template_file.name
                st.session_state.template_path = template_path
                merged_config['paths']['template'] = template_path
            # else: use default template path from config
            
            # Handle assets path - use session directory if custom assets selected
            if assets_source == "Upload custom assets" and st.session_state.custom_assets_dir:
                # Verify session directory has files
                session_dir = Path(st.session_state.custom_assets_dir)
                session_files = get_session_asset_files(session_dir)
                
                if session_files:
                    merged_config['paths']['assets_dir'] = st.session_state.custom_assets_dir
                    st.info(f"Using {len(session_files)} custom asset(s) from session directory")
                else:
                    st.warning("⚠️ No custom assets found. Using default assets directory.")
                    # Fall back to default assets_dir
            # else: use default assets_dir from config
            
            # Update settings
            merged_config['settings']['overwrite_output'] = overwrite
            merged_config['settings']['logging']['level'] = st.session_state.get('log_level', 'INFO')
            
            # Handle output path
            if use_temp_output:
                temp_dir = Path(tempfile.mkdtemp(prefix='iltci_pptx_'))
                output_path = temp_dir / output_filename
            else:
                output_path = project_root / "output" / output_filename
                output_path.parent.mkdir(parents=True, exist_ok=True)
            
            merged_config['paths']['output'] = str(output_path)
            
            # Remove styles_overrides path when "none" mode selected
            # This prevents Config.from_dict from merging any style overrides
            if style_mode == "none":
                if 'paths' in merged_config:
                    merged_config['paths'].pop('styles_overrides', None)
            
            # Create Config and generate
            with st.spinner('🔄 Generating presentation...'):
                cfg = Config.from_dict(merged_config, config_dir)
                generator = PresentationGenerator(cfg)
                generator.generate(style_overrides=st.session_state.style_overrides)
            
            # Read generated file
            pptx_bytes = output_path.read_bytes()
            st.session_state.pptx_bytes = pptx_bytes
            st.session_state.output_filename = output_filename
            
            st.success("✅ PowerPoint generated successfully!")
            
        except FileNotFoundError as e:
            st.error(f"❌ File not found: {e}")
        except Exception as e:
            st.error(f"❌ Generation failed: {e}")
            st.exception(e)
        
        finally:
            # Cleanup temp files
            if temp_content_file and Path(temp_content_file.name).exists():
                try:
                    Path(temp_content_file.name).unlink()
                except Exception:
                    pass
            
            if temp_template_file and Path(temp_template_file.name).exists():
                try:
                    Path(temp_template_file.name).unlink()
                except Exception:
                    pass
            
            if temp_dir and temp_dir.exists():
                try:
                    shutil.rmtree(temp_dir)
                except Exception:
                    pass
    
    # === Download Section (single location, no duplicates) ===
    if st.session_state.pptx_bytes is not None:
        st.divider()
        st.subheader("📥 Download")
        st.download_button(
            label=f"📥 Download {st.session_state.output_filename}",
            data=st.session_state.pptx_bytes,
            file_name=st.session_state.output_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
            key="download_button"
        )
    
    # === Advanced Settings (collapsible) - shown after Generate button ===
    st.divider()
    with st.expander("🔧 Advanced Settings", expanded=False):
        # Logging level (demoted to advanced)
        st.markdown("##### Logging")
        log_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR']
        current_level = base_config.get('settings', {}).get('logging', {}).get('level', 'INFO')
        default_index = log_levels.index(current_level) if current_level in log_levels else 1
        
        st.selectbox(
            "Log level",
            options=log_levels,
            index=default_index,
            key='log_level',
            help="Verbosity of logging output"
        )
        
        # Template paths (if enabled in config)
        if advanced_config.get('show_template_paths', False):
            st.markdown("##### Paths")
            st.caption("Template and configuration paths (relative to project root)")
            
            paths_config = base_config.get('paths', {})
            
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


if __name__ == "__main__":
    main()
