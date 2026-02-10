"""Asset management service - handles file uploads and session assets."""

import logging
import shutil
import tempfile
import uuid
import zipfile
from pathlib import Path

import streamlit as st

from app.constants import ALLOWED_ASSET_EXTENSIONS, SessionKeys
from app.utils.fs_safety import is_safe_filename, strip_assets_prefix
from app.state import get_state_value, set_state_value, has_state_key


def get_or_create_session_assets_dir() -> Path:
    """Get or create a session-scoped temporary assets directory.
    
    Uses the system temp directory to ensure session data is temporary
    and doesn't persist in the repository.
    
    Returns:
        Path to the session assets directory
    """
    if not has_state_key(SessionKeys.SESSION_ID):
        set_state_value(SessionKeys.SESSION_ID, str(uuid.uuid4()))
    
    # Check if we already have a temp directory for this session
    if not has_state_key(SessionKeys.SESSION_ASSETS_DIR):
        # Create a new temp directory for this session
        temp_dir = tempfile.mkdtemp(prefix='iltci_assets_')
        set_state_value(SessionKeys.SESSION_ASSETS_DIR, temp_dir)
        logging.info(f"Created temporary session assets directory: {temp_dir}")
    
    session_dir = Path(get_state_value(SessionKeys.SESSION_ASSETS_DIR))
    
    # Ensure directory exists (in case it was manually deleted)
    session_dir.mkdir(parents=True, exist_ok=True)
    return session_dir


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
        rel_path = strip_assets_prefix(original_path)
        
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
            stripped_member = strip_assets_prefix(member)
            
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


def sync_session_files_with_uploaders(
    session_dir: Path,
    uploaded_files: list | None
) -> bool:
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
            normalized = strip_assets_prefix(f.name)
        else:
            normalized = f.name
        current_files.add(normalized)
    
    # Get tracked saved files
    saved_files = get_state_value(SessionKeys.SAVED_FILES, set())
    
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
        set_state_value(SessionKeys.SAVED_FILES, saved_files - removed_files)
    
    # Clean up empty subdirectories
    if files_deleted:
        _cleanup_empty_dirs(session_dir)
    
    return files_deleted


def clear_upload_widget_state() -> None:
    """Clear the file uploader widget states from session_state."""
    keys_to_clear = ['file_folder_uploader', 'zip_uploader']
    for key in keys_to_clear:
        if key in st.session_state:
            del st.session_state[key]
