"""File system safety utilities - pure functions for path validation."""

import logging


def is_safe_filename(filename: str) -> bool:
    """Check if a filename is safe (no path traversal).
    
    Pure function that validates filenames for security.
    
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


def strip_assets_prefix(rel_path: str) -> str:
    """Strip leading 'assets/' prefix from a path if present.
    
    This normalizes uploaded folder paths to match how resolve_asset_ref
    looks up assets (it also strips the 'assets/' prefix).
    
    Pure function for path normalization.
    
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
