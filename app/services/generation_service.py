"""Presentation generation service."""

import copy
import shutil
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import streamlit as st

from app.constants import CONFIG_DIR, PROJECT_ROOT, SessionKeys
from app.state import (
    get_state_value,
    set_state_value,
    get_base_config,
    get_style_overrides,
    set_pptx_bytes,
    set_output_filename,
)
from app.services.assets_service import get_session_asset_files


@dataclass
class GenerationResult:
    """Result of a presentation generation attempt."""
    success: bool
    pptx_bytes: bytes | None = None
    error_message: str | None = None
    exception: Exception | None = None


def _build_merged_config(
    content_source: str,
    template_source: str,
    assets_source: str,
    style_mode: str,
    output_filename: str,
    use_temp_output: bool,
    overwrite: bool,
    uploaded_content_path: str | None = None,
    uploaded_template_path: str | None = None,
) -> tuple[dict[str, Any], Path]:
    """Build the merged configuration for generation.
    
    Args:
        content_source: "Default" or "Upload custom content"
        template_source: "Default" or "Upload custom template"
        assets_source: "Default" or "Upload custom assets"
        style_mode: "none", "default", or "custom overrides"
        output_filename: Name for output file
        use_temp_output: Whether to use temp directory for output
        overwrite: Whether to allow overwriting
        uploaded_content_path: Path to uploaded content file (if any)
        uploaded_template_path: Path to uploaded template file (if any)
        
    Returns:
        Tuple of (merged_config, output_path)
    """
    merged_config = copy.deepcopy(get_base_config())
    
    # Handle content path
    if content_source == "Upload custom content" and uploaded_content_path:
        merged_config['paths']['content'] = uploaded_content_path
    
    # Handle template path
    if template_source == "Upload custom template" and uploaded_template_path:
        merged_config['paths']['template'] = uploaded_template_path
        set_state_value(SessionKeys.TEMPLATE_PATH, uploaded_template_path)
    
    # Handle assets path - use session directory if custom assets selected
    custom_assets_dir = get_state_value(SessionKeys.CUSTOM_ASSETS_DIR)
    if assets_source == "Upload custom assets" and custom_assets_dir:
        session_dir = Path(custom_assets_dir)
        session_files = get_session_asset_files(session_dir)
        
        if session_files:
            merged_config['paths']['assets_dir'] = custom_assets_dir
            st.info(f"Using {len(session_files)} custom asset(s) from session directory")
        else:
            st.warning("⚠️ No custom assets found. Using default assets directory.")
    
    # Update settings
    merged_config['settings']['overwrite_output'] = overwrite
    merged_config['settings']['logging']['level'] = get_state_value(SessionKeys.LOG_LEVEL, 'INFO')
    
    # Handle output path
    if use_temp_output:
        temp_dir = Path(tempfile.mkdtemp(prefix='iltci_pptx_'))
        output_path = temp_dir / output_filename
    else:
        output_path = PROJECT_ROOT / "output" / output_filename
        output_path.parent.mkdir(parents=True, exist_ok=True)
    
    merged_config['paths']['output'] = str(output_path)
    
    # Remove styles_overrides path when "none" mode selected
    if style_mode == "none":
        if 'paths' in merged_config:
            merged_config['paths'].pop('styles_overrides', None)
    
    return merged_config, output_path


def generate_presentation(
    content_source: str,
    template_source: str,
    assets_source: str,
    style_mode: str,
    output_filename: str,
    use_temp_output: bool,
    overwrite: bool,
    uploaded_content_path: str | None = None,
    uploaded_template_path: str | None = None,
) -> GenerationResult:
    """Generate a PowerPoint presentation.
    
    Args:
        content_source: "Default" or "Upload custom content"
        template_source: "Default" or "Upload custom template"
        assets_source: "Default" or "Upload custom assets"
        style_mode: "none", "default", or "custom overrides"
        output_filename: Name for output file
        use_temp_output: Whether to use temp directory for output
        overwrite: Whether to allow overwriting
        uploaded_content_path: Path to uploaded content file (if any)
        uploaded_template_path: Path to uploaded template file (if any)
        
    Returns:
        GenerationResult with success status and data
    """
    # Import here to avoid import issues before path setup
    from iltci_pptx.config import Config
    from iltci_pptx.generator import PresentationGenerator
    
    temp_dir: Path | None = None
    
    try:
        merged_config, output_path = _build_merged_config(
            content_source=content_source,
            template_source=template_source,
            assets_source=assets_source,
            style_mode=style_mode,
            output_filename=output_filename,
            use_temp_output=use_temp_output,
            overwrite=overwrite,
            uploaded_content_path=uploaded_content_path,
            uploaded_template_path=uploaded_template_path,
        )
        
        # Track temp_dir for cleanup
        if use_temp_output:
            temp_dir = output_path.parent
        
        # Create Config and generate
        cfg = Config.from_dict(merged_config, CONFIG_DIR)
        generator = PresentationGenerator(cfg)
        generator.generate(style_overrides=get_style_overrides())
        
        # Read generated file
        pptx_bytes = output_path.read_bytes()
        
        # Update session state
        set_pptx_bytes(pptx_bytes)
        set_output_filename(output_filename)
        
        return GenerationResult(
            success=True,
            pptx_bytes=pptx_bytes,
        )
        
    except FileNotFoundError as e:
        return GenerationResult(
            success=False,
            error_message=f"File not found: {e}",
            exception=e,
        )
    except Exception as e:
        return GenerationResult(
            success=False,
            error_message=f"Generation failed: {e}",
            exception=e,
        )
    finally:
        # Cleanup temp directory
        if temp_dir and temp_dir.exists():
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass


def write_temp_file(content: bytes, suffix: str) -> str:
    """Write content to a temporary file.
    
    Args:
        content: File content as bytes
        suffix: File suffix (e.g., '.md', '.pptx')
        
    Returns:
        Path to the temporary file
    """
    temp_file = tempfile.NamedTemporaryFile(
        mode='wb',
        suffix=suffix,
        delete=False
    )
    temp_file.write(content)
    temp_file.close()
    return temp_file.name


def cleanup_temp_file(path: str | None) -> None:
    """Clean up a temporary file.
    
    Args:
        path: Path to the temporary file (or None)
    """
    if path and Path(path).exists():
        try:
            Path(path).unlink()
        except Exception:
            pass
