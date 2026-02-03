"""Streamlit UI for PowerPoint Generator.

This is the main entrypoint for the Streamlit app. It serves as a thin
orchestrator that imports and calls modular components.

Architecture:
    - bootstrap.py: App initialization and page configuration
    - state.py: Session state management with typed dataclasses
    - config_loader.py: Configuration loading utilities
    - constants.py: Constants and paths
    - services/: Business logic (assets_service, generation_service)
    - components/: UI components (content_source, assets_source, etc.)
    - utils/: Pure utility functions (fs_safety)
"""

import sys
from pathlib import Path

# Add project root to Python path for module imports
# This is needed because Streamlit runs scripts in a way that doesn't
# automatically include the parent directory
_app_dir = Path(__file__).parent
_project_root = _app_dir.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

# Now we can import from the app package
from app.bootstrap import bootstrap_app

# Import UI components
from app.components import (
    render_content_source_section,
    render_assets_source_section,
    render_template_source_section,
    render_style_overrides_section,
    render_output_config_section,
    render_generate_section,
    render_download_section,
    render_advanced_settings,
)


def main():
    """Main Streamlit application.
    
    Orchestrates the UI by calling modular components in sequence.
    """
    # === Bootstrap ===
    # Sets up Python path, initializes session state, configures page
    base_config = bootstrap_app()
    
    # === Content Source Section ===
    content_source, uploaded_file = render_content_source_section(base_config)
    
    # === Assets Source Section ===
    assets_source = render_assets_source_section(base_config)
    
    # === Template Source Section ===
    template_source, uploaded_template = render_template_source_section(base_config)
    
    # === Style Overrides Section ===
    style_mode = render_style_overrides_section(base_config)
    
    # === Output Configuration Section ===
    output_filename, use_temp_output, overwrite = render_output_config_section(base_config)
    
    # === Generate Button ===
    render_generate_section(
        content_source=content_source,
        template_source=template_source,
        assets_source=assets_source,
        style_mode=style_mode,
        output_filename=output_filename,
        use_temp_output=use_temp_output,
        overwrite=overwrite,
        uploaded_file=uploaded_file,
        uploaded_template=uploaded_template,
    )
    
    # === Download Section ===
    render_download_section()
    
    # === Advanced Settings ===
    render_advanced_settings(base_config)


if __name__ == "__main__":
    main()
