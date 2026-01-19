"""Streamlit UI for PowerPoint Generator."""

import streamlit as st
import yaml
import sys
import tempfile
import shutil
from pathlib import Path

# Add src directory to path for imports
app_dir = Path(__file__).parent
project_root = app_dir.parent
src_dir = project_root / "src"
if str(src_dir) not in sys.path:
    sys.path.insert(0, str(src_dir))

from iltci_pptx.config import Config
from iltci_pptx.generator import PresentationGenerator


def load_base_config() -> dict:
    """Load the base configuration from config.yaml."""
    config_path = app_dir / "config.yaml"
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
            ["Default", "Upload custom file"],
            horizontal=True,
            key="content_source"
        )
    
    uploaded_file = None
    if content_source == "Upload custom file":
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
        if content_source == "Upload custom file" and uploaded_file is None:
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
            if content_source == "Upload custom file" and uploaded_file:
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
            if template_source == "Upload" and uploaded_template:
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
            
            # Create Config and generate
            with st.spinner('🔄 Generating presentation...'):
                cfg = Config.from_dict(merged_config, app_dir)
                generator = PresentationGenerator(cfg)
                generator.generate()
            
            # Read generated file
            pptx_bytes = output_path.read_bytes()
            st.session_state.pptx_bytes = pptx_bytes
            st.session_state.output_filename = output_filename
            
            st.success("✅ PowerPoint generated successfully!")
            
            # Download button
            st.download_button(
                label="📥 Download PowerPoint",
                data=pptx_bytes,
                file_name=output_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
            
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
    
    # Show previous download if available (when not just generated)
    if st.session_state.pptx_bytes is not None and not generate_clicked:
        st.info("💾 Previous generation available for download:")
        st.download_button(
            label=f"📥 Download {st.session_state.output_filename}",
            data=st.session_state.pptx_bytes,
            file_name=st.session_state.output_filename,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True
        )


if __name__ == "__main__":
    main()
