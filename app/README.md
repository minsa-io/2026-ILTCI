# Streamlit App for 2026 ILTCI PPTX Generator

This provides a user-friendly web UI to configure and generate PowerPoint presentations using the ILTCI library.

## Modular Architecture Overview

The app follows a modular architecture for maintainability and testability:

```
app/
├── app.py              # Thin entrypoint/orchestrator
├── bootstrap.py        # App initialization and page configuration
├── constants.py        # Constants, paths, and session keys
├── config_loader.py    # Configuration loading utilities
├── state.py            # Session state management with typed dataclasses
├── components/         # UI components
│   ├── __init__.py
│   ├── content_source.py      # Content source selection
│   ├── assets_source.py       # Assets source/upload management
│   ├── template_source.py     # Template source selection
│   ├── style_overrides.py     # Style overrides configuration
│   ├── output_config.py       # Output filename/options
│   ├── generate_button.py     # Generate button and logic
│   ├── download_section.py    # Download generated file
│   └── advanced_settings.py   # Logging and path settings
├── services/           # Business logic
│   ├── __init__.py
│   ├── assets_service.py      # Asset file handling (upload, extract, sync)
│   └── generation_service.py  # Presentation generation logic
└── utils/              # Pure utility functions
    ├── __init__.py
    └── fs_safety.py           # Filename validation and path safety
```

## Usage

1. Install dependencies (see root README.md): `uv sync`
2. Run the app: `uv run streamlit run app/app.py`
3. Select parameters in the UI (defaults from `config.yaml`).
4. Click Generate and download the PPTX.

## Module Descriptions

### Core Modules

- **[`app.py`](app.py)** - Thin orchestrator entrypoint that imports and calls modular components
- **[`bootstrap.py`](bootstrap.py)** - Sets up Python path, initializes session state, configures Streamlit page
- **[`constants.py`](constants.py)** - All constants including paths, allowed extensions, session keys, defaults
- **[`config_loader.py`](config_loader.py)** - Loads YAML configuration files (base config, style overrides)
- **[`state.py`](state.py)** - Typed dataclasses (`AppState`, `GenerationRequest`, `Choices`) and session state wrappers

### Services

- **[`services/assets_service.py`](services/assets_service.py)** - Handles all asset operations: session directory management, file uploads, zip extraction, file syncing
- **[`services/generation_service.py`](services/generation_service.py)** - Orchestrates presentation generation, builds merged config, calls `iltci_pptx` generator

### Components

Each component renders a UI section and returns relevant state:

- **[`components/content_source.py`](components/content_source.py)** - Content file selection (default or upload)
- **[`components/assets_source.py`](components/assets_source.py)** - Custom asset uploads (files, folders, zips)
- **[`components/template_source.py`](components/template_source.py)** - PowerPoint template selection
- **[`components/style_overrides.py`](components/style_overrides.py)** - Style override mode selection
- **[`components/output_config.py`](components/output_config.py)** - Output filename and generation options
- **[`components/generate_button.py`](components/generate_button.py)** - Generate button with validation
- **[`components/download_section.py`](components/download_section.py)** - Download button for generated file
- **[`components/advanced_settings.py`](components/advanced_settings.py)** - Log level and optional path display

### Utilities

- **[`utils/fs_safety.py`](utils/fs_safety.py)** - Pure functions for filename validation and path traversal prevention

## Configuration

Edit `config.yaml` directly or override via UI.

For CLI usage, the library now defaults to `app/config.yaml`.

## Generic Template Support

The app supports any PowerPoint template (`.pptx` or `.potx`). Layouts are discovered dynamically from the selected template. To use a custom template:

1. Place your template in the `templates/` directory.
2. Discover available layouts:
   ```bash
   python scripts/inspect_template.py templates/your_template.potx
   ```
3. Update [`assets/layout-specs.yaml`](../assets/layout-specs.yaml) with image placement specs for custom layouts.
4. Use layout names in your markdown frontmatter (see main README).

The underlying library uses [`layout_discovery`](../src/iltci_pptx/layout_discovery.py) and [`placeholder_resolver`](../src/iltci_pptx/placeholder_resolver.py) modules for template-agnostic slide generation.

See root [README.md](../README.md) for full project details, migration guide, and generic template documentation.

## Development

### Adding New Components

1. Create a new file in `components/` following the pattern of existing components
2. Export it from `components/__init__.py`
3. Import and call it in `app.py`

### Modifying Services

Services handle business logic separate from UI. Each service should:
- Be importable without Streamlit dependencies where possible
- Use state wrappers from `state.py` for session state access
- Use constants from `constants.py`

### Testing

Run the app locally to test changes:
```bash
uv run streamlit run app/app.py
```
