# Streamlit App for 2026 ILTCI PPTX Generator

This provides a user-friendly web UI to configure and generate PowerPoint presentations using the ILTCI library.

## Usage

1. Install dependencies (see root README.md): `uv sync`
2. Run the app: `uv run streamlit run app/app.py`
3. Select parameters in the UI (defaults from `config.yaml`).
4. Click Generate and download the PPTX.

## Configuration

Edit `config.yaml` directly or override via UI.

For CLI usage, the library now defaults to `app/config.yaml`.

See root [README.md](../README.md) for full project details.
