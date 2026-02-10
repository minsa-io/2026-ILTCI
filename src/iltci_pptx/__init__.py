"""ILTCI PowerPoint Presentation Generator package."""

from .layout_discovery import (
    LayoutRegistry,
    load_layout_registry,
    get_available_layout_names,
    validate_layout_name,
)
from .markdown_parser import (
    SlideData,
    parse_slides,
    parse_markdown_file,
    parse_document_frontmatter,
    parse_slide_frontmatter,
    SPACER_MARKER,
)
from .placeholder_resolver import (
    get_placeholder,
    get_placeholders,
    resolve_placeholders,
    PlaceholderNotFoundError,
)
from .slide_builders import (
    build_slide,
    populate_slide,
    build_rich_content,
)
from .images import (
    get_picture_placeholders,
    add_images_for_layout,
)

__all__ = [
    # Layout discovery
    "LayoutRegistry",
    "load_layout_registry",
    "get_available_layout_names",
    "validate_layout_name",
    # Markdown parsing
    "SlideData",
    "parse_slides",
    "parse_markdown_file",
    "parse_document_frontmatter",
    "parse_slide_frontmatter",
    "SPACER_MARKER",
    # Placeholder resolution
    "get_placeholder",
    "get_placeholders",
    "resolve_placeholders",
    "PlaceholderNotFoundError",
    # Slide building
    "build_slide",
    "populate_slide",
    "build_rich_content",
    # Image handling
    "get_picture_placeholders",
    "add_images_for_layout",
]
