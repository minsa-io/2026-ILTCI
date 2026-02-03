"""Unified slide construction functionality.

This module provides generic functions for building and populating slides
using the layout registry and placeholder resolver.
"""

from __future__ import annotations

import re
import logging
from typing import TYPE_CHECKING

from pptx.util import Pt

from .config import Config
from .images import add_images_for_layout
from .layout_discovery import LayoutRegistry, validate_layout_name
from .markdown_parser import SlideData, SPACER_MARKER
from .placeholder_resolver import resolve_placeholders, PlaceholderNotFoundError
from .rich_text import add_formatted_text, add_bullet, remove_bullet, add_numbering

if TYPE_CHECKING:
    from pptx.presentation import Presentation
    from pptx.slide import Slide
    from pptx.text.text import TextFrame

logger = logging.getLogger(__name__)


def build_slide(layout_name: str, prs: "Presentation", registry: LayoutRegistry) -> "Slide":
    """Build a slide using the specified layout from the registry.
    
    This is the primary function for creating new slides. It validates the
    layout name against the registry, retrieves the layout index, and creates
    a new slide.
    
    Args:
        layout_name: Exact layout name from the template (must exist in registry).
        prs: PowerPoint presentation object.
        registry: LayoutRegistry mapping layout names to indices.
        
    Returns:
        Created slide object.
        
    Raises:
        ValueError: If layout_name is not found in the registry.
        
    Example:
        >>> registry = load_layout_registry(prs)
        >>> slide = build_slide("Title and Content", prs, registry)
    """
    # Validate layout name against registry
    validate_layout_name(layout_name, registry, raise_on_missing=True)
    
    # Get layout index and create slide
    layout_idx = registry[layout_name]
    layout = prs.slide_layouts[layout_idx]
    slide = prs.slides.add_slide(layout)
    
    logger.info(f"Built slide using layout '{layout_name}' (index {layout_idx})")
    logger.debug(f"  Slide has {len(slide.shapes)} shapes")
    
    return slide


def populate_slide(
    slide: "Slide",
    data: SlideData,
    config: Config,
    registry: LayoutRegistry | None = None,
) -> None:
    """Populate a slide with content from SlideData.
    
    Uses the placeholder resolver to find title and content placeholders,
    then populates them with the provided data. If images are present and
    a registry is provided, uses add_images_for_layout for config-driven
    image placement.
    
    Args:
        slide: Slide object to populate.
        data: SlideData containing title, content_blocks, images, etc.
        config: Configuration object for fonts and formatting.
        registry: Optional LayoutRegistry for image placement validation.
        
    Raises:
        PlaceholderNotFoundError: If required placeholders cannot be found.
        
    Example:
        >>> slide = build_slide("Title and Content", prs, registry)
        >>> populate_slide(slide, slide_data, config, registry)
    """
    # Handle images using config-driven layout specs
    if data.images:
        if registry is None:
            logger.warning(
                f"Slide '{data.title or data.layout_name}' has {len(data.images)} images "
                "but no registry provided. Images will not be placed."
            )
        else:
            add_images_for_layout(data, slide, config, registry)
    
    # Try to resolve placeholders with fallback for title type variations
    # Title slides typically use CENTER_TITLE, content slides use TITLE
    phs: dict = {}
    title_ph = None
    content_ph = None
    
    # Try to find title placeholder (TITLE or CENTER_TITLE)
    if data.title:
        for title_type in ["title", "center_title"]:
            try:
                result = resolve_placeholders(slide, {"title": title_type})
                title_ph = result.get("title")
                if title_ph:
                    logger.debug(f"  Found title placeholder with type '{title_type}'")
                    break
            except PlaceholderNotFoundError:
                continue
        
        if title_ph is None:
            logger.warning(f"  No title placeholder found for slide '{data.layout_name}'")
    
    # Try to find content placeholder (BODY, OBJECT, or SUBTITLE for title slides)
    # Note: Some templates use OBJECT type for content placeholders instead of BODY
    if data.content_blocks:
        for content_type in ["body", "object", "subtitle"]:
            try:
                result = resolve_placeholders(slide, {"content": content_type})
                content_ph = result.get("content")
                if content_ph:
                    logger.debug(f"  Found content placeholder with type '{content_type}'")
                    break
            except PlaceholderNotFoundError:
                continue
        
        if content_ph is None:
            logger.warning(f"  No content placeholder found for slide '{data.layout_name}'")
    
    # Populate title if found
    if title_ph and data.title:
        title_ph.text_frame.text = data.title
        logger.debug(f"  Set title: '{data.title}'")
    
    # Populate content if found
    if content_ph and data.content_blocks:
        text_frame = content_ph.text_frame
        build_rich_content(text_frame, data.content_blocks, config)
        logger.debug(f"  Populated {len(data.content_blocks)} content blocks")


def build_rich_content(
    text_frame: "TextFrame",
    content_blocks: list[str],
    config: Config,
) -> None:
    """Build rich content from content blocks into a text frame.
    
    Processes content blocks (paragraphs, bullets, headers, spacers) and
    applies appropriate formatting based on configuration.
    
    Supports:
    - Headers: ## (H2), ### (H3), #### (H4), ##### (H5)
    - Bullets: - (level 0), "  - " (level 1)
    - Numbered lists: 1. 2. 3. etc.
    - Spacers: SPACER_MARKER for vertical spacing
    - Plain text with markdown formatting (**bold**, *italic*, etc.)
    
    Args:
        text_frame: PowerPoint text frame to populate.
        content_blocks: List of content strings to process.
        config: Configuration object for fonts and formatting.
        
    Example:
        >>> build_rich_content(text_frame, ["## Introduction", "- Point 1", "- Point 2"], config)
    """
    # Clear existing content
    text_frame.clear()
    
    # Get font sizes from config
    h2_size = config.get("fonts.content_slide.h2_header", 32)
    h3_size = config.get("fonts.content_slide.h3_header", 24)
    h4_size = config.get("fonts.content_slide.h4_header", 20)
    h5_size = config.get("fonts.content_slide.h5_header", 18)
    body_size = config.get("fonts.content_slide.body_text", 24)
    bullet_size = config.get("fonts.content_slide.bullet", 24)
    numbered_size = config.get("fonts.content_slide.numbered", 24)
    spacer_size = config.get("fonts.content_slide.spacer", 12)
    numbering_type = config.get("bullets.numbering_type", "arabicPeriod")
    
    # Get bold settings from config
    h2_bold = config.get("formatting.h2_bold", True)
    h3_bold = config.get("formatting.h3_bold", False)
    h4_bold = config.get("formatting.h4_bold", False)
    h5_bold = config.get("formatting.h5_bold", False)
    
    # Process each content block
    for block in content_blocks:
        # Each block may contain multiple lines
        for line in block.split("\n"):
            line_stripped = line.strip()
            if not line_stripped:
                continue
            
            _add_content_line(
                text_frame,
                line_stripped,
                h2_size=h2_size,
                h3_size=h3_size,
                h4_size=h4_size,
                h5_size=h5_size,
                body_size=body_size,
                bullet_size=bullet_size,
                numbered_size=numbered_size,
                spacer_size=spacer_size,
                numbering_type=numbering_type,
                h2_bold=h2_bold,
                h3_bold=h3_bold,
                h4_bold=h4_bold,
                h5_bold=h5_bold,
            )


def _add_content_line(
    text_frame: "TextFrame",
    line: str,
    *,
    h2_size: int,
    h3_size: int,
    h4_size: int,
    h5_size: int,
    body_size: int,
    bullet_size: int,
    numbered_size: int,
    spacer_size: int,
    numbering_type: str,
    h2_bold: bool,
    h3_bold: bool,
    h4_bold: bool,
    h5_bold: bool,
) -> None:
    """Add a single content line to a text frame with appropriate formatting.
    
    Internal helper function that handles the different line types:
    spacers, headers, bullets, numbered lists, and plain text.
    
    Args:
        text_frame: PowerPoint text frame to add content to.
        line: Single stripped line of content.
        h2_size - h5_bold: Formatting parameters from config.
    """
    # Handle spacer markers (blank lines in markdown)
    if line == SPACER_MARKER:
        p = text_frame.add_paragraph()
        # Use a space character to ensure the paragraph renders with height
        p.text = " "
        remove_bullet(p)
        for run in p.runs:
            run.font.size = Pt(spacer_size)
        p.space_before = Pt(spacer_size)
        p.space_after = Pt(0)
        logger.debug(f"  Added spacer paragraph ({spacer_size}pt)")
        return
    
    # Handle H5 header (##### )
    if line.startswith("##### "):
        p = text_frame.add_paragraph()
        add_formatted_text(p, line[6:])
        p.level = 0
        remove_bullet(p)
        for run in p.runs:
            run.font.size = Pt(h5_size)
            if h5_bold:
                run.font.bold = True
        logger.debug(f"  Added H5: {line[6:]}")
        return
    
    # Handle H4 header (#### )
    if line.startswith("#### "):
        p = text_frame.add_paragraph()
        add_formatted_text(p, line[5:])
        p.level = 0
        remove_bullet(p)
        for run in p.runs:
            run.font.size = Pt(h4_size)
            if h4_bold:
                run.font.bold = True
        logger.debug(f"  Added H4: {line[5:]}")
        return
    
    # Handle H3 header (### )
    if line.startswith("### "):
        p = text_frame.add_paragraph()
        add_formatted_text(p, line[4:])
        p.level = 0
        remove_bullet(p)
        for run in p.runs:
            run.font.size = Pt(h3_size)
            if h3_bold:
                run.font.bold = True
        logger.debug(f"  Added H3: {line[4:]}")
        return
    
    # Handle H2 header (## )
    if line.startswith("## "):
        p = text_frame.add_paragraph()
        add_formatted_text(p, line[3:])
        p.level = 0
        remove_bullet(p)
        for run in p.runs:
            run.font.size = Pt(h2_size)
            if h2_bold:
                run.font.bold = True
        logger.debug(f"  Added H2: {line[3:]}")
        return
    
    # Handle level-1 bullet (indented: "  - ")
    if line.startswith("  - "):
        p = text_frame.add_paragraph()
        add_formatted_text(p, line[4:])
        p.level = 1
        add_bullet(p, level=1)
        for run in p.runs:
            run.font.size = Pt(bullet_size)
        logger.debug(f"  Added sub-bullet: {line[4:]}")
        return
    
    # Handle level-0 bullet ("- ")
    if line.startswith("- "):
        p = text_frame.add_paragraph()
        add_formatted_text(p, line[2:])
        p.level = 0
        add_bullet(p, level=0)
        for run in p.runs:
            run.font.size = Pt(bullet_size)
        logger.debug(f"  Added bullet: {line[2:]}")
        return
    
    # Handle numbered lists (e.g., "1. ", "2. ")
    numbered_match = re.match(r"^(\d+)\.\s+(.*)$", line)
    if numbered_match:
        num = int(numbered_match.group(1))
        text = numbered_match.group(2)
        p = text_frame.add_paragraph()
        add_formatted_text(p, text)
        p.level = 0
        add_numbering(p, start_at=num, numbering_type=numbering_type)
        for run in p.runs:
            run.font.size = Pt(numbered_size)
        logger.debug(f"  Added numbered item {num}: {text}")
        return
    
    # Plain text (default)
    p = text_frame.add_paragraph()
    add_formatted_text(p, line)
    remove_bullet(p)
    for run in p.runs:
        run.font.size = Pt(body_size)
    logger.debug(f"  Added text: {line}")
