"""Markdown parsing functionality for slide content.

This module parses markdown files with per-slide YAML frontmatter into
SlideData structures that can be used by the generator.

Expected format per slide:
    ---
    layout: "Layout Name"
    title: "Optional Override"
    images: ["path1.png", "path2.png"]
    ---
    
    # Slide Content
    - Bullet points...
"""

import re
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

from .config import Config
from .layout_discovery import LayoutRegistry, validate_layout_name, get_available_layout_names

logger = logging.getLogger(__name__)

# Spacer marker used to represent intentional blank lines for spacing
SPACER_MARKER = '<!-- spacer -->'


@dataclass
class SlideData:
    """Data structure representing a parsed slide.
    
    Attributes:
        layout_name: Exact layout name from the template (required).
        title: Slide title, either from frontmatter or first H1.
        content_blocks: List of content blocks (paragraphs, bullet lists, etc.).
        images: List of image paths specified in frontmatter or content.
        section_name: Optional section marker for navigation.
        raw_content: Original markdown content after frontmatter removal.
        options: Additional options from frontmatter (image_fit, bg_image, etc.).
    """
    layout_name: str
    title: str | None = None
    content_blocks: list[str] = field(default_factory=list)
    images: list[str] = field(default_factory=list)
    section_name: str = ""
    raw_content: str = ""
    options: dict[str, Any] = field(default_factory=dict)


def parse_document_frontmatter(content: str, delimiter: str = '---') -> tuple[dict[str, Any], str]:
    """Extract and parse document-level YAML frontmatter from markdown content.
    
    This parses the frontmatter at the very beginning of a markdown document,
    which contains document-level metadata (title, author, theme, etc.).
    
    Args:
        content: Full markdown content.
        delimiter: Frontmatter delimiter (default '---').
        
    Returns:
        Tuple of (frontmatter_dict, remaining_content).
    """
    lines = content.split('\n')
    
    # Document frontmatter must start at line 0
    if not lines or lines[0].strip() != delimiter:
        return {}, content
    
    end_idx = -1
    for i in range(1, len(lines)):
        if lines[i].strip() == delimiter:
            end_idx = i
            break
    
    if end_idx == -1:
        return {}, content
    
    # Extract frontmatter YAML
    frontmatter_lines = lines[1:end_idx]
    frontmatter_text = '\n'.join(frontmatter_lines)
    
    try:
        frontmatter = yaml.safe_load(frontmatter_text) or {}
    except yaml.YAMLError as e:
        logger.warning(f"Failed to parse document YAML frontmatter: {e}")
        frontmatter = {}
    
    # Remaining content after frontmatter
    remaining = '\n'.join(lines[end_idx + 1:])
    return frontmatter, remaining


def parse_slide_frontmatter(slide_content: str) -> tuple[dict[str, Any], str]:
    """Extract YAML frontmatter from the start of a slide.
    
    Per-slide frontmatter uses --- delimiters at the start of slide content:
    
        ---
        layout: "Title and Content"
        title: "My Title"
        images: ["img1.png"]
        ---
        
        Content here...
    
    Args:
        slide_content: Content of a single slide (after document split).
        
    Returns:
        Tuple of (frontmatter_dict, content_without_frontmatter).
    """
    lines = slide_content.strip().split('\n')
    
    if not lines:
        return {}, slide_content
    
    # Check if slide starts with ---
    if lines[0].strip() != '---':
        return {}, slide_content
    
    # Find closing ---
    end_idx = -1
    for i in range(1, len(lines)):
        if lines[i].strip() == '---':
            end_idx = i
            break
    
    if end_idx == -1:
        # No closing delimiter - not valid frontmatter
        return {}, slide_content
    
    # Extract and parse YAML
    yaml_lines = lines[1:end_idx]
    yaml_text = '\n'.join(yaml_lines)
    
    try:
        frontmatter = yaml.safe_load(yaml_text) or {}
    except yaml.YAMLError as e:
        logger.warning(f"Failed to parse slide YAML frontmatter: {e}")
        frontmatter = {}
    
    # Remaining content
    remaining = '\n'.join(lines[end_idx + 1:]).strip()
    return frontmatter, remaining


def _extract_images_from_content(content: str) -> list[str]:
    """Extract image paths from HTML image tags in content.
    
    Looks for <img src="..."> patterns and extracts paths.
    
    Args:
        content: Markdown/HTML content.
        
    Returns:
        List of image paths found.
    """
    images = []
    
    # Match HTML img tags with src attribute
    img_pattern = re.compile(r'<img[^>]+src=["\']([^"\']+)["\']', re.IGNORECASE)
    for match in img_pattern.finditer(content):
        images.append(match.group(1))
    
    # Also match markdown image syntax: ![alt](path)
    md_img_pattern = re.compile(r'!\[[^\]]*\]\(([^)]+)\)')
    for match in md_img_pattern.finditer(content):
        images.append(match.group(1))
    
    return images


def _parse_content_blocks(content: str) -> tuple[str | None, list[str], str]:
    """Parse slide content into title and content blocks.
    
    Args:
        content: Markdown content (frontmatter already removed).
        
    Returns:
        Tuple of (title, content_blocks, section_name).
    """
    lines = content.split('\n')
    title = None
    section_name = ""
    content_blocks = []
    has_content_started = False
    current_block: list[str] = []
    
    def flush_block():
        nonlocal current_block
        if current_block:
            block_text = '\n'.join(current_block).strip()
            if block_text:
                content_blocks.append(block_text)
            current_block = []
    
    for line in lines:
        line_stripped = line.strip()
        
        # Handle blank lines - preserve as spacers after content starts
        if not line_stripped:
            if has_content_started:
                # Flush current block and add spacer
                flush_block()
                content_blocks.append(SPACER_MARKER)
            continue
        
        # Check for section marker: <!-- section: Name -->
        section_match = re.match(r'<!--\s*section:\s*(.+?)\s*-->', line_stripped)
        if section_match:
            section_name = section_match.group(1)
            logger.debug(f"  -> Section name: {section_name}")
            continue
        
        # Extract title from H1 (# Title)
        if line_stripped.startswith('# ') and not line_stripped.startswith('## '):
            if title is None:
                title = line_stripped[2:].strip()
                logger.debug(f"  -> Title (H1): {title}")
            else:
                # Second H1 becomes content
                flush_block()
                current_block.append(line_stripped)
                has_content_started = True
            continue
        
        # H2 can be title if no H1, otherwise content
        if line_stripped.startswith('## '):
            if title is None:
                title = line_stripped[3:].strip()
                logger.debug(f"  -> Title (H2): {title}")
            else:
                flush_block()
                current_block.append(line_stripped)
                has_content_started = True
            continue
        
        # Everything else is content
        current_block.append(line_stripped)
        has_content_started = True
    
    # Flush remaining block
    flush_block()
    
    # Strip trailing spacer markers
    while content_blocks and content_blocks[-1] == SPACER_MARKER:
        content_blocks.pop()
    
    return title, content_blocks, section_name


def _split_into_slides(markdown_content: str, slide_separator: str = '---') -> list[str]:
    """Split markdown content into individual slide segments.
    
    Handles the complexity of `---` being used both as slide separator AND
    as YAML frontmatter delimiters within slides.
    
    Strategy: Process line by line, tracking whether we're inside frontmatter.
    A standalone `---` that is NOT part of frontmatter marks a slide boundary.
    
    Args:
        markdown_content: Markdown content (document frontmatter already removed).
        slide_separator: Separator between slides (default '---').
        
    Returns:
        List of slide content strings.
    """
    lines = markdown_content.split('\n')
    slides: list[str] = []
    current_slide_lines: list[str] = []
    in_frontmatter = False
    
    def _looks_like_yaml_start(line_idx: int) -> bool:
        """Check if lines following line_idx look like YAML content."""
        check_idx = line_idx + 1
        # Skip empty lines
        while check_idx < len(lines) and not lines[check_idx].strip():
            check_idx += 1
        if check_idx >= len(lines):
            return False
        next_line = lines[check_idx].strip()
        # YAML keys look like "key:" or "key: value" but not markdown headers
        return ':' in next_line and not next_line.startswith('#')
    
    def _has_only_blank_lines(line_list: list[str]) -> bool:
        """Check if list contains only blank/whitespace lines."""
        return all(not line.strip() for line in line_list)
    
    i = 0
    while i < len(lines):
        line = lines[i]
        line_stripped = line.strip()
        
        # Check for separator line
        if line_stripped == slide_separator:
            if in_frontmatter:
                # This closes the frontmatter block
                current_slide_lines.append(line)
                in_frontmatter = False
                i += 1
                continue
            
            # Not in frontmatter - is this a slide boundary or frontmatter start?
            # Check if current slide buffer is empty or only has blanks
            if not current_slide_lines or _has_only_blank_lines(current_slide_lines):
                # At start of a slide - check if this begins frontmatter
                if _looks_like_yaml_start(i):
                    # This is the opening of YAML frontmatter
                    in_frontmatter = True
                    # Clear any blank lines we might have accumulated
                    current_slide_lines = [line]
                    i += 1
                    continue
                else:
                    # Just a separator with no frontmatter following
                    # Skip it (don't start a new empty slide)
                    i += 1
                    continue
            else:
                # We have actual content - this is a slide boundary
                slide_content = '\n'.join(current_slide_lines).strip()
                if slide_content:
                    slides.append(slide_content)
                current_slide_lines = []
                
                # Now check if this same --- starts frontmatter for next slide
                if _looks_like_yaml_start(i):
                    in_frontmatter = True
                    current_slide_lines = [line]
                
                i += 1
                continue
        
        # Regular line - add to current slide
        current_slide_lines.append(line)
        i += 1
    
    # Don't forget the last slide
    if current_slide_lines:
        slide_content = '\n'.join(current_slide_lines).strip()
        if slide_content:
            slides.append(slide_content)
    
    return slides


def parse_slides(
    markdown_content: str,
    registry: LayoutRegistry,
    *,
    strict: bool = True,
    slide_separator: str = '---',
) -> list[SlideData]:
    """Parse markdown content into a list of SlideData objects.
    
    Each slide must have YAML frontmatter with at least a 'layout' field
    specifying the exact layout name from the template.
    
    Args:
        markdown_content: Full markdown content (document frontmatter removed).
        registry: LayoutRegistry for validating layout names.
        strict: If True, raise error for missing/invalid layouts.
                If False, log warnings and skip invalid slides.
        slide_separator: Separator between slides (default '---').
        
    Returns:
        List of SlideData objects, one per valid slide.
        
    Raises:
        ValueError: If strict=True and a slide has missing/invalid layout.
    """
    # Split into slide segments using smart separator detection
    raw_slides = _split_into_slides(markdown_content, slide_separator)
    
    slides: list[SlideData] = []
    available = get_available_layout_names(registry)
    
    for idx, raw_slide in enumerate(raw_slides):
        raw_slide = raw_slide.strip()
        if not raw_slide:
            logger.debug(f"Slide {idx}: Empty, skipping")
            continue
        
        logger.debug(f"--- Parsing Slide {idx} ---")
        logger.debug(f"Raw content preview: {raw_slide[:100]}...")
        
        # Parse per-slide frontmatter
        frontmatter, content = parse_slide_frontmatter(raw_slide)
        logger.debug(f"Frontmatter: {frontmatter}")
        
        # Extract layout name (required)
        layout_name = frontmatter.get('layout')
        
        if not layout_name:
            msg = (
                f"Slide {idx + 1} missing required 'layout' in frontmatter. "
                f"Available layouts: {', '.join(available)}"
            )
            if strict:
                raise ValueError(msg)
            logger.warning(msg)
            continue
        
        # Validate layout name against registry
        if not validate_layout_name(layout_name, registry, raise_on_missing=False):
            msg = (
                f"Slide {idx + 1}: Unknown layout '{layout_name}'. "
                f"Available layouts: {', '.join(available)}"
            )
            if strict:
                raise ValueError(msg)
            logger.warning(msg)
            continue
        
        # Parse content blocks
        title_from_content, content_blocks, section_name = _parse_content_blocks(content)
        
        # Title priority: frontmatter > first H1 in content
        title = frontmatter.get('title') or title_from_content
        
        # Extract images: frontmatter images list + images found in content
        images = list(frontmatter.get('images', []))
        images.extend(_extract_images_from_content(content))
        
        # Collect additional options
        options = {
            k: v for k, v in frontmatter.items()
            if k not in ('layout', 'title', 'images')
        }
        
        slide_data = SlideData(
            layout_name=layout_name,
            title=title,
            content_blocks=content_blocks,
            images=images,
            section_name=section_name,
            raw_content=content,
            options=options,
        )
        
        logger.debug(f"Created SlideData: layout={layout_name}, title={title}, "
                    f"blocks={len(content_blocks)}, images={len(images)}")
        slides.append(slide_data)
    
    logger.info(f"Parsed {len(slides)} slides from markdown")
    return slides


def parse_markdown_file(
    md_file: Path,
    registry: LayoutRegistry,
    config: Config | None = None,
    *,
    strict: bool = True,
) -> tuple[dict[str, Any], list[SlideData]]:
    """Parse a markdown file into document metadata and slides.
    
    This is the main entry point for parsing slide markdown files.
    
    Args:
        md_file: Path to the markdown file.
        registry: LayoutRegistry for validating layout names.
        config: Optional Config object for additional settings.
        strict: If True, raise on invalid layouts. If False, warn and skip.
        
    Returns:
        Tuple of (document_frontmatter, list_of_SlideData).
        
    Raises:
        FileNotFoundError: If markdown file doesn't exist.
        ValueError: If strict=True and slides have invalid layouts.
    """
    if not md_file.exists():
        raise FileNotFoundError(f"Markdown file not found: {md_file}")
    
    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    logger.info(f"Parsing markdown file: {md_file} ({len(content)} chars)")
    
    # Get separator from config if provided
    slide_separator = '---'
    if config:
        slide_separator = config.get('markdown.slide_separator', '---')
    
    # Parse document-level frontmatter
    doc_frontmatter, remaining_content = parse_document_frontmatter(content)
    logger.info(f"Document frontmatter keys: {list(doc_frontmatter.keys())}")
    
    # Parse individual slides
    slides = parse_slides(
        remaining_content,
        registry,
        strict=strict,
        slide_separator=slide_separator,
    )
    
    return doc_frontmatter, slides


# =============================================================================
# Legacy API (deprecated, for backward compatibility during migration)
# =============================================================================

def parse_yaml_frontmatter(content: str, delimiter: str = '---') -> tuple[dict[str, Any], str]:
    """[DEPRECATED] Use parse_document_frontmatter instead.
    
    This function is kept for backward compatibility.
    """
    logger.warning("parse_yaml_frontmatter is deprecated; use parse_document_frontmatter")
    return parse_document_frontmatter(content, delimiter)


def parse_markdown_slides(
    md_file: Path,
    config: Config
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """[DEPRECATED] Legacy parser that returns dict-based slide data.
    
    This function maintains backward compatibility with code that expects
    the old dict-based return format. New code should use parse_markdown_file
    which returns SlideData objects.
    
    Args:
        md_file: Path to markdown file.
        config: Configuration object.
        
    Returns:
        Tuple of (frontmatter_meta, list of parsed slide dictionaries).
        
    Note:
        This function does NOT validate layouts since it doesn't have
        access to a LayoutRegistry. It preserves the old dict format
        including is_title, subtitle, content keys.
    """
    logger.warning(
        "parse_markdown_slides is deprecated; use parse_markdown_file with LayoutRegistry"
    )
    
    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    frontmatter_delim = config.get('markdown.frontmatter_delimiter', '---')
    slide_separator = config.get('markdown.slide_separator', '---')
    title_class = config.get('markdown.title_class_marker', '<!-- _class: title -->')
    
    # Parse document frontmatter
    frontmatter, content = parse_document_frontmatter(content, frontmatter_delim)
    
    # Split by slide separators
    separator_pattern = f'\n{slide_separator}\n'
    raw_slides = re.split(separator_pattern, content)
    
    parsed_slides = []
    for idx, slide in enumerate(raw_slides):
        slide = slide.strip()
        if not slide:
            continue
        
        # Check for title class directive (legacy)
        is_title = title_class in slide
        slide = slide.replace(title_class, '').strip()
        
        # Try to parse slide frontmatter for layout directive
        fm, remaining = parse_slide_frontmatter(slide)
        layout = fm.get('layout')
        
        # Also check for legacy HTML comment directives
        if not layout:
            layout_match = re.search(r'<!--\s*_layout:\s*([^\s]+)\s*-->', remaining)
            if layout_match:
                layout = layout_match.group(1).strip()
                remaining = re.sub(r'<!--\s*_layout:\s*[^\s]+\s*-->', '', remaining).strip()
        
        # Parse image fit
        image_fit = fm.get('image_fit')
        if not image_fit:
            fit_match = re.search(r'<!--\s*_image_fit:\s*([^\s]+)\s*-->', remaining)
            if fit_match:
                image_fit = fit_match.group(1).strip()
                remaining = re.sub(r'<!--\s*_image_fit:\s*[^\s]+\s*-->', '', remaining).strip()
        
        # Parse bg_image
        bg_image = fm.get('bg_image')
        if not bg_image:
            bg_match = re.search(r'<!--\s*_bg_image:\s*([^\s]+)\s*-->', remaining)
            if bg_match:
                bg_image = bg_match.group(1).strip()
                remaining = re.sub(r'<!--\s*_bg_image:\s*[^\s]+\s*-->', '', remaining).strip()
        
        # Parse content using old logic
        slide_data = _parse_legacy_slide_content(remaining, is_title)
        
        # Add directive values
        if layout:
            slide_data['layout'] = layout
        if image_fit:
            slide_data['image_fit'] = image_fit
        if bg_image:
            slide_data['bg_image'] = bg_image
        
        parsed_slides.append(slide_data)
    
    return frontmatter, parsed_slides


def _parse_legacy_slide_content(slide: str, is_title: bool) -> dict[str, Any]:
    """Parse slide content into legacy dict format.
    
    This preserves the old behavior for backward compatibility.
    """
    lines = slide.split('\n')
    title = ''
    subtitle_lines: list[str] = []
    content_lines: list[str] = []
    section_name = ''
    has_content_started = False
    
    for line in lines:
        line_stripped = line.strip()
        
        if not line_stripped:
            if has_content_started and not is_title:
                content_lines.append(SPACER_MARKER)
            elif has_content_started and is_title:
                subtitle_lines.append(SPACER_MARKER)
            continue
        
        line = line_stripped
        
        if line.startswith('<!-- section:') and line.endswith('-->'):
            section_name = line.replace('<!-- section:', '').replace('-->', '').strip()
        elif line.startswith('# ') and not line.startswith('## '):
            title = line.replace('#', '').strip()
        elif line.startswith('## '):
            if is_title:
                subtitle_lines.append(line)
                has_content_started = True
            else:
                if not title:
                    title = line.replace('##', '').strip()
                else:
                    content_lines.append(line)
                    has_content_started = True
        elif line.startswith('### '):
            if is_title:
                subtitle_lines.append(line)
                has_content_started = True
            else:
                content_lines.append(line)
                has_content_started = True
        elif line.startswith('#### ') or line.startswith('##### '):
            content_lines.append(line)
            has_content_started = True
        else:
            if is_title:
                subtitle_lines.append(line)
                has_content_started = True
            else:
                content_lines.append(line)
                has_content_started = True
    
    while content_lines and content_lines[-1] == SPACER_MARKER:
        content_lines.pop()
    while subtitle_lines and subtitle_lines[-1] == SPACER_MARKER:
        subtitle_lines.pop()
    
    return {
        'is_title': is_title,
        'section_name': section_name,
        'title': title,
        'subtitle': '\n'.join(subtitle_lines),
        'content': '\n'.join(content_lines)
    }
