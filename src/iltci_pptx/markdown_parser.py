"""Markdown parsing functionality for slide content."""

import re
import logging
import yaml
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional
from .config import Config


def parse_yaml_frontmatter(content: str, delimiter: str = '---') -> Tuple[Dict[str, Any], str]:
    """Extract and parse YAML frontmatter from markdown content.
    
    Args:
        content: Full markdown content
        delimiter: Frontmatter delimiter (default '---')
        
    Returns:
        Tuple of (frontmatter_dict, remaining_content)
    """
    lines = content.split('\n')
    start_idx = -1
    end_idx = -1
    
    for i, line in enumerate(lines):
        if line.strip() == delimiter:
            if start_idx == -1:
                start_idx = i
            else:
                end_idx = i
                break
    
    if start_idx >= 0 and end_idx > start_idx:
        # Extract frontmatter YAML
        frontmatter_lines = lines[start_idx + 1:end_idx]
        frontmatter_text = '\n'.join(frontmatter_lines)
        
        try:
            frontmatter = yaml.safe_load(frontmatter_text) or {}
        except yaml.YAMLError as e:
            logging.warning(f"Failed to parse YAML frontmatter: {e}")
            frontmatter = {}
        
        # Remaining content after frontmatter
        remaining = '\n'.join(lines[end_idx + 1:])
        return frontmatter, remaining
    
    return {}, content


def parse_slide_directives(slide_content: str) -> Tuple[Dict[str, Any], str]:
    """Extract slide directives from HTML comments.
    
    Supported directives:
    - <!-- _layout: layout-name -->
    - <!-- _image_fit: cover|contain -->
    - <!-- _bg_image: path/to/image.png -->
    
    Args:
        slide_content: Content of a single slide
        
    Returns:
        Tuple of (directives_dict, content_without_directives)
    """
    directives = {}
    content = slide_content
    
    # Pattern for layout directive: <!-- _layout: name -->
    layout_match = re.search(r'<!--\s*_layout:\s*([^\s]+)\s*-->', content)
    if layout_match:
        directives['layout'] = layout_match.group(1).strip()
        content = content.replace(layout_match.group(0), '').strip()
    
    # Pattern for image fit directive: <!-- _image_fit: mode -->
    fit_match = re.search(r'<!--\s*_image_fit:\s*([^\s]+)\s*-->', content)
    if fit_match:
        directives['image_fit'] = fit_match.group(1).strip()
        content = content.replace(fit_match.group(0), '').strip()
    
    # Pattern for background image directive: <!-- _bg_image: path -->
    bg_match = re.search(r'<!--\s*_bg_image:\s*([^\s]+)\s*-->', content)
    if bg_match:
        directives['bg_image'] = bg_match.group(1).strip()
        content = content.replace(bg_match.group(0), '').strip()
    
    return directives, content


def parse_markdown_slides(md_file: Path, config: Config) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """Parse markdown file into individual slides with metadata.
    
    Args:
        md_file: Path to markdown file
        config: Configuration object
        
    Returns:
        Tuple of (frontmatter_meta, list of parsed slide dictionaries)
    """
    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    logging.info(f"Original content length: {len(content)}")
    
    # Get configuration settings
    frontmatter_delim = config.get('markdown.frontmatter_delimiter', '---')
    slide_separator = config.get('markdown.slide_separator', '---')
    title_class = config.get('markdown.title_class_marker', '<!-- _class: title -->')
    
    # Parse YAML frontmatter and get remaining content
    frontmatter, content = parse_yaml_frontmatter(content, frontmatter_delim)
    
    logging.info(f"Frontmatter keys: {list(frontmatter.keys())}")
    logging.info(f"Content after frontmatter removal length: {len(content)}")
    logging.debug(f"Content preview: {content[:200]}")
    
    # Split by slide separators
    separator_pattern = f'\n{slide_separator}\n'
    slides = re.split(separator_pattern, content)
    
    logging.info(f"Number of slides found: {len(slides)}")
    
    parsed_slides = []
    for idx, slide in enumerate(slides):
        slide = slide.strip()
        if not slide:
            logging.debug(f"Slide {idx}: Empty, skipping")
            continue
        
        logging.debug(f"--- Parsing Slide {idx} ---")
        logging.debug(f"Slide content preview: {slide[:100]}")
        
        # Check for class directive
        is_title = title_class in slide
        slide = slide.replace(title_class, '').strip()
        
        logging.debug(f"Is title slide: {is_title}")
        
        # Extract slide directives (layout, image_fit, bg_image)
        directives, slide = parse_slide_directives(slide)
        logging.debug(f"Slide directives: {directives}")
        
        # Parse slide content
        slide_data = _parse_slide_content(slide, is_title)
        
        # Merge directives into slide_data
        slide_data.update(directives)
        
        logging.debug(f"Parsed slide data: {slide_data}")
        parsed_slides.append(slide_data)
    
    logging.info(f"Total parsed slides: {len(parsed_slides)}")
    return frontmatter, parsed_slides


# Spacer marker used to represent intentional blank lines for spacing
SPACER_MARKER = '<!-- spacer -->'


def _parse_slide_content(slide: str, is_title: bool) -> Dict[str, Any]:
    """Parse individual slide content into structured data.
    
    Blank lines are converted to spacer markers to allow control over
    vertical spacing in the rendered presentation. Consecutive blank
    lines are preserved.
    
    Args:
        slide: Slide content string
        is_title: Whether this is a title slide
        
    Returns:
        Dictionary with slide data
    """
    lines = slide.split('\n')
    title = ''
    subtitle_lines = []
    content_lines = []
    section_name = ''
    has_content_started = False  # Track if we've seen non-header content
    
    for line in lines:
        line_stripped = line.strip()
        
        # Handle blank lines - convert to spacer markers after content has started
        if not line_stripped:
            if has_content_started and not is_title:
                content_lines.append(SPACER_MARKER)
                logging.debug("  -> Added spacer marker for blank line")
            elif has_content_started and is_title:
                subtitle_lines.append(SPACER_MARKER)
                logging.debug("  -> Added spacer marker for blank line (title slide)")
            continue
        
        line = line_stripped
        
        logging.debug(f"  Processing line: {line[:50]}")
        
        # Check for explicit section name marker (e.g., <!-- section: Name -->)
        if line.startswith('<!-- section:') and line.endswith('-->'):
            section_name = line.replace('<!-- section:', '').replace('-->', '').strip()
            logging.debug(f"  -> Section name: {section_name}")
        elif line.startswith('# ') and not line.startswith('## '):
            title = line.replace('#', '').strip()
            logging.debug(f"  -> Title: {title}")
        elif line.startswith('## '):
            if is_title:
                # Preserve ## marker for font size differentiation in slide builder
                subtitle_lines.append(line)
                has_content_started = True
                logging.debug(f"  -> Subtitle (##): {line}")
            else:
                # For content slides, ## can also be a title
                if not title:
                    title = line.replace('##', '').strip()
                    logging.debug(f"  -> Title (from ##): {title}")
                else:
                    # If title already exists, treat ## as content
                    content_lines.append(line)
                    has_content_started = True
                    logging.debug(f"  -> Content (##): {line}")
        elif line.startswith('### '):
            if is_title:
                # Preserve ### marker for font size differentiation in slide builder
                subtitle_lines.append(line)
                has_content_started = True
                logging.debug(f"  -> Subtitle (###): {line}")
            else:
                content_lines.append(line)
                has_content_started = True
                logging.debug(f"  -> Content (###): {line}")
        elif line.startswith('#### ') or line.startswith('##### '):
            # H4 and H5 always go to content
            content_lines.append(line)
            has_content_started = True
            logging.debug(f"  -> Content (####/####): {line}")
        else:
            # For title slides, plain text goes to subtitle area (e.g., author, date)
            # For content slides, plain text goes to content area
            if is_title:
                subtitle_lines.append(line)
                has_content_started = True
                logging.debug(f"  -> Subtitle (plain text): {line[:50]}")
            else:
                content_lines.append(line)
                has_content_started = True
                logging.debug(f"  -> Content: {line[:50]}")
    
    # Strip trailing spacer markers (they don't add value at the end)
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
