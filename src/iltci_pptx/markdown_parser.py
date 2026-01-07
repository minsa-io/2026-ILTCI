"""Markdown parsing functionality for slide content."""

import re
import logging
from pathlib import Path
from typing import List, Dict, Any
from .config import Config


def parse_markdown_slides(md_file: Path, config: Config) -> List[Dict[str, Any]]:
    """Parse markdown file into individual slides.
    
    Args:
        md_file: Path to markdown file
        config: Configuration object
        
    Returns:
        List of parsed slide dictionaries
    """
    with open(md_file, 'r', encoding='utf-8') as f:
        content = f.read()
    
    logging.info(f"Original content length: {len(content)}")
    
    # Get configuration settings
    frontmatter_delim = config.get('markdown.frontmatter_delimiter', '---')
    slide_separator = config.get('markdown.slide_separator', '---')
    title_class = config.get('markdown.title_class_marker', '<!-- _class: title -->')
    
    # Remove YAML frontmatter (everything between first --- and second ---)
    lines = content.split('\n')
    start_idx = 0
    end_idx = 0
    found_first = False
    
    for i, line in enumerate(lines):
        if line.strip() == frontmatter_delim:
            if not found_first:
                start_idx = i
                found_first = True
            else:
                end_idx = i
                break
    
    if end_idx > start_idx:
        # Remove frontmatter lines
        content_lines = lines[end_idx + 1:]
        content = '\n'.join(content_lines)
    
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
        
        # Parse slide content
        slide_data = _parse_slide_content(slide, is_title)
        
        logging.debug(f"Parsed slide data: {slide_data}")
        parsed_slides.append(slide_data)
    
    logging.info(f"Total parsed slides: {len(parsed_slides)}")
    return parsed_slides


def _parse_slide_content(slide: str, is_title: bool) -> Dict[str, Any]:
    """Parse individual slide content into structured data.
    
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
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
        
        logging.debug(f"  Processing line: {line[:50]}")
        
        if line.startswith('### ') and not section_name and is_title:
            # First h3 is section name on title slide
            section_name = line.replace('###', '').strip()
            logging.debug(f"  -> Section name: {section_name}")
        elif line.startswith('# '):
            title = line.replace('#', '').strip()
            logging.debug(f"  -> Title: {title}")
        elif line.startswith('## '):
            if is_title:
                subtitle_lines.append(line.replace('##', '').strip())
                logging.debug(f"  -> Subtitle: {line.replace('##', '').strip()}")
            else:
                # For content slides, ## can also be a title
                if not title:
                    title = line.replace('##', '').strip()
                    logging.debug(f"  -> Title (from ##): {title}")
                else:
                    # If title already exists, treat ## as content
                    content_lines.append(line)
                    logging.debug(f"  -> Content (##): {line}")
        elif line.startswith('### '):
            if is_title:
                subtitle_lines.append(line.replace('###', '').strip())
                logging.debug(f"  -> Subtitle: {line.replace('###', '').strip()}")
            else:
                content_lines.append(line)
                logging.debug(f"  -> Content (###): {line}")
        elif line.startswith('#### ') or line.startswith('##### '):
            # H4 and H5 always go to content
            content_lines.append(line)
            logging.debug(f"  -> Content (####/####): {line}")
        else:
            content_lines.append(line)
            logging.debug(f"  -> Content: {line[:50]}")
    
    return {
        'is_title': is_title,
        'section_name': section_name,
        'title': title,
        'subtitle': '\n'.join(subtitle_lines),
        'content': '\n'.join(content_lines)
    }
