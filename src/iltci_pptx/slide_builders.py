"""Slide construction functionality for title and content slides."""

import re
import logging
from pptx.util import Inches, Pt
from typing import Dict, Any, TYPE_CHECKING
from .config import Config
from .rich_text import add_formatted_text, add_bullet, remove_bullet, add_numbering
from .html_media import has_html_content, extract_images_from_html, remove_html_tags
from .images import add_images_to_slide
from pathlib import Path

if TYPE_CHECKING:
    from pptx.presentation import Presentation
    from pptx.slide import Slide


def build_title_slide(prs: 'Presentation', slide_data: Dict[str, Any], config: Config, all_layouts: list) -> 'Slide':
    """Build a title slide from slide data.
    
    Args:
        prs: PowerPoint presentation object
        slide_data: Dictionary with slide content
        config: Configuration object
        all_layouts: List of all available slide layouts
        
    Returns:
        Created slide object
    """
    layout_idx = config.get('layouts.title_slide_index', 0)
    
    if len(all_layouts) == 0:
        raise IndexError("No slide layouts available in template")
    
    slide = prs.slides.add_slide(all_layouts[layout_idx])
    logging.info(f"Using layout {layout_idx} ({all_layouts[layout_idx].name}), {len(slide.shapes)} shapes")
    
    # Debug: print all shapes
    for i, shape in enumerate(slide.shapes):
        logging.debug(f"  Shape {i}: {shape.name}, has_text_frame: {hasattr(shape, 'text_frame')}, "
                     f"is_placeholder: {hasattr(shape, 'is_placeholder') and shape.is_placeholder}")
    
    # Check if the slide has any shapes
    if len(slide.shapes) == 0:
        # No shapes available - manually add text boxes
        _add_title_slide_textboxes(slide, slide_data, config)
    else:
        # Shapes are available - try to use them
        _populate_title_slide_shapes(slide, slide_data, config)
    
    return slide


def _add_title_slide_textboxes(slide: 'Slide', slide_data: Dict[str, Any], config: Config) -> None:
    """Manually add text boxes for title slide when no shapes available."""
    logging.info("No shapes found in title slide, manually adding text boxes...")
    
    # Get positioning from config
    section_pos = {
        'left': config.get('title_slide_positions.section_name.left', 0.5),
        'top': config.get('title_slide_positions.section_name.top', 0.5),
        'width': config.get('title_slide_positions.section_name.width', 9.0),
        'height': config.get('title_slide_positions.section_name.height', 1.0)
    }
    title_pos = {
        'left': config.get('title_slide_positions.title.left', 0.5),
        'top': config.get('title_slide_positions.title.top', 1.8),
        'width': config.get('title_slide_positions.title.width', 9.0),
        'height': config.get('title_slide_positions.title.height', 1.5)
    }
    subtitle_pos = {
        'left': config.get('title_slide_positions.subtitle.left', 0.5),
        'top': config.get('title_slide_positions.subtitle.top', 3.5),
        'width': config.get('title_slide_positions.subtitle.width', 9.0),
        'height': config.get('title_slide_positions.subtitle.height', 1.0)
    }
    
    # Get font sizes from config
    section_font_size = config.get('fonts.title_slide.section_name', 40)
    title_font_size = config.get('fonts.title_slide.title', 50)
    subtitle_font_size = config.get('fonts.title_slide.subtitle', 24)
    
    # Add section name at the top
    if slide_data['section_name']:
        section_box = slide.shapes.add_textbox(
            Inches(section_pos['left']), Inches(section_pos['top']),
            Inches(section_pos['width']), Inches(section_pos['height'])
        )
        section_frame = section_box.text_frame
        section_frame.text = slide_data['section_name']
        section_frame.word_wrap = True
        # Format section name
        for paragraph in section_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(section_font_size)
                if config.get('formatting.section_bold', True):
                    run.font.bold = True
        logging.info("Added section name text box")
    
    # Add main title
    if slide_data['title']:
        title_box = slide.shapes.add_textbox(
            Inches(title_pos['left']), Inches(title_pos['top']),
            Inches(title_pos['width']), Inches(title_pos['height'])
        )
        title_frame = title_box.text_frame
        title_frame.text = slide_data['title']
        title_frame.word_wrap = True
        # Format title
        for paragraph in title_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(title_font_size)
                if config.get('formatting.title_bold', True):
                    run.font.bold = True
        logging.info("Added title text box")
    
    # Add subtitle
    if slide_data['subtitle']:
        subtitle_box = slide.shapes.add_textbox(
            Inches(subtitle_pos['left']), Inches(subtitle_pos['top']),
            Inches(subtitle_pos['width']), Inches(subtitle_pos['height'])
        )
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.text = slide_data['subtitle']
        subtitle_frame.word_wrap = True
        # Format subtitle
        for paragraph in subtitle_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(subtitle_font_size)
                if config.get('formatting.subtitle_bold', True):
                    run.font.bold = True
        logging.info("Added subtitle text box")


def _populate_title_slide_shapes(slide: 'Slide', slide_data: Dict[str, Any], config: Config) -> None:
    """Populate existing shapes on title slide."""
    # Set title - use the title placeholder if available
    if slide.shapes.title:
        slide.shapes.title.text = slide_data['title']
        logging.info("Set title in title placeholder")
    
    # Find and set section name and subtitle in other text placeholders
    section_name_set = False
    subtitle_set = False
    
    for shape in slide.shapes:
        if not hasattr(shape, 'text_frame'):
            continue
        
        # Skip the title placeholder (already set)
        if shape == slide.shapes.title:
            continue
        
        # Try to identify shapes by their placeholder type or position
        if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
            placeholder_type = shape.placeholder_format.type
            logging.debug(f"  Found placeholder: type={placeholder_type}, name={shape.name}")
            
            # Try to set section name in a subtitle or body placeholder
            if not section_name_set and slide_data['section_name']:
                shape.text = slide_data['section_name']
                section_name_set = True
                logging.info("Set section name in placeholder")
            elif not subtitle_set and slide_data['subtitle']:
                shape.text = slide_data['subtitle']
                subtitle_set = True
                logging.info("Set subtitle in placeholder")
        else:
            # For non-placeholder shapes, try to set content if not already set
            if not section_name_set and slide_data['section_name']:
                shape.text = slide_data['section_name']
                section_name_set = True
                logging.info("Set section name in non-placeholder shape")
            elif not subtitle_set and slide_data['subtitle']:
                shape.text = slide_data['subtitle']
                subtitle_set = True
                logging.info("Set subtitle in non-placeholder shape")


def build_content_slide(prs: 'Presentation', slide_data: Dict[str, Any], config: Config, all_layouts: list) -> 'Slide':
    """Build a content slide from slide data.
    
    Args:
        prs: PowerPoint presentation object
        slide_data: Dictionary with slide content
        config: Configuration object
        all_layouts: List of all available slide layouts
        
    Returns:
        Created slide object
    """
    layout_idx = config.get('layouts.content_slide_index', 1)
    
    # Check if layout exists
    if len(all_layouts) < 2 and layout_idx > 0:
        logging.warning(f"Only {len(all_layouts)} layout(s) available. Using layout 0 for content slide")
        layout_idx = 0
    
    slide = prs.slides.add_slide(all_layouts[layout_idx])
    logging.info(f"Using layout {layout_idx} ({all_layouts[layout_idx].name}), {len(slide.shapes)} shapes")
    
    # Debug: print all shapes
    for i, shape in enumerate(slide.shapes):
        logging.debug(f"  Shape {i}: {shape.name}, has_text_frame: {hasattr(shape, 'text_frame')}")
    
    # Set title
    if slide.shapes.title:
        slide.shapes.title.text = slide_data['title']
        logging.info("Set title")
    
    # Set content - find the content placeholder
    content_shape = None
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame') and shape != slide.shapes.title:
            if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                content_shape = shape
                break
    
    if content_shape and hasattr(content_shape, 'text_frame'):
        _populate_content_text_frame(content_shape.text_frame, slide_data['content'], slide, config)
    
    return slide


def _populate_content_text_frame(text_frame, content: str, slide: 'Slide', config: Config) -> None:
    """Populate a text frame with parsed content.
    
    Args:
        text_frame: PowerPoint text frame object
        content: Content string to parse and add
        slide: Slide object (for adding images)
        config: Configuration object
    """
    text_frame.clear()
    logging.info("Adding content to text frame...")
    
    # Check if content has HTML with images
    images = []
    if has_html_content(content):
        logging.info("Detected HTML content, extracting images...")
        images = extract_images_from_html(content)
        logging.info(f"Found {len(images)} images")
        # Remove HTML from content
        content = remove_html_tags(content)
    
    # Get font sizes from config
    h2_size = config.get('fonts.content_slide.h2_header', 32)
    h3_size = config.get('fonts.content_slide.h3_header', 24)
    h4_size = config.get('fonts.content_slide.h4_header', 20)
    h5_size = config.get('fonts.content_slide.h5_header', 18)
    body_size = config.get('fonts.content_slide.body_text', 24)
    bullet_size = config.get('fonts.content_slide.bullet', 24)
    numbered_size = config.get('fonts.content_slide.numbered', 24)
    numbering_type = config.get('bullets.numbering_type', 'arabicPeriod')
    
    # Parse content (bullets, numbered lists, etc.)
    for line in content.split('\n'):
        line_stripped = line.strip()
        if not line_stripped:
            continue
        
        # Handle headers (h2-h5) with different sizes
        if line_stripped.startswith('##### '):
            # H5 header
            p = text_frame.add_paragraph()
            add_formatted_text(p, line_stripped[6:])
            p.level = 0
            remove_bullet(p)
            for run in p.runs:
                run.font.size = Pt(h5_size)
            logging.debug(f"  Added h5 header: {line_stripped[6:]}")
        elif line_stripped.startswith('#### '):
            # H4 header
            p = text_frame.add_paragraph()
            add_formatted_text(p, line_stripped[5:])
            p.level = 0
            remove_bullet(p)
            for run in p.runs:
                run.font.size = Pt(h4_size)
            logging.debug(f"  Added h4 header: {line_stripped[5:]}")
        elif line_stripped.startswith('### '):
            # H3 header
            p = text_frame.add_paragraph()
            add_formatted_text(p, line_stripped[4:])
            p.level = 0
            remove_bullet(p)
            for run in p.runs:
                run.font.size = Pt(h3_size)
            logging.debug(f"  Added h3 header: {line_stripped[4:]}")
        elif line_stripped.startswith('## '):
            # H2 header
            p = text_frame.add_paragraph()
            add_formatted_text(p, line_stripped[3:])
            p.level = 0
            remove_bullet(p)
            for run in p.runs:
                run.font.size = Pt(h2_size)
            logging.debug(f"  Added h2 header: {line_stripped[3:]}")
        # Handle bullet points
        elif line_stripped.startswith('- '):
            p = text_frame.add_paragraph()
            add_formatted_text(p, line_stripped[2:])
            p.level = 0
            # Explicitly add bullet formatting
            add_bullet(p, level=0)
            # Set font size for bullet text
            for run in p.runs:
                run.font.size = Pt(bullet_size)
            logging.debug(f"  Added bullet: {line_stripped[2:]}")
        elif line_stripped.startswith('  - '):
            p = text_frame.add_paragraph()
            add_formatted_text(p, line_stripped[4:])
            p.level = 1
            # Explicitly add bullet formatting for sub-bullets
            add_bullet(p, level=1)
            # Set font size for sub-bullet text
            for run in p.runs:
                run.font.size = Pt(bullet_size)
            logging.debug(f"  Added sub-bullet: {line_stripped[4:]}")
        # Handle numbered lists (e.g., "1. ", "2. ")
        elif re.match(r'^\d+\.\s+', line_stripped):
            p = text_frame.add_paragraph()
            # Extract the number and text
            match = re.match(r'^(\d+)\.\s+(.*)$', line_stripped)
            if match:
                num = int(match.group(1))
                text = match.group(2)
            else:
                num = 1
                text = re.sub(r'^\d+\.\s+', '', line_stripped)
            add_formatted_text(p, text)
            p.level = 0
            # Add automatic numbering
            add_numbering(p, start_at=num, numbering_type=numbering_type)
            # Set font size for numbered text
            for run in p.runs:
                run.font.size = Pt(numbered_size)
            logging.debug(f"  Added numbered item {num}: {text}")
        else:
            p = text_frame.add_paragraph()
            add_formatted_text(p, line_stripped)
            # Turn off bullets for regular text
            remove_bullet(p)
            # Set font size for regular text
            for run in p.runs:
                run.font.size = Pt(body_size)
            logging.debug(f"  Added text: {line_stripped}")
    
    # Add images if any were found
    if images:
        logging.info(f"Adding {len(images)} images to slide...")
        add_images_to_slide(slide, images, config, base_path=Path('.'))
