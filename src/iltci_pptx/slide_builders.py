"""Slide construction functionality for title and content slides."""

import re
import logging
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from typing import Dict, Any, TYPE_CHECKING
from .config import Config
from .rich_text import add_formatted_text, add_bullet, remove_bullet, add_numbering
from .html_media import has_html_content, extract_images_from_html, remove_html_tags
from .images import add_images_to_slide, add_images_for_layout, add_background_image, add_overlay_rectangle, LAYOUT_SPECS
from .markdown_parser import SPACER_MARKER
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
    subtitle_h2_font_size = config.get('fonts.title_slide.subtitle_h2', 32)
    subtitle_h3_font_size = config.get('fonts.title_slide.subtitle_h3', 24)
    subtitle_text_font_size = config.get('fonts.title_slide.subtitle_text', 20)
    
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
        subtitle_frame.word_wrap = True
        
        # Get spacer size for title slides
        spacer_size = config.get('fonts.title_slide.spacer', 8)
        
        # Parse subtitle lines and apply different font sizes based on header level
        subtitle_lines = slide_data['subtitle'].split('\n')
        first_line = True
        for line in subtitle_lines:
            line_stripped = line.strip()
            if not line_stripped:
                continue
            
            # Handle spacer markers (blank lines in markdown)
            if line_stripped == SPACER_MARKER:
                if first_line:
                    p = subtitle_frame.paragraphs[0]
                    p.text = ' '  # Use space to ensure paragraph has height
                    first_line = False
                else:
                    p = subtitle_frame.add_paragraph()
                    p.text = ' '  # Use space to ensure paragraph has height
                # Set font size small to minimize visual impact of the space character
                for run in p.runs:
                    run.font.size = Pt(spacer_size)
                p.space_before = Pt(spacer_size)
                p.space_after = Pt(0)
                logging.debug(f"  Added spacer paragraph in title slide subtitle ({spacer_size}pt)")
                continue
            
            # Determine font size based on header marker
            if line_stripped.startswith('## '):
                text = line_stripped[3:].strip()
                font_size = subtitle_h2_font_size
            elif line_stripped.startswith('### '):
                text = line_stripped[4:].strip()
                font_size = subtitle_h3_font_size
            else:
                # Plain text (author, date, etc.) - use subtitle_text size
                text = line_stripped
                font_size = subtitle_text_font_size
            
            if first_line:
                # Use existing first paragraph
                p = subtitle_frame.paragraphs[0]
                p.text = text
                first_line = False
            else:
                # Add new paragraph for subsequent lines
                p = subtitle_frame.add_paragraph()
                p.text = text
            
            # Format the paragraph
            for run in p.runs:
                run.font.size = Pt(font_size)
                if config.get('formatting.subtitle_bold', True):
                    run.font.bold = True
        
        logging.info("Added subtitle text box with hierarchical formatting")


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
    spacer_size = config.get('fonts.content_slide.spacer', 12)  # Smaller font for spacer lines
    
    # Parse content (bullets, numbered lists, etc.)
    for line in content.split('\n'):
        line_stripped = line.strip()
        if not line_stripped:
            continue
        
        # Handle spacer markers (blank lines in markdown)
        if line_stripped == SPACER_MARKER:
            p = text_frame.add_paragraph()
            # Use a space character instead of empty string to ensure the paragraph renders
            # with height. Empty paragraphs can collapse to zero height in PowerPoint.
            p.text = ' '
            remove_bullet(p)
            # Set font size small to minimize visual impact of the space character
            for run in p.runs:
                run.font.size = Pt(spacer_size)
            # Add spacing before to create vertical gap
            p.space_before = Pt(spacer_size)
            p.space_after = Pt(0)
            logging.debug(f"  Added spacer paragraph for vertical spacing ({spacer_size}pt)")
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


def build_layout_slide(prs: 'Presentation', slide_data: Dict[str, Any], 
                       config: Config, all_layouts: list, layout_map: dict) -> 'Slide':
    """Build a slide using a custom layout (image-side, content-bg, title-bg).
    
    Args:
        prs: PowerPoint presentation object
        slide_data: Dictionary with slide content including 'layout' key
        config: Configuration object
        all_layouts: List of all available slide layouts
        layout_map: Dictionary mapping layout names to indices
        
    Returns:
        Created slide object
    """
    layout_name = slide_data.get('layout', 'image-side')
    layout_spec = LAYOUT_SPECS.get(layout_name)
    
    if not layout_spec:
        logging.warning(f"Unknown layout '{layout_name}', falling back to content slide")
        return build_content_slide(prs, slide_data, config, all_layouts)
    
    # Use the 'Title and Content' layout as base (index 1, or find by name)
    base_layout_idx = config.get('layouts.content_slide_index', 1)
    if 'Title and Content' in layout_map:
        base_layout_idx = layout_map['Title and Content']
    
    if base_layout_idx >= len(all_layouts):
        base_layout_idx = 0
    
    slide = prs.slides.add_slide(all_layouts[base_layout_idx])
    logging.info(f"Building {layout_name} slide using base layout {base_layout_idx}")
    
    # Extract images from content
    content = slide_data.get('content', '')
    images = []
    if has_html_content(content):
        images = extract_images_from_html(content)
        content = remove_html_tags(content)
    
    # Also check for bg_image directive
    if slide_data.get('bg_image'):
        images.insert(0, {'src': slide_data['bg_image']})
    
    # Get fit mode from directive
    fit_mode = slide_data.get('image_fit', 'contain')
    
    if layout_name == 'image-side':
        _build_image_side_slide(slide, slide_data, content, images, config, fit_mode)
    elif layout_name == 'content-bg':
        _build_content_bg_slide(slide, slide_data, content, images, config, fit_mode)
    elif layout_name == 'title-bg':
        _build_title_bg_slide(slide, slide_data, content, images, config, fit_mode)
    elif layout_name == 'dual-image-text-bottom':
        _build_dual_image_slide(slide, slide_data, content, images, config, fit_mode)
    
    return slide


def _build_image_side_slide(slide: 'Slide', slide_data: Dict[str, Any],
                            content: str, images: list, config: Config,
                            fit_mode: str) -> None:
    """Build an image-side layout slide (text left, image right).
    
    Args:
        slide: PowerPoint slide object
        slide_data: Slide data dictionary
        content: Cleaned content text (HTML removed)
        images: List of image info dictionaries
        config: Configuration object
        fit_mode: Image fit mode ('contain' or 'cover')
    """
    # Set title
    if slide.shapes.title:
        slide.shapes.title.text = slide_data.get('title', '')
        logging.info("Set title for image-side slide")
    
    # Find and populate the content placeholder with narrower body area
    content_shape = None
    for shape in slide.shapes:
        if hasattr(shape, 'text_frame') and shape != slide.shapes.title:
            if hasattr(shape, 'is_placeholder') and shape.is_placeholder:
                content_shape = shape
                break
    
    # Resize content placeholder to left 60% for image-side layout
    layout_spec = LAYOUT_SPECS['image-side']
    body_spec = layout_spec.get('body', {})
    
    if content_shape:
        # Adjust content shape size to make room for image
        content_shape.left = Inches(body_spec.get('left', 0.5))
        content_shape.top = Inches(body_spec.get('top', 1.2))
        content_shape.width = Inches(body_spec.get('width', 6.5))
        content_shape.height = Inches(body_spec.get('height', 5.5))
        
        # Populate content
        _populate_content_text_frame(content_shape.text_frame, content, slide, config)
    else:
        # Create a text box manually
        textbox = slide.shapes.add_textbox(
            Inches(body_spec.get('left', 0.5)),
            Inches(body_spec.get('top', 1.2)),
            Inches(body_spec.get('width', 6.5)),
            Inches(body_spec.get('height', 5.5))
        )
        _populate_content_text_frame(textbox.text_frame, content, slide, config)
    
    # Add images to right side
    if images:
        add_images_for_layout(slide, images, 'image-side', config, 
                              base_path=Path('.'), fit_mode=fit_mode)


def _build_content_bg_slide(slide: 'Slide', slide_data: Dict[str, Any],
                            content: str, images: list, config: Config,
                            fit_mode: str) -> None:
    """Build a content-bg layout slide (full background with overlay).
    
    Args:
        slide: PowerPoint slide object
        slide_data: Slide data dictionary
        content: Cleaned content text (HTML removed)
        images: List of image info dictionaries
        config: Configuration object
        fit_mode: Image fit mode
    """
    layout_spec = LAYOUT_SPECS['content-bg']
    
    # Add background image first (will be at back)
    if images:
        first_img = images[0]
        img_src = first_img.get('src', '')
        if img_src:
            img_path = Path('.') / img_src
            bg_spec = layout_spec.get('background', {})
            add_background_image(slide, img_path, 
                                width=bg_spec.get('width', 13.33),
                                height=bg_spec.get('height', 7.5))
    
    # Add overlay rectangle
    overlay_spec = layout_spec.get('overlay', {})
    add_overlay_rectangle(
        slide,
        left=overlay_spec.get('left', 0.5),
        top=overlay_spec.get('top', 0.5),
        width=overlay_spec.get('width', 8.0),
        height=overlay_spec.get('height', 6.5),
        fill_color=overlay_spec.get('fill_color', (255, 255, 255)),
        transparency=overlay_spec.get('transparency', 0.25)
    )
    
    # Clear existing placeholders (they may conflict with our layout)
    # Add title text box on top of overlay
    title_box = slide.shapes.add_textbox(
        Inches(0.75), Inches(0.75),
        Inches(7.5), Inches(0.8)
    )
    title_frame = title_box.text_frame
    title_frame.text = slide_data.get('title', '')
    title_frame.word_wrap = True
    for p in title_frame.paragraphs:
        for run in p.runs:
            run.font.size = Pt(config.get('fonts.content_slide.title', 32))
            run.font.bold = True
    
    # Add body text box on overlay
    body_box = slide.shapes.add_textbox(
        Inches(0.75), Inches(1.75),
        Inches(7.5), Inches(5.0)
    )
    _populate_content_text_frame(body_box.text_frame, content, slide, config)


def _build_title_bg_slide(slide: 'Slide', slide_data: Dict[str, Any],
                          content: str, images: list, config: Config,
                          fit_mode: str) -> None:
    """Build a title-bg layout slide (full background with title overlay at bottom).
    
    Args:
        slide: PowerPoint slide object
        slide_data: Slide data dictionary
        content: Cleaned content text (used as subtitle)
        images: List of image info dictionaries
        config: Configuration object
        fit_mode: Image fit mode
    """
    layout_spec = LAYOUT_SPECS['title-bg']
    
    # Add background image first (will be at back)
    if images:
        first_img = images[0]
        img_src = first_img.get('src', '')
        if img_src:
            img_path = Path('.') / img_src
            bg_spec = layout_spec.get('background', {})
            add_background_image(slide, img_path,
                                width=bg_spec.get('width', 13.33),
                                height=bg_spec.get('height', 7.5))
    
    # Add overlay strip at bottom
    overlay_spec = layout_spec.get('overlay', {})
    add_overlay_rectangle(
        slide,
        left=overlay_spec.get('left', 0),
        top=overlay_spec.get('top', 5.0),
        width=overlay_spec.get('width', 13.33),
        height=overlay_spec.get('height', 2.5),
        fill_color=overlay_spec.get('fill_color', (0, 0, 0)),
        transparency=overlay_spec.get('transparency', 0.5)
    )
    
    # Add large title text box
    title_box = slide.shapes.add_textbox(
        Inches(0.5), Inches(5.25),
        Inches(12.33), Inches(1.5)
    )
    title_frame = title_box.text_frame
    title_frame.text = slide_data.get('title', '')
    title_frame.word_wrap = True
    for p in title_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        for run in p.runs:
            run.font.size = Pt(44)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    
    # Add subtitle if there's content
    if content.strip():
        subtitle_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(6.75),
            Inches(12.33), Inches(0.5)
        )
        subtitle_frame = subtitle_box.text_frame
        # Use first line of content as subtitle
        subtitle_text = content.strip().split('\n')[0]
        # Clean up any markdown formatting
        subtitle_text = re.sub(r'^[-*#]+\s*', '', subtitle_text)
        subtitle_frame.text = subtitle_text
        subtitle_frame.word_wrap = True
        for p in subtitle_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for run in p.runs:
                run.font.size = Pt(24)
                run.font.color.rgb = RGBColor(255, 255, 255)


def _build_dual_image_slide(slide: 'Slide', slide_data: Dict[str, Any],
                             content: str, images: list, config: Config,
                             fit_mode: str) -> None:
    """Build a dual-image layout slide (two images on top, text below).
    
    Args:
        slide: PowerPoint slide object
        slide_data: Slide data dictionary
        content: Cleaned content text (HTML removed)
        images: List of image info dictionaries
        config: Configuration object
        fit_mode: Image fit mode ('contain' or 'cover')
    """
    layout_spec = LAYOUT_SPECS['dual-image-text-bottom']
    
    # Set title
    if slide.shapes.title:
        slide.shapes.title.text = slide_data.get('title', '')
        logging.info("Set title for dual-image slide")
    
    # Add images using add_images_for_layout (handles both images)
    if images:
        add_images_for_layout(slide, images, 'dual-image-text-bottom', config,
                              base_path=Path('.'), fit_mode=fit_mode)
    
    # Add body text box below images
    body_spec = layout_spec.get('body', {})
    body_box = slide.shapes.add_textbox(
        Inches(body_spec.get('left', 0.5)),
        Inches(body_spec.get('top', 5.5)),
        Inches(body_spec.get('width', 12.0)),
        Inches(body_spec.get('height', 1.75))
    )
    body_frame = body_box.text_frame
    body_frame.word_wrap = True
    
    # Populate with content if present
    if content.strip():
        _populate_content_text_frame(body_frame, content, slide, config)
    
    # Center-align all paragraphs in the body text box
    for paragraph in body_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
    
    logging.info("Built dual-image-text-bottom slide")
