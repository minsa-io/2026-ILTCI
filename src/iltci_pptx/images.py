"""Image handling for PowerPoint slides."""

import re
import logging
from pathlib import Path
from pptx.util import Inches, Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from typing import List, Dict, Any, Optional, Tuple, TYPE_CHECKING
from pptx.enum.text import PP_ALIGN
from .config import Config

if TYPE_CHECKING:
    from pptx.slide import Slide
    from pptx.shapes.picture import Picture

# Default image style settings for borders and rounded corners
IMAGE_STYLE_DEFAULTS = {
    'border_width': Pt(2),           # 2pt border width
    'border_color': RGBColor(68, 84, 106),  # #44546A - dark blue-gray
    'corner_radius': Inches(0.1),    # Rounded corner radius
    'border_enabled': True,          # Enable border by default
    'rounded_enabled': True,         # Enable rounded corners by default
}

# Style class mappings for per-image overrides
STYLE_CLASS_MAP = {
    'no-border': {'border_enabled': False},
    'no-rounded': {'rounded_enabled': False},
    'border-thin': {'border_width': Pt(1)},
    'border-thick': {'border_width': Pt(4)},
    'border-light': {'border_color': RGBColor(180, 198, 231)},  # #B4C6E7
    'border-dark': {'border_color': RGBColor(47, 62, 78)},      # #2F3E4E
    'rounded-sm': {'corner_radius': Inches(0.05)},
    'rounded-lg': {'corner_radius': Inches(0.2)},
}


def parse_style_classes(class_attr: str) -> Dict[str, Any]:
    """Parse CSS class attribute and return merged style settings.
    
    Args:
        class_attr: Space-separated CSS class names from HTML img tag
        
    Returns:
        Dictionary of style settings merged from defaults and class overrides
    """
    style = IMAGE_STYLE_DEFAULTS.copy()
    
    if not class_attr:
        return style
    
    classes = class_attr.split()
    for cls in classes:
        if cls in STYLE_CLASS_MAP:
            style.update(STYLE_CLASS_MAP[cls])
    
    return style


def apply_image_style(picture: 'Picture', style: Dict[str, Any]) -> None:
    """Apply border and rounded corner styling to a PowerPoint Picture shape.
    
    Args:
        picture: PowerPoint Picture shape object
        style: Style dictionary with border_enabled, border_width, border_color,
               rounded_enabled, corner_radius keys
    """
    if picture is None:
        return
    
    try:
        # Apply border styling
        if style.get('border_enabled', True):
            line = picture.line
            line.width = style.get('border_width', Pt(2))
            line.color.rgb = style.get('border_color', RGBColor(68, 84, 106))
        else:
            # Remove border
            picture.line.fill.background()
        
        # Apply rounded corners using adjustments
        # PowerPoint uses shape adjustments (adj) for rounded rectangles
        # For pictures, we need to use XML manipulation for soft edges/rounded corners
        if style.get('rounded_enabled', True):
            corner_radius = style.get('corner_radius', Inches(0.1))
            _apply_rounded_corners(picture, corner_radius)
        
        logging.debug(f"Applied image style: border={style.get('border_enabled')}, rounded={style.get('rounded_enabled')}")
        
    except Exception as e:
        logging.warning(f"Could not apply image style: {e}")


def _apply_rounded_corners(picture: 'Picture', radius: int) -> None:
    """Apply rounded corners to a picture using XML manipulation.
    
    Args:
        picture: PowerPoint Picture shape
        radius: Corner radius in EMUs
    """
    try:
        from pptx.oxml.ns import qn
        from lxml import etree
        
        # Get the spPr (shape properties) element
        spPr = picture._pic.spPr
        
        # Check if prstGeom exists, if not create it
        prstGeom = spPr.find(qn('a:prstGeom'))
        if prstGeom is None:
            # Create preset geometry for rounded rectangle
            prstGeom = etree.SubElement(spPr, qn('a:prstGeom'))
        
        # Set to rounded rectangle preset
        prstGeom.set('prst', 'roundRect')
        
        # Add or update adjustment values for corner radius
        avLst = prstGeom.find(qn('a:avLst'))
        if avLst is None:
            avLst = etree.SubElement(prstGeom, qn('a:avLst'))
        
        # Clear existing adjustments
        for child in list(avLst):
            avLst.remove(child)
        
        # Calculate adjustment value (PowerPoint uses 0-50000 scale for roundRect)
        # Convert radius to percentage of shortest side
        # Assuming ~10% rounding as default
        adj_val = 10000  # 10% rounding
        if isinstance(radius, int):  # If EMUs provided
            # Rough conversion - smaller radius = smaller adj value
            adj_val = min(int(radius / 914400 * 100000), 50000)  # Cap at 50%
        
        gd = etree.SubElement(avLst, qn('a:gd'))
        gd.set('name', 'adj')
        gd.set('fmla', f'val {adj_val}')
        
        logging.debug(f"Applied rounded corners with adj={adj_val}")
        
    except ImportError:
        logging.warning("lxml not available, cannot apply rounded corners via XML")
    except Exception as e:
        logging.debug(f"Could not apply rounded corners: {e}")


# Layout specifications for image-aware layouts
# These match the specs in assets/layout-specs.yaml
LAYOUT_SPECS = {
    "image-side": {
        "description": "Text on left (60%), image on right (40%)",
        "picture": {"left": 7.5, "top": 1.2, "width": 5.33, "height": 5.5},
        "body": {"left": 0.5, "top": 1.2, "width": 6.5, "height": 5.5},
    },
    "content-bg": {
        "description": "Full background image with semi-transparent content overlay",
        "background": {"left": 0, "top": 0, "width": 13.33, "height": 7.5},
        "overlay": {"left": 0.5, "top": 0.5, "width": 8.0, "height": 6.5,
                   "fill_color": (255, 255, 255), "transparency": 0.25},
    },
    "title-bg": {
        "description": "Full background image with title overlay at bottom",
        "background": {"left": 0, "top": 0, "width": 13.33, "height": 7.5},
        "overlay": {"left": 0, "top": 5.0, "width": 13.33, "height": 2.5,
                   "fill_color": (0, 0, 0), "transparency": 0.5},
    },
    "dual-image-text-bottom": {
        "description": "Two side-by-side images (top ~70%), text below (bottom ~30%)",
        "picture_left": {"left": 0.75, "top": 1.2, "width": 5.5, "height": 4.0},
        "picture_right": {"left": 6.75, "top": 1.2, "width": 5.5, "height": 4.0},
        "body": {"left": 0.5, "top": 5.5, "width": 12.0, "height": 1.75},
        "title": {"left": 0.5, "top": 0.2, "width": 12.0, "height": 0.8},
    },
}


def add_background_image(slide: 'Slide', img_path: Path, 
                         width: float = 13.33, height: float = 7.5) -> None:
    """Add a full-bleed background image to slide.
    
    The image is added at position (0,0) and should be moved to back.
    
    Args:
        slide: PowerPoint slide object
        img_path: Path to the image file
        width: Slide width in inches (default: 13.33 for widescreen)
        height: Slide height in inches (default: 7.5)
    """
    if not img_path.exists():
        logging.warning(f"Background image not found: {img_path}")
        return
    
    try:
        # Add picture at full slide size
        picture = slide.shapes.add_picture(
            str(img_path),
            Inches(0),
            Inches(0),
            width=Inches(width),
            height=Inches(height)
        )
        
        # Move to back by adjusting z-order
        # In python-pptx, we need to move the shape to the beginning of shapes
        sp = picture._element
        sp.getparent().insert(0, sp)
        
        logging.info(f"Added background image: {img_path}")
    except Exception as e:
        logging.error(f"Error adding background image {img_path}: {e}")


def add_overlay_rectangle(slide: 'Slide', left: float, top: float, 
                          width: float, height: float,
                          fill_color: tuple = (255, 255, 255),
                          transparency: float = 0.25) -> None:
    """Add a semi-transparent overlay rectangle to slide.
    
    Args:
        slide: PowerPoint slide object
        left, top, width, height: Position and size in inches
        fill_color: RGB tuple (r, g, b) with values 0-255
        transparency: Transparency value 0.0 (opaque) to 1.0 (fully transparent)
    """
    try:
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(left),
            Inches(top),
            Inches(width),
            Inches(height)
        )
        
        # Set fill color
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(*fill_color)
        
        # Set transparency (python-pptx uses 0-100000 scale for transparency)
        # transparency 0.25 = 25% transparent = 75% opaque
        fill.fore_color.brightness = 0  # No brightness adjustment
        
        # Note: python-pptx doesn't directly support fill transparency easily
        # We need to use XML manipulation for proper transparency
        from pptx.oxml.ns import qn
        spPr = shape._sp.spPr
        solidFill = spPr.find(qn('a:solidFill'))
        if solidFill is not None:
            srgbClr = solidFill.find(qn('a:srgbClr'))
            if srgbClr is not None:
                # Add alpha element for transparency
                # alpha value is 0-100000 where 100000 = fully opaque
                alpha_val = int((1 - transparency) * 100000)
                alpha = srgbClr.makeelement(qn('a:alpha'), {})
                alpha.set('val', str(alpha_val))
                srgbClr.append(alpha)
        
        # Remove outline
        shape.line.fill.background()
        
        logging.info(f"Added overlay rectangle at ({left}, {top}) with {transparency*100}% transparency")
    except Exception as e:
        logging.error(f"Error adding overlay rectangle: {e}")


def add_image_to_area(slide: 'Slide', img_path: Path,
                      left: float, top: float, width: float, height: float,
                      fit_mode: str = 'contain',
                      class_attr: Optional[str] = None) -> Optional['Picture']:
    """Add an image to a specific area, scaling to fit.
    
    Args:
        slide: PowerPoint slide object
        img_path: Path to image file
        left, top, width, height: Target area in inches
        fit_mode: 'contain' (fit within, preserve aspect) or 'cover' (fill, may crop)
        class_attr: CSS class attribute for style overrides (e.g., 'no-border rounded-lg')
        
    Returns:
        Picture shape or None if failed
    """
    if not img_path.exists():
        logging.warning(f"Image not found: {img_path}")
        return None
    
    # Parse CSS classes to get image style settings
    image_style = parse_style_classes(class_attr or '')
    
    try:
        from PIL import Image
        
        # Get original image dimensions
        with Image.open(img_path) as img:
            orig_width, orig_height = img.size
        
        # Calculate aspect ratios
        orig_ratio = orig_width / orig_height
        target_ratio = width / height
        
        if fit_mode == 'contain':
            # Scale to fit within bounds, preserving aspect ratio
            if orig_ratio > target_ratio:
                # Image is wider - constrain by width
                final_width = width
                final_height = width / orig_ratio
            else:
                # Image is taller - constrain by height
                final_height = height
                final_width = height * orig_ratio
            
            # Center within target area
            final_left = left + (width - final_width) / 2
            final_top = top + (height - final_height) / 2
        else:
            # Cover mode - fill area (may exceed bounds)
            if orig_ratio > target_ratio:
                # Image is wider - constrain by height, center horizontally
                final_height = height
                final_width = height * orig_ratio
                final_top = top
                final_left = left + (width - final_width) / 2
            else:
                # Image is taller - constrain by width, center vertically
                final_width = width
                final_height = width / orig_ratio
                final_left = left
                final_top = top + (height - final_height) / 2
        
        picture = slide.shapes.add_picture(
            str(img_path),
            Inches(final_left),
            Inches(final_top),
            width=Inches(final_width),
            height=Inches(final_height)
        )
        
        # Apply image styling (border and rounded corners)
        apply_image_style(picture, image_style)
        
        logging.info(f"Added image {img_path} to area ({left}, {top}) {width}x{height} [mode={fit_mode}]")
        return picture
        
    except ImportError:
        # PIL not available, use simple placement
        logging.warning("PIL not available, using simple image placement")
        picture = slide.shapes.add_picture(
            str(img_path),
            Inches(left),
            Inches(top),
            width=Inches(width)
        )
        
        # Apply image styling (border and rounded corners)
        apply_image_style(picture, image_style)
        return picture
        
    except Exception as e:
        logging.error(f"Error adding image to area: {e}")
        return None


def add_images_for_layout(slide: 'Slide', images: List[Dict[str, Any]], 
                          layout_name: str, config: Config,
                          base_path: Path = Path('.'),
                          fit_mode: str = 'contain') -> None:
    """Add images to slide based on layout type.
    
    Args:
        slide: PowerPoint slide object
        images: List of image dictionaries with 'src' key
        layout_name: Layout type ('image-side', 'content-bg', 'title-bg')
        config: Configuration object
        base_path: Base path for resolving image paths
        fit_mode: 'contain' or 'cover' for image fitting
    """
    if not images:
        return
    
    layout_spec = LAYOUT_SPECS.get(layout_name)
    if not layout_spec:
        logging.warning(f"Unknown layout: {layout_name}, using fallback placement")
        add_images_to_slide(slide, images, config, base_path)
        return
    
    # Get first image path
    first_img = images[0]
    img_src = first_img.get('src', '')
    if not img_src:
        return
    
    img_path = base_path / img_src
    
    if layout_name in ('content-bg', 'title-bg'):
        # Background image layouts
        bg_spec = layout_spec.get('background', {})
        add_background_image(
            slide, img_path,
            width=bg_spec.get('width', 13.33),
            height=bg_spec.get('height', 7.5)
        )
        
        # Add overlay
        overlay_spec = layout_spec.get('overlay', {})
        if overlay_spec:
            add_overlay_rectangle(
                slide,
                left=overlay_spec.get('left', 0),
                top=overlay_spec.get('top', 0),
                width=overlay_spec.get('width', 13.33),
                height=overlay_spec.get('height', 7.5),
                fill_color=overlay_spec.get('fill_color', (255, 255, 255)),
                transparency=overlay_spec.get('transparency', 0.25)
            )
    
    elif layout_name == 'image-side':
        # Side image layout with optional caption
        pic_spec = layout_spec.get('picture', {})
        
        # Get caption and class from data attributes
        first_caption = first_img.get('data-caption', '')
        first_class = first_img.get('class', '')
        
        add_image_with_caption(
            slide, img_path,
            left=pic_spec.get('left', 7.5),
            top=pic_spec.get('top', 1.2),
            width=pic_spec.get('width', 5.33),
            height=pic_spec.get('height', 5.5),
            caption=first_caption,
            fit_mode=fit_mode,
            class_attr=first_class
        )
        
        # Handle additional images (fallback placement below main)
        if len(images) > 1:
            logging.info(f"Layout {layout_name} has {len(images)} images, placing extras below")
            # Place remaining images using fallback
            add_images_to_slide(slide, images[1:], config, base_path)
    
    elif layout_name == 'dual-image-text-bottom':
        # Two side-by-side images at top, with optional captions
        pic_left_spec = layout_spec.get('picture_left', {})
        
        # Get caption and class from data attributes
        first_caption = first_img.get('data-caption', '')
        first_class = first_img.get('class', '')
        
        add_image_with_caption(
            slide, img_path,
            left=pic_left_spec.get('left', 0.75),
            top=pic_left_spec.get('top', 1.2),
            width=pic_left_spec.get('width', 5.5),
            height=pic_left_spec.get('height', 4.0),
            caption=first_caption,
            fit_mode=fit_mode,
            class_attr=first_class
        )
        
        # Add second image if present
        if len(images) > 1:
            second_img = images[1]
            img_src2 = second_img.get('src', '')
            if img_src2:
                img_path2 = base_path / img_src2
                pic_right_spec = layout_spec.get('picture_right', {})
                second_caption = second_img.get('data-caption', '')
                second_class = second_img.get('class', '')
                add_image_with_caption(
                    slide, img_path2,
                    left=pic_right_spec.get('left', 6.75),
                    top=pic_right_spec.get('top', 1.2),
                    width=pic_right_spec.get('width', 5.5),
                    height=pic_right_spec.get('height', 4.0),
                    caption=second_caption,
                    fit_mode=fit_mode,
                    class_attr=second_class
                )


def add_images_to_slide(slide: 'Slide', images: List[Dict[str, Any]], 
                        config: Config, base_path: Path = Path('.'),
                        layout_name: Optional[str] = None,
                        fit_mode: str = 'contain') -> None:
    """Add images to a slide based on extracted image information.
    
    Args:
        slide: PowerPoint slide object
        images: List of image dictionaries with attributes
        config: Configuration object
        base_path: Base path for resolving image paths
        layout_name: Optional layout name for layout-aware placement
        fit_mode: 'contain' or 'cover' for image fitting
    """
    if not images:
        return
    
    # If layout specified, use layout-aware placement
    if layout_name and layout_name in LAYOUT_SPECS:
        add_images_for_layout(slide, images, layout_name, config, base_path, fit_mode)
        return
    
    # Fallback: Position-based placement with improved positioning
    default_height = config.get('image_layout.default_height', 3.5)  # Taller images
    default_width = config.get('image_layout.default_width', 3.0)    # Wider images  
    gap_between = config.get('image_layout.gap_between', 0.5)
    top_position = config.get('image_layout.top_position', 2.0)  # Higher position (was 4.0)
    slide_width = config.get('image_layout.slide_width', 13.33)  # Widescreen
    pixels_to_inches_factor = config.get('image_layout.pixels_to_inches', 72)
    
    # Calculate layout for multiple images
    num_images = len(images)
    
    # Adjust dimensions based on number of images
    if num_images == 1:
        # Single image: place on right side like image-side layout
        img_height = Inches(5.0)
        img_width = Inches(5.0)
        start_left = Inches(7.5)
        top_pos = Inches(1.5)
        total_gap = Inches(0)
    elif num_images == 2:
        # Two images: side by side, centered, higher up
        img_height = Inches(default_height)
        img_width = Inches(default_width)
        total_gap = Inches(gap_between)
        total_width = (img_width * num_images) + total_gap
        start_left = (Inches(slide_width) - total_width) / 2
        top_pos = Inches(top_position)
    else:
        # More images: original layout
        img_height = Inches(default_height)
        img_width = Inches(default_width)
        total_gap = Inches(gap_between) * (num_images - 1) if num_images > 1 else Inches(0)
        total_width = (img_width * num_images) + total_gap
        start_left = (Inches(slide_width) - total_width) / 2
        top_pos = Inches(top_position)
    
    for idx, img_info in enumerate(images):
        img_src = img_info.get('src', '')
        if not img_src:
            continue
        
        # Construct full path
        img_path = base_path / img_src
        
        if not img_path.exists():
            logging.warning(f"Image not found: {img_path}")
            continue
        
        # Calculate position
        if num_images == 1:
            left = start_left
        else:
            left = start_left + (idx * (img_width + Inches(gap_between)))
        
        # Extract height from style if present
        style = img_info.get('style', '')
        height_match = re.search(r'height:\s*(\d+)px', style)
        current_img_height = img_height
        if height_match:
            height_px = int(height_match.group(1))
            current_img_height = Inches(height_px / pixels_to_inches_factor)
        
        # Get CSS class attribute for styling
        class_attr = img_info.get('class', '')
        image_style = parse_style_classes(class_attr)
        
        try:
            # Add the image to the slide
            picture = slide.shapes.add_picture(
                str(img_path),
                left,
                top_pos,
                height=current_img_height
            )
            
            # Apply image styling (border and rounded corners)
            apply_image_style(picture, image_style)
            
            logging.info(f"Added image: {img_path} at position ({left}, {top_pos})")
        except Exception as e:
            logging.error(f"Error adding image {img_path}: {e}")


# Default caption style settings
CAPTION_STYLE = {
    'color': RGBColor(51, 51, 51),  # #333333 - off-black
    'font_size': 12,  # 12pt
    'align': PP_ALIGN.LEFT,
}


def add_image_caption(slide: 'Slide', caption: str, 
                      left: float, top: float, width: float,
                      style: Optional[Dict[str, Any]] = None) -> None:
    """Add a caption text box below an image.
    
    Args:
        slide: PowerPoint slide object
        caption: Caption text to display
        left: Left position in inches (should match image left)
        top: Top position in inches (should be below image bottom)
        width: Width in inches (should match image width)
        style: Optional style overrides {color, font_size, align}
    """
    if not caption:
        return
    
    # Merge default style with any overrides
    caption_style = CAPTION_STYLE.copy()
    if style:
        if 'color' in style:
            caption_style['color'] = style['color']
        if 'font_size' in style:
            caption_style['font_size'] = style['font_size']
        if 'align' in style:
            caption_style['align'] = style['align']
    
    try:
        # Create caption text box
        caption_box = slide.shapes.add_textbox(
            Inches(left),
            Inches(top),
            Inches(width),
            Inches(0.4)  # Height for ~1-2 lines of caption
        )
        
        caption_frame = caption_box.text_frame
        caption_frame.word_wrap = True
        
        # Add caption text (interpret \n as newlines)
        caption_text = caption.replace('\\n', '\n')
        p = caption_frame.paragraphs[0]
        p.text = caption_text
        p.alignment = caption_style['align']
        
        # Style the text
        for run in p.runs:
            run.font.size = Pt(caption_style['font_size'])
            run.font.color.rgb = caption_style['color']
        
        # Remove any fill/background on the textbox
        caption_box.fill.background()
        
        logging.info(f"Added caption: '{caption}' at ({left}, {top})")
        
    except Exception as e:
        logging.error(f"Error adding caption: {e}")


def add_image_with_caption(slide: 'Slide', img_path: Path,
                           left: float, top: float, width: float, height: float,
                           caption: Optional[str] = None,
                           fit_mode: str = 'contain',
                           caption_style: Optional[Dict[str, Any]] = None,
                           class_attr: Optional[str] = None) -> Tuple[Optional['Picture'], float]:
    """Add an image to a specific area with optional caption below.
    
    Args:
        slide: PowerPoint slide object
        img_path: Path to image file
        left, top, width, height: Target area in inches
        caption: Optional caption text to display below image
        fit_mode: 'contain' (fit within, preserve aspect) or 'cover' (fill, may crop)
        caption_style: Optional style overrides for caption
        class_attr: CSS class attribute for style overrides (e.g., 'no-border rounded-lg')
        
    Returns:
        Tuple of (picture shape or None, actual bottom position of image in inches)
    """
    if not img_path.exists():
        logging.warning(f"Image not found: {img_path}")
        return None, top + height
    
    # Parse CSS classes to get image style settings
    image_style = parse_style_classes(class_attr or '')
    
    try:
        from PIL import Image
        
        # Get original image dimensions
        with Image.open(img_path) as img:
            orig_width, orig_height = img.size
        
        # Calculate aspect ratios
        orig_ratio = orig_width / orig_height
        target_ratio = width / height
        
        if fit_mode == 'contain':
            # Scale to fit within bounds, preserving aspect ratio
            if orig_ratio > target_ratio:
                # Image is wider - constrain by width
                final_width = width
                final_height = width / orig_ratio
            else:
                # Image is taller - constrain by height
                final_height = height
                final_width = height * orig_ratio
            
            # Center within target area
            final_left = left + (width - final_width) / 2
            final_top = top + (height - final_height) / 2
        else:
            # Cover mode - fill area (may exceed bounds)
            if orig_ratio > target_ratio:
                final_height = height
                final_width = height * orig_ratio
                final_top = top
                final_left = left + (width - final_width) / 2
            else:
                final_width = width
                final_height = width / orig_ratio
                final_left = left
                final_top = top + (height - final_height) / 2
        
        picture = slide.shapes.add_picture(
            str(img_path),
            Inches(final_left),
            Inches(final_top),
            width=Inches(final_width),
            height=Inches(final_height)
        )
        
        # Apply image styling (border and rounded corners)
        apply_image_style(picture, image_style)
        
        # Calculate actual bottom of image
        actual_bottom = final_top + final_height
        
        logging.info(f"Added image {img_path} to area ({left}, {top}) {width}x{height} [mode={fit_mode}]")
        
        # Add caption if provided
        if caption:
            # Position caption just below the image, aligned with image left edge
            caption_top = actual_bottom + 0.05  # Small gap below image
            add_image_caption(slide, caption, final_left, caption_top, final_width, caption_style)
        
        return picture, actual_bottom
        
    except ImportError:
        # PIL not available, use simple placement
        logging.warning("PIL not available, using simple image placement")
        picture = slide.shapes.add_picture(
            str(img_path),
            Inches(left),
            Inches(top),
            width=Inches(width)
        )
        
        # Apply image styling (border and rounded corners)
        apply_image_style(picture, image_style)
        
        actual_bottom = top + height
        
        if caption:
            caption_top = actual_bottom + 0.05
            add_image_caption(slide, caption, left, caption_top, width, caption_style)
        
        return picture, actual_bottom
        
    except Exception as e:
        logging.error(f"Error adding image to area: {e}")
        return None, top + height
