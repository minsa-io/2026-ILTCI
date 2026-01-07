"""Image handling for PowerPoint slides."""

import re
import logging
from pathlib import Path
from pptx.util import Inches
from typing import List, Dict, Any, TYPE_CHECKING
from .config import Config

if TYPE_CHECKING:
    from pptx.slide import Slide


def add_images_to_slide(slide: 'Slide', images: List[Dict[str, Any]], config: Config, base_path: Path = Path('.')) -> None:
    """Add images to a slide based on extracted image information.
    
    Args:
        slide: PowerPoint slide object
        images: List of image dictionaries with attributes
        config: Configuration object
        base_path: Base path for resolving image paths
    """
    if not images:
        return
    
    # Get configuration settings
    default_height = config.get('image_layout.default_height', 3.0)
    default_width = config.get('image_layout.default_width', 2.5)
    gap_between = config.get('image_layout.gap_between', 0.5)
    top_position = config.get('image_layout.top_position', 4.0)
    slide_width = config.get('image_layout.slide_width', 10.0)
    pixels_to_inches_factor = config.get('image_layout.pixels_to_inches', 72)
    
    # Calculate layout for multiple images
    num_images = len(images)
    
    # Default image dimensions
    img_height = Inches(default_height)
    img_width = Inches(default_width)
    
    # Calculate spacing and positions for horizontal layout
    total_gap = Inches(gap_between) * (num_images - 1) if num_images > 1 else Inches(0)
    
    # Center the images horizontally
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
        left = start_left + (idx * (img_width + Inches(gap_between)))
        
        # Extract height from style if present
        style = img_info.get('style', '')
        height_match = re.search(r'height:\s*(\d+)px', style)
        current_img_height = img_height
        if height_match:
            height_px = int(height_match.group(1))
            current_img_height = Inches(height_px / pixels_to_inches_factor)
        
        try:
            # Add the image to the slide
            picture = slide.shapes.add_picture(
                str(img_path),
                left,
                top_pos,
                height=current_img_height
            )
            logging.info(f"Added image: {img_path} at position ({left}, {top_pos})")
        except Exception as e:
            logging.error(f"Error adding image {img_path}: {e}")
