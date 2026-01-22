"""HTML content processing and image extraction."""

import re
from html.parser import HTMLParser
from typing import List, Dict, Any


class ImageExtractor(HTMLParser):
    """Extract image tags and their attributes from HTML."""
    
    def __init__(self):
        super().__init__()
        self.images = []
    
    def handle_starttag(self, tag, attrs):
        if tag == 'img':
            attr_dict = dict(attrs)
            self.images.append(attr_dict)


def extract_images_from_html(html_content: str) -> List[Dict[str, Any]]:
    """Extract image information from HTML content.
    
    Args:
        html_content: HTML string to parse
        
    Returns:
        List of dictionaries containing image attributes
    """
    parser = ImageExtractor()
    parser.feed(html_content)
    return parser.images


def has_html_content(text: str) -> bool:
    """Check if text contains HTML tags.
    
    Args:
        text: Text to check
        
    Returns:
        True if HTML tags are found
    """
    return bool(re.search(r'<[^>]+>', text))


def remove_html_tags(text: str, preserve_markers: list = None) -> str:
    """Remove HTML tags from text but preserve the content structure.
    
    Args:
        text: Text with HTML tags
        preserve_markers: List of markers to preserve (default: ['<!-- spacer -->'])
        
    Returns:
        Text with HTML tags removed
    """
    if preserve_markers is None:
        preserve_markers = ['<!-- spacer -->']
    
    # Temporarily replace markers we want to preserve
    preserved = {}
    for i, marker in enumerate(preserve_markers):
        placeholder = f'__PRESERVED_MARKER_{i}__'
        preserved[placeholder] = marker
        text = text.replace(marker, placeholder)
    
    # Remove complete HTML blocks that we've processed (like img tags in divs)
    text = re.sub(r'<div[^>]*>.*?</div>', '', text, flags=re.DOTALL)
    # Remove any remaining standalone tags
    text = re.sub(r'<[^>]+>', '', text)
    
    # Restore preserved markers
    for placeholder, marker in preserved.items():
        text = text.replace(placeholder, marker)
    
    return text.strip()
