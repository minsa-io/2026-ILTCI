"""Rich text formatting for PowerPoint paragraphs."""

import re
from pptx.oxml.xmlchemy import OxmlElement
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from pptx.text.text import _Paragraph


def add_bullet(paragraph: '_Paragraph', level: int = 0) -> None:
    """Add bullet formatting to a paragraph.
    
    Args:
        paragraph: PowerPoint paragraph object
        level: Indentation level (not currently used)
    """
    pPr = paragraph._element.get_or_add_pPr()
    # Create bullet element
    buChar = OxmlElement('a:buChar')
    buChar.set('char', '•')  # Bullet character
    pPr.insert(0, buChar)


def remove_bullet(paragraph: '_Paragraph') -> None:
    """Remove bullet formatting from a paragraph.
    
    Args:
        paragraph: PowerPoint paragraph object
    """
    pPr = paragraph._element.get_or_add_pPr()
    pPr.insert(0, OxmlElement('a:buNone'))


def add_numbering(paragraph: '_Paragraph', start_at: int = 1, numbering_type: str = 'arabicPeriod') -> None:
    """Add automatic numbering to a paragraph.
    
    Args:
        paragraph: PowerPoint paragraph object
        start_at: Starting number
        numbering_type: Numbering style (e.g., 'arabicPeriod' for 1. 2. 3.)
    """
    pPr = paragraph._element.get_or_add_pPr()
    # Create buAutoNum element for numbering
    buAutoNum = OxmlElement('a:buAutoNum')
    buAutoNum.set('type', numbering_type)
    if start_at > 1:
        buAutoNum.set('startAt', str(start_at))
    pPr.insert(0, buAutoNum)


def add_formatted_text(paragraph: '_Paragraph', text: str) -> None:
    """Add text to a paragraph with markdown formatting support.
    
    Supports **bold**, *italic*, ***bold+italic***, and [text](url) markdown syntax.
    
    Args:
        paragraph: PowerPoint paragraph object
        text: Text with markdown formatting
    """
    # Clear any existing text
    paragraph.text = ""
    
    # Pattern to match markdown formatting including links
    # Matches [text](url) links, ***text*** (bold+italic), **text** (bold), or *text* (italic)
    pattern = r'(\[.*?\]\(.*?\)|\*\*\*.*?\*\*\*|\*\*.*?\*\*|\*.*?\*)'
    
    # Split text by markdown patterns
    parts = re.split(pattern, text)
    
    for part in parts:
        if not part:
            continue
        
        # Check what type of formatting this part has
        if part.startswith('[') and part.endswith(')') and '](' in part:
            # Markdown link [text](url)
            # Extract the link text and URL
            link_match = re.match(r'\[(.*?)\]\((.*?)\)', part)
            if link_match:
                link_text = link_match.group(1)
                link_url = link_match.group(2)
                run = paragraph.add_run()
                run.text = link_text
                # Add hyperlink
                run.hyperlink.address = link_url
                # Add underline for visibility (PowerPoint will handle color)
                run.font.underline = True
        elif part.startswith('***') and part.endswith('***'):
            # Bold and italic
            run = paragraph.add_run()
            run.text = part[3:-3]
            run.font.bold = True
            run.font.italic = True
        elif part.startswith('**') and part.endswith('**'):
            # Bold
            run = paragraph.add_run()
            run.text = part[2:-2]
            run.font.bold = True
        elif part.startswith('*') and part.endswith('*'):
            # Italic
            run = paragraph.add_run()
            run.text = part[1:-1]
            run.font.italic = True
        else:
            # Regular text
            run = paragraph.add_run()
            run.text = part
