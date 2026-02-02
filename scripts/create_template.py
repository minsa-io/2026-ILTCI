#!/usr/bin/env python3
"""
Create templates/template.potx with 1 slide master and 4 layouts:
- "Title": TITLE and SUBTITLE placeholders
- "Text": TITLE and BODY placeholders  
- "Image Right": TITLE, BODY (left side) - no picture placeholder
- "Dual Image": TITLE, BODY (bottom) - no picture placeholders

Images are added programmatically via add_images_for_layout() based on layout_name.
"""

from pptx import Presentation
from pptx.util import Inches, Emu, Pt
from pptx.enum.shapes import MSO_SHAPE, PP_PLACEHOLDER
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree
import os

# Widescreen dimensions
SLIDE_WIDTH_IN = 13.33
SLIDE_HEIGHT_IN = 7.5


def create_template():
    """Create a new template with 1 master and 4 layouts."""
    
    # Create a new presentation with widescreen dimensions
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)
    
    # Get the default master
    master = prs.slide_masters[0]
    
    print(f"Initial layouts: {len(master.slide_layouts)}")
    for i, layout in enumerate(master.slide_layouts):
        print(f"  Layout {i}: {layout.name}")
    
    # Remove extra layouts - we need exactly 4
    # Access the slide layout ID list in the master and remove extras
    while len(master.slide_layouts) > 4:
        layout_id_lst = master.part._element.find(qn('p:sldLayoutIdLst'))
        if layout_id_lst is not None and len(layout_id_lst) > 0:
            # Get the rId of the last layout
            last_layout_id = layout_id_lst[-1]
            rId = last_layout_id.get(qn('r:id'))
            
            # Remove from the layout ID list
            layout_id_lst.remove(last_layout_id)
            
            # Also need to remove the relationship and the layout part
            # This is more complex - for now just remove from the list
    
    # Configure the 4 layouts
    layout_configs = [
        ("Title", configure_title_layout),
        ("Text", configure_text_layout),
        ("Image Right", configure_image_right_layout),
        ("Dual Image", configure_dual_image_layout),
    ]
    
    for idx, (name, config_func) in enumerate(layout_configs):
        if idx < len(master.slide_layouts):
            layout = master.slide_layouts[idx]
            layout.name = name
            print(f"Configuring layout {idx}: {name}")
            
            # Clear existing placeholders
            clear_placeholders(layout)
            
            # Configure the layout
            config_func(layout)
    
    # Save the template
    output_path = "templates/template.potx"
    os.makedirs("templates", exist_ok=True)
    prs.save(output_path)
    print(f"\nTemplate saved to: {output_path}")
    
    return output_path


def clear_placeholders(layout):
    """Remove all placeholder shapes from layout."""
    shapes_to_remove = []
    for shape in layout.shapes:
        if shape.is_placeholder:
            shapes_to_remove.append(shape._element)
    
    for sp in shapes_to_remove:
        sp.getparent().remove(sp)


def add_placeholder_shape(layout, ph_type, left, top, width, height, idx=None):
    """
    Add a placeholder shape to the layout using low-level XML.
    """
    # Get the spTree element
    spTree = layout.shapes._spTree
    
    # ph type mapping
    ph_type_map = {
        PP_PLACEHOLDER.TITLE: ('title', 1),
        PP_PLACEHOLDER.BODY: ('body', 2),
        PP_PLACEHOLDER.CENTER_TITLE: ('ctrTitle', 3),
        PP_PLACEHOLDER.SUBTITLE: ('subTitle', 4),
    }
    
    ph_info = ph_type_map.get(ph_type, ('body', 2))
    ph_type_str = ph_info[0]
    
    # Generate unique ID
    existing_ids = [int(sp.get('id')) for sp in spTree.findall(qn('p:sp')) if sp.get('id')]
    new_id = max(existing_ids) + 1 if existing_ids else 2
    
    # Create shape XML with proper namespaces
    NSMAP = {
        'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }
    
    # Build the XML element
    sp_elem = etree.Element(qn('p:sp'), nsmap=NSMAP)
    
    # nvSpPr
    nvSpPr = etree.SubElement(sp_elem, qn('p:nvSpPr'))
    cNvPr = etree.SubElement(nvSpPr, qn('p:cNvPr'))
    cNvPr.set('id', str(new_id))
    cNvPr.set('name', f'{ph_type_str} {idx if idx else 1}')
    
    cNvSpPr = etree.SubElement(nvSpPr, qn('p:cNvSpPr'))
    spLocks = etree.SubElement(cNvSpPr, qn('a:spLocks'))
    spLocks.set('noGrp', '1')
    
    nvPr = etree.SubElement(nvSpPr, qn('p:nvPr'))
    ph = etree.SubElement(nvPr, qn('p:ph'))
    ph.set('type', ph_type_str)
    if idx is not None:
        ph.set('idx', str(idx))
    
    # spPr
    spPr = etree.SubElement(sp_elem, qn('p:spPr'))
    xfrm = etree.SubElement(spPr, qn('a:xfrm'))
    
    off = etree.SubElement(xfrm, qn('a:off'))
    off.set('x', str(int(left * 914400)))
    off.set('y', str(int(top * 914400)))
    
    ext = etree.SubElement(xfrm, qn('a:ext'))
    ext.set('cx', str(int(width * 914400)))
    ext.set('cy', str(int(height * 914400)))
    
    prstGeom = etree.SubElement(spPr, qn('a:prstGeom'))
    prstGeom.set('prst', 'rect')
    avLst = etree.SubElement(prstGeom, qn('a:avLst'))
    
    # txBody
    txBody = etree.SubElement(sp_elem, qn('p:txBody'))
    bodyPr = etree.SubElement(txBody, qn('a:bodyPr'))
    lstStyle = etree.SubElement(txBody, qn('a:lstStyle'))
    p_elem = etree.SubElement(txBody, qn('a:p'))
    endParaRPr = etree.SubElement(p_elem, qn('a:endParaRPr'))
    endParaRPr.set('lang', 'en-US')
    
    # Append to spTree
    spTree.append(sp_elem)
    
    return sp_elem


def configure_title_layout(layout):
    """Configure 'Title' layout with TITLE and SUBTITLE placeholders."""
    # Title placeholder - centered
    add_placeholder_shape(layout, PP_PLACEHOLDER.CENTER_TITLE, 
                         0.5, 2.5, 12.33, 1.0, idx=0)
    # Subtitle placeholder
    add_placeholder_shape(layout, PP_PLACEHOLDER.SUBTITLE, 
                         0.5, 3.75, 12.33, 1.0, idx=1)


def configure_text_layout(layout):
    """Configure 'Text' layout with TITLE and BODY placeholders."""
    # Title at top
    add_placeholder_shape(layout, PP_PLACEHOLDER.TITLE, 
                         0.5, 0.3, 12.33, 0.8, idx=0)
    # Body for content
    add_placeholder_shape(layout, PP_PLACEHOLDER.BODY, 
                         0.5, 1.3, 12.33, 5.5, idx=1)


def configure_image_right_layout(layout):
    """Configure 'Image Right' layout - TITLE, BODY on left side."""
    # Title at top
    add_placeholder_shape(layout, PP_PLACEHOLDER.TITLE, 
                         0.5, 0.2, 12.33, 0.8, idx=0)
    # Body on left side (images added programmatically on right)
    add_placeholder_shape(layout, PP_PLACEHOLDER.BODY, 
                         0.5, 1.2, 6.5, 5.5, idx=1)


def configure_dual_image_layout(layout):
    """Configure 'Dual Image' layout - TITLE, BODY at bottom."""
    # Title at top
    add_placeholder_shape(layout, PP_PLACEHOLDER.TITLE, 
                         0.5, 0.2, 12.0, 0.8, idx=0)
    # Body at bottom (images added programmatically at top)
    add_placeholder_shape(layout, PP_PLACEHOLDER.BODY, 
                         0.5, 5.5, 12.0, 1.75, idx=1)


if __name__ == "__main__":
    create_template()
