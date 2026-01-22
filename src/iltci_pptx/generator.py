"""Main presentation generation orchestration."""

import logging
from pathlib import Path
from pptx import Presentation
from typing import List, Dict, Any, Optional
from .config import Config
from .markdown_parser import parse_markdown_slides
from .slide_builders import build_title_slide, build_content_slide, build_layout_slide


class PresentationGenerator:
    """Orchestrates the creation of PowerPoint presentations from markdown."""
    
    def __init__(self, config: Config):
        """Initialize the generator with configuration.
        
        Args:
            config: Configuration object
        """
        self.config = config
    
    def generate(self, template_override: Optional[Path] = None) -> None:
        """Generate the PowerPoint presentation from markdown content.
        
        Args:
            template_override: Optional path to override the template from config
        """
        # Validate paths exist
        self.config.validate_paths()
        
        # Parse markdown content first to get frontmatter
        content_path = self.config.content_path
        frontmatter, parsed_slides = parse_markdown_slides(content_path, self.config)
        
        # Determine template path: override > frontmatter > config
        if template_override:
            template_path = template_override
        elif 'template' in frontmatter:
            # Resolve template path from frontmatter relative to project root
            template_path = self.config.project_root / frontmatter['template']
        else:
            template_path = self.config.template_path
        
        logging.info(f"Loading template: {template_path}")
        prs = Presentation(str(template_path))
        
        # Collect all layouts from all slide masters
        all_layouts = []
        for master in prs.slide_masters:
            all_layouts.extend(master.slide_layouts)
        
        # Build layout name to index map for layout-by-name selection
        layout_map = {layout.name: idx for idx, layout in enumerate(all_layouts)}
        
        logging.info(f"Template has {len(prs.slide_masters)} slide master(s)")
        logging.info(f"Template has {len(all_layouts)} total layout(s) across all masters:")
        for i, layout in enumerate(all_layouts):
            logging.debug(f"  Layout {i}: {layout.name}")
        logging.info(f"Template has {len(prs.slides)} existing slides")
        
        # Remove existing content slides (keep master/layouts)
        logging.info("Removing existing slides...")
        while len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[0]
        
        logging.info(f"Creating {len(parsed_slides)} new slides...")
        
        # Create slides
        for idx, slide_data in enumerate(parsed_slides):
            logging.info(f"\nCreating slide {idx + 1}...")
            logging.info(f"  Title: {slide_data['title']}")
            logging.info(f"  Is title slide: {slide_data['is_title']}")
            
            # Check for custom layout directive
            layout_name = slide_data.get('layout')
            if layout_name:
                logging.info(f"  Custom layout: {layout_name}")
            
            # Determine which builder to use
            if layout_name in ('image-side', 'content-bg', 'title-bg', 'dual-image-text-bottom'):
                # Use new layout-aware builder
                build_layout_slide(prs, slide_data, self.config, all_layouts, layout_map)
            elif slide_data['is_title']:
                build_title_slide(prs, slide_data, self.config, all_layouts)
            else:
                build_content_slide(prs, slide_data, self.config, all_layouts)
        
        # Save presentation
        output_path = self.config.output_path
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
        logging.info(f"\nSaving presentation to {output_path}...")
        prs.save(str(output_path))
        logging.info("✓ Presentation saved successfully!")
        logging.info(f"  Total slides created: {len(prs.slides)}")
