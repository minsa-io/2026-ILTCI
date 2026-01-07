"""Command-line interface for ILTCI presentation generator."""

import argparse
import sys
import logging
from pathlib import Path
from .config import Config
from .generator import PresentationGenerator


def parse_arguments() -> argparse.Namespace:
    """Parse command-line arguments.
    
    Returns:
        Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description='Generate PowerPoint presentation from markdown content and template.'
    )
    
    # Config file path
    parser.add_argument(
        '--config',
        default='config.yaml',
        help='Path to configuration file (default: config.yaml)'
    )
    
    # Path overrides (note: not implemented yet, would require CLI override logic in Config)
    parser.add_argument(
        '--template',
        help='Path to PowerPoint template file (overrides config)'
    )
    
    parser.add_argument(
        '--content',
        help='Path to markdown content file (overrides config)'
    )
    
    parser.add_argument(
        '--output',
        help='Path to output PowerPoint file (overrides config)'
    )
    
    return parser.parse_args()


def main() -> int:
    """Main entry point for the CLI.
    
    Returns:
        Exit code (0 for success, 1 for error)
    """
    # Parse command-line arguments
    args = parse_arguments()
    
    # Load configuration
    try:
        config = Config(args.config)
    except FileNotFoundError as e:
        print(f"Error: {e}")
        print(f"Please ensure the configuration file exists at: {args.config}")
        return 1
    except Exception as e:
        print(f"Error loading configuration: {e}")
        return 1
    
    # Apply CLI overrides if provided
    # Note: This is a simple implementation; a more robust version would modify Config
    if args.template:
        config._config['paths']['template'] = args.template
    if args.content:
        config._config['paths']['content'] = args.content
    if args.output:
        config._config['paths']['output'] = args.output
    
    # Print banner and configuration
    print("=" * 60)
    print("ILTCI Presentation Generator")
    print("=" * 60)
    print(f"Configuration: {args.config}")
    print(f"Template:      {config.template_path}")
    print(f"Content:       {config.content_path}")
    print(f"Output:        {config.output_path}")
    print("=" * 60)
    
    # Generate presentation
    try:
        generator = PresentationGenerator(config)
        generator.generate()
    except FileNotFoundError as e:
        print(f"\nError: {e}")
        return 1
    except Exception as e:
        logging.exception("Error generating presentation")
        print(f"\nError generating presentation: {e}")
        return 1
    
    print("\n" + "=" * 60)
    print("Done!")
    print("=" * 60)
    return 0


if __name__ == '__main__':
    sys.exit(main())
