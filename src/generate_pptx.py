#!/usr/bin/env python3
"""
Script to apply markdown content to the ILTCI PowerPoint template.
This is a thin wrapper around the iltci_pptx package.
"""

import sys
from iltci_pptx.cli import main

if __name__ == '__main__':
    sys.exit(main())
