# ILTCI Template Styling Guide

## Overview

You have two approaches for creating your presentation with the ILTCI template styling:

### Approach 1: Marp CLI (Automated, CSS-based styling)
- ✅ Fast and automated
- ✅ Content in Markdown
- ❌ Background images are rendered into slides (not editable in PowerPoint)
- ❌ Images become part of the slide image, not separate objects

### Approach 2: Python Script (Template-based)
- ✅ Preserves all template elements (images, backgrounds, layouts)
- ✅ Fully editable in PowerPoint
- ✅ Perfect template fidelity
- ❌ Requires python-pptx library

### Approach 3: Manual Copy (Highest fidelity)
- ✅ Perfect template preservation
- ✅ Full PowerPoint control
- ❌ Manual work required

---

## Approach 1: Marp CLI (Current Setup)

### Files
- **slides.md** - Your content in Markdown
- **iltci-theme.css** - Marp theme matching template colors/fonts
- **title_slide_bg_image1.png** - Title slide background
- **title_slide_bg_image2.png** - Logo for content slides

### Usage
```bash
# Generate PPTX
marp slides.md --pptx --theme-set ./iltci-theme.css -o presentation.pptx

# Preview in browser
marp slides.md --html --theme-set ./iltci-theme.css -o preview.html

# Watch mode
marp slides.md --pptx --theme-set ./iltci-theme.css --watch
```

### Styling Applied
✅ Correct colors (navy #00274B, green #3D6B38, mint #9CCCB4)
✅ Correct fonts (Calibri Light for titles, Calibri for body)
✅ Correct font sizes (50pt titles, 28pt body)
✅ Title slide background image
✅ Content slide header bar and logo
✅ Footer with session title and slide numbers

### Limitations
- Background images are rendered as part of the slide (not separate editable objects)
- To edit images in PowerPoint, you'd need to replace the entire slide background
- This is a Marp limitation, not a theme issue

---

## Approach 2: Python Script (Recommended for Full Template Fidelity)

### Setup

**Option A: Using uv (Recommended - Fast & Simple)**
```bash
# Install uv if needed
curl -LsSf https://astral.sh/uv/install.sh | sh

# Run the script (uv handles everything automatically)
uv run apply_content_to_template.py
```

**Option B: Traditional virtual environment**
```bash
sudo apt install python3.12-venv
python3 -m venv venv
source venv/bin/activate
pip install python-pptx
python apply_content_to_template.py
deactivate
```

See [`INSTALLATION.md`](INSTALLATION.md) for detailed instructions.

### What It Does
- Takes your markdown content from **slides.md**
- Applies it to **2026 ILTCI AF PPT template.pptx**
- Creates **2026_ILTCI_presentation_from_template.pptx**
- Preserves ALL template elements:
  - Background images as original objects
  - All layers and styling
  - Editable text boxes
  - Logos positioned correctly

### Result
A PowerPoint file with:
- Your markdown content
- Full template styling
- Editable images and backgrounds
- Professional ILTCI template appearance

---

## Approach 3: Manual Process

If you prefer full control or can't use the Python script:

### Steps
1. Open **2026 ILTCI AF PPT template.pptx** in PowerPoint
2. Open **slides.md** in VS Code or text editor
3. For each slide:
   - Duplicate the appropriate template slide (title or content)
   - Copy text from markdown
   - Paste into PowerPoint placeholders
   - Format lists/bullets as needed
4. Save as your final presentation

### Advantages
- Perfect template preservation
- Full control over every element
- No tool dependencies
- Can add custom elements easily

---

## Controlling Vertical Spacing

The presentation generator supports blank lines in markdown for controlling vertical spacing between content elements.

### How It Works

Blank lines in your markdown content (after the first line of content) are converted to spacer elements in the rendered slides. This allows you to:

- Add visual separation between text blocks
- Create breathing room between headers and body text
- Control the vertical rhythm of your slides

### Example

```markdown
# New skills

**Using with Agents**

You will need to learn how to integrate LLMs directly into your workspace

This is a separate paragraph with space above it
```

In this example, the blank line between the bold text and the following paragraph creates vertical spacing in the rendered slide.

### Configuration

The amount of vertical spacing can be configured in `assets/template-config.yaml`:

```yaml
fonts:
  title_slide:
    spacer: 8    # Points of spacing for blank lines in title slides
  content_slide:
    spacer: 12   # Points of spacing for blank lines in content slides
```

Increase these values for more vertical space, decrease for less.

### Notes

- Blank lines before the first content line are ignored (leading blank lines are trimmed)
- Multiple consecutive blank lines will each add spacing
- This feature works with both title slides and content slides

---

## Comparison Table

| Feature | Marp CLI | Python Script | Manual |
|---------|----------|---------------|--------|
| Speed | ⭐⭐⭐ | ⭐⭐ | ⭐ |
| Template Fidelity | ⭐⭐ | ⭐⭐⭐ | ⭐⭐⭐ |
| Editable Images | ❌ | ✅ | ✅ |
| Markdown Workflow | ✅ | ✅ | ❌ |
| Setup Required | None | pip install | None |

---

## Recommendation

**For this presentation:**

Since you want the images to be editable in PowerPoint:

1. **Best Option**: Use the Python script (Approach 2) with uv
   - Install uv: `curl -LsSf https://astral.sh/uv/install.sh | sh`
   - Run: `uv run apply_content_to_template.py`
   - uv handles all dependencies automatically!

2. **Quick Option**: Use Marp CLI and accept that backgrounds are rendered
   - Good for drafts and quick iterations
   - Images are present but embedded in slide renders

3. **Manual Option**: Copy content manually into the template
   - Best for final presentation with custom adjustments

---

## Current File Organization

```
2026 ILTCI/
├── slides.md                              # Your content
├── iltci-theme.css                        # Marp theme
├── title_slide_bg_image1.png             # Title background
├── title_slide_bg_image2.png             # Logo
├── template_styles_extracted.md           # Style documentation
├── 2026 ILTCI AF PPT template.pptx       # Original template
├── apply_content_to_template.py           # Python script
├── pyproject.toml                         # Python dependencies for uv
├── INSTALLATION.md                        # Installation guide
└── README_STYLING.md                      # This file
```

---

## Next Steps

Choose your approach based on your needs:
- **For drafts/iteration**: Use Marp CLI (already set up)
- **For final presentation with editable images**: Use Python script with uv
- **For full PowerPoint control**: Manual method

### Quick Start (Recommended):
```bash
# Install uv (if needed)
curl -LsSf https://astral.sh/uv/install.sh | sh

# Run the script
uv run apply_content_to_template.py
```

See [`INSTALLATION.md`](INSTALLATION.md) for more details!