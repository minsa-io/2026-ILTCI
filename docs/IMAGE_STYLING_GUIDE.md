# Image Styling Guide

This guide explains how to use the CSS classes defined in [`assets/image-layout.css`](../assets/image-layout.css) to style images in your presentation markdown.

## Basic Usage

Instead of using inline styles, you can now use predefined CSS classes:

### Before (inline styles):
```html
<div style="display: flex; justify-content: center; gap: 50px; margin-top: 30px;">
  <img src="assets/image1.png" alt="Image 1" style="height: 300px; object-fit: contain;">
  <img src="assets/image2.png" alt="Image 2" style="height: 300px; object-fit: contain;">
</div>
```

### After (CSS classes):
```html
<div class="image-container">
  <img src="assets/image1.png" alt="Image 1" class="img-medium">
  <img src="assets/image2.png" alt="Image 2" class="img-medium">
</div>
```

## Container Classes

### `.image-container`
Centers images horizontally with a 50px gap. Perfect for side-by-side images.

```html
<div class="image-container">
  <img src="assets/image1.png" alt="Image 1" class="img-medium">
  <img src="assets/image2.png" alt="Image 2" class="img-medium">
</div>
```

### Gap Variations
- `.image-container.small-gap` - 20px gap
- `.image-container.large-gap` - 80px gap

```html
<div class="image-container small-gap">
  <img src="assets/image1.png" alt="Image 1" class="img-medium">
  <img src="assets/image2.png" alt="Image 2" class="img-medium">
</div>
```

## Image Size Classes

Apply these classes directly to `<img>` tags:

- `.img-small` - 200px height
- `.img-medium` - 300px height
- `.img-large` - 400px height
- `.img-full-width` - 100% width, auto height

```html
<img src="assets/screenshot.png" alt="Screenshot" class="img-large">
```

## Image Alignment

- `.img-left` - Align image to the left
- `.img-right` - Align image to the right
- `.img-center` - Center the image

```html
<img src="assets/logo.png" alt="Logo" class="img-small img-center">
```

## Grid Layouts

### 2-Column Grid
```html
<div class="image-grid-2">
  <img src="assets/image1.png" alt="Image 1">
  <img src="assets/image2.png" alt="Image 2">
</div>
```

### 3-Column Grid
```html
<div class="image-grid-3">
  <img src="assets/image1.png" alt="Image 1">
  <img src="assets/image2.png" alt="Image 2">
  <img src="assets/image3.png" alt="Image 3">
</div>
```

### 4-Image Grid (2x2)
```html
<div class="image-grid-4">
  <img src="assets/image1.png" alt="Image 1">
  <img src="assets/image2.png" alt="Image 2">
  <img src="assets/image3.png" alt="Image 3">
  <img src="assets/image4.png" alt="Image 4">
</div>
```

## Flexbox Layouts

### Row Layout (horizontal)
```html
<div class="image-row">
  <img src="assets/image1.png" alt="Image 1" class="img-medium">
  <img src="assets/image2.png" alt="Image 2" class="img-medium">
  <img src="assets/image3.png" alt="Image 3" class="img-medium">
</div>
```

### Column Layout (vertical)
```html
<div class="image-column">
  <img src="assets/image1.png" alt="Image 1" class="img-medium">
  <img src="assets/image2.png" alt="Image 2" class="img-medium">
</div>
```

## Special Layouts

### Side-by-Side (Auto-sized images)
Similar to `.image-container` but with predefined image sizing:
```html
<div class="side-by-side">
  <img src="assets/image1.png" alt="Image 1">
  <img src="assets/image2.png" alt="Image 2">
</div>
```

### Image with Caption
```html
<div class="image-with-caption">
  <img src="assets/screenshot.png" alt="Screenshot" class="img-medium">
  <div class="caption">Figure 1: Application Screenshot</div>
</div>
```

## Combining Classes

You can combine multiple classes for more control:

```html
<!-- Small gap between medium-sized images -->
<div class="image-container small-gap">
  <img src="assets/image1.png" alt="Image 1" class="img-medium">
  <img src="assets/image2.png" alt="Image 2" class="img-medium">
  <img src="assets/image3.png" alt="Image 3" class="img-medium">
</div>
```

## Border and Rounded Corner Styling

By default, all images in the generated PowerPoint presentation have:
- **Rounded corners** (8px/0.1in radius)
- **Border** (2pt solid, color #44546A - dark blue-gray)

### Override Classes

You can override these defaults per-image using the following classes:

#### Border Controls
- `.no-border` - Remove the border completely
- `.border-thin` - Use a thinner border (1pt)
- `.border-thick` - Use a thicker border (4pt)
- `.border-light` - Use a lighter border color (#B4C6E7)
- `.border-dark` - Use a darker border color (#2F3E4E)

#### Rounded Corner Controls
- `.no-rounded` - Remove rounded corners (square edges)
- `.rounded-sm` - Use smaller rounded corners (4px/0.05in)
- `.rounded-lg` - Use larger rounded corners (16px/0.2in)

### Examples

```html
<!-- Image with no border -->
<div class="image-container">
  <img src="assets/screenshot.png" alt="Screenshot" class="img-large no-border">
</div>

<!-- Image with thick border and larger corners -->
<div class="image-container">
  <img src="assets/diagram.png" alt="Diagram" class="img-medium border-thick rounded-lg">
</div>

<!-- Image with no border and no rounded corners (square) -->
<div class="image-container">
  <img src="assets/photo.png" alt="Photo" class="img-large no-border no-rounded">
</div>

<!-- Combining multiple style classes -->
<div class="image-container">
  <img src="assets/chart.png" alt="Chart" class="img-medium border-thin border-light rounded-sm">
</div>
```

### When to Override

- **Screenshots**: Often look better without borders (`no-border`)
- **Photos**: May benefit from no corners (`no-rounded`) for a natural look
- **Diagrams**: Consider thicker borders (`border-thick`) for emphasis
- **Side-by-side comparisons**: Use consistent styling for both images

## Tips

1. **Consistency**: Use the same classes throughout your presentation for a cohesive look
2. **Sizing**: Start with `.img-medium` (300px) and adjust as needed
3. **Spacing**: The default gap (50px) works well for most cases
4. **Testing**: Preview your slides after making changes to ensure proper layout
5. **Borders**: The default border provides a clean framing effect; override with `no-border` for screenshots that should blend seamlessly

## Customization

If you need custom sizes or layouts, edit [`assets/image-layout.css`](../assets/image-layout.css) to add new classes or modify existing ones.

For deeper customization of border and corner styles, modify the `IMAGE_STYLE_DEFAULTS` and `STYLE_CLASS_MAP` constants in [`src/iltci_pptx/images.py`](../src/iltci_pptx/images.py).
