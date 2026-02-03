#!/usr/bin/env python3
from PIL import Image, ImageDraw, ImageFont

def create_icon(size, filename):
    # Create blue square with white "CT" text
    img = Image.new('RGB', (size, size), color='#0078D4')
    draw = ImageDraw.Draw(img)
    
    # Try to use a decent font, fallback to default
    try:
        font_size = int(size * 0.4)
        font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)
    except:
        font = ImageFont.load_default()
    
    # Draw "CT" text centered
    text = "CT"
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    
    position = ((size - text_width) / 2, (size - text_height) / 2 - bbox[1])
    draw.text(position, text, fill='white', font=font)
    
    img.save(filename)
    print(f"Created {filename}")

create_icon(16, '/Users/jason/c0d3t3xt/TaskPaneMod/assets/icon-16.png')
create_icon(32, '/Users/jason/c0d3t3xt/TaskPaneMod/assets/icon-32.png')
create_icon(64, '/Users/jason/c0d3t3xt/TaskPaneMod/assets/icon-64.png')
create_icon(80, '/Users/jason/c0d3t3xt/TaskPaneMod/assets/icon-80.png')
