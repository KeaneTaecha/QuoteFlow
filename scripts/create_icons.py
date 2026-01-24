#!/usr/bin/env python3
"""
Icon Generator Script
Creates proper ICO and ICNS files with rounded corners from PNG source
"""

from PIL import Image, ImageDraw
import os
import subprocess
import sys

def add_rounded_corners(image, radius=200):
    """Add rounded corners to an image with extremely rounded edges"""
    # Create a mask for rounded corners
    mask = Image.new('L', image.size, 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([0, 0, image.size[0], image.size[1]], radius=radius, fill=255)
    
    # Apply the mask
    output = Image.new('RGBA', image.size, (0, 0, 0, 0))
    output.paste(image, mask=mask)
    
    return output

def create_ico_file(png_path, output_path, sizes=[16, 32, 48, 64, 128, 256]):
    """Create ICO file with multiple sizes"""
    try:
        # Open the PNG file
        img = Image.open(png_path)
        img = img.convert("RGBA")
        
        # Add extremely rounded corners (increased radius to 200)
        rounded_img = add_rounded_corners(img, radius=200)
        
        # Create ICO file with multiple sizes
        ico_sizes = []
        for size in sizes:
            resized = rounded_img.resize((size, size), Image.LANCZOS)
            ico_sizes.append(resized)
        
        # Save as ICO
        ico_sizes[0].save(output_path, format='ICO', sizes=[(s.width, s.height) for s in ico_sizes])
        print(f"✅ Created ICO file: {output_path}")
        return True
        
    except Exception as e:
        print(f"❌ Error creating ICO: {e}")
        return False

def create_icns_file(png_path, output_path):
    """Create ICNS file for macOS"""
    try:
        # Open the PNG file
        img = Image.open(png_path)
        img = img.convert("RGBA")
        
        # Add extremely rounded corners (increased radius to 200)
        rounded_img = add_rounded_corners(img, radius=200)
        
        # Create iconset directory
        iconset_dir = "icon.iconset"
        os.makedirs(iconset_dir, exist_ok=True)
        
        # Create different sizes for ICNS
        sizes = {
            "icon_16x16.png": 16,
            "icon_16x16@2x.png": 32,
            "icon_32x32.png": 32,
            "icon_32x32@2x.png": 64,
            "icon_128x128.png": 128,
            "icon_128x128@2x.png": 256,
            "icon_256x256.png": 256,
            "icon_256x256@2x.png": 512,
            "icon_512x512.png": 512,
            "icon_512x512@2x.png": 1024
        }
        
        for filename, size in sizes.items():
            resized = rounded_img.resize((size, size), Image.LANCZOS)
            resized.save(os.path.join(iconset_dir, filename))
        
        # Convert iconset to ICNS using iconutil
        try:
            subprocess.run(["iconutil", "-c", "icns", iconset_dir, "-o", output_path], check=True)
            print(f"✅ Created ICNS file: {output_path}")
            
            # Clean up iconset directory
            import shutil
            shutil.rmtree(iconset_dir)
            return True
            
        except subprocess.CalledProcessError:
            print("❌ iconutil not available, creating ICNS manually...")
            # Fallback: create a simple ICNS by saving as PNG
            rounded_img.save(output_path.replace('.icns', '_fallback.png'))
            print(f"✅ Created fallback PNG: {output_path.replace('.icns', '_fallback.png')}")
            return False
            
    except Exception as e:
        print(f"❌ Error creating ICNS: {e}")
        return False

def main():
    """Main function"""
    print("🎨 Icon Generator Script - Extremely Rounded Edges")
    print("=" * 55)
    
    # Check if source PNG exists
    png_path = "assets/icon.png"
    if not os.path.exists(png_path):
        print(f"❌ Source PNG not found: {png_path}")
        return
    
    print(f"📁 Source PNG: {png_path}")
    print(f"🔵 Radius: 200 pixels (extremely rounded edges)")
    
    # Create assets directory if it doesn't exist
    os.makedirs("assets", exist_ok=True)
    
    # Create ICO file
    ico_path = "assets/icon.ico"
    print(f"\n🪟 Creating Windows ICO file...")
    create_ico_file(png_path, ico_path)
    
    # Create ICNS file
    icns_path = "assets/icon.icns"
    print(f"\n🍎 Creating macOS ICNS file...")
    create_icns_file(png_path, icns_path)
    
    print(f"\n✅ Icon generation complete!")
    print(f"📁 Check the assets/ folder for your new icon files")
    print(f"🔵 Icons now have extremely rounded edges (radius: 200px)")

if __name__ == "__main__":
    main()