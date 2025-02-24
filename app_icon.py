import os
import win32com.client
import win32gui
import win32ui
import win32con
from PIL import Image
from configparser import ConfigParser

def get_target_and_icon_from_lnk(lnk_path):
    """Extracts the target path and icon path from a .lnk file."""
    try:
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(lnk_path)
        target_path = shortcut.Targetpath
        icon_path = shortcut.IconLocation.split(',')[0]
        
        # Use the target path if icon path is not valid
        if not icon_path or not os.path.exists(icon_path):
            icon_path = target_path
        
        if icon_path and os.path.exists(icon_path):
            print(f"Extracted icon path: {icon_path}")
            return icon_path
        else:
            print(f"No valid icon found in {lnk_path}")
    except Exception as e:
        print(f"Failed to extract target and icon from {lnk_path}: {e}")
    return None

def extract_icon(icon_path):
    """Extracts an icon from an .exe, .ico, or .dll file and returns it as an Image object."""
    try:
        # Extract icons from the file
        large, small = win32gui.ExtractIconEx(icon_path, 0)
        if not large and not small:
            print(f"No icons found in {icon_path}")
            return None
        
        hicon = large[0] if large else small[0]
        
        # Set up device context
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hdc_mem = hdc.CreateCompatibleDC()
        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(hdc, 256, 256)
        hdc_mem.SelectObject(bmp)
        
        # Draw the icon into the bitmap
        win32gui.DrawIconEx(hdc_mem.GetSafeHdc(), 0, 0, hicon, 256, 256, 0, None, win32con.DI_NORMAL)
        
        # Convert the bitmap to an image
        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)
        img = Image.frombuffer(
            'RGBA',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRA', 0, 1
        )
        
        return img  # Return the Image object directly for further processing
    except Exception as e:
        print(f"Failed to extract icon from {icon_path}: {e}")
    return None

def process_icon(img, base_name, output_folder):
    """Processes the icon image to create specific variations with transparency."""
    try:
        # Ensure the image is in RGBA mode
        img = img.convert("RGBA")
        width, height = img.size

        # Method 1: Delete lighter parts, make darker parts white
        img_contrast = Image.new("RGBA", img.size, (255, 255, 255, 0))
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 0:  # If not fully transparent
                    brightness = (r + g + b) / 3
                    if brightness < 128:  # Darker parts
                        img_contrast.putpixel((x, y), (255, 255, 255, a))

        # Save first variation
        contrast_path = os.path.join(output_folder, f"{base_name}_white.ico")
        img_contrast.save(contrast_path, format='ICO')
        print(f"Saved contrast white version: {contrast_path}")

        # Method 2: Delete darker parts, make lighter parts white
        img_contrast_alt = Image.new("RGBA", img.size, (255, 255, 255, 0))
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 0:  # If not fully transparent
                    brightness = (r + g + b) / 3
                    if brightness >= 128:  # Lighter parts
                        img_contrast_alt.putpixel((x, y), (255, 255, 255, a))

        # Save second variation
        contrast_alt_path = os.path.join(output_folder, f"{base_name}_white_alt.ico")
        img_contrast_alt.save(contrast_alt_path, format='ICO')
        print(f"Saved contrast white alt version: {contrast_alt_path}")

        # Method 3: Original image turned white while respecting transparency
        img_white = Image.new("RGBA", img.size, (255, 255, 255, 0))
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 0:  # If not fully transparent
                    img_white.putpixel((x, y), (255, 255, 255, a))

        # Save third variation
        white_path = os.path.join(output_folder, f"{base_name}_white_original.ico")
        img_white.save(white_path, format='ICO')
        print(f"Saved original white version: {white_path}")

        # Method 4: Remove black parts (make transparent), rest becomes white
        img_no_black = Image.new("RGBA", img.size, (255, 255, 255, 0))
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 0:  # If not fully transparent
                    # Check if pixel is close to black (using average RGB)
                    avg_color = (r + g + b) / 3
                    if avg_color <= 30:  # If pixel is very dark
                        continue  # Keep transparent
                    else:
                        img_no_black.putpixel((x, y), (255, 255, 255, a))

        # Save fourth variation
        no_black_path = os.path.join(output_folder, f"{base_name}_white_no_black.ico")
        img_no_black.save(no_black_path, format='ICO')
        print(f"Saved no-black white version: {no_black_path}")

        # Method 5: Remove white parts (make transparent), rest becomes white
        img_no_white = Image.new("RGBA", img.size, (255, 255, 255, 0))
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 0:  # If not fully transparent
                    # Check if pixel is close to white (using average RGB)
                    avg_color = (r + g + b) / 3
                    if avg_color >= 225:  # If pixel is very light
                        continue  # Keep transparent
                    else:
                        img_no_white.putpixel((x, y), (255, 255, 255, a))

        # Save fifth variation
        no_white_path = os.path.join(output_folder, f"{base_name}_white_no_white.ico")
        img_no_white.save(no_white_path, format='ICO')
        print(f"Saved no-white white version: {no_white_path}")

        # Method 6: Keep only white parts, make everything else transparent
        img_only_white = Image.new("RGBA", img.size, (255, 255, 255, 0))
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 0:  # If not fully transparent
                    # Check if pixel is close to white (using average RGB)
                    avg_color = (r + g + b) / 3
                    if avg_color >= 225:  # If pixel is very light/white
                        img_only_white.putpixel((x, y), (255, 255, 255, a))
                    else:
                        continue  # Keep transparent

        # Save sixth variation
        only_white_path = os.path.join(output_folder, f"{base_name}_white_only.ico")
        img_only_white.save(only_white_path, format='ICO')
        print(f"Saved white-only version: {only_white_path}")

        # Method 7: Enhanced pixel art processing specifically for detailed icons
        img_pixel = Image.new("RGBA", img.size, (255, 255, 255, 0))
        
        # Enhanced parameters for detailed pixel art
        light_threshold = 200  # For white/light pixels
        dark_threshold = 60    # For dark/detail pixels
        mid_threshold = 130    # For midtone detection
        alpha_threshold = 30   # Minimum alpha to consider
        
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a >= alpha_threshold:
                    brightness = (r + g + b) / 3
                    
                    # Detect important features (like cat's outline and details)
                    is_feature = (
                        # Catch darker details
                        (brightness <= dark_threshold) or
                        # Catch midtones that form important features
                        (dark_threshold < brightness < mid_threshold) or
                        # Catch light details that are part of the main shape
                        (brightness >= light_threshold and a > 200)
                    )
                    
                    # Check surrounding pixels for edge detection
                    if is_feature:
                        # Preserve the pixel with original alpha
                        img_pixel.putpixel((x, y), (255, 255, 255, a))

        # Save enhanced pixel art version
        pixel_path = os.path.join(output_folder, f"{base_name}_white_pix.ico")
        img_pixel.save(pixel_path, format='ICO')
        print(f"Saved enhanced pixel art version: {pixel_path}")

    except Exception as e:
        print(f"Failed to process icon for {base_name}: {e}")

def create_characteristic_variations(img_path, output_folder):
    """Creates two opposing variations based on dominant image characteristics."""
    try:
        # Open and convert image
        img = Image.open(img_path).convert("RGBA")
        base_name = os.path.splitext(os.path.basename(img_path))[0]
        width, height = img.size
        
        # Analysis arrays
        color_data = []
        total_pixels = 0
        
        # Collect color data
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 30:  # Only consider visible pixels
                    total_pixels += 1
                    brightness = (r + g + b) / 3
                    saturation = max(r, g, b) - min(r, g, b)
                    color_temp = (r - b)  # Simple warm-cool measure
                    color_data.append({
                        'pos': (x, y),
                        'brightness': brightness,
                        'saturation': saturation,
                        'temperature': color_temp,
                        'alpha': a
                    })
        
        if not color_data:
            return
        
        # Calculate variances for each characteristic
        avg_bright = sum(c['brightness'] for c in color_data) / total_pixels
        avg_sat = sum(c['saturation'] for c in color_data) / total_pixels
        avg_temp = sum(c['temperature'] for c in color_data) / total_pixels
        
        var_bright = sum((c['brightness'] - avg_bright) ** 2 for c in color_data)
        var_sat = sum((c['saturation'] - avg_sat) ** 2 for c in color_data)
        var_temp = sum((c['temperature'] - avg_temp) ** 2 for c in color_data)
        
        # Determine dominant characteristic
        characteristics = {
            'brightness': (var_bright, avg_bright, '_light', '_dark'),
            'saturation': (var_sat, avg_sat, '_saturated', '_muted'),
            'temperature': (var_temp, avg_temp, '_warm', '_cool')
        }
        
        dominant_char = max(characteristics.items(), key=lambda x: x[1][0])
        
        # Create two opposing images
        img_type1 = Image.new("RGBA", img.size, (255, 255, 255, 0))
        img_type2 = Image.new("RGBA", img.size, (255, 255, 255, 0))
        
        # Split pixels based on dominant characteristic
        avg_value = dominant_char[1][1]
        suffix1 = dominant_char[1][2]
        suffix2 = dominant_char[1][3]
        
        for y in range(height):
            for x in range(width):
                r, g, b, a = img.getpixel((x, y))
                if a > 30:
                    value = {
                        'brightness': (r + g + b) / 3,
                        'saturation': max(r, g, b) - min(r, g, b),
                        'temperature': (r - b)
                    }[dominant_char[0]]
                    
                    if value > avg_value:
                        img_type1.putpixel((x, y), (255, 255, 255, a))
                    else:
                        img_type2.putpixel((x, y), (255, 255, 255, a))
        
        # Save variations
        type1_path = os.path.join(output_folder, f"{base_name}{suffix1}.ico")
        type2_path = os.path.join(output_folder, f"{base_name}{suffix2}.ico")
        img_type1.save(type1_path, format='ICO')
        img_type2.save(type2_path, format='ICO')
        print(f"Created characteristic variations: {suffix1} and {suffix2}")
        
    except Exception as e:
        print(f"Failed to create characteristic variations: {e}")

def find_steam_libraries():
    """Finds all Steam library folders on the system."""
    steam_libraries = []
    drives = ['C:', 'D:', 'E:', 'F:', 'G:']  # Add or remove drives as needed
    
    for drive in drives:
        steam_path = os.path.join(drive, os.sep, 'Program Files (x86)', 'Steam')
        if os.path.exists(steam_path):
            steam_libraries.append(steam_path)
            # Check for additional library folders in libraryfolders.vdf
            libraryfolders_path = os.path.join(steam_path, 'steamapps', 'libraryfolders.vdf')
            if os.path.exists(libraryfolders_path):
                with open(libraryfolders_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        if '"path"' in line:
                            library_path = line.split('"path"')[1].split('"')[1]
                            if os.path.exists(library_path):
                                steam_libraries.append(library_path)
    
    return steam_libraries

def find_steam_app_icons(steam_libraries):
    """Finds Steam app icons in the given Steam libraries."""
    steam_icons = []
    for library in steam_libraries:
        steamapps_path = os.path.join(library, 'steamapps', 'common')
        if os.path.exists(steamapps_path):
            for root, dirs, files in os.walk(steamapps_path):
                for file in files:
                    if file.lower().endswith('.exe'):
                        steam_icons.append(os.path.join(root, file))
    return steam_icons

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_folder = os.path.join(script_dir, "Processed_Icons")
    os.makedirs(output_folder, exist_ok=True)
    
    files_processed = False
    
    # Search for .lnk, .exe, .dll, and .ico files
    valid_extensions = ('.lnk', '.exe', '.dll', '.ico', '.url')
    for file in os.listdir(script_dir):
        if file.lower().endswith(valid_extensions):
            file_path = os.path.join(script_dir, file)
            base_name = os.path.splitext(file)[0]
            
            if file.lower().endswith('.lnk'):
                icon_path = get_target_and_icon_from_lnk(file_path)
            else:
                icon_path = file_path  # Directly use .exe, .dll, or .ico files
            
            if icon_path:
                img = extract_icon(icon_path)
                if img:
                    # Save original icon
                    icon_save_path = os.path.join(output_folder, f"{base_name}.ico")
                    img.save(icon_save_path, format='ICO')
                    process_icon(img, base_name, output_folder)
                    files_processed = True
                else:
                    print(f"Skipping {file}, could not extract icon.")
            else:
                print(f"Skipping {file}, no valid icon found.")
    
    # Search for Steam app icons
    steam_libraries = find_steam_libraries()
    steam_icons = find_steam_app_icons(steam_libraries)
    for icon_path in steam_icons:
        base_name = os.path.splitext(os.path.basename(icon_path))[0]
        img = extract_icon(icon_path)
        if img:
            # Save original icon
            icon_save_path = os.path.join(output_folder, f"{base_name}.ico")
            img.save(icon_save_path, format='ICO')
            process_icon(img, base_name, output_folder)
            files_processed = True
        else:
            print(f"Skipping {icon_path}, could not extract icon.")
    
    if files_processed:
        print("Processing complete. Check the 'Processed_Icons' folder.")
    else:
        print("No icons were processed.")

if __name__ == "__main__":
    main()