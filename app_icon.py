
import win32com.client
import win32gui
import win32ui
import win32con
from configparser import ConfigParser
from PIL import Image, ImageFilter, ImageOps
import os
from PIL import Image, ImageFilter, ImageOps
import os
import numpy as np
import cv2


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

def enhanced_remove_artifacts(image, min_cluster_size=5):
    width, height = image.size
    cleaned = image.copy()
    visited = set()
    
    def get_cluster(x, y):
        cluster = set()
        stack = [(x, y)]
        
        while stack:
            cx, cy = stack.pop()
            if (cx, cy) in cluster:
                continue
                
            cluster.add((cx, cy))
            
            # Check 8-connected neighbors
            for dx, dy in [(-1,-1), (-1,0), (-1,1), (0,-1), (0,1), (1,-1), (1,0), (1,1)]:
                nx, ny = cx + dx, cy + dy
                if (0 <= nx < width and 0 <= ny < height and 
                    (nx, ny) not in cluster and 
                    image.getpixel((nx, ny))[3] > 0):
                    stack.append((nx, ny))
        
        return cluster
    
    # Find and remove small clusters
    for y in range(height):
        for x in range(width):
            if (x, y) not in visited and image.getpixel((x, y))[3] > 0:
                cluster = get_cluster(x, y)
                visited.update(cluster)
                
                if len(cluster) < min_cluster_size:
                    for cx, cy in cluster:
                        cleaned.putpixel((cx, cy), (255, 255, 255, 0))
    
    return cleaned

def remove_artifacts(image):
    width, height = image.size
    cleaned = image.copy()
    
    for y in range(height):
        for x in range(width):
            r, g, b, a = image.getpixel((x, y))
            if a > 0 and r == 255 and g == 255 and b == 255:
                has_white_neighbor = False
                for j in range(max(0, y - 1), min(height, y + 2)):
                    for i in range(max(0, x - 1), min(width, x + 2)):
                        if (i, j) == (x, y):
                            continue
                        nr, ng, nb, na = image.getpixel((i, j))
                        if na > 0 and nr == 255 and ng == 255 and nb == 255:
                            has_white_neighbor = True
                            break
                    if has_white_neighbor:
                        break
                if not has_white_neighbor:
                    cleaned.putpixel((x, y), (255, 255, 255, 0))
    return cleaned.filter(ImageFilter.SMOOTH_MORE)

def apply_antialiasing(image):
    return image.filter(ImageFilter.SMOOTH_MORE)

def edge_detection(img):
    grayscale = img.convert("L")
    edges = grayscale.filter(ImageFilter.FIND_EDGES)
    edges = edges.point(lambda p: 255 if p > 50 else 0)  # Thresholding
    edges = ImageOps.invert(edges).convert("RGBA")
    return edges

def extract_edges_and_lines(image):
    # Convert PIL Image to numpy array for OpenCV processing
    img_array = np.array(image)
    
    # Convert to grayscale if needed
    if len(img_array.shape) == 3:
        gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    else:
        gray = img_array
    
    # Apply Canny edge detection
    edges = cv2.Canny(gray, 50, 150, apertureSize=3)
    
    # Apply probabilistic Hough transform
    lines = cv2.HoughLinesP(
        edges,
        rho=1,
        theta=np.pi/180,
        threshold=50,
        minLineLength=20,
        maxLineGap=10
    )
    
    # Create blank image for lines
    line_image = np.zeros_like(img_array)
    
    if lines is not None:
        for line in lines:
            x1, y1, x2, y2 = line[0]
            cv2.line(line_image, (x1, y1), (x2, y2), (255, 255, 255), 2)
    
    # Convert back to PIL Image
    return Image.fromarray(line_image)

def process_icon_with_edges(img, base_name, output_folder):
    try:
        img = img.convert("RGBA")
        
        # Edge detection based processing
        edges = edge_detection(img)
        edges = enhanced_remove_artifacts(edges)
        edges = apply_antialiasing(edges)
        
        # Extract lines using enhanced edge detection
        line_image = extract_edges_and_lines(img)
        line_image = line_image.convert("RGBA")
        
        # Apply post-processing to line image
        line_image = enhanced_remove_artifacts(line_image)
        line_image = apply_antialiasing(line_image)
        
        # Save edge-based results
        edge_path = os.path.join(output_folder, f"{base_name}_white_edges.ico")
        line_path = os.path.join(output_folder, f"{base_name}_white_lines.ico")
        
        edges.save(edge_path, format='ICO')
        line_image.save(line_path, format='ICO')
        
        print(f"Saved edge detection versions: {edge_path}, {line_path}")
        
    except Exception as e:
        print(f"Failed to process edge detection for {base_name}: {e}")

def process_icon(img, base_name, output_folder):
    try:
        img = img.convert("RGBA")
        width, height = img.size

        def process_variation(condition_fn, filename_suffix):
            img_variant = Image.new("RGBA", img.size, (255, 255, 255, 0))
            for y in range(height):
                for x in range(width):
                    r, g, b, a = img.getpixel((x, y))
                    if a > 0 and condition_fn(r, g, b):
                        img_variant.putpixel((x, y), (255, 255, 255, a))
            img_variant = enhanced_remove_artifacts(img_variant)
            img_variant = apply_antialiasing(img_variant)
            output_path = os.path.join(output_folder, f"{base_name}_{filename_suffix}.ico")
            img_variant.save(output_path, format='ICO')
            print(f"Saved {filename_suffix} version: {output_path}")

        # Standard variations
        process_variation(lambda r, g, b: (r + g + b) / 3 < 128, "white")
        process_variation(lambda r, g, b: (r + g + b) / 3 >= 128, "white_alt")
        process_variation(lambda r, g, b: True, "white_original")
        process_variation(lambda r, g, b: (r + g + b) / 3 > 30, "white_no_black")
        process_variation(lambda r, g, b: (r + g + b) / 3 < 225, "white_no_white")
        process_variation(lambda r, g, b: (r + g + b) / 3 >= 225, "white_only")
        process_variation(lambda r, g, b: (r + g + b) / 3 <= 60 or 60 < (r + g + b) / 3 < 130 or ((r + g + b) / 3 >= 200 and r > 200 and g > 200 and b > 200), "white_pix")
        
        # Edge and line detection processing
        process_icon_with_edges(img, base_name, output_folder)
        
        # Characteristic variations
        create_characteristic_variations(img_path=None, output_folder=output_folder, img=img, base_name=base_name)

    except Exception as e:
        print(f"Failed to process icon for {base_name}: {e}")

def create_characteristic_variations(img_path=None, output_folder=None, img=None, base_name=None):
    """Creates two opposing variations based on dominant image characteristics."""
    try:
        # Handle both direct image input and path input
        if img_path is not None:
            img = Image.open(img_path).convert("RGBA")
            base_name = os.path.splitext(os.path.basename(img_path))[0]
        elif img is None or base_name is None:
            raise ValueError("Either img_path or both img and base_name must be provided")
            
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
        
        # Apply enhanced processing to characteristic variations
        img_type1 = enhanced_remove_artifacts(img_type1)
        img_type2 = enhanced_remove_artifacts(img_type2)
        img_type1 = apply_antialiasing(img_type1)
        img_type2 = apply_antialiasing(img_type2)
        
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