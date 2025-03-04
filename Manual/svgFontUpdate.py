import os
import base64
import subprocess
from lxml import etree

def convert_font_to_base64(font_path):
    """Converts the font file to a base64 encoded string."""
    with open(font_path, 'rb') as font_file:
        base64_encoded_font = base64.b64encode(font_file.read()).decode('utf-8')
    return base64_encoded_font

def convert_svg_to_plain(svg_path):
    """Converts an SVG file to plain SVG format using Inkscape."""
    temp_path = svg_path + '_temp'
    
    try:
        print(f"Running Inkscape to convert {svg_path} to plain SVG...")
        result = subprocess.run([
            'inkscape',
            '--export-plain-svg=' + temp_path,
            svg_path
        ], check=True, capture_output=True, text=True)
        
        print(f"Inkscape output: {result.stdout}")
        print(f"Inkscape error (if any): {result.stderr}")
        
        # Verify the temporary file was created
        if os.path.exists(temp_path):
            print(f"Temporary file created: {temp_path}")
            os.replace(temp_path, svg_path)
            print(f"Converted and replaced {svg_path}.")
        else:
            print(f"Temporary file not found: {temp_path}")
    except subprocess.CalledProcessError as e:
        print(f"Inkscape command failed: {e}")
    except FileNotFoundError as e:
        print(f"File operation failed: {e}")

def embed_font_in_svg(svg_path, font_path):
    """Embeds the font into an SVG file by modifying its content."""
    base64_font = convert_font_to_base64(font_path)
    
    try:
        # Parse the SVG file
        parser = etree.XMLParser(remove_blank_text=True)
        tree = etree.parse(svg_path, parser)
        root = tree.getroot()

        # Define the style content
        style_content = f"""
        @font-face {{
            font-family: 'Titillium Web';
            src: url('data:font/truetype;charset=utf-8;base64,{base64_font}') format('truetype');
            font-weight: normal;
            font-style: normal;
        }}
        text {{
            font-family: 'Titillium Web', sans-serif;
        }}
        """
        
        # Check if there's already a <style> tag in the SVG
        namespaces = {'svg': 'http://www.w3.org/2000/svg'}
        style_elements = root.findall('.//svg:style', namespaces)
        
        if not style_elements:
            # Create new <style> tag if not present
            style_element = etree.Element('{http://www.w3.org/2000/svg}style', type="text/css")
            style_element.text = style_content
            root.insert(0, style_element)
        else:
            # Update existing <style> tag
            style_elements[0].text += style_content

        # Ensure that all text elements use the correct font-family
        for text in root.findall('.//svg:text', namespaces):
            if 'font-family' not in text.attrib:
                text.set('font-family', 'Titillium Web')
        
        # Save the updated SVG
        tree.write(svg_path, pretty_print=True, xml_declaration=True, encoding='utf-8')
        print(f"Font embedded in: {svg_path}")
    except etree.XMLSyntaxError as e:
        print(f"Error parsing SVG file {svg_path}: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

def has_text(svg_file):
    """Checks if the SVG file contains text elements."""
    try:
        tree = etree.parse(svg_file)
        root = tree.getroot()
        
        namespaces = {'svg': 'http://www.w3.org/2000/svg'}
        for element in root.findall('.//svg:text', namespaces):
            return True
        return False
    except etree.XMLSyntaxError as e:
        print(f"Error parsing SVG file {svg_file}: {e}")
        return False

def process_svgs_in_directory(directory, font_path):
    """Converts SVG files to plain SVG and embeds font if text is present."""
    for root_dir, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.svg'):
                svg_path = os.path.join(root_dir, file)
                
                # Convert SVG to plain SVG format (overwriting original file)
                convert_svg_to_plain(svg_path)
                
                if has_text(svg_path):
                    embed_font_in_svg(svg_path, font_path)

if __name__ == "__main__":
    current_directory = os.path.dirname(os.path.realpath(__file__))
    font_path = os.path.join(current_directory, 'TitilliumWeb-Regular.ttf')
    
    # Process SVG files in the current directory
    process_svgs_in_directory(current_directory, font_path)
