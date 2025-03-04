import os
import base64
from pathlib import Path
import logging

logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

def embed_font_in_svg(svg_path, font_path):
    """Embeds a font in Base64 format into an SVG image.

    Args:
        svg_path: Path to the SVG file.
        font_path: Path to the font file.
    """

    logger.info(f"Processing file: {svg_path}")

    try:
        # Open SVG file with 'utf-8' encoding (may need adjustment)
        with open(svg_path, "r", encoding="utf-8") as f:
            svg_content = f.read()

        # Check if SVG contains text
        if "<text" in svg_content:
            logger.debug(f"SVG file contains text: {svg_path}")

            # Read font file and base64 encode it
            with open(font_path, "rb") as f:
                font_data = f.read()
            font_b64 = base64.b64encode(font_data).decode('utf-8')

            # Create the font-face style
            font_face_style = f"""
            <style type="text/css"><![CDATA[
                @font-face {{
                    font-family: 'EmbeddedTitilliumWeb';  # Unique font family name
                    font-style: normal;
                    font-weight: normal;
                    src: url(data:font/woff;base64,{font_b64}) format('woff');
                }}
            ]]></style>
            """

            # Insert the font-face style before the closing </svg> tag
            svg_content = svg_content.replace("</svg>", font_face_style + "</svg>")

            # Replace all font-family references with the embedded font name
            svg_content = svg_content.replace('font-family:\'Titillium Web\';', 'font-family:\'EmbeddedTitilliumWeb\';')
            svg_content = svg_content.replace('-inkscape-font-specification:\'Titillium Web\';', '')

        else:
            logger.debug(f"SVG file does not contain text: {svg_path}")

    except UnicodeDecodeError as e:
        logger.error(f"Error decoding SVG file: {svg_path} ({e})")
    except Exception as e:  # Catch other potential errors
        logger.error(f"Error processing file: {svg_path} ({e})")

    else:  # Successful processing (if no exceptions)
        logger.info(f"Processed successfully: {svg_path}")

        # Overwrite the original SVG file with the modified content
        with open(svg_path, "w", encoding="utf-8") as f:
            f.write(svg_content)

def process_svg_files(directory, font_path):
    """Processes SVG files in a directory and its subdirectories.

    Args:
        directory: The root directory to search for SVG files.
        font_path: Path to the font file.
    """

    logger.info(f"Processing directory: {directory}")

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(".svg"):
                svg_path = os.path.join(root, file)
                embed_font_in_svg(svg_path, font_path)

# Example usage
font_path = "TitilliumWeb-Regular.woff"
directory = "."  # Current directory
process_svg_files(directory, font_path)
