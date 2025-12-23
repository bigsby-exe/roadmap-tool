"""
Configuration module - CUSTOMIZE YOUR BRANDING HERE
All branding and styling options are in this file.
"""

from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

################################################################################
### CONFIGURATION SECTION - CUSTOMIZE YOUR BRANDING HERE ###
################################################################################

# BRANDING CONFIG - Colors (RGB values 0-255)
BRAND_PRIMARY_COLOR = RGBColor(0, 51, 102)      # Deep blue - CHANGE THIS
BRAND_SECONDARY_COLOR = RGBColor(0, 102, 204)   # Medium blue - CHANGE THIS
BRAND_ACCENT_COLOR = RGBColor(255, 153, 0)      # Orange accent - CHANGE THIS
BRAND_TEXT_COLOR = RGBColor(51, 51, 51)         # Dark gray text - CHANGE THIS
BRAND_BACKGROUND_COLOR = RGBColor(255, 255, 255) # White background - CHANGE THIS

# Logo Configuration
LOGO_PATH = None  # Set to path of your logo file (e.g., "logo.png") or None to skip
LOGO_POSITION = "top_right"  # Options: "top_left", "top_right", "bottom_left", "bottom_right", "center"

# Font Configuration
TITLE_FONT_NAME = "Calibri"  # CHANGE THIS to your brand font
BODY_FONT_NAME = "Calibri"   # CHANGE THIS to your brand font
TITLE_FONT_SIZE = Pt(44)
SUBTITLE_FONT_SIZE = Pt(28)
HEADING_FONT_SIZE = Pt(32)
BODY_FONT_SIZE = Pt(18)

# SLIDE LAYOUT CONFIG - Dimensions and spacing
SLIDE_WIDTH = Inches(10)
SLIDE_HEIGHT = Inches(7.5)
TITLE_TOP_MARGIN = Inches(1)
CONTENT_TOP_MARGIN = Inches(1.5)
SIDE_MARGIN = Inches(0.5)
BOTTOM_MARGIN = Inches(0.5)

# VISUAL STYLE CONFIG
USE_SHAPES = True  # Use rounded rectangles for content boxes
SHAPE_CORNER_RADIUS = Inches(0.1)
CONTENT_BOX_COLOR = RGBColor(245, 245, 245)  # Light gray for content boxes

# PAGINATION CONFIG
KEY_ELEMENT_HEIGHT_ESTIMATE = Inches(0.5)  # Estimated height per key element item
NORTH_STAR_MIN_HEIGHT = Inches(0.8)  # Minimum height for north star box
NORTH_STAR_MAX_HEIGHT = Inches(3.0)  # Maximum height for north star box
TEXT_BOX_MARGIN = Inches(0.2)  # Consistent margin for text boxes
CHARS_PER_LINE_ESTIMATE = 80  # Estimated characters per line for text wrapping

# TEMPLATE CONFIG
TITLE_SLIDE_TEMPLATE = None  # Path to template file for title slide (.pptx or .potx) or None to skip
CONTENT_SLIDE_TEMPLATE = None  # Path to template file for content slides (.pptx or .potx) or None to skip
TEMPLATE_SLIDE_INDEX = 0  # Which slide from template to use (0 = first slide)

# OVERVIEW SLIDE CONFIG
OVERVIEW_TIMELINE_SHAPE_COLOR = RGBColor(0, 51, 102)  # Timeline shape fill color
OVERVIEW_TIMELINE_TEXT_COLOR = RGBColor(255, 255, 255)  # Timeline text color (should contrast with shape color)
OVERVIEW_CHEVRON_COLOR = RGBColor(255, 153, 0)  # Chevron connector color (kept for potential future use)
OVERVIEW_SHAPE_HEIGHT = Inches(1.3)  # Height of timeline shapes (increased from 1.0)
OVERVIEW_SHAPE_WIDTH = Inches(2.2)  # Width of timeline shapes (increased from 1.75)
OVERVIEW_CHEVRON_WIDTH = Inches(2.0)  # Width of chevron connectors (kept for potential future use)
OVERVIEW_CHEVRON_HEIGHT = Inches(1.0)  # Height of chevron connectors (kept for potential future use)

################################################################################
### END CONFIGURATION SECTION ###
################################################################################

