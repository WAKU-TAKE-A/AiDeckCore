from pptx.util import Pt
from pptx.dml.color import RGBColor

class Theme:
    def __init__(self, name: str = "default"):
        self.name = name
        
        # Simple modern business palette
        self.color_primary = RGBColor(0x0F, 0x4C, 0x81)    # Classic Blue
        self.color_text = RGBColor(0x33, 0x33, 0x33)       # Dark Gray
        self.color_text_light = RGBColor(0x66, 0x66, 0x66) # Medium Gray
        self.color_background = RGBColor(0xFF, 0xFF, 0xFF) # White
        self.color_accent = RGBColor(0x00, 0xA8, 0x6B)     # Jade
        self.color_surface = RGBColor(0xF5, 0xF7, 0xFA)    # Light surface
        self.color_border = RGBColor(0xCC, 0xD3, 0xDC)     # Soft border
        
        # Typography
        self.font_name = "Calibri" # Usually safe on most systems
        self.size_title = Pt(44)
        self.size_subtitle = Pt(32)
        self.size_body = Pt(24)
        self.size_body_small = Pt(18)
