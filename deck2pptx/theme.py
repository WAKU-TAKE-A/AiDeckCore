"""Theme loading.

All tunable numbers (colors, fonts, spacing, element-specific sizes) live
in `themes/<name>.ini`, not scattered across the renderer source. This
module only loads the ini and exposes the values as namespaced attribute
groups, e.g.:

    theme.color.primary       # RGBColor
    theme.font.size_body      # Pt
    theme.flow.node_width     # Inches (a Length, i.e. EMU)
    theme.calibration.emu_per_inch  # plain float

To tune a value, edit the ini file. Do not reintroduce literals into the
renderer/element_renderers modules; add a new ini key + attribute instead.
"""

import configparser
from pathlib import Path

from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor

_THEMES_DIR = Path(__file__).parent / "themes"


def _color(cp, section, key) -> RGBColor:
    return RGBColor.from_string(cp.get(section, key))


def _pt(cp, section, key) -> Pt:
    return Pt(cp.getfloat(section, key))


def _in(cp, section, key) -> Inches:
    return Inches(cp.getfloat(section, key))


def _f(cp, section, key) -> float:
    return cp.getfloat(section, key)


def _i(cp, section, key) -> int:
    return cp.getint(section, key)


def _s(cp, section, key) -> str:
    return cp.get(section, key)


class ColorTheme:
    """Namespace: theme.color.*"""
    def __init__(self, cp):
        s = "color"
        self.primary = _color(cp, s, "primary")
        self.text = _color(cp, s, "text")
        self.text_light = _color(cp, s, "text_light")
        self.background = _color(cp, s, "background")
        self.accent = _color(cp, s, "accent")
        self.surface = _color(cp, s, "surface")
        self.border = _color(cp, s, "border")
        self.flow_fill = _color(cp, s, "flow_fill")
        self.flow_line = _color(cp, s, "flow_line")
        self.flow_text = _color(cp, s, "flow_text")


class FontTheme:
    """Namespace: theme.font.*"""
    def __init__(self, cp):
        s = "font"
        self.name = _s(cp, s, "name")
        self.name_code = _s(cp, s, "name_code")
        self.size_title = _pt(cp, s, "size_title")
        self.size_subtitle = _pt(cp, s, "size_subtitle")
        self.size_body = _pt(cp, s, "size_body")
        self.size_body_semi_small = _pt(cp, s, "size_body_semi_small")
        self.size_body_small = _pt(cp, s, "size_body_small")
        self.size_body_extra_small = _pt(cp, s, "size_body_extra_small")
        # Defensive fallback only; see note in height_estimator._get_font_size.
        self.fallback_size_level0 = _f(cp, s, "fallback_size_level0")
        self.fallback_size_other = _f(cp, s, "fallback_size_other")


class LayoutTheme:
    """Namespace: theme.layout.* (slide-level layout, not per-element)"""
    def __init__(self, cp):
        s = "layout"
        self.margin_x = _in(cp, s, "margin_x")
        self.margin_y = _in(cp, s, "margin_y")
        self.title_height = _in(cp, s, "title_height")
        self.title_body_gap = _in(cp, s, "title_body_gap")
        self.element_gap = _in(cp, s, "element_gap")
        self.min_remaining_height = _in(cp, s, "min_remaining_height")
        self.slide_width_landscape = _in(cp, s, "slide_width_landscape")
        self.slide_height_landscape = _in(cp, s, "slide_height_landscape")
        self.slide_width_portrait = _in(cp, s, "slide_width_portrait")
        self.slide_height_portrait = _in(cp, s, "slide_height_portrait")
        self.align_offset_small = _in(cp, s, "align_offset_small")
        self.align_offset_large = _in(cp, s, "align_offset_large")
        self.title_slide_title_box_height = _in(cp, s, "title_slide_title_box_height")
        self.title_slide_subtitle_box_height = _in(cp, s, "title_slide_subtitle_box_height")
        self.title_slide_position_ratio = _f(cp, s, "title_slide_position_ratio")


class TextTheme:
    """Namespace: theme.text.* (Text element rendering + height estimation)"""
    def __init__(self, cp):
        s = "text"
        self.line_height = _in(cp, s, "line_height")
        self.padding = _in(cp, s, "padding")
        self.fallback_chars_per_line = _i(cp, s, "fallback_chars_per_line")


class BulletTheme:
    """Namespace: theme.bullet.*"""
    def __init__(self, cp):
        s = "bullet"
        self.line_height = _in(cp, s, "line_height")
        self.padding = _in(cp, s, "padding")
        self.indent_per_level = _f(cp, s, "indent_per_level")
        self.level_weight_default = _f(cp, s, "level_weight_default")
        self.level_weight_indented = _f(cp, s, "level_weight_indented")


class ImageTheme:
    """Namespace: theme.image.*"""
    def __init__(self, cp):
        s = "image"
        self.caption_height = _in(cp, s, "caption_height")
        self.gallery_padding = _in(cp, s, "gallery_padding")


class TableTheme:
    """Namespace: theme.table.*"""
    def __init__(self, cp):
        s = "table"
        self.placeholder_init_height = _in(cp, s, "placeholder_init_height")
        self.header_fill_color = _color(cp, s, "header_fill_color")
        self.header_font_size = _pt(cp, s, "header_font_size")
        self.cell_font_size = _pt(cp, s, "cell_font_size")
        self.border_color = _s(cp, s, "border_color")
        self.border_width_pt = _f(cp, s, "border_width_pt")


class FlowTheme:
    """Namespace: theme.flow.*"""
    def __init__(self, cp):
        s = "flow"
        self.node_width = _in(cp, s, "node_width")
        self.node_height = _in(cp, s, "node_height")
        self.node_gap = _in(cp, s, "node_gap")
        self.arrow_min_length = _in(cp, s, "arrow_min_length")
        self.arrow_thickness = _in(cp, s, "arrow_thickness")
        self.line_offset = _in(cp, s, "line_offset")


class TreeTheme:
    """Namespace: theme.tree.*"""
    def __init__(self, cp):
        s = "tree"
        self.node_width = _in(cp, s, "node_width")
        self.node_height = _in(cp, s, "node_height")
        self.vertical_gap = _in(cp, s, "vertical_gap")
        self.horizontal_gap = _in(cp, s, "horizontal_gap")
        self.connector_line_width = _pt(cp, s, "connector_line_width_pt")


class CodeTheme:
    """Namespace: theme.code.*"""
    def __init__(self, cp):
        s = "code"
        self.caption_height = _in(cp, s, "caption_height")
        self.line_height_factor = _f(cp, s, "line_height_factor")
        self.height_padding = _f(cp, s, "height_padding")


class ComparisonTheme:
    """Namespace: theme.comparison.* (also covers Timeline title/columns part)"""
    def __init__(self, cp):
        s = "comparison"
        self.title_height = _in(cp, s, "title_height")
        self.padding = _in(cp, s, "padding")
        self.fallback_chars_per_line = _i(cp, s, "fallback_chars_per_line")


class TimelineTheme:
    """Namespace: theme.timeline.*"""
    def __init__(self, cp):
        s = "timeline"
        self.event_height = _in(cp, s, "event_height")
        self.label_width = _in(cp, s, "label_width")
        self.line_x_offset = _in(cp, s, "line_x_offset")
        self.line_width = _in(cp, s, "line_width")
        self.line_vertical_padding = _in(cp, s, "line_vertical_padding")
        self.content_x_offset = _in(cp, s, "content_x_offset")


class SplitTheme:
    """Namespace: theme.split.*"""
    def __init__(self, cp):
        s = "split"
        self.gap = _in(cp, s, "gap")
        self.title_box_height = _in(cp, s, "title_box_height")
        self.title_advance = _in(cp, s, "title_advance")


class MermaidTheme:
    """Namespace: theme.mermaid.*"""
    def __init__(self, cp):
        s = "mermaid"
        self.default_height = _in(cp, s, "default_height")


class CalibrationTheme:
    """Namespace: theme.calibration.*

    Plain floats only (not Length/Pt objects): these are used in raw
    arithmetic (EMU<->inch conversion, fallback chars-per-inch, etc.),
    not assigned directly to python-pptx font/size properties.
    """
    def __init__(self, cp):
        s = "calibration"
        self.emu_per_inch = _f(cp, s, "emu_per_inch")
        self.pt_to_emu = _f(cp, s, "pt_to_emu")
        self.line_height_factor = _f(cp, s, "line_height_factor")
        self.fallback_cpi = _f(cp, s, "fallback_cpi")
        self.fallback_width_inches = _f(cp, s, "fallback_width_inches")


class Theme:
    """Aggregates all namespaced theme sub-objects, loaded from
    `themes/<name>.ini` (falls back to `themes/default.ini` if a theme
    with the given name does not exist).
    """

    def __init__(self, name: str = "default"):
        self.name = name
        cp = configparser.ConfigParser()
        ini_path = _THEMES_DIR / f"{name}.ini"
        if not ini_path.is_file():
            ini_path = _THEMES_DIR / "default.ini"

        read_files = cp.read(ini_path, encoding="utf-8")
        if not read_files:
            raise FileNotFoundError(
                f"Theme ini file not found: {ini_path} "
                f"(expected a default theme at {_THEMES_DIR / 'default.ini'})."
            )

        self.color = ColorTheme(cp)
        self.font = FontTheme(cp)
        self.layout = LayoutTheme(cp)
        self.text = TextTheme(cp)
        self.bullet = BulletTheme(cp)
        self.image = ImageTheme(cp)
        self.table = TableTheme(cp)
        self.flow = FlowTheme(cp)
        self.tree = TreeTheme(cp)
        self.code = CodeTheme(cp)
        self.comparison = ComparisonTheme(cp)
        self.timeline = TimelineTheme(cp)
        self.split = SplitTheme(cp)
        self.mermaid = MermaidTheme(cp)
        self.calibration = CalibrationTheme(cp)
