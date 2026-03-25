"""
style_engine.py — Maps AI visual style analysis to concrete DOCX formatting

Takes the JSON from Gemini's style analysis call and produces
a StyleMap used by the DOCX builder to apply precise formatting.
"""

from dataclasses import dataclass, field
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ─── Font families ─────────────────────────────────────────────────────────────
FONT_MAP = {
    "sans-serif": "Calibri",
    "serif":      "Georgia",
    "mono":       "Courier New",
    "handwriting": "Calibri",   # fallback
    "default":    "Calibri",
}

# ─── Size categories → Pt ──────────────────────────────────────────────────────
SIZE_MAP = {
    "xlarge": 20,
    "large":  16,
    "medium": 13,
    "normal": 11,
    "small":  10,
    "tiny":    9,
}

# ─── Alignment ─────────────────────────────────────────────────────────────────
ALIGN_MAP = {
    "center":  WD_ALIGN_PARAGRAPH.CENTER,
    "right":   WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    "left":    WD_ALIGN_PARAGRAPH.LEFT,
}

# ─── Spacing presets (space_before, space_after, line_spacing multiplier) ──────
SPACING_MAP = {
    "loose":  (Pt(12), Pt(10), 1.5),
    "normal": (Pt(6),  Pt(6),  1.15),
    "tight":  (Pt(2),  Pt(2),  1.0),
}

# ─── Per-doc-type defaults ─────────────────────────────────────────────────────
DOC_TYPE_DEFAULTS = {
    "OFFICIAL_NOTICE": {
        "body_font": "Calibri", "body_size": 11,
        "body_align": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "body_spacing": "normal",
        "title_font": "Calibri", "title_size": 18, "title_align": WD_ALIGN_PARAGRAPH.CENTER,
    },
    "FORMAL_DOCUMENT": {
        "body_font": "Calibri", "body_size": 11,
        "body_align": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "body_spacing": "normal",
        "title_font": "Calibri", "title_size": 14, "title_align": WD_ALIGN_PARAGRAPH.CENTER,
    },
    "PRODUCT_LABEL": {
        "body_font": "Calibri", "body_size": 10,
        "body_align": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "body_spacing": "tight",
        "title_font": "Calibri", "title_size": 14, "title_align": WD_ALIGN_PARAGRAPH.CENTER,
    },
    "TEXTBOOK_SCIENCE": {
        "body_font": "Calibri", "body_size": 11,
        "body_align": WD_ALIGN_PARAGRAPH.LEFT,
        "body_spacing": "normal",
        "title_font": "Calibri", "title_size": 14, "title_align": WD_ALIGN_PARAGRAPH.LEFT,
    },
    "TEXTBOOK_GENERAL": {
        "body_font": "Calibri", "body_size": 11,
        "body_align": WD_ALIGN_PARAGRAPH.JUSTIFY,
        "body_spacing": "normal",
        "title_font": "Calibri", "title_size": 14, "title_align": WD_ALIGN_PARAGRAPH.LEFT,
    },
    "FORM_TABLE": {
        "body_font": "Calibri", "body_size": 11,
        "body_align": WD_ALIGN_PARAGRAPH.LEFT,
        "body_spacing": "tight",
        "title_font": "Calibri", "title_size": 14, "title_align": WD_ALIGN_PARAGRAPH.CENTER,
    },
    "MIXED": {
        "body_font": "Calibri", "body_size": 11,
        "body_align": WD_ALIGN_PARAGRAPH.LEFT,
        "body_spacing": "normal",
        "title_font": "Calibri", "title_size": 14, "title_align": WD_ALIGN_PARAGRAPH.CENTER,
    },
}


@dataclass
class ElementStyle:
    font_name:    str   = "Calibri"
    font_size:    float = 11
    bold:         bool  = False
    italic:       bool  = False
    align:        any   = WD_ALIGN_PARAGRAPH.LEFT
    space_before: any   = Pt(4)
    space_after:  any   = Pt(4)
    line_spacing: float = 1.15
    color:        any   = None


@dataclass
class StyleMap:
    """Complete style map for a document, built from AI analysis + doc type."""
    doc_type:   str = "MIXED"
    base_font:  str = "Calibri"
    base_size:  int = 11
    base_align: any = field(default_factory=lambda: WD_ALIGN_PARAGRAPH.LEFT)

    # Per-role element styles
    styles:     dict = field(default_factory=dict)

    def get(self, role: str) -> ElementStyle:
        return self.styles.get(role, self.styles.get("body", ElementStyle()))


def build_style_map(doc_type: str, style_json: dict) -> StyleMap:
    """
    Build a StyleMap by merging:
      1. Doc-type defaults (always applied)
      2. AI visual analysis (overrides where confident)
    """
    defaults = DOC_TYPE_DEFAULTS.get(doc_type, DOC_TYPE_DEFAULTS["MIXED"])
    sm = StyleMap(doc_type=doc_type)
    sm.base_font  = defaults["body_font"]
    sm.base_size  = defaults["body_size"]
    sm.base_align = defaults["body_align"]

    # ── Default styles per role ────────────────────────────────────────────────
    body_sp_before, body_sp_after, body_line = SPACING_MAP[defaults["body_spacing"]]

    sm.styles = {
        "body": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            align=defaults["body_align"],
            space_before=body_sp_before,
            space_after=body_sp_after,
            line_spacing=body_line,
        ),
        "main_title": ElementStyle(
            font_name=defaults["title_font"],
            font_size=defaults["title_size"],
            bold=True,
            align=defaults["title_align"],
            space_before=Pt(0), space_after=Pt(6),
            line_spacing=1.0,
        ),
        "subtitle": ElementStyle(
            font_name=defaults["title_font"],
            font_size=defaults["title_size"] - 2,
            bold=False,
            align=defaults["title_align"],
            space_before=Pt(2), space_after=Pt(8),
        ),
        "heading": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"] + 3,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            space_before=Pt(10), space_after=Pt(4),
        ),
        "subheading": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"] + 1,
            bold=True,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            space_before=Pt(8), space_after=Pt(3),
        ),
        "ref_line": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            bold=True,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            space_before=Pt(6), space_after=Pt(2),
        ),
        "subject_line": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            bold=True,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            space_before=Pt(6), space_after=Pt(6),
        ),
        "list_item": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            align=WD_ALIGN_PARAGRAPH.LEFT,
            space_before=Pt(1), space_after=Pt(1),
            line_spacing=1.1,
        ),
        "table_header": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        ),
        "table_cell": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            align=WD_ALIGN_PARAGRAPH.LEFT,
        ),
        "label": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            bold=True,
        ),
        "footnote": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"] - 1,
            italic=True,
            space_before=Pt(2), space_after=Pt(2),
        ),
        "signature_name": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"],
            bold=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        ),
        "signature_role": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"] - 1,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        ),
        "struct": ElementStyle(
            font_name="Courier New",
            font_size=10,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            space_before=Pt(4), space_after=Pt(4),
            line_spacing=1.0,
        ),
        "caption": ElementStyle(
            font_name=defaults["body_font"],
            font_size=defaults["body_size"] - 1,
            italic=True,
            align=WD_ALIGN_PARAGRAPH.CENTER,
        ),
    }

    # ── Override with AI visual analysis where available ───────────────────────
    if style_json and "elements" in style_json:
        # Override base font if AI detected it confidently
        ai_font = style_json.get("doc_font_style")
        if ai_font and ai_font in FONT_MAP:
            ai_font_name = FONT_MAP[ai_font]
            # Update all body-type styles
            for role in ("body", "list_item", "ref_line", "subject_line"):
                if role in sm.styles:
                    sm.styles[role].font_name = ai_font_name
            sm.base_font = ai_font_name

        for elem in style_json["elements"]:
            role      = elem.get("role", "")
            font_s    = elem.get("font_style", "")
            size_cat  = elem.get("size_category", "")
            bold      = elem.get("bold")
            italic    = elem.get("italic")
            align_s   = elem.get("align", "")
            spacing_s = elem.get("spacing", "")

            if role not in sm.styles:
                sm.styles[role] = ElementStyle()

            es = sm.styles[role]
            if font_s in FONT_MAP:
                es.font_name = FONT_MAP[font_s]
            if size_cat in SIZE_MAP:
                es.font_size = SIZE_MAP[size_cat]
            if bold is not None:
                es.bold = bold
            if italic is not None:
                es.italic = italic
            if align_s in ALIGN_MAP:
                es.align = ALIGN_MAP[align_s]
            if spacing_s in SPACING_MAP:
                sb, sa, ls = SPACING_MAP[spacing_s]
                es.space_before = sb
                es.space_after  = sa
                es.line_spacing = ls

    return sm
