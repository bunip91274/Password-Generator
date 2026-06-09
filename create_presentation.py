from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn, nsmap
from lxml import etree
import copy


def set_slide_bg(slide, r, g, b):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(r, g, b)


def set_gradient_bg(slide, color1, color2, angle=270):
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_angle = angle * 60000
    stops = fill.gradient_stops
    stops[0].color.rgb = RGBColor(*color1)
    stops[0].position = 0.0
    stops[1].color.rgb = RGBColor(*color2)
    stops[1].position = 1.0


def add_transition(slide, trans_type="fade", duration=700, advance_time=None):
    """Add slide transition via XML manipulation."""
    P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
    P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"

    trans_elem = etree.SubElement(slide._element, f"{{{P_NS}}}transition")
    trans_elem.set("spd", "med")
    trans_elem.set(f"{{{P14_NS}}}dur", str(duration))

    if advance_time:
        trans_elem.set("advTm", str(advance_time))

    if trans_type == "fade":
        child = etree.SubElement(trans_elem, f"{{{P_NS}}}fade")
        child.set("thruBlk", "1")
    elif trans_type == "push":
        child = etree.SubElement(trans_elem, f"{{{P_NS}}}push")
        child.set("dir", "l")
    elif trans_type == "wipe":
        child = etree.SubElement(trans_elem, f"{{{P_NS}}}wipe")
        child.set("dir", "d")
    elif trans_type == "cover":
        child = etree.SubElement(trans_elem, f"{{{P_NS}}}cover")
        child.set("dir", "l")
    elif trans_type == "split":
        child = etree.SubElement(trans_elem, f"{{{P_NS}}}split")
        child.set("orient", "horz")
        child.set("dir", "out")
    elif trans_type == "strips":
        child = etree.SubElement(trans_elem, f"{{{P_NS}}}strips")
        child.set("dir", "rd")


def add_text_box(slide, left, top, width, height, text, font_size=18,
                 bold=False, color=RGBColor(255, 255, 255), alignment=PP_ALIGN.LEFT,
                 font_name="Calibri Light", italic=False):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.font.italic = italic
    p.alignment = alignment
    return tf


def add_paragraph(text_frame, text, font_size=18, bold=False,
                  color=RGBColor(255, 255, 255), alignment=PP_ALIGN.LEFT,
                  font_name="Calibri Light", space_before=Pt(10), italic=False):
    p = text_frame.add_paragraph()
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.font.italic = italic
    p.alignment = alignment
    p.space_before = space_before
    return p


def add_accent_line(slide, left, top, width, color=RGBColor(100, 180, 255), thickness=Pt(3)):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, thickness)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    shape.shadow.inherit = False


def add_decorative_circle(slide, left, top, size, color, alpha=40):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    # Make semi-transparent via XML
    spPr = shape._element.find(qn("p:spPr"))
    solidFill = spPr.find(qn("a:solidFill"))
    if solidFill is not None:
        srgb = solidFill.find(qn("a:srgbClr"))
        if srgb is not None:
            alpha_elem = etree.SubElement(srgb, qn("a:alpha"))
            alpha_elem.set("val", str(alpha * 1000))


# Color palette
BG_DARK = (8, 8, 28)
BG_MID = (12, 14, 38)
BG_SECTION = (16, 18, 48)
BG_BOTTOM = (4, 4, 18)
ACCENT_BLUE = RGBColor(90, 170, 255)
ACCENT_CYAN = RGBColor(80, 220, 230)
ACCENT_GOLD = RGBColor(255, 195, 60)
ACCENT_WARM = RGBColor(255, 140, 60)
TEXT_WHITE = RGBColor(250, 250, 255)
TEXT_LIGHT = RGBColor(210, 215, 230)
TEXT_DIM = RGBColor(150, 155, 180)
SUBTLE_BLUE = RGBColor(30, 40, 80)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ─── Slide 1: Title ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

# Decorative elements
add_decorative_circle(slide, Inches(-1), Inches(-1), Inches(4), ACCENT_BLUE, alpha=8)
add_decorative_circle(slide, Inches(10), Inches(5), Inches(5), ACCENT_GOLD, alpha=6)

add_accent_line(slide, Inches(4.5), Inches(2.3), Inches(4.3), ACCENT_BLUE, Pt(3))

add_text_box(slide, Inches(1), Inches(2.5), Inches(11.3), Inches(1.5),
             "INTERSTELLAR", font_size=60, bold=True, color=TEXT_WHITE,
             alignment=PP_ALIGN.CENTER, font_name="Calibri")
add_text_box(slide, Inches(1), Inches(3.9), Inches(11.3), Inches(0.8),
             "Hard Science Fiction in Film", font_size=26, bold=False,
             color=ACCENT_CYAN, alignment=PP_ALIGN.CENTER)

add_accent_line(slide, Inches(5.5), Inches(5.0), Inches(2.3), ACCENT_BLUE, Pt(2))

add_text_box(slide, Inches(1), Inches(5.3), Inches(11.3), Inches(0.6),
             "Directed by Christopher Nolan  •  2014", font_size=18,
             color=TEXT_DIM, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(6.2), Inches(11.3), Inches(0.6),
             "ISU Genre Study  |  Companion to Project Hail Mary by Andy Weir",
             font_size=14, color=TEXT_DIM, alignment=PP_ALIGN.CENTER, italic=True)

add_transition(slide, "fade", duration=1000)

# ─── Slide 2: What Is Hard Sci-Fi? ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_decorative_circle(slide, Inches(10.5), Inches(-1.5), Inches(4), ACCENT_CYAN, alpha=6)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "What Is Hard Science Fiction?", font_size=34, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.3), Inches(3.5), ACCENT_BLUE)

# Key statement in a box
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(0.8), Inches(1.8), Inches(11.5), Inches(1.2))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(20, 25, 55)
shape.line.color.rgb = ACCENT_GOLD
shape.line.width = Pt(1.5)
tf = shape.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "  The science is the story — not the backdrop."
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = ACCENT_GOLD
p.font.name = "Calibri Light"
p.alignment = PP_ALIGN.CENTER

tf = add_text_box(slide, Inches(1.2), Inches(3.4), Inches(11), Inches(4),
                  "", font_size=12)
add_paragraph(tf, "▸  Soft sci-fi uses spaceships and robots as decoration",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(18))
add_paragraph(tf, "▸  Hard sci-fi: remove the physics or biology → the plot collapses",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(18))
add_paragraph(tf, "▸  Always begins with a plausible \"what if\" premise",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(18))
add_paragraph(tf, "▸  Uses that premise to examine something true about the real world",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(18))

add_transition(slide, "fade", duration=800)

# ─── Slide 3: Section Break ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_SECTION)

add_decorative_circle(slide, Inches(0), Inches(2), Inches(6), ACCENT_GOLD, alpha=5)
add_decorative_circle(slide, Inches(8), Inches(0.5), Inches(6), ACCENT_BLUE, alpha=5)

add_accent_line(slide, Inches(4), Inches(2.6), Inches(5.3), ACCENT_GOLD, Pt(2))

add_text_box(slide, Inches(1), Inches(2.8), Inches(11.3), Inches(1.2),
             "THREE HALLMARKS", font_size=46, bold=True,
             color=ACCENT_GOLD, alignment=PP_ALIGN.CENTER, font_name="Calibri")
add_text_box(slide, Inches(1), Inches(4.1), Inches(11.3), Inches(0.7),
             "OF HARD SCIENCE FICTION", font_size=22, bold=False,
             color=TEXT_DIM, alignment=PP_ALIGN.CENTER)

add_accent_line(slide, Inches(4), Inches(5.0), Inches(5.3), ACCENT_GOLD, Pt(2))

add_text_box(slide, Inches(1.5), Inches(5.3), Inches(10.3), Inches(0.8),
             "Each demonstrated in Interstellar — with one additional genre example",
             font_size=16, color=TEXT_DIM, alignment=PP_ALIGN.CENTER, italic=True)

add_transition(slide, "wipe", duration=900)

# ─── Slide 4: Three Hallmarks Overview ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Three Hallmarks — Overview", font_size=34, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.3), Inches(3.5), ACCENT_BLUE)

hallmarks = [
    ("01", "Science as Plot Engine", "The science drives every plot event — remove it and the story collapses"),
    ("02", "Isolated Protagonist as Social Mirror", "Strips away civilization to ask what it was actually for"),
    ("03", "\"What If\" Premise as Social Critique", "Uses a plausible premise to interrogate the present world"),
]

for i, (num, title, desc) in enumerate(hallmarks):
    y = 2.0 + i * 1.75
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(1.0), Inches(y), Inches(11.3), Inches(1.45))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(18, 22, 55)
    shape.line.color.rgb = ACCENT_BLUE
    shape.line.width = Pt(1.2)

    # Number accent
    num_box = add_text_box(slide, Inches(1.3), Inches(y + 0.15), Inches(0.9), Inches(1.0),
                           num, font_size=32, bold=True, color=ACCENT_GOLD,
                           alignment=PP_ALIGN.CENTER, font_name="Calibri")

    # Title
    add_text_box(slide, Inches(2.4), Inches(y + 0.15), Inches(9.5), Inches(0.6),
                 title, font_size=22, bold=True, color=TEXT_WHITE, font_name="Calibri")

    # Description
    add_text_box(slide, Inches(2.4), Inches(y + 0.75), Inches(9.5), Inches(0.6),
                 desc, font_size=16, color=TEXT_DIM)

add_transition(slide, "fade", duration=800)

# ─── Slide 5: Hallmark 1 — Science as Plot Engine ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_decorative_circle(slide, Inches(11), Inches(-1), Inches(3.5), ACCENT_BLUE, alpha=7)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(9), Inches(0.8),
             "Hallmark 1: Science as Plot Engine", font_size=30, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.15), Inches(6), ACCENT_BLUE, Pt(2))

# Interstellar section
add_text_box(slide, Inches(0.8), Inches(1.5), Inches(5), Inches(0.5),
             "INTERSTELLAR", font_size=16, bold=True, color=ACCENT_GOLD,
             font_name="Calibri")

tf = add_text_box(slide, Inches(0.8), Inches(2.0), Inches(7.5), Inches(4.5),
                  "", font_size=12)
add_paragraph(tf, "▸  Time dilation near Gargantua = applied general relativity",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "▸  1 hour on Miller's Planet = 7 years on Earth",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "▸  Cooper doesn't choose to lose his daughter — the physics does",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "▸  Kip Thorne (Nobel laureate) as scientific advisor",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "▸  First scientifically accurate black hole simulation in film",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))

# Genre parallel - right panel
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(8.6), Inches(1.5), Inches(4.3), Inches(3.8))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(20, 25, 50)
shape.line.color.rgb = ACCENT_GOLD
shape.line.width = Pt(1.2)

add_text_box(slide, Inches(9.0), Inches(1.8), Inches(3.6), Inches(0.5),
             "GENRE PARALLEL", font_size=12, bold=True, color=TEXT_DIM,
             alignment=PP_ALIGN.CENTER, font_name="Calibri")
add_text_box(slide, Inches(9.0), Inches(2.3), Inches(3.6), Inches(0.5),
             "The Martian", font_size=20, bold=True, color=ACCENT_GOLD,
             alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(9.0), Inches(2.9), Inches(3.6), Inches(0.4),
             "Andy Weir", font_size=14, color=TEXT_DIM,
             alignment=PP_ALIGN.CENTER, italic=True)

tf = add_text_box(slide, Inches(9.0), Inches(3.5), Inches(3.8), Inches(1.5),
                  "", font_size=12)
add_paragraph(tf, "Watney survives using real chemistry (hydrazine → water)",
              font_size=15, color=TEXT_LIGHT, space_before=Pt(6))
add_paragraph(tf, "The plot stops the second the science stops",
              font_size=15, color=TEXT_LIGHT, space_before=Pt(10))

add_transition(slide, "push", duration=800)

# ─── Slide 6: Hallmark 2 — Isolated Protagonist ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_decorative_circle(slide, Inches(-1.5), Inches(4), Inches(4), ACCENT_CYAN, alpha=6)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(9), Inches(0.8),
             "Hallmark 2: The Isolated Protagonist", font_size=30, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.15), Inches(6), ACCENT_BLUE, Pt(2))

add_text_box(slide, Inches(0.8), Inches(1.5), Inches(5), Inches(0.5),
             "INTERSTELLAR", font_size=16, bold=True, color=ACCENT_GOLD,
             font_name="Calibri")

tf = add_text_box(slide, Inches(0.8), Inches(2.0), Inches(7.5), Inches(4.5),
                  "", font_size=12)
add_paragraph(tf, "▸  Cooper separated from his family for decades by physics itself",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "▸  Strips away every support system civilization provides",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "▸  Then asks: what was that civilization actually for?",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "", font_size=8)
add_paragraph(tf, "Central question:", font_size=17, bold=True, color=ACCENT_CYAN,
              space_before=Pt(14))
add_paragraph(tf, "What does a parent owe a child when humanity itself is at stake?",
              font_size=19, color=TEXT_WHITE, italic=True, space_before=Pt(6))

# Genre parallel - right panel
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(8.6), Inches(1.5), Inches(4.3), Inches(3.8))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(20, 25, 50)
shape.line.color.rgb = ACCENT_GOLD
shape.line.width = Pt(1.2)

add_text_box(slide, Inches(9.0), Inches(1.8), Inches(3.6), Inches(0.5),
             "GENRE PARALLEL", font_size=12, bold=True, color=TEXT_DIM,
             alignment=PP_ALIGN.CENTER, font_name="Calibri")
add_text_box(slide, Inches(9.0), Inches(2.3), Inches(3.6), Inches(0.5),
             "Contact (1997)", font_size=20, bold=True, color=ACCENT_GOLD,
             alignment=PP_ALIGN.CENTER)

tf = add_text_box(slide, Inches(9.0), Inches(3.0), Inches(3.8), Inches(2.0),
                  "", font_size=12)
add_paragraph(tf, "Ellie Arroway — only human to make first contact",
              font_size=15, color=TEXT_LIGHT, space_before=Pt(6))
add_paragraph(tf, "Isolation becomes a lens for science vs. faith, individual vs. institution",
              font_size=15, color=TEXT_LIGHT, space_before=Pt(10))

add_transition(slide, "push", duration=800)

# ─── Slide 7: Hallmark 3 — The What If Premise ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_decorative_circle(slide, Inches(9), Inches(4.5), Inches(4), ACCENT_GOLD, alpha=6)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(9), Inches(0.8),
             "Hallmark 3: The \"What If\" Premise", font_size=30, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.15), Inches(6), ACCENT_BLUE, Pt(2))

add_text_box(slide, Inches(0.8), Inches(1.5), Inches(5), Inches(0.5),
             "INTERSTELLAR", font_size=16, bold=True, color=ACCENT_GOLD,
             font_name="Calibri")

tf = add_text_box(slide, Inches(0.8), Inches(2.0), Inches(7.5), Inches(4.5),
                  "", font_size=12)
add_paragraph(tf, "▸  What if Earth became uninhabitable and humanity had one",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "    launch window?",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(2))
add_paragraph(tf, "▸  Forces the question: who gets to decide who lives?",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "▸  Brand's monologue: love as the only force that transcends",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(8))
add_paragraph(tf, "    dimensions — the film's emotional answer & political argument",
              font_size=17, color=TEXT_LIGHT, space_before=Pt(2))

# Genre parallel - right panel
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(8.6), Inches(1.5), Inches(4.3), Inches(3.8))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(20, 25, 50)
shape.line.color.rgb = ACCENT_GOLD
shape.line.width = Pt(1.2)

add_text_box(slide, Inches(9.0), Inches(1.8), Inches(3.6), Inches(0.5),
             "GENRE PARALLEL", font_size=12, bold=True, color=TEXT_DIM,
             alignment=PP_ALIGN.CENTER, font_name="Calibri")
add_text_box(slide, Inches(9.0), Inches(2.3), Inches(3.6), Inches(0.5),
             "Arrival (2016)", font_size=20, bold=True, color=ACCENT_GOLD,
             alignment=PP_ALIGN.CENTER)

tf = add_text_box(slide, Inches(9.0), Inches(3.0), Inches(3.8), Inches(2.0),
                  "", font_size=12)
add_paragraph(tf, "What if aliens arrived and the only tool was language?",
              font_size=15, color=TEXT_LIGHT, space_before=Pt(6))
add_paragraph(tf, "Forces questions about communication, determinism, and choice",
              font_size=15, color=TEXT_LIGHT, space_before=Pt(10))

add_transition(slide, "push", duration=800)

# ─── Slide 8: Section Break — The Review ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_SECTION)

add_decorative_circle(slide, Inches(5), Inches(1.5), Inches(7), ACCENT_WARM, alpha=4)
add_decorative_circle(slide, Inches(-2), Inches(3), Inches(5), ACCENT_BLUE, alpha=5)

add_accent_line(slide, Inches(4.5), Inches(2.8), Inches(4.3), ACCENT_WARM, Pt(2))

add_text_box(slide, Inches(1), Inches(3.0), Inches(11.3), Inches(1.2),
             "THE REVIEW", font_size=50, bold=True,
             color=ACCENT_GOLD, alignment=PP_ALIGN.CENTER, font_name="Calibri")
add_text_box(slide, Inches(1.5), Inches(4.4), Inches(10.3), Inches(0.7),
             "Plot Synopsis  &  Evaluation", font_size=20,
             color=TEXT_DIM, alignment=PP_ALIGN.CENTER)

add_accent_line(slide, Inches(4.5), Inches(5.3), Inches(4.3), ACCENT_WARM, Pt(2))

add_transition(slide, "wipe", duration=900)

# ─── Slide 9: Synopsis ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.8),
             "Plot Synopsis", font_size=34, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.15), Inches(3), ACCENT_BLUE, Pt(2))

tf = add_text_box(slide, Inches(1.0), Inches(1.6), Inches(11.3), Inches(3.5),
                  "", font_size=12)
add_paragraph(tf, "▸  Earth is dying from crop blight",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(14))
add_paragraph(tf, "▸  Former NASA pilot Cooper recruited for interstellar mission",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(14))
add_paragraph(tf, "▸  Travels through a wormhole near Saturn → another galaxy",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(14))
add_paragraph(tf, "▸  Time dilation, black holes, gravitational physics, and wormhole",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(14))
add_paragraph(tf, "    mechanics drive every major plot event",
              font_size=20, color=TEXT_LIGHT, space_before=Pt(2))

# Bottom quote
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(1.0), Inches(5.5), Inches(11.3), Inches(1.4))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(20, 22, 50)
shape.line.color.rgb = ACCENT_GOLD
shape.line.width = Pt(1.2)
tf = shape.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "  Cooper loses decades with his daughter not because he chose to —"
p.font.size = Pt(18)
p.font.bold = False
p.font.italic = True
p.font.color.rgb = ACCENT_GOLD
p.font.name = "Calibri Light"
p.alignment = PP_ALIGN.CENTER
p2 = tf.add_paragraph()
p2.text = "  but because the universe operates on rules that don't care about feelings."
p2.font.size = Pt(18)
p2.font.bold = False
p2.font.italic = True
p2.font.color.rgb = ACCENT_GOLD
p2.font.name = "Calibri Light"
p2.alignment = PP_ALIGN.CENTER

add_transition(slide, "fade", duration=800)

# ─── Slide 10: Quote ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_SECTION)

add_decorative_circle(slide, Inches(0), Inches(0), Inches(5), ACCENT_GOLD, alpha=4)
add_decorative_circle(slide, Inches(9), Inches(4), Inches(5), ACCENT_BLUE, alpha=4)

# Large quotation mark visual
add_text_box(slide, Inches(1.5), Inches(0.8), Inches(2), Inches(1.5),
             "\u201C", font_size=100, color=ACCENT_GOLD, bold=True,
             alignment=PP_ALIGN.LEFT, font_name="Georgia")

add_text_box(slide, Inches(2.0), Inches(1.8), Inches(9.3), Inches(2.5),
             "Love is the one thing we're capable of\nperceiving that transcends\ndimensions of time and space.",
             font_size=28, bold=False, color=TEXT_WHITE, alignment=PP_ALIGN.CENTER,
             italic=True, font_name="Georgia")

add_text_box(slide, Inches(2.0), Inches(4.2), Inches(9.3), Inches(0.5),
             "— Dr. Amelia Brand", font_size=16, color=TEXT_DIM,
             alignment=PP_ALIGN.CENTER, italic=True)

add_accent_line(slide, Inches(5), Inches(5.0), Inches(3.3), ACCENT_GOLD, Pt(1.5))

tf = add_text_box(slide, Inches(1.5), Inches(5.3), Inches(10.3), Inches(1.8),
                  "The film's most debated line. It sits in tension with everything the science",
                  font_size=16, color=TEXT_LIGHT, alignment=PP_ALIGN.CENTER)
add_paragraph(tf, "has told us. Nolan is asking whether reason alone is enough —",
              font_size=16, color=TEXT_LIGHT, alignment=PP_ALIGN.CENTER, space_before=Pt(4))
add_paragraph(tf, "exactly the kind of question hard sci-fi is built to ask.",
              font_size=16, color=TEXT_LIGHT, alignment=PP_ALIGN.CENTER, space_before=Pt(4))

add_transition(slide, "fade", duration=1000)

# ─── Slide 11: Clip / Key Scene ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.8),
             "Key Scene: Miller's Planet", font_size=34, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.15), Inches(4.5), ACCENT_BLUE, Pt(2))

# Timestamp badge
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(0.8), Inches(1.6), Inches(5.5), Inches(0.7))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(25, 28, 55)
shape.line.color.rgb = ACCENT_CYAN
shape.line.width = Pt(1)
tf = shape.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = "    \u25B6  Timestamp: ~1:05:00   |   Duration: ~5 minutes"
p.font.size = Pt(16)
p.font.color.rgb = ACCENT_CYAN
p.font.name = "Calibri Light"

add_text_box(slide, Inches(0.8), Inches(2.7), Inches(5), Inches(0.5),
             "WHAT TO WATCH FOR", font_size=15, bold=True, color=ACCENT_GOLD,
             font_name="Calibri")

tf = add_text_box(slide, Inches(0.8), Inches(3.2), Inches(11.5), Inches(2.5),
                  "", font_size=12)
add_paragraph(tf, "▸  Cooper and Brand land on a water world in Gargantua's gravitational field",
              font_size=18, color=TEXT_LIGHT, space_before=Pt(10))
add_paragraph(tf, "▸  Every hour on the surface = 7 years on Earth",
              font_size=18, color=TEXT_LIGHT, space_before=Pt(10))
add_paragraph(tf, "▸  When they return, Romilly has aged 23 years waiting for them",
              font_size=18, color=TEXT_LIGHT, space_before=Pt(10))

# Bottom quote
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(1.5), Inches(5.6), Inches(10.3), Inches(1.2))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(20, 22, 50)
shape.line.color.rgb = ACCENT_GOLD
shape.line.width = Pt(1.5)
tf = shape.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
tf.word_wrap = True
p = tf.paragraphs[0]
p.text = "  \"The science is not decorating the tragedy — it IS the tragedy.\""
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = ACCENT_GOLD
p.font.name = "Calibri Light"
p.alignment = PP_ALIGN.CENTER

add_transition(slide, "cover", duration=800)

# ─── Slide 12: Rating ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_decorative_circle(slide, Inches(10), Inches(-1), Inches(4), ACCENT_GOLD, alpha=5)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.8),
             "Evaluation", font_size=34, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT, font_name="Calibri")
add_accent_line(slide, Inches(0.8), Inches(1.15), Inches(2.5), ACCENT_BLUE, Pt(2))

ratings = [
    ("Scientific Accuracy", "9/10", "Thorne's physics are real; ending stretches deliberately"),
    ("\"What If\" Quality", "8/10", "Genuinely plausible, more immediate than most sci-fi"),
    ("Science as Plot Engine", "8/10", "Third act requires understanding gravitational physics"),
    ("Genre Hallmarks", "9/10", "Isolated protagonist, social critique, plausible premise"),
    ("Watchability", "9/10", "Visually stunning, emotionally driven, scientifically serious"),
]

for i, (category, score, note) in enumerate(ratings):
    y = 1.6 + i * 1.0

    # Row background (alternating subtle difference)
    row_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                       Inches(0.6), Inches(y), Inches(12.0), Inches(0.8))
    row_shape.fill.solid()
    if i % 2 == 0:
        row_shape.fill.fore_color.rgb = RGBColor(16, 18, 45)
    else:
        row_shape.fill.fore_color.rgb = RGBColor(12, 14, 38)
    row_shape.line.fill.background()

    # Category
    add_text_box(slide, Inches(0.9), Inches(y + 0.1), Inches(3.8), Inches(0.6),
                 category, font_size=17, bold=True, color=TEXT_WHITE, font_name="Calibri")

    # Score with gold highlight
    score_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                         Inches(4.9), Inches(y + 0.1), Inches(1.2), Inches(0.55))
    score_shape.fill.solid()
    score_shape.fill.fore_color.rgb = RGBColor(40, 35, 15)
    score_shape.line.color.rgb = ACCENT_GOLD
    score_shape.line.width = Pt(1)
    stf = score_shape.text_frame
    stf.vertical_anchor = MSO_ANCHOR.MIDDLE
    sp = stf.paragraphs[0]
    sp.text = score
    sp.font.size = Pt(16)
    sp.font.bold = True
    sp.font.color.rgb = ACCENT_GOLD
    sp.font.name = "Calibri"
    sp.alignment = PP_ALIGN.CENTER

    # Note
    add_text_box(slide, Inches(6.4), Inches(y + 0.1), Inches(6.2), Inches(0.6),
                 note, font_size=15, color=TEXT_DIM)

# Final verdict
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(2.5), Inches(6.5), Inches(8.3), Inches(0.75))
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(25, 22, 10)
shape.line.color.rgb = ACCENT_GOLD
shape.line.width = Pt(2)
tf = shape.text_frame
tf.vertical_anchor = MSO_ANCHOR.MIDDLE
p = tf.paragraphs[0]
p.text = "One of the best hard sci-fi films ever made."
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = ACCENT_GOLD
p.font.name = "Calibri Light"
p.alignment = PP_ALIGN.CENTER

add_transition(slide, "split", duration=900)

# Register p14 namespace for transitions
etree.register_namespace("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main")

# Save
prs.save("/workspace/Interstellar_Hard_SciFi_Presentation.pptx")
print("Presentation saved successfully!")
print(f"Slides: {len(prs.slides)}")
