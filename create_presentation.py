from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import copy


def set_slide_bg(slide, r, g, b):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(r, g, b)


def add_text_box(slide, left, top, width, height, text, font_size=18,
                 bold=False, color=RGBColor(255, 255, 255), alignment=PP_ALIGN.LEFT,
                 font_name="Calibri"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = alignment
    return tf


def add_paragraph(text_frame, text, font_size=18, bold=False,
                  color=RGBColor(255, 255, 255), alignment=PP_ALIGN.LEFT,
                  font_name="Calibri", space_before=Pt(12)):
    p = text_frame.add_paragraph()
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = alignment
    p.space_before = space_before
    return p


# Color palette
BG_DARK = (10, 10, 30)
BG_SECTION = (20, 20, 50)
ACCENT_BLUE = RGBColor(100, 180, 255)
ACCENT_GOLD = RGBColor(255, 200, 80)
TEXT_WHITE = RGBColor(255, 255, 255)
TEXT_LIGHT = RGBColor(220, 220, 230)
TEXT_DIM = RGBColor(170, 170, 190)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ─── Slide 1: Title ───
slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(1), Inches(1.5), Inches(11), Inches(1.2),
             "INTERSTELLAR", font_size=54, bold=True, color=ACCENT_BLUE,
             alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(2.8), Inches(11), Inches(0.8),
             "Hard Science Fiction in Film", font_size=28, bold=False,
             color=TEXT_LIGHT, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(4.0), Inches(11), Inches(0.6),
             "Directed by Christopher Nolan (2014)", font_size=20,
             color=TEXT_DIM, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1), Inches(5.2), Inches(11), Inches(0.8),
             "ISU Genre Study — Companion to Project Hail Mary by Andy Weir",
             font_size=16, color=TEXT_DIM, alignment=PP_ALIGN.CENTER)

# ─── Slide 2: What Is Hard Sci-Fi? ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "What Is Hard Science Fiction?", font_size=36, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

tf = add_text_box(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(5),
                  "The science is the story — not the backdrop.", font_size=24,
                  bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "", font_size=12)
add_paragraph(tf, "• Soft sci-fi uses spaceships and robots as decoration",
              font_size=20, color=TEXT_LIGHT)
add_paragraph(tf, "• Hard sci-fi: remove the physics or biology, and the plot collapses",
              font_size=20, color=TEXT_LIGHT)
add_paragraph(tf, "• Always starts with a plausible \"what if\" premise",
              font_size=20, color=TEXT_LIGHT)
add_paragraph(tf, "• Uses that premise to examine something true about the real world",
              font_size=20, color=TEXT_LIGHT)

# ─── Slide 3: Section Break ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_SECTION)

add_text_box(slide, Inches(1), Inches(2.8), Inches(11), Inches(1.0),
             "THREE HALLMARKS OF HARD SCI-FI", font_size=40, bold=True,
             color=ACCENT_GOLD, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1.5), Inches(4.2), Inches(10), Inches(1.0),
             "Each demonstrated in Interstellar — with one additional genre example",
             font_size=20, color=TEXT_DIM, alignment=PP_ALIGN.CENTER)

# ─── Slide 4: Three Hallmarks Overview ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Three Hallmarks — Overview", font_size=36, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

# Hallmark boxes
for i, (num, title) in enumerate([
    ("1", "Science as Plot Engine"),
    ("2", "The Isolated Protagonist as Social Mirror"),
    ("3", "The \"What If\" Premise as Social Critique")
]):
    y = 2.0 + i * 1.7
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Inches(1.2), Inches(y), Inches(10.5), Inches(1.3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(30, 30, 70)
    shape.line.color.rgb = ACCENT_BLUE
    shape.line.width = Pt(1.5)
    tf = shape.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.text = f"  {num}.  {title}"
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = TEXT_WHITE
    p.font.name = "Calibri"

# ─── Slide 5: Hallmark 1 — Science as Plot Engine ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.9),
             "Hallmark 1: Science as Plot Engine", font_size=32, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

tf = add_text_box(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(5.5),
                  "Interstellar", font_size=22, bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "• Time dilation near Gargantua is applied special relativity",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• 1 hour on Miller's Planet = 7 years on Earth",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Cooper doesn't choose to lose his daughter — the physics does it for him",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Kip Thorne (Nobel laureate) served as scientific advisor",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Gargantua: first scientifically accurate black hole simulation in film",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "", font_size=10)
add_paragraph(tf, "Genre Parallel: The Martian (Andy Weir)", font_size=22,
              bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "• Mark Watney survives using real chemistry (hydrazine → water)",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• The plot stops the second the science stops",
              font_size=18, color=TEXT_LIGHT)

# ─── Slide 6: Hallmark 2 — Isolated Protagonist ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.9),
             "Hallmark 2: The Isolated Protagonist", font_size=32, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

tf = add_text_box(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(5.5),
                  "Interstellar", font_size=22, bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "• Cooper is separated from his family for decades by the physics of space",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Strips away every support system civilization provides",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Central question: What does a parent owe a child when humanity is at stake?",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "", font_size=10)
add_paragraph(tf, "Genre Parallel: Contact (1997)", font_size=22,
              bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "• Ellie Arroway — only human to make first contact",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Isolation becomes a lens for science vs. faith, individual vs. institution",
              font_size=18, color=TEXT_LIGHT)

# ─── Slide 7: Hallmark 3 — The What If Premise ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.9),
             "Hallmark 3: The \"What If\" Premise", font_size=32, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

tf = add_text_box(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(5.5),
                  "Interstellar", font_size=22, bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "• What if Earth became uninhabitable and humanity had one launch window?",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Forces the question: who gets to decide who lives?",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Brand's monologue — love as the only force transcending dimensions",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• The film's emotional answer and political argument",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "", font_size=10)
add_paragraph(tf, "Genre Parallel: Arrival (2016)", font_size=22,
              bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "• What if aliens arrived and the only tool was language?",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Forces questions about communication, determinism, and choice",
              font_size=18, color=TEXT_LIGHT)

# ─── Slide 8: Section Break — The Review ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_SECTION)

add_text_box(slide, Inches(1), Inches(2.5), Inches(11), Inches(1.0),
             "THE REVIEW", font_size=44, bold=True,
             color=ACCENT_GOLD, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1.5), Inches(4.0), Inches(10), Inches(1.0),
             "Plot Synopsis & Evaluation", font_size=22,
             color=TEXT_DIM, alignment=PP_ALIGN.CENTER)

# ─── Slide 9: Synopsis ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.9),
             "Plot Synopsis", font_size=36, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

tf = add_text_box(slide, Inches(0.8), Inches(1.6), Inches(11.5), Inches(5.5),
                  "• Earth is dying from crop blight", font_size=20, color=TEXT_LIGHT)
add_paragraph(tf, "• Former NASA pilot Cooper recruited for interstellar mission",
              font_size=20, color=TEXT_LIGHT)
add_paragraph(tf, "• Travels through a wormhole near Saturn to find a habitable planet",
              font_size=20, color=TEXT_LIGHT)
add_paragraph(tf, "• Time dilation, black holes, gravitational physics, and wormhole mechanics drive every major plot event",
              font_size=20, color=TEXT_LIGHT)
add_paragraph(tf, "", font_size=10)
add_paragraph(tf, "Cooper loses decades with his daughter not because he chose to,\nbut because the universe operates on rules that don't care about feelings.",
              font_size=20, bold=True, color=ACCENT_GOLD)

# ─── Slide 10: Quote ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_SECTION)

add_text_box(slide, Inches(1.5), Inches(1.5), Inches(10), Inches(3.0),
             "\"Love is the one thing we're capable of\nperceiving that transcends\ndimensions of time and space.\"",
             font_size=30, bold=True, color=ACCENT_GOLD, alignment=PP_ALIGN.CENTER)
add_text_box(slide, Inches(1.5), Inches(4.3), Inches(10), Inches(0.6),
             "— Dr. Amelia Brand", font_size=18, color=TEXT_DIM,
             alignment=PP_ALIGN.CENTER)

tf = add_text_box(slide, Inches(1.2), Inches(5.2), Inches(10.5), Inches(2.0),
                  "The film's most debated line. It sits in tension with everything the science has told us — Nolan is asking whether reason alone is enough.",
                  font_size=18, color=TEXT_LIGHT, alignment=PP_ALIGN.CENTER)

# ─── Slide 11: Clip ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.9),
             "Key Scene: Miller's Planet", font_size=36, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

tf = add_text_box(slide, Inches(0.8), Inches(1.6), Inches(11.5), Inches(5.5),
                  "Timestamp: ~1:05:00  |  Duration: ~5 minutes", font_size=18,
                  bold=True, color=TEXT_DIM)
add_paragraph(tf, "", font_size=10)
add_paragraph(tf, "What to watch for:", font_size=22, bold=True, color=ACCENT_GOLD)
add_paragraph(tf, "• Cooper and Brand land on a water world in Gargantua's gravitational field",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• Every hour on the surface = 7 years on Earth",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "• When they return, Romilly has aged 23 years waiting",
              font_size=18, color=TEXT_LIGHT)
add_paragraph(tf, "", font_size=10)
add_paragraph(tf, "\"The science is not decorating the tragedy — it IS the tragedy.\"",
              font_size=22, bold=True, color=ACCENT_GOLD)

# ─── Slide 12: Rating ───
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, *BG_DARK)

add_text_box(slide, Inches(0.8), Inches(0.4), Inches(11), Inches(0.9),
             "Evaluation", font_size=36, bold=True,
             color=ACCENT_BLUE, alignment=PP_ALIGN.LEFT)

ratings = [
    ("Scientific Accuracy", "9 / 10", "Kip Thorne's physics are real; ending stretches deliberately"),
    ("\"What If\" Quality", "8 / 10", "Genuinely plausible, more immediate than most sci-fi"),
    ("Science as Plot Engine", "8 / 10", "Third act requires understanding gravitational physics"),
    ("Genre Hallmarks", "9 / 10", "Isolated protagonist, social critique, plausible premise"),
    ("Watchability", "9 / 10", "Visually stunning, emotionally driven, scientifically serious"),
]

for i, (category, score, note) in enumerate(ratings):
    y = 1.6 + i * 1.1
    # Category label
    add_text_box(slide, Inches(0.8), Inches(y), Inches(4.0), Inches(0.5),
                 category, font_size=18, bold=True, color=TEXT_WHITE)
    # Score
    add_text_box(slide, Inches(5.0), Inches(y), Inches(1.5), Inches(0.5),
                 score, font_size=18, bold=True, color=ACCENT_GOLD,
                 alignment=PP_ALIGN.CENTER)
    # Note
    add_text_box(slide, Inches(6.8), Inches(y), Inches(6.0), Inches(0.5),
                 note, font_size=15, color=TEXT_DIM)

add_text_box(slide, Inches(0.8), Inches(6.6), Inches(11.5), Inches(0.6),
             "One of the best hard sci-fi films ever made.",
             font_size=22, bold=True, color=ACCENT_GOLD, alignment=PP_ALIGN.CENTER)

# Save
prs.save("/workspace/Interstellar_Hard_SciFi_Presentation.pptx")
print("Presentation saved successfully!")
