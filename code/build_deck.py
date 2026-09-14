from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image
import os

IMG = "/Users/sr320/GitHub/talk-pcsga-2026/images"
OUT = "/Users/sr320/GitHub/talk-pcsga-2026/Future-Shellf-PCSGA-2026.pptx"

NAVY   = RGBColor(0x1B, 0x3A, 0x5C)
TEAL   = RGBColor(0x0E, 0x7C, 0x86)
PINK   = RGBColor(0xC2, 0x18, 0x7A)
GRAY   = RGBColor(0x44, 0x4A, 0x52)
LIGHT  = RGBColor(0xEF, 0xF3, 0xF6)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height

BLANK = prs.slide_layouts[6]

def slide():
    return prs.slides.add_slide(BLANK)

def textbox(s, l, t, w, h):
    tb = s.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    return tf

def title(s, text, sub=None):
    tf = textbox(s, 0.7, 0.45, 11.9, 1.0)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(34); p.font.bold = True; p.font.color.rgb = NAVY
    # accent bar
    bar = s.shapes.add_shape(1, Inches(0.7), Inches(1.45), Inches(1.6), Inches(0.06))
    bar.fill.solid(); bar.fill.fore_color.rgb = TEAL; bar.line.fill.background()
    bar.shadow.inherit = False
    if sub:
        tf2 = textbox(s, 0.7, 1.6, 11.9, 0.5)
        p2 = tf2.paragraphs[0]
        p2.text = sub
        p2.font.size = Pt(18); p2.font.italic = True; p2.font.color.rgb = TEAL
    return s

def bullets(s, items, left=0.85, top=2.15, width=11.6, height=4.6, size=22, gap=12):
    tf = textbox(s, left, top, width, height)
    first = True
    for it in items:
        if isinstance(it, tuple):
            text, lvl = it
        else:
            text, lvl = it, 0
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = "• " + text if lvl == 0 else "– " + text
        p.level = lvl
        p.font.size = Pt(size if lvl == 0 else size - 4)
        p.font.color.rgb = GRAY if lvl else RGBColor(0x22,0x26,0x2B)
        p.space_after = Pt(gap)
        if lvl:
            p.space_before = Pt(2)
    return tf

def picture(s, name, left, top, max_w, max_h):
    path = os.path.join(IMG, name)
    iw, ih = Image.open(path).size
    ar = iw / ih
    w, h = max_w, max_w / ar
    if h > max_h:
        h = max_h; w = max_h * ar
    l = left + (max_w - w) / 2
    t = top + (max_h - h) / 2
    s.shapes.add_picture(path, Inches(l), Inches(t), Inches(w), Inches(h))

def place(s, name, left, top, height):
    """Place a figure at an exact position, scaled to a fixed height."""
    path = os.path.join(IMG, name)
    iw, ih = Image.open(path).size
    s.shapes.add_picture(path, Inches(left), Inches(top),
                         Inches(height * iw / ih), Inches(height))

def section(text, sub=None):
    s = slide()
    bg = s.shapes.add_shape(1, 0, 0, SW, SH)
    bg.fill.solid(); bg.fill.fore_color.rgb = NAVY; bg.line.fill.background()
    bg.shadow.inherit = False
    tf = textbox(s, 1.2, 2.9, 11.0, 1.6)
    p = tf.paragraphs[0]; p.text = text
    p.font.size = Pt(40); p.font.bold = True; p.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
    if sub:
        p2 = tf.add_paragraph(); p2.text = sub
        p2.font.size = Pt(20); p2.font.color.rgb = RGBColor(0x9F,0xD3,0xD8)
        p2.space_before = Pt(10)
    return s

def notes(s, text):
    s.notes_slide.notes_text_frame.text = text

def callout(s, text, top=6.15, color=TEAL):
    box = s.shapes.add_shape(1, Inches(0.85), Inches(top), Inches(11.6), Inches(0.85))
    box.fill.solid(); box.fill.fore_color.rgb = LIGHT; box.line.color.rgb = color
    box.line.width = Pt(1.25); box.shadow.inherit = False
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text; p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(19); p.font.bold = True; p.font.color.rgb = NAVY

# ---------------------------------------------------------------- 1 TITLE
s = slide()
band = s.shapes.add_shape(1, 0, 0, SW, Inches(3.35))
band.fill.solid(); band.fill.fore_color.rgb = NAVY; band.line.fill.background()
band.shadow.inherit = False
tf = textbox(s, 0.9, 1.05, 11.6, 1.6)
p = tf.paragraphs[0]; p.text = "Future Shellf"
p.font.size = Pt(54); p.font.bold = True; p.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
p2 = tf.add_paragraph(); p2.text = "Predictive Phenotyping for Oyster Performance"
p2.font.size = Pt(30); p2.font.color.rgb = RGBColor(0x9F,0xD3,0xD8)
p2.space_before = Pt(8)

tf = textbox(s, 0.9, 4.14, 11.6, 1.84)
p = tf.paragraphs[0]
p.text = "Steven Roberts  ·  Ariana Huffmyer  ·  Louis Plough  ·  Mackenzie Gavery  ·  Neil Thompson"
p.font.size = Pt(20); p.font.color.rgb = GRAY
p3 = tf.add_paragraph()
p3.text = "University of Washington  ·  USDA-ARS  ·  NOAA NWFSC"
p3.font.size = Pt(15); p3.font.italic = True; p3.font.color.rgb = TEAL; p3.space_before = Pt(10)
p4 = tf.add_paragraph()
p4.text = "PCSGA Annual Conference · 2026\v\vSORMI.SCIENCE"
p4.font.size = Pt(15); p4.font.color.rgb = GRAY; p4.space_before = Pt(18)
notes(s, "12-minute talk. Two things growers can use: predict phenotype before you grow it out, and read stress on the farm in real time.")

# ---------------------------------------------------------------- 2 WHY
s = slide(); title(s, "Two questions")
bullets(s, [
    "Which seed should I buy — which family lines actually perform on my site?",
    "How stressed is my crop right now, before I start seeing mortality?",
], top=2.3, size=26, gap=26)
tf = textbox(s, 0.85, 3.9, 11.6, 2.0)
p = tf.paragraphs[0]
p.text = "Today, answering either one means waiting."
p.font.size = Pt(24); p.font.bold = True; p.font.color.rgb = NAVY
p2 = tf.add_paragraph(); p2.text = "Breeding trials take a season. Mortality events tell you after the fact."
p2.font.size = Pt(20); p2.font.color.rgb = GRAY; p2.space_before = Pt(8)
callout(s, "We need a fast, cheap, non-lethal measure of how an oyster is doing — right now.")
notes(s, "Set up the grower problem before introducing the tool.")

# ---------------------------------------------------------------- 3 ASSAY INTRO
s = slide(); title(s, "From blue to pink: the resazurin assay",
                   "A redox dye that reports whole-organism metabolism")
picture(s, "resazurin_plates_blue_to_pink.png", 0.85, 2.15, 7.0, 3.9)
tf = textbox(s, 8.15, 2.3, 4.3, 3.8)
lines = [
    ("Resazurin is blue and non-fluorescent.", True),
    ("Living cells reduce it to resorufin — pink and strongly fluorescent.", False),
    ("More metabolism = more pink = more signal.", True),
    ("The animal stays alive. Nothing is sacrificed.", False),
]
first = True
for text, bold in lines:
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    first = False
    p.text = text
    p.font.size = Pt(19); p.font.bold = bold
    p.font.color.rgb = PINK if bold else GRAY
    p.space_after = Pt(16)
notes(s, "Left plate: early in the incubation. Right plate: 4-6 hours later. You can literally see metabolism.")

# ---------------------------------------------------------------- 4 HOW IT RUNS
s = slide(); title(s, "How a run works", "Roughly half a day, bench-top, ~$1 of dye per plate")
bullets(s, [
    "Photograph animals for shell length (size normalization), then load one oyster per well.",
    "Fill wells with resazurin working solution; empty wells serve as blanks.",
    "Read fluorescence on a plate reader (ex 530 / em 590 nm) — baseline, then hourly for 4–6 h.",
    "Hold at whatever condition you want to probe: ambient, heat, low salinity, freshwater.",
    "Correct to blanks, normalize to shell length → a metabolism curve for every individual.",
], top=2.35, size=21, gap=16)
callout(s, "96 individuals per plate, non-lethal, no respirometry hardware required.")
notes(s, "Emphasize throughput: this is the reason it is useful operationally.")

# ---------------------------------------------------------------- 5 PUBLISHED
s = slide(); title(s, "This is published and validated",
                   "Huffmyer et al. 2026, PeerJ 14:e21542 — DOI 10.7717/peerj.21542")
bullets(s, [
    "Resazurin fluorescence correlates with oxygen consumption on the same individuals.",
    "Signal comes from live animals — live oysters far exceed empty shells and blanks (p < 0.001).",
    "Scales with body size, so we normalize to shell length (rho = 0.62, p < 0.001).",
    "Thermal performance curves behave as expected: metabolism peaks near 36 °C, then depresses.",
    "Works in both Pacific (C. gigas) and Eastern (C. virginica) oysters, seed through juveniles.",
], top=2.35, size=20, gap=14)
callout(s, "Peer-reviewed, open access, protocol included as an appendix — anyone can run it.", color=PINK)
notes(s, "Highlight that the method paper is out — this is not preliminary.")

# ---------------------------------------------------------------- 6 SECTION USE 1
section("Use #1 — Predicting phenotype",
        "Read metabolism today, forecast field performance later")

# ---------------------------------------------------------------- 7 PAPER RESULTS
s = slide(); title(s, "Metabolism already tracks outcomes",
                   "Key findings from the PeerJ study")
bullets(s, [
    "Under acute heat stress, oysters that survived had ~52% lower cumulative metabolism than those that died.",
    "High metabolic output under stress = higher risk of mortality (p < 0.001).",
    "Families differ significantly in metabolic response — the trait is genetically variable.",
    "In 50 selectively bred C. virginica families, metabolism correlated with predicted survival in the environment they were selected for (r = 0.56, p = 0.004).",
], top=2.35, size=21, gap=18)
callout(s, "Metabolism is not just a health readout — it carries information about future performance.")
notes(s, "Objectives 3, 4, 5 of the paper, compressed.")

# ---------------------------------------------------------------- 8 TRIALS
s = slide(); title(s, "Testing the idea head-on: survival trials",
                   "Nine USDA C. gigas families, run side by side")
rows = [
    ("Trial", "Stressor", "Oysters", "% died"),
    ("Apr 27", "33 °C", "128", "49%"),
    ("May 11", "33 °C", "84", "99%"),
    ("Jun 04", "36 °C", "39", "95%"),
    ("Jun 09", "36 °C", "77", "86%"),
    ("Jul 06", "35 °C", "378", "97%"),
    ("Jul 06", "35 °C + low salinity", "204", "84%"),
    ("Jul 07", "Freshwater → 36 °C", "185", "33%"),
]
gt = s.shapes.add_table(len(rows), 4, Inches(2.6), Inches(2.2), Inches(8.1), Inches(3.5)).table
gt.columns[0].width = Inches(1.6); gt.columns[1].width = Inches(3.3)
gt.columns[2].width = Inches(1.6); gt.columns[3].width = Inches(1.6)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = gt.cell(r, c)
        cell.text = val
        para = cell.text_frame.paragraphs[0]
        para.font.size = Pt(15)
        para.font.bold = (r == 0)
        para.alignment = PP_ALIGN.CENTER if c else PP_ALIGN.LEFT
callout(s, "Harshness varies enormously — raw survival can't be compared across trials.")
notes(s, "20 oysters per family per trial; time of death logged repeatedly.")

# ---------------------------------------------------------------- 9 SCORE
s = slide(); title(s, "A fair hardiness score", "0–100, comparable across trials")
bullets(s, [
    "Within each trial, rank families on (1) fraction surviving and (2) how long they lasted.",
    "Convert each family's standing to a percentile against the other families in that same trial — this cancels out trial harshness.",
    "Average those percentiles across all seven trials.",
], top=2.4, size=22, gap=20)
callout(s, "Higher score = consistently tougher across many kinds of stress.")
notes(s, "Keep this brief — it's plumbing, not the point.")

# ---------------------------------------------------------------- 10 RANKING FIG
s = slide(); title(s, "Families really do differ")
picture(s, "family_ranking.png", 2.4, 1.75, 8.5, 4.3)
callout(s, "Family 5 is the toughest; family 7 the most fragile — and the spread repeats across trials.")
notes(s, "This is the truth we want to predict.")

# ---------------------------------------------------------------- 11 BEST PREDICTOR
s = slide(); title(s, "Tough families crash later", "Single best metabolic predictor")
place(s, "best_single_predictor_scatter.png", 1.63, 2.25, 4.30)
tf = textbox(s, 7.7, 2.0, 4.9, 4.4)
items = [
    "Each dot is a family.",
    "X: time to the steepest metabolic decline under freshwater + heat. Further right = crashes later.",
    "Y: survival score (higher = tougher).",
    "rho = 0.90; cross-validated rho = 0.81 when each family is predicted from the other eight.",
]
first = True
for it in items:
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    first = False
    p.text = "• " + it
    p.font.size = Pt(17); p.font.color.rgb = GRAY; p.space_after = Pt(14)
notes(s, "Leave-one-family-out CV is the honest number.")

# ---------------------------------------------------------------- 12 SURPRISE
s = slide(); title(s, "The counterintuitive part", "Hardy oysters stay calm")
bullets(s, [
    "Families whose metabolism spikes hardest when heated survive the least.",
    "Hardy families keep metabolism restrained — they don't overreact.",
    "They decline slowly and late rather than crashing early.",
    "The signal is strongest when the assay is run during stress, not at rest.",
], top=2.5, size=24, gap=24)
notes(s, "Metabolic depression as a resilience strategy — consistent with the PeerJ survival result.")

# ---------------------------------------------------------------- 13 INDEX
s = slide(); title(s, "The resazurin index", "Four curve features, blended into one score per family")
rows = [
    ("Feature", "Plain meaning", "Survivor-like value"),
    ("Time to steepest drop", "how long before metabolism crashes", "later"),
    ("Steadiness", "how much metabolism bounces around", "steadier"),
    ("Early rise", "how fast metabolism climbs at first", "faster"),
    ("Depth of collapse", "how far it falls from its peak", "smaller"),
]
gt = s.shapes.add_table(len(rows), 3, Inches(1.0), Inches(2.35), Inches(11.3), Inches(2.6)).table
gt.columns[0].width = Inches(3.2); gt.columns[1].width = Inches(5.3); gt.columns[2].width = Inches(2.8)
for r, row in enumerate(rows):
    for c, val in enumerate(row):
        cell = gt.cell(r, c)
        cell.text = val
        para = cell.text_frame.paragraphs[0]
        para.font.size = Pt(16)
        para.font.bold = (r == 0) or (c == 2 and r > 0)
tf = textbox(s, 1.0, 5.15, 11.3, 0.9)
p = tf.paragraphs[0]
p.text = "Each feature is put on a common scale, flipped so higher = tougher, and the four are averaged."
p.font.size = Pt(18); p.font.color.rgb = GRAY
notes(s, "Measured under freshwater + heat. Blending averages out individual noise.")

# ---------------------------------------------------------------- 14 INDEX FIG
s = slide(); title(s, "The index tracks real survival")
picture(s, "resazurin_index_scatter.png", 0.8, 1.9, 6.6, 4.3)
tf = textbox(s, 7.7, 2.2, 4.9, 4.0)
items = [
    "X: combined resazurin index (further right = more survivor-like metabolism).",
    "Y: survival score from the lethal trials.",
    "Leave-one-family-out rho = 0.79 — strong prediction for families the model never saw.",
    "Toughest (5) and most fragile (7) land at opposite ends.",
]
first = True
for it in items:
    p = tf.paragraphs[0] if first else tf.add_paragraph()
    first = False
    p.text = "• " + it
    p.font.size = Pt(17); p.font.color.rgb = GRAY; p.space_after = Pt(16)
notes(s, "Half a day of dye work stands in for a season of grow-out.")

# ---------------------------------------------------------------- 15 SECTION USE 2
section("Use #2 — Real-time stress on the farm",
        "A thermometer for physiological condition")

# ---------------------------------------------------------------- 16 FARM USE
s = slide(); title(s, "Reading stress while it's happening",
                   "Same assay, different question")
bullets(s, [
    "Pull a small sample of animals, run a plate, get an answer the same day.",
    "Elevated metabolism under a heat or salinity event flags a crop under load — before mortality shows up.",
    "Compare bags, tidal heights, gear types, or sites on the same afternoon.",
    "Track recovery after an event: does metabolism return to baseline, and how fast?",
    "Non-lethal and cheap enough to repeat weekly through a risky season.",
], top=2.35, size=21, gap=16)
callout(s, "Turns \"the water got hot last week\" into a number you can act on.", color=PINK)
notes(s, "This is the operational pitch: early warning + management comparison.")

# ---------------------------------------------------------------- 17 NEXT
s = slide(); title(s, "What's next")
bullets(s, [
    "Broaden the index with additional assays:",
    ("Glycogen (heat vs. control, 3 h)", 1),
    ("Na+/K+ ATPase", 1),
    ("Citrate synthase", 1),
    "Validate predictions against field outplants across PNW sites.",
    "Package a grower-facing protocol and a simple scoring sheet.",
], top=2.3, size=21, gap=12)
notes(s, "Short — leave time for questions.")

# ---------------------------------------------------------------- 18 TAKEAWAYS
s = slide(); title(s, "Takeaways")
bullets(s, [
    "Resazurin turns metabolism into a color change — cheap, non-lethal, 96 animals at a time.",
    "Published and validated against oxygen consumption (PeerJ 2026).",
    "Predicts phenotype: metabolic curve shape forecasts family hardiness (cross-validated rho ≈ 0.8).",
    "Measures stress in real time: a same-day readout of how hard your crop is working.",
], top=2.4, size=22, gap=22)
callout(s, "FUNDING  - USDA-WRAC - sormi.science")
notes(s, "Land on the two grower uses.")

if os.path.exists(OUT):
    import shutil, datetime
    bak = OUT.replace(".pptx", "-backup-%s.pptx" % datetime.datetime.now().strftime("%Y%m%d-%H%M%S"))
    shutil.copy2(OUT, bak)
    print("backed up existing deck ->", bak)

prs.save(OUT)
print("wrote", OUT, "slides:", len(prs.slides.__iter__.__self__._sldIdLst))
