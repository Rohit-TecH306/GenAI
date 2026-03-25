from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE


prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Theme colors
BG_LIGHT = RGBColor(246, 252, 248)
PRIMARY = RGBColor(22, 101, 52)
PRIMARY_DARK = RGBColor(14, 78, 40)
ACCENT = RGBColor(74, 222, 128)
TEXT_DARK = RGBColor(17, 24, 39)
TEXT_MUTED = RGBColor(75, 85, 99)
WARN = RGBColor(245, 158, 11)
WHITE = RGBColor(255, 255, 255)


def add_bg(slide, color=BG_LIGHT):
    rect = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = color
    rect.line.fill.background()


def add_top_bar(slide, title):
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, Inches(0.9)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = PRIMARY
    bar.line.fill.background()

    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.18), Inches(9.5), Inches(0.6))
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title
    p.font.name = "Calibri"
    p.font.size = Pt(30)
    p.font.bold = True
    p.font.color.rgb = WHITE


def add_subtitle(slide, text):
    box = slide.shapes.add_textbox(Inches(0.6), Inches(1.05), Inches(12.0), Inches(0.6))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Calibri"
    p.font.size = Pt(18)
    p.font.color.rgb = TEXT_MUTED


def add_bullets(slide, bullets, x=0.9, y=1.8, w=7.0, h=4.8, size=24):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.clear()
    for i, line in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.level = 0
        p.font.name = "Calibri"
        p.font.size = Pt(size)
        p.font.color.rgb = TEXT_DARK
        p.space_after = Pt(8)


def add_metric_card(slide, x, y, w, h, label, value, color=PRIMARY):
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    card.fill.solid()
    card.fill.fore_color.rgb = WHITE
    card.line.color.rgb = ACCENT
    card.line.width = Pt(1.5)

    val = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.15), Inches(w - 0.4), Inches(0.6))
    tfv = val.text_frame
    pv = tfv.paragraphs[0]
    pv.text = value
    pv.font.name = "Calibri"
    pv.font.bold = True
    pv.font.size = Pt(28)
    pv.font.color.rgb = color

    lbl = slide.shapes.add_textbox(Inches(x + 0.2), Inches(y + 0.8), Inches(w - 0.4), Inches(h - 0.9))
    tfl = lbl.text_frame
    pl = tfl.paragraphs[0]
    pl.text = label
    pl.font.name = "Calibri"
    pl.font.size = Pt(15)
    pl.font.color.rgb = TEXT_MUTED


def add_footer(slide, text="AI Sustainable Agri Platform | Hackathon 2026"):
    footer = slide.shapes.add_textbox(Inches(0.5), Inches(7.1), Inches(12.3), Inches(0.3))
    tf = footer.text_frame
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Calibri"
    p.font.size = Pt(11)
    p.font.color.rgb = TEXT_MUTED
    p.alignment = PP_ALIGN.RIGHT


# Slide 1: Title
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, PRIMARY_DARK)
hero = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.6), prs.slide_width, Inches(1.9))
hero.fill.solid()
hero.fill.fore_color.rgb = RGBColor(9, 59, 30)
hero.line.fill.background()

box = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(2.4))
tf = box.text_frame
p = tf.paragraphs[0]
p.text = "India's First AI-Powered\nSustainable Agriculture\nDecision Platform"
p.font.name = "Calibri"
p.font.size = Pt(52)
p.font.bold = True
p.font.color.rgb = WHITE

sub = slide.shapes.add_textbox(Inches(0.9), Inches(4.4), Inches(11.0), Inches(1.0))
tf2 = sub.text_frame
p2 = tf2.paragraphs[0]
p2.text = "Better yield today. Healthier soil tomorrow."
p2.font.name = "Calibri"
p2.font.size = Pt(24)
p2.font.color.rgb = ACCENT

add_footer(slide, "Team: <Your Team Name> | National Hackathon Pitch")

# Slide 2: Problem
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide)
add_top_bar(slide, "The National Problem We Are Solving")
add_subtitle(slide, "Three connected gaps hurting farmer income and long-term food security")

add_metric_card(slide, 0.8, 2.0, 3.9, 2.2, "Farmers miss schemes due to low discoverability", "Low Scheme Access")
add_metric_card(slide, 4.9, 2.0, 3.9, 2.2, "Vendors cannot reliably reach verified rural buyers", "Market Access Gap")
add_metric_card(slide, 9.0, 2.0, 3.5, 2.2, "Short-term yield choices degrade long-term soil health", "Soil Risk", WARN)

add_bullets(
    slide,
    [
        "Our biggest innovation focus: Soil health negligence at decision time.",
        "Current systems optimize transactions, not sustainability outcomes.",
    ],
    x=0.9,
    y=4.8,
    w=11.8,
    h=1.8,
    size=20,
)
add_footer(slide)

# Slide 3: Solution Architecture
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide)
add_top_bar(slide, "Platform Architecture")
add_subtitle(slide, "Two-sided platform powered by a sustainability-first AI engine")

# Farmers block
f = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), Inches(3.9), Inches(4.8))
f.fill.solid(); f.fill.fore_color.rgb = WHITE; f.line.color.rgb = ACCENT
ftext = slide.shapes.add_textbox(Inches(1.05), Inches(2.1), Inches(3.4), Inches(4.4)).text_frame
for i, t in enumerate([
    "Farmers (Free)",
    "• AI Scheme Recommendations",
    "• Product Recommendations",
    "• Soil Health Tracking",
    "• Profit & Risk Insights",
]):
    p = ftext.paragraphs[0] if i == 0 else ftext.add_paragraph(); p.text = t
    p.font.name = "Calibri"; p.font.size = Pt(18 if i == 0 else 15)
    p.font.bold = (i == 0); p.font.color.rgb = (PRIMARY if i == 0 else TEXT_DARK)

# AI core
c = slide.shapes.add_shape(MSO_SHAPE.HEXAGON, Inches(4.95), Inches(2.15), Inches(3.4), Inches(3.8))
c.fill.solid(); c.fill.fore_color.rgb = PRIMARY; c.line.fill.background()
ct = slide.shapes.add_textbox(Inches(5.25), Inches(2.65), Inches(2.9), Inches(2.8)).text_frame
for i, t in enumerate([
    "AI CORE",
    "Scheme Matching",
    "Product Reco",
    "Soil Impact Score",
    "Behavior Nudges",
]):
    p = ct.paragraphs[0] if i == 0 else ct.add_paragraph(); p.text = t
    p.font.name = "Calibri"; p.font.size = Pt(17 if i == 0 else 14)
    p.font.bold = (i == 0); p.alignment = PP_ALIGN.CENTER; p.font.color.rgb = WHITE

# Vendors block
v = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.7), Inches(1.9), Inches(3.9), Inches(4.8))
v.fill.solid(); v.fill.fore_color.rgb = WHITE; v.line.color.rgb = ACCENT
vtext = slide.shapes.add_textbox(Inches(8.95), Inches(2.1), Inches(3.4), Inches(4.4)).text_frame
for i, t in enumerate([
    "Vendors (Paid)",
    "• Product Listing",
    "• Reach Verified Farmers",
    "• Analytics Dashboard",
    "• Featured Sustainability Slots",
]):
    p = vtext.paragraphs[0] if i == 0 else vtext.add_paragraph(); p.text = t
    p.font.name = "Calibri"; p.font.size = Pt(18 if i == 0 else 15)
    p.font.bold = (i == 0); p.font.color.rgb = (PRIMARY if i == 0 else TEXT_DARK)

add_footer(slide)

# Slide 4: Soil Impact Score
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide)
add_top_bar(slide, "Key Innovation: Soil Health Impact Score")
add_subtitle(slide, "A nutritional-label style sustainability score for every agri input")

card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(1.9), Inches(7.2), Inches(4.9))
card.fill.solid(); card.fill.fore_color.rgb = WHITE; card.line.color.rgb = ACCENT

title = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(6.7), Inches(0.5)).text_frame
pt = title.paragraphs[0]
pt.text = "Product: XYZ Fast-Growth Fertilizer"
pt.font.name = "Calibri"; pt.font.size = Pt(20); pt.font.bold = True; pt.font.color.rgb = PRIMARY_DARK

items = [
    "Short-term yield boost: HIGH",
    "Long-term soil impact: LOW ⚠",
    "Soil pH effect: +0.8 (acidic)",
    "Recovery time needed: 6-8 months",
    "Recommended for: Sandy soil, Wheat",
    "Avoid if: Used 2+ consecutive seasons",
]
add_bullets(slide, items, x=1.2, y=2.9, w=6.5, h=3.5, size=16)

right = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.5), Inches(2.0), Inches(3.9), Inches(4.7))
right.fill.solid(); right.fill.fore_color.rgb = RGBColor(236, 253, 245); right.line.color.rgb = ACCENT
rt = slide.shapes.add_textbox(Inches(8.8), Inches(2.35), Inches(3.3), Inches(3.9)).text_frame
for i, t in enumerate([
    "Why Judges Care",
    "• First-mover transparency",
    "• Vendor accountability",
    "• Behavioral change for farmers",
    "• Strong sustainability story",
    "• Data moat over time",
]):
    p = rt.paragraphs[0] if i == 0 else rt.add_paragraph(); p.text = t
    p.font.name = "Calibri"; p.font.size = Pt(18 if i == 0 else 15)
    p.font.bold = (i == 0); p.font.color.rgb = (PRIMARY_DARK if i == 0 else TEXT_DARK)

add_footer(slide)

# Slide 5: Soil Passport
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide)
add_top_bar(slide, "Soil Health Passport (Per Farmer)")
add_subtitle(slide, "Transforms marketplace behavior into long-term advisory intelligence")

profile = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.9), Inches(4.1), Inches(4.9))
profile.fill.solid(); profile.fill.fore_color.rgb = WHITE; profile.line.color.rgb = ACCENT
add_bullets(
    slide,
    [
        "Farmer: Ramesh",
        "Land: 3 acres",
        "Location: Vidarbha",
        "Current Soil Score: 62/100",
        "Last Year: 71/100",
    ],
    x=1.1, y=2.2, w=3.4, h=3.4, size=17
)

trend = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(1.9), Inches(7.3), Inches(2.2))
trend.fill.solid(); trend.fill.fore_color.rgb = WHITE; trend.line.color.rgb = ACCENT
add_bullets(
    slide,
    [
        "Usage History: Jan 2025 Product A → Yield +30%, pH dropped",
        "Jun 2025 Product B → Yield stable, soil score improved",
    ],
    x=5.5, y=2.2, w=6.8, h=1.5, size=15
)

rec = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.2), Inches(4.3), Inches(7.3), Inches(2.5))
rec.fill.solid(); rec.fill.fore_color.rgb = RGBColor(236, 253, 245); rec.line.color.rgb = ACCENT
add_bullets(
    slide,
    [
        "AI Recommendation:",
        "• Avoid nitrogen-heavy fertilizers for next 2 crop cycles",
        "• Eligible scheme auto-link: Soil Health Card support",
    ],
    x=5.5, y=4.65, w=6.8, h=1.8, size=15
)
add_footer(slide)

# Slide 6: Business Model
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide)
add_top_bar(slide, "Business Model with Ethical Incentives")
add_subtitle(slide, "Revenue aligned with sustainability outcomes")

streams = [
    ("Vendor Listing Fee", "Recurring"),
    ("Featured Placement", "Soil-score gated"),
    ("Vendor Analytics", "Subscription"),
    ("Premium Farmer Insights", "Optional"),
    ("Government Partnerships", "Future scale"),
]
for i, (a, b) in enumerate(streams):
    y = 1.9 + i * 0.95
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.9), Inches(y), Inches(7.6), Inches(0.75))
    card.fill.solid(); card.fill.fore_color.rgb = WHITE; card.line.color.rgb = ACCENT
    t = slide.shapes.add_textbox(Inches(1.2), Inches(y + 0.18), Inches(6.9), Inches(0.4)).text_frame
    p = t.paragraphs[0]; p.text = f"{a}  |  {b}"
    p.font.name = "Calibri"; p.font.size = Pt(16); p.font.color.rgb = TEXT_DARK

eth = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.9), Inches(2.0), Inches(3.6), Inches(4.7))
eth.fill.solid(); eth.fill.fore_color.rgb = PRIMARY; eth.line.fill.background()
eth_t = slide.shapes.add_textbox(Inches(9.2), Inches(2.4), Inches(3.0), Inches(3.8)).text_frame
for i, t in enumerate([
    "Ethical Lock-In",
    "Products with poor",
    "soil health ratings",
    "get lower organic",
    "visibility.",
    "\nVendors are nudged",
    "to improve quality.",
]):
    p = eth_t.paragraphs[0] if i == 0 else eth_t.add_paragraph(); p.text = t
    p.font.name = "Calibri"; p.font.size = Pt(19 if i == 0 else 15)
    p.font.bold = (i == 0); p.font.color.rgb = WHITE

add_footer(slide)

# Slide 7: GTM + Impact
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide)
add_top_bar(slide, "Go-To-Market and National Impact")
add_subtitle(slide, "Pilot fast, prove outcomes, then scale by district")

add_metric_card(slide, 0.9, 1.9, 3.7, 1.6, "Farming households in India", "140M+")
add_metric_card(slide, 4.9, 1.9, 3.7, 1.6, "Pilot districts (phase-1)", "1-2")
add_metric_card(slide, 8.9, 1.9, 3.7, 1.6, "Initial verified vendors", "5-10")

add_bullets(
    slide,
    [
        "Phase 1: District pilot with FPO/cooperative partnerships",
        "Phase 2: Vernacular + voice + WhatsApp + SMS fallback",
        "Phase 3: Hyperlocal soil clustering and seasonal intelligence",
        "Policy alignment: soil restoration, climate resilience, Digital India goals",
    ],
    x=0.9, y=3.9, w=11.8, h=2.6, size=18
)
add_footer(slide)

# Slide 8: Moat + Closing Ask
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(PRIMARY_DARK if False else slide)  # keeps same helper usage pattern
# Manually set dark background for finale
bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
bg.fill.solid(); bg.fill.fore_color.rgb = PRIMARY_DARK; bg.line.fill.background()

title = slide.shapes.add_textbox(Inches(0.7), Inches(0.6), Inches(12.0), Inches(0.8)).text_frame
p = title.paragraphs[0]; p.text = "Why We Win and What We Need"
p.font.name = "Calibri"; p.font.size = Pt(40); p.font.bold = True; p.font.color.rgb = WHITE

moat = slide.shapes.add_textbox(Inches(0.9), Inches(1.8), Inches(6.2), Inches(4.2)).text_frame
for i, t in enumerate([
    "Competitive Moat",
    "• AI scheme matching engine (already built)",
    "• Compounding Soil Health Passport data",
    "• Vendor sustainability ratings",
    "• Behavioral nudging intelligence",
    "• Policy and public-system alignment",
]):
    q = moat.paragraphs[0] if i == 0 else moat.add_paragraph(); q.text = t
    q.font.name = "Calibri"; q.font.size = Pt(22 if i == 0 else 16)
    q.font.bold = (i == 0); q.font.color.rgb = (ACCENT if i == 0 else WHITE)

ask = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.4), Inches(2.0), Inches(5.0), Inches(3.4))
ask.fill.solid(); ask.fill.fore_color.rgb = RGBColor(236, 253, 245); ask.line.fill.background()
ask_t = slide.shapes.add_textbox(Inches(7.7), Inches(2.35), Inches(4.4), Inches(2.8)).text_frame
for i, t in enumerate([
    "Hackathon Ask",
    "• Pilot district introductions",
    "• Data/API partnership mentorship",
    "• Govt and FPO connect support",
]):
    q = ask_t.paragraphs[0] if i == 0 else ask_t.add_paragraph(); q.text = t
    q.font.name = "Calibri"; q.font.size = Pt(20 if i == 0 else 15)
    q.font.bold = (i == 0); q.font.color.rgb = (PRIMARY_DARK if i == 0 else TEXT_DARK)

close = slide.shapes.add_textbox(Inches(0.8), Inches(6.3), Inches(11.8), Inches(0.8)).text_frame
cp = close.paragraphs[0]
cp.text = "We are not building a marketplace. We are building India's sustainable agriculture decision infrastructure."
cp.font.name = "Calibri"; cp.font.size = Pt(22); cp.font.bold = True; cp.font.color.rgb = ACCENT
cp.alignment = PP_ALIGN.CENTER

out_path = r"d:\VIIT\Github\GenAI\AgriTech_National_Hackathon_Pitch_Designed.pptx"
prs.save(out_path)
print(out_path)
