#!/usr/bin/env python3
"""Build the ISO/TC 154 DIS proposal presentation for lightweight-doc."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pathlib import Path

INK = RGBColor(0x0A, 0x16, 0x28)
PARCHMENT = RGBColor(0xF3, 0xEB, 0xE0)
PAPER = RGBColor(0xFF, 0xFA, 0xF2)
COPPER = RGBColor(0xB8, 0x5A, 0x2A)
CYAN = RGBColor(0x2F, 0x6F, 0x82)
MIST = RGBColor(0x8A, 0x9A, 0xAB)
SLATE = RGBColor(0x24, 0x36, 0x47)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]


def slide(bg=PAPER):
    s = prs.slides.add_slide(BLANK)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = bg
    return s


def box(s, x, y, w, h, text, size=18, color=INK, bold=False, mono=False, align=None):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    first = True
    for line in text.split("\n"):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.text = line
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bold
        if mono:
            p.font.name = "Courier New"
        if align:
            p.alignment = align
    return tb


def band(s=None, color=COPPER, h=0.14):
    from pptx.enum.shapes import MSO_SHAPE
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = color; r.line.fill.background()


def eyebrow(s, text, y=0.55, color=CYAN):
    box(s, 0.9, y, 11.5, 0.4, text, size=13, color=color, mono=True)


def title(s, text, y=0.95, size=40, color=INK):
    box(s, 0.9, y, 11.5, 1.1, text, size=size, color=color, bold=True)


def footer(s, n):
    box(s, 0.9, 7.0, 11.5, 0.4,
        f"CalConnect proposal to ISO/TC 154  ·  CC/ISO 36010  ·  {n}",
        size=10, color=MIST)


# ---------- 1 · Title ----------
s = slide(INK)
band(s, COPPER, 0.18)
box(s, 0.9, 1.5, 11.5, 0.5, "PROPOSAL FOR A DRAFT INTERNATIONAL STANDARD", size=15, color=CYAN, mono=True)
box(s, 0.9, 2.1, 11.5, 2.0,
    "Lightweight document —\nDocument metamodel", size=48, color=PAPER, bold=True)
box(s, 0.9, 4.3, 11.5, 0.6, "One model for every markup: harmonized, specializable, extensible",
    size=20, color=RGBColor(0xC5, 0xDD, 0xE4))
box(s, 0.9, 5.6, 11.5, 1.0,
    "Proposed by CalConnect (TC VCARD) to ISO/TC 154\nBase text: CC/ISO 36010 · Model source: basicdoc-models · Atlas: metanorma.github.io/basicdoc-models",
    size=14, color=MIST)

# ---------- 2 · The problem ----------
s = slide(); band(s); eyebrow(s, "THE PROBLEM")
title(s, "Structured text has no interoperability layer")
items = [
    ("Fragmented markups", "Markdown, AsciiDoc, reStructuredText — same ideas, incompatible syntaxes and semantics"),
    ("Lossy conversion", "Format-to-format converters drop semantics; exchange degrades to plain text, the lowest common denominator"),
    ("No shared reference model", "HTML/DocBook/TEI describe rendering or publishing, not a lightweight interchange structure"),
    ("Collaboration gap", "Documents edited by many parties need defined, incremental change — no common model to patch"),
]
y = 2.0
for head, body in items:
    box(s, 0.9, y, 3.4, 0.5, head, size=17, bold=True, color=COPPER)
    box(s, 4.5, y, 8.0, 0.9, body, size=15, color=SLATE)
    y += 1.15
footer(s, 2)

# ---------- 3 · The answer ----------
s = slide(); band(s); eyebrow(s, "THE ANSWER")
title(s, "BasicDocument: a lightweight document metamodel")
box(s, 0.9, 1.9, 11.5, 0.8,
    "A minimal, complete model of generic documents — documents are sections,\nsections are blocks, blocks are inline elements. The tiers never mix.",
    size=17, color=SLATE)
rows = [
    ("Harmonized", "Every markup construct maps onto one shared model (Annex A: AsciiDoc, Markdown/GFM/Pandoc, RST)"),
    ("Lightweight", "Deliberately not DocBook/TEI: a base others map to, specialize upon, or use for interchange"),
    ("Specializable", "Markup-specific kinds arrive as type values and attribute overrides — never parallel classes"),
    ("Extensible", "An open attribute register, opaque raw content, and variable references — no model change needed"),
    ("Collaborative", "Change models apply patches incrementally: insert, delete, move, modify"),
]
y = 3.1
for k, v in rows:
    box(s, 0.9, y, 2.6, 0.5, k, size=16, bold=True, color=CYAN)
    box(s, 3.7, y, 8.8, 0.7, v, size=14, color=SLATE)
    y += 0.72
footer(s, 3)

# ---------- 4 · Proven in production ----------
s = slide(); band(s); eyebrow(s, "EVIDENCE")
title(s, "Proven in production, not a paper model")
ev = [
    ("Metanorma", "The model underlies standards authoring and publishing across 30+ document flavours (relaton-models, standoc-models)"),
    ("Executable conformance", "Lint, parity, and twin XML/YAML instance gates run in CI — a construct without a serializable instance is unfinished (Annex D)"),
    ("Grammar + model in lockstep", "LML definition modules are the definitive expression; RelaxNG grammars compile and pass regeneration-parity in CI"),
    ("Live model atlas", "metanorma.github.io/basicdoc-models — every diagram plate with full model definitions, hyperlinked types"),
    ("Registry integration", "Localized strings identify spelling systems per ISO 24229; romanization schemes by system code"),
]
y = 2.0
for k, v in ev:
    box(s, 0.9, y, 3.6, 0.5, k, size=16, bold=True, color=COPPER)
    box(s, 4.7, y, 7.8, 0.9, v, size=14, color=SLATE)
    y += 0.95
footer(s, 4)

# ---------- 5 · One model, many markups ----------
s = slide(); band(s); eyebrow(s, "HARMONIZED MODELS")
title(s, "One model, many markups (Annex A)")
box(s, 0.9, 1.85, 11.5, 0.5, "Construct-by-construct mapping tables, normative:", size=15, color=SLATE)
cols = [
    ("AsciiDoc", "header → bibdata + register\nadmonitions → AdmonitionBlock + type\nlists (continuations) → relaxed ListItem\npassthroughs → FormattedString\n{attr} → ReferenceToVariable"),
    ("Markdown / GFM / Pandoc", "GFM tables → TableBlock\n{#id .class} → attribute register\n::: divs → register + any container\n$math$ → StemElement (LaTeX)\nraw HTML → FormattedString + format"),
    ("reStructuredText", "directives → register + relaxed content\n|substitution| → ReferenceToVariable\nroles → text-element specializations\nfield lists → document register\nadmonitions → type specializations"),
]
x = 0.9
for head, body in cols:
    box(s, x, 2.45, 3.7, 0.5, head, size=16, bold=True, color=CYAN)
    box(s, x, 3.0, 3.7, 2.6, body, size=12.5, color=SLATE, mono=True)
    x += 3.95
box(s, 0.9, 6.15, 11.5, 0.6,
    "Unlisted constructs are covered by the conformance formula:\nconstruct · type-specialization · register + relaxed or opaque content",
    size=14, color=COPPER, bold=True)
footer(s, 5)

# ---------- 6 · Specialization & extension ----------
s = slide(); band(s); eyebrow(s, "SPECIALIZATION & EXTENSION")
title(s, "Accommodate dialects without forking the model")
box(s, 0.9, 1.9, 5.6, 0.4, "SPECIALIZE (Annex B)", size=15, bold=True, color=CYAN, mono=True)
box(s, 0.9, 2.4, 5.6, 3.4,
    "· Type values: admonition kinds, list numeration,\n  stem languages, document classes\n· Attribute overrides: add, remove, retighten\n  (e.g. proscribing hanging paragraphs)\n· Class subclassing under the construct roots\n· Constraint specialization: PureTextElement\n  recursion guarantees",
    size=14, color=SLATE)
box(s, 7.0, 1.9, 5.4, 0.4, "EXTEND (Annex C)", size=15, bold=True, color=COPPER, mono=True)
box(s, 7.0, 2.4, 5.4, 3.4,
    "· Attribute register: open key-value entries on any\n  construct; well-known keys (id, class, lang,\n  unnumbered, format…), dialect-bound keys\n· Raw content: FormattedString / format-qualified —\n  the unknown stays lossless\n· Unknown directives decompose into register +\n  relaxed content (the entire Sphinx ecosystem)\n· A document is a block: whole documents compose",
    size=14, color=SLATE)
box(s, 0.9, 6.2, 11.5, 0.5, "OCP by construction: the basis is defined once; specialization and extension never modify it.",
    size=14, color=COPPER, bold=True)
footer(s, 6)

# ---------- 7 · Standardization readiness ----------
s = slide(); band(s); eyebrow(s, "READINESS")
title(s, "Built to be a standard")
rd = [
    ("Complete base text", "CC/ISO 36010: model clauses, conformance clause with a testable formula, normative mapping annex, specialization and extension annexes, abstract test suite (Annex D)"),
    ("Machine-checked", "Every claim that can be executed is executed in CI: model lint, diagram parity, grammar compilation and regeneration parity, twin-instance validation"),
    ("Aligned ecosystem", "ISO 639 / 15924 / 3166 / 8601 code sets; ISO 24229 spelling systems and system codes; Relaton (ISO 690 lineage) for bibliography"),
    ("Open governance artifacts", "Model sources, grammars, fixtures, atlas site, and CI are public: metanorma/basicdoc-models"),
    ("Existing consumers", "relaton-models, standoc-models, Metanorma flavours — the model already moves real standards today"),
]
y = 1.95
for k, v in rd:
    box(s, 0.9, y, 3.3, 0.5, k, size=15, bold=True, color=CYAN)
    box(s, 4.4, y, 8.1, 1.0, v, size=13, color=SLATE)
    y += 0.98
footer(s, 7)

# ---------- 8 · The ask ----------
s = slide(INK); band(s, COPPER, 0.18)
box(s, 0.9, 1.2, 11.5, 0.5, "THE ASK", size=15, color=CYAN, mono=True)
box(s, 0.9, 1.75, 11.5, 1.6,
    "Adopt CC/ISO 36010 as the base text for a\nDraft International Standard",
    size=36, color=PAPER, bold=True)
asks = [
    "Recognize the lightweight document metamodel as the interoperability layer for structured text",
    "Progress the CalConnect base text through NP → WD → CD → DIS at ISO/TC 154",
    "Designate the model repository and its executable conformance suite as the maintenance vehicle",
    "CalConnect (TC VCARD) continues as proposer and editor in liaison with TC 154",
]
y = 3.7
for a in asks:
    box(s, 1.1, y, 11.0, 0.6, "→  " + a, size=16, color=RGBColor(0xC5, 0xDD, 0xE4))
    y += 0.72
box(s, 0.9, 6.7, 11.5, 0.4,
    "metanorma.github.io/basicdoc-models  ·  github.com/metanorma/basicdoc-models  ·  CC/ISO 36010",
    size=12, color=MIST, mono=True)

out = Path(__file__).resolve().parent.parent / "presentations" / "lightweight-doc-tc154-dis-proposal.pptx"
out.parent.mkdir(exist_ok=True)
prs.save(str(out))
print(f"saved: {out}")
