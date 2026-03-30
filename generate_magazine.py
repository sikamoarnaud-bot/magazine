"""
generate_magazine.py
Generates magazine.docx — a professional French-language magazine in MS Word format.
Uses python-docx (pip install python-docx).
"""

import os
import sys

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY       = RGBColor(0x1A, 0x1A, 0x2E)   # dark navy (headings, borders)
GOLD       = RGBColor(0xE2, 0xB9, 0x6F)   # gold accent
DARK_GREY  = RGBColor(0x33, 0x33, 0x33)   # body text
MID_GREY   = RGBColor(0x77, 0x77, 0x77)   # captions / secondary
LIGHT_BG   = RGBColor(0xF4, 0xF2, 0xEE)   # card / tip background (unused in docx)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)

FONT_SERIF = "Georgia"
FONT_SANS  = "Calibri"

OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "magazine.docx")

# ── Helpers ───────────────────────────────────────────────────────────────────

def set_paragraph_spacing(para, before=0, after=6, line_rule=None, line_val=None):
    pf = para.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after  = Pt(after)
    if line_rule and line_val:
        pf.line_spacing_rule = line_rule
        pf.line_spacing = line_val


def add_run(para, text, bold=False, italic=False, size=None, color=None,
            font=FONT_SERIF, underline=False):
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.name = font
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return run


def set_cell_background(cell, hex_color):
    """Fill a table cell with a solid background colour."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def add_border_bottom(para, color="E2B96F", size=12):
    """Add a bottom border to a paragraph (used for section rules)."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size))
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_border_left(para, color="E2B96F", size=18):
    """Add a left border to a paragraph (used for pull quotes / standfirsts)."""
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), color)
    pBdr.append(left)
    pPr.append(pBdr)


def shade_paragraph(para, hex_color="F4F2EE"):
    """Shade a paragraph's background (for tips / cards)."""
    pPr = para._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    pPr.append(shd)


def page_break(doc):
    para = doc.add_paragraph()
    run = para.add_run()
    run.add_break(WD_BREAK.PAGE)
    set_paragraph_spacing(para, 0, 0)


# ── Section helpers ───────────────────────────────────────────────────────────

def cover_page(doc):
    """Dark-themed cover page using a 1-column table for visual effect."""
    # We simulate the dark cover with a 1×1 table spanning the page
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "1A1A2E")

    # badge
    p = cell.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    add_run(p, "N° 1  •  Printemps 2026", size=9, color=WHITE, font=FONT_SANS)
    set_paragraph_spacing(p, 0, 6)

    # spacer
    sp = cell.add_paragraph()
    set_paragraph_spacing(sp, 0, 30)

    # category label
    cat = cell.add_paragraph()
    cat.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(cat, "VOTRE MAGAZINE EN LIGNE", size=9, color=GOLD, font=FONT_SANS)
    set_paragraph_spacing(cat, 0, 4)

    # Title
    title = cell.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(title, "MAGAZINE", bold=True, size=54, color=WHITE, font=FONT_SERIF)
    set_paragraph_spacing(title, 0, 4)

    # Tagline
    tag = cell.add_paragraph()
    tag.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(tag, "Culture  ·  Société  ·  Innovation  ·  Lifestyle",
            size=11, color=RGBColor(0xCC, 0xCC, 0xCC), font=FONT_SANS)
    set_paragraph_spacing(tag, 0, 24)

    # Teasers
    teasers = [
        ("🌍 ", "Comprendre le monde", " — Les grands enjeux de demain"),
        ("💡 ", "Innovations", " — Quand la technologie change nos vies"),
        ("🎨 ", "Culture & Art", " — Les nouvelles créations qui font le buzz"),
        ("🌿 ", "Bien-être", " — Conseils pour une vie équilibrée"),
    ]
    for icon, bold_txt, rest in teasers:
        tp = cell.add_paragraph()
        tp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_run(tp, icon + bold_txt, bold=True, size=11, color=WHITE, font=FONT_SANS)
        add_run(tp, rest, size=11, color=RGBColor(0xBB, 0xBB, 0xBB), font=FONT_SANS)
        set_paragraph_spacing(tp, 0, 4)

    # spacer
    sp2 = cell.add_paragraph()
    set_paragraph_spacing(sp2, 0, 24)

    # footer
    foot = cell.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(foot, "www.magazine.io", size=9,
            color=RGBColor(0x88, 0x88, 0x88), font=FONT_SANS)
    set_paragraph_spacing(foot, 0, 0)

    # Remove default first empty paragraph that Word adds
    # (already handled by the table being first element)
    doc.add_paragraph()  # small gap after table


def toc_section(doc):
    page_break(doc)

    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(h, "SOMMAIRE", size=9, color=MID_GREY, font=FONT_SANS)
    add_border_bottom(h, "D5CFB7", 6)
    set_paragraph_spacing(h, 0, 18)

    entries = [
        ("03", "Éditorial — Bienvenue dans ce premier numéro"),
        ("06", "Dossier — Comprendre le monde de demain"),
        ("12", "Innovations — La technologie au service de l'humain"),
        ("18", "Culture & Art — Les créateurs qui inspirent"),
        ("24", "Société — Nouvelles façons de vivre ensemble"),
        ("30", "Bien-être — Prendre soin de soi au quotidien"),
    ]
    for num, label in entries:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_run(p, num + "  ", bold=True, size=11, color=GOLD, font=FONT_SANS)
        add_run(p, label, size=11, color=DARK_GREY, font=FONT_SERIF)
        add_border_bottom(p, "D5CFB7", 4)
        set_paragraph_spacing(p, 0, 6)


def article_header(doc, rubric, title, standfirst=None):
    """Print the coloured rubric bar, then the article title and optional standfirst."""
    # rubric bar
    bar = doc.add_paragraph()
    bar.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(bar, rubric.upper(), bold=True, size=8, color=NAVY, font=FONT_SANS)
    add_border_bottom(bar, "1A1A2E", 18)
    set_paragraph_spacing(bar, 0, 8)

    # title
    t = doc.add_paragraph()
    t.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(t, title, bold=True, size=26, color=NAVY, font=FONT_SERIF)
    set_paragraph_spacing(t, 0, 10)

    # standfirst / lead
    if standfirst:
        sf = doc.add_paragraph()
        sf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        add_run(sf, standfirst, italic=True, size=12, color=DARK_GREY, font=FONT_SERIF)
        add_border_left(sf, "E2B96F", 18)
        sf.paragraph_format.left_indent = Cm(0.6)
        set_paragraph_spacing(sf, 4, 12)


def body_paragraph(doc, text, size=11):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(p, text, size=size, color=DARK_GREY, font=FONT_SERIF)
    set_paragraph_spacing(p, 0, 8)
    return p


def pull_quote(doc, text):
    pq = doc.add_paragraph()
    pq.alignment = WD_ALIGN_PARAGRAPH.LEFT
    add_run(pq, text, italic=True, size=13, color=NAVY, font=FONT_SERIF)
    add_border_left(pq, "E2B96F", 24)
    shade_paragraph(pq, "F4F2EE")
    pq.paragraph_format.left_indent = Cm(0.6)
    pq.paragraph_format.right_indent = Cm(0.6)
    set_paragraph_spacing(pq, 10, 10)


def card(doc, icon, heading, text):
    """Simple card block using a shaded table cell."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "F4F2EE")
    # top gold border via paragraph
    h = cell.add_paragraph()
    add_border_bottom(h, "0F3460", 18)
    set_paragraph_spacing(h, 0, 0)

    hp = cell.add_paragraph()
    add_run(hp, icon + "  " + heading, bold=True, size=12, color=NAVY, font=FONT_SANS)
    set_paragraph_spacing(hp, 4, 4)

    tp = cell.add_paragraph()
    add_run(tp, text, size=10, color=DARK_GREY, font=FONT_SERIF)
    set_paragraph_spacing(tp, 0, 6)

    doc.add_paragraph()  # gap after card


def tip_block(doc, number, heading, text):
    """Numbered tip row with a gold left accent."""
    tbl = doc.add_table(rows=1, cols=2)
    tbl.style = "Table Grid"
    # number cell
    nc = tbl.cell(0, 0)
    set_cell_background(nc, "FAF9F7")
    nc.width = Cm(1.6)
    np = nc.add_paragraph()
    np.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(np, number, bold=True, size=20, color=GOLD, font=FONT_SANS)
    set_paragraph_spacing(np, 8, 8)

    # content cell
    cc = tbl.cell(0, 1)
    set_cell_background(cc, "FAF9F7")
    h = cc.add_paragraph()
    add_run(h, heading, bold=True, size=11, color=NAVY, font=FONT_SANS)
    set_paragraph_spacing(h, 6, 2)

    bp = cc.add_paragraph()
    add_run(bp, text, size=10, color=DARK_GREY, font=FONT_SERIF)
    set_paragraph_spacing(bp, 0, 6)

    doc.add_paragraph()  # gap


def back_cover(doc):
    page_break(doc)
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = "Table Grid"
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "0F3460")

    sp = cell.add_paragraph()
    set_paragraph_spacing(sp, 0, 40)

    logo = cell.add_paragraph()
    logo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(logo, "MAGAZINE", bold=True, size=42, color=WHITE, font=FONT_SERIF)
    set_paragraph_spacing(logo, 0, 6)

    tag = cell.add_paragraph()
    tag.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(tag, "Culture  ·  Société  ·  Innovation  ·  Lifestyle",
            size=10, color=RGBColor(0xAA, 0xAA, 0xAA), font=FONT_SANS)
    set_paragraph_spacing(tag, 0, 18)

    nxt = cell.add_paragraph()
    nxt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(nxt, "Prochain numéro — Été 2026", size=11, color=GOLD, font=FONT_SANS)
    set_paragraph_spacing(nxt, 0, 40)

    sep = cell.add_paragraph()
    add_border_bottom(sep, "FFFFFF40", 4)
    set_paragraph_spacing(sep, 0, 6)

    foot = cell.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(foot, "www.magazine.io   |   contact@magazine.io   |   ISSN 0000-0000",
            size=8, color=RGBColor(0x77, 0x77, 0x77), font=FONT_SANS)
    set_paragraph_spacing(foot, 0, 0)


# ── Build the document ────────────────────────────────────────────────────────

def build():
    doc = Document()

    # ── Page setup: A4, narrow margins ──
    section = doc.sections[0]
    section.page_width  = Cm(21)
    section.page_height = Cm(29.7)
    section.left_margin   = Cm(2.0)
    section.right_margin  = Cm(2.0)
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)

    # ── 1. COVER ──────────────────────────────────────────────────────────────
    cover_page(doc)

    # ── 2. TABLE OF CONTENTS ──────────────────────────────────────────────────
    toc_section(doc)

    # ── 3. ÉDITORIAL ─────────────────────────────────────────────────────────
    page_break(doc)
    article_header(
        doc, "Éditorial", "Bienvenue dans ce premier numéro",
        standfirst="Cette publication est née d'une envie simple : partager des idées, "
                   "des histoires et des perspectives qui enrichissent notre quotidien."
    )
    body_paragraph(doc,
        "Dans un monde en perpétuel mouvement, il est parfois difficile de prendre le recul "
        "nécessaire pour analyser les grandes tendances qui façonnent nos sociétés. C'est "
        "précisément le rôle que nous souhaitons jouer : celui d'un compagnon de réflexion, "
        "d'exploration et d'inspiration.")
    body_paragraph(doc,
        "Chaque numéro abordera des sujets variés — de la culture à l'innovation, du bien-être "
        "à la société — avec le souci constant d'apporter de la profondeur et de la nuance à des "
        "questions qui nous touchent tous.")
    body_paragraph(doc,
        "Nous espérons que ces pages sauront éveiller votre curiosité, provoquer des débats "
        "constructifs et, pourquoi pas, vous donner envie d'agir à votre échelle.")
    body_paragraph(doc, "Bonne lecture !")
    sig = doc.add_paragraph()
    add_run(sig, "— La Rédaction", italic=True, size=10, color=MID_GREY, font=FONT_SANS)
    set_paragraph_spacing(sig, 4, 0)

    # ── 4. DOSSIER ────────────────────────────────────────────────────────────
    page_break(doc)
    article_header(
        doc, "Dossier", "Comprendre le monde de demain",
        standfirst="Crise climatique, géopolitique en recomposition, nouvelles puissances "
                   "émergentes… Quels sont les grands défis qui redessineront la carte du monde "
                   "dans les prochaines décennies ?"
    )
    body_paragraph(doc,
        "Le XXIe siècle s'annonce comme l'un des plus décisifs de l'histoire humaine. "
        "Les transitions — énergétique, démographique, numérique — se superposent et "
        "interagissent, créant des dynamiques à la fois inédites et imprévisibles.")
    body_paragraph(doc,
        "La question climatique occupe désormais le premier rang des préoccupations mondiales. "
        "Les événements météorologiques extrêmes se multiplient, forçant les gouvernements, "
        "les entreprises et les citoyens à reconsidérer leurs modes de vie et de production.")
    body_paragraph(doc,
        "Sur le plan géopolitique, l'ordre mondial multipolaire qui se dessine remet en cause "
        "les équilibres établis depuis la fin de la Guerre froide. De nouvelles alliances se "
        "nouent, de nouvelles tensions émergent, et les questions de souveraineté — numérique, "
        "alimentaire, sanitaire — sont au cœur des débats.")
    body_paragraph(doc,
        "Face à ces défis, des solutions émergent pourtant : technologies vertes, économie "
        "circulaire, coopération internationale renforcée. L'avenir n'est pas écrit ; il se "
        "construit dans les choix que nous faisons collectivement aujourd'hui.")
    pull_quote(doc,
        "« L'avenir appartient à ceux qui comprennent que les crises sont aussi des "
        "opportunités de transformation. »")

    # ── 5. INNOVATIONS ───────────────────────────────────────────────────────
    page_break(doc)
    article_header(
        doc, "Innovations", "La technologie au service de l'humain",
        standfirst="Intelligence artificielle, biotechnologies, énergies renouvelables : "
                   "comment les innovations récentes transforment-elles notre quotidien et "
                   "redéfinissent-elles les frontières du possible ?"
    )
    card(doc, "🤖", "Intelligence artificielle",
         "L'IA générative révolutionne la création de contenu, la médecine diagnostique et "
         "l'industrie manufacturière. Si les opportunités sont immenses, les enjeux éthiques "
         "restent au cœur des débats.")
    card(doc, "🧬", "Biotechnologies",
         "L'édition génomique (CRISPR), les thérapies cellulaires et les vaccins à ARNm "
         "ouvrent de nouvelles perspectives dans le traitement des maladies rares et des cancers.")
    card(doc, "☀️", "Énergies vertes",
         "Le coût des énergies renouvelables continue de chuter. Solaire, éolien, hydrogène "
         "vert : la transition énergétique s'accélère, portée par des investissements massifs.")

    # ── 6. CULTURE & ART ─────────────────────────────────────────────────────
    page_break(doc)
    article_header(
        doc, "Culture & Art", "Les créateurs qui inspirent",
        standfirst="De la peinture numérique aux installations immersives, en passant par la "
                   "littérature engagée, un panorama des artistes et des œuvres qui marquent "
                   "notre époque."
    )
    body_paragraph(doc,
        "L'art contemporain n'a jamais été aussi accessible. Les plateformes numériques ont "
        "démocratisé la création et la diffusion, permettant à des artistes du monde entier de "
        "toucher des audiences mondiales sans passer par les circuits traditionnels.")
    body_paragraph(doc,
        "Les NFT, bien que controversés, ont ouvert un débat profond sur la valeur, l'originalité "
        "et la propriété des œuvres numériques. Certains y voient une révolution ; d'autres, une "
        "bulle spéculative. Dans tous les cas, ils ont bousculé l'écosystème.")
    body_paragraph(doc,
        "La littérature francophone connaît elle aussi un renouveau remarquable. De nouveaux "
        "auteurs, portés par des maisons d'édition indépendantes et les réseaux sociaux, proposent "
        "des récits qui explorent des territoires jusque-là peu représentés.")
    body_paragraph(doc,
        "Le cinéma, quant à lui, se réinvente face à la concurrence des plateformes de streaming. "
        "Les formats hybrides — séries-films, documentaires-fictions — brouillent les frontières "
        "traditionnelles et enrichissent l'expérience du spectateur.")

    # ── 7. SOCIÉTÉ ───────────────────────────────────────────────────────────
    page_break(doc)
    article_header(
        doc, "Société", "Nouvelles façons de vivre ensemble",
        standfirst="Télétravail, habitat participatif, économie collaborative : comment nos "
                   "modes de vie et de travail se transforment-ils, et quelles nouvelles formes "
                   "de lien social inventons-nous ?"
    )
    body_paragraph(doc,
        "La pandémie de Covid-19 a agi comme un accélérateur de tendances déjà à l'œuvre. "
        "Le télétravail, marginal avant 2020, est désormais une réalité pour des millions de "
        "travailleurs. Cette transformation recompose les rapports entre vie professionnelle et "
        "vie personnelle, mais aussi les dynamiques urbaines.")
    body_paragraph(doc,
        "L'habitat participatif et le cohabitat (coliving) connaissent un essor inédit. "
        "Partager des espaces communs, mutualiser des ressources, construire une communauté de "
        "voisinage : ces pratiques répondent à la fois à des enjeux économiques et à un besoin "
        "profond de lien social.")
    body_paragraph(doc,
        "L'économie collaborative, portée par des plateformes numériques, a transformé nos "
        "manières de consommer, de se déplacer et d'apprendre. Si ses promesses initiales de "
        "partage et de solidarité n'ont pas toujours été tenues, elle a durablement modifié nos "
        "comportements.")
    body_paragraph(doc,
        "La question du vivre-ensemble dans des sociétés de plus en plus diverses reste centrale. "
        "Comment construire une cohésion sociale tout en célébrant les différences ? C'est l'un "
        "des grands défis politiques et culturels de notre temps.")

    # ── 8. BIEN-ÊTRE ─────────────────────────────────────────────────────────
    page_break(doc)
    article_header(
        doc, "Bien-être", "Prendre soin de soi au quotidien",
        standfirst="Dans un monde hyperconnecté et souvent stressant, comment trouver des "
                   "équilibres durables pour notre santé physique, mentale et émotionnelle ?"
    )
    tips = [
        ("01", "Pratiquer la déconnexion",
         "Réserver des plages horaires sans écrans favorise la concentration, améliore la "
         "qualité du sommeil et réduit le stress."),
        ("02", "Bouger chaque jour",
         "Trente minutes d'activité physique quotidienne suffisent pour réduire "
         "significativement les risques cardiovasculaires et améliorer l'humeur."),
        ("03", "Cultiver la pleine conscience",
         "La méditation de pleine conscience, même pratiquée cinq minutes par jour, aide à "
         "mieux gérer les émotions et à renforcer la résilience."),
        ("04", "Soigner son alimentation",
         "Un régime varié, riche en végétaux et pauvre en aliments ultra-transformés, est le "
         "socle d'une bonne santé à long terme."),
        ("05", "Entretenir ses liens sociaux",
         "Les relations sociales de qualité sont l'un des meilleurs prédicteurs de longévité "
         "et de bien-être mental selon de nombreuses études."),
    ]
    for num, h, t in tips:
        tip_block(doc, num, h, t)

    # ── 9. BACK COVER ────────────────────────────────────────────────────────
    back_cover(doc)

    # ── Save ─────────────────────────────────────────────────────────────────
    doc.save(OUTPUT_PATH)
    print(f"✅  Saved: {OUTPUT_PATH}")


if __name__ == "__main__":
    build()
