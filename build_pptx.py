"""
build_pptx.py
Generates fasano-mamucabo.pptx — a fully editable A4 portrait PPTX
replicating the Fasano Mamucabo HTML one-pager design.
"""

from pptx import Presentation
from pptx.util import Emu, Pt, Mm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import copy
import os
from lxml import etree

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
ASSETS = "/home/user/pptx/assets"

# Slide size: A4 portrait
SLIDE_W = 7560072   # 210 mm
SLIDE_H = 10692792  # 297 mm

# Colors
NAVY   = RGBColor(0x20, 0x1A, 0x47)
PURPLE = RGBColor(0x88, 0x7C, 0xD3)
GOLD   = RGBColor(0xC8, 0xA9, 0x6D)
GRAY   = RGBColor(0x7F, 0x80, 0x8F)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
DARK   = RGBColor(0x1A, 0x1A, 0x2E)

def mm(val):
    return Emu(int(val * 36000))  # 1 mm = 36000 EMU

def rgb_hex(r, g, b):
    return RGBColor(r, g, b)

# ---------------------------------------------------------------------------
# Helper: add solid rectangle
# ---------------------------------------------------------------------------
def add_rect(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        left, top, width, height
    )
    shape.line.fill.background()
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = fill_color
    return shape

# ---------------------------------------------------------------------------
# Helper: add textbox with full control
# ---------------------------------------------------------------------------
def add_textbox(slide, left, top, width, height, text,
                font_name="Calibri", font_size=8, bold=False, italic=False,
                color=WHITE, align=PP_ALIGN.LEFT, word_wrap=True,
                space_before=0, space_after=0, line_spacing=None):
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = word_wrap
    tf = txb.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.alignment = align
    if space_before:
        p.space_before = Pt(space_before)
    if space_after:
        p.space_after = Pt(space_after)
    if line_spacing:
        p.line_spacing = Pt(line_spacing)
    run = p.add_run()
    run.text = text
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return txb

# ---------------------------------------------------------------------------
# Helper: add textbox with multiple paragraphs
# ---------------------------------------------------------------------------
def add_textbox_paragraphs(slide, left, top, width, height, paragraphs,
                            word_wrap=True):
    """
    paragraphs: list of dicts with keys:
      text, font_size, bold, italic, color, align, space_before, space_after,
      line_spacing, bullet (bool)
    """
    txb = slide.shapes.add_textbox(left, top, width, height)
    txb.word_wrap = word_wrap
    tf = txb.text_frame
    tf.word_wrap = word_wrap

    first = True
    for pdata in paragraphs:
        if first:
            p = tf.paragraphs[0]
            first = False
        else:
            p = tf.add_paragraph()

        p.alignment = pdata.get("align", PP_ALIGN.LEFT)
        sb = pdata.get("space_before", 0)
        sa = pdata.get("space_after", 0)
        ls = pdata.get("line_spacing", None)
        if sb:
            p.space_before = Pt(sb)
        if sa:
            p.space_after = Pt(sa)
        if ls:
            p.line_spacing = Pt(ls)

        run = p.add_run()
        run.text = pdata.get("text", "")
        run.font.name = pdata.get("font_name", "Calibri")
        run.font.size = Pt(pdata.get("font_size", 8))
        run.font.bold = pdata.get("bold", False)
        run.font.italic = pdata.get("italic", False)
        run.font.color.rgb = pdata.get("color", WHITE)

    return txb

# ---------------------------------------------------------------------------
# Helper: add image fitted to box
# ---------------------------------------------------------------------------
def add_image(slide, path, left, top, width, height):
    pic = slide.shapes.add_picture(path, left, top, width, height)
    return pic

# ---------------------------------------------------------------------------
# Helper: add vertical line
# ---------------------------------------------------------------------------
def add_vline(slide, x, top, height, color=PURPLE, width_pt=0.5):
    from pptx.util import Pt as PtU
    line = slide.shapes.add_shape(1, x, top, Emu(int(0.5 * 12700)), height)
    line.fill.background()
    line.line.color.rgb = color
    line.line.width = PtU(width_pt)
    return line

# ---------------------------------------------------------------------------
# Section label helper
# ---------------------------------------------------------------------------
def section_label(slide, left, top, width, text):
    add_textbox(slide, left, top, width, mm(5),
                text, font_size=6, bold=True, color=GOLD,
                align=PP_ALIGN.LEFT)

# ---------------------------------------------------------------------------
# Bullet list helper
# ---------------------------------------------------------------------------
def bullet_list(slide, left, top, width, height, items,
                font_size=7, color=WHITE, bullet_color=PURPLE):
    paragraphs = []
    for i, item in enumerate(items):
        paragraphs.append({
            "text": f"•  {item}",
            "font_size": font_size,
            "bold": False,
            "color": color,
            "align": PP_ALIGN.LEFT,
            "space_before": 1 if i > 0 else 0,
            "space_after": 0,
            "line_spacing": 9,
        })
    add_textbox_paragraphs(slide, left, top, width, height, paragraphs)

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def build():
    prs = Presentation()
    prs.slide_width  = Emu(SLIDE_W)
    prs.slide_height = Emu(SLIDE_H)

    blank_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(blank_layout)

    # -----------------------------------------------------------------------
    # DIMENSIONS (all in mm from top)
    # -----------------------------------------------------------------------
    HEADER_H  = 22   # mm
    HERO_H    = 30   # mm
    INTRO_H   = 14   # mm
    BOTTOM_H  = 12   # mm
    FOOTER_H  = 18   # mm
    BODY_TOP  = HEADER_H + HERO_H + INTRO_H
    BODY_H    = 297 - BODY_TOP - BOTTOM_H - FOOTER_H  # ~201mm
    COL_W     = 210 / 3  # ~70mm each

    MARGIN    = 3    # mm inner margin
    IMG_H     = 22   # mm for paired images in columns

    # -----------------------------------------------------------------------
    # 1. HEADER
    # -----------------------------------------------------------------------
    add_rect(slide, mm(0), mm(0), mm(210), mm(HEADER_H), NAVY)

    # Logo
    logo_path = os.path.join(ASSETS, "logo_mb_ib_white.png")
    add_image(slide, logo_path, mm(3), mm(2), mm(25), mm(18))

    # Title
    add_textbox(slide, mm(32), mm(2), mm(120), mm(11),
                "FASANO MAMUCABO",
                font_size=18, bold=True, color=GOLD,
                align=PP_ALIGN.LEFT)

    # Subtitle
    add_textbox(slide, mm(32), mm(13), mm(130), mm(7),
                "Grupo Prima × Fasano | Oportunidade de Coinvestimento",
                font_size=8, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT)

    # Confidential label
    add_textbox(slide, mm(160), mm(7), mm(47), mm(8),
                "CONFIDENCIAL",
                font_size=7, bold=True, color=GOLD,
                align=PP_ALIGN.RIGHT)

    # -----------------------------------------------------------------------
    # 2. HERO IMAGE STRIP
    # -----------------------------------------------------------------------
    hero_path = os.path.join(ASSETS, "hero.jpg")
    add_image(slide, hero_path, mm(0), mm(HEADER_H), mm(210), mm(HERO_H))

    # -----------------------------------------------------------------------
    # 3. INTRO BAR
    # -----------------------------------------------------------------------
    intro_top = HEADER_H + HERO_H
    add_rect(slide, mm(0), mm(intro_top), mm(210), mm(INTRO_H), NAVY)

    intro_text = ("Projeto residencial de ultra-luxo desenvolvido pelo Grupo Prima em parceria "
                  "com a rede hoteleira Fasano, localizado em Mamucabo, litoral sul do Rio de Janeiro.")
    add_textbox(slide, mm(5), mm(intro_top + 2), mm(200), mm(INTRO_H - 2),
                intro_text,
                font_size=8, bold=False, color=WHITE,
                align=PP_ALIGN.CENTER)

    # -----------------------------------------------------------------------
    # 4. BODY BACKGROUND
    # -----------------------------------------------------------------------
    add_rect(slide, mm(0), mm(BODY_TOP), mm(210), mm(BODY_H), DARK)

    # Column separators
    add_vline(slide, mm(COL_W), mm(BODY_TOP + 2), mm(BODY_H - 4))
    add_vline(slide, mm(COL_W * 2), mm(BODY_TOP + 2), mm(BODY_H - 4))

    # -----------------------------------------------------------------------
    # COLUMN 1
    # -----------------------------------------------------------------------
    c1_left = mm(MARGIN)
    c1_w    = mm(COL_W - MARGIN * 1.5)
    cy      = BODY_TOP + 3   # current y in mm

    # --- Section A: O PROJETO ---
    section_label(slide, c1_left, mm(cy), c1_w, "O PROJETO")
    cy += 5

    add_textbox(slide, c1_left, mm(cy), c1_w, mm(7),
                "Fasano Mamucabo",
                font_size=9, bold=True, color=WHITE)
    cy += 7

    items_a = [
        "Empreendimento residencial de alto padrão",
        "43 unidades residenciais — studios a penthouses",
        "Área privativa: 45m² a 500m²",
        "Serviços hoteleiros Fasano integrados",
        "Localização: Mamucabo, RJ (entre Paraty e Angra)",
    ]
    bullet_list(slide, c1_left, mm(cy), c1_w, mm(30), items_a, font_size=7, color=WHITE)
    cy += 28

    # Two images side by side
    img_w = (COL_W - MARGIN * 1.5) / 2 - 1
    add_image(slide, os.path.join(ASSETS, "pool.jpg"),
              c1_left, mm(cy), mm(img_w), mm(IMG_H))
    add_image(slide, os.path.join(ASSETS, "entrance.jpg"),
              c1_left + mm(img_w + 1), mm(cy), mm(img_w), mm(IMG_H))
    cy += IMG_H + 3

    # --- Section B: GRUPO PRIMA ---
    section_label(slide, c1_left, mm(cy), c1_w, "GRUPO PRIMA")
    cy += 5

    items_b = [
        "Incorporadora com +20 anos de experiência",
        "Portfólio: +R$2Bi em VGV lançado",
        "51% do capital do projeto",
        "AF Terreno Residencial (840.834m²): Outorgado para 4k unidades",
        "AF Terreno Residencial (240.829m²): Outorgado para 4k unidades",
        "AF Terreno Hoteleiro (300.076m²)",
    ]
    bullet_list(slide, c1_left, mm(cy), c1_w, mm(36), items_b, font_size=7, color=WHITE)
    cy += 34

    # Two images side by side
    add_image(slide, os.path.join(ASSETS, "villa.jpg"),
              c1_left, mm(cy), mm(img_w), mm(IMG_H))
    add_image(slide, os.path.join(ASSETS, "lobby.jpg"),
              c1_left + mm(img_w + 1), mm(cy), mm(img_w), mm(IMG_H))

    # -----------------------------------------------------------------------
    # COLUMN 2
    # -----------------------------------------------------------------------
    c2_left = mm(COL_W + MARGIN * 0.5)
    c2_w    = mm(COL_W - MARGIN)
    cy2     = BODY_TOP + 3

    # --- Section C: FASANO ---
    section_label(slide, c2_left, mm(cy2), c2_w, "FASANO")
    cy2 += 5

    items_c = [
        "Rede hoteleira de ultra-luxo referência no Brasil",
        "Presente em São Paulo, Rio, Trancoso, Punta del Este, Lisboa",
        "Gestão hoteleira + brand licensing do projeto",
        "49% do capital do projeto",
    ]
    bullet_list(slide, c2_left, mm(cy2), c2_w, mm(25), items_c, font_size=7, color=WHITE)
    cy2 += 23

    # Two images side by side
    img_w2 = (COL_W - MARGIN) / 2 - 1
    add_image(slide, os.path.join(ASSETS, "room.jpg"),
              c2_left, mm(cy2), mm(img_w2), mm(IMG_H))
    add_image(slide, os.path.join(ASSETS, "aerial.jpg"),
              c2_left + mm(img_w2 + 1), mm(cy2), mm(img_w2), mm(IMG_H))
    cy2 += IMG_H + 3

    # --- Section D: ESTRUTURA CRI ---
    section_label(slide, c2_left, mm(cy2), c2_w, "ESTRUTURA CRI")
    cy2 += 5

    # Sub-label
    add_textbox(slide, c2_left, mm(cy2), c2_w, mm(5),
                "GARANTIAS CRI",
                font_size=6, bold=True, color=NAVY,
                align=PP_ALIGN.LEFT)
    cy2 += 5

    items_d = [
        "Cessão de 100% das cotas das Devedoras",
        "AF Terreno Residencial (840.834m²): Outorgado para 4k unidades",
        "AF Terreno Hoteleiro (300.076m²)",
        "Cessão fiduciária de recebíveis e das quotas de sociedades hoteleiras",
    ]
    bullet_list(slide, c2_left, mm(cy2), c2_w, mm(28), items_d, font_size=7, color=WHITE)

    # -----------------------------------------------------------------------
    # COLUMN 3
    # -----------------------------------------------------------------------
    c3_left = mm(COL_W * 2 + MARGIN * 0.5)
    c3_w    = mm(COL_W - MARGIN - 1)
    cy3     = BODY_TOP + 3

    # --- Section E: INDICADORES ---
    section_label(slide, c3_left, mm(cy3), c3_w, "INDICADORES")
    cy3 += 5

    metrics = [
        ("R$ 330Mi", "Captação Total CRI"),
        ("R$ 1,25Bi", "VGV Total do Projeto"),
        ("13,5% a.a.", "Taxa CRI (IPCA+)"),
        ("48 meses", "Prazo"),
        ("Mai/2025", "Emissão"),
    ]

    for val, label in metrics:
        add_textbox(slide, c3_left, mm(cy3), c3_w, mm(8),
                    val,
                    font_size=14, bold=True, color=PURPLE,
                    align=PP_ALIGN.LEFT)
        cy3 += 7
        add_textbox(slide, c3_left, mm(cy3), c3_w, mm(5),
                    label,
                    font_size=6.5, bold=False, color=GRAY,
                    align=PP_ALIGN.LEFT)
        cy3 += 5

    cy3 += 2

    # --- Section F: MARCOS CONCLUÍDOS ---
    section_label(slide, c3_left, mm(cy3), c3_w, "MARCOS CONCLUÍDOS")
    cy3 += 5

    items_f = [
        "Aquisição do terreno (2020/2008) — a valor de mercado",
        "Licença Prévia emitida 2012 — renovada até mai/2031",
        "Licença de Instalação emitida jun/2024 — válida até jun/2026",
        "Alvará de Construção emitido — Projetos arquitetônicos concluídos",
    ]
    bullet_list(slide, c3_left, mm(cy3), c3_w, mm(26), items_f, font_size=7, color=WHITE)
    cy3 += 26

    # --- Section G: PRÓXIMAS ETAPAS ---
    section_label(slide, c3_left, mm(cy3), c3_w, "PRÓXIMAS ETAPAS")
    cy3 += 5

    items_g = [
        "Integralização final do equity — investidor 4%",
        "Estruturação e emissão do CRI (R$ 330MM)",
        "Lançamento Residencial Fase 1 — out/2025",
        "Início obras Hotel + Residencial — set/2025 / jan/2027",
        "Abertura Hotel Fasano — jan/2026",
    ]
    bullet_list(slide, c3_left, mm(cy3), c3_w, mm(30), items_g, font_size=7, color=WHITE)

    # -----------------------------------------------------------------------
    # 5. BOTTOM BAR
    # -----------------------------------------------------------------------
    bottom_top = 297 - BOTTOM_H - FOOTER_H
    add_rect(slide, mm(0), mm(bottom_top), mm(210), mm(BOTTOM_H), NAVY)

    add_textbox(slide, mm(5), mm(bottom_top + 1), mm(100), mm(BOTTOM_H - 2),
                "Aval cruzado das holdings e subholdings do Grupo Prima",
                font_size=6.5, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT)

    add_textbox(slide, mm(5), mm(bottom_top + 1), mm(200), mm(BOTTOM_H - 2),
                "Negociação Fasano concluída — Formalização em fase final",
                font_size=7, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)

    # -----------------------------------------------------------------------
    # 6. FOOTER
    # -----------------------------------------------------------------------
    footer_top = 297 - FOOTER_H
    add_rect(slide, mm(0), mm(footer_top), mm(210), mm(FOOTER_H), DARK)

    # Footer logo
    add_image(slide, logo_path, mm(3), mm(footer_top + 2), mm(20), mm(14))

    # Disclaimer text
    disclaimer = (
        "Material preparado pela Monte Bravo Investimentos para uso exclusivo e confidencial. "
        "Este documento não constitui oferta de valores mobiliários. "
        "As informações aqui contidas são baseadas em fontes consideradas confiáveis, "
        "porém não garantidas. Investimentos envolvem riscos e retornos passados não garantem resultados futuros. "
        "Distribuição ou reprodução proibida sem autorização expressa."
    )
    add_textbox(slide, mm(26), mm(footer_top + 2), mm(155), mm(FOOTER_H - 2),
                disclaimer,
                font_size=5.5, bold=False, color=GRAY,
                align=PP_ALIGN.CENTER)

    add_textbox(slide, mm(183), mm(footer_top + 6), mm(24), mm(6),
                "montebravo.com.br",
                font_size=6, bold=False, color=GOLD,
                align=PP_ALIGN.RIGHT)

    # -----------------------------------------------------------------------
    # Save
    # -----------------------------------------------------------------------
    out_path = "/home/user/pptx/fasano-mamucabo.pptx"
    prs.save(out_path)
    print(f"Saved: {out_path}")
    return out_path


if __name__ == "__main__":
    build()
