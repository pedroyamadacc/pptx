"""
Fasano Mamucabo — A4 PPTX one-pager
Fiel ao HTML: mesmo conteúdo, mesma hierarquia visual.
"""
from pptx import Presentation
from pptx.util import Mm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os
from lxml import etree
from pptx.oxml.ns import qn

NAVY   = RGBColor(0x20,0x1A,0x47)
PURPLE = RGBColor(0x88,0x7C,0xD3)
GOLD   = RGBColor(0xC8,0xA9,0x6D)
GRAY   = RGBColor(0x7F,0x80,0x8F)
WHITE  = RGBColor(0xFF,0xFF,0xFF)
LGRAY  = RGBColor(0xE8,0xE8,0xF0)

def rgb(r,g,b): return RGBColor(r,g,b)

ASSETS = "/home/user/pptx/assets/"

def add_rect(slide, l,t,w,h, fill=None, line=None, line_w=Pt(0)):
    shp = slide.shapes.add_shape(1, l,t,w,h)
    shp.line.fill.background()
    if fill:
        shp.fill.solid(); shp.fill.fore_color.rgb = fill
    else:
        shp.fill.background()
    if line:
        shp.line.color.rgb = line
        shp.line.width = line_w
    else:
        shp.line.fill.background()
    return shp

def add_pic(slide, path, l,t,w,h):
    if os.path.exists(path):
        return slide.shapes.add_picture(path, l,t,w,h)

def txb(slide, l,t,w,h):
    tb = slide.shapes.add_textbox(l,t,w,h)
    tb.word_wrap = True
    return tb

def para(tf, text, size=Pt(7), bold=False, color=WHITE, align=PP_ALIGN.LEFT,
         space_before=Pt(0), space_after=Pt(0), italic=False):
    p = tf.add_paragraph()
    p.alignment = align
    p.space_before = space_before
    p.space_after  = space_after
    r = p.add_run()
    r.text = text
    r.font.size   = size
    r.font.bold   = bold
    r.font.color.rgb = color
    r.font.italic = italic
    r.font.name   = "Calibri"
    return p

def bullet_para(tf, text, size=Pt(6.5), color=WHITE, bold_part=None, space_before=Pt(1.5)):
    p = tf.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    p.space_before = space_before
    p.space_after  = Pt(0)
    r0 = p.add_run(); r0.text = "• "; r0.font.size=size; r0.font.color.rgb=PURPLE; r0.font.name="Calibri"
    if bold_part and bold_part in text:
        rb = p.add_run(); rb.text = bold_part; rb.font.size=size; rb.font.bold=True
        rb.font.color.rgb=NAVY; rb.font.name="Calibri"
        rest = text[len(bold_part):]
        r1 = p.add_run(); r1.text=rest; r1.font.size=size; r1.font.color.rgb=color; r1.font.name="Calibri"
    else:
        r1 = p.add_run(); r1.text=text; r1.font.size=size; r1.font.color.rgb=color; r1.font.name="Calibri"
    return p

def sec_label(tf, text):
    p = tf.add_paragraph()
    p.space_before=Pt(0); p.space_after=Pt(1)
    r=p.add_run(); r.text=text.upper(); r.font.size=Pt(5); r.font.bold=True
    r.font.color.rgb=GOLD; r.font.name="Calibri"

def sec_title(tf, text):
    p = tf.add_paragraph()
    p.space_before=Pt(0); p.space_after=Pt(2)
    r=p.add_run(); r.text=text; r.font.size=Pt(8); r.font.bold=True
    r.font.color.rgb=NAVY; r.font.name="Calibri"

def rule_line(slide, l,t,w):
    add_rect(slide, l,t,w,Mm(0.3), fill=GOLD)

def body_sec_header(slide, col_x, top, label, title, col_inner, col_pad):
    lx = col_x + col_pad
    tb = txb(slide, lx, top, col_inner, Mm(12))
    tf = tb.text_frame; tf.word_wrap=True
    sec_label(tf, label)
    sec_title(tf, title)
    rule_line(slide, lx, top+Mm(7.5), col_inner)
    return top + Mm(10)

# ── Presentation ─────────────────────────────────────────────────────────────
prs = Presentation()
prs.slide_width  = Mm(210)
prs.slide_height = Mm(297)

slide = prs.slides.add_slide(prs.slide_layouts[6])
W = Mm(210)
H = Mm(297)

add_rect(slide, 0,0,W,H, fill=rgb(0xF4,0xF4,0xF7))

# ── HEADER ───────────────────────────────────────────────────────────────────
HD_H = Mm(20)
add_rect(slide, 0,0,W,HD_H, fill=NAVY)

logo = ASSETS+"logo_mb_ib_white.png"
add_pic(slide, logo, Mm(4), Mm(3.5), Mm(28), Mm(13))

tb = txb(slide, Mm(36), Mm(1.5), Mm(130), HD_H)
tf = tb.text_frame; tf.word_wrap=True
p=tf.paragraphs[0]; p.space_before=Pt(0); p.space_after=Pt(1)
r=p.add_run(); r.text="Fasano Mamucabo"; r.font.size=Pt(16); r.font.bold=True
r.font.color.rgb=GOLD; r.font.name="Calibri"
para(tf,"Branded Residences & Hotel · Overview de Investimento",Pt(6.5),False,WHITE)

tb2 = txb(slide, Mm(155), Mm(3), Mm(50), HD_H)
tf2=tb2.text_frame
p2=tf2.paragraphs[0]; p2.alignment=PP_ALIGN.RIGHT; p2.space_before=Pt(0)
r2=p2.add_run(); r2.text="CONFIDENCIAL"; r2.font.size=Pt(6); r2.font.bold=True
r2.font.color.rgb=GOLD; r2.font.name="Calibri"
p3=tf2.add_paragraph(); p3.alignment=PP_ALIGN.RIGHT
r3=p3.add_run(); r3.text="Março 2026"; r3.font.size=Pt(6); r3.font.color.rgb=GRAY; r3.font.name="Calibri"

# ── HERO ──────────────────────────────────────────────────────────────────────
HERO_T = HD_H
HERO_H = Mm(32)
hero_img = ASSETS+"hero.jpg"
if os.path.exists(hero_img):
    add_pic(slide, hero_img, 0, HERO_T, W, HERO_H)
else:
    add_rect(slide, 0,HERO_T,W,HERO_H, fill=rgb(0x2A,0x1E,0x50))

# dark overlay for text readability
add_rect(slide, 0,HERO_T, Mm(115),HERO_H, fill=rgb(0x14,0x11,0x30))

tb_h = txb(slide, Mm(5), HERO_T+Mm(3), Mm(108), HERO_H-Mm(5))
tf_h = tb_h.text_frame; tf_h.word_wrap=True
p=tf_h.paragraphs[0]; p.space_before=Pt(0)
r=p.add_run(); r.text="Litoral Norte da Bahia · Costa dos Coqueiros · 108 km de Salvador"
r.font.size=Pt(5.5); r.font.color.rgb=GOLD; r.font.name="Calibri"
p2=tf_h.add_paragraph(); p2.space_before=Pt(3)
r2=p2.add_run(); r2.text="O maior destino turístico-imobiliário de luxo do Brasil"
r2.font.size=Pt(13); r2.font.bold=True; r2.font.color.rgb=WHITE; r2.font.name="Calibri"
p3=tf_h.add_paragraph(); p3.space_before=Pt(3)
r3=p3.add_run(); r3.text="Complexo Fasano · Mamucabo · Esplanada, BA · Licença de Instalação Emitida"
r3.font.size=Pt(5.5); r3.font.color.rgb=LGRAY; r3.font.name="Calibri"

tags = ["IPCA + 12,5% a.a.", "TIR 30,9%", "LI Emitida Jan/2024"]
for i,tag in enumerate(tags):
    tx = Mm(116) + i*Mm(31)
    ty = HERO_T + HERO_H - Mm(9)
    add_rect(slide, tx,ty, Mm(29),Mm(7), fill=NAVY, line=GOLD, line_w=Pt(0.5))
    tb_t=txb(slide, tx+Mm(1),ty+Mm(1), Mm(27),Mm(6))
    tf_t=tb_t.text_frame
    p=tf_t.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=tag; r.font.size=Pt(5.5); r.font.bold=True
    r.font.color.rgb=GOLD; r.font.name="Calibri"

# ── KPI BAR ──────────────────────────────────────────────────────────────────
KPI_T = HERO_T + HERO_H
KPI_H = Mm(14)
add_rect(slide, 0,KPI_T,W,KPI_H, fill=rgb(0x18,0x14,0x3A))

kpis = [
    ("R$411,6MM","VGV Total Residencial"),
    ("R$170MM",  "Volume CRI"),
    ("30,9%",    "TIR Equity (49%)"),
    ("48",       "Unidades Residenciais"),
    ("40",       "Unidades Hoteleiras"),
]
kw = W // len(kpis)
for i,(val,lbl) in enumerate(kpis):
    kx = i*kw
    tb_k = txb(slide, kx+Mm(1), KPI_T+Mm(1.5), kw-Mm(2), KPI_H-Mm(2))
    tf_k=tb_k.text_frame; tf_k.word_wrap=True
    p=tf_k.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=val
    r.font.size=Pt(11); r.font.bold=True
    r.font.color.rgb = GOLD if i==2 else WHITE
    r.font.name="Calibri"
    p2=tf_k.add_paragraph(); p2.alignment=PP_ALIGN.CENTER
    r2=p2.add_run(); r2.text=lbl; r2.font.size=Pt(5); r2.font.color.rgb=LGRAY; r2.font.name="Calibri"
    if i<len(kpis)-1:
        add_rect(slide, kx+kw-Mm(0.15), KPI_T+Mm(2), Mm(0.3), KPI_H-Mm(4), fill=PURPLE)

# ── TIMELINE ─────────────────────────────────────────────────────────────────
TL_T = KPI_T + KPI_H
TL_H = Mm(11)
add_rect(slide, 0,TL_T,W,TL_H, fill=rgb(0xF0,0xEE,0xF8))

timeline = [
    ("Out/2026","Lançamento Res. F1\nInício Obras Hotel",False),
    ("Abr/2027","Início Obras\nResidencial F1",False),
    ("Jan/2028","Lançamento\nResidencial F2",False),
    ("Jan/2029","Abertura\nHotel Fasano",True),
    ("Ago/2029","Entrega Res. F1\nPós-Chaves",False),
    ("Jun/2030","Entrega Res. F2\nFull Cash Sweep",False),
    ("Ago/2031","Encerramento\nOperação Dívida",True),
]
tw = W // len(timeline)
add_rect(slide, Mm(5), TL_T+Mm(5), W-Mm(10), Mm(0.4), fill=PURPLE)

for i,(date,ev,is_navy) in enumerate(timeline):
    tx = i*tw + tw//2
    dot_c = NAVY if is_navy else PURPLE
    add_rect(slide, tx-Mm(1.3), TL_T+Mm(3.8), Mm(2.6),Mm(2.6), fill=dot_c)
    tb_d=txb(slide, tx-Mm(11), TL_T+Mm(0.5), Mm(22), Mm(3.5))
    tf_d=tb_d.text_frame
    p=tf_d.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=date; r.font.size=Pt(5); r.font.bold=True
    r.font.color.rgb=NAVY; r.font.name="Calibri"
    tb_e=txb(slide, tx-Mm(12), TL_T+Mm(6.5), Mm(24), Mm(5))
    tf_e=tb_e.text_frame
    p=tf_e.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=ev; r.font.size=Pt(4.5); r.font.color.rgb=GRAY; r.font.name="Calibri"

# ── BODY ─────────────────────────────────────────────────────────────────────
BODY_T = TL_T + TL_H
BOT_H  = Mm(20)
FT_H   = Mm(14)
BODY_H = H - BODY_T - BOT_H - FT_H

add_rect(slide, 0,BODY_T,W,BODY_H, fill=WHITE)

COL_W     = W // 3
COL_PAD   = Mm(3)
COL_INNER = COL_W - COL_PAD*2

for i in [1,2]:
    add_rect(slide, i*COL_W-Mm(0.2), BODY_T+Mm(3), Mm(0.4), BODY_H-Mm(6), fill=LGRAY)

# ── COL 0: Tese + Prima/Fasano ────────────────────────────────────────────────
CX=0
cx0 = CX*COL_W
cy = BODY_T + Mm(2)

cy = body_sec_header(slide, cx0, cy, "Por que investir", "Tese de Investimento", COL_INNER, COL_PAD)

tese = [
    ("Marca Fasano",              "— Parceria com grupo centenário, sinônimo de luxo e hospitalidade premium na América Latina. Contrato negociado, formalização final em curso."),
    ("Destino exclusivo e escasso","— Mamucabo reservado exclusivamente para projetos AAA. 13 km de praias intocadas com baixíssima densidade."),
    ("Risco regulatório zerado",  "— Terreno desde 2009. LP emitida 2012, LI emitida jan/2024 e Alvará concluído. Investidor ingressa na fase de execução."),
    ("Garantias robustas",        "— AF de cotas, terrenos e recebíveis. Razão mínima de 120%. Full cash sweep pós-entrega residencial."),
    ("Track record Prima",        "— 20 anos, R$ 1,36 BI VGV em execução, R$ 280MM receita e R$ 138MM EBITDA em 2025. Premiado ADEMI-BA."),
]
tb_t=txb(slide, cx0+COL_PAD, cy, COL_INNER, Mm(38))
tf_t=tb_t.text_frame; tf_t.word_wrap=True
tf_t.paragraphs[0].space_before=Pt(0)
for bp,bt in tese:
    p=tf_t.add_paragraph(); p.space_before=Pt(1.5); p.space_after=Pt(0)
    r0=p.add_run(); r0.text="• "; r0.font.size=Pt(6); r0.font.color.rgb=PURPLE; r0.font.name="Calibri"
    rb=p.add_run(); rb.text=bp+" "; rb.font.size=Pt(6); rb.font.bold=True; rb.font.color.rgb=NAVY; rb.font.name="Calibri"
    r1=p.add_run(); r1.text=bt; r1.font.size=Pt(6); r1.font.color.rgb=rgb(0x33,0x33,0x55); r1.font.name="Calibri"
cy += Mm(39)

# images
for i,img in enumerate([ASSETS+"aerial.jpg", ASSETS+"pool.jpg"]):
    ix = cx0+COL_PAD + i*(COL_INNER//2+Mm(1))
    add_pic(slide, img, ix, cy, COL_INNER//2-Mm(1), Mm(17)) or add_rect(slide,ix,cy,COL_INNER//2-Mm(1),Mm(17),fill=LGRAY)
cy += Mm(19)

cy = body_sec_header(slide, cx0, cy, "Desenvolvedor & Operador", "Grupo Prima & Fasano", COL_INNER, COL_PAD)

# Stats
stats=[("R$280MM","Receita 2025"),("R$138MM","EBITDA 2025"),("R$1,36BI","VGV em Exec.")]
sw=COL_INNER//3
for i,(v,l) in enumerate(stats):
    sx=cx0+COL_PAD+i*sw
    add_rect(slide, sx+Mm(0.3),cy, sw-Mm(0.6),Mm(8.5), fill=rgb(0xF0,0xEE,0xF8))
    tb_s=txb(slide, sx+Mm(0.3),cy+Mm(0.5), sw-Mm(0.6),Mm(8.5))
    tf_s=tb_s.text_frame
    p=tf_s.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=v; r.font.size=Pt(7); r.font.bold=True; r.font.color.rgb=NAVY; r.font.name="Calibri"
    p2=tf_s.add_paragraph(); p2.alignment=PP_ALIGN.CENTER
    r2=p2.add_run(); r2.text=l; r2.font.size=Pt(4.5); r2.font.color.rgb=GRAY; r2.font.name="Calibri"
cy += Mm(10)

tb_p=txb(slide, cx0+COL_PAD, cy, COL_INNER, Mm(12))
tf_p=tb_p.text_frame; tf_p.word_wrap=True
tf_p.paragraphs[0].space_before=Pt(0)
for txt in ["Fundado em 2005 · 6 famílias industriais espanholas · 570 colaboradores · Governança formal",
            "Hotel Fasano Salvador · 3 hotéis em operação · 4 em desenvolvimento em Mamucabo/Baixio"]:
    bullet_para(tf_p, txt, Pt(6), color=rgb(0x33,0x33,0x55))
cy += Mm(10)

add_rect(slide, cx0+COL_PAD, cy, COL_INNER, Mm(0.3), fill=LGRAY)
cy += Mm(2)

tb_f=txb(slide, cx0+COL_PAD, cy, COL_INNER, Mm(16))
tf_f=tb_f.text_frame; tf_f.word_wrap=True
p=tf_f.paragraphs[0]; p.space_before=Pt(0)
r=p.add_run(); r.text="Fasano — Desde 1902"; r.font.size=Pt(6.5); r.font.bold=True; r.font.color.rgb=NAVY; r.font.name="Calibri"
for txt in ["Marca centenária de luxo e gastronomia premium · Hotéis icônicos em SP, RJ, Punta del Este e Salvador",
            "Agrega prestígio internacional e forte valorização imobiliária ao complexo Mamucabo"]:
    bullet_para(tf_f, txt, Pt(6), color=rgb(0x33,0x33,0x55))
cy += Mm(14)

for i,img in enumerate([ASSETS+"entrance.jpg", ASSETS+"room.jpg"]):
    ix = cx0+COL_PAD + i*(COL_INNER//2+Mm(1))
    add_pic(slide, img, ix, cy, COL_INNER//2-Mm(1), Mm(16)) or add_rect(slide,ix,cy,COL_INNER//2-Mm(1),Mm(16),fill=LGRAY)

# ── COL 1: CRI + Destino ──────────────────────────────────────────────────────
CX=1
cx1=CX*COL_W
cy1=BODY_T+Mm(2)

cy1=body_sec_header(slide, cx1, cy1, "Operação Financeira", "Estrutura de Investimento · CRI", COL_INNER, COL_PAD)

cri_rows=[
    ("Instrumento", "CRI · Oferta pública (Res. CVM 160/22)"),
    ("Volume",      "R$ 170MM · 5 integralizações semestrais"),
    ("Devedoras",   "FAS Mamucabo Hotelaria + Empreendimentos Ltda."),
    ("Remuneração", "IPCA + 12,5% a.a."),
    ("Juros",       "Mensal, sem carência"),
    ("Prazo",       "Até 84 meses (ou 12m após entrega)"),
    ("Amortização", "Full cash sweep pós-entrega Residencial"),
    ("Duration",    "3,06 anos · Cobertura mínima 120%"),
]
row_h=Mm(4)
for j,(k,v) in enumerate(cri_rows):
    ry=cy1+j*row_h; rx=cx1+COL_PAD
    add_rect(slide, rx,ry, COL_INNER,row_h-Mm(0.2), fill=rgb(0xF5,0xF3,0xFB) if j%2==0 else WHITE)
    tb_r=txb(slide, rx+Mm(0.8),ry+Mm(0.4), COL_INNER-Mm(1.5),row_h-Mm(0.5))
    tf_r=tb_r.text_frame
    p=tf_r.paragraphs[0]; p.space_before=Pt(0)
    rk=p.add_run(); rk.text=k+"  "; rk.font.size=Pt(5.5); rk.font.color.rgb=GRAY; rk.font.name="Calibri"
    rv2=p.add_run(); rv2.text=v; rv2.font.size=Pt(5.5); rv2.font.bold=True
    rv2.font.color.rgb=NAVY if k=="Remuneração" else rgb(0x22,0x1A,0x50); rv2.font.name="Calibri"
cy1+=len(cri_rows)*row_h+Mm(1.5)

add_rect(slide, cx1+COL_PAD,cy1, COL_INNER,Mm(10), fill=NAVY)
tb_tir=txb(slide, cx1+COL_PAD+Mm(1),cy1+Mm(0.5), COL_INNER-Mm(2),Mm(10))
tf_tir=tb_tir.text_frame; tf_tir.word_wrap=True
p=tf_tir.paragraphs[0]
r=p.add_run(); r.text="30,9% a.a."; r.font.size=Pt(14); r.font.bold=True; r.font.color.rgb=GOLD; r.font.name="Calibri"
p2=tf_tir.add_paragraph()
r2=p2.add_run(); r2.text="TIR Equity — Investidor (49%) · Prazo 2026–2038"; r2.font.size=Pt(5.5); r2.font.color.rgb=WHITE; r2.font.name="Calibri"
cy1+=Mm(12)

eq_w=COL_INNER//2-Mm(1)
for i,(pct,name,amt,c) in enumerate([
    ("49%","Investidor","R$ 44,2MM",PURPLE),
    ("51%","Grupo Prima","R$ 46,0MM",rgb(0x44,0x3C,0x7A)),
]):
    ex=cx1+COL_PAD+i*(eq_w+Mm(2))
    add_rect(slide, ex,cy1, eq_w,Mm(12), fill=c)
    tb_eq=txb(slide, ex+Mm(1),cy1+Mm(0.5), eq_w-Mm(2),Mm(12))
    tf_eq=tb_eq.text_frame
    p=tf_eq.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=pct; r.font.size=Pt(13); r.font.bold=True; r.font.color.rgb=WHITE; r.font.name="Calibri"
    p2=tf_eq.add_paragraph(); p2.alignment=PP_ALIGN.CENTER
    r2=p2.add_run(); r2.text=name; r2.font.size=Pt(5.5); r2.font.color.rgb=WHITE; r2.font.name="Calibri"
    p3=tf_eq.add_paragraph(); p3.alignment=PP_ALIGN.CENTER
    r3=p3.add_run(); r3.text=amt; r3.font.size=Pt(5.5); r3.font.color.rgb=GOLD; r3.font.name="Calibri"
cy1+=Mm(14)

tb_en=txb(slide, cx1+COL_PAD,cy1, COL_INNER,Mm(5))
tf_en=tb_en.text_frame
p=tf_en.paragraphs[0]; p.space_before=Pt(0)
r=p.add_run(); r.text="Equity total R$ 90,2MM. R$ 32,9MM já investidos pelo Prima (terreno + licenças + projetos)."
r.font.size=Pt(5); r.font.color.rgb=GRAY; r.font.name="Calibri"
cy1+=Mm(7)

cy1=body_sec_header(slide, cx1, cy1, "Localização & Destino", "Mamucabo · Destino AAA", COL_INNER, COL_PAD)

tb_loc=txb(slide, cx1+COL_PAD,cy1, COL_INNER,Mm(14))
tf_loc=tb_loc.text_frame; tf_loc.word_wrap=True
tf_loc.paragraphs[0].space_before=Pt(0)
p=tf_loc.paragraphs[0]; p.space_before=Pt(0)
r0=p.add_run(); r0.text="• "; r0.font.size=Pt(6); r0.font.color.rgb=PURPLE; r0.font.name="Calibri"
rb=p.add_run(); rb.text="Esplanada, BA"; rb.font.size=Pt(6); rb.font.bold=True; rb.font.color.rgb=NAVY; rb.font.name="Calibri"
r1=p.add_run(); r1.text=" · 108 km de Salvador · 13 km de praias intocadas · Ecossistemas de lagoas e mata nativa"; r1.font.size=Pt(6); r1.font.color.rgb=rgb(0x33,0x33,0x55); r1.font.name="Calibri"
bullet_para(tf_loc,"Área licenciada: +63MM m² · Comercializável: +25MM m² · Pista de pouso, heliponto e utilities completas",Pt(6),color=rgb(0x33,0x33,0x55))
cy1+=Mm(10)

add_rect(slide, cx1+COL_PAD,cy1, COL_INNER,Mm(0.3), fill=LGRAY)
cy1+=Mm(2)

tb_br=txb(slide, cx1+COL_PAD,cy1, COL_INNER,Mm(5))
tf_br=tb_br.text_frame
p=tf_br.paragraphs[0]; p.space_before=Pt(0)
r=p.add_run(); r.text="Branded Residences · 48 Unidades · VGV R$ 411,6MM"
r.font.size=Pt(6.5); r.font.bold=True; r.font.color.rgb=NAVY; r.font.name="Calibri"
cy1+=Mm(6)

for i,(v,l) in enumerate([("F1 · 22 unid.","R$ 156MM · Out/2026"),("F2 · 26 unid.","R$ 285MM · Jan/2028")]):
    px=cx1+COL_PAD+i*(COL_INNER//2+Mm(1))
    add_rect(slide, px,cy1, COL_INNER//2-Mm(1),Mm(9), fill=rgb(0xF5,0xE8,0xD0))
    tb_ph=txb(slide, px+Mm(0.5),cy1+Mm(0.5), COL_INNER//2-Mm(2),Mm(8))
    tf_ph=tb_ph.text_frame
    p=tf_ph.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r=p.add_run(); r.text=v; r.font.size=Pt(7); r.font.bold=True; r.font.color.rgb=NAVY; r.font.name="Calibri"
    p2=tf_ph.add_paragraph(); p2.alignment=PP_ALIGN.CENTER
    r2=p2.add_run(); r2.text=l; r2.font.size=Pt(5); r2.font.color.rgb=GRAY; r2.font.name="Calibri"
cy1+=Mm(11)

tb_note=txb(slide, cx1+COL_PAD,cy1, COL_INNER,Mm(5))
tf_note=tb_note.text_frame
p=tf_note.paragraphs[0]; p.space_before=Pt(0)
r=p.add_run(); r.text="15+10 Casas Rental Pool · 17 Lotes 2.200m² · 16 Lotes 2.500m² · Terreno 260.350m²"
r.font.size=Pt(5); r.font.color.rgb=GRAY; r.font.name="Calibri"
cy1+=Mm(7)

for i,img in enumerate([ASSETS+"villa.jpg", ASSETS+"lobby.jpg"]):
    ix=cx1+COL_PAD+i*(COL_INNER//2+Mm(1))
    add_pic(slide, img, ix, cy1, COL_INNER//2-Mm(1), Mm(16)) or add_rect(slide,ix,cy1,COL_INNER//2-Mm(1),Mm(16),fill=LGRAY)

# ── COL 2: dark right panel (Garantias + Marcos + Etapas) ─────────────────────
CX=2
cx2=CX*COL_W
add_rect(slide, cx2,BODY_T, COL_W,BODY_H, fill=rgb(0x20,0x1A,0x47))

cy2=BODY_T+Mm(2)

garantias=[
    "AF de 100% das cotas das Devedoras",
    "AF terreno Residencial (260.350m²) c/ conversão para AF de unidades",
    "AF terreno Hoteleiro (306.017m²)",
    "Cessão fiduciária de recebíveis do Residencial e das receitas hoteleiras",
    "Aval cruzado das holdings e subholdings do Grupo Prima",
]
marcos=[
    "Aquisição de terreno (2008/2009) — a valor de mercado",
    "Licença Prévia emitida 2012 — renovada até mar/2030",
    "Licença de Instalação emitida jan/2024 — válida até jan/2030",
    "Alvará de Construção emitido · Projetos arquitetônicos concluídos",
    "Negociação Fasano concluída — formalização em fase final",
]
etapas=[
    "Integralização final do equity — investidor 49%",
    "Estruturação e emissão do CRI (R$ 170MM)",
    "Lançamento Residencial Fase 1 — out/2026",
    "Início obras Hotel + Residencial — out/2026 / abr/2027",
    "Abertura Hotel Fasano — jan/2029",
]

for label,items in [("Garantias CRI",garantias),("Marcos Concluídos",marcos),("Próximas Etapas",etapas)]:
    tb_l=txb(slide, cx2+COL_PAD,cy2, COL_INNER,Mm(5))
    tf_l=tb_l.text_frame; sec_label(tf_l, label)
    rule_line(slide, cx2+COL_PAD, cy2+Mm(4.2), COL_INNER)
    cy2+=Mm(6)
    tb_i=txb(slide, cx2+COL_PAD,cy2, COL_INNER,Mm(25))
    tf_i=tb_i.text_frame; tf_i.word_wrap=True
    tf_i.paragraphs[0].space_before=Pt(0)
    for item in items:
        bullet_para(tf_i, item, Pt(6))
    cy2+=Mm(27)
    add_rect(slide, cx2+COL_PAD,cy2, COL_INNER,Mm(0.3), fill=PURPLE)
    cy2+=Mm(2)

# ── BOTTOM STRIP ─────────────────────────────────────────────────────────────
BOT_T=H-BOT_H-FT_H
add_rect(slide, 0,BOT_T,W,BOT_H, fill=rgb(0x18,0x14,0x3A))

for i,(lbl,items) in enumerate([("Garantias CRI",garantias),("Marcos Concluídos",marcos),("Próximas Etapas",etapas)]):
    bx=i*COL_W+COL_PAD
    tb_bl=txb(slide, bx,BOT_T+Mm(1), COL_INNER,Mm(4))
    tf_bl=tb_bl.text_frame; sec_label(tf_bl, lbl)
    tb_bi=txb(slide, bx,BOT_T+Mm(5), COL_INNER,BOT_H-Mm(6))
    tf_bi=tb_bi.text_frame; tf_bi.word_wrap=True
    tf_bi.paragraphs[0].space_before=Pt(0)
    for item in items:
        bullet_para(tf_bi, item, Pt(5.5))

# ── FOOTER ───────────────────────────────────────────────────────────────────
FT_T=H-FT_H
add_rect(slide, 0,FT_T,W,FT_H, fill=NAVY)
add_pic(slide, logo, Mm(4),FT_T+Mm(2), Mm(22),Mm(10))

tb_disc=txb(slide, Mm(30),FT_T+Mm(2), Mm(140),FT_H-Mm(4))
tf_disc=tb_disc.text_frame; tf_disc.word_wrap=True
p=tf_disc.paragraphs[0]; p.space_before=Pt(0)
r=p.add_run()
r.text=("Material elaborado pela Monte Bravo Investment Banking para fins informativos. "
        "Não constitui oferta ou recomendação de investimento. "
        "Destinado exclusivamente a investidores qualificados (Res. CVM 30). "
        "Informações baseadas em documentos do Grupo Prima. "
        "Investimentos em CRI e participações imobiliárias envolvem riscos. "
        "Vedada a reprodução sem autorização prévia.")
r.font.size=Pt(4.5); r.font.color.rgb=GRAY; r.font.name="Calibri"

tb_meta=txb(slide, Mm(175),FT_T+Mm(3), Mm(30),Mm(8))
tf_meta=tb_meta.text_frame
p=tf_meta.paragraphs[0]; p.alignment=PP_ALIGN.RIGHT
r=p.add_run(); r.text="Versão Indicativa\nMarço 2026"; r.font.size=Pt(5); r.font.color.rgb=GRAY; r.font.name="Calibri"

# ── Save ─────────────────────────────────────────────────────────────────────
out="/home/user/pptx/fasano-mamucabo.pptx"
prs.save(out)
print(f"Saved → {out}  ({os.path.getsize(out)//1024} KB)")
