# -*- coding: utf-8 -*-
"""올더베러 클린 템플릿 컴포넌트 라이브러리.
레드 비즈니스 템플릿 구조 + 올더베러 브랜드 색/폰트(Pretendard)."""
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---- brand tokens (공식 가이드: 메인=딥브라운 주색) ----
MAIN    = RGBColor(0x3D,0x2B,0x1B)  # 메인 컬러 (주색)
BROWN2  = RGBColor(0x6B,0x4E,0x2A)  # 강조색 1
GREENA  = RGBColor(0x0E,0x5A,0x30)  # 강조색 2 (그린)
OLIVE   = RGBColor(0x9A,0x9A,0x3C)  # 서브 컬러 1
LIME    = RGBColor(0xC7,0xCE,0x3E)  # 서브 컬러 2
PINK    = RGBColor(0xF0,0xA4,0x93)  # 서브 컬러 3
BEIGE   = RGBColor(0xF3,0xF2,0xE7)
INK     = RGBColor(0x14,0x14,0x12)
GRAY    = RGBColor(0x86,0x86,0x80)
LGRAY   = RGBColor(0xC9,0xC4,0xB8)
WHITE   = RGBColor(0xFF,0xFF,0xFF)
# 컴포넌트 코드 호환 별칭: GREEN=주색(브라운), GREEN_B=그린강조, YELLOW=라임
GREEN   = MAIN
GREEN_B = GREENA
YELLOW  = LIME
FONT = "Pretendard"

def _set_font(run, size, color, bold=False, font=FONT, spacing=None):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    # set East Asian font too
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn('a:ea'))
    if ea is None:
        ea = rPr.makeelement(qn('a:ea'), {}); rPr.append(ea)
    ea.set('typeface', font)
    if spacing is None:
        spacing = -1.0  # 가이드: 자간 좁게 1pt
    rPr.set('spc', str(int(spacing*100)))

def textbox(slide, x, y, w, h, lines, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, wrap=True):
    """lines: list of (text, size, color, bold) or (text,size,color,bold,spacing)"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame; tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for m in (tf.margin_left, ): pass
    tf.margin_left=Pt(0); tf.margin_right=Pt(0); tf.margin_top=Pt(0); tf.margin_bottom=Pt(0)
    first=True
    for spec in lines:
        text,size,color,bold = spec[0],spec[1],spec[2],spec[3]
        sp = spec[4] if len(spec)>4 else None
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first=False
        p.alignment = align
        run = p.add_run(); run.text = text
        _set_font(run, size, color, bold, spacing=sp)
    return tb

def rect(slide, x,y,w,h, fill=None, line=None, line_w=1.0, rounded=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE,
                                 Inches(x),Inches(y),Inches(w),Inches(h))
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb = fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb = line; shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    if rounded:
        try: shp.adjustments[0]=0.08
        except: pass
    return shp

def oval(slide, x,y,d, fill=None, line=None, line_w=1.0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x),Inches(y),Inches(d),Inches(d))
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb=fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb=line; shp.line.width=Pt(line_w)
    shp.shadow.inherit=False
    return shp

def hline(slide, x,y,w, color=GREEN, weight=2.0):
    ln = slide.shapes.add_connector(2, Inches(x),Inches(y),Inches(x+w),Inches(y))
    ln.line.color.rgb=color; ln.line.width=Pt(weight); ln.shadow.inherit=False
    return ln

def rtri(slide, x,y,w,h, rot=0, fill=GREEN):
    """right triangle (대각선 분할용)"""
    shp = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, Inches(x),Inches(y),Inches(w),Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb=fill; shp.line.fill.background()
    shp.rotation=rot; shp.shadow.inherit=False
    return shp

# ---- composite components ----
def section_label(slide, word1, word2="", sub=""):
    """top-center: 'WORD1 word2' + green underline (template 헤더)"""
    full = word1 + (("  "+word2) if word2 else "")
    tb = slide.shapes.add_textbox(Inches(0), Inches(0.34), Inches(13.33), Inches(0.34))
    tf=tb.text_frame; tf.word_wrap=False
    tf.margin_top=Pt(0); tf.margin_bottom=Pt(0)
    p=tf.paragraphs[0]; p.alignment=PP_ALIGN.CENTER
    r1=p.add_run(); r1.text=word1+" "; _set_font(r1,15,GREEN,True,spacing=1.5)
    if word2:
        r2=p.add_run(); r2.text=word2; _set_font(r2,15,INK,True,spacing=1.5)
    hline(slide, 13.33/2-0.32, 0.74, 0.64, GREEN, 2.2)
    if sub:
        textbox(slide, 0, 0.82, 13.33, 0.3, [(sub,11,GRAY,False)], align=PP_ALIGN.CENTER)

def page_num(slide, n):
    textbox(slide, 13.33/2-0.5, 7.04, 1.0, 0.3, [(str(n),10,LGRAY,False)], align=PP_ALIGN.CENTER)

def num_circle(slide, cx, y, num, title, desc, d=0.92):
    oval(slide, cx-d/2, y, d, fill=GREEN)
    textbox(slide, cx-d/2, y, d, d, [(num,22,WHITE,True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    textbox(slide, cx-1.4, y+d+0.12, 2.8, 0.34, [(title,14,INK,True)], align=PP_ALIGN.CENTER)
    textbox(slide, cx-1.5, y+d+0.5, 3.0, 0.9, [(desc,11,GRAY,False)], align=PP_ALIGN.CENTER)

def icon_card(slide, x,y,w,h, title, desc):
    rect(slide, x,y,w,h, fill=WHITE, line=LGRAY, line_w=1.0, rounded=True)
    oval(slide, x+0.28, y+0.28, 0.5, fill=GREEN)
    textbox(slide, x+0.95, y+0.28, w-1.1, 0.4, [(title,13,INK,True)], anchor=MSO_ANCHOR.MIDDLE)
    textbox(slide, x+0.3, y+0.95, w-0.6, h-1.1, [(desc,10.5,GRAY,False)])

def vbar_chart(slide, x, y, w, h, data):
    """data: list of (label,value0-1). green bars."""
    n=len(data); gap=0.18; bw=(w-gap*(n-1))/n
    base=y+h
    hline(slide, x, base, w, LGRAY, 1.0)
    for i,(lab,val) in enumerate(data):
        bx=x+i*(bw+gap); bh=h*val
        rect(slide, bx, base-bh, bw, bh, fill=GREEN if i%2==0 else GREEN_B)
        textbox(slide, bx-0.1, base-bh-0.32, bw+0.2, 0.3, [(f"{int(val*100)}%",12,INK,True)],align=PP_ALIGN.CENTER)
        textbox(slide, bx-0.1, base+0.08, bw+0.2, 0.5, [(lab,10,GRAY,False)],align=PP_ALIGN.CENTER)

# ===== extended components =====
from pptx.oxml.ns import qn as _qn

def content_header(slide, label_kr, label_en, headline, sub=""):
    section_label(slide, label_kr, label_en)
    textbox(slide, 0.9, 1.08, 11.53, 0.66, [(headline,26,INK,True)], align=PP_ALIGN.CENTER)
    if sub:
        textbox(slide, 1.4, 1.82, 10.53, 0.5, [(sub,12,GRAY,False)], align=PP_ALIGN.CENTER)

def divider(slide, part_no, en, title, sub=""):
    rect(slide, 0,0,13.33,7.5, fill=MAIN)        # 메인 브라운 주색 필드
    rtri(slide, -1.2, 0, 9.8, 7.5, rot=0, fill=GREENA)  # 그린 대각선 강조
    textbox(slide, 0.95, 2.5, 7, 0.45, [("PART "+part_no+"   "+en.upper(),15,YELLOW,True,3)])
    textbox(slide, 0.9, 3.0, 9, 1.4, [(title,38,WHITE,True)])
    if sub:
        textbox(slide, 0.95, 4.35, 8.5, 0.6, [(sub,14,RGBColor(0xD9,0xE5,0xDD),False)])
    oval(slide, 10.95, 0.85, 1.35, fill=WHITE)
    textbox(slide, 10.95, 0.85, 1.35, 1.35, [(part_no,26,GREEN,True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

def arrow(slide, x, y, w=0.4):
    shp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(x), Inches(y), Inches(w), Inches(0.28))
    shp.fill.solid(); shp.fill.fore_color.rgb=GREEN_B; shp.line.fill.background(); shp.shadow.inherit=False
    return shp

def step_flow(slide, x, y, w, steps, num=True):
    """steps: list of (title, desc). 가로 넘버드 스텝 + chevron."""
    n=len(steps); gap=0.5; bw=(w-gap*(n-1))/n
    for i,(t,d) in enumerate(steps):
        bx=x+i*(bw+gap)
        if num:
            oval(slide, bx+bw/2-0.34, y, 0.68, fill=GREEN)
            textbox(slide, bx+bw/2-0.34, y, 0.68, 0.68, [(f"{i+1:02d}",17,WHITE,True)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        textbox(slide, bx-0.1, y+0.8, bw+0.2, 0.5, [(t,13.5,INK,True)], align=PP_ALIGN.CENTER)
        textbox(slide, bx-0.05, y+1.32, bw+0.1, 1.2, [(d,10.5,GRAY,False)], align=PP_ALIGN.CENTER)
        if i<n-1:
            arrow(slide, bx+bw+0.05, y+0.2, gap-0.1)

def two_col(slide, y, h, left, right, gap=0.6):
    """left/right: dict {tag,title,rows:[(label,text)]}"""
    x=0.9; w=(13.33-1.8-gap)/2
    for col,(x0) in [(left,x),(right,x+w+gap)]:
        rect(slide, x0, y, w, h, fill=WHITE, line=LGRAY, line_w=1.0, rounded=True)
        rect(slide, x0, y, w, 0.62, fill=GREEN, rounded=False)
        textbox(slide, x0+0.3, y, w-0.6, 0.62, [(col["title"],15,WHITE,True)], anchor=MSO_ANCHOR.MIDDLE)
        cy=y+0.85
        for lab,txt in col["rows"]:
            textbox(slide, x0+0.3, cy, w-0.6, 0.3, [(lab,11,GREEN,True)])
            tb=textbox(slide, x0+0.3, cy+0.3, w-0.6, 0.9, [(txt,10.5,GRAY,False)])
            cy+=1.18

def stat_badge(slide, x, y, w, num, label, color=GREEN, num_size=32):
    textbox(slide, x, y, w, 0.7, [(num,num_size,color,True)], align=PP_ALIGN.CENTER, wrap=False)
    textbox(slide, x, y+0.74, w, 0.5, [(label,11.5,INK,False)], align=PP_ALIGN.CENTER)

def bullets(slide, x, y, w, h, items, size=11.5, color=INK, gap_extra=0.0, mark="·"):
    tb=slide.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h))
    tf=tb.text_frame; tf.word_wrap=True
    tf.margin_left=Pt(0);tf.margin_right=Pt(0);tf.margin_top=Pt(0);tf.margin_bottom=Pt(0)
    first=True
    for it in items:
        p=tf.paragraphs[0] if first else tf.add_paragraph(); first=False
        p.space_after=Pt(6+gap_extra*72)
        r=p.add_run(); r.text=f"{mark}  {it}" if mark else it
        _set_font(r, size, color, False)
    return tb

def styled_table(slide, x, y, w, data, col_w=None, header=True, fs=10, h_fs=11, row_h=0.34, first_col_bold=False):
    rows=len(data); cols=len(data[0])
    from pptx.util import Inches as _I
    gtbl=slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w), Inches(row_h*rows))
    tbl=gtbl.table
    # disable default style banding
    tbl.first_row=False; tbl.horz_banding=False
    if col_w:
        tot=sum(col_w)
        for i,c in enumerate(tbl.columns): c.width=Inches(w*col_w[i]/tot)
    for ri,row in enumerate(data):
        tbl.rows[ri].height=Inches(row_h)
        for ci,val in enumerate(row):
            cell=tbl.cell(ri,ci)
            cell.margin_left=Pt(5);cell.margin_right=Pt(5);cell.margin_top=Pt(2);cell.margin_bottom=Pt(2)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            if header and ri==0:
                cell.fill.solid(); cell.fill.fore_color.rgb=GREEN
                col=WHITE; bold=True; sz=h_fs
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb=WHITE if ri%2==1 else BEIGE
                col=INK; bold=(first_col_bold and ci==0); sz=fs
            tf=cell.text_frame; tf.word_wrap=True
            p=tf.paragraphs[0]; p.alignment=PP_ALIGN.LEFT
            r=p.add_run(); r.text=str(val); _set_font(r, sz, col, bold)
    return gtbl
