# -*- coding: utf-8 -*-
"""경량 PPTX->PNG 미리보기 렌더러 (PIL). 레이아웃/색/구조 검증용.
플랫 도형(autoshape/textbox/oval/line/picture/freeform) + 1단계 그룹 지원."""
import sys, math
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

PXIN = 96  # px per inch
EMU = 914400
FONT = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"

def load_font(sz):
    try: return ImageFont.truetype(FONT, max(8, int(sz)))
    except: return ImageFont.load_default()

def hexof(color):
    try: return str(color.rgb)
    except: return None

def rgb(h, default=(60,60,60)):
    if not h: return default
    try: return (int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
    except: return default

def emu_px(v): return int((v or 0)/EMU*PXIN)

def shape_fill(sh):
    try:
        if sh.fill.type==1: return rgb(hexof(sh.fill.fore_color), None)
    except: pass
    return None

def shape_line(sh):
    try:
        if sh.line.fill.type==1:
            c=rgb(hexof(sh.line.color), None)
            w=1
            try:
                if sh.line.width: w=max(1,int(sh.line.width/EMU*PXIN))
            except: pass
            return c,w
    except: pass
    return None,0

def run_color(r, default=(30,30,30)):
    try:
        if r.font.color is not None and r.font.color.type is not None:
            return rgb(str(r.font.color.rgb), default)
    except: pass
    return default

def draw_text(draw, sh, x, y, w, h):
    tf = sh.text_frame
    cy = y + 4
    for pa in tf.paragraphs:
        runs = pa.runs
        if not runs:
            cy += 16; continue
        text = "".join(r.text for r in runs)
        if not text.strip():
            cy += 14; continue
        r0 = runs[0]
        sz = 14
        if r0.font.size: sz = r0.font.size.pt
        col = run_color(r0)
        bold = bool(r0.font.bold)
        f = load_font(sz*PXIN/72)
        align = str(pa.alignment) if pa.alignment else ""
        # naive wrap (respect hard newlines)
        maxw = w - 10
        lines=[]
        for seg in text.split("\n"):
            line = ""
            for ch in seg:
                test = line+ch
                if draw.textlength(test, font=f) > maxw and line:
                    lines.append(line); line=ch
                else:
                    line=test
            lines.append(line)
        for ln in lines:
            tw = draw.textlength(ln, font=f)
            if "CENTER" in align: tx = x + (w-tw)/2
            elif "RIGHT" in align: tx = x + w - tw - 5
            else: tx = x + 5
            draw.text((tx, cy), ln, font=f, fill=col)
            if bold:
                draw.text((tx+1, cy), ln, font=f, fill=col)
            cy += int(sz*PXIN/72*1.25)

def draw_shape(draw, sh, ox=0, oy=0, sx=1.0, sy=1.0):
    st = sh.shape_type
    # tables
    if sh.has_table:
        tbl = sh.table
        x0 = ox + emu_px(sh.left); y0 = oy + emu_px(sh.top)
        col_w = [emu_px(c.width) for c in tbl.columns]
        row_h = [emu_px(r.height) for r in tbl.rows]
        cy = y0
        for ri, row in enumerate(tbl.rows):
            cx = x0
            for ci in range(len(col_w)):
                cell = tbl.cell(ri, ci)
                cw, ch = col_w[ci], row_h[ri]
                fill = None
                try:
                    if cell.fill.type == 1: fill = rgb(hexof(cell.fill.fore_color), None)
                except: pass
                draw.rectangle([cx, cy, cx+cw, cy+ch], fill=fill, outline=(210,210,205))
                # text
                txt = cell.text_frame
                sz = 10; col = (30,30,30); bold=False
                for pa in txt.paragraphs:
                    for r in pa.runs:
                        if r.font.size: sz = r.font.size.pt
                        if r.font.bold is not None: bold = r.font.bold
                        col = run_color(r, (30,30,30)); break
                    break
                f = load_font(sz*PXIN/72)
                s = cell.text.replace("\n"," ")
                # wrap
                maxw = cw-8; line=""; ty=cy+3
                for chh in s:
                    if draw.textlength(line+chh, font=f) > maxw and line:
                        draw.text((cx+4,ty), line, font=f, fill=col); ty+=int(sz*PXIN/72*1.2); line=chh
                    else: line+=chh
                if line: draw.text((cx+4,ty), line, font=f, fill=col)
                cx += cw
            cy += row_h[ri]
        return
    if st == MSO_SHAPE_TYPE.GROUP:
        # compute child transform from xml
        try:
            g = sh.element.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm')
        except: g=None
        # fallback: just recurse with same transform using shape offsets
        gx = ox + emu_px(sh.left)*1
        gy = oy + emu_px(sh.top)*1
        for c in sh.shapes:
            draw_shape(draw, c, ox, oy, sx, sy)
        return
    x = ox + emu_px(sh.left); y = oy + emu_px(sh.top)
    w = emu_px(sh.width); h = emu_px(sh.height)
    fill = shape_fill(sh); lc, lw = shape_line(sh)
    name = ""
    try: name = sh.name.lower()
    except: pass
    is_oval = False
    try:
        is_oval = sh.adjustments is not None and ("oval" in name or "ellipse" in name)
    except: pass
    # detect oval via auto_shape_type
    try:
        from pptx.enum.shapes import MSO_SHAPE
        if sh.auto_shape_type == MSO_SHAPE.OVAL: is_oval=True
    except: pass
    if st == MSO_SHAPE_TYPE.PICTURE:
        draw.rectangle([x,y,x+w,y+h], fill=(225,225,225), outline=(180,180,180))
        f=load_font(11); draw.text((x+6,y+6),"[IMG]",font=f,fill=(120,120,120))
        return
    if st == MSO_SHAPE_TYPE.LINE or "connector" in name:
        c = lc or (40,40,40)
        x1,y1,x2,y2 = x,y,x+w,y+h
        try:
            xf = sh._element.find('.//'+qn('a:xfrm'))
            if xf is not None:
                if xf.get('flipH')=='1': x1,x2 = x+w,x
                if xf.get('flipV')=='1': y1,y2 = y+h,y
        except Exception: pass
        draw.line([x1,y1,x2,y2], fill=c, width=max(1,lw))
        return
    # right triangle (대각선 분할) with rotation
    try:
        from pptx.enum.shapes import MSO_SHAPE as _MS
        if sh.auto_shape_type == _MS.RIGHT_TRIANGLE:
            rot = (sh.rotation or 0) % 360
            # base (rot=0): right angle bottom-left -> pts (x,y+h),(x+w,y+h),(x,y)
            pts = [(x,y+h),(x+w,y+h),(x,y)]
            cx,cy = x+w/2, y+h/2
            if rot:
                a = math.radians(rot)
                pts = [(cx+(px-cx)*math.cos(a)-(py-cy)*math.sin(a),
                        cy+(px-cx)*math.sin(a)+(py-cy)*math.cos(a)) for px,py in pts]
            draw.polygon(pts, fill=fill, outline=lc)
            return
    except Exception:
        pass
    if fill or lc:
        if is_oval:
            draw.ellipse([x,y,x+w,y+h], fill=fill, outline=lc, width=max(1,lw) if lc else 1)
        else:
            draw.rectangle([x,y,x+w,y+h], fill=fill, outline=lc, width=max(1,lw) if lc else 1)
    if sh.has_text_frame and sh.text_frame.text.strip():
        draw_text(draw, sh, x, y, w, h)

def render(path, outdir, prefix="s", only=None):
    import os
    os.makedirs(outdir, exist_ok=True)
    prs = Presentation(path)
    W = emu_px(prs.slide_width); H = emu_px(prs.slide_height)
    outs=[]
    for i, slide in enumerate(prs.slides):
        if only and (i+1) not in only: continue
        img = Image.new("RGB", (W,H), (255,255,255))
        draw = ImageDraw.Draw(img)
        for sh in slide.shapes:
            try: draw_shape(draw, sh)
            except Exception as e: pass
        p = os.path.join(outdir, f"{prefix}{i+1:02d}.png")
        img.save(p); outs.append(p)
    return outs

if __name__ == "__main__":
    src = sys.argv[1]
    out = sys.argv[2] if len(sys.argv)>2 else "/tmp/prev"
    only = [int(x) for x in sys.argv[3].split(",")] if len(sys.argv)>3 else None
    outs = render(src, out, only=only)
    print(f"rendered {len(outs)} -> {out}")
