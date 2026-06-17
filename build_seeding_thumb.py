# -*- coding: utf-8 -*-
"""멜라나잇 구미 — 마이크로 인플루언서 시딩 릴스 썸네일 목업 (9:16)."""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import math

W,H = 1080,1920
F = "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc"
def font(sz): return ImageFont.truetype(F, sz)

def btext(d, xy, s, fnt, fill, anchor="la", bold=0, spacing=0):
    if spacing and len(s)>1:
        # manual letter spacing
        x,y=xy;
        for ch in s:
            d.text((x,y),ch,font=fnt,fill=fill,anchor="la")
            x+=d.textlength(ch,font=fnt)+spacing
        return
    for dx in range(-bold,bold+1):
        for dy in range(-bold,bold+1):
            d.text((xy[0]+dx,xy[1]+dy), s, font=fnt, fill=fill, anchor=anchor)

# ---- background: night gradient ----
bg = Image.new("RGB",(W,H))
top=(58,28,74); bot=(22,15,26)
for y in range(H):
    t=y/H
    r=int(top[0]+(bot[0]-top[0])*t); g=int(top[1]+(bot[1]-top[1])*t); b=int(top[2]+(bot[2]-top[2])*t)
    for x in range(0,W,W): pass
    ImageDraw.Draw(bg).line([(0,y),(W,y)], fill=(r,g,b))

# warm lamp glow (upper right)
glow=Image.new("L",(W,H),0); gd=ImageDraw.Draw(glow)
gd.ellipse([W-560,-260,W+260,560], fill=120)
glow=glow.filter(ImageFilter.GaussianBlur(160))
warm=Image.new("RGB",(W,H),(252,226,150))
bg=Image.composite(warm,bg,glow)

# soft bokeh dots
dots=Image.new("RGBA",(W,H),(0,0,0,0)); dd=ImageDraw.Draw(dots)
import random; random.seed(7)
for _ in range(26):
    x=random.randint(0,W); y=random.randint(0,int(H*0.55)); rr=random.randint(6,26)
    a=random.randint(20,70)
    dd.ellipse([x-rr,y-rr,x+rr,y+rr], fill=(255,238,190,a))
dots=dots.filter(ImageFilter.GaussianBlur(3))
bg=Image.alpha_composite(bg.convert("RGBA"),dots).convert("RGB")

d=ImageDraw.Draw(bg)

# ---- top bar: creator handle + 광고 badge ----
# avatar
d.ellipse([54,70,150,166], fill=(227,90,110))
btext(d,(102,118),"Y",font(54),(255,255,255),anchor="mm",bold=1)
btext(d,(170,84),"yujin.nightlog",font(40),(255,255,255),bold=1)
btext(d,(170,132),"팔로워 2.4만 · 웰니스 데일리로그",font(28),(220,210,225))
# 광고 badge
bw=190
d.rounded_rectangle([W-bw-54,78,W-54,150],28,fill=(252,255,161))
btext(d,(W-bw/2-54,114),"광고·유료광고",font(30),(60,45,30),anchor="mm",bold=1)

# ---- hook text ----
btext(d,(64,300),"잠 들기 전, 마지막으로…",font(52),(252,255,200),bold=1)
btext(d,(60,378),"할 일이 ",font(118),(255,255,255),bold=2)
wlen=d.textlength("할 일이 ",font(118))
btext(d,(60+wlen,378),"있어요",font(118),(252,255,161),bold=2)
# accent underline
d.rounded_rectangle([66,532,360,548],8,fill=(252,255,161))
btext(d,(66,580),"불 끄기 전, 나만의 밤 마무리 리추얼",font(40),(232,222,238))

# ---- product pouch (simplified packaging) ----
px,py,pw,ph = 300,820,480,720
# pouch body cream
d.rounded_rectangle([px,py,px+pw,py+ph],46,fill=(243,239,227))
# top purple band
d.rounded_rectangle([px,py,px+pw,py+230],46,fill=(120,40,140))
d.rectangle([px,py+150,px+pw,py+230],fill=(120,40,140))
# notch
d.ellipse([px+pw/2-26,py+24,px+pw/2+26,py+76],fill=(243,239,227))
btext(d,(px+pw/2,py+118),"Tasty Wellness, Better Days",font(22),(245,240,255),anchor="mm")
btext(d,(px+pw/2,py+178),"ALL THE BETTER",font(50),(252,255,200),anchor="mm",bold=2)
# gummies (3 red)
gy=py+360
for i,gx in enumerate([px+120,px+pw/2,px+pw-120]):
    off=[-14,-30,-14][i]
    d.rounded_rectangle([gx-66,gy+off-50,gx+66,gy+off+58],40,fill=(226,58,70))
    d.rounded_rectangle([gx-66,gy+off-50,gx+66,gy+off+10],40,fill=(238,82,92))
    btext(d,(gx,gy+off+6),"BETTER",font(22),(250,180,180),anchor="mm")
btext(d,(px+pw/2,py+560),"식물성 멜라토닌 함유 · 멜라나잇 구미",font(28),(60,45,55),anchor="mm",bold=1)
btext(d,(px+pw/2,py+610),"Melatonin Gummies · Tart Cherry Flavor",font(24),(120,100,110),anchor="mm")
btext(d,(px+pw/2,py+660),"7일분 42g (3g x 14구미)",font(22),(150,135,145),anchor="mm")
# flavor tag
d.rounded_rectangle([px+pw-150,py+250,px+pw-30,py+312],20,fill=(150,135,120))
btext(d,(px+pw-90,py+281),"타트체리",font(24),(255,255,255),anchor="mm",bold=1)

# ---- bottom caption + hashtags ----
d.rounded_rectangle([54,H-360,W-54,H-180],34,fill=(255,255,255))
btext(d,(96,H-330),"“오늘도 잘 채웠다, 내일도 better 하게.”",font(40),(40,30,45),bold=1)
btext(d,(96,H-262),"#오늘도배러  #자기전루틴  #멜라나잇구미  #타트체리",font(34),(120,60,140),bold=1)
# reels play hint
d.ellipse([W-180,H-336,W-92,H-248],fill=(120,40,140))
d.polygon([(W-150,H-312),(W-150,H-272),(W-116,H-292)],fill=(255,255,255))

# ---- footer label ----
btext(d,(60,H-152),"마이크로 인플루언서 시딩 · 콘셉트 썸네일 예시",font(28),(200,190,205),bold=1)
btext(d,(60,H-110),"일반식품 — 효능 표현 배제 / TPO·루틴 소구 · 이미지 예시",font(24),(160,150,168))
btext(d,(W-60,H-130),"ALL THE BETTER",font(28),(252,255,161),anchor="ra",bold=1)

bg.save("/tmp/melanight_thumb.png")
print("saved /tmp/melanight_thumb.png", bg.size)
