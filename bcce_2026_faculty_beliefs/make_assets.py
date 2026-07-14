#!/usr/bin/env python3
"""Build redesigned figure assets in the restrained B&W + flamingo palette."""
import os
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
PINK = (0xF5, 0x62, 0x8C)
INK  = (0x17, 0x14, 0x11)
GRAY = (0x6B, 0x66, 0x60)
FAINT= (0xB2, 0xAC, 0xA3)
SERIF   = "/usr/share/fonts/truetype/liberation/LiberationSerif-Regular.ttf"
SERIFB  = "/usr/share/fonts/truetype/liberation/LiberationSerif-Bold.ttf"
SERIFI  = "/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf"
SANS    = "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
SANSB   = "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
def F(p,s): return ImageFont.truetype(p,s)

# ---------- 1. Desaturated speech-bubble opener ----------
src = Image.open(os.path.join(HERE, "image1.png")).convert("RGBA")
_w,_h = src.size
src = src.crop((4,4,_w-4,_h-4))            # strip the 1px edge frame
gray = Image.new("RGBA", src.size, (0,0,0,0))
px = src.load(); gp = gray.load()
w,h = src.size
for y in range(h):
    for x in range(w):
        r,g,b,a = px[x,y]
        lum = 0.299*r + 0.587*g + 0.114*b
        if a < 12 or lum > 228:        # drop background + near-white rim
            continue
        v = int(120 + (lum-120)*0.5)   # soft warm-gray silhouette
        v = max(70, min(205, v))
        gp[x,y] = (v, v-4, v-8, 255)
gray.save(os.path.join(HERE, "speech_bubbles_gray.png"))
print("saved speech_bubbles_gray.png")

# ---------- 2. Monochrome + flamingo boundary map ----------
W, H = 1560, 1740
img = Image.new("RGBA", (W,H), (0,0,0,0))
d = ImageDraw.Draw(img)

def arrow(draw, x0, x1, y, color, wgt=3, head=11):
    draw.line([(x0,y),(x1,y)], fill=color, width=wgt)
    for xx, dirn in ((x0,1),(x1,-1)):
        draw.polygon([(xx,y),(xx+dirn*head, y-head*0.7),(xx+dirn*head, y+head*0.7)], fill=color)

themes = [
    ("Theme 1", "The doctorate develops a self-directed scientist in service of the student’s goals",
     [("IC","answering existing questions","asking new questions"),
      ("UN","weight on independent science","weight on career preparation"),
      ("SG","career preparation as the defining purpose","career preparation as a minor benefit"),
      ("FE","independence largely reached by the PhD","the PhD as a foundation only")]),
    ("Theme 2", "Students learn by doing scaffolded research, shown in performance",
     [("LE","unsupported struggle is productive","struggle should be scaffolded"),
      ("LO","publication is a reliable indicator","publication is an unreliable indicator"),
      ("AV","active checking","calibrated trust")]),
    ("Theme 3", "Development is a shared responsibility within a two-way relationship",
     [("MB","hands-on mentoring","hands-off mentoring"),
      ("FA","faculty-guided development","self-directed development"),
      ("MP","motivation is innate","motivation is developed"),
      ("OE","affective growth is a high priority","affective growth is less essential"),
      ("RS","responsibility stays shared","responsibility fully transfers to the student"),
      ("SD","the capacity is built through initiative","the capacity is partly given")]),
    ("Theme 4", "Structure shapes advising and leaves gaps the program owns",
     [("PS","the evolved system mostly works","curricula and funding need real change"),
      ("GA","professional and career gaps","teaching as a gap")]),
]

f_intro = F(SERIFI, 30)
f_hd    = F(SERIFB, 30)
f_hdn   = F(SANSB, 24)
f_pole  = F(SERIF, 27)
f_tag   = F(SANSB, 21)

MX = 40
AX0, AX1 = 470, W-MX          # axis column
d.text((MX, 20), "Within each shared theme, faculty differ along the axes below.",
       font=f_intro, fill=GRAY)
d.text((MX, 58), "Each axis names the two poles of variation; nothing is plotted on it.",
       font=f_intro, fill=FAINT)

y = 130
for (tn, prop, rows) in themes:
    # header bar
    barh = 58
    d.rectangle([MX, y, W-MX, y+barh], fill=INK)
    d.text((MX+18, y+10), tn+".", font=f_hdn, fill=PINK)
    # wrap prop if long
    tw = d.textbbox((0,0), prop, font=f_hd)[2]
    d.text((MX+18+d.textbbox((0,0),tn+".  ",font=f_hdn)[2], y+13), prop, font=f_hd, fill=(255,255,255))
    y += barh + 26
    for (tag, lp, rp) in rows:
        ay = y + 6
        arrow(d, AX0, AX1, ay, PINK, wgt=3, head=11)
        # code tag at far left
        d.text((MX+6, ay-16), tag, font=f_tag, fill=GRAY)
        # poles under the arrow
        d.text((AX0, ay+12), lp, font=f_pole, fill=INK)
        rpw = d.textbbox((0,0), rp, font=f_pole)[2]
        d.text((AX1-rpw, ay+12), rp, font=f_pole, fill=INK)
        y += 78
    y += 22

img = img.crop((0,0,W, min(H, y+10)))
img.save(os.path.join(HERE, "boundary_map_v2.png"))
print("saved boundary_map_v2.png", img.size)
