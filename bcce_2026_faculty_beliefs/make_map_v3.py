#!/usr/bin/env python3
"""Boundary map in the reference house style: green headers, gray axes."""
import os
from PIL import Image, ImageDraw, ImageFont
HERE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
GREEN=(0x7D,0x92,0x63); DARK=(0x2C,0x2C,0x2C); GRAY=(0x59,0x59,0x59)
FAINT=(0x9A,0x94,0x8B); WHITE=(255,255,255)
SERIF="/usr/share/fonts/truetype/liberation/LiberationSerif-Italic.ttf"
SANS="/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"
SANSB="/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf"
def F(p,s): return ImageFont.truetype(p,s)

themes=[("Theme 1","The doctorate develops a self-directed scientist in service of the student’s goals",
  [("IC","answering existing questions","asking new questions"),
   ("UN","weight on independent science","weight on career preparation"),
   ("SG","career preparation as the defining purpose","career preparation as a minor benefit"),
   ("FE","independence largely reached by the PhD","the PhD as a foundation only")]),
 ("Theme 2","Students learn by doing scaffolded research, shown in performance",
  [("LE","unsupported struggle is productive","struggle should be scaffolded"),
   ("LO","publication is a reliable indicator","publication is an unreliable indicator"),
   ("AV","active checking","calibrated trust")]),
 ("Theme 3","Development is a shared responsibility within a two-way relationship",
  [("MB","hands-on mentoring","hands-off mentoring"),
   ("FA","faculty-guided development","self-directed development"),
   ("MP","motivation is innate","motivation is developed"),
   ("OE","affective growth is a high priority","affective growth is less essential"),
   ("RS","responsibility stays shared","responsibility fully transfers to the student"),
   ("SD","the capacity is built through initiative","the capacity is partly given")]),
 ("Theme 4","Structure shapes advising and leaves gaps the program owns",
  [("PS","the evolved system mostly works","curricula and funding need real change"),
   ("GA","professional and career gaps","teaching as a gap")]),
]
W,H=1560,1740
img=Image.new("RGBA",(W,H),(0,0,0,0)); d=ImageDraw.Draw(img)
def arrow(x0,x1,y,color,wgt=3,head=11):
    d.line([(x0,y),(x1,y)],fill=color,width=wgt)
    for xx,dr in ((x0,1),(x1,-1)):
        d.polygon([(xx,y),(xx+dr*head,y-head*0.7),(xx+dr*head,y+head*0.7)],fill=color)
f_intro=F(SERIF,30); f_hd=F(SANSB,28); f_hdn=F(SANSB,24); f_pole=F(SANS,27); f_tag=F(SANSB,21)
MX=40; AX0,AX1=470,W-MX
d.text((MX,20),"Within each shared theme, faculty differ along the axes below.",font=f_intro,fill=GRAY)
d.text((MX,58),"Each axis names the two poles of variation; nothing is plotted on it.",font=f_intro,fill=FAINT)
y=130
for (tn,prop,rows) in themes:
    barh=60
    d.rounded_rectangle([MX,y,W-MX,y+barh],radius=8,fill=GREEN)
    d.text((MX+20,y+13),tn+".",font=f_hdn,fill=(0xEC,0xF0,0xE4))
    d.text((MX+20+d.textbbox((0,0),tn+".  ",font=f_hdn)[2],y+15),prop,font=f_hd,fill=WHITE)
    y+=barh+26
    for (tag,lp,rp) in rows:
        ay=y+6
        arrow(AX0,AX1,ay,GRAY,wgt=3,head=11)
        d.text((MX+6,ay-16),tag,font=f_tag,fill=GREEN)
        d.text((AX0,ay+12),lp,font=f_pole,fill=DARK)
        rpw=d.textbbox((0,0),rp,font=f_pole)[2]
        d.text((AX1-rpw,ay+12),rp,font=f_pole,fill=DARK)
        y+=78
    y+=22
img=img.crop((0,0,W,min(H,y+10)))
img.save(os.path.join(HERE,"boundary_map_v3.png"))
print("saved boundary_map_v3.png",img.size)
