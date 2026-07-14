#!/usr/bin/env python3
"""BCCE 2026 deck v3 — matched to the lab house style (assertion-evidence,
sage-green boxes, Century Gothic, terrazzo background)."""
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
A = os.path.join(HERE, "assets")
BGTEX = os.path.join(A, "bg_texture.png")
MAP = os.path.join(A, "boundary_map_v3.png")

GREEN = RGBColor(0x7D,0x92,0x63)
GREEND= RGBColor(0x64,0x77,0x4E)
DARK  = RGBColor(0x2C,0x2C,0x2C)
GRAY  = RGBColor(0x59,0x59,0x59)
FAINT = RGBColor(0x8C,0x89,0x83)
WHITE = RGBColor(0xFF,0xFF,0xFF)
CREAM = RGBColor(0xF7,0xEE,0xCE)
PAPER = RGBColor(0xF7,0xF7,0xF7)
CG = "Century Gothic"
TNR = "Times New Roman"

prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
SW, SH = 13.333, 7.5
BLANK = prs.slide_layouts[6]
LM = 0.75

def slide():
    s = prs.slides.add_slide(BLANK)
    r = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0,0, prs.slide_width, prs.slide_height)
    r.fill.solid(); r.fill.fore_color.rgb = PAPER; r.line.fill.background(); r.shadow.inherit=False
    sp=r._element; sp.getparent().remove(sp); s.shapes._spTree.insert(2, sp)
    try:
        p = s.shapes.add_picture(BGTEX, 0,0, width=prs.slide_width, height=prs.slide_height)
        pe=p._element; pe.getparent().remove(pe); s.shapes._spTree.insert(3, pe)
    except Exception: pass
    return s

def txt(slide,l,t,w,h,runs,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP,line_spacing=1.0,space_after=0):
    tb=slide.shapes.add_textbox(Inches(l),Inches(t),Inches(w),Inches(h))
    tf=tb.text_frame; tf.word_wrap=True; tf.vertical_anchor=anchor
    for m in ('left','right','top','bottom'): setattr(tf,'margin_'+m,0)
    for i,para in enumerate(runs):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.alignment=align
        if line_spacing: p.line_spacing=line_spacing
        p.space_before=0; p.space_after=Pt(space_after)
        for (text,st) in para:
            r=p.add_run(); r.text=text; f=r.font
            f.name=st.get('font',CG); f.size=Pt(st.get('size',18))
            f.bold=st.get('bold',False); f.italic=st.get('italic',False)
            f.color.rgb=st.get('color',DARK)
    return tb

def title(slide, lines, size=27, y=0.5, color=DARK):
    txt(slide, LM, y, 12.0, 1.4, [[(l, dict(font=CG, size=size, color=color))] for l in lines], line_spacing=1.08)

def citation(slide, text):
    txt(slide, LM, 7.02, 9, 0.3, [[(text, dict(font=CG, size=10, color=FAINT))]])

def pagenum(slide, n):
    txt(slide, SW-LM-0.6, 7.02, 0.6, 0.3, [[(str(n), dict(font=CG, size=10, color=FAINT))]], align=PP_ALIGN.RIGHT)

def box(slide,l,t,w,h,lines,fill=GREEN,tcolor=WHITE,size=15,bold=False,radius=0.14,
        sharp=False,line=None,line_w=1.25,italic=False,anchor=MSO_ANCHOR.MIDDLE,font=CG):
    shp=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE if sharp else MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(l),Inches(t),Inches(w),Inches(h))
    if fill is None: shp.fill.background()
    else: shp.fill.solid(); shp.fill.fore_color.rgb=fill
    if line is None: shp.line.fill.background()
    else: shp.line.color.rgb=line; shp.line.width=Pt(line_w)
    shp.shadow.inherit=False
    if not sharp:
        try: shp.adjustments[0]=radius
        except Exception: pass
    if isinstance(lines,str): lines=[lines]
    txt(slide,l+0.08,t,w-0.16,h,[[(ln,dict(font=font,size=size,color=tcolor,bold=bold,italic=italic))] for ln in lines],
        align=PP_ALIGN.CENTER,anchor=anchor,line_spacing=1.02)
    return shp

def oval(slide,l,t,w,h,text,fill=GREEN,tcolor=WHITE,size=14,bold=False):
    shp=slide.shapes.add_shape(MSO_SHAPE.OVAL,Inches(l),Inches(t),Inches(w),Inches(h))
    shp.fill.solid(); shp.fill.fore_color.rgb=fill; shp.line.fill.background(); shp.shadow.inherit=False
    txt(slide,l,t,w,h,[[(text,dict(font=CG,size=size,color=tcolor,bold=bold))]],align=PP_ALIGN.CENTER,anchor=MSO_ANCHOR.MIDDLE)
    return shp

def arrow(slide,x1,y1,x2,y2,color=GRAY,w=1.75,double=False):
    c=slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2))
    c.line.color.rgb=color; c.line.width=Pt(w); c.shadow.inherit=False
    ln=c.line._get_or_add_ln()
    tags=('a:headEnd','a:tailEnd') if double else ('a:tailEnd',)
    for tg in tags: ln.append(ln.makeelement(qn(tg),{'type':'triangle','w':'med','len':'med'}))
    return c

# ============================================= S1 TITLE
s=slide()
txt(s,1.0,1.05,11.33,0.4,[[("BCCE 2026", dict(font=CG,size=13,color=GREEND))]],align=PP_ALIGN.CENTER)
txt(s,1.0,2.35,11.33,2.0,
    [[("What chemistry faculty believe", dict(font=CG,size=38,color=DARK))],
     [("about doctoral education", dict(font=CG,size=38,color=DARK))]],align=PP_ALIGN.CENTER,line_spacing=1.1)
# green underline accent
u=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(5.67),Inches(4.15),Inches(2.0),Inches(0.05))
u.fill.solid(); u.fill.fore_color.rgb=GREEN; u.line.fill.background(); u.shadow.inherit=False
txt(s,1.0,4.4,11.33,0.5,[[("A phenomenographic and thematic characterization",dict(font=CG,size=19,italic=True,color=GRAY))]],align=PP_ALIGN.CENTER)
txt(s,1.0,5.5,11.33,0.4,[[("Zoe Rahimi Pirkoohi",dict(font=CG,size=18,color=DARK))]],align=PP_ALIGN.CENTER)
txt(s,1.0,5.98,11.33,0.4,[[("Jordan Harshman  ·  University of Iowa",dict(font=CG,size=15,color=GRAY))]],align=PP_ALIGN.CENTER)
txt(s,1.0,6.95,11.33,0.35,[[("Supported by NSF CAREER #2142873 & #2602955  ·  Jordan Harshman, PI",dict(font=CG,size=10,color=FAINT))]],align=PP_ALIGN.CENTER)

# ============================================= S2 ENTERPRISE
s=slide()
title(s,["Doctoral education produces the chemists","who staff the discipline."])
data=[("~2,900","chemistry PhDs each year"),("~196","U.S. doctoral programs"),("decades","one largely unchanged structure")]
bw=3.7; gap=0.3; x=LM+ (12.0-(bw*3+gap*2))/2 -0.0
x=LM
for (big,small) in data:
    box(s,x,2.7,bw,1.5,[big,small],size=15,fill=GREEN, anchor=MSO_ANCHOR.MIDDLE)
    # big number larger: overlay
    x+=bw+gap
# redo with two-line internal styling
# (numbers as bold larger handled by separate overlay)
txt(s,LM,4.7,12.0,1.0,
    [[("Yet the same record documents its strains — competing responsibilities, no standard "
       "assessment of outcomes, uneven implementation, and inconsistent mentorship.",dict(font=CG,size=16,color=DARK))]],line_spacing=1.3)
citation(s,"Donkor & Harshman, 2023  ·  Collini et al., 2025"); pagenum(s,2)

# ============================================= S3 THE GAP
s=slide()
title(s,["Beliefs about advising a PhD","have never been characterized."])
box(s,LM,3.0,5.6,1.7,["Classroom teaching","","Beliefs are well characterized —","and predict what happens in class"],fill=GREEN,size=15)
box(s,LM+6.2,3.0,5.6,1.7,["Research advising","","A different practice — and","never characterized"],fill=DARK,sharp=True,size=15)
txt(s,LM,5.15,12.0,0.8,
    [[("Donkor et al. (2024) set faculty beliefs aside — a proper investigation ",dict(font=CG,size=15,color=DARK)),
      ("“warranted its own study.”",dict(font=TNR,size=16,italic=True,color=GREEND)),
      ("  This is that study.",dict(font=CG,size=15,color=DARK,bold=True))]],line_spacing=1.25)
citation(s,"Gibbons; Lund & Stains; Popova; Mack & Towns  ·  Kagan, 1992"); pagenum(s,3)

# ============================================= S4 RESEARCH QUESTION
s=slide()
txt(s,1.0,1.5,11.33,0.4,[[("RESEARCH QUESTION",dict(font=CG,size=13,color=GREEND))]],align=PP_ALIGN.CENTER)
txt(s,1.0,2.5,11.33,1.6,
    [[("What do chemistry faculty believe",dict(font=CG,size=32,color=DARK))],
     [("about doctoral education?",dict(font=CG,size=32,color=DARK))]],align=PP_ALIGN.CENTER,line_spacing=1.12)
areas=["what it’s for","how students learn","who develops which skill","the structure they work in"]
bw=2.7; gap=0.28; total=bw*4+gap*3; x=(SW-total)/2
for a in areas:
    box(s,x,4.9,bw,0.85,[a],size=13.5,fill=GREEN); x+=bw+gap
pagenum(s,4)

# ============================================= S5 TCSR
s=slide()
title(s,["Whether a reform reaches practice","depends on what teachers believe."])
els=[("general context","of reform",GREEN),("the teacher’s","personal profile",GREEN),
     ("structural &","cultural context",GREEN),("teacher thinking","→ beliefs",DARK)]
bw=2.72; gap=0.2; x=LM
for (a,b,fill) in els:
    box(s,x,2.85,bw,1.2,[a,b],fill=fill,sharp=(fill==DARK),size=14); x+=bw+gap
txt(s,LM,4.35,12.0,0.4,[[("All four shape whether a reform reaches practice — the model’s ultimate test.",dict(font=CG,size=14,color=GRAY,italic=True))]],align=PP_ALIGN.CENTER)
box(s,2.9,5.15,7.5,1.0,["Beliefs are self-perpetuating — they can","absorb a reform without yielding to it."],fill=CREAM,tcolor=DARK,size=17,bold=False)
citation(s,"Woodbury & Gess-Newsome, 2002"); pagenum(s,5)

# ============================================= S6 BELIEF RULE
s=slide()
title(s,["A belief is a proposition held true","that guides how a person acts."])
gates=[("Propositional","held to be true"),("Evaluative","guides decisions"),
       ("Personal","not group consensus"),("Resistant","survives counter-evidence")]
bw=2.5; gap=0.32; x=LM; ytop=3.0
for i,(term,gloss) in enumerate(gates):
    box(s,x,ytop,bw,1.15,[term,"",gloss],fill=GREEN,size=14.5,bold=False)
    if i<3: arrow(s,x+bw+0.02,ytop+0.575,x+bw+gap-0.02,ytop+0.575,color=GRAY,w=1.5)
    x+=bw+gap
box(s,5.4,4.6,2.5,0.7,["= a belief"],fill=DARK,sharp=True,size=16,bold=True)
citation(s,"Rokeach, 1968 · Green, 1971 · Nespor, 1987 · Pajares, 1992"); pagenum(s,6)

# ============================================= S7 METHODS
s=slide()
title(s,["An interpretive design built to capture","the range of faculty belief."])
m=[("Twenty faculty","U.S. doctoral programs"),("Interviews + card sort","~60-min, semi-structured"),
   ("Phenomenography +","thematic analysis")]
bw=3.7; gap=0.3; x=LM
for (a,b) in m:
    box(s,x,2.8,bw,1.15,[a,"",b],fill=GREEN,size=14.5); x+=bw+gap
pipe=[("990","sub-codes"),("15","codes"),("4","themes")]
x=2.15; ytop=4.65
for i,(big,small) in enumerate(pipe):
    box(s,x,ytop,2.0,0.95,[big+"  "+small],fill=None,tcolor=DARK,size=17,line=GREEN,line_w=1.5)
    if i<2: arrow(s,x+2.02,ytop+0.475,x+2.5-0.02,ytop+0.475,color=GREEN,w=1.75)
    x+=2.5
txt(s,LM,5.85,12.0,0.4,[[("No new codes appeared after interview 7 — the result held from 10 participants to all 20.",dict(font=CG,size=13,color=GRAY,italic=True))]],align=PP_ALIGN.CENTER)
citation(s,"Braun & Clarke, 2006 · Åkerlind, 2005"); pagenum(s,7)

# ============================================= S8 SAMPLE
s=slide()
title(s,["Recruited for maximum variation","across profile and subfield."])
stats=[("N = 20","participants"),("13 / 7","men / women"),("13 / 7","White / non-White"),
       ("15","full professors"),("14","at R1 institutions"),("18","U.S.-trained")]
bw=3.7; bh=1.0; gap=0.3; gy=0.35
for i,(big,small) in enumerate(stats):
    col=i%3; row=i//3
    x=LM+col*(bw+gap); y=2.75+row*(bh+gy)
    box(s,x,y,bw,bh,[big,small],fill=GREEN,size=14)
txt(s,LM,5.5,12.0,0.4,[[("Subdisciplines: ",dict(font=CG,size=13,color=DARK,bold=True)),
    ("Inorganic (5) · Chemistry education (2) · Materials (2) · Physical (2) · Biophysical (2) · Organic (2) · + five singletons",dict(font=CG,size=13,color=GRAY))]])
txt(s,LM,6.0,12.0,0.4,[[("Analytical generalization and transferability, not population estimates; the sample skews senior and male.",dict(font=CG,size=11,color=FAINT,italic=True))]])
pagenum(s,8)

# ============================================= S9 THE FINDING
s=slide()
title(s,["Faculty agree on the propositions —","and divide on where the boundaries fall."])
box(s,LM,3.0,5.6,1.6,["AGREE","on the four shared propositions"],fill=GREEN,size=16)
box(s,LM+6.2,3.0,5.6,1.6,["DIVIDE","on where the boundaries fall"],fill=DARK,sharp=True,size=16)
txt(s,LM,5.05,12.0,1.0,
    [[("“What differs is not which beliefs faculty hold, but where they draw the lines those beliefs require.”",dict(font=TNR,size=18,italic=True,color=GREEND))]],align=PP_ALIGN.CENTER,line_spacing=1.2)
pagenum(s,9)

# ============================================= S10-13 THEMES
themes=[
 dict(n=1,label="What the doctorate is for",codes=["IC","UN","SG","FE"],
   prop=["The doctorate develops a self-directed scientist","in service of the student’s own goals."],
   quote="When they start bringing ideas forward without necessarily being prompted.",who="P2 · on what independence looks like",
   left="answering existing questions",right="asking new questions",cap="The realistic goal versus the higher bar for the degree."),
 dict(n=2,label="How students learn",codes=["LE","LO","AV"],
   prop=["Students learn by doing scaffolded research —","and learning shows in performance, not grades."],
   quote="I do. We do. You do — spread out over X number of years.",who="P5 · on how research is learned",
   left="unsupported struggle helps",right="struggle should be scaffolded",cap="Same formative deprivation, opposite lessons — P13 vs. P7."),
 dict(n=3,label="Who develops which skill — the largest theme",codes=["MB","FA","OE","RS","SD","MP"],
   prop=["Development is a shared responsibility in a two-way","relationship — and who develops which skill varies."],
   quote="I go from the leader to the follower, and they start as the follower and end up as the leader.",who="P6 · on how responsibility transfers",
   left="faculty’s to teach",right="student’s to build",cap="The same skill — creativity — is one faculty’s to teach (P5, P17), another’s for the student (P9, P11)."),
 dict(n=4,label="The structural conditions",codes=["PS","GA"],
   prop=["Structure shapes advising — and the gaps it leaves","are the program’s to fix, not the student’s."],
   quote="Those professional skills often get kind of forgotten.",who="P2 · on the program’s gap",
   left="the evolved system works",right="curricula & funding need change",cap="How much of the structure deserves changing — P14 vs. P18."),
]
for i,d in enumerate(themes):
    s=slide()
    txt(s,LM,0.5,12.0,0.3,[[(f"THEME {d['n']} / 4    ·    {d['label'].upper()}",dict(font=CG,size=10.5,color=GREEND))]])
    title(s,d["prop"],size=25,y=0.95)
    # quote
    txt(s,LM,2.75,11.8,1.1,[[(d["quote"],dict(font=TNR,size=21,italic=True,color=DARK))]],line_spacing=1.16)
    txt(s,LM,3.95,11.8,0.35,[[("— "+d["who"],dict(font=CG,size=13,color=GRAY))]])
    # codes as green boxes
    cx=LM
    for c in d["codes"]:
        box(s,cx,4.5,0.72,0.42,[c],fill=GREEN,size=12.5,bold=True,radius=0.2); cx+=0.84
    # boundary
    txt(s,LM,5.25,8,0.3,[[("WHERE THEY DIVIDE",dict(font=CG,size=10.5,color=GREEND))]])
    arrow(s,2.2,5.85,10.9,5.85,color=GRAY,w=1.75,double=True)
    txt(s,2.1,5.95,4.6,0.4,[[(d["left"],dict(font=CG,size=13.5,color=DARK))]])
    txt(s,6.4,5.95,4.6,0.4,[[(d["right"],dict(font=CG,size=13.5,color=DARK))]],align=PP_ALIGN.RIGHT)
    txt(s,LM,6.4,12.0,0.4,[[(d["cap"],dict(font=CG,size=12,italic=True,color=GRAY))]])
    pagenum(s,10+i)

# ============================================= S14 MAP
s=slide()
s.shapes.add_picture(MAP,Inches(0.62),Inches(0.85),height=Inches(6.2))
RX=7.15
title([s][0],["The full outcome space"],size=24,y=1.1) if False else None
txt(s,RX,1.1,5.5,0.4,[[("THE FULL OUTCOME SPACE",dict(font=CG,size=11,color=GREEND))]])
txt(s,RX,1.65,5.6,2.6,
    [[("Four shared themes.",dict(font=CG,size=24,color=DARK))],
     [("Fifteen codes.",dict(font=CG,size=24,color=DARK))],
     [("An axis of division",dict(font=CG,size=24,color=GREEND))],
     [("in every one.",dict(font=CG,size=24,color=DARK))]],line_spacing=1.16)
txt(s,RX,4.7,5.6,1.8,[[("The agreement is real — and so is the disagreement one level down. Every faculty member lands somewhere on each axis.",dict(font=CG,size=14,color=GRAY))]],line_spacing=1.3)
pagenum(s,14)

# ============================================= S15 WHAT THIS ADDS
s=slide()
title(s,["A belief baseline","for changing doctoral education."])
pts=[("Content for the model","Gives the Teacher-Centered Systemic Reform model the doctoral-education content it lacked — the shared ground and the fault lines a reform must engage."),
     ("Nests with prior work","The goals record and the challenge record now sit beside what faculty believe. Three answers to one enterprise — they nest, not compete."),
     ("Conviction vs. practice","Thirteen of twenty hold stable convictions but refine implementation with experience — so supports aimed at implementation may fit faculty best.")]
bw=3.7; gap=0.3; x=LM
for (h,g) in pts:
    box(s,x,2.75,bw,0.7,[h],fill=GREEN,size=15,bold=True)
    txt(s,x+0.05,3.6,bw-0.1,2.4,[[(g,dict(font=CG,size=13,color=DARK))]],line_spacing=1.25)
    x+=bw+gap
citation(s,"Woodbury & Gess-Newsome, 2002 · Donkor et al.; Collini et al."); pagenum(s,15)

# ============================================= S16 IMPLICATIONS + THANKS
s=slide()
title(s,["Two levels, two lessons."])
box(s,LM,2.4,5.6,1.7,["SHARED PROPOSITIONS","","Common ground a reform can take as","given — arguing from where faculty stand."],fill=GREEN,size=14.5)
box(s,LM+6.2,2.4,5.6,1.7,["BOUNDARY DISAGREEMENTS","","The harder, more useful target — where","aligned colleagues quietly diverge."],fill=DARK,sharp=True,size=14.5)
u=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(LM),Inches(4.75),Inches(1.6),Inches(0.05))
u.fill.solid(); u.fill.fore_color.rgb=GREEN; u.line.fill.background(); u.shadow.inherit=False
txt(s,LM,4.95,11.8,0.7,[[("Thank you.",dict(font=CG,size=28,color=DARK))]])
txt(s,LM,5.95,11.8,0.35,[[("Zoe Rahimi Pirkoohi  ·  Jordan Harshman  ·  University of Iowa",dict(font=CG,size=13,color=GRAY))]])
txt(s,LM,6.4,11.8,0.3,[[("Supported by NSF CAREER #2142873 & #2602955  ·  Jordan Harshman, PI",dict(font=CG,size=10,color=FAINT))]])

out=os.path.join(HERE,"BCCE2026_faculty_beliefs.pptx")
prs.save(out)
print("saved",out,len(prs.slides._sldIdLst),"slides")
