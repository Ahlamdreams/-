import cairo, numpy as np, math, sys, os, subprocess
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from multiprocessing import Pool
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import timeline as TL
ROOT=os.path.dirname(os.path.abspath(__file__))+'/..'
W,H,FPS=1080,1920,30
OFF=TL.OFF; VEND=OFF+TL.VOICE_LEN; OUTRO=VEND+3.9; DUR=OUTRO+4.6; NF=int(DUR*FPS)
P=None
NAVY1=(0.035,0.075,0.19); NAVY2=(0.075,0.16,0.36); CARD=(0.11,0.21,0.44)
GOLD=(0.96,0.75,0.24); GOLD2=(1.0,0.85,0.45); WHITE=(1,1,1); SOFT=(0.62,0.70,0.86)
HANDLE="@ahlamtrainer"
VC=(540,650)   # visual centre
# ---------------- easing ----------------
def cl(x): return 0.0 if x<0 else 1.0 if x>1 else x
def seg(t,a,b): return cl((t-a)/(b-a)) if b>a else float(t>=a)
def eo(x): x=cl(x); return 1-(1-x)**3
def eio(x): x=cl(x); return x*x*(3-2*x)
def back(x,c=1.7):
    x=cl(x); return 1+(c+1)*(x-1)**3+c*(x-1)**2
def elastic(x):
    x=cl(x)
    if x in (0,1): return x
    return 2**(-10*x)*math.sin((x*10-0.75)*2*math.pi/3)+1
def lerp(a,b,k): return a+(b-a)*k
def mix(c1,c2,k): return tuple(lerp(a,b,k) for a,b in zip(c1,c2))
# ---------------- drawing helpers ----------------
def rgba(c,a,x): x.set_source_rgba(c[0],c[1],c[2],a)
def circle(x,cx,cy,r,c,a=1,fill=True,lw=4):
    if r<=0 or a<=0: return
    x.new_path(); x.arc(cx,cy,r,0,2*math.pi); rgba(c,a,x)
    if fill: x.fill()
    else: x.set_line_width(lw); x.stroke()
def rrect(x,x0,y0,w,h,r,c,a=1,fill=True,lw=4):
    if a<=0 or w<=0 or h<=0: return
    r=min(r,w/2,h/2); x.new_path()
    x.arc(x0+w-r,y0+r,r,-math.pi/2,0); x.arc(x0+w-r,y0+h-r,r,0,math.pi/2); x.arc(x0+r,y0+h-r,r,math.pi/2,math.pi); x.arc(x0+r,y0+r,r,math.pi,1.5*math.pi); x.close_path()
    rgba(c,a,x)
    if fill: x.fill()
    else: x.set_line_width(lw); x.stroke()
def poly(x,pts,c,a=1,fill=True,lw=4,close=True):
    if a<=0 or len(pts)<2: return
    x.new_path(); x.move_to(*pts[0])
    for p in pts[1:]: x.line_to(*p)
    if close: x.close_path()
    rgba(c,a,x)
    if fill: x.fill()
    else: x.set_line_width(lw); x.set_line_join(cairo.LINE_JOIN_ROUND); x.set_line_cap(cairo.LINE_CAP_ROUND); x.stroke()
def partial(pts,k):
    """first fraction k (by length) of polyline"""
    if k<=0: return pts[:1]
    pts=np.asarray(pts,float); d=np.r_[0,np.cumsum(np.hypot(*np.diff(pts,axis=0).T))]; L=d[-1]*cl(k)
    i=np.searchsorted(d,L); 
    if i>=len(pts): return pts.tolist()
    f=(L-d[i-1])/max(d[i]-d[i-1],1e-6); return pts[:i].tolist()+[list(pts[i-1]+(pts[i]-pts[i-1])*f)]
def glowline(x,pts,c,a,lw,k=1.0,close=False):
    pp=partial(pts,k) if k<1 else pts
    for wmul,am in ((4.2,0.07),(2.4,0.14),(1,1)):
        poly(x,pp,c,a*am,fill=False,lw=lw*wmul,close=close and k>=1)
def arcp(cx,cy,r,a0,a1,n=64): return [(cx+r*math.cos(a0+(a1-a0)*i/n),cy+r*math.sin(a0+(a1-a0)*i/n)) for i in range(n+1)]
def radial(x,cx,cy,r,c,a):
    g=cairo.RadialGradient(cx,cy,0,cx,cy,r); g.add_color_stop_rgba(0,c[0],c[1],c[2],a); g.add_color_stop_rgba(1,c[0],c[1],c[2],0)
    x.set_source(g); x.new_path(); x.arc(cx,cy,r,0,2*math.pi); x.fill()
def star_pts(cx,cy,r,ri=None,n=5,rot=-math.pi/2):
    ri=ri or r*0.45; return [(cx+(r if i%2==0 else ri)*math.cos(rot+i*math.pi/n),cy+(r if i%2==0 else ri)*math.sin(rot+i*math.pi/n)) for i in range(2*n)]
def shape_pts(kind,r,n=96):
    out=[]
    for i in range(n):
        u=i/n; a=-math.pi/2+u*2*math.pi
        if kind=='circle': out.append((r*math.cos(a),r*math.sin(a)))
        else:
            verts={'square':[(r,-r),(r,r),(-r,r),(-r,-r)],'triangle':[(0,-r*1.15),(r*1.05,r*0.75),(-r*1.05,r*0.75)],
                   'arrow':[(0,-r*1.3),(r*0.95,-r*0.15),(r*0.38,-r*0.15),(r*0.38,r*1.1),(-r*0.38,r*1.1),(-r*0.38,-r*0.15),(-r*0.95,-r*0.15)]}[kind]
            if kind=='square': verts=[(0,-r)]+[(r,-r),(r,r),(-r,r),(-r,-r)]+[(0,-r)]; verts=verts[:-1]
            vv=verts+[verts[0]]; seglen=[math.dist(vv[j],vv[j+1]) for j in range(len(verts))]; L=sum(seglen); s=u*L
            for j,sl in enumerate(seglen):
                if s<=sl: f=s/sl; out.append((vv[j][0]+(vv[j+1][0]-vv[j][0])*f,vv[j][1]+(vv[j+1][1]-vv[j][1])*f)); break
                s-=sl
            else: out.append(vv[0])
    return np.array(out)
def person(x,cx,cy,s,c,a=1):
    circle(x,cx,cy-62*s,22*s,c,a)
    x.new_path(); x.move_to(cx-38*s,cy+20*s); x.curve_to(cx-38*s,cy-38*s,cx+38*s,cy-38*s,cx+38*s,cy+20*s); x.close_path(); rgba(c,a,x); x.fill()
def lock(x,cx,cy,s,a,open_k=0):
    rrect(x,cx-34*s,cy-22*s,68*s,56*s,10*s,GOLD,a)
    x.new_path(); x.arc(cx,cy-22*s-open_k*26*s,22*s,math.pi,2*math.pi); x.set_line_width(9*s); rgba(WHITE,a,x); x.stroke()
    x.new_path(); x.move_to(cx-22*s,cy-22*s-open_k*26*s); x.line_to(cx-22*s,cy-18*s-open_k*26*s); x.stroke()
    x.new_path(); x.move_to(cx+22*s,cy-22*s-open_k*26*s); x.line_to(cx+22*s,cy-14*s); x.stroke()
    circle(x,cx,cy+4*s,7*s,NAVY1,a)
def key(x,cx,cy,s,rot,a):
    x.save(); x.translate(cx,cy); x.rotate(rot); x.scale(s,s)
    circle(x,-50,0,26,GOLD,a,fill=False,lw=12); rrect(x,-26,-7,110,14,6,GOLD,a); rrect(x,60,0,12,26,3,GOLD,a); rrect(x,80,0,12,20,3,GOLD,a)
    x.restore()
def gear(x,cx,cy,r,rot,c,a,teeth=10):
    pts=[]
    for i in range(teeth*4):
        ang=rot+i*2*math.pi/(teeth*4); rr=r if (i%4) in (0,1) else r*0.78; pts.append((cx+rr*math.cos(ang),cy+rr*math.sin(ang)))
    poly(x,pts,c,a); circle(x,cx,cy,r*0.32,NAVY1,a)
def heart(x,cx,cy,s,c,a):
    x.new_path(); x.move_to(cx,cy+40*s)
    x.curve_to(cx-70*s,cy-5*s,cx-40*s,cy-60*s,cx,cy-25*s); x.curve_to(cx+40*s,cy-60*s,cx+70*s,cy-5*s,cx,cy+40*s); rgba(c,a,x); x.fill()
def arrow_shape(x,x0,y0,x1,y1,c,a,lw=8,head=26):
    poly(x,[(x0,y0),(x1,y1)],c,a,fill=False,lw=lw,close=False)
    ang=math.atan2(y1-y0,x1-x0)
    poly(x,[(x1+math.cos(ang)*head*0.6,y1+math.sin(ang)*head*0.6),(x1+math.cos(ang+2.5)*head,y1+math.sin(ang+2.5)*head),(x1+math.cos(ang-2.5)*head,y1+math.sin(ang-2.5)*head)],c,a)
def clock(x,cx,cy,r,t,a):
    circle(x,cx,cy,r,WHITE,a,fill=False,lw=8)
    for i in range(12):
        ang=i*math.pi/6; poly(x,[(cx+r*0.78*math.cos(ang),cy+r*0.78*math.sin(ang)),(cx+r*0.88*math.cos(ang),cy+r*0.88*math.sin(ang))],WHITE,a*0.7,fill=False,lw=4,close=False)
    for L,sp,lw in ((0.55,1.2,8),(0.75,7,5)):
        ang=t*sp-math.pi/2; poly(x,[(cx,cy),(cx+r*L*math.cos(ang),cy+r*L*math.sin(ang))],GOLD,a,fill=False,lw=lw,close=False)
def ban(x,cx,cy,r,k,a,lw=14):
    if k<=0: return
    glowline(x,arcp(cx,cy,r,-math.pi*0.75,-math.pi*0.75+2*math.pi*cl(k*1.4)),GOLD,a,lw)
    if k>0.6:
        kk=(k-0.6)/0.4; d=r*0.707
        glowline(x,partial([(cx-d,cy-d),(cx+d,cy+d)],kk),GOLD,a,lw)
# ---------------- text sprites (PIL + raqm -> cairo) ----------------
SPR={}
KEEP=[]
FONTS={}
def pfont(size,weight=b'Bold',kind='cairo'):
    k=(size,weight,kind)
    if k not in FONTS:
        path={'cairo':'Cairo.ttf','fa':'fa-brands.ttf','mi':'MaterialIcons.ttf'}[kind]
        f=ImageFont.truetype(ROOT+'/fonts/'+path,size,layout_engine=ImageFont.Layout.RAQM)
        if kind=='cairo': f.set_variation_by_name(weight)
        FONTS[k]=f
    return FONTS[k]
def to_surface(im):
    a=np.asarray(im.convert('RGBA'),np.float32); al=a[...,3:4]/255
    b=np.empty(a.shape,np.uint8); b[...,0]=(a[...,2]*al[...,0]); b[...,1]=(a[...,1]*al[...,0]); b[...,2]=(a[...,0]*al[...,0]); b[...,3]=a[...,3]
    h,w=b.shape[:2]; stride=cairo.ImageSurface.format_stride_for_width(cairo.FORMAT_ARGB32,w)
    buf=bytearray(stride*h); arr=np.frombuffer(buf,np.uint8).reshape(h,stride); arr[:,:w*4]=b.reshape(h,w*4)
    s=cairo.ImageSurface.create_for_data(buf,cairo.FORMAT_ARGB32,w,h,stride); KEEP.append(buf); return s
def sprite(txt,size,col=(255,255,255),weight=b'Bold',glow=None,kind='cairo',rtl=True):
    k=(txt,size,col,weight,glow,kind)
    if k in SPR: return SPR[k]
    f=pfont(size,weight,kind); kw=dict(direction='rtl',language='ar') if rtl and kind=='cairo' else {}
    bb=f.getbbox(txt,**kw); pad=34; w=bb[2]-bb[0]+2*pad; h=bb[3]-bb[1]+2*pad
    m=Image.new('L',(w,h),0); ImageDraw.Draw(m).text((pad-bb[0],pad-bb[1]),txt,font=f,fill=255,**kw)
    im=Image.new('RGBA',(w,h),(0,0,0,0))
    if glow:
        g=m.filter(ImageFilter.GaussianBlur(14)); gl=Image.new('RGBA',(w,h),glow+(0,)); gl.putalpha(g.point(lambda v:min(255,int(v*0.75)))); im=Image.alpha_composite(im,gl)
    sh=m.filter(ImageFilter.GaussianBlur(5)); shl=Image.new('RGBA',(w,h),(5,12,35,0)); shl.putalpha(sh.point(lambda v:int(v*0.6)))
    im=Image.alpha_composite(im,shl.transform(shl.size,Image.AFFINE,(1,0,0,0,1,-4)))
    fg=Image.new('RGBA',(w,h),col+(0,)); fg.putalpha(m); im=Image.alpha_composite(im,fg)
    adv=f.getlength(txt,**kw)
    SPR[k]=dict(s=to_surface(im),w=w,h=h,adv=adv,cy=pad+(bb[3]-bb[1])/2)
    return SPR[k]
def blit(x,sp,cx,cy,sc=1,a=1,rot=0):
    if a<=0.003 or sc<=0.001: return
    x.save(); x.translate(cx,cy); x.rotate(rot); x.scale(sc,sc); x.set_source_surface(sp['s'],-sp['w']/2,-sp['h']/2); x.paint_with_alpha(min(1,a)); x.restore()
LOGO=None
def logo_surface(h):
    k=('logo',int(h))
    if k not in SPR:
        im=Image.open(ROOT+'/logo.png').convert('RGBA'); w=int(im.width*h/im.height); SPR[k]=dict(s=to_surface(im.resize((w,int(h)),Image.LANCZOS)),w=w,h=int(h))
    return SPR[k]
# ---------------- scenes (v5: initiatives & impact) ----------------
def ph(i): return P[i]
def wt(i,j): return P[i]['words'][j]['start']
def scene_span(si):
    a=P[TL.SCENES[si][1]]['start']-0.55
    b=P[TL.SCENES[si+1][1]]['start']-0.55 if si+1<len(TL.SCENES) else OUTRO
    return a,b
def S_(v): return max(1e-3,v)
def bulb(x,cx,cy,s,lit,a,t=0):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    if lit>0:
        radial(x,0,-40,260,GOLD,0.45*lit*a)
        for i in range(10):
            ang=-math.pi/2+(i-4.5)*0.32; r0=150; r1=150+40*lit+8*math.sin(t*6+i)
            poly(x,[(r0*math.cos(ang),-40+r0*math.sin(ang)),(r1*math.cos(ang),-40+r1*math.sin(ang))],GOLD2,a*lit,fill=False,lw=7,close=False)
    x.new_path(); x.arc(0,-40,105,math.radians(140),math.radians(40)+2*math.pi*0); 
    x.new_path(); x.arc(0,-40,105,math.radians(135),math.radians(405)); x.line_to(48,90); x.line_to(-48,90); x.close_path()
    path=x.copy_path(); rgba(mix(CARD,GOLD2,lit),a,x); x.fill()
    x.append_path(path); x.set_line_width(8); rgba(WHITE,a,x); x.stroke()
    for k in range(3): rrect(x,-50+k*4,98+k*22,100-k*8,16,8,WHITE,a)
    poly(x,[(-25,60),(-12,0),(0,30),(12,0),(25,60)],NAVY1 if lit>0.5 else WHITE,a*0.8,fill=False,lw=5,close=False)
    x.restore()
def check(x,cx,cy,s,k,c,a,lw=8):
    if k>0: glowline(x,partial([(cx-22*s,cy),(cx-6*s,cy+16*s),(cx+24*s,cy-18*s)],k),c,a,lw)
def magnifier(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    circle(x,0,0,80,CARD,a*0.8); circle(x,0,0,80,WHITE,a,fill=False,lw=12)
    poly(x,[(58,58),(130,130)],GOLD,a,fill=False,lw=24,close=False)
    q=sprite("؟",110,(246,196,72),b'Black'); blit(x,q,0,4,1,a)
    x.restore()
def trophy(x,cx,cy,s,a,t):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    radial(x,0,-20,240,GOLD,0.35*a)
    x.new_path(); x.move_to(-90,-120); x.line_to(90,-120); x.curve_to(90,10,40,50,0,55); x.curve_to(-40,50,-90,10,-90,-120); x.close_path(); rgba(GOLD,a,x); x.fill()
    for sg in (-1,1):
        x.new_path(); x.arc(sg*92,-70,38,math.pi/2 if sg>0 else -math.pi/2, -math.pi/2 if sg>0 else math.pi/2) if False else None
        x.new_path(); x.arc_negative(sg*90,-75,36,math.pi*0.5,math.pi*1.5) if sg<0 else x.arc(sg*90,-75,36,-math.pi*0.5,math.pi*0.5)
        x.set_line_width(12); rgba(GOLD,a,x); x.stroke()
    rrect(x,-14,50,28,50,6,GOLD,a); rrect(x,-60,96,120,26,10,GOLD2,a); rrect(x,-75,118,150,22,8,WHITE,a)
    poly(x,star_pts(0,-50,32,rot=-math.pi/2),WHITE,a)
    for i in range(5):
        ang=t*0.8+i*1.26; r=170+10*math.sin(t*3+i); poly(x,star_pts(r*math.cos(ang),-30+r*0.6*math.sin(ang),10,rot=t+i),GOLD2,a*0.8)
    x.restore()
def tile(x,cx,cy,w,h,a,active):
    rrect(x,cx-w/2,cy-h/2+8,w,h,28,(0,0,0),0.25*a); rrect(x,cx-w/2,cy-h/2,w,h,28,CARD,a)
    rrect(x,cx-w/2,cy-h/2,w,h,28,GOLD if active else (0.35,0.47,0.75),a,fill=False,lw=5 if active else 2.5)
def num_node(x,cx,cy,r,n,a,active,k=1):
    if k<=0: return
    rr=r*k*(1.18 if active else 1)
    if active: radial(x,cx,cy,rr*2.6,GOLD,0.35*a)
    circle(x,cx,cy,rr,GOLD if active else CARD,a); circle(x,cx,cy,rr,WHITE if not active else GOLD2,a,fill=False,lw=4)
    sp=sprite('٠١٢٣٤٥'[n],int(r*1.1),(14,32,80) if active else (255,255,255),b'Black'); blit(x,sp,cx,cy-r*0.08,k*(1.18 if active else 1),a)
def mi(code,size,col=(255,255,255)): return sprite(chr(code),size,col,kind='mi',rtl=False)
def chip(x,txt,cx,cy,k,a,gold=False,icon=None,fs=40):
    if k<=0: return
    sp=sprite(txt,fs,(14,32,80) if gold else (255,255,255),b'Bold'); w=sp['adv']+60+(56 if icon else 0); h=fs+34
    x.save(); x.translate(cx,cy); x.scale(S_(k),S_(k))
    rrect(x,-w/2,-h/2+6,w,h,h/2,(0,0,0),0.25*a); rrect(x,-w/2,-h/2,w,h,h/2,GOLD if gold else CARD,a)
    if not gold: rrect(x,-w/2,-h/2,w,h,h/2,GOLD,a,fill=False,lw=3)
    if icon: ic=mi(icon,fs+4,(14,32,80) if gold else (246,196,72)); blit(x,ic,w/2-38,0,1,a); blit(x,sp,-28,-3,1,a)
    else: blit(x,sp,0,-3,1,a)
    x.restore()
def icon_disc(x,code,cx,cy,r,k,a,gold=False,glow=0):
    if k<=0: return
    if glow>0: radial(x,cx,cy,r*2.6*k,GOLD,0.4*glow*a)
    circle(x,cx,cy,r*k,GOLD if gold else CARD,a); circle(x,cx,cy,r*k,WHITE if not gold else GOLD2,a*0.8,fill=False,lw=4)
    blit(x,mi(code,int(r*1.15),(14,32,80) if gold else (255,255,255)),cx,cy,k,a)
def monitor(x,cx,cy,s,a,glow=0,t=0):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    if glow>0: radial(x,0,-20,260,(0.4,0.6,1.0),0.35*glow*a)
    rrect(x,-160,-120,320,210,18,WHITE,a); rrect(x,-146,-106,292,182,10,(0.12,0.3,0.62),a)
    for i in range(3): rrect(x,-120,-80+i*44,lerp(80,220,(math.sin(t*2+i)+1)/2),22,8,mix((0.3,0.5,0.9),GOLD,i==1),a*0.9)
    rrect(x,-20,90,40,40,4,WHITE,a); rrect(x,-80,126,160,16,8,WHITE,a); x.restore()
def board(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    rrect(x,-170,-120,340,220,12,(0.55,0.4,0.25),a); rrect(x,-156,-106,312,192,6,(0.12,0.3,0.22),a)
    for i in range(3): poly(x,[(-120,-60+i*45),(lerp(0,110,i/2),-60+i*45)],WHITE,a*0.7,fill=False,lw=5,close=False)
    x.restore()
def book(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    poly(x,[(0,-90),(-170,-110),(-170,100),(0,120)],WHITE,a); poly(x,[(0,-90),(170,-110),(170,100),(0,120)],(0.9,0.92,0.97),a)
    for i in range(4):
        poly(x,[(-140,-60+i*40),(-30,-50+i*40)],SOFT,a,fill=False,lw=4,close=False); poly(x,[(30,-50+i*40),(140,-60+i*40)],SOFT,a,fill=False,lw=4,close=False)
    poly(x,[(0,-90),(0,120)],NAVY2,a,fill=False,lw=4,close=False); x.restore()
def tablet(x,cx,cy,s,a,t):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    rrect(x,-170,-120,340,240,26,WHITE,a); rrect(x,-152,-102,304,204,12,(0.12,0.3,0.62),a)
    for i in range(3):
        pr=((t*0.8+i/3)%1); circle(x,40,10,10+pr*70,GOLD2,(1-pr)*a,fill=False,lw=4)
    circle(x,40,10,12,GOLD,a); rrect(x,-130,-80,90,60,10,GOLD,a*0.9); rrect(x,-130,0,90,60,10,(0.3,0.5,0.9),a*0.9); x.restore()
def paper(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    poly(x,[(-110,-150),(70,-150),(110,-110),(110,150),(-110,150)],WHITE,a)
    for i in range(6): poly(x,[(-80,-100+i*40),(lerp(80,20,i%2),-100+i*40)],SOFT,a,fill=False,lw=5,close=False)
    x.restore()
def phone(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    rrect(x,-95,-170,190,340,30,WHITE,a); rrect(x,-82,-145,164,280,14,(0.12,0.3,0.62),a)
    for r in range(3):
        for c in range(3): rrect(x,-66+c*48,-125+r*60,38,38,10,GOLD if (r+c)%3==0 else (0.35,0.55,0.95),a)
    circle(x,0,152,9,NAVY2,a); x.restore()
def morph_pair(x,t,t0,draw_old,draw_new,cx,cy,s,hold_end=None):
    """old object appears, flips into new at t0"""
    k=seg(t,t0-0.15,t0+0.45)
    if k<0.5:
        sx=1-2*k if k>0 else 1; x.save(); x.translate(cx,cy); x.scale(S_(sx),1); x.translate(-cx,-cy); draw_old(x,cx,cy,s,1); x.restore()
    else:
        sx=2*k-1; x.save(); x.translate(cx,cy); x.scale(S_(sx),1); x.translate(-cx,-cy); draw_new(x,cx,cy,s,1); x.restore()
        if k<1:
            for i in range(8):
                ang=i*math.pi/4; r=150*s+80*k; poly(x,star_pts(cx+r*math.cos(ang),cy+r*math.sin(ang),10*(1-k)+2,rot=t),GOLD2,1-k)
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    kb=back(seg(t,wt(0,3)-0.2,wt(0,3)+0.3)); km=back(seg(t,wt(1,3)-0.2,wt(1,3)+0.3))
    k0=eo(seg(t,s0-0.4,s0+0.2))
    if k0<=0: return
    glowline(x,[(0,-330),(0,330)],SOFT,0.4*k0,3)
    if kb>0:
        icon_disc(x,0xea4a,230,-40,140,kb,1,gold=True,glow=1)
        sp=sprite("التعلّم",48,(246,196,72),b'Black'); blit(x,sp,230,170,kb,1)
    if km>0:
        monitor(x,-230,-40,0.85*km,1,glow=0.6,t=t)
        sp=sprite("شكل العرض",44,(255,255,255),b'Black'); blit(x,sp,-230,170,km,1)
    kq=back(seg(t,ph(1)['end']-0.4,ph(1)['end']))
    if kq>0:
        circle(x,0,-40,70*kq,NAVY1,1); circle(x,0,-40,70*kq,GOLD,1,fill=False,lw=5)
        q=sprite("؟",100,(246,196,72),b'Black'); blit(x,q,0,-44,kq*(1+0.05*math.sin(t*5)),1)
def scene_2(x,t):
    s2,s3,s4,s5=[ph(i)['start'] for i in (2,3,4,5)]
    row=eio(seg(t,s5-0.2,s5+0.5))
    pos=[(lerp(0,250,row),lerp(-60,-260,row)),(lerp(0,0,row),lerp(-60,-260,row)),(lerp(0,-250,row),lerp(-60,-260,row))]
    sc=lerp(1.1,0.5,row)
    items=[(s2,lambda x,cx,cy,s,a:board(x,cx,cy,s,a),lambda x,cx,cy,s,a:monitor(x,cx,cy,s,a,glow=0.6,t=t)),
           (s3,lambda x,cx,cy,s,a:book(x,cx,cy,s,a),lambda x,cx,cy,s,a:tablet(x,cx,cy,s,a,t)),
           (s4,lambda x,cx,cy,s,a:paper(x,cx,cy,s,a*0.95),lambda x,cx,cy,s,a:phone(x,cx,cy,s*0.85,a))]
    for i,(st,o,n) in enumerate(items):
        nxt=items[i+1][0] if i<2 else s5
        k=back(seg(t,st-0.3,st+0.1))
        if k<=0: continue
        if row<=0:
            if t>=nxt-0.2 and i<2: continue
            ko=1-eo(seg(t,nxt-0.45,nxt-0.2)) if i<2 else 1
            morph_pair(x,t,ph(i+2)['words'][-1]['start'],o,n,0,-60,1.1*k*ko)
        else:
            cx,cy=pos[i]; n(x,cx,cy,sc,1)
    # learning meter unchanged
    if row>0:
        a=row; base=320
        h1=170; h2=170*back(seg(t,wt(5,3)-0.1,wt(5,3)+0.5))
        glowline(x,[(-240,base),(240,base)],WHITE,a,4)
        rrect(x,50,base-h1,130,h1,14,SOFT,a); rrect(x,-180,base-h2,130,h2,14,GOLD,a)
        for lab,lx in (("قبل",115),("بعد",-115)): blit(x,sprite(lab,38,(255,255,255),b'Bold'),lx,base+36,1,a)
        sp=sprite("التعلّم",40,(246,196,72),b'Black'); blit(x,sp,0,base-230,1,a)
        kq=back(seg(t,wt(5,4)-0.1,wt(5,4)+0.4))
        if kq>0: q=sprite("؟",120,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,0,base-120,kq*(1+0.05*math.sin(t*5)),a)
def scene_2_row(x,t):
    pass
def scene_3(x,t):
    s6,s7,s8=[ph(i)['start'] for i in (6,7,8)]
    # lesson cards, one glows
    kc=seg(t,s6-0.3,s6+0.3); kco=eio(seg(t,s7-0.2,s7+0.3))
    if kc>0 and kco<1:
        a=eo(kc)*(1-kco)
        for i in range(3):
            k=back(seg(t,s6+i*0.18,s6+0.4+i*0.18)); cx=(1-i)*240
            win=seg(t,ph(6)['words'][-1]['start']-0.1,ph(6)['words'][-1]['start']+0.4) if i==1 else 0
            if k<=0: continue
            x.save(); x.translate(cx,-20); s=k*(1+0.12*win); x.scale(S_(s),S_(s))
            if win>0: radial(x,0,0,260,GOLD,0.45*win*a)
            rrect(x,-100,-140,200,280,22,GOLD if win>0.5 else CARD,a); rrect(x,-100,-140,200,280,22,WHITE,a*0.6,fill=False,lw=3)
            blit(x,mi(0xe30c,90,(14,32,80) if win>0.5 else (255,255,255)),0,-40,1,a)
            blit(x,sprite("حصة",36,(14,32,80) if win>0.5 else (255,255,255),b'Bold'),0,70,1,a)
            x.restore()
            if win>=1:
                for j in range(6):
                    ang=t+j*1.05; poly(x,star_pts(cx+170*math.cos(ang),-20+190*math.sin(ang),10,rot=t),GOLD2,a)
    # classroom: screen present, students passive -> active
    kr=seg(t,s7-0.25,s7+0.35)
    if kr>0:
        a=eo(kr); act=eo(seg(t,wt(8,3)-0.2,wt(8,3)+0.5))
        monitor(x,0,-230,0.8,a,glow=1-act*0.5,t=t)
        for r in range(2):
            for c in range(3):
                j=r*3+c; px,py=(c-1)*200,70+r*170
                rrect(x,px-70,py+40,140,20,8,(0.4,0.48,0.65),a)
                person(x,px,py+36,0.9,mix(SOFT,GOLD if j%2==0 else WHITE,act),a)
                if act>0:
                    k=back(seg(t,wt(8,3)+j*0.1,wt(8,3)+0.4+j*0.1))
                    if k>0:
                        ic=[0xe925,0xe0f0,0xe2db,0xe0bf,0xe913,0xe925][j]
                        icon_disc(x,ic,px+55,py-75,28,k,a,gold=(j%2==0))
def scene_4(x,t):
    rows=[(9,0xe8f4,"المشاهدة",0xe925,"المشاركة"),(10,0xe156,"التلقّي",0xe2db,"التحليل"),(11,0xe88b,"انتظار التصحيح",0xe0bf,"تغذية راجعة")]
    for i,(pi,ic1,l1,ic2,l2) in enumerate(rows):
        st=ph(pi)['start']; y=-250+i*230
        k1=back(seg(t,st-0.3,st+0.15))
        if k1<=0: continue
        rrect(x,-420,y-90,840,180,40,CARD,0.55*eo(k1)); rrect(x,-420,y-90,840,180,40,WHITE,0.15,fill=False,lw=2)
        icon_disc(x,ic1,300,y-15,62,k1,0.85)
        blit(x,sprite(l1,32,(200,210,235),b'Bold'),300,y+68,k1,0.9)
        to=ph(pi)['words'][-1]['start'] if pi!=11 else ph(pi)['words'][-3]['start']
        ka=eo(seg(t,to-0.6,to-0.05))
        if ka>0:
            pts=[(220,y-15),(lerp(220,-200,ka),y-15)]; glowline(x,pts,GOLD,1,7)
            if ka>0.1: arrow_shape(x,pts[1][0]+20,y-15,pts[1][0],y-15,GOLD,1,lw=1,head=26)
        k2=back(seg(t,to-0.1,to+0.35))
        if k2>0:
            icon_disc(x,ic2,-300,y-15,70,k2,1,gold=True,glow=1)
            blit(x,sprite(l2,34,(246,196,72),b'Black'),-300,y+70,k2,1)
            if pi==11:
                kt=seg(t,ph(11)['end']-0.6,ph(11)['end']+0.3)
                glowline(x,partial([(-220,y+40),(-170,y+10),(-130,y+25),(-80,y-40)],eo(kt)),GOLD2,1,5)
def scene_5(x,t):
    s12,s13,s14,s15=[ph(i)['start'] for i in (12,13,14,15)]
    cur=-1 if t<s13-0.1 else 0 if t<s14-0.1 else 1 if t<s15-0.1 else 2
    xs=[220,0,-220]; ny=-310
    kin=eo(seg(t,wt(12,len(P[12]['words'])-2)-0.3,wt(12,len(P[12]['words'])-2)+0.2))
    for i in range(3):
        kk=back(seg(t,wt(12,len(P[12]['words'])-2)-0.1+i*0.15,wt(12,len(P[12]['words'])-2)+0.3+i*0.15))
        if i<2 and kin>0: poly(x,[(xs[i]-60,ny),(xs[i+1]+60,ny)],GOLD if cur>i else SOFT,(0.9 if cur>i else 0.35)*kin,fill=False,lw=6,close=False)
        if kk>0: num_node(x,xs[i],ny,52,i+1,1,i==cur,kk)
    # P12 tool with question
    k=back(seg(t,s12-0.3,s12+0.2)); ko=eio(seg(t,s13-0.2,s13+0.25))
    if k>0 and ko<1:
        tablet(x,0,60,1.0*k,1-ko,t); q=sprite("؟",120,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,200,-110,k,1-ko)
    # P13 what will learner do
    k=back(seg(t,s13-0.2,s13+0.3)); ko=eio(seg(t,s14-0.2,s14+0.25))
    if k>0 and ko<1:
        a=1-ko; person(x,-130,200,2.0*k,WHITE,a); tablet(x,120,90,0.65*k,a,t)
        for j,ic in enumerate((0xe913,0xe2db,0xe0bf)):
            kk=back(seg(t,wt(13,1)+j*0.15,wt(13,1)+0.35+j*0.15)); icon_disc(x,ic,-260+j*130,-130,46,kk,a,gold=(j==0))
    # P14 how will I know: quiz + check
    k=back(seg(t,s14-0.2,s14+0.3)); ko=eio(seg(t,s15-0.2,s15+0.25))
    if k>0 and ko<1:
        a=1-ko; x.save(); x.translate(0,60); x.scale(S_(k),S_(k))
        rrect(x,-200,-200,400,400,30,WHITE,a); rrect(x,-200,-200,400,80,30,GOLD,a)
        blit(x,mi(0xf04c,54,(14,32,80)),0,-160,1,a)
        for j in range(3):
            y=-80+j*90; rrect(x,-150,y-14,220,28,10,SOFT,a); rrect(x,100,y-24,48,48,10,NAVY2,a,fill=False,lw=4)
            check(x,124,y,0.85,eo(seg(t,wt(14,2)+j*0.2,wt(14,2)+0.3+j*0.2)),GOLD,a,lw=6)
        x.restore()
    # P15 added value: without vs with
    k=seg(t,s15-0.2,s15+0.3)
    if k>0:
        a=eo(k); base=330
        glowline(x,[(-260,base),(260,base)],WHITE,a,4)
        h1=150*back(seg(t,s15,s15+0.5)); h2=330*back(seg(t,wt(15,3)-0.1,wt(15,3)+0.6),1.2)
        rrect(x,60,base-h1,150,h1,16,SOFT,a); rrect(x,-210,base-h2,150,h2,16,GOLD,a)
        blit(x,sprite("دونها",36,(255,255,255),b'Bold'),135,base+36,1,a); blit(x,sprite("معها",36,(246,196,72),b'Black'),-135,base+36,1,a)
        if h2>200: icon_disc(x,0xe147,-135,base-h2-70,44,back(seg(t,wt(15,3)+0.3,wt(15,3)+0.7)),a,gold=True,glow=1)
def scene_6(x,t):
    s16,s17,s18=[ph(i)['start'] for i in (16,17,18)]
    kn=back(seg(t,s16-0.3,s16+0.2)); kno=eio(seg(t,s17-0.15,s17+0.3))
    if kn>0 and kno<1:
        a=1-kno; x.save(); x.translate(0,20); x.rotate(-0.05); x.scale(S_(kn),S_(kn))
        rrect(x,-170,-160,340,320,16,(0,0,0),0.25*a); rrect(x,-170,-170,340,320,16,GOLD2,a)
        for i in range(3): rrect(x,-120,-60+i*60,240-i*50,14,7,NAVY2,0.5*a)
        x.restore()
        kp=eo(seg(t,s16,s16+0.35)); py=lerp(-480,-150,kp)
        circle(x,0,py,34,(0.85,0.2,0.25),a); circle(x,-10,py-10,10,WHITE,0.6*a); poly(x,[(0,py+30),(0,py+70)],SOFT,a,fill=False,lw=6,close=False)
    # changing the screen is not it
    k=back(seg(t,s17-0.2,s17+0.3)); ko=eio(seg(t,s18-0.15,s18+0.35))
    if k>0 and ko<1:
        a=1-ko; flip=seg(t,wt(17,4)-0.2,wt(17,4)+0.4)
        x.save(); x.translate(0,-30); x.scale(S_(abs(math.cos(flip*math.pi))),1); x.translate(0,30)
        monitor(x,0,-30,1.1*k,a,glow=0.5,t=t); x.restore()
        ban(x,0,-30,240,seg(t,wt(17,5)-0.05,wt(17,5)+0.5),a,lw=13)
    # learning experience orbit
    k=seg(t,s18-0.2,s18+0.4)
    if k>0:
        a=eo(k); icon_disc(x,0xea4a,0,-30,120,back(k),a,gold=True,glow=1)
        for j,ic in enumerate((0xe925,0xe2db,0xe0bf,0xe0f0,0xe913,0xe8e5)):
            kk=back(seg(t,wt(18,1)+j*0.1,wt(18,1)+0.35+j*0.1)); ang=t*0.5+j*math.pi/3
            px,py=250*math.cos(ang),-30+230*math.sin(ang)
            if kk>0: glowline(x,[(0,-30),(px,py)],GOLD,0.3*kk,3); icon_disc(x,ic,px,py,46,kk,a)
def scene_7(x,t):
    s19,s20=ph(19)['start'],ph(20)['start']
    fin=eio(seg(t,VEND+0.1,VEND+0.8))
    # goal crossed
    k=back(seg(t,s19-0.3,s19+0.2)); ko=eio(seg(t,s20-0.2,s20+0.3))
    if k>0 and ko<1:
        a=1-ko
        for i,(r,c) in enumerate(((170,WHITE),(125,CARD),(80,WHITE),(38,GOLD))): circle(x,0,-20,r*k,c,a)
        blit(x,sprite("غاية",40,(14,32,80),b'Black'),0,190,k,a)
        ban(x,0,-20,240,seg(t,wt(19,2)-0.1,wt(19,2)+0.5),a,lw=13)
    # means -> deeper learning, lasting impact
    k=seg(t,s20-0.2,s20+0.4)
    if k>0 and fin<1:
        a=eo(k)*(1-fin)
        icon_disc(x,0xe869,240,-150,70,back(k),a,gold=False); blit(x,sprite("وسيلة",34,(255,255,255),b'Bold'),240,-50,1,a)
        kd=eo(seg(t,wt(20,3)-0.2,wt(20,3)+0.6))
        for j in range(5):
            r=60+j*45; circle(x,-60,120,r*kd,mix(GOLD,(0.1,0.2,0.45),j/5),a*(0.9-j*0.12),fill=False,lw=8)
        icon_disc(x,0xea4a,-60,120,55,kd,a,gold=True,glow=kd)
        if kd>0: glowline(x,partial([(180,-110),(40,10),(-10,60)],kd),GOLD2,a,5)
        kr=seg(t,wt(20,len(P[20]['words'])-1)-0.1,VEND+2)
        for j in range(3):
            pr=((t-s20)*0.5+j/3)%1
            if kr>0: circle(x,-60,120,100+pr*380,GOLD,(1-pr)*0.5*min(1,kr*3)*a,fill=False,lw=4)
    # hero statement
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        monitor(x,230,-230,0.55,a*0.6,glow=0,t=t); ban(x,230,-230,130,seg(t,VEND+0.3,VEND+0.9),a*0.8,lw=9)
        icon_disc(x,0xea4a,-230,-230,95,back(seg(t,VEND+0.9,VEND+1.4)),a,gold=True,glow=1)
        l1=sprite("لا تقِس نجاح التقنية",70,(255,255,255),b'Black'); l2=sprite("بما تعرضه الشاشة…",62,(255,255,255),b'Bold')
        l3=sprite("بل بما يحدث في",66,(246,196,72),b'Black',glow=(245,180,40)); l4=sprite("عقل المتعلّم",86,(246,196,72),b'Black',glow=(245,180,40))
        e1=eo(fin); e2=eo(seg(t,VEND+1.0,VEND+1.7))
        blit(x,l1,0,-20+(1-e1)*30,1,a); blit(x,l2,0,70+(1-e1)*30,1,a*e1)
        blit(x,l3,0,190+(1-e2)*30,1,a*e2); blit(x,l4,0,290+(1-e2)*30,1+0.03*math.sin(t*3),a*e2)
        glowline(x,[(-280*e2,360),(280*e2,360)],GOLD,a*0.7,4)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5,scene_6,scene_7]

# ---------------- kinetic typography ----------------
TEXT_Y=1330
def kinetic(x,t):
    for i,p in enumerate(P):
        nxt=P[i+1]['start']-0.12 if i+1<len(P) else VEND+0.15
        if t<p['start']-0.3 or t>nxt+0.45: continue
        words=p['words']; sps=[]
        for w in words:
            if w['key']: sps.append(sprite(w['word'],112,(246,196,72),b'Black',glow=(245,180,40)))
            else: sps.append(sprite(w['word'],90,(255,255,255),b'Bold'))
        gap=34; lines=[[]]; lw=[0]
        for w,sp in zip(words,sps):
            if lines[-1] and lw[-1]+gap+sp['adv']+(16 if w['key'] else 0)>960: lines.append([]); lw.append(0)
            lw[-1]+=(gap if lines[-1] else 0)+sp['adv']+(16 if w['key'] else 0); lines[-1].append((w,sp))
        lh=150; y0=TEXT_Y-(len(lines)-1)*lh/2
        out=seg(t,nxt,nxt+0.35)
        for li,ln in enumerate(lines):
            xr=W/2+lw[li]/2
            for j,(w,sp) in enumerate(ln):
                ex=16 if w['key'] else 0; cx=xr-(sp['adv']+ex)/2; xr-=sp['adv']+ex+gap
                k=seg(t,w['start']-0.12,w['start']+0.3); 
                if k<=0: continue
                sc=back(k,2.2)*0.4+0.6 if k<1 else 1
                if w['key']: sc*= 1+0.05*math.sin(seg(t,w['start'],w['start']+0.6)*math.pi)
                ko=eo(cl(out*1.6-j*0.12))
                cy=y0+li*lh+lerp(40,0,eo(k))-ko*70
                blit(x,sp,cx,cy,sc*(1-0.25*ko),eo(k)*(1-ko))
                if w['key']:
                    ku=eo(seg(t,w['start']+0.15,w['start']+0.55))*(1-ko)
                    if ku>0:
                        half=sp['adv']/2; pts=[(cx+half-2*half*u,cy+58+8*math.sin(u*math.pi)) for u in np.linspace(0,1,24)]
                        glowline(x,partial(pts,ku),GOLD,1-ko,6)
# ---------------- chrome: background, logo, progress, waveform ----------------
G={}
def init():
    global P
    P=TL.compute(); rng=np.random.default_rng(11)
    G['pt']=np.c_[rng.uniform(0,W,70),rng.uniform(0,H,70),rng.uniform(1.5,4,70),rng.uniform(0,6.3,70),rng.uniform(10,35,70)]
    w=wave.open(ROOT+'/v8/voice.wav'); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
    sr=16000; nb=24; f=np.zeros((NF,nb)); edges=np.geomspace(90,5000,nb+1); fr=np.fft.rfftfreq(1024,1/sr); win=np.hanning(1024)
    for i in range(NF):
        c=int((i/FPS-OFF)*sr)
        if c<0 or c+1024>len(a): continue
        sp=np.abs(np.fft.rfft(a[c:c+1024]*win)); f[i]=[sp[(fr>=edges[k])&(fr<edges[k+1])].mean() for k in range(nb)]
    f=np.log1p(f*4); f/=np.percentile(f[f>0],98)+1e-6
    for i in range(1,NF): f[i]=np.maximum(f[i],f[i-1]*0.8)
    G['feat']=np.clip(f,0,1.1)
import wave
def background(x,t):
    g=cairo.LinearGradient(0,0,0,H); g.add_color_stop_rgb(0,*NAVY1); g.add_color_stop_rgb(1,*NAVY2); x.set_source(g); x.paint()
    # scene accent glow
    radial(x,VC[0]+80*math.sin(t*0.3),VC[1]-40,620,(0.2,0.35,0.7),0.35)
    radial(x,W*0.85,H*0.85+60*math.sin(t*0.4),520,GOLD,0.06)
    # grid dots
    rgba(WHITE,0.06,x)
    for yy in range(40,H,60):
        for xx in range(30,W,60): x.rectangle(xx-1.5,yy-1.5,3,3)
    x.fill()
    for (px,py,s,ph_,v) in G['pt']:
        yy=(py-v*t)%H; a=0.25+0.25*math.sin(t*1.5+ph_); circle(x,px,yy,s,GOLD if s>2.8 else WHITE,a*0.6)
def corner_logo(x,t):
    k=back(seg(t,0.15,0.85)); ko=eio(seg(t,OUTRO-0.1,OUTRO+0.5))
    if k<=0 or ko>=1: return
    lg=logo_surface(138); R=92
    cx=W-40-R; cy=44+R
    x.save(); x.translate(cx,cy); s=max(1e-3,k*(1-ko)); x.scale(s,s)
    circle(x,0,7,R,(0,0,0),0.28)
    circle(x,0,0,R,WHITE,1); circle(x,0,0,R,GOLD,1,fill=False,lw=5)
    x.set_source_surface(lg['s'],-lg['w']/2,-lg['h']/2+2); x.paint(); x.restore()
def hud(x,t):
    a=eo(seg(t,0.4,1.0))*(1-eio(seg(t,OUTRO-0.2,OUTRO+0.3)))
    if a<=0: return
    pr=cl((t-OFF)/TL.VOICE_LEN); cx,cy=110,128
    circle(x,cx,cy,48,WHITE,0.15*a,fill=False,lw=8)
    x.new_path(); x.arc(cx,cy,48,-math.pi/2,-math.pi/2+2*math.pi*pr); x.set_line_width(8); x.set_line_cap(cairo.LINE_CAP_ROUND); rgba(GOLD,a,x); x.stroke()
    si=max([i for i in range(len(TL.SCENES)) if t>=scene_span(i)[0]] or [0])
    n=sprite('٠١٢٣٤٥٦٧٨٩'[si+1],46,(255,255,255),b'Black'); blit(x,n,cx,cy-4,1,a)
    tt=sprite(TL.SCENES[si][0],40,(246,196,72),b'Bold'); blit(x,tt,cx+80+tt['adv']/2,cy-4,1,a*eo(seg(t,scene_span(si)[0],scene_span(si)[0]+0.5)))
    # waveform
    f=G['feat'][min(NF-1,int(t*FPS))]; n=len(f); va=a*eo(seg(t,OFF-0.3,OFF+0.3))
    for k in range(n):
        hh=6+56*f[k]
        for sg in (-1,1):
            xx=W/2+sg*(10+k*16); rrect(x,xx-4,1710-hh/2,8,hh,4,GOLD if k%4==0 else WHITE,va*(0.75-0.45*k/n))
def outro(x,t):
    if t<OUTRO-0.1: return
    lt=t-OUTRO
    # logo badge entrance
    k=back(seg(lt,0.1,0.9),1.4); cx,cy=540,690
    if k>0:
        radial(x,cx,cy,520,GOLD,0.25*cl(lt))
        for i in range(3):
            r=300+i*34; a0=lt*(0.9 if i%2 else -0.7)+i; span=math.pi*(0.6+0.25*i)*eo(seg(lt,0.3+i*0.15,1.3+i*0.15))
            x.new_path(); x.arc(cx,cy,r*k,a0,a0+span); x.set_line_width(6-i*1.5); x.set_line_cap(cairo.LINE_CAP_ROUND); rgba(GOLD if i!=1 else WHITE,0.9,x); x.stroke()
            x.new_path(); x.arc(cx,cy,r*k,a0+math.pi,a0+math.pi+span*0.5); x.stroke()
        circle(x,cx,cy+10,262*k,(0,0,0),0.25); circle(x,cx,cy,262*k,WHITE,1); circle(x,cx,cy,262*k,GOLD,1,fill=False,lw=7)
        lg=logo_surface(400); x.save(); x.translate(cx,cy+6); s=max(1e-3,k*eo(seg(lt,0.1,0.5))); x.scale(s,s); x.set_source_surface(lg['s'],-lg['w']/2,-lg['h']/2); x.paint(); x.restore()
    kt=eo(seg(lt,1.1,1.6))
    if kt>0:
        sp=sprite("تابعونا على",64,(246,196,72),b'Bold'); blit(x,sp,540,1085+(1-kt)*30,1,kt)
        glowline(x,[(540-260*kt,1140),(540+260*kt,1140)],GOLD,kt*0.6,3)
    kk=seg(lt,1.45,2.25)
    if kk>0:
        e=back(kk,1.4); a=eo(kk*1.5)
        x.save(); x.translate(540,1290); x.scale(max(1e-3,lerp(0.6,1,e)),max(1e-3,lerp(0.6,1,e)))
        social_row(x,0,0,a,r=50,fs=62,pill=True,t=t,lt=lt)
        x.restore()
def social_icon(x,kind,cx,cy,r,a):
    if kind=='ig':
        g=cairo.LinearGradient(cx-r,cy+r,cx+r,cy-r); g.add_color_stop_rgb(0,0.99,0.69,0.27); g.add_color_stop_rgb(0.5,0.88,0.19,0.42); g.add_color_stop_rgb(1,0.51,0.23,0.71)
        x.save(); x.new_path(); x.arc(cx,cy,r,0,2*math.pi); x.clip(); x.set_source(g); x.paint_with_alpha(a); x.restore()
        ic=0xf16d
    else:
        circle(x,cx,cy,r,(0,0,0),a); circle(x,cx,cy,r,WHITE,a*0.9,fill=False,lw=max(1.5,r*0.06)); ic=0xe61b
    isp=sprite(chr(ic),max(12,int(r*1.12)),(255,255,255),kind='fa',rtl=False); blit(x,isp,cx,cy,1,a)
def social_row(x,cx,cy,a,r,fs,pill=False,t=0,lt=0,col=(14,32,80)):
    hs=sprite(HANDLE,fs,col,b'Bold',rtl=False); gap=r*0.35; sp=r*0.7
    total=4*r+gap+sp+hs['adv']; x0=cx-total/2
    if pill:
        pw=total+2*r*1.1; ph=2*r+36
        rrect(x,cx-pw/2,cy-ph/2+8,pw,ph,ph/2,(0,0,0),0.25*a); rrect(x,cx-pw/2,cy-ph/2,pw,ph,ph/2,WHITE,a)
        sw=seg(lt,2.7,3.5)
        if 0<sw<1:
            x.save(); x.new_path(); x.rectangle(cx-pw/2,cy-ph/2,pw,ph); x.clip(); sx=lerp(cx-pw/2-100,cx+pw/2+100,sw)
            g=cairo.LinearGradient(sx-90,0,sx+90,0); g.add_color_stop_rgba(0,1,0.9,0.6,0); g.add_color_stop_rgba(0.5,1,0.85,0.4,0.45); g.add_color_stop_rgba(1,1,0.9,0.6,0)
            x.set_source(g); x.paint(); x.restore()
    # handle on the left, icons on the right (X then Instagram)
    blit(x,hs,x0+hs['adv']/2,cy-fs*0.06,1,a)
    ix=x0+hs['adv']+sp+r
    social_icon(x,'x',ix,cy,r,a); social_icon(x,'ig',ix+2*r+gap,cy,r,a)
def footer(x,t):
    a=eo(seg(t,0.6,1.3))*(1-eio(seg(t,OUTRO-0.2,OUTRO+0.3)))
    if a<=0: return
    rrect(x,540-250,1838-38,500,76,38,WHITE,0.08*a)
    social_row(x,540,1838,a*0.95,r=22,fs=34,col=(255,255,255))
def transition_ring(x,t):
    for si in range(1,len(TL.SCENES)):
        b=scene_span(si)[0]+0.4; k=seg(t,b-0.25,b+0.45)
        if 0<k<1: circle(x,VC[0],VC[1],60+900*eo(k),GOLD,0.55*(1-k),fill=False,lw=10*(1-k)+2)
    k=seg(t,OUTRO-0.3,OUTRO+0.5)
    if 0<k<1: circle(x,540,690,60+1100*eo(k),GOLD,0.6*(1-k),fill=False,lw=12*(1-k)+2)
def frame(fi):
    t=fi/FPS
    surf=cairo.ImageSurface(cairo.FORMAT_ARGB32,W,H); x=cairo.Context(surf); x.set_antialias(cairo.ANTIALIAS_BEST)
    background(x,t)
    # intro sweep lines
    ki=seg(t,0,1.1)
    if 0<ki<1:
        for k in range(3): glowline(x,[(lerp(-200,1300,eo(cl(ki*1.3-k*0.12))),620+k*40),(lerp(-400,1100,eo(cl(ki*1.3-k*0.12))),620+k*40)],GOLD,1-ki,4)
    for si in range(len(SCENE_FN)):
        a,b=scene_span(si)
        if t<a-0.2 or t>b+0.6: continue
        zin=eo(seg(t,a-0.1,a+0.55)); zout=eio(seg(t,b-0.15,b+0.45))
        if zin<=0 or zout>=1: continue
        s=lerp(0.78,1,zin)*lerp(1,1.4,zout); alpha=zin*(1-zout)
        x.push_group(); x.save(); x.translate(*VC); x.scale(s,s)
        SCENE_FN[si](x,t)
        x.restore(); x.pop_group_to_source(); x.paint_with_alpha(alpha)
    transition_ring(x,t)
    kinetic(x,t); hud(x,t); footer(x,t); corner_logo(x,t); outro(x,t)
    fade=eo(seg(t,0,0.35))*(1-eio(seg(t,DUR-0.6,DUR)))
    if fade<1: x.set_source_rgba(NAVY1[0],NAVY1[1],NAVY1[2],1-fade); x.paint()
    surf.flush(); a=np.ndarray((H,W,4),np.uint8,surf.get_data())
    return a[...,[2,1,0]].tobytes()
if __name__=='__main__':
    out=sys.argv[1]
    if len(sys.argv)>2:
        init()
        for tt in sys.argv[2].split(','): Image.frombytes('RGB',(W,H),frame(int(float(tt)*FPS))).save(f"{out}_{tt}.png")
        sys.exit()
    print('DUR',DUR,flush=True)
    p=subprocess.Popen(['ffmpeg','-y','-loglevel','error','-f','rawvideo','-pix_fmt','rgb24','-s',f'{W}x{H}','-r',str(FPS),'-i','-','-c:v','libx264','-preset','slow','-crf','19','-pix_fmt','yuv420p',out],stdin=subprocess.PIPE)
    with Pool(4,initializer=init) as pool:
        for b in pool.imap(frame,range(NF),chunksize=4): p.stdin.write(b)
    p.stdin.close(); p.wait(); print('done')
