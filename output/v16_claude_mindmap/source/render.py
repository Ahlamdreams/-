import cairo, numpy as np, math, sys, os, subprocess
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from multiprocessing import Pool
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import timeline as TL
ROOT=os.path.dirname(os.path.abspath(__file__))+'/..'
W,H,FPS=1080,1920,30
OFF=TL.OFF; VEND=OFF+TL.VOICE_LEN; OUTRO=VEND+3.4; DUR=OUTRO+4.6; NF=int(DUR*FPS)
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

def top5(x,t,cur,k=1):
    xs=[320,160,0,-160,-320]; ny=-330
    for i in range(4): poly(x,[(xs[i]-44,ny),(xs[i+1]+44,ny)],GOLD if cur>i else SOFT,(0.9 if cur>i else 0.35)*k,fill=False,lw=5,close=False)
    for i in range(5): num_node(x,xs[i],ny,40,i+1,k,i==cur,1)
def pinnote(x,t,s0,s1):
    kn=back(seg(t,s0-0.3,s0+0.2)); kno=eio(seg(t,s1-0.15,s1+0.3))
    if kn>0 and kno<1:
        a=1-kno; x.save(); x.translate(0,20); x.rotate(-0.05); x.scale(S_(kn),S_(kn))
        rrect(x,-170,-160,340,320,16,(0,0,0),0.25*a); rrect(x,-170,-170,340,320,16,GOLD2,a)
        for i in range(3): rrect(x,-120,-60+i*60,240-i*50,14,7,NAVY2,0.5*a)
        x.restore()
        kp=eo(seg(t,s0,s0+0.35)); py=lerp(-480,-150,kp)
        circle(x,0,py,34,(0.85,0.2,0.25),a); circle(x,-10,py-10,10,WHITE,0.6*a); poly(x,[(0,py+30),(0,py+70)],SOFT,a,fill=False,lw=6,close=False)
def bar(x,y,label,v,col,a,k):
    blit(x,sprite(label,34,(255,255,255),b'Bold'),250,y,k,a)
    rrect(x,-300,y-20,460,40,20,CARD,a*k); rrect(x,-300+460*(1-v),y-20,460*v,40,20,col,a*k)

PARTS=[("الدور",0xe7fd,"أنت معلّم علوم للصف الخامس",GOLD),("المهمة",0xe2e6,"اكتب ١٠ أسئلة عن دورة الماء",(0.45,0.65,1.0)),("المستويات",0xe8e5,"من التذكّر إلى التحليل والتطبيق",GOLD2),("الشكل",0xe6b1,"اختيار من متعدد مع الإجابة",(0.85,0.9,1.0))]
def tag(x,label,col,cx,cy,k,a,fs=28):
    sp=sprite(label,fs,(14,32,80),b'Black'); w=sp['adv']+34; h=fs+20
    x.save(); x.translate(cx,cy); x.scale(S_(k),S_(k)); rrect(x,-w/2,-h/2,w,h,h/2,col,a); blit(x,sp,0,-2,1,a); x.restore(); return w
def typed(x,sp,rx,cy,prog,a):
    """reveal sprite right-to-left, right edge at rx"""
    if prog<=0: return
    w=sp['adv']; x.save(); x.new_path(); x.rectangle(rx-w*prog-6,cy-40,w*prog+12,80); x.clip(); blit(x,sp,rx-w/2,cy,1,a); x.restore()
    if prog<1: rrect(x,rx-w*prog-10,cy-20,4,40,2,GOLD2,a*(0.5+0.5*math.sin(t_now[0]*12)))
t_now=[0]
def window(x,a,t):
    rrect(x,-410,-330,820,700,36,(0,0,0),0.25*a); rrect(x,-410,-340,820,700,36,(0.08,0.15,0.34),a)
    rrect(x,-410,-340,820,90,36,(0.12,0.24,0.52),a); rrect(x,-410,-290,820,40,0,(0.12,0.24,0.52),a)
    rrect(x,-410,-340,820,700,36,GOLD,a,fill=False,lw=3)
    icon_disc(x,0xf06c,350,-295,30,1,a,gold=True)
    blit(x,sprite("مساعد الذكاء الاصطناعي",30,(255,255,255),b'Bold'),180,-297,1,a)
    for j in range(3): circle(x,-370+j*28,-295,8,(0.4,0.5,0.75),a)

def cwin(x,a,t,x0=-410,y0=-340,w=820,h=700):
    rrect(x,x0,y0+10,w,h,36,(0,0,0),0.25*a); rrect(x,x0,y0,w,h,36,(0.08,0.15,0.34),a)
    rrect(x,x0,y0,w,90,36,(0.12,0.24,0.52),a); rrect(x,x0,y0+50,w,40,0,(0.12,0.24,0.52),a)
    rrect(x,x0,y0,w,h,36,GOLD,a,fill=False,lw=3)
    icon_disc(x,0xe65f,x0+w-60,y0+45,30,1,a,gold=True)
    blit(x,sprite("Claude",36,(255,255,255),b'Black',rtl=False),x0+w-150,y0+43,1,a)
    for j in range(3): circle(x,x0+40+j*28,y0+45,8,(0.4,0.5,0.75),a)
def game(x,cx,cy,s,a,t,vivid=0,timer=0,ans=1,score=30):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    g=cairo.LinearGradient(0,-300,0,300)
    c1=mix((0.1,0.2,0.45),(0.95,0.35,0.55),vivid); c2=mix((0.06,0.12,0.3),(0.98,0.6,0.2),vivid)
    g.add_color_stop_rgba(0,c1[0],c1[1],c1[2],a); g.add_color_stop_rgba(1,c2[0],c2[1],c2[2],a)
    x.new_path(); x.rectangle(-230,-300,460,600); x.set_source(g)
    x.save(); rrect(x,-230,-300,460,600,34,WHITE,0); x.new_path()
    r=34; x.arc(230-r,-300+r,r,-math.pi/2,0); x.arc(230-r,300-r,r,0,math.pi/2); x.arc(-230+r,300-r,r,math.pi/2,math.pi); x.arc(-230+r,-300+r,r,math.pi,1.5*math.pi); x.close_path(); x.clip()
    x.set_source(g); x.paint(); x.restore()
    rrect(x,-230,-300,460,600,34,WHITE,a*0.8,fill=False,lw=5)
    blit(x,mi(0xe838,40,(246,196,72)),150,-250,1,a)
    blit(x,sprite(''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(score)),36,(255,255,255),b'Black'),95,-252,1,a)
    blit(x,sprite("سؤال ٣ من ١٠",26,(220,230,255),b'Bold'),-100,-252,1,a)
    if timer>0:
        circle(x,-170,-180,34*timer,WHITE,a*0.25,fill=False,lw=6)
        x.new_path(); x.arc(-170,-180,34*timer,-math.pi/2,-math.pi/2+2*math.pi*(1-((t*0.2)%1))); x.set_line_width(6); rgba(GOLD2,a,x); x.stroke()
    rrect(x,-190,-150,380,120,22,WHITE,a)
    blit(x,sprite("١/٤ + ١/٤ = ؟",50,(14,32,80),b'Black'),0,-92,1,a)
    opts=["١/٢","٢/٨","١/٨","٢/٤"]
    for j,o in enumerate(opts):
        r_,q=j//2,j%2; px=100-q*200; py=20+r_*110
        ok=(j==0); sel=ans>0.5 and ok
        rrect(x,px-85,py-44,170,88,20,GOLD if sel else (1,1,1),a*(1 if sel else 0.22)); rrect(x,px-85,py-44,170,88,20,WHITE,a*0.7,fill=False,lw=3)
        blit(x,sprite(o,40,(14,32,80) if sel else (255,255,255),b'Black'),px,py-3,1,a)
    if ans>0.5:
        e=back(cl((ans-0.5)*2)); x.save(); x.translate(0,250); x.scale(S_(e),S_(e))
        rrect(x,-140,-36,280,72,36,WHITE,a); blit(x,sprite("أحسنت!",36,(14,32,80),b'Black'),20,-3,1,a); blit(x,mi(0xea65,40,(246,196,72)),-90,0,1,a); x.restore()
    x.restore()

BR=[("التبخّر",0xe430,(1.0,0.62,0.25),"الشمس تسخّن الماء",(170,-190)),("التكاثف",0xe2bd,(0.62,0.78,1.0),"تتكوّن الغيوم",(-170,-190)),("الهطول",0xe798,(0.35,0.6,1.0),"مطر وثلج",(-170,190)),("التجمّع",0xe176,(0.3,0.8,0.75),"أنهار وبحار",(170,190))]
def bookpage(x,cx,cy,s,a,rot=0,dense=1):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.rotate(rot); x.scale(S_(s),S_(s))
    rrect(x,-170,-230,340,460,14,(0,0,0),0.25*a); rrect(x,-170,-240,340,460,14,WHITE,a)
    blit(x,sprite("دورة الماء",34,(14,32,80),b'Black'),0,-200,1,a)
    for j in range(int(14*dense)):
        y=-160+j*26; w=[280,250,270,230,260][j%5]
        rrect(x,140-w,y,w,10,5,(0.72,0.76,0.86),a)
    x.restore()
def mindmap(x,cx,cy,t,a,kc,kb,kl,sc=1):
    x.save(); x.translate(cx,cy); x.scale(S_(sc),S_(sc))
    for i,(lab,ic,col,leaf,(bx,by)) in enumerate(BR):
        k=kb[i]
        if k<=0: continue
        ex,ey=bx*eo(k),by*eo(k)
        x.new_path(); x.move_to(0,0); x.curve_to(bx*0.2,by*0.7,bx*0.6,by*0.3,ex,ey); x.set_line_width(10); rgba(col,a,x); x.set_line_cap(cairo.LINE_CAP_ROUND); x.stroke()
        if k>=0.95:
            e=back(cl((k-0.95)*20))
            x.save(); x.translate(bx,by); x.scale(S_(e),S_(e))
            rrect(x,-95,-42,190,84,42,col,a); blit(x,sprite(lab,32,(14,32,80),b'Black'),18,-3,1,a); blit(x,mi(ic,36,(14,32,80)),-58,0,1,a); x.restore()
            kk=kl[i]
            if kk>0:
                lx,ly=bx*1.0,by+(90 if by>0 else -90)
                poly(x,[(bx,by+(42 if by>0 else -42)),(lx,ly+(-26 if by>0 else 26))],col,a*0.7,fill=False,lw=4,close=False)
                rrect(x,lx-110*kk,ly-26,220*kk,52,26,WHITE,a); blit(x,sprite(leaf,24,(14,32,80),b'Bold'),lx,ly-2,kk,a)
    if kc>0:
        e=back(kc); x.save(); x.scale(S_(e),S_(e)); radial(x,0,0,200,GOLD,0.5*a)
        circle(x,0,0,92,GOLD,a); circle(x,0,0,92,WHITE,a,fill=False,lw=5); blit(x,sprite("دورة الماء",32,(14,32,80),b'Black'),0,-3,1,a); x.restore()
    x.restore()
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    k=back(seg(t,s0-0.4,s0+0.2)); m=eio(seg(t,P[1]['words'][-3]['start']-0.3,P[1]['words'][-3]['start']+0.7))
    if k>0 and m<1:
        bookpage(x,0,-60,1.25*k*(1-m*0.5),1-m,rot=-0.03)
        kq=back(seg(t,P[0]['words'][-2]['start']-0.1,P[0]['words'][-2]['start']+0.4))
        if kq>0 and m<1:
            person(x,250,330,1.2*kq,SOFT,1-m)
            for j in range(3): blit(x,sprite("؟",60+j*12,(246,196,72),b'Black'),200+j*45,150-j*35,kq,(1-m))
    k=seg(t,s1-0.2,s1+0.3)
    if k>0:
        a=eo(k)
        x.save(); x.translate(0,-330); rrect(x,-140,-45,280,90,45,WHITE,a); blit(x,sprite("Claude",44,(14,32,80),b'Black',rtl=False),25,-3,1,a); blit(x,mi(0xe65f,44,(246,196,72)),-90,0,1,a); x.restore()
        if m>0:
            kb=[eo(seg(m,0.2+i*0.15,0.5+i*0.15)) for i in range(4)]
            mindmap(x,0,20,t,a,m,kb,[0,0,0,0],sc=0.9)
        kt=back(seg(t,P[1]['words'][-1]['start']-0.1,P[1]['words'][-1]['start']+0.4))
        if kt>0: clock(x,280,-330,44*kt,t*5,a)
def scene_2(x,t):
    t_now[0]=t
    s2,s3=ph(2)['start'],ph(3)['start']
    k=seg(t,s2-0.2,s2+0.3); ko=eio(seg(t,s3-0.2,s3+0.2))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        bookpage(x,0,40,1.0,a)
        snap=seg(t,P[2]['words'][1]['start']-0.1,P[2]['words'][1]['start']+0.4)
        x.save(); x.translate(lerp(380,0,eio(seg(t,s2,s2+0.6))),40)
        rrect(x,-200,-300,400,600,40,(0.05,0.08,0.15),a*0.35); rrect(x,-200,-300,400,600,40,WHITE,a,fill=False,lw=8)
        for (px,py,sx,sy) in ((-150,-250,1,1),(150,-250,-1,1),(-150,250,1,-1),(150,250,-1,-1)):
            poly(x,[(px,py+sy*40),(px,py),(px+sx*40,py)],GOLD,a,fill=False,lw=6,close=False)
        circle(x,0,340,36,WHITE,a); circle(x,0,340,28,GOLD,a)
        x.restore()
        if 0<snap<1: rgba(WHITE,0.8*(1-snap)*a,x); x.paint() if False else rrect(x,-540,-700,1080,1400,0,WHITE,0.7*(1-snap)*a)
        up=eio(seg(t,P[2]['words'][-2]['start']-0.2,ph(2)['end']+0.2))
        if up>0:
            icon_disc(x,0xe2c6,0,-330,60,back(up),a,gold=True,glow=0.6)
            rrect(x,-140,-430,280,70,35,WHITE,a*up); blit(x,sprite("Claude",34,(14,32,80),b'Black',rtl=False),0,-397,up,a)
    k=back(seg(t,s3-0.3,s3+0.1))
    if k>0:
        x.save(); x.scale(S_(k),S_(k)); a=1
        cwin(x,a,t)
        rrect(x,-380,-220,760,500,26,(0.05,0.1,0.25),a)
        x.save(); x.translate(240,-130); x.rotate(0.05); x.scale(0.35,0.35); bookpage(x,0,0,1,a); x.restore()
        blit(x,mi(0xe3f4,40,(246,196,72)),110,-200,1,a)
        lines=[("حوّل هذه الصفحة إلى",0.0),("خريطة ذهنية ملوّنة",0.33),("لطلاب الصف الخامس",0.66)]
        d=ph(3)['end']-P[3]['words'][2]['start'] if len(P[3]['words'])>2 else 2
        for i,(txt,f) in enumerate(lines):
            st=P[3]['words'][2]['start']+d*f; en=st+d*0.33
            if t<st-0.1: continue
            typed(x,sprite(txt,34,(255,255,255) if i!=1 else (246,196,72),b'Bold'),330,20+i*80,eo(seg(t,st,en)),a)
        pulse=seg(t,ph(4)['start']-0.1,ph(4)['start']+0.3)
        circle(x,-330,330,34*(1+0.25*math.sin(pulse*math.pi)),GOLD,a); blit(x,mi(0xe163,38,(14,32,80)),-330,330,1,a)
        x.restore()
def scene_3(x,t):
    s4,s5=ph(4)['start'],ph(5)['start']
    k=seg(t,s4-0.2,s4+0.3)
    if k<=0: return
    a=eo(k)
    kc=seg(t,P[4]['words'][4]['start']-0.3,P[4]['words'][4]['start']+0.3) if len(P[4]['words'])>4 else seg(t,s4+0.5,s4+1)
    br0=P[4]['words'][-2]['start']-0.6
    kb=[seg(t,br0+i*0.3,br0+0.6+i*0.3) for i in range(4)]
    kl=[eo(seg(t,ph(4)['end']-0.2+i*0.15,ph(4)['end']+0.2+i*0.15)) for i in range(4)]
    sh=eio(seg(t,s5,s5+0.6))
    mindmap(x,0,lerp(0,-120,sh),t,a,kc,kb,kl,sc=lerp(1.0,0.75,sh))
    if sh>0:
        qs=["ما سبب تبخّر الماء؟","كيف تتكوّن الغيوم؟","أين يتجمّع الماء؟"]
        for j,q in enumerate(qs):
            kk=back(seg(t,P[5]['words'][-3]['start']+j*0.2,P[5]['words'][-3]['start']+0.4+j*0.2)); y=230+j*80
            if kk<=0: continue
            rrect(x,-320,y-32,640*kk,64,18,WHITE,a); circle(x,280,y,22,GOLD,a)
            blit(x,sprite('٠١٢٣'[j+1],26,(14,32,80),b'Black'),280,y-2,kk,a); blit(x,sprite(q,28,(14,32,80),b'Bold'),20,y-2,kk,a)
def scene_4(x,t):
    s6=ph(6)['start']
    k=seg(t,s6-0.2,s6+0.3)
    if k<=0: return
    a=eo(k)
    bookpage(x,190,-60,0.8,a,rot=0.03)
    mindmap(x,-190,-60,t,a,1,[1,1,1,1],[0,0,0,0],sc=0.5)
    u=seg(t,s6+0.2,ph(6)['end']); magnifier(x,lerp(190,-190,eio(u)),-60+60*math.sin(u*math.pi*2),0.6,a)
    for i in range(4):
        ck=seg(t,s6+0.6+i*0.3,s6+0.9+i*0.3)
        if ck>0:
            px,py=-190+BR[i][4][0]*0.5+40,-60+BR[i][4][1]*0.5-20
            circle(x,px,py,16*back(ck),GOLD,a); check(x,px,py-1,0.45,1,NAVY1,a,lw=4)
    kk=back(seg(t,ph(6)['end']-0.4,ph(6)['end']+0.1))
    if kk>0: blit(x,sprite("مطابقة للكتاب",40,(246,196,72),b'Black'),0,290,kk,a)
def scene_5(x,t):
    s7,s8=ph(7)['start'],ph(8)['start']
    k=seg(t,s7-0.2,s7+0.3); ko=eio(seg(t,s8-0.15,s8+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        bookpage(x,200,-40,0.75,a*0.85,rot=0.05)
        ka=eo(seg(t,s7+0.6,s7+1.2)); glowline(x,partial([(80,-40),(-60,-40)],ka),GOLD,a,7)
        if ka>0.9: arrow_shape(x,-40,-40,-70,-40,GOLD,a,lw=1,head=26)
        km=seg(t,P[7]['words'][-3]['start']-0.3,P[7]['words'][-3]['start']+0.6)
        mindmap(x,-230,-40,t,a,km,[km]*4,[0,0,0,0],sc=0.42)
        kh=back(seg(t,P[7]['words'][-1]['start']-0.2,P[7]['words'][-1]['start']+0.3))
        if kh>0: person(x,-230,330,1.1*kh,GOLD,a); icon_disc(x,0xe0f0,-230,170,40,kh,a,gold=True,glow=0.6)
    fin=eio(seg(t,VEND+0.1,VEND+0.7))
    k=seg(t,s8-0.15,s8+0.4)
    if k>0 and fin<1:
        a=eo(k)*(1-fin)
        icon_disc(x,0xe866,0,-80,120,back(k),a,gold=True,glow=1); blit(x,sprite("احفظ",44,(246,196,72),b'Black'),0,80,back(k),a)
        for j,(ic,px,py) in enumerate(((0xe97a,-240,-200),(0xe87d,240,-190),(0xe0b7,-230,120),(0xe8dc,250,140))):
            kk=back(seg(t,s8+0.5+j*0.2,s8+0.9+j*0.2)); icon_disc(x,ic,px,py+8*math.sin(t*3+j),44,kk,a,gold=(j%2==1))
        kd=seg(t,P[8]['words'][-1]['start']-0.4,VEND+1)
        if kd>0:
            for j in range(3): blit(x,mi(0xe5db,64,(246,196,72)),0,190+j*50+8*math.sin(t*5),1,a*((math.sin(t*5-j)+1)/2))
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        e0=back(seg(t,VEND+0.1,VEND+0.6))
        bookpage(x,190,-220,0.45*e0,a,rot=0.05); mindmap(x,-190,-220,t,a,e0,[e0]*4,[0,0,0,0],sc=0.38)
        glowline(x,[(100,-220),(-80,-220)],GOLD,a*e0,6)
        e1=eo(seg(t,VEND+0.6,VEND+1.2)); e2=eo(seg(t,VEND+1.0,VEND+1.6))
        blit(x,sprite("صوّر الصفحة…",74,(255,255,255),b'Black'),0,60+(1-e1)*30,1,a*e1)
        blit(x,sprite("تحصل على خريطة ذهنية",76,(246,196,72),b'Black',glow=(245,180,40)),0,190+(1-e2)*30,1+0.03*math.sin(t*3),a*e2)
        blit(x,sprite("مع Claude",44,(200,210,235),b'Bold'),0,300,1,a*e2)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
def OVERLAY(x,t): pass

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
    w=wave.open(ROOT+'/v16/voice.wav'); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
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
    transition_ring(x,t); OVERLAY(x,t)
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
