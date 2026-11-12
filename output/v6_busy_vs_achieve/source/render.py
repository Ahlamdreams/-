import cairo, numpy as np, math, sys, os, subprocess
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from multiprocessing import Pool
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import timeline as TL
ROOT=os.path.dirname(os.path.abspath(__file__))+'/..'
W,H,FPS=1080,1920,30
OFF=TL.OFF; VEND=OFF+TL.VOICE_LEN; OUTRO=VEND+0.2; DUR=OUTRO+5.8; NF=int(DUR*FPS)
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
        path={'cairo':'Cairo.ttf','fa':'fa-brands.ttf'}[kind]
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
def cloud(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s)); x.new_path()
    for (px,py,r) in ((-110,10,70),(-40,-45,88),(55,-38,80),(125,12,62),(0,20,80)): x.arc(px,py,r,0,2*math.pi); x.new_sub_path()
    rgba(WHITE,a,x); x.fill(); rrect(x,-175,10,350,72,36,WHITE,a)
    circle(x,-150,130,18,WHITE,a); circle(x,-185,175,10,WHITE,a); x.restore()
def nums(n): return ''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(n))
def card_row(x,x0,y,w,a,k,gold=False):
    rrect(x,x0,y-26,w,52,14,GOLD if gold else WHITE,a*(0.95 if gold else 0.9))
    rrect(x,x0+w-46,y-14,28,28,7,NAVY2,a*0.8,fill=False,lw=4)
    rrect(x,x0+20,y-7,(w-90)*k,14,7,NAVY2 if gold else SOFT,a)
def diamond(x,cx,cy,r,a,t):
    radial(x,cx,cy,r*2.3,GOLD,0.45*a)
    top=[(cx-r,cy-r*0.35),(cx-r*0.55,cy-r*0.85),(cx+r*0.55,cy-r*0.85),(cx+r,cy-r*0.35),(cx,cy+r)]
    poly(x,top,GOLD,a); poly(x,[(cx-r,cy-r*0.35),(cx+r,cy-r*0.35),(cx,cy+r)],GOLD2,a*0.9)
    poly(x,[(cx-r*0.55,cy-r*0.85),(cx-r*0.2,cy-r*0.35),(cx,cy-r*0.85),(cx+r*0.2,cy-r*0.35),(cx+r*0.55,cy-r*0.85)],WHITE,a*0.35)
    poly(x,top,WHITE,a,fill=False,lw=4)
    for i in range(6):
        ang=t*0.9+i*1.05; rr=r*1.6+12*math.sin(t*3+i); poly(x,star_pts(cx+rr*math.cos(ang),cy+rr*math.sin(ang),9,rot=t),GOLD2,a*0.85)
def scene_1(x,t):
    s0,s1,s2,s3=[ph(i)['start'] for i in (0,1,2,3)]
    # thought cloud with ?
    kc=back(seg(t,s0-0.45,s0+0.15)); kco=eio(seg(t,s1-0.15,s1+0.35))
    if kc>0 and kco<1:
        s=kc*lerp(1.2,0.5,kco); cy=lerp(-20,-300,kco)
        radial(x,0,cy,300*s,GOLD,0.18); cloud(x,0,cy,s,1-kco)
        q=sprite("؟",150,(246,196,72),b'Black'); blit(x,q,0,cy-10*s,s*(1+0.06*math.sin(t*5)),1-kco)
    # busy gears + clock
    kg=back(seg(t,s1-0.25,s1+0.3)); kgo=eio(seg(t,s3-0.15,s3+0.35))
    up=eio(seg(t,s2-0.1,s2+0.4))
    if kg>0 and kgo<1:
        a=1-kgo; x.save(); x.translate(0,lerp(0,-170,up)); s=lerp(1,0.62,up); x.scale(s,s)
        sp=1+3*seg(t,wt(1,1),wt(1,1)+0.6)
        gear(x,-150,-20,95*kg,t*1.2*sp,GOLD,a); gear(x,10,-110,62*kg,-t*1.9*sp+0.2,WHITE,a,teeth=8); gear(x,40,60,48*kg,-t*2.4*sp,SOFT,a,teeth=8)
        clock(x,190,-10,92*kg,t*2.5*sp,a)
        x.restore()
    # aspiration star not reached
    km=seg(t,s2-0.1,s2+0.3); kmo=eio(seg(t,s3-0.15,s3+0.3))
    if km>0 and kmo<1:
        a=eo(km)*(1-kmo); sx,sy=230,170
        radial(x,sx,sy,170,GOLD,0.35*a); poly(x,star_pts(sx,sy,70*back(km),rot=-math.pi/2+t*0.3),GOLD2,a)
        kp=seg(t,wt(2,4)-0.3,ph(2)['end'])
        pts=[(lerp(-300,150,u),lerp(330,190,u)-120*math.sin(u*math.pi)) for u in np.linspace(0,eo(kp),30)]
        for j in range(0,len(pts)-1,2): poly(x,[pts[j],pts[j+1]],WHITE,a*0.8,fill=False,lw=5,close=False)
        if kp>=1:
            kx=seg(t,ph(2)['end'],ph(2)['end']+0.3); d=26
            glowline(x,partial([(150-d,190-d),(150+d,190+d)],kx),GOLD,a,9); glowline(x,partial([(150+d,190-d),(150-d,190+d)],kx),GOLD,a,9)
    # busy != achievement
    kn=seg(t,s3-0.2,s3+0.3)
    if kn>0:
        a=eo(kn)
        k1=back(seg(t,wt(3,1)-0.1,wt(3,1)+0.35)); k2=back(seg(t,wt(3,5)-0.15,wt(3,5)+0.35)); ke=back(seg(t,wt(3,3)-0.1,wt(3,3)+0.3))
        if k1>0:
            circle(x,220,0,140*k1,CARD,a); circle(x,220,0,140*k1,WHITE,a*0.7,fill=False,lw=4)
            gear(x,220,-20,60*k1,t*3,SOFT,a); sp=sprite("انشغال",44,(255,255,255),b'Bold'); blit(x,sp,220,80,k1,a)
        if k2>0:
            radial(x,-220,0,300*k2,GOLD,0.35*a); circle(x,-220,0,140*k2,GOLD,a)
            check(x,-220,-25,2.2*k2,eo(seg(t,wt(3,5),wt(3,5)+0.5)),NAVY1,a,lw=12)
            sp=sprite("إنجاز",46,(14,32,80),b'Black'); blit(x,sp,-220,80,k2,a)
        if ke>0:
            poly(x,[(-30,-25),(30,-25)],GOLD2,a,fill=False,lw=12,close=False); poly(x,[(-30,25),(30,25)],GOLD2,a,fill=False,lw=12,close=False)
            glowline(x,[(26,-60*ke),(-26,60*ke)],GOLD2,a,10)
def scene_2(x,t):
    s4,s5,s6,s7,s8=[ph(i)['start'] for i in (4,5,6,7,8)]
    out=eio(seg(t,s7-0.15,s7+0.35))
    # schedule board
    kb=back(seg(t,s4-0.35,s4+0.2))
    if kb>0 and out<1:
        a=1-out; x.save(); x.translate(0,lerp(0,-80,out)); x.scale(S_(kb),S_(kb))
        rrect(x,-330,-330,660,620,30,(0,0,0),0.25*a); rrect(x,-330,-340,660,620,30,(0.09,0.17,0.38),a); rrect(x,-330,-340,660,80,30,GOLD,a)
        days=["أحد","اثنين","ثلاثاء","أربعاء","خميس"]
        for i,d in enumerate(days):
            sp=sprite(d,30,(14,32,80),b'Bold'); blit(x,sp,260-i*130,-302,1,a)
        # tasks (s4)
        for j in range(15):
            c,r=j%5,j//5; k=back(seg(t,s4+0.15+j*0.07,s4+0.4+j*0.07))
            if k>0: rrect(x,205-c*130,-240+r*58,110*k,46,10,WHITE,a*0.9); rrect(x,215-c*130,-224+r*58,70*k,12,6,SOFT,a)
        # meetings (s5)
        for j in range(10):
            c,r=j%5,j//5; k=back(seg(t,s5+0.1+j*0.08,s5+0.35+j*0.08))
            if k>0:
                bx,by=205-c*130,-66+r*80
                rrect(x,bx,by,110*k,68,10,GOLD,a*0.9)
                for q in range(3): circle(x,bx+25+q*30,by+26,9*k,NAVY1,a); 
        # flowing work (s6)
        kf=seg(t,s6-0.1,s6+0.3)
        if kf>0:
            for j in range(8):
                u=((t-s6)*0.6+j/8)%1; px=lerp(300,-300,u); al=a*eo(kf)*math.sin(u*math.pi)
                rrect(x,px-50,120,100,110,14,WHITE,al*0.9); rrect(x,px-35,140,70,10,5,SOFT,al); rrect(x,px-35,160,50,10,5,SOFT,al)
        x.restore()
    # magnifier over dots -> one makes the difference
    kq=seg(t,s7-0.2,s7+0.3)
    if kq>0:
        a=eo(kq)
        for r in range(5):
            for c in range(7):
                j=r*7+c; bx,by=-300+c*100,-200+r*100
                special=(r==2 and c==4)
                k=seg(t,wt(8,2)-0.1,wt(8,2)+0.5) if special else 0
                if special and k>0:
                    rr=lerp(18,70,back(k)); radial(x,bx,by,rr*3,GOLD,0.5*a); circle(x,bx,by,rr,GOLD,a)
                    check(x,bx,by-4,1.6*back(k),eo(seg(t,wt(8,3)-0.1,wt(8,3)+0.4)),NAVY1,a,lw=9)
                else: circle(x,bx,by,18,SOFT,a*(0.75-0.45*seg(t,wt(8,2),wt(8,2)+0.5)))
        km=back(seg(t,wt(7,2)-0.15,wt(7,2)+0.3)); kmo=eio(seg(t,wt(8,2)-0.2,wt(8,2)+0.2))
        if km>0 and kmo<1:
            u=seg(t,wt(7,2),wt(8,2)); mx=lerp(-250,100,eio(u))+40*math.sin(t*3); my=lerp(150,0,eio(u))
            magnifier(x,mx,my,km*(1-kmo)*0.9,1)
def scene_3(x,t):
    s9,s10,s11=[ph(i)['start'] for i in (9,10,11)]
    # clock + priority bars
    kc=back(seg(t,s9-0.35,s9+0.2)); kco=eio(seg(t,s10-0.15,s10+0.35))
    if kc>0 and kco<1:
        a=1-kco; cx=170
        radial(x,cx,-30,280,GOLD,0.25*a); circle(x,cx,-30,150*kc,CARD,a)
        calm=eio(seg(t,wt(9,4)-0.2,wt(9,4)+0.6)); clock(x,cx,-30,150*kc,lerp(t*3,0.0,calm)+math.pi/2*0,a)
        circle(x,cx,-30,150*kc,GOLD,a,fill=False,lw=6)
        for i in range(3):
            k=back(seg(t,wt(9,5)-0.1+i*0.2,wt(9,5)+0.3+i*0.2))
            if k>0:
                y=-150+i*120; w=(260-i*60)*k
                rrect(x,-160-w,y-38,w,76,20,GOLD if i==0 else CARD,a); rrect(x,-160-w,y-38,w,76,20,WHITE,a*0.5,fill=False,lw=3)
                num_node(x,-160-w-10,y,34,i+1,a,i==0,k)
    # person asking 3 questions
    kp=back(seg(t,s10-0.25,s10+0.25)); kpo=eio(seg(t,s11-0.15,s11+0.35))
    if kp>0 and kpo<1:
        a=1-kpo; person(x,0,150,2.4*kp,WHITE,a)
        for i in range(3):
            k=back(seg(t,wt(10,4)-0.1+i*0.18,wt(10,4)+0.3+i*0.18))
            if k>0:
                bx,by=(1-i)*200,-170-(40 if i==1 else 0)
                circle(x,bx,by,62*k,GOLD if i==1 else CARD,a); circle(x,bx,by,62*k,WHITE,a*0.7,fill=False,lw=3)
                q=sprite("؟",70,(14,32,80) if i==1 else (246,196,72),b'Black'); blit(x,q,bx,by-4,k,a)
    # sunrise: start your day
    ks=seg(t,s11-0.2,s11+0.3)
    if ks>0:
        a=eo(ks); rise=eo(seg(t,s11-0.1,ph(11)['end']+0.3)); cy=330-260*rise; R=120
        x.save(); x.new_path(); x.rectangle(-600,-600,1200,930); x.clip()
        radial(x,0,cy,420,GOLD,0.5*a)
        for k in range(20):
            ang=k/20*2*math.pi+t*0.2; r1=R+24; r2=R+80+30*math.sin(t*2+k)
            poly(x,[(r1*math.cos(ang),cy+r1*math.sin(ang)),(r2*math.cos(ang),cy+r2*math.sin(ang))],GOLD2,a*0.7,fill=False,lw=5,close=False)
        circle(x,0,cy,R,GOLD2,a); x.restore()
        glowline(x,[(-420,330),(420,330)],GOLD,a,4)
def scene_4(x,t):
    s12,s13,s14=[ph(i)['start'] for i in (12,13,14)]
    cur=0 if t<s13-0.1 else 1 if t<s14-0.1 else 2
    xs=[220,0,-220]; ny=-310; kin=eo(seg(t,s12-0.4,s12))
    pr=[seg(t,s13-0.3,s13+0.2),seg(t,s14-0.3,s14+0.2)]
    for i in range(2):
        poly(x,[(xs[i]-60,ny),(xs[i+1]+60,ny)],SOFT,0.35*kin,fill=False,lw=6,close=False)
        if pr[i]>0: glowline(x,partial([(xs[i]-60,ny),(xs[i+1]+60,ny)],eo(pr[i])),GOLD,kin,6)
    for i in range(3): num_node(x,xs[i],ny,52,i+1,kin,i==cur,1)
    # 1. biggest impact task
    k1=eo(seg(t,s12-0.2,s12+0.3)); k1o=eio(seg(t,s13-0.2,s13+0.25))
    if k1>0 and k1o<1:
        a=k1*(1-k1o); base=300; hs=[150,220,380,180]; hi=2
        for i,h in enumerate(hs):
            g=back(seg(t,wt(12,2)-0.2+i*0.12,wt(12,2)+0.25+i*0.12),1.2)
            win=seg(t,wt(12,5)-0.1,wt(12,5)+0.4)
            if g>0:
                bx=210-i*140; hh=h*g
                c=GOLD if (i==hi and win>0) else CARD
                rrect(x,bx-55,base-hh,110,hh,14,c,a*(1 if i==hi or win==0 else 1-0.5*win)); rrect(x,bx-55,base-hh,110,hh,14,WHITE,a*0.4,fill=False,lw=3)
                if i==hi and win>0:
                    radial(x,bx,base-hh-60,150,GOLD,0.5*a*win); poly(x,star_pts(bx,base-hh-60,46*back(win),rot=-math.pi/2+t*0.4),GOLD2,a)
        glowline(x,[(-400,base),(300,base)],WHITE,a,4)
    # 2. delegate / drop
    k2=eo(seg(t,s13-0.2,s13+0.3)); k2o=eio(seg(t,s14-0.2,s14+0.25))
    if k2>0 and k2o<1:
        a=k2*(1-k2o)
        for i in range(3): card_row(x,-170,-120+i*90,340,a,0.7,gold=(i==0))
        kd=eio(seg(t,wt(13,5)-0.2,wt(13,5)+0.5))   # تفويضه -> to person
        person(x,-300,230,1.3,WHITE,a); 
        card_row(x,lerp(-170,-380,kd),lerp(60,110,kd),lerp(340,160,kd),a*(1-0.3*kd),0.6)
        kt=eio(seg(t,wt(13,7)-0.2,wt(13,7)+0.5))   # الاستغناء -> trash
        bx,by=280,250
        rrect(x,bx-55,by-60,110,130,14,SOFT,a,fill=False,lw=6); rrect(x,bx-70,by-82,140,16,6,SOFT,a)
        for q in (-25,0,25): poly(x,[(bx+q,by-35),(bx+q,by+45)],SOFT,a,fill=False,lw=4,close=False)
        card_row(x,lerp(-170,bx-60,kt),lerp(150,by-110,kt)+60*math.sin(kt*math.pi)*-1,lerp(340,120,kt),a*(1-kt*0.9),0.5)
    # 3. real result: ring to 100 + check
    k3=eo(seg(t,s14-0.2,s14+0.3))
    if k3>0:
        a=k3; kp=eio(seg(t,wt(14,3),ph(14)['end']+0.2)); cy=40
        circle(x,0,cy,190,CARD,a); circle(x,0,cy,190,WHITE,a*0.2,fill=False,lw=26)
        x.new_path(); x.arc(0,cy,190,-math.pi/2,-math.pi/2+2*math.pi*max(0.001,kp)); x.set_line_width(26); x.set_line_cap(cairo.LINE_CAP_ROUND); rgba(GOLD,a,x); x.stroke()
        if kp<1:
            sp=sprite(nums(int(kp*100))+'٪',90,(255,255,255),b'Black'); blit(x,sp,0,cy-4,1,a)
        else:
            kc=seg(t,ph(14)['end']+0.2,ph(14)['end']+0.7); radial(x,0,cy,300,GOLD,0.4*a*kc)
            check(x,0,cy-10,4.2*back(kc),eo(kc),GOLD2,a,lw=16)
def scene_5(x,t):
    s15,s16,s17,s18,s19=[ph(i)['start'] for i in (15,16,17,18,19)]
    kn=back(seg(t,s15-0.3,s15+0.2)); kno=eio(seg(t,s16-0.15,s16+0.3))
    if kn>0 and kno<1:
        a=1-kno; x.save(); x.translate(0,20); x.rotate(-0.05); x.scale(S_(kn),S_(kn))
        rrect(x,-170,-160,340,320,16,(0,0,0),0.25*a); rrect(x,-170,-170,340,320,16,GOLD2,a)
        for i in range(3): rrect(x,-120,-60+i*60,240-i*50,14,7,NAVY2,0.5*a)
        x.restore()
        kp=eo(seg(t,s15,s15+0.35)); py=lerp(-480,-150,kp)
        circle(x,0,py,34,(0.85,0.2,0.25),a); circle(x,-10,py-10,10,WHITE,0.6*a); poly(x,[(0,py+30),(0,py+70)],SOFT,a,fill=False,lw=6,close=False)
    # many checked tasks -> one diamond of value
    kg=seg(t,s16-0.2,s16+0.2); col=eio(seg(t,s17-0.1,s17+0.6))
    if kg>0 and t<s18+0.3:
        n=0
        for r in range(5):
            for c in range(7):
                j=r*7+c; k=back(seg(t,s16+j*0.04,s16+0.25+j*0.04))
                if k<=0: continue
                n+=1; bx,by=-300+c*100,-230+r*100; px,py=lerp(bx,0,col),lerp(by,0,col)
                dimk=eio(seg(t,wt(16,4)-0.1,wt(16,4)+0.4))
                al=(1-0.5*dimk)*(1-col)
                rrect(x,px-30*k,py-30*k,60*k,60*k,12,CARD,al); check(x,px,py,0.9*k,1,WHITE,al,lw=5)
        if col<1: sp=sprite(nums(n),96,(255,255,255),b'Black'); blit(x,sp,0,-340,1-col,1-col)
    kd=seg(t,s17-0.1,s17+0.8); kdo=eio(seg(t,s18-0.1,s18+0.4))
    if kd>0 and kdo<1:
        r=lerp(10,170,back(seg(t,wt(17,3)-0.15,wt(17,3)+0.6),1.2)); diamond(x,0,0,r,1-kdo,t)
        sp=sprite("النتائج",44,(255,255,255),b'Black'); blit(x,sp,0,r+70,seg(t,wt(17,3),wt(17,3)+0.4),1-kdo)
    # busy gears slowing & fading
    kb=seg(t,s18-0.2,s18+0.3); kbo=eio(seg(t,s19-0.1,s19+0.4))
    if kb>0 and kbo<1:
        a=eo(kb)*(1-kbo)*lerp(1,0.4,seg(t,wt(18,1),ph(18)['end']))
        slow=lerp(3,0.2,eio(seg(t,s18,ph(18)['end']+0.2)))
        gear(x,-110,0,110,t*slow,SOFT,a); gear(x,90,-80,70,-t*slow*1.6,SOFT,a,teeth=8); gear(x,100,110,56,-t*slow*2,SOFT,a,teeth=8)
        ban(x,0,0,250,seg(t,wt(18,1)-0.1,wt(18,1)+0.5),a,lw=13)
    # achieve what matters
    kz=seg(t,s19-0.15,s19+0.4)
    if kz>0:
        e=back(kz,1.3); radial(x,0,-10,420*e,GOLD,0.5)
        for i in range(16):
            ang=i*math.pi/8+t*0.25; r1=230*e; r2=r1+70+25*math.sin(t*3+i)
            poly(x,[(r1*math.cos(ang),-10+r1*math.sin(ang)),(r2*math.cos(ang),-10+r2*math.sin(ang))],GOLD2,0.75,fill=False,lw=6,close=False)
        circle(x,0,-10,200*e,GOLD,1); circle(x,0,-10,200*e,WHITE,1,fill=False,lw=6)
        check(x,0,-20,4.6*e,eo(seg(t,wt(19,1)-0.1,wt(19,1)+0.5)),NAVY1,1,lw=18)
        kst=seg(t,wt(19,3)-0.1,wt(19,3)+0.9)
        for i in range(12):
            a2=i*math.pi/6+0.2; r=230+200*eo(kst)
            poly(x,star_pts(r*math.cos(a2),-10+r*math.sin(a2),14*(1-kst*0.7),rot=t),GOLD2,1-kst)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]

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
    w=wave.open(ROOT+'/v6/voice.wav'); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
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
    si=max([i for i in range(5) if t>=scene_span(i)[0]] or [0])
    n=sprite('٠١٢٣٤٥'[si+1],46,(255,255,255),b'Black'); blit(x,n,cx,cy-4,1,a)
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
    for si in range(1,5):
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
    for si in range(5):
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
