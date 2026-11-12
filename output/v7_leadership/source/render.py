import cairo, numpy as np, math, sys, os, subprocess
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from multiprocessing import Pool
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import timeline as TL
ROOT=os.path.dirname(os.path.abspath(__file__))+'/..'
W,H,FPS=1080,1920,30
OFF=TL.OFF; VEND=OFF+TL.VOICE_LEN; OUTRO=VEND+2.9; DUR=OUTRO+4.6; NF=int(DUR*FPS)
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
def leader(x,cx,cy,s,c,a,halo=0):
    if halo>0: radial(x,cx,cy-40*s,170*s,GOLD,0.45*halo*a)
    person(x,cx,cy,s,c,a)
    poly(x,star_pts(cx,cy-112*s,14*s,rot=-math.pi/2),GOLD2 if c!=GOLD else WHITE,a)
def chip(x,txt,cx,cy,k,a,gold=False,icon=None,fs=40):
    if k<=0: return
    sp=sprite(txt,fs,(14,32,80) if gold else (255,255,255),b'Bold'); w=sp['adv']+60+(56 if icon else 0); h=fs+34
    x.save(); x.translate(cx,cy); x.scale(S_(k),S_(k))
    rrect(x,-w/2,-h/2+6,w,h,h/2,(0,0,0),0.25*a); rrect(x,-w/2,-h/2,w,h,h/2,GOLD if gold else CARD,a)
    if not gold: rrect(x,-w/2,-h/2,w,h,h/2,GOLD,a,fill=False,lw=3)
    if icon:
        ic=mi(icon,fs+4,(14,32,80) if gold else (246,196,72)); blit(x,ic,w/2-38,0,1,a); blit(x,sp,-28,-3,1,a)
    else: blit(x,sp,0,-3,1,a)
    x.restore()
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    kl=back(seg(t,s0-0.45,s0+0.15))
    if kl<=0: return
    # balance scale
    tilt=0.18*math.sin((t-s0)*2.2)*(1-seg(t,ph(1)['end'],ph(1)['end']+1.2))
    py=-170
    poly(x,[(0,py),(0,230)],WHITE,kl,fill=False,lw=10,close=False); rrect(x,-110,225,220,26,12,WHITE,kl)
    x.save(); x.translate(0,py); x.rotate(tilt)
    poly(x,[(-280*kl,0),(280*kl,0)],GOLD,1,fill=False,lw=12,close=False); circle(x,0,0,22,GOLD2,1)
    pans=[]
    for sg in (1,-1):
        ex=sg*270*kl; pans.append(ex)
    x.restore()
    for j,sg in enumerate((1,-1)):
        ex=sg*270*kl*math.cos(tilt); ey=py+sg*270*kl*math.sin(tilt)
        poly(x,[(ex,ey),(ex-80,ey+150)],SOFT,kl,fill=False,lw=3,close=False); poly(x,[(ex,ey),(ex+80,ey+150)],SOFT,kl,fill=False,lw=3,close=False)
        x.new_path(); x.arc(ex,ey+150,95,0,math.pi); rgba(GOLD if j==1 else CARD,kl,x); x.fill()
        if j==0:
            k=back(seg(t,wt(0,5)-0.15,wt(0,5)+0.35))
            if k>0:
                leader(x,ex,ey+130,0.85*k,WHITE,1,halo=1); sp=sprite("الحضور",42,(255,255,255),b'Black'); blit(x,sp,ex,ey+250,k,1)
        else:
            k=back(seg(t,wt(1,2)-0.15,wt(1,2)+0.35))
            if k>0:
                for q,dx in enumerate((-50,0,50)): person(x,ex+dx,ey+135-(14 if q==1 else 0),0.6*k,NAVY1 if q!=1 else WHITE,1)
                sp=sprite("التمكين",42,(246,196,72),b'Black'); blit(x,sp,ex,ey+250,k,1)
    kq=back(seg(t,wt(1,2),wt(1,2)+0.4))
    if kq>0:
        for k2 in range(3):
            pr=((t*0.8+k2/3)%1); circle(x,0,-330,50+pr*120,GOLD,(1-pr)*0.5*kq,fill=False,lw=3)
        q=sprite("؟",150,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,0,-335,kq*(1+0.06*math.sin(t*5)),1)
NODES=[(0xe2c7,"ملفات"),(0xe878,"اجتماعات"),(0xe2e6,"مهام"),(0xe158,"رسائل"),(0xf233,"فريق"),(0xe90e,"قرارات")]
def scene_2(x,t):
    s2,s3,s4=[ph(i)['start'] for i in (2,3,4)]
    calm=eio(seg(t,ph(5)['start']-0.6,ph(5)['start']+0.2))
    kl=back(seg(t,s2-0.4,s2+0.2))
    stress=seg(t,s4,ph(4)['end'])
    jx=4*math.sin(t*40)*stress*(1-calm)
    R=330
    for i,(ic,lab) in enumerate(NODES):
        k=back(seg(t,s2+0.4+i*0.3,s2+0.8+i*0.3))
        if k<=0: continue
        ang=-math.pi/2+i*2*math.pi/6+0.25; nx,ny=R*math.cos(ang),R*math.sin(ang)*0.95
        pull=0.06*math.sin(t*6+i)*stress
        if stress>0:
            for w_ in (8,3): poly(x,[(jx,-20),(nx*(1-pull),ny*(1-pull))],GOLD2,(0.25 if w_==8 else 0.9)*k*(1-calm),fill=False,lw=w_,close=False)
        else:
            glowline(x,partial([(0,-20),(nx,ny)],eo(k)),GOLD,k*(1-calm),4)
        circle(x,nx,ny,62*k,CARD,1-calm); circle(x,nx,ny,62*k,WHITE,(1-calm)*0.7,fill=False,lw=3)
        blit(x,mi(ic,56),nx,ny,k,1-calm)
        sp=sprite(lab,30,(200,210,235),b'SemiBold'); blit(x,sp,nx,ny+92,k,1-calm)
    # floating words
    for j,(wd,tt) in enumerate((("متابعة",s3),("تفاصيل",s2+1.0),("قرارات",s3+0.6),("تنفيذ",s4))):
        k=back(seg(t,tt,tt+0.35))
        if k>0:
            px=[-160,170,190,-180][j]; pyy=[-150,-170,130,120][j]+8*math.sin(t*2+j)
            chip(x,wd,px,pyy,k*0.9,1-calm,gold=(j%2==0),fs=34)
    if kl>0: leader(x,jx,40,1.25*kl,GOLD,1,halo=0.6+0.4*stress)
    if stress>0 and calm<1:
        for k2 in range(6):
            ang=k2*1.05+t; poly(x,[(jx+120*math.cos(ang),-10+120*math.sin(ang)),(jx+150*math.cos(ang),-10+150*math.sin(ang))],WHITE,0.6*stress*(1-calm),fill=False,lw=4,close=False)
def scene_3(x,t):
    s5,s6,s7=[ph(i)['start'] for i in (5,6,7)]
    k=eo(seg(t,s5-0.3,s5+0.5))
    radial(x,0,0,520,GOLD,0.18*k)
    # statement card
    kc=back(seg(t,wt(5,1)-0.1,wt(5,1)+0.4)); kco=eio(seg(t,s7-0.2,s7+0.3))
    if kc>0 and kco<1:
        a=1-kco; x.save(); x.translate(0,lerp(-40,-300,eio(seg(t,s6,s6+0.6)))); x.scale(S_(kc),S_(kc))
        sp=sprite("القيادة ليست احتكارًا للإنجاز",56,(14,32,80),b'Black'); w=sp['adv']+90
        rrect(x,-w/2,-62,w,124,40,(0,0,0),0.25*a); rrect(x,-w/2,-70,w,124,40,WHITE,a); rrect(x,-w/2,-70,w,124,40,GOLD,a,fill=False,lw=5)
        blit(x,sp,0,-12,1,a); x.restore()
    # self-done tasks (by himself)
    ks=seg(t,s6-0.1,s6+0.4); kso=eio(seg(t,s7-0.2,s7+0.3))
    if ks>0 and kso<1:
        a=eo(ks)*(1-kso); leader(x,0,120,1.3,WHITE,a,halo=0.5)
        for j in range(5):
            kk=back(seg(t,s6+0.3+j*0.2,s6+0.6+j*0.2)); ang=-math.pi*0.9+j*math.pi*0.2
            circle(x,190*math.cos(ang),40+190*math.sin(ang),26*kk,CARD,a); check(x,190*math.cos(ang),36+190*math.sin(ang),0.8*kk,1,GOLD,a,lw=5)
    # arrows morph from leader to team
    kt=seg(t,s7-0.2,s7+0.3)
    if kt>0:
        a=eo(kt); ly=-200; leader(x,0,ly,1.0,GOLD,a,halo=0.7)
        xs=[-240,0,240]
        for i,tx in enumerate(xs):
            kk=eo(seg(t,wt(7,3)-0.3+i*0.12,wt(7,3)+0.3+i*0.12)); kp=back(seg(t,wt(7,3)-0.1+i*0.12,wt(7,3)+0.3+i*0.12))
            if kk>0:
                pts=[(0,ly+40),(tx*0.5,ly+120),(tx,190)]
                pp=partial([(lerp(0,tx,u),lerp(ly+40,190,u)-40*math.sin(u*math.pi)) for u in np.linspace(0,1,30)],kk)
                glowline(x,pp,GOLD2,a,6)
                if len(pp)>1: arrow_shape(x,pp[-2][0],pp[-2][1],pp[-1][0],pp[-1][1],GOLD2,a,lw=1,head=26)
            if kp>0: radial(x,tx,250,110*kp,GOLD,0.3*a); person(x,tx,300,1.1*kp,WHITE,a)
def scene_4(x,t):
    s8,s9,s10,s11=[ph(i)['start'] for i in (8,9,10,11)]
    # P8: present leader doing a lot alone
    k8=back(seg(t,s8-0.35,s8+0.2)); side=eio(seg(t,s9-0.1,s9+0.6))
    if k8>0:
        lx=lerp(0,330,side); ly=lerp(80,-260,side); ls=lerp(1.4,0.75,side)*k8
        leader(x,lx,ly,ls,lerp(1,0.6,side) and (mix(WHITE,SOFT,side)),1,halo=0.6*(1-side))
        if side<1:
            for j in range(8):
                kk=back(seg(t,wt(8,3)-0.2+j*0.08,wt(8,3)+0.1+j*0.08))
                ang=-math.pi*0.95+j*math.pi*0.27; r=220
                rrect(x,lx+r*math.cos(ang)-26*kk,ly-40+r*math.sin(ang)*0.8-26*kk,52*kk,52*kk,10,GOLD if j%3==0 else CARD,1-side)
    # growth stairs behind team
    kg=seg(t,s9,s9+0.8)
    if kg>0:
        for i in range(3):
            g=back(seg(t,s9+i*0.15,s9+0.5+i*0.15),1.1); h=(80+i*70)*g
            rrect(x,180-i*200-80,330-h,160,h,14,CARD,0.85); rrect(x,180-i*200-80,330-h,160,12,6,GOLD2,0.9*g)
    # team with chips
    xs=[180,-20,-220]; tops=[330-80,330-150,330-220]
    labs=[("ينجز",0xe2e6,wt(9,6) if len(P[9]['words'])>6 else s9+0.8),("يقرر",0xe90e,wt(10,1)),("يتحمل المسؤولية",0xe9e0,wt(11,1))]
    for i in range(3):
        kp=back(seg(t,s9+0.2+i*0.15,s9+0.6+i*0.15))
        if kp>0:
            lit=seg(t,labs[i][2]-0.1,labs[i][2]+0.3)
            person(x,xs[i],tops[i]-10,1.05*kp,mix(WHITE,GOLD,lit),1)
            if lit>0: radial(x,xs[i],tops[i]-60,120,GOLD,0.3*lit)
        kc=back(seg(t,labs[i][2]-0.1,labs[i][2]+0.35))
        if kc>0: chip(x,labs[i][0],xs[i]+(40 if i==2 else 0),tops[i]-200,kc,1,gold=True,icon=labs[i][1],fs=34)
    kb=back(seg(t,wt(11,len(P[11]['words'])-1)-0.1,wt(11,len(P[11]['words'])-1)+0.4))
    if kb>0:
        for i in range(5):
            ang=t*0.7+i*1.26; poly(x,star_pts(-20+300*math.cos(ang),40+150*math.sin(ang),12*kb,rot=t),GOLD2,0.9)
def scene_5(x,t):
    s12,s13,s14,s15=[ph(i)['start'] for i in (12,13,14,15)]
    # stage/spotlight: leader steps aside
    kst=eo(seg(t,s12-0.3,s12+0.3)); step=eio(seg(t,wt(12,3),ph(12)['end']))
    road=eo(seg(t,s13-0.3,s13+0.5))
    if kst>0 and road<1:
        a=kst*(1-road)
        poly(x,[(-60,-420),(60,-420),(230,200),(-230,200)],GOLD,0.12*a); x.save(); x.translate(0,200); x.scale(1,0.22); circle(x,0,0,240,GOLD,0.35*a); x.restore()
        leader(x,lerp(0,-300,step),170,1.3,mix(GOLD,WHITE,step),a,halo=1-step)
        sp=sprite("الواجهة",44,(246,196,72),b'Black'); blit(x,sp,0,-300,1,a*(1-step*0.6))
    if road>0:
        hy=-230; by=340
        poly(x,[(-18,hy),(18,hy),(300,by),(-300,by)],CARD,road*0.9)
        glowline(x,[(-18,hy),(-300,by)],WHITE,road,4); glowline(x,[(18,hy),(300,by)],WHITE,road,4)
        for k in range(8):
            u=((k/8)+t*0.4)%1; u2=min(1,u+0.05); y1=hy+(by-hy)*u**2; y2=hy+(by-hy)*u2**2
            poly(x,[(0,y1),(0,y2)],GOLD,road*min(1,u*4),fill=False,lw=2+10*u,close=False)
        radial(x,0,hy,240,GOLD,0.3*road)
        walk=seg(t,s13,ph(15)['end']+0.5)
        for i,dx in enumerate((-90,0,90)):
            u=0.85-0.5*walk; yy=hy+(by-hy)*u**2; sc=0.35+0.9*u**2
            person(x,dx*u+0*i,yy+5*math.sin(t*8+i*2),sc,WHITE if i!=1 else GOLD2,road)
        leader(x,-330,280,0.9,SOFT,road)
        poly(x,[(-290,250),(-200,180)],GOLD,road,fill=False,lw=6,close=False)
        for j,(lab,ic,tt,cx,cy) in enumerate((("يعرف الطريق",0xe55b,s13,250,-200),("يملك الأدوات",0xf10b,s14,270,-60),("يواصل الأثر",0xef49,s15,250,80))):
            kc=back(seg(t,tt-0.05,tt+0.35)); chip(x,lab,cx,cy,kc,1,gold=(j==2),icon=ic,fs=32)
        kr=seg(t,wt(15,1),wt(15,1)+1.5)
        if kr>0:
            for k in range(3):
                pr=((t-wt(15,1))*0.7+k/3)%1; circle(x,0,hy,30+pr*260,GOLD,(1-pr)*0.6,fill=False,lw=4)
def scene_6(x,t):
    s16,s17,s18=[ph(i)['start'] for i in (16,17,18)]
    k=eo(seg(t,s16-0.35,s16+0.4))
    if k<=0: return
    cy=-40
    x.save(); x.translate(0,cy); x.scale(1,1.35); radial(x,0,0,260,(0.35,0.5,0.85),0.35*k)
    circle(x,0,0,200*k,(0.07,0.14,0.32),1); x.restore()
    x.save(); x.translate(0,cy); x.scale(1,1.35)
    x.new_path(); x.arc(0,0,200*k,0,2*math.pi); x.set_line_width(8); rgba(GOLD,k,x); x.stroke()
    x.new_path(); x.arc(0,0,222*k,math.pi*1.1,math.pi*1.6); x.set_line_width(3); rgba(GOLD2,0.6*k,x); x.stroke()
    x.restore()
    # glare
    poly(x,[(-120,cy-200),(-60,cy-230),(-150,cy+40),(-200,cy+10)],WHITE,0.08*k)
    person(x,0,cy+90,1.6*k,mix(SOFT,WHITE,seg(t,s17,s17+0.6)),1)
    kt=seg(t,wt(18,3)-0.1,wt(18,3)+0.8)
    for i,(dx,dy) in enumerate(((-120,60),(120,60),(-70,140),(70,140))):
        kk=back(seg(kt,i*0.2,i*0.2+0.5))
        if kk>0: person(x,dx,cy+dy+40,0.85*kk,GOLD2 if i<2 else WHITE,0.95)
    if kt>0: radial(x,0,cy,300,GOLD,0.25*kt)
    kq=seg(t,s17-0.1,s17+0.4)
    if kq>0:
        q=sprite("؟",110,(246,196,72),b'Black'); blit(x,q,250,cy-230,eo(kq)*(1+0.05*math.sin(t*3)),eo(kq))
def scene_7(x,t):
    s19,s20=ph(19)['start'],ph(20)['start']
    kin=eo(seg(t,s19-0.35,s19+0.3))
    if kin<=0: return
    fin=eio(seg(t,VEND+0.1,VEND+0.8))
    x.save(); x.translate(0,lerp(0,-140,fin)); s=lerp(1,0.75,fin); x.scale(s,s)
    hal=seg(t,wt(20,2)-0.1,wt(20,2)+0.5)
    leader(x,0,-120,1.2*kin,mix(WHITE,GOLD,hal),1,halo=0.4+0.6*hal)
    for i,dx in enumerate((-230,0,230)):
        kp=back(seg(t,s19+0.1+i*0.15,s19+0.5+i*0.15))
        if kp>0:
            lit=seg(t,wt(20,3)-0.1+i*0.12,wt(20,3)+0.4+i*0.12)
            if lit>0:
                glowline(x,partial([(0,-150),(dx,140)],eo(lit)),GOLD2,0.8,5)
                radial(x,dx,140,140,GOLD,0.35*lit)
            person(x,dx,230,1.45*kp,mix(WHITE,GOLD,lit),1)
    kr=seg(t,wt(20,len(P[20]['words'])-1)-0.1,VEND+3)
    if kr>0:
        for k in range(4):
            pr=((t-wt(20,len(P[20]['words'])-1))*0.45+k/4)%1; circle(x,0,150,80+pr*480,GOLD,(1-pr)*0.5*min(1,kr*3),fill=False,lw=4)
    x.restore()
    # closing statement
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        l1=sprite("القيادة أثر…",74,(255,255,255),b'Black'); l2=sprite("والتمكين امتداد هذا الأثر",66,(246,196,72),b'Black',glow=(245,180,40))
        blit(x,l1,0,300+(1-eo(fin))*30,1,a); blit(x,l2,0,410+(1-eo(seg(t,VEND+0.5,VEND+1.2)))*30,1,a*eo(seg(t,VEND+0.5,VEND+1.2)))
        glowline(x,[(-260*fin,480),(260*fin,480)],GOLD,a*0.7,4)
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
    w=wave.open(ROOT+'/v7/voice.wav'); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
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
