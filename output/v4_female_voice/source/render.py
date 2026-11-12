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
# ---------------- scenes ----------------
def ph(i): return P[i]
def wt(i,j): return P[i]['words'][j]['start']
SC_B=[]  # scene boundaries (video time)
def scene_span(si):
    a=P[TL.SCENES[si][1]]['start']-0.55
    b=P[TL.SCENES[si+1][1]]['start']-0.55 if si+1<len(TL.SCENES) else OUTRO
    return a,b
def scene_A(x,t):
    s0,s1,s2=ph(0)['start'],ph(1)['start'],ph(2)['start']
    # hourglass + ban (لا تنتظر)
    kin=back(seg(t,s0-0.45,s0+0.15)); kout=eio(seg(t,s1-0.15,s1+0.35))
    if kin>0 and kout<1:
        x.save(); x.translate(0,-10); x.rotate(0.06*math.sin(t*2)-kout*1.2); sc=kin*(1-kout); x.scale(sc,sc)
        flow=seg(t,s0-0.2,s1+0.3); lev=1-flow
        poly(x,[(-110,-170),(110,-170)],GOLD,1,fill=False,lw=16,close=False); poly(x,[(-110,170),(110,170)],GOLD,1,fill=False,lw=16,close=False)
        glass=[(-90,-160),(90,-160),(10,-6),(10,6),(90,160),(-90,160),(-10,6),(-10,-6)]
        poly(x,glass,CARD,0.6); 
        ys=-160*lev; hw=lambda y:10+80*abs(y)/160
        if lev>0.02: poly(x,[(-hw(ys)+6,ys),(hw(ys)-6,ys),(0,-4)],GOLD2,0.95)
        b=flow; poly(x,[(-84*min(1,b*1.4),156),(84*min(1,b*1.4),156),(0,156-120*b)],GOLD2,0.95)
        if 0.02<lev: poly(x,[(0,-4),(0,156-120*b)],GOLD2,0.9,fill=False,lw=4,close=False)
        poly(x,glass,WHITE,1,fill=False,lw=7)
        x.restore()
        ban(x,0,-10,240*(1-kout),seg(t,wt(0,1)-0.05,wt(0,1)+0.55),1-kout)
    # road (أن يمنحك الطريق فرصة)
    kin=eo(seg(t,s1-0.25,s1+0.45)); kout=eio(seg(t,s2-0.2,s2+0.3))
    if kin>0 and kout<1:
        a=kin*(1-kout); hy=-170; by=330
        x.save(); x.translate(0,lerp(80,0,kin)); s=lerp(1,1.25,kout); x.scale(s,s)
        poly(x,[(-14,hy),(14,hy),(330,by),(-330,by)],CARD,a*0.9)
        glowline(x,[(-14,hy),(-330,by)],WHITE,a,5,k=kin); glowline(x,[(14,hy),(330,by)],WHITE,a,5,k=kin)
        for k in range(9):
            u=((k/9)+t*0.55)%1; u2=min(1,u+0.05); y1=hy+(by-hy)*u**2; y2=hy+(by-hy)*u2**2
            poly(x,[(0,y1),(0,y2)],GOLD,a*min(1,u*4),fill=False,lw=2+12*u,close=False)
        glowline(x,[(-420,hy),(420,hy)],GOLD,a*0.7,3)
        radial(x,0,hy,260,GOLD,0.25*a)
        x.restore()
    # star opportunity: pops at "فرصة", flies to chart arrow tip
    ks=back(seg(t,wt(1,3)-0.05,wt(1,3)+0.45))
    if ks>0:
        fly=eio(seg(t,s2+0.6,s2+1.3)); cx=lerp(0,215,fly); cy=lerp(-170,-260,fly); r=lerp(62,40,fly)*ks
        endk=1-eio(seg(t,scene_span(0)[1]-0.3,scene_span(0)[1]+0.2))
        radial(x,cx,cy,r*3,GOLD,0.45*endk)
        for i in range(12):
            ang=i*math.pi/6+t*0.8; poly(x,[(cx+r*1.35*math.cos(ang),cy+r*1.35*math.sin(ang)),(cx+r*1.9*math.cos(ang),cy+r*1.9*math.sin(ang))],GOLD2,0.8*endk,fill=False,lw=4,close=False)
        poly(x,star_pts(cx,cy,r,rot=-math.pi/2+t*0.5),GOLD2,endk)
    # bar chart (اصنع فرصتك بنفسك)
    kin=seg(t,s2-0.2,s2+0.3)
    if kin>0:
        a=eo(kin); base=250
        glowline(x,[(-280,base),(280,base)],WHITE,a,4,k=eo(seg(t,s2-0.2,s2+0.4)))
        hs=[110,190,280,370]; xs=[-195,-65,65,195]; tops=[]
        for i,(bx,hh) in enumerate(zip(xs,hs)):
            g=back(seg(t,s2+0.05+i*0.22,s2+0.6+i*0.22),1.2)
            h=hh*g; tops.append((bx,base-h-30))
            if h>1:
                rrect(x,bx-42,base-h,84,h,12,GOLD if i==3 else CARD,a); rrect(x,bx-42,base-h,84,min(h,14),7,GOLD2,a*(1 if i<3 else 0))
                if i<3: rrect(x,bx-42,base-h,84,h,12,WHITE,a*0.35,fill=False,lw=3)
        ka=seg(t,wt(2,1),wt(2,1)+0.9)
        if ka>0:
            pts=[(-260,base-60)]+tops; pp=partial(pts,eo(ka)); glowline(x,pp,GOLD2,a,7)
            if len(pp)>1: arrow_shape(x,pp[-2][0],pp[-2][1],pp[-1][0],pp[-1][1],GOLD2,a,lw=1,head=30)
        kp=back(seg(t,wt(2,2)-0.05,wt(2,2)+0.4))
        if kp>0: radial(x,195,base-370-90,130,GOLD,0.35*kp); person(x,195,base-370-38,1.05*kp,WHITE,a)
def scene_B(x,t):
    s3,s4,s5=ph(3)['start'],ph(4)['start'],ph(5)['start']
    kinds=['circle','square','triangle']; xs=[-230,0,230]
    conv=eio(seg(t,s4-0.1,s4+0.6))
    kout=eio(seg(t,s5-0.2,s5+0.4))
    if kout<1:
        # starting lines
        glowline(x,[(-360,170),(360,170)],SOFT,0.5*eo(seg(t,s3-0.3,s3+0.3))*(1-conv),3)
        m=max(0,(t-wt(3,3)+0.1)*1.3)
        for i in range(3):
            g=back(seg(t,wt(3,min(i+1,3))-0.25,wt(3,min(i+1,3))+0.25))
            if g<=0: continue
            mi=int(m); fr=eio(m-mi) if m>0 else 0
            A=shape_pts(kinds[(i+mi)%3],70); B=shape_pts(kinds[(i+mi+1)%3],70); pts=A*(1-fr)+B*fr
            # converge into arrow
            ARW=shape_pts('arrow',105)
            ang=math.pi*lerp(1,0,eio(seg(t,wt(4,1)+0.55,wt(4,1)+1.1)))   # arrow points down (back) then flips up
            R=np.array([[math.cos(ang),-math.sin(ang)],[math.sin(ang),math.cos(ang)]]); ARW=ARW@R.T
            pts=pts*(1-conv)+ARW*conv
            cx=lerp(xs[i],0,conv); cy=lerp(70,-20,conv)
            pts=pts*g+np.array([cx,cy])
            col=mix(WHITE,GOLD,[0.0,1,0.0][i]) if conv<0.5 else mix(WHITE,GOLD,seg(t,wt(4,1)+0.55,wt(4,1)+1.0))
            poly(x,pts.tolist(),col,(1-kout)*(0.95 if conv>0 else 1))
            if conv<0.9:
                fa=(1-conv)*g*(1-kout); fx=cx; fy=cy-95
                poly(x,[(fx,fy+20),(fx,fy-60)],WHITE,fa,fill=False,lw=5,close=False)
                wv=math.sin(t*6+i)*8; poly(x,[(fx,fy-60),(fx+55,fy-48+wv),(fx,fy-34)],GOLD,fa)
        # cross over the backward arrow
        kc=seg(t,wt(4,1)-0.05,wt(4,1)+0.4); kcx=1-eio(seg(t,wt(4,1)+0.5,wt(4,1)+0.8))
        if kc>0 and kcx>0:
            for (a0,a1) in (((-120,-140),(120,100)),((120,-140),(-120,100))):
                glowline(x,partial([a0,a1],eo(kc)),GOLD,kcx*(1-kout),16)
        ku=seg(t,wt(4,1)+0.9,wt(4,1)+1.4)
        if ku>0:
            for k in range(3):
                yy=-200-((t*160+k*60)%180); poly(x,[(-40,yy+30),(0,yy),(40,yy+30)],GOLD2,ku*(1-kout)*0.8,fill=False,lw=7,close=False)
    # winding foggy path
    kin=seg(t,s5-0.3,s5+0.3)
    if kin>0:
        a=eo(kin)
        pts=[(140*math.sin(u*3.4+0.3)*(1-u*0.45),330-u*600) for u in np.linspace(0,1,120)]
        solid=pts[:55]; rest=pts[54:]
        kd=eo(seg(t,s5-0.1,wt(5,2)))
        glowline(x,solid,GOLD,a,10,k=kd)
        if kd>=1:
            kr=seg(t,wt(5,2),wt(5,4))
            for j,p in enumerate(rest[::5]):
                if j/len(rest[::5])>kr: break
                circle(x,p[0],p[1],6*(1-j/14),WHITE,a*(0.8-0.05*j))
        # fog
        for k in range(5):
            fx=-260+130*k+60*math.sin(t*0.6+k); fy=-170+30*math.sin(t*0.8+k*2)
            radial(x,fx,fy,190,(0.75,0.82,0.95),0.22*a*seg(t,wt(5,3)-0.5,wt(5,3)+0.3))
        # pin walking along
        u=eo(seg(t,s5,wt(5,3)))*54; i=int(min(53,u)); p=solid[i]
        x.save(); x.translate(p[0],p[1]-36); 
        x.new_path(); x.arc(0,-6,22,math.pi*0.85,math.pi*0.15+2*math.pi); x.line_to(0,30); x.close_path(); rgba(WHITE,a,x); x.fill(); circle(x,0,-6,8,NAVY1,a); x.restore()
        kq=back(seg(t,wt(5,3)-0.1,wt(5,3)+0.4))
        if kq>0:
            q=sprite("؟",170,(245,192,62),b'Black',glow=(245,190,60)); blit(x,q,0,-300,kq*(1+0.06*math.sin(t*5)),a)
def door(x,cx,base,w,h,a,openk=0,dim=1):
    top=base-h; r=w/2
    x.new_path(); x.move_to(cx-r,base); x.line_to(cx-r,top+r); x.arc(cx,top+r,r,math.pi,2*math.pi); x.line_to(cx+r,base); x.close_path()
    path=x.copy_path()
    if openk>0:
        x.save(); x.append_path(path); x.clip()
        g=cairo.LinearGradient(cx,top,cx,base); g.add_color_stop_rgba(0,1,0.92,0.7,a); g.add_color_stop_rgba(1,0.96,0.75,0.24,a); x.set_source(g); x.paint_with_alpha(openk)
        x.restore()
    pw=w*(1-0.82*openk)
    x.new_path(); x.move_to(cx-r,base); x.line_to(cx-r,top+r); x.arc(cx,top+r,r,math.pi,2*math.pi); x.line_to(cx+r,base); x.close_path()
    if openk<=0: rgba(CARD,a*dim,x); x.fill()
    else:
        x.save(); x.append_path(path); x.clip(); rrect(x,cx-r,top-10,pw,h+20,0,CARD,a*dim); x.restore()
    x.append_path(path); x.set_line_width(7); rgba(mix(WHITE,GOLD,openk),a*dim,x); x.stroke()
    circle(x,cx-r+pw-18,base-h*0.42,7,GOLD,a*dim*(1-openk))
def scene_C(x,t):
    s6,s7,s8=ph(6)['start'],ph(7)['start'],ph(8)['start']
    xs=[-235,0,235]; base=250; dh=300; dw=170
    kkey=eo(seg(t,s8+0.05,wt(8,2)+0.15))
    openk=eio(seg(t,wt(8,2)+0.2,wt(8,2)+0.9))
    if openk>0:
        for i in range(14):
            ang=-math.pi/2+(i-6.5)*0.17; L=lerp(150,520,openk)
            poly(x,[(0,base-120),(L*math.cos(ang)*1.1,base-120+L*math.sin(ang))],GOLD2,0.18*openk,fill=False,lw=10,close=False)
        radial(x,0,base-140,360*openk,GOLD,0.35*openk)
    for i,cx in enumerate(xs):
        g=back(seg(t,s6-0.3+i*0.14,s6+0.25+i*0.14),1.3)
        if g<=0: continue
        x.save(); x.translate(cx,base); x.scale(g,g); x.translate(-cx,-base)
        dim=1-0.55*openk if i!=1 else 1
        door(x,cx,base,dw,dh,1,openk if i==1 else 0,dim)
        x.restore()
        # locks drop
        kl=seg(t,s7-0.25+i*0.08,s7+0.25+i*0.08)
        if kl>0:
            yy=lerp(-520,base-140,eo(kl)) + (8*math.sin(kl*12)*(1-kl) if kl<1 else 0)
            fall=eio(seg(t,wt(8,2)+0.25,wt(8,2)+0.8)) if i==1 else 0
            ok=eo(seg(t,wt(8,2)+0.05,wt(8,2)+0.3)) if i==1 else 0
            lock(x,cx,yy+fall*260,1.15,(1-fall)*(dim if i!=1 else 1),ok)
    if 0<kkey and openk<1:
        kx=lerp(480,40,kkey); ky=lerp(-420,base-140,kkey); key(x,kx,ky,1.0,lerp(-1.2,math.pi,kkey),1-openk)
    # attempts (المحاولة)
    ka=seg(t,wt(8,4)-0.1,wt(8,4)+0.8)
    if ka>0:
        for i in range(3):
            k=back(seg(ka,i*0.25,i*0.25+0.3)); circle(x,(i-1)*70,-200,24*k,GOLD if i==2 else WHITE,1)
            circle(x,(i-1)*70,-200,30*k,GOLD,0.6,fill=False,lw=3)
        kc=seg(t,wt(8,4)+0.75,wt(8,4)+1.2)
        if kc>0:
            circle(x,0,-310,52*back(kc),GOLD,1); glowline(x,partial([(-24,-310),(-6,-292),(26,-330)],eo(kc)),NAVY1,1,9)
def cloud(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(s,s)
    x.new_path()
    for (px,py,r) in ((-110,10,70),(-40,-45,88),(55,-38,80),(125,12,62),(0,20,80)): x.arc(px,py,r,0,2*math.pi); x.new_sub_path()
    rgba(WHITE,a,x); x.fill()
    rrect(x,-175,10,350,72,36,WHITE,a)
    circle(x,-150,130,18,WHITE,a); circle(x,-185,175,10,WHITE,a)
    x.restore()
def scene_D(x,t):
    s9,s10,s11,s12=[ph(i)['start'] for i in (9,10,11,12)]
    up=eio(seg(t,s10-0.1,s10+0.5))
    kc=back(seg(t,s9-0.3,s9+0.35))
    if kc>0:
        cx,cy,s=lerp(0,-170,up),lerp(-20,-300,up),kc*lerp(1.15,0.55,up)
        radial(x,cx,cy,300*s,GOLD,0.18)
        cloud(x,cx,cy,s,1)
        for i,(px,py,r) in enumerate(((-60,-10,26),(30,-30,20),(90,20,16),(-120,30,14),(10,40,18))):
            k=back(seg(t,wt(9,1)+i*0.1,wt(9,1)+0.3+i*0.1))
            rr=r*k*(0.85+0.15*math.sin(t*5+i))*s; poly(x,star_pts(cx+px*s,cy+py*s,rr,rot=t*0.6+i),GOLD,1)
    # target
    kt=back(seg(t,s10-0.2,s10+0.35))
    tx,ty=60,90
    if kt>0:
        for i,(r,c) in enumerate(((200,WHITE),(150,CARD),(100,WHITE),(52,CARD))):
            circle(x,tx,ty,r*kt,c,1)
        circle(x,tx,ty,22*kt,GOLD,1)
    # missed arrow (لا تتحقق)
    km=seg(t,wt(10,1)-0.1,wt(10,1)+0.45); kmo=1-eio(seg(t,s11,s11+0.4))
    if km>0 and kmo>0:
        e=eo(km); x0,y0=-460,-60; x1,y1=lerp(x0,tx+250,e),lerp(y0,ty-230,e)
        traj=[(lerp(x0,tx+250,u),lerp(y0,ty-230,u)-60*math.sin(u*math.pi)) for u in np.linspace(0,e,30)]
        for j in range(0,len(traj)-1,3): poly(x,[traj[j],traj[j+1]],SOFT,0.6*kmo,fill=False,lw=4,close=False)
        if len(traj)>1: arrow_shape(x,traj[-2][0],traj[-2][1],traj[-1][0],traj[-1][1],WHITE,kmo,lw=6)
        if km>=1:
            kx=seg(t,wt(10,1)+0.45,wt(10,1)+0.8); d=36
            glowline(x,partial([(tx+250-d,ty-230-d),(tx+250+d,ty-230+d)],kx),GOLD,kmo,10); glowline(x,partial([(tx+250+d,ty-230-d),(tx+250-d,ty-230+d)],kx),GOLD,kmo,10)
    # heart (يؤمن بها)
    kh=back(seg(t,wt(11,2)-0.15,wt(11,2)+0.3)); kho=1-eio(seg(t,ph(12)['end']-0.6,ph(12)['end']-0.3))
    if kh>0 and kho>0:
        beat=1+0.12*max(0,math.sin(t*7))**8
        for k in range(3):
            pr=((t*0.9+k/3)%1); circle(x,tx,ty,60+pr*220,GOLD,(1-pr)*0.5*kh*kho,fill=False,lw=4)
        heart(x,tx,ty+4,1.25*kh*beat*kho,GOLD2,1)
    # gears + progress (ويعمل لأجلها)
    kg=back(seg(t,s12-0.2,s12+0.3))
    if kg>0:
        gear(x,-290,300,72*kg,t*1.6,GOLD,1); gear(x,-178,348,46*kg,-t*2.5+0.2,WHITE,1,teeth=8)
        kp=eio(seg(t,s12+0.1,ph(12)['end']-0.25))
        rrect(x,-100,320,380,26,13,CARD,kg); rrect(x,-100,320,380*kp,26,13,GOLD,kg)
        pct=sprite(''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(int(kp*100)))+'٪',44,(255,255,255),b'Bold'); blit(x,pct,90,280,kg,kg)
    # final hit
    khit=seg(t,ph(12)['end']-0.6,ph(12)['end']-0.25)
    if khit>0:
        e=eo(khit); x0,y0=-460,-200; ax,ay=lerp(x0,tx,e),lerp(y0,ty,e); ang=math.atan2(ty-y0,tx-x0)
        arrow_shape(x,ax-170*math.cos(ang),ay-170*math.sin(ang),ax,ay,GOLD2,1,lw=8,head=34)
        if khit>=1:
            kb=seg(t,ph(12)['end']-0.25,ph(12)['end']+0.4)
            for i in range(14):
                a2=i*2*math.pi/14; r0=40+150*eo(kb); r1=r0+50*(1-kb)
                poly(x,[(tx+r0*math.cos(a2),ty+r0*math.sin(a2)),(tx+r1*math.cos(a2),ty+r1*math.sin(a2))],GOLD2,1-kb,fill=False,lw=6,close=False)
def scene_E(x,t):
    s13,s14,s15,s16=[ph(i)['start'] for i in (13,14,15,16)]
    clear=eio(seg(t,s16-0.2,s16+0.3))
    xs=[-300,-150,0,150,300]
    lead=eo(seg(t,wt(13,1)-0.1,wt(13,1)+0.4))
    if clear<1:
        # zoom-out feel on "فالحياة تتسع"
        z=lerp(1,0.82,eio(seg(t,s15-0.1,s15+0.6)))
        x.save(); x.scale(z,z)
        kr=seg(t,s15-0.1,s15+1.2)
        if kr>0:
            for k in range(5):
                pr=cl(kr*1.4-k*0.12); circle(x,0,60,80+pr*520,GOLD if k%2==0 else WHITE,(1-pr)*0.7*(1-clear),fill=False,lw=5)
            ke=back(seg(t,wt(15,1)-0.1,wt(15,1)+0.35))
            for k in range(4):
                ang=math.pi/4+k*math.pi/2; r0=330*ke; arrow_shape(x,r0*0.8*math.cos(ang),60+r0*0.8*math.sin(ang),r0*1.05*math.cos(ang),60+r0*1.05*math.sin(ang),GOLD2,ke*(1-clear),lw=8,head=28)
        if lead>0: radial(x,0,-60,260,GOLD,0.35*lead*(1-clear)); 
        if lead>0: poly(x,[(-40,-420),(40,-420),(170,180),(-170,180)],GOLD,0.08*lead*(1-clear))
        for i,px in enumerate(xs):
            g=back(seg(t,s13-0.3+abs(i-2)*0.08,s13+0.2+abs(i-2)*0.08))
            if g<=0: continue
            if i==2:
                person(x,px,lerp(160,90,lead),lerp(1,1.45,lead)*g,mix(WHITE,GOLD,lead),1-clear)
            else:
                fade=lerp(0.85,0.3,eio(seg(t,s14-0.1,s14+0.4)))
                person(x,px,160,g,SOFT,fade*(1-clear))
        kf=back(seg(t,wt(13,2)-0.05,wt(13,2)+0.4))
        if kf>0:
            fx,fy=36,-80; poly(x,[(fx,fy+30),(fx,fy-110*kf)],WHITE,1-clear,fill=False,lw=6,close=False)
            wv=math.sin(t*6)*10; poly(x,[(fx,fy-110*kf),(fx+80*kf,fy-92*kf+wv),(fx,fy-70*kf)],GOLD,1-clear)
        # waiting clock crossed (لا منتظر الفرصة)
        kk=back(seg(t,s14-0.15,s14+0.3)); kko=1-eio(seg(t,s15-0.2,s15+0.2))
        if kk>0 and kko>0:
            clock(x,250,-250,85*kk,t*3,kko); ban(x,250,-250,120*kk,seg(t,wt(14,1)-0.1,wt(14,1)+0.5),kko,lw=11)
        x.restore()
    # stairs + climber (لمن يملك شجاعة السعي)
    kin=seg(t,s16-0.25,s16+0.25)
    if kin>0:
        steps=5; sw=120; sh=95; x0=-300; y0=330
        tops=[]
        for i in range(steps):
            g=back(seg(t,s16-0.1+i*0.16,s16+0.35+i*0.16),1.2)
            bx=x0+i*sw; by=y0-(i+1)*sh; tops.append((bx+sw/2,by))
            if g>0:
                hh=(i+1)*sh*g
                rrect(x,bx+3,y0-hh,sw-6,hh,10,GOLD if i==steps-1 else CARD,1); rrect(x,bx+3,y0-hh,sw-6,12,6,GOLD2,1)
        # dotted rising line
        kl=seg(t,wt(16,1),wt(16,3)); pts=[(x0-20,y0-30)]+[(tx,ty-50) for tx,ty in tops]
        pp=partial(pts,kl)
        for j in range(len(pp)-1): poly(x,[pp[j],pp[j+1]],GOLD2,0.6,fill=False,lw=3,close=False)
        # climber
        prog=eio(seg(t,wt(16,1)-0.1,wt(16,3)+0.25))*(steps-1); i=int(prog); f=prog-i
        p0=tops[i]; p1=tops[min(i+1,steps-1)]; hop=math.sin(f*math.pi)*40
        px,py=lerp(p0[0],p1[0],f),lerp(p0[1],p1[1],f)-hop
        person(x,px,py-4,0.95*eo(kin),GOLD if prog>=steps-1.01 else WHITE,1)
        kf=back(seg(t,wt(16,3),wt(16,3)+0.5))
        if kf>0:
            fx,fy=tops[-1][0]+40,tops[-1][1]; poly(x,[(fx,fy),(fx,fy-150*kf)],WHITE,1,fill=False,lw=6,close=False)
            wv=math.sin(t*7)*10; poly(x,[(fx,fy-150*kf),(fx+95*kf,fy-128*kf+wv),(fx,fy-104*kf)],GOLD,1)
            for k in range(10):
                a2=k*0.63+0.2; r=60+120*eo(seg(t,wt(16,3)+0.2,wt(16,3)+0.9))
                circle(x,fx+r*math.cos(a2),fy-120+r*math.sin(a2),5*(1-seg(t,wt(16,3)+0.6,wt(16,3)+1.3)),GOLD2,1)
SCENE_FN=[scene_A,scene_B,scene_C,scene_D,scene_E]
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
    w=wave.open(ROOT+'/v4/takeA.wav'); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
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
