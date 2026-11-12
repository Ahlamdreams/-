import numpy as np, wave, sys, math, subprocess, os
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from multiprocessing import Pool
S=os.path.dirname(os.path.abspath(__file__))+'/..'
W,H=1080,1920; W2,H2=540,960; FPS=30
OFF=1.5; VOICE_END=31.35; DUR=OFF+VOICE_END+4.2
NF=int(DUR*FPS)
FONT=S+'/fonts/Cairo.ttf'; AMIRI=S+'/fonts/Amiri-Bold.ttf'
GOLD=(246,200,110); WHITE=(250,246,238)
# ---------- script (times = original voice time) ----------
scenes=[
 [("لا تنتظر",0.96,2.43),("أن يمنحك الطريق فرصة",2.58,3.42),("اصنع فرصتك بنفسك",3.93,5.61)],
 [("لا تخشَ البدايات المختلفة",6.27,8.13),("ولا تتراجع",8.52,9.57),("لأن الطريق لم تتضح ملامحه بعد",9.87,12.69)],
 [("فبعض الأبواب",13.32,14.28),("لا تُفتح",14.55,14.97),("إلا لمن يجرؤ على المحاولة",15.15,17.52)],
 [("وبعض الأحلام",17.97,19.02),("لا تتحقق",19.29,20.22),("إلا لمن يؤمن بها",20.37,21.9),("ويعمل لأجلها",22.2,23.58)],
 [("كن صاحب المبادرة",24.21,25.62),("لا منتظر الفرصة",25.95,27.21),("فالحياة تتسع",27.81,28.44),("لمن يملك شجاعة السعي",28.77,31.3)],
]
HIGHLIGHT={"فرصتك","بنفسك","البدايات","الأبواب","يجرؤ","الأحلام","صاحب","المبادرة","شجاعة","السعي","تتسع"}
SC_START=[OFF+s[0][1]-0.25 for s in scenes]
SC_END=[SC_START[i+1]-0.15 for i in range(len(scenes)-1)]+[OFF+VOICE_END+0.6]
OUTRO=OFF+VOICE_END+0.6
def font(sz,path=FONT,var=b'Bold'):
    f=ImageFont.truetype(path,sz,layout_engine=ImageFont.Layout.RAQM)
    if path==FONT:
        try: f.set_variation_by_name(var)
        except Exception: f.set_variation_by_name(var.decode())
    return f
def ease(x): x=min(1,max(0,x)); return 1-(1-x)**3
def smooth(x): x=min(1,max(0,x)); return x*x*(3-2*x)
# ---------- word sprites ----------
def word_sprite(txt,f,col):
    bb=f.getbbox(txt,direction='rtl',language='ar'); pad=40
    w=bb[2]-bb[0]+2*pad; h=bb[3]-bb[1]+2*pad
    im=Image.new('L',(w,h),0); ImageDraw.Draw(im).text((pad-bb[0],pad-bb[1]),txt,font=f,fill=255,direction='rtl',language='ar')
    glow=im.filter(ImageFilter.GaussianBlur(14)); soft=im.filter(ImageFilter.GaussianBlur(6))
    return dict(a=np.asarray(im,np.float32)/255,g=np.asarray(glow,np.float32)/255,s=np.asarray(soft,np.float32)/255,col=np.array(col,np.float32),w=w,h=h,pad=pad,bb=bb)
def build_layout():
    f=font(96); fs=f.getlength(' ')
    blocks=[]
    for si,sc in enumerate(scenes):
        lines=[]
        for (txt,t0,t1) in sc:
            words=txt.split(); lens=[len(w)+1 for w in words]; tot=sum(lens); acc=0
            # wrap if too wide
            ws=[]
            for w,l in zip(words,lens):
                ws.append(dict(t=w,start=OFF+t0+(t1-t0)*acc/tot-0.12,sp=word_sprite(w,f,GOLD if any(h in w for h in HIGHLIGHT) else WHITE))); acc+=l
            # break into lines of max width 940
            cur=[];cw=0
            for wd in ws:
                ww=f.getlength(wd['t'],direction='rtl',language='ar')
                if cur and cw+fs+ww>940: lines.append(cur); cur=[]; cw=0
                wd['adv']=ww; cw+= (fs if cur else 0)+ww; cur.append(wd)
            lines.append(cur)
        lh=150; total=lh*len(lines); y0=900-total/2
        for li,ln in enumerate(lines):
            lw=sum(w['adv'] for w in ln)+fs*(len(ln)-1); x=W/2+lw/2
            for w in ln:  # RTL: place from right
                x-=w['adv']; w['x']=x; w['y']=y0+li*lh; x-=fs
        blocks.append(dict(lines=lines,start=SC_START[si],end=SC_END[si]))
    return blocks
# ---------- audio features ----------
def voice_features():
    w=wave.open(S+'/voice16.wav'); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
    sr=16000; hop=sr//FPS; nb=24; feats=np.zeros((NF,nb)); rms=np.zeros(NF)
    edges=np.geomspace(90,5000,nb+1); fr=np.fft.rfftfreq(1024,1/sr); win=np.hanning(1024)
    for i in range(NF):
        c=int((i/FPS-OFF)*sr)
        if c<0 or c+1024>len(a): continue
        x=a[c:c+1024]; rms[i]=np.sqrt(np.mean(x**2)); sp=np.abs(np.fft.rfft(x*win))
        feats[i]=[sp[(fr>=edges[k])&(fr<edges[k+1])].mean() for k in range(nb)]
    feats=np.log1p(feats*4); feats/=np.percentile(feats[feats>0],98)+1e-6
    rms/=np.percentile(rms,98)+1e-6
    # smooth
    for arr in (feats,rms):
        for i in range(1,NF): arr[i]=np.maximum(arr[i],arr[i-1]*0.82)
    return np.clip(feats,0,1.2),np.clip(rms,0,1.2)
# ---------- palette keyframes (top,bottom) ----------
PAL=[(0,(6,8,24),(18,16,48)),(SC_START[0],(8,12,36),(26,38,84)),(SC_START[1],(8,20,34),(24,64,74)),
     (SC_START[2],(22,10,34),(78,38,58)),(SC_START[3],(6,8,30),(38,28,86)),(SC_START[4],(26,12,30),(150,70,42)),(OUTRO,(30,14,26),(170,86,44)),(DUR,(30,14,26),(170,86,44))]
def palette(t):
    for i in range(len(PAL)-1):
        if PAL[i][0]<=t<PAL[i+1][0]:
            k=smooth(min(1,(t-PAL[i][0])/1.2)); a,b=PAL[i-1] if i>0 else PAL[0],PAL[i]
            top=np.array(a[1])*(1-k)+np.array(b[1])*k; bot=np.array(a[2])*(1-k)+np.array(b[2])*k
            return top,bot
    return np.array(PAL[-1][1],float),np.array(PAL[-1][2],float)
def scene_weight(si,t,fi=0.9,fo=0.9):
    s=SC_START[si]-0.3; e=SC_END[si]+(3.5 if si==4 else 0.3)
    return smooth((t-s)/fi)*smooth((e-t)/fo)
# ---------- globals per worker ----------
G={}
def init():
    rng=np.random.default_rng(3)
    G['blocks']=build_layout(); G['feats'],G['rms']=voice_features()
    yy,xx=np.mgrid[0:H,0:W]; r=np.sqrt(((xx-W/2)/(W*0.75))**2+((yy-H/2)/(H*0.62))**2)
    G['vig']=np.clip(1.15-r**2*0.75,0.25,1)[...,None].astype(np.float32)
    G['grain']=[(rng.standard_normal((H//2,W//2)).astype(np.float32)) for _ in range(6)]
    G['parts']=dict(x=rng.uniform(0,W2,160),y=rng.uniform(0,H2,160),v=rng.uniform(6,22,160),s=rng.uniform(0.6,2.2,160),ph=rng.uniform(0,6.28,160),dr=rng.uniform(-6,6,160))
    G['stars']=np.c_[rng.uniform(20,520,140),rng.uniform(20,620,140),rng.uniform(0.5,1.8,140),rng.uniform(0,6.28,140)]
    G['const']=[(90,190),(170,120),(250,170),(330,100),(420,150),(470,250),(380,300),(300,260)]
    G['title']=word_sprite("خطوتك الأولى",font(120,AMIRI),GOLD)
    G['final1']=word_sprite("اصنع فرصتك بنفسك",font(118,AMIRI),GOLD)
    G['final2']=word_sprite("كن صاحب المبادرة",font(64,FONT,b'SemiBold'),WHITE)
def blobs(t,top,bot):
    h,w=240,135; yy,xx=np.mgrid[0:h,0:w]/np.array([h,w])[:,None,None]
    base=top[None,None,:]*(1-yy[...,None])+bot[None,None,:]*yy[...,None]
    acc=np.zeros((h,w,1))
    for k,(cx,cy,sp,rad,amp) in enumerate([(0.3,0.3,0.11,0.35,0.5),(0.75,0.6,0.08,0.4,0.45),(0.5,0.85,0.06,0.3,0.6)]):
        x=cx+0.18*math.sin(t*sp*2+k*2); y=cy+0.1*math.cos(t*sp*1.7+k)
        acc+=amp*np.exp(-(((xx-x)/rad)**2+((yy-y)/(rad*0.56))**2))[...,None]
    col=base*(1+0.9*acc)+np.array([20,14,30])*acc
    return np.asarray(Image.fromarray(np.clip(col,0,255).astype(np.uint8)).resize((W,H),Image.BILINEAR),np.float32)
def motif(t,fi):
    L=Image.new('RGB',(W2,H2),(0,0,0)); d=ImageDraw.Draw(L); g=lambda a:tuple(int(c*a) for c in GOLD)
    # intro ornament
    wi=smooth(t/1.2)*smooth((SC_START[0]+0.3-t)/0.6)
    if wi>0:
        ln=int(170*ease(t/1.4)); d.line([(270-ln,610),(270+ln,610)],fill=g(wi),width=2)
        r=9*ease((t-0.4)/0.6); 
        if r>0: d.polygon([(270,610-r),(270+r,610),(270,610+r),(270-r,610)],outline=g(wi),width=2)
    # S1 road
    w=scene_weight(0,t)
    if w>0:
        vp=(270,640); pr=ease((t-SC_START[0]+0.3)/1.5)
        for bx in (-160,700):
            ex=vp[0]+(bx-vp[0])*pr; ey=vp[1]+(H2-vp[1])*pr; d.line([vp,(ex,ey)],fill=g(w*0.9),width=2)
        for k in range(10):
            s=((k/10)+t*0.18)%1; s2=min(1,s+0.035); y1=vp[1]+(H2-vp[1])*s**2; y2=vp[1]+(H2-vp[1])*s2**2
            d.line([(270,y1),(270,y2)],fill=g(w*min(1,s*3)),width=max(1,int(1+5*s)))
        for k in range(1,7):
            gx=vp[0]+(-160-vp[0])*1.0 + k*0; 
        for k in range(6):
            yy=vp[1]+(H2-vp[1])*(((k/6)+t*0.18)%1)**2; d.line([(0,yy),(W2,yy)],fill=g(w*0.12),width=1)
        d.line([(0,vp[1]),(W2,vp[1])],fill=g(w*0.5),width=1)
    # S2 winding path of dots
    w=scene_weight(1,t)
    if w>0:
        rev=ease((t-SC_START[1]+0.2)/(SC_END[1]-SC_START[1]))
        for k in range(60):
            u=k/60
            if u>rev: break
            x=270+110*math.sin(5*u+0.6)*(1-u*0.7); y=950-u*430; r=4.5*(1-u)+1
            a=w*(0.35+0.65*(1-u))*(0.6+0.4*math.sin(t*3-k*0.4))
            d.ellipse([x-r,y-r,x+r,y+r],fill=g(a))
        d.line([(0,520),(W2,520)],fill=g(w*0.25),width=1)
    # S3a arch door opening
    w=scene_weight(2,t)
    if w>0:
        cx,top,bot,hw=270,170,800,175; op=ease((t-SC_START[2]-0.6)/2.8)
        pts=[]
        for i in range(31):
            a=i/30; # pointed arch: two circular arcs
            if a<0.5: th=math.pi*(1-a*2)*0.5; x=cx-hw+hw*1.25*(1-math.cos(th*1.0))*0.8; 
        # simple pointed arch polyline
        arch=[(cx-hw,bot),(cx-hw,330)]
        for i in range(1,16):
            a=i/15; arch.append((cx-hw+hw*math.sin(a*math.pi/2), 330-(330-top)*(1-math.cos(a*math.pi/2))**0.9 if False else 330-(330-top)*math.sin(a*math.pi/2)**1.6))
        right=[(2*cx-x,y) for (x,y) in arch[::-1]]
        outline=arch+right
        # light inside (opening)
        if op>0:
            lm=Image.new('L',(W2,H2),0); ImageDraw.Draw(lm).polygon(outline,fill=int(255*op*w)); lm=lm.filter(ImageFilter.GaussianBlur(25))
            lit=Image.new('RGB',(W2,H2),(255,190,120)); L.paste(lit,(0,0),lm.point(lambda v:int(v*0.32)))
            # rays
            for k in range(-6,7):
                ang=k*0.13; ln=520*op
                d.line([(cx,bot),(cx+ln*math.sin(ang)*1.6,bot+ln*0.25)],fill=g(w*op*0.3),width=2)
        d.line(outline,fill=g(w),width=3)
        # door panels shrinking (perspective open)
        pw=hw*(1-op*0.88)
        for sgn in (-1,1):
            x0=cx+sgn*hw; x1=x0-sgn*pw
            d.polygon([(x0,bot),(x0,330),(x1,330+40*op),(x1,bot-20*op)],outline=g(w*0.8),fill=(int(40*(1-op)),int(26*(1-op)),int(20*(1-op))))
        d.line([(60,bot),(480,bot)],fill=g(w*0.6),width=2)
    # S3b stars & constellation
    w=scene_weight(3,t)
    if w>0:
        for (x,y,s,ph) in G['stars']:
            a=w*(0.35+0.65*(0.5+0.5*math.sin(t*2.2+ph))); r=s
            d.ellipse([x-r,y-r,x+r,y+r],fill=(int(255*a),int(240*a),int(210*a)))
        c=G['const']; pr=(t-SC_START[3]-0.3)/(SC_END[3]-SC_START[3]-1.2)*(len(c)-1)
        for i in range(len(c)-1):
            k=min(1,max(0,pr-i))
            if k<=0: break
            x=c[i][0]+(c[i+1][0]-c[i][0])*k; y=c[i][1]+(c[i+1][1]-c[i][1])*k
            d.line([c[i],(x,y)],fill=g(w*0.8),width=2)
        for i,(x,y) in enumerate(c):
            if pr>=i-0.05:
                r=5+2*math.sin(t*4+i); d.ellipse([x-r,y-r,x+r,y+r],fill=g(w))
    # S4 + outro sunrise
    w=scene_weight(4,t,fo=0.01)
    if w>0:
        rise=ease((t-SC_START[4]+0.3)/3.0); cy=1040-330*rise; R=95
        for k in range(24):
            ang=k/24*2*math.pi+t*0.15; r1=R+25; r2=R+70+40*math.sin(t*2+k)
            d.line([(270+r1*math.cos(ang),cy+r1*math.sin(ang)),(270+r2*math.cos(ang),cy+r2*math.sin(ang))],fill=g(w*0.5),width=2)
        d.ellipse([270-R,cy-R,270+R,cy+R],fill=(int(255*w),int(196*w),int(120*w)))
        for k in range(4):
            ph=((t*0.35)+k/4)%1; rr=R+40+ph*420
            d.ellipse([270-rr,cy-rr,270+rr,cy+rr],outline=g(w*(1-ph)*0.6),width=2)
        d.rectangle([0,712,W2,H2],fill=(0,0,0))  # horizon mask below
        d.line([(0,712),(W2,712)],fill=g(w),width=2)
        for k in range(1,8):
            yy=712+(H2-712)*((k/8+t*0.1)%1)**2; d.line([(0,yy),(W2,yy)],fill=g(w*0.15),width=1)
    # particles
    p=G['parts']; py=(p['y']-p['v']*t)%H2; px=(p['x']+p['dr']*math.sin(t*0.5))%W2
    for x,y,s,ph in zip(px,py,p['s'],p['ph']):
        a=0.35+0.45*math.sin(t*2+ph); d.ellipse([x-s,y-s,x+s,y+s],fill=g(max(0,a)))
    # audio bars
    f=G['feats'][fi]; va=smooth((t-OFF)/1.0)*smooth((OFF+VOICE_END+0.5-t)/0.8)
    if va>0:
        n=len(f)
        for k in range(n):
            hgt=2+48*f[k]
            for sgn in (-1,1):
                x=270+sgn*(6+k*8)
                d.line([(x,905-hgt/2),(x,905+hgt/2)],fill=g(va*(0.75-0.4*k/n)),width=3)
    arr=np.asarray(L,np.float32); glow=np.asarray(L.filter(ImageFilter.GaussianBlur(9)),np.float32)
    out=arr+glow*1.6
    return np.asarray(Image.fromarray(np.clip(out,0,255).astype(np.uint8)).resize((W,H),Image.BILINEAR),np.float32)
def put(img,sp,x,y,alpha,blur=0.0,scale_col=1.0):
    # composite sprite (a,g,s) at top-left x,y with alpha, additive glow
    x=int(round(x-sp['pad']+sp['bb'][0])); y=int(round(y-sp['pad']+sp['bb'][1]))
    h,w=sp['a'].shape; x0,y0=max(0,x),max(0,y); x1,y1=min(W,x+w),min(H,y+h)
    if x1<=x0 or y1<=y0 or alpha<=0: return
    sl=(slice(y0-y,y1-y),slice(x0-x,x1-x))
    a=(sp['a'][sl]*(1-blur)+sp['s'][sl]*blur)*alpha; gl=sp['g'][sl]*alpha
    reg=img[y0:y1,x0:x1]
    reg*=(1-np.clip(sp['g'][sl]*2.2,0,1)*alpha*0.55)[...,None]
    reg+=gl[...,None]*sp['col']*(0.55+0.9*blur)
    reg[:]=reg*(1-a[...,None])+sp['col']*scale_col*a[...,None]
def frame(fi):
    t=fi/FPS
    top,bot=palette(t)
    img=blobs(t,top,bot)
    img+=motif(t,fi)
    # shadow band behind text for readability
    for b in G['blocks']:
        bw=smooth((t-b['start'])/0.5)*smooth((b['end']-t)/0.4)
        if bw<=0: continue
        out=1-smooth((b['end']-t)/0.4)
        for ln in b['lines']:
            for w in ln:
                p=(t-w['start'])/0.55
                if p<=0: continue
                e=ease(p); al=e*(1-out)
                put(img,w['sp'],w['x'],w['y']+(1-e)*46-out*30,al,blur=max(1-e*1.6,0,out))
    # intro title
    ti=smooth((t-0.3)/0.7)*smooth((SC_START[0]-t)/0.5)
    if ti>0:
        sp=G['title']; put(img,sp,W/2-(sp['bb'][2]-sp['bb'][0])/2-sp['bb'][0],1020+ (1-ease((t-0.3)/1))*30,ti,blur=1-ease((t-0.3)/0.8))
    # outro
    if t>OUTRO-0.2:
        k=ease((t-OUTRO)/0.9); sp=G['final1']
        put(img,sp,W/2-(sp['bb'][2]-sp['bb'][0])/2-sp['bb'][0],760+(1-k)*40,k,blur=1-k)
        k2=ease((t-OUTRO-0.6)/0.9); sp=G['final2']
        put(img,sp,W/2-(sp['bb'][2]-sp['bb'][0])/2-sp['bb'][0],940+(1-k2)*30,k2*0.9,blur=1-k2)
        # gold divider
        ln=int(220*ease((t-OUTRO-0.3)/1.0)); img[915:918,W//2-ln:W//2+ln]=img[915:918,W//2-ln:W//2+ln]*0.3+np.array(GOLD)*0.7
    img*=G['vig']
    gr=G['grain'][fi%6]; img+=np.repeat(np.repeat(gr,2,0),2,1)[...,None]*5
    fade=smooth(t/0.6)*smooth((DUR-t)/1.0)
    img*=fade
    return np.clip(img,0,255).astype(np.uint8).tobytes()
if __name__=='__main__':
    out=sys.argv[1]; only=sys.argv[2] if len(sys.argv)>2 else None
    if only:  # preview stills
        init()
        for tt in only.split(','):
            fi=int(float(tt)*FPS); Image.frombytes('RGB',(W,H),frame(fi)).save(f"{out}_{tt}.png")
        sys.exit()
    p=subprocess.Popen(['ffmpeg','-y','-loglevel','error','-f','rawvideo','-pix_fmt','rgb24','-s',f'{W}x{H}','-r',str(FPS),'-i','-','-c:v','libx264','-preset','medium','-crf','18','-pix_fmt','yuv420p',out],stdin=subprocess.PIPE)
    with Pool(4,initializer=init) as pool:
        for i,b in enumerate(pool.imap(frame,range(NF),chunksize=4)):
            p.stdin.write(b)
            if i%150==0: print(i,'/',NF,flush=True)
    p.stdin.close(); p.wait(); print('video done',NF/FPS)
