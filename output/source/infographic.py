import numpy as np, wave, sys, math, subprocess, os
from PIL import Image, ImageDraw, ImageFont, ImageFilter
from multiprocessing import Pool
S=os.path.dirname(os.path.abspath(__file__))+'/..'
W,H=1080,1920; FPS=30
OFF=1.5; VOICE_END=31.35; OUTRO=OFF+VOICE_END+0.5; DUR=OUTRO+6.0
NF=int(DUR*FPS)
BLUE=(33,80,168); YEL=(244,190,66); NAVY=(20,40,92); BG=(246,248,253); GRAY=(140,150,175)
HANDLE="@ahlamtrainer"
scenes=[
 ("الفرصة",[("لا تنتظر",0.96,2.43,0xe88b),("أن يمنحك الطريق فرصة",2.58,3.42,0xeacd),("اصنع فرصتك بنفسك",3.93,5.61,0xe65f)]),
 ("البدايات",[("لا تخشَ البدايات المختلفة",6.27,8.13,0xe153),("ولا تتراجع",8.52,9.57,0xe8e5),("لأن الطريق لم تتضح ملامحه بعد",9.87,12.69,0xe8fd)]),
 ("الأبواب",[("فبعض الأبواب",13.32,14.28,0xeffd),("لا تُفتح",14.55,14.97,0xe897),("إلا لمن يجرؤ على المحاولة",15.15,17.52,0xe0da)]),
 ("الأحلام",[("وبعض الأحلام",17.97,19.02,0xea46),("لا تتحقق",19.29,20.22,0xe8e1),("إلا لمن يؤمن بها",20.37,21.9,0xe87d),("ويعمل لأجلها",22.2,23.58,0xe8b8)]),
 ("المبادرة",[("كن صاحب المبادرة",24.21,25.62,0xeb9b),("لا منتظر الفرصة",25.95,27.21,0xef53),("فالحياة تتسع",27.81,28.44,0xf1ce),("لمن يملك شجاعة السعي",28.77,31.3,0xe50a)]),
]
HL={"فرصتك","بنفسك","البدايات","الأبواب","يجرؤ","المحاولة","الأحلام","يؤمن","صاحب","المبادرة","تتسع","شجاعة","السعي","تتراجع","تنتظر"}
AR="٠١٢٣٤٥٦٧٨٩"
def arnum(n): return ''.join(AR[int(c)] for c in str(n))
SC_START=[OFF+s[1][0][1]-0.45 for s in scenes]
SC_END=[SC_START[i+1]-0.05 for i in range(len(scenes)-1)]+[OUTRO-0.2]
FC={}
def font(sz,kind='cairo',var=b'Bold'):
    k=(sz,kind,var)
    if k in FC: return FC[k]
    p={'cairo':'Cairo.ttf','mi':'MaterialIcons.ttf','fa':'fa-brands.ttf'}[kind]
    f=ImageFont.truetype(S+'/fonts/'+p,max(4,int(sz)),layout_engine=ImageFont.Layout.RAQM)
    if kind=='cairo': f.set_variation_by_name(var)
    FC[k]=f; return f
def clamp(x): return min(1,max(0,x))
def ease(x): x=clamp(x); return 1-(1-x)**3
def smooth(x): x=clamp(x); return x*x*(3-2*x)
def back(x):
    x=clamp(x); c=1.9; return 1+(c+1)*(x-1)**3+c*(x-1)**2
def lerp(a,b,k): return a+(b-a)*k
# ---------- layout ----------
CX0,CX1=80,1000; ICX=CX1-92; TXR=CX1-185; TMAX=640
def layout():
    f=font(56); sp=f.getlength(' ')
    blocks=[]
    for si,(title,ph) in enumerate(scenes):
        cards=[]
        for (txt,t0,t1,ic) in ph:
            words=txt.split(); lens=[len(w)+1 for w in words]; tot=sum(lens); acc=0; ws=[]
            for w,l in zip(words,lens):
                ws.append(dict(t=w,start=OFF+t0+(t1-t0)*acc/tot-0.1,adv=f.getlength(w,direction='rtl',language='ar'),hl=w in HL)); acc+=l
            lines=[];cur=[];cw=0
            for wd in ws:
                if cur and cw+sp+wd['adv']>TMAX: lines.append(cur);cur=[];cw=0
                cw+=(sp if cur else 0)+wd['adv']; cur.append(wd)
            lines.append(cur)
            h=150 if len(lines)==1 else 222
            for li,ln in enumerate(lines):
                x=TXR
                for wd in ln:
                    x-=wd['adv']; wd['x']=x; wd['ly']=li; x-=sp
            cards.append(dict(words=ws,lines=len(lines),h=h,start=OFF+t0-0.4,icon=ic,end=OFF+t1))
        gap=34; tot=sum(c['h'] for c in cards)+gap*(len(cards)-1); y=1080-tot/2+20
        for c in cards: c['y']=y; y+=c['h']+gap
        blocks.append(dict(title=title,cards=cards,start=SC_START[si],end=SC_END[si]))
    return blocks
# ---------- audio ----------
def voice_features():
    w=wave.open(S+'/voice16.wav'); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
    sr=16000; nb=28; feats=np.zeros((NF,nb)); edges=np.geomspace(90,5000,nb+1); fr=np.fft.rfftfreq(1024,1/sr); win=np.hanning(1024)
    for i in range(NF):
        c=int((i/FPS-OFF)*sr)
        if c<0 or c+1024>len(a): continue
        sp=np.abs(np.fft.rfft(a[c:c+1024]*win)); feats[i]=[sp[(fr>=edges[k])&(fr<edges[k+1])].mean() for k in range(nb)]
    feats=np.log1p(feats*4); feats/=np.percentile(feats[feats>0],98)+1e-6
    for i in range(1,NF): feats[i]=np.maximum(feats[i],feats[i-1]*0.8)
    return np.clip(feats,0,1.1)
G={}
def init():
    rng=np.random.default_rng(5)
    G['blocks']=layout(); G['feats']=voice_features()
    bg=Image.new('RGBA',(W,H),BG+(255,)); d=ImageDraw.Draw(bg)
    for y in range(20,H,44):
        for x in range(20,W,44): d.ellipse([x-1.6,y-1.6,x+1.6,y+1.6],fill=(214,220,236,255))
    G['bg']=bg
    blob=Image.new('L',(600,600),0); ImageDraw.Draw(blob).ellipse([100,100,500,500],fill=255); blob=blob.filter(ImageFilter.GaussianBlur(70))
    G['blobs']=[]
    for col,a in ((BLUE,40),(YEL,70),(BLUE,30)):
        im=Image.new('RGBA',(600,600),col+(0,)); im.putalpha(blob.point(lambda v,a=a:v*a//255)); G['blobs'].append(im)
    G['sq']=np.c_[rng.uniform(0,W,26),rng.uniform(0,H,26),rng.uniform(10,26,26),rng.uniform(0,6.28,26),rng.integers(0,2,26),rng.uniform(8,22,26)]
    logo=Image.open(S+'/logo.png').convert('RGBA'); G['logo']=logo
    sh=Image.new('L',(960,300),0); ImageDraw.Draw(sh).rounded_rectangle([30,30,930,270],40,fill=110); G['shadow']=sh.filter(ImageFilter.GaussianBlur(18))
def logo_at(img,cx,top,h,alpha):
    lg=G['logo']; w=int(lg.width*h/lg.height)
    im=lg.resize((w,int(h)),Image.LANCZOS)
    if alpha<1: im.putalpha(im.getchannel('A').point(lambda v:int(v*alpha)))
    img.alpha_composite(im,(int(cx-w/2),int(top)))
def fade_layer(layer,a):
    if a>=1: return layer
    layer.putalpha(layer.getchannel('A').point(lambda v:int(v*a))); return layer
def draw_card(img,c,num,t,ox,alpha,active):
    h=int(c['h']); cw=CX1-CX0+60
    L=Image.new('RGBA',(cw,h+60),(0,0,0,0))
    sh=G['shadow'].resize((cw,h+60)); shl=Image.new('RGBA',(cw,h+60),(30,50,110,0)); shl.putalpha(sh.point(lambda v:int(v*0.55)))
    L.alpha_composite(shl,(0,10))
    d=ImageDraw.Draw(L,'RGBA'); x0,y0,x1,y1=30,30,cw-30,30+h
    d.rounded_rectangle([x0,y0,x1,y1],34,fill=(255,255,255,255),outline=(YEL+(255,)) if active else (226,231,243,255),width=5 if active else 2)
    # accent bar (right, RTL start)
    d.rounded_rectangle([x1-12,y0+26,x1-4,y1-26],4,fill=BLUE+(255,))
    # icon circle with pop
    ip=back((t-c['start']-0.15)/0.5); icx=ICX-CX0+30; icy=y0+h/2
    if ip>0:
        r=58*ip; d.ellipse([icx-r,icy-r,icx+r,icy+r],fill=(BLUE if not active else BLUE)+(255,))
        d.ellipse([icx-r-9,icy-r-9,icx+r+9,icy+r+9],outline=(YEL+(255,)),width=3)
        f=font(64*ip,'mi'); ch=chr(c['icon']); bb=f.getbbox(ch)
        d.text((icx-(bb[0]+bb[2])/2,icy-(bb[1]+bb[3])/2),ch,font=f,fill=(255,255,255,255))
        # number badge
        bx,by,br=icx+44,icy-46,22*ip
        d.ellipse([bx-br,by-br,bx+br,by+br],fill=YEL+(255,)); f2=font(28*ip,'cairo',b'Black'); s=arnum(num); bb=f2.getbbox(s)
        d.text((bx-(bb[0]+bb[2])/2,by-(bb[1]+bb[3])/2),s,font=f2,fill=NAVY+(255,))
    # words
    f=font(56); lh=72; ty0=y0+h/2-(c['lines']*lh)/2-14
    for w in c['words']:
        p=(t-w['start'])/0.38
        if p<=0: continue
        e=ease(p); x=w['x']-CX0+30; y=ty0+w['ly']*lh+(1-e)*18
        if w['hl']:
            m=ease((t-w['start']-0.15)/0.45)
            if m>0:
                bx1=x+w['adv']+6; bx0=bx1-(w['adv']+12)*m
                d.rounded_rectangle([bx0,y+44,bx1,y+78],8,fill=YEL+(int(170*e),))
        d.text((x,y),w['t'],font=f,fill=NAVY+(int(255*e),),direction='rtl',language='ar')
    img.alpha_composite(fade_layer(L,alpha),(int(CX0-30+ox),int(c['y']-30)))
def background(img,t):
    for k,bl in enumerate(G['blobs']):
        x=[150,820,540][k]+120*math.sin(t*0.25+k*2); y=[420,1150,1700][k]+90*math.cos(t*0.2+k)
        img.alpha_composite(bl,(int(x-300),int(y-300)))
    d=ImageDraw.Draw(img,'RGBA')
    for x,y,s,ph,ci,v in G['sq']:
        yy=(y-v*t)%H; a=int(45+25*math.sin(t+ph)); col=(BLUE if ci else YEL)+(a,)
        r=s/2; ang=ph+t*0.3; pts=[(x+r*1.4*math.cos(ang+q*math.pi/2),yy+r*1.4*math.sin(ang+q*math.pi/2)) for q in range(4)]
        d.polygon(pts,fill=col)
def frame(fi):
    t=fi/FPS
    img=G['bg'].copy(); background(img,t); d=ImageDraw.Draw(img,'RGBA')
    # ---- logo: intro center -> header -> outro center
    if t<OUTRO:
        k=smooth((t-1.15)/0.85); a=ease(t/0.7); sc=lerp(0.85,1,ease(t/0.9))
        h=lerp(620*sc,250,k); top=lerp(960-310*sc,48,k)
        logo_at(img,W/2,top,h,a)
    else:
        k=smooth((t-OUTRO)/1.0); h=lerp(250,560,k); top=lerp(48,250,k)
        logo_at(img,W/2,top,h,1)
    main=smooth((t-1.8)/0.6)*smooth((OUTRO+0.3-t)/0.5)
    if main>0:
        # progress bar (RTL)
        pr=clamp((t-OFF)/VOICE_END); y=335
        d.rounded_rectangle([CX0,y-4,CX1,y+4],4,fill=(222,228,242,int(255*main)))
        xp=CX1-(CX1-CX0)*pr; d.rounded_rectangle([xp,y-4,CX1,y+4],4,fill=BLUE+(int(255*main),))
        d.ellipse([xp-11,y-11,xp+11,y+11],fill=YEL+(int(255*main),),outline=(255,255,255,int(255*main)),width=3)
        # scene dots
        cur=max([i for i in range(5) if t>=SC_START[i]-0.3] or [0])
        for i in range(5):
            cx=W/2+(2-i)*34; r=8 if i==cur else 5
            d.ellipse([cx-r,380-r,cx+r,380+r],fill=(YEL if i==cur else (200,208,228))+(int(255*main),))
        # footer handle + waveform
        f=font(30,'cairo',b'SemiBold'); bb=f.getbbox(HANDLE); d.text((W/2-(bb[0]+bb[2])/2,1840),HANDLE,font=f,fill=GRAY+(int(220*main),))
        fe=G['feats'][fi]; n=len(fe)
        for k in range(n):
            hh=4+70*fe[k]
            for sgn in (-1,1):
                x=W/2+sgn*(8+k*13)
                col=YEL if k%3==0 else BLUE
                d.rounded_rectangle([x-3.5,1760-hh/2,x+3.5,1760+hh/2],3,fill=col+(int(230*main*(1-0.5*k/n)),))
    # ---- scenes
    for si,b in enumerate(G['blocks']):
        if t<b['start']-0.6 or t>b['end']+0.8: continue
        # chip
        ca=ease((t-b['start'])/0.5); co=ease((t-b['end']+0.35)/0.45); al=ca*(1-co)
        if al>0:
            f=font(46,'cairo',b'Black'); tt=b['title']; tw=f.getlength(tt,direction='rtl',language='ar')
            cw=tw+150; x1=W/2+cw/2; x0=W/2-cw/2; y0=440+(1-ca)*25-co*25
            d.rounded_rectangle([x0,y0,x1,y0+84],42,fill=BLUE+(int(255*al),))
            d.ellipse([x1-74,y0+10,x1-10,y0+74],fill=YEL+(int(255*al),))
            f2=font(36,'cairo',b'Black'); s=arnum(si+1); bb=f2.getbbox(s)
            d.text((x1-42-(bb[0]+bb[2])/2,y0+42-(bb[1]+bb[3])/2),s,font=f2,fill=NAVY+(int(255*al),))
            bb=f.getbbox(tt,direction='rtl',language='ar')
            d.text((x1-90-tw,y0+42-(bb[1]+bb[3])/2),tt,font=f,fill=(255,255,255,int(255*al)),direction='rtl',language='ar')
        # timeline connector
        cards=b['cards']
        for i in range(len(cards)-1):
            a,c2=cards[i],cards[i+1]; p=ease((t-c2['start']+0.1)/0.5)*(1-co)
            if p<=0: continue
            ya=a['y']+a['h']/2; yb=c2['y']+c2['h']/2; ye=lerp(ya,yb,p)
            yy=ya
            while yy<ye:
                d.rounded_rectangle([ICX-3,yy,ICX+3,min(ye,yy+12)],3,fill=YEL+(int(255*p),)); yy+=22
        for i,c in enumerate(cards):
            p=ease((t-c['start'])/0.6)
            if p<=0: continue
            ex=ease((t-b['end']+0.45-i*0.07)/0.45)
            al=p*(1-ex)
            if al<=0.01: continue
            nxt=cards[i+1]['start']+0.3 if i+1<len(cards) else b['end']
            active=c['start']+0.3<=t<nxt
            draw_card(img,c,i+1,t,(1-p)*140-ex*160,al if active or t<nxt else al*0.92,active)
    # ---- outro: socials
    if t>OUTRO:
        k=ease((t-OUTRO-0.9)/0.7)
        f=font(46,'cairo',b'Bold'); s="تابعونا على"; tw=f.getlength(s,direction='rtl',language='ar')
        d.text((W/2-tw/2,1130+(1-k)*30),s,font=f,fill=NAVY+(int(255*k),),direction='rtl',language='ar')
        for j,(ic,bgc) in enumerate(((0xe61b,(0,0,0)),(0xf16d,(214,41,118)))):
            k2=back((t-OUTRO-1.2-j*0.25)/0.7); a=int(255*clamp((t-OUTRO-1.2-j*0.25)/0.4))
            if k2<=0: continue
            y0=1240+j*150+(1-clamp(k2))*40; x0,x1=200,880
            d.rounded_rectangle([x0,y0,x1,y0+118],59,fill=(255,255,255,a),outline=(226,231,243,a),width=3)
            cx,cy,r=x1-62,y0+59,46*min(1.1,k2)
            if j==1:
                g=Image.new('RGBA',(100,100)); gd=ImageDraw.Draw(g)
                for q in range(100): gd.line([(0,q),(100,q)],fill=(int(lerp(131,252,q/99)),int(lerp(58,175,q/99)),int(lerp(180,69,q/99)),255))
                m=Image.new('L',(100,100),0); ImageDraw.Draw(m).ellipse([0,0,99,99],fill=a); g.putalpha(m)
                g=g.resize((max(1,int(2*r)),max(1,int(2*r)))); img.alpha_composite(g,(int(cx-r),int(cy-r))); d=ImageDraw.Draw(img,'RGBA')
            else:
                d.ellipse([cx-r,cy-r,cx+r,cy+r],fill=bgc+(a,))
            fi_=font(r*1.0,'fa'); ch=chr(ic); bb=fi_.getbbox(ch)
            d.text((cx-(bb[0]+bb[2])/2,cy-(bb[1]+bb[3])/2),ch,font=fi_,fill=(255,255,255,a))
            f3=font(50,'cairo',b'Bold'); bb=f3.getbbox(HANDLE)
            d.text(((x0+x1-124)/2-(bb[0]+bb[2])/2,cy-(bb[1]+bb[3])/2),HANDLE,font=f3,fill=NAVY+(a,))
        # ring accent around logo
        rk=0
        if rk>0:
            R=330; cy=250+280
            d.arc([W/2-R,cy-R,W/2+R,cy+R],start=-90,end=-90+360*rk,fill=YEL+(140,),width=4)
    out=np.asarray(img.convert('RGB'),np.float32)
    fade=smooth(t/0.4)*smooth((DUR-t)/0.6)
    if fade<1: out=out*fade+np.array(BG,np.float32)*(1-fade)
    return out.astype(np.uint8).tobytes()
if __name__=='__main__':
    out=sys.argv[1]; only=sys.argv[2] if len(sys.argv)>2 else None
    if only:
        init()
        for tt in only.split(','): Image.frombytes('RGB',(W,H),frame(int(float(tt)*FPS))).save(f"{out}_{tt}.png")
        sys.exit()
    print('DUR',DUR,'scenes',[round(x,2) for x in SC_START],flush=True)
    p=subprocess.Popen(['ffmpeg','-y','-loglevel','error','-f','rawvideo','-pix_fmt','rgb24','-s',f'{W}x{H}','-r',str(FPS),'-i','-','-c:v','libx264','-preset','slow','-crf','20','-pix_fmt','yuv420p',out],stdin=subprocess.PIPE)
    with Pool(4,initializer=init) as pool:
        for i,b in enumerate(pool.imap(frame,range(NF),chunksize=4)):
            p.stdin.write(b)
    p.stdin.close(); p.wait(); print('video done',NF/FPS)
