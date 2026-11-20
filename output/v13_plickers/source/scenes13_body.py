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

PAT=[[1,1,1,0,1],[1,0,1,0,1],[1,0,1,1,1],[1,0,0,0,1],[1,1,0,1,1]]
LET=["أ","ب","ج","د"]
def pcard(x,cx,cy,s,rot,a,lit=0,letters=True):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.rotate(rot); x.scale(S_(s),S_(s))
    if lit>0: radial(x,0,0,200,GOLD,0.5*lit*a)
    rrect(x,-100,-100,200,200,14,WHITE,a); rrect(x,-100,-100,200,200,14,GOLD if lit>0.5 else (0.8,0.84,0.92),a,fill=False,lw=4)
    c=26
    for r in range(5):
        for q in range(5):
            if PAT[r][q]: x.rectangle(-65+q*c,-65+r*c,c+0.5,c+0.5)
    rgba((0.06,0.08,0.15),a,x); x.fill()
    if letters:
        for k,(lx,ly) in enumerate(((0,-84),(84,0),(0,84),(-84,0))):
            x.save(); x.translate(lx,ly); x.rotate(k*math.pi/2); blit(x,sprite(LET[k],22,(14,32,80),b'Black'),0,0,1,a*0.9); x.restore()
    x.restore()
def phone_cam(x,cx,cy,s,a,t,scan=0):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    rrect(x,-100,-180,200,360,30,WHITE,a); rrect(x,-86,-155,172,300,14,(0.06,0.1,0.22),a)
    if scan>0:
        y=-150+300*((t*0.9)%1); rrect(x,-86,y,172,6,3,GOLD2,a); radial(x,0,y,120,GOLD,0.4*a)
    for k in ((-60,-120),(60,-120),(-60,120),(60,120)):
        sx=1 if k[0]>0 else -1; sy=1 if k[1]>0 else -1
        poly(x,[(k[0],k[1]-sy*24),(k[0],k[1]),(k[0]-sx*24,k[1])],GOLD,a,fill=False,lw=5,close=False)
    x.restore()
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    k=seg(t,s0-0.3,s0+0.3)
    if k<=0: return
    a=eo(k); nod=eo(seg(t,P[0]['words'][2]['start']-0.2,P[0]['words'][2]['start']+0.4))
    for r in range(5):
        for q in range(6):
            j=r*6+q; kk=back(seg(t,s0+j*0.025,s0+0.25+j*0.025))
            if kk>0: person(x,-250+q*100,-250+r*110,0.55*kk,mix(SOFT,WHITE,0.2),a)
    if nod>0:
        icon_disc(x,0xe324,0,-10,110,nod,a,glow=0.5); ban(x,0,-10,150,nod,a,lw=12)
    kt=back(seg(t,s1-0.2,s1+0.3))
    if kt>0:
        p=seg(t,s1,ph(1)['end']); cx,cy=0,330
        circle(x,cx,cy,90*kt,CARD,a); x.new_path(); x.arc(cx,cy,90*kt,-math.pi/2,-math.pi/2+2*math.pi*max(0.001,1-p)); x.set_line_width(12); rgba(GOLD,a,x); x.stroke()
        blit(x,sprite(str(int(30*(1-p))),64,(255,255,255),b'Black',rtl=False),cx,cy-4,kt,a)
def scene_2(x,t):
    s2,s3=ph(2)['start'],ph(3)['start']
    k=back(seg(t,s2-0.35,s2+0.2)); up=eio(seg(t,s3-0.1,s3+0.4))
    if k>0:
        sp=sprite("Plickers",lerp(120,70,up).__int__(),(246,196,72),b'Black',glow=(245,180,40),rtl=False)
        blit(x,sp,0,lerp(-60,-330,up),k,1)
        if up<1: blit(x,sprite("أداة تقويم مجانية",38,(255,255,255),b'Bold'),0,60,k,1-up)
    kc=back(seg(t,s3-0.1,s3+0.4))
    if kc>0:
        rot=0
        kl=seg(t,P[3]['words'][-4]['start']-0.1,ph(3)['end'])
        if kl>0: rot=-math.pi/2*min(3,int(kl*4))*1.0 + 0
        pcard(x,0,-20,1.6*kc,rot*eo(1) ,1,lit=kl,letters=True)
        for j in range(4):
            kk=back(seg(t,P[3]['words'][-4]['start']+j*0.3,P[3]['words'][-4]['start']+0.3+j*0.3))
            if kk>0:
                px=(1.5-j)*150; icon_disc(x,0,px,300,0,0,0)
                circle(x,px,300,46*kk,GOLD if j==int(min(3,kl*4)) else CARD,1); blit(x,sprite(LET[j],44,(14,32,80) if j==int(min(3,kl*4)) else (255,255,255),b'Black'),px,296,kk,1)
        ks=seg(t,P[3]['words'][1]['start'],P[3]['words'][1]['start']+0.6)
        if 0<ks<1: blit(x,mi(0xe8ad,70,(246,196,72)),230,-220,back(ks),1)
def scene_3(x,t):
    s4,s5,s6=[ph(i)['start'] for i in (4,5,6)]
    cur=0 if t<s5-0.1 else 1 if t<s6-0.1 else 2
    xs=[200,0,-200]; ny=-340
    for i in range(3):
        if i<2: poly(x,[(xs[i]-50,ny),(xs[i+1]+50,ny)],GOLD if cur>i else SOFT,0.9 if cur>i else 0.35,fill=False,lw=5,close=False)
        num_node(x,xs[i],ny,44,i+1,1,i==cur,1)
    k=back(seg(t,s4-0.3,s4+0.2)); ko=eio(seg(t,s5-0.15,s5+0.25))
    if k>0 and ko<1:
        a=1-ko; x.save(); x.translate(0,lerp(0,-40,ko)); x.scale(S_(k),S_(k))
        rrect(x,-330,-230,660,420,24,WHITE,a); rrect(x,-310,-210,620,380,14,(0.12,0.28,0.6),a)
        blit(x,sprite("ما حالة الماء عند ١٠٠°؟",38,(255,255,255),b'Black'),0,-150,1,a)
        for j,(l,txt) in enumerate(zip(LET,("صلبة","سائلة","غازية","بلازما"))):
            kk=eo(seg(t,s4+0.6+j*0.15,s4+0.9+j*0.15)); r,q=j//2,j%2; px=150-q*300; py=-50+r*110
            rrect(x,px-130,py-40,260,80,16,GOLD if j==2 and t>s4+1.5 else (0.2,0.38,0.75),a*kk)
            blit(x,sprite(l+") "+txt,32,(14,32,80) if j==2 and t>s4+1.5 else (255,255,255),b'Bold'),px,py-3,kk,a)
        x.restore()
    k=seg(t,s5-0.2,s5+0.3); ko=eio(seg(t,ph(6)['end']+0.1,ph(6)['end']+0.5))
    if k>0 and ko<1:
        a=eo(k)*(1-ko); raise_=eio(seg(t,P[5]['words'][2]['start']-0.2,P[5]['words'][2]['start']+0.5))
        turn=eio(seg(t,P[5]['words'][-1]['start']-0.6,P[5]['words'][-1]['start']+0.3))
        ans=[2,2,1,2,0,2,3,2,2]
        for j in range(9):
            r,q=j//3,j%3; px=230-q*230; py=-170+r*230
            person(x,px,py+110,0.75,WHITE,a)
            cy=py+lerp(60,-30,raise_)
            scanned=seg(t,s6+0.4+j*0.12,s6+0.6+j*0.12) if t>s6 else 0
            pcard(x,px,cy,0.38,-math.pi/2*ans[j]*turn,a,lit=scanned,letters=False)
            if scanned>0: circle(x,px+48,cy-48,16*back(scanned),GOLD,a); check(x,px+48,cy-49,0.45,1,NAVY1,a,lw=4)
        kp=seg(t,s6-0.2,s6+0.3)
        if kp>0:
            u=seg(t,s6+0.2,ph(6)['end']); px=lerp(330,-330,eio(u))
            phone_cam(x,px,lerp(380,320,eo(kp)),0.75*eo(kp),a,t,scan=1)
def scene_4(x,t):
    s7,s8=ph(7)['start'],ph(8)['start']
    k=seg(t,s7-0.2,s7+0.3); ko=eio(seg(t,s8-0.15,s8+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko); base=80
        vals=[2,4,20,4]
        for j,v in enumerate(vals):
            g=back(seg(t,s7+0.3+j*0.12,s7+0.7+j*0.12),1.2); h=v*14*g; px=225-j*150
            rrect(x,px-50,base-h,100,h,12,GOLD if j==2 else (0.4,0.52,0.8),a)
            blit(x,sprite(LET[j],40,(255,255,255),b'Black'),px,base+36,1,a)
        pv=int(77*eio(seg(t,s7+0.5,s7+1.6)))
        blit(x,sprite(''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(pv))+'٪',100,(246,196,72),b'Black',glow=(245,180,40)),0,-300,1,a)
        blit(x,sprite("إجابات صحيحة",34,(255,255,255),b'Bold'),0,-210,1,a)
        names=["سارة","نورة","ريم","هند"]; ok=[1,1,0,1]
        for j,nm in enumerate(names):
            kk=eo(seg(t,P[7]['words'][-3]['start']+j*0.15,P[7]['words'][-3]['start']+0.35+j*0.15)); y=180+j*62
            rrect(x,-260,y-26,520,52,14,WHITE,a*kk*0.95); blit(x,sprite(nm,28,(14,32,80),b'Bold'),170,y-2,kk,a)
            blit(x,sprite(LET[2 if ok[j] else 1],28,(14,32,80),b'Black'),0,y-2,kk,a)
            circle(x,-200,y,18*kk,GOLD if ok[j] else (0.85,0.3,0.35),a)
            if ok[j]: check(x,-200,y-1,0.5,1,NAVY1,a,lw=4)
    k=seg(t,s8-0.2,s8+0.3)
    if k>0:
        a=eo(k)
        icon_disc(x,0xe7fd,0,-220,70,back(k),a,gold=True,glow=0.6)
        for j,(lab,ic,px) in enumerate((("الدرس التالي",0xe5c8,200),("إعادة الشرح",0xe042,-200))):
            kk=back(seg(t,P[8]['words'][1+j*3 if 1+j*3<len(P[8]['words']) else -1]['start']-0.1,P[8]['words'][1+j*3 if 1+j*3<len(P[8]['words']) else -1]['start']+0.35))
            if kk<=0: continue
            glowline(x,partial([(0,-150),(px,60)],eo(kk)),GOLD,a,5)
            icon_disc(x,ic,px,120,64,kk,a,gold=(j==0)); blit(x,sprite(lab,36,(255,255,255),b'Bold'),px,230,kk,a)
def scene_5(x,t):
    s9,s10=ph(9)['start'],ph(10)['start']
    k=seg(t,s9-0.2,s9+0.3); ko=eio(seg(t,s10-0.15,s10+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        for j,(lab,ic) in enumerate((("تقويم فوري",0xe425),("بلا أجهزة للطلاب",0xe644),("بلا أوراق تصحيح",0xf0c5))):
            kk=back(seg(t,P[9]['words'][min(j*2,len(P[9]['words'])-1)]['start']-0.1,P[9]['words'][min(j*2,len(P[9]['words'])-1)]['start']+0.35)); y=-220+j*190
            if kk<=0: continue
            x.save(); x.translate(0,y); x.scale(S_(kk),S_(kk))
            rrect(x,-300,-70,600,140,70,GOLD if j==0 else CARD,a); rrect(x,-300,-70,600,140,70,GOLD,a,fill=False,lw=3)
            icon_disc(x,ic,220,0,48,1,a,gold=(j!=0)); blit(x,sprite(lab,44,(14,32,80) if j==0 else (255,255,255),b'Black'),-40,-3,1,a); x.restore()
    fin=eio(seg(t,VEND+0.1,VEND+0.7))
    k=seg(t,s10-0.15,s10+0.4)
    if k>0 and fin<1:
        a=eo(k)*(1-fin)
        icon_disc(x,0xe866,0,-80,120,back(k),a,gold=True,glow=1); blit(x,sprite("احفظ",44,(246,196,72),b'Black'),0,80,back(k),a)
        for j,(ic,px,py) in enumerate(((0xe0bf,-240,-200),(0xe87d,240,-190),(0xe0b7,-230,120),(0xe8dc,250,140))):
            kk=back(seg(t,s10+0.6+j*0.2,s10+1.0+j*0.2)); icon_disc(x,ic,px,py+8*math.sin(t*3+j),44,kk,a,gold=(j%2==1))
        kd=seg(t,P[10]['words'][-1]['start']-0.4,VEND+1)
        if kd>0:
            for j in range(3): blit(x,mi(0xe5db,64,(246,196,72)),0,190+j*50+8*math.sin(t*5),1,a*((math.sin(t*5-j)+1)/2))
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        pcard(x,0,-200,1.2*back(seg(t,VEND+0.1,VEND+0.7)),math.sin(t*1.5)*0.12,a,lit=1)
        e1=eo(seg(t,VEND+0.5,VEND+1.1)); e2=eo(seg(t,VEND+0.9,VEND+1.5))
        blit(x,sprite("قيّم صفك كله",76,(255,255,255),b'Black'),0,60+(1-e1)*30,1,a*e1)
        blit(x,sprite("في ٣٠ ثانية",96,(246,196,72),b'Black',glow=(245,180,40)),0,190+(1-e2)*30,1+0.03*math.sin(t*3),a*e2)
        blit(x,sprite("Plickers",44,(200,210,235),b'Bold',rtl=False),0,300,1,a*e2)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
def OVERLAY(x,t): pass
