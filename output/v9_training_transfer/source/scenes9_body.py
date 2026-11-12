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

def meter(x,cx,cy,v,a,label=None):
    rrect(x,cx-30,cy-150,60,300,30,CARD,a); rrect(x,cx-30,cy-150,60,300,30,WHITE,a*0.5,fill=False,lw=3)
    h=max(0,min(1,v))*280; rrect(x,cx-20,cy+140-h,40,h,20,mix((0.6,0.65,0.8),GOLD,v),a)
    blit(x,mi(0xea0b,46,(246,196,72)),cx,cy-195,1,a)
def cliff(x,x0,x1,y,a,col):
    poly(x,[(x0,y),(x1,y),(x1+(30 if x1>x0 else -30),y+260),(x0,y+260)],col,a)
    poly(x,[(x0,y),(x1,y)],GOLD2,a,fill=False,lw=6,close=False)
def scene_1(x,t):
    s0,s1,s2,s3=[ph(i)['start'] for i in range(4)]
    # training room
    k=back(seg(t,s0-0.4,s0+0.2)); ko=eio(seg(t,s1-0.2,s1+0.3))
    if k>0 and ko<1:
        a=1-ko; x.save(); x.scale(S_(k),S_(k))
        rrect(x,-220,-300,440,230,16,WHITE,a); rrect(x,-206,-286,412,202,8,(0.12,0.3,0.62),a)
        blit(x,mi(0xe80c,90,(246,196,72)),0,-190,1,a)
        person(x,280,-120,1.0,GOLD,a)
        for r in range(2):
            for c in range(4): person(x,-240+c*160,90+r*150,0.9,mix(SOFT,WHITE,0.3),a)
        x.restore()
    # ideas + enthusiasm
    k=back(seg(t,s1-0.3,s1+0.2)); ko=eio(seg(t,s2-0.2,s2+0.3))
    en=eo(seg(t,s1,ph(1)['end']))*(1-0.85*eo(seg(t,s2+0.2,ph(3)['end'])))
    if k>0:
        m_on=1-eio(seg(t,ph(3)['end']+0.1,ph(3)['end']+0.5))
        meter(x,330,-20,en,k*m_on)
    if k>0 and ko<1:
        a=1-ko; person(x,0,170,2.2*k,WHITE,a)
        for j in range(5):
            kk=back(seg(t,wt(1,3)-0.1+j*0.12,wt(1,3)+0.3+j*0.12)); ang=-math.pi*0.9+j*math.pi*0.2
            icon_disc(x,0xe0f0 if j%2==0 else 0xea0b,190*math.cos(ang),-60+190*math.sin(ang),42,kk,a,gold=(j%2==0),glow=0.5)
    # back to work desk
    k=back(seg(t,s2-0.2,s2+0.3)); ko=eio(seg(t,ph(3)['end']+0.1,ph(3)['end']+0.5))
    if k>0 and ko<1:
        a=1-ko; walk=eio(seg(t,s2,ph(2)['end']))
        rrect(x,-260,120,300,24,8,(0.55,0.42,0.3),a); rrect(x,-240,144,20,140,6,(0.45,0.34,0.25),a); rrect(x,0,144,20,140,6,(0.45,0.34,0.25),a)
        blit(x,mi(0xe320,120,(255,255,255)),-110,60,k,a)
        person(x,lerp(200,40,walk),170,1.4*k,mix(WHITE,SOFT,seg(t,s3,s3+0.6)),a)
        kq=back(seg(t,wt(3,2)-0.1,wt(3,2)+0.4))
        if kq>0:
            q=sprite("؟",150,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,-100,-200,kq*(1+0.05*math.sin(t*5)),a)
            glowline(x,[(-300,-60),(100,-60)],SOFT,a*kq*0.8,4)
            blit(x,sprite("لا تغيير",36,(200,210,235),b'Bold'),-100,-20,kq,a)
def scene_2(x,t):
    s4,s5,s6,s7=[ph(i)['start'] for i in (4,5,6,7)]
    k=back(seg(t,s4-0.35,s4+0.2)); ko=eio(seg(t,s5-0.2,s5+0.3))
    if k>0 and ko<1:
        magnifier(x,60*math.sin(t*2),-20,1.6*k*(1-ko),1)
    k=seg(t,s5-0.2,s5+0.3); ko=eio(seg(t,s6-0.2,s6+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        k1=back(seg(t,wt(5,1)-0.15,wt(5,1)+0.3)); k2=back(seg(t,wt(5,5)-0.15,wt(5,5)+0.3)); ke=back(seg(t,wt(5,3)-0.1,wt(5,3)+0.3))
        if k1>0: icon_disc(x,0xea19,220,-20,120,k1,a); blit(x,sprite("المعرفة",40,(255,255,255),b'Bold'),220,150,k1,a)
        if k2>0: icon_disc(x,0xe8e5,-220,-20,120,k2,a,gold=True,glow=1); blit(x,sprite("التغيير",40,(246,196,72),b'Black'),-220,150,k2,a)
        if ke>0:
            poly(x,[(-30,-45),(30,-45)],GOLD2,a,fill=False,lw=12,close=False); poly(x,[(-30,5),(30,5)],GOLD2,a,fill=False,lw=12,close=False)
            glowline(x,[(26,-80*ke),(-26,40*ke)],GOLD2,a,10)
    # rich program: 5 stars + info cards
    k=seg(t,s6-0.2,s6+0.3); ko=eio(seg(t,s7-0.2,s7+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        rrect(x,-300,-260,600,360,30,CARD,a); rrect(x,-300,-260,600,360,30,WHITE,a*0.4,fill=False,lw=3)
        for j in range(4):
            kk=eo(seg(t,s6+0.2+j*0.2,s6+0.5+j*0.2)); rrect(x,-250,-210+j*70,420*kk,34,12,SOFT if j%2 else WHITE,a*0.9)
        for j in range(5):
            kk=back(seg(t,wt(6,len(P[6]['words'])-2)+j*0.12,wt(6,len(P[6]['words'])-2)+0.3+j*0.12))
            poly(x,star_pts(-200+j*100,190,40*kk,rot=-math.pi/2),GOLD,a)
    # inside the hall vs outside
    k=seg(t,s7-0.2,s7+0.3)
    if k>0:
        a=eo(k); out=eio(seg(t,P[7]['words'][-2]['start']-0.3,P[7]['words'][-1]['start']+0.3))
        rrect(x,40,-250,360,440,24,CARD,a); rrect(x,40,-250,360,440,24,WHITE,a*0.7,fill=False,lw=5)
        blit(x,sprite("داخل القاعة",34,(255,255,255),b'Bold'),220,-290,1,a)
        rrect(x,30,-40,24,150,6,GOLD,a)
        blit(x,mi(0xea19,90,(200,210,235)),220,-110,1,a*(1-0.5*out))
        person(x,lerp(220,-200,out),120,1.4,mix(WHITE,GOLD,out),a)
        if out>0:
            for j,ic in enumerate((0xea3d,0xe2e6,0xf092)):
                kk=back(seg(out,0.5+j*0.15,0.8+j*0.15)); icon_disc(x,ic,-320+j*120,-140,46,kk,a,gold=(j==1),glow=0.6)
            blit(x,sprite("خارجها",38,(246,196,72),b'Black'),-200,-260,out,a)
def scene_3(x,t):
    s8,s9=ph(8)['start'],ph(9)['start']
    k=seg(t,s8-0.3,s8+0.4)
    if k<=0: return
    a=eo(k); gap=eo(seg(t,wt(8,len(P[8]['words'])-4)-0.2,wt(8,len(P[8]['words'])-4)+0.6))
    g=lerp(10,170,gap)
    cliff(x,420,g,40,a,CARD); cliff(x,-420,-g,40,a,(0.16,0.3,0.6))
    blit(x,sprite("التعلّم",42,(255,255,255),b'Black'),220+g/3,0,1,a); blit(x,sprite("التطبيق",42,(246,196,72),b'Black'),-220-g/3,0,1,a)
    kq=back(seg(t,wt(8,3)-0.1,wt(8,3)+0.4))
    if kq>0:
        icon_disc(x,0xe7af,280,-160,56,kq,a,gold=True); check(x,350,-210,0.8*kq,1,GOLD2,a,lw=5)
    if gap>0:
        for j in range(6): poly(x,[(0,60+j*40),(0,80+j*40)],GOLD2,a*gap*0.7,fill=False,lw=4,close=False)
        blit(x,sprite("الفجوة",40,(246,196,72),b'Black'),0,-60,gap,a)
    person(x,300,30,1.0,WHITE,a)
    kb=seg(t,s9-0.1,s9+0.8)
    if kb>0:
        for j in range(3):
            kk=back(seg(kb,j*0.25,j*0.25+0.4)); px=g-(2*g)*(j+0.5)/3
            rrect(x,px-g/3+4,36,2*g/3-8,22*kk,6,(0.55,0.42,0.3),a*0.6)
        q=sprite("؟",120,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,0,-200,back(kb)*(1+0.05*math.sin(t*5)),a)
def scene_4(x,t):
    s10,s11,s12,s13=[ph(i)['start'] for i in (10,11,12,13)]
    cur=-1 if t<s11-0.1 else 0 if t<s12-0.1 else 1 if t<s13-0.1 else 2
    xs=[220,0,-220]; ny=-320
    for i in range(3):
        kk=back(seg(t,s10+0.1+i*0.15,s10+0.5+i*0.15))
        if i<2 and kk>0: poly(x,[(xs[i]-60,ny),(xs[i+1]+60,ny)],GOLD if cur>i else SOFT,0.9 if cur>i else 0.35,fill=False,lw=6,close=False)
        if kk>0: num_node(x,xs[i],ny,50,i+1,1,i==cur,kk)
    g=170; y=230
    cliff(x,420,g,y,1,CARD); cliff(x,-420,-g,y,1,(0.16,0.3,0.6))
    starts=[s11,s12,s13]
    for j in range(3):
        kk=back(seg(t,starts[j]+0.3,starts[j]+0.8)); px=g-(2*g)*(j+0.5)/3
        if kk>0: rrect(x,px-g/3+3,y-4,2*g/3-6,30*kk,6,GOLD if j==cur else (0.62,0.48,0.32),1)
    cross=eio(seg(t,P[13]['words'][-2]['start'],ph(13)['end']+0.3))
    person(x,lerp(300,-300,cross),y-10,1.0,mix(WHITE,GOLD,cross),1)
    # step visuals above
    ca=[(s11,s12),(s12,s13),(s13,ph(13)['end']+1.5)]
    for j,(a0,a1) in enumerate(ca):
        k=back(seg(t,a0-0.1,a0+0.35)); ko=eio(seg(t,a1-0.2,a1+0.2)) if j<2 else 0
        if k<=0 or ko>=1: continue
        a=1-ko; cy=-60
        if j==0:
            mo=seg(t,P[11]['words'][-4]['start'] if len(P[11]['words'])>4 else a0+0.8,P[11]['words'][-1]['start']+0.3)
            icon_disc(x,0xe0f0,120,cy,80,k*(1-0.4*mo),a*(1-0.6*mo))
            glowline(x,partial([(30,cy),(-60,cy)],eo(mo)),GOLD,a,6)
            if mo>0: icon_disc(x,0xe2e6,-150,cy,90,back(mo),a,gold=True,glow=1)
        elif j==1:
            icon_disc(x,0xe536,180,cy,60,k,a)
            k2=back(seg(t,P[12]['words'][-3]['start']-0.1,P[12]['words'][-3]['start']+0.3)); icon_disc(x,0xe878,0,cy,70,k2,a,gold=True,glow=0.6)
            k3=back(seg(t,P[12]['words'][-2]['start']-0.1,P[12]['words'][-2]['start']+0.3)); icon_disc(x,0xe55f,-180,cy,70,k3,a,gold=True,glow=0.6)
        else:
            base=60; pts=[(220,base),(120,base-40),(20,base-30),(-80,base-110),(-200,base-170)]
            glowline(x,partial(pts,eo(seg(t,a0,a0+1.2))),GOLD2,a,7)
            for p in pts[:int(1+4*eo(seg(t,a0,a0+1.2)))]: circle(x,p[0],p[1],10,GOLD,a)
            magnifier(x,-200,-200,0.7*back(seg(t,P[13]['words'][-3]['start'],P[13]['words'][-3]['start']+0.4)),a)
def scene_5(x,t):
    s14,s15,s16=[ph(i)['start'] for i in (14,15,16)]
    kn=back(seg(t,s14-0.3,s14+0.2)); kno=eio(seg(t,s15-0.15,s15+0.3))
    if kn>0 and kno<1:
        a=1-kno; x.save(); x.translate(0,20); x.rotate(-0.05); x.scale(S_(kn),S_(kn))
        rrect(x,-170,-160,340,320,16,(0,0,0),0.25*a); rrect(x,-170,-170,340,320,16,GOLD2,a)
        for i in range(3): rrect(x,-120,-60+i*60,240-i*50,14,7,NAVY2,0.5*a)
        x.restore()
        kp=eo(seg(t,s14,s14+0.35)); py=lerp(-480,-150,kp)
        circle(x,0,py,34,(0.85,0.2,0.25),a); circle(x,-10,py-10,10,WHITE,0.6*a); poly(x,[(0,py+30),(0,py+70)],SOFT,a,fill=False,lw=6,close=False)
    k=seg(t,s15-0.2,s15+0.3)
    if k>0:
        a=eo(k); dim=eio(seg(t,s16+0.2,s16+0.9))
        kc=back(seg(t,wt(15,0)-0.1,wt(15,0)+0.4)); kb=back(seg(t,wt(15,3)-0.1,wt(15,3)+0.4))
        x.save(); x.translate(230,lerp(-60,-250,dim)); s=kc*lerp(1,0.6,dim); x.scale(S_(s),S_(s))
        rrect(x,-150,-110,300,220,14,WHITE,a*(1-0.4*dim)); rrect(x,-135,-95,270,190,8,GOLD,a*(1-0.4*dim),fill=False,lw=4)
        for j in range(3): rrect(x,-90,-50+j*34,180-j*40,12,6,SOFT,a*(1-0.4*dim))
        circle(x,90,80,34,GOLD,a*(1-0.4*dim)); x.restore()
        blit(x,sprite("الشهادة",36,(255,255,255),b'Bold'),230,lerp(110,-120,dim),kc,a*(1-0.5*dim))
        icon_disc(x,0xea19,-230,lerp(-60,-250,dim),lerp(110,66,dim),kb,a*(1-0.4*dim))
        blit(x,sprite("المعرفة",36,(255,255,255),b'Bold'),-230,lerp(110,-120,dim),kb,a*(1-0.5*dim))
        kp=seg(t,s16-0.1,s16+0.6)
        if kp>0:
            gear(x,-40,150,80*back(kp),t*1.5,GOLD,a); gear(x,70,210,52*back(kp),-t*2.3,WHITE,a,teeth=8)
            blit(x,sprite("الممارسة",44,(246,196,72),b'Black'),0,310,back(kp),a)
            kr=seg(t,P[16]['words'][-2]['start'],P[16]['words'][-1]['start']+1.5)
            for j in range(3):
                pr=((t-s16)*0.6+j/3)%1
                if kr>0: circle(x,0,180,90+pr*360,GOLD,(1-pr)*0.5*min(1,kr*3)*a,fill=False,lw=4)
def scene_6(x,t):
    s17,s18=ph(17)['start'],ph(18)['start']
    k=seg(t,s17-0.2,s17+0.4); ko=eio(seg(t,s18+0.4,s18+0.9))
    if k>0 and ko<1:
        a=1-ko; y=60
        glowline(x,[(380,y),(lerp(380,-120,eo(k)),y)],WHITE,a*0.9,8)
        for j in range(4): circle(x,380-j*160,y,14,GOLD if j==0 else WHITE,a)
        ke=back(seg(t,P[17]['words'][-1]['start']-0.1,P[17]['words'][-1]['start']+0.35))
        if ke>0:
            rrect(x,-140,y-70*ke,30,140*ke,8,GOLD,a)
            blit(x,sprite("محطة النهاية",38,(255,255,255),b'Bold'),-125,y+110,ke,a*0.9)
    kr=seg(t,s18-0.15,s18+0.5)
    if kr>0:
        fly=eio(seg(t,wt(18,2)-0.2,wt(18,2)+1.2)); ry=lerp(240,-180,fly)
        for j in range(8):
            u=j/8; circle(x,0+8*math.sin(t*20+j),ry+90+u*200*fly,(22-j*2)*eo(kr),GOLD2 if j%2 else GOLD,(1-u)*0.8)
        icon_disc(x,0xeb9b,0,ry,90,back(kr),1,gold=True,glow=1)
        ki=seg(t,P[18]['words'][-1]['start']-0.1,P[18]['words'][-1]['start']+1.2)
        if ki>0:
            pts=[(230*math.sin(u),ry-200+110*math.sin(u)*math.cos(u)) for u in np.linspace(0,2*math.pi,140)]
            glowline(x,pts,GOLD,1,10,k=eo(ki),close=True)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5,scene_6]
