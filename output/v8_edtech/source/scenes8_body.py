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
