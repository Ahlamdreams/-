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
