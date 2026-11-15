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

def box(x,cx,cy,w,L,a,glow=0,bars=0,t=0):
    """open isometric box: top rim centred at (cx,cy), half-width w, wall height L"""
    d=w*0.5
    P1=(cx,cy-d); P2=(cx+w,cy); P3=(cx,cy+d); P4=(cx-w,cy)
    poly(x,[P4,P1,P2,P3],(0.05,0.1,0.25),a)             # inside floor shadow
    poly(x,[P4,P1,(P1[0],P1[1]+L*0.55),(P4[0],P4[1]+L*0.55)],(0.1,0.19,0.42),a)   # inner back-left
    poly(x,[P1,P2,(P2[0],P2[1]+L*0.55),(P1[0],P1[1]+L*0.55)],(0.08,0.16,0.36),a)  # inner back-right
    if glow>0:
        radial(x,cx,cy,w*1.9,GOLD,0.55*glow*a)
        for i in range(9):
            ang=-math.pi/2+(i-4)*0.16; r=w*(1.3+0.25*math.sin(t*3+i))
            poly(x,[(cx,cy),(cx+r*1.2*math.cos(ang),cy+r*math.sin(ang)*1.4)],GOLD2,0.25*glow*a,fill=False,lw=8,close=False)
    poly(x,[P4,P3,(P3[0],P3[1]+L),(P4[0],P4[1]+L)],(0.16,0.3,0.62),a)
    poly(x,[P3,P2,(P2[0],P2[1]+L),(P3[0],P3[1]+L)],(0.12,0.24,0.52),a)
    for e_ in ([P4,P1,P2,P3],[P4,(P4[0],P4[1]+L),(P3[0],P3[1]+L),(P2[0],P2[1]+L),P2],[P3,(P3[0],P3[1]+L)]):
        poly(x,e_,GOLD,a,fill=False,lw=5,close=(len(e_)==4))
    if bars>0:
        for j in range(5):
            bx=cx-w*0.8+j*w*0.4; poly(x,[(bx,cy-d*0.6),(bx,cy-d*0.6-180*bars)],SOFT,a*bars,fill=False,lw=8,close=False)
        poly(x,[(cx-w*0.9,cy-d*0.6-180*bars),(cx+w*0.9,cy-d*0.6-180*bars)],SOFT,a*bars,fill=False,lw=8,close=False)
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    k=back(seg(t,s0-0.45,s0+0.15))
    if k<=0: return
    back_in=eio(seg(t,wt(1,len(P[1]['words'])-2)-0.3,wt(1,len(P[1]['words'])-2)+0.6))
    glow=seg(t,wt(1,len(P[1]['words'])-2),wt(1,len(P[1]['words'])-2)+0.6)
    box(x,0,80,210*k,200*k,1,glow=glow,t=t)
    out=eo(seg(t,wt(0,3)-0.2,wt(0,3)+0.6))
    bx=lerp(0,300,out)*(1-back_in); by=lerp(40,-250,out)*(1-back_in)+lerp(0,40,back_in)*0
    by=lerp(lerp(40,-250,out),20,back_in)
    if out>0 and back_in<1:
        glowline(x,partial([(40,0),(300,-250)],out),WHITE,0.5*(1-back_in),4)
    icon_disc(x,0xe0f0,bx,by,60,k,1,gold=True,glow=0.7)
    kq=back(seg(t,ph(1)['end']-0.4,ph(1)['end']))
    if kq>0: q=sprite("؟",130,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,0,-280,kq*(1+0.05*math.sin(t*5)),1)
def scene_2(x,t):
    s2,s3=ph(2)['start'],ph(3)['start']
    k=back(seg(t,s2-0.4,s2+0.15))
    if k<=0: return
    bars=1-eio(seg(t,wt(2,2)+0.2,wt(2,2)+0.8))
    box(x,0,110,230*k,210*k,1,bars=bars*seg(t,s2-0.2,s2+0.3),glow=0.3*seg(t,s3,s3+0.8),t=t)
    if bars>0.05 and bars<1:
        for j in range(6):
            ang=j*1.05; r=(1-bars)*300; poly(x,star_pts(r*math.cos(ang),-80+r*0.6*math.sin(ang),10,rot=t),SOFT,bars)
    items=[(0xe8b5,"وقتك",4),(0xe1a1,"مواردك",5),(0xf233,"فريقك",6)]
    ex=eio(seg(t,s3-0.2,s3+0.4))
    for j,(ic,lab,wi) in enumerate(items):
        if wi>=len(P[2]['words']): wi=len(P[2]['words'])-1
        kk=back(seg(t,wt(2,wi)-0.1,wt(2,wi)+0.35))
        if kk<=0: continue
        px=lerp((1-j)*250,(1-j)*110,ex); py=lerp(-230,60,ex)
        icon_disc(x,ic,px,py,lerp(60,44,ex),kk,1,gold=(j==1),glow=0.4)
        blit(x,sprite(lab,34,(255,255,255),b'Bold'),px,py+lerp(90,0,ex),kk,1-ex)
    km=back(seg(t,wt(3,len(P[3]['words'])-3)-0.2,wt(3,len(P[3]['words'])-3)+0.3))
    if km>0:
        mx=lerp(-180,80,eio(seg(t,s3+0.6,ph(3)['end'])))
        magnifier(x,mx,-40,0.85*km,1)
        kg=seg(t,wt(3,len(P[3]['words'])-1)-0.1,wt(3,len(P[3]['words'])-1)+0.6)
        for j in range(5):
            gk=back(seg(kg,j*0.12,j*0.12+0.4))
            if gk>0: icon_disc(x,0xead5,-200+j*100,-260,34,gk,1,gold=True,glow=0.6)
def scene_3(x,t):
    s4=ph(4)['start']
    k=seg(t,s4-0.3,s4+0.3)
    if k<=0: return
    a=eo(k); focus=eio(seg(t,wt(4,len(P[4]['words'])-2)-0.3,wt(4,len(P[4]['words'])-2)+0.5))
    # funnel walls
    fw=eo(seg(t,wt(4,3)-0.3,wt(4,3)+0.3))
    if fw>0:
        glowline(x,partial([(-320,-260),(-60,40),(-60,140)],fw),GOLD,a,7); glowline(x,partial([(320,-260),(60,40),(60,140)],fw),GOLD,a,7)
    rng=np.random.default_rng(4)
    for j in range(40):
        ox,oy=rng.uniform(-330,330),rng.uniform(-300,-80); ph_=rng.uniform(0,6.28)
        u=((t-s4)*0.45+j/40)%1 if focus>0 else 0
        # scattered drift vs. channelled flow
        sx,sy=ox+20*math.sin(t*1.5+ph_),oy+20*math.cos(t*1.3+ph_)
        if u<0.55: fx,fy=lerp(ox,0,u/0.55)*(1-0.0),lerp(oy,40,u/0.55)
        else: fx,fy=0,lerp(40,330,(u-0.55)/0.45)
        px,py=lerp(sx,fx,focus),lerp(sy,fy,focus)
        circle(x,px,py,7 if focus<0.5 else 6,mix(SOFT,GOLD2,focus),a*0.9)
    if focus>0:
        radial(x,0,320,200,GOLD,0.5*focus*a)
        icon_disc(x,0xe65f,0,300,64,back(seg(t,wt(4,len(P[4]['words'])-1)-0.1,wt(4,len(P[4]['words'])-1)+0.4)),a,gold=True,glow=1)
    kx=seg(t,wt(4,2)-0.1,wt(4,2)+0.4)
    if kx>0 and focus<0.5:
        blit(x,sprite("القيود",40,(255,255,255),b'Bold'),0,-330,eo(kx),a*(1-focus*2))
def scene_4(x,t):
    s5,s6,s7,s8=[ph(i)['start'] for i in (5,6,7,8)]
    cur=-1 if t<s6-0.1 else 0 if t<s7-0.1 else 1 if t<s8-0.1 else 2
    xs=[220,0,-220]; ny=-320
    for i in range(3):
        kk=back(seg(t,wt(5,len(P[5]['words'])-2)-0.1+i*0.15,wt(5,len(P[5]['words'])-2)+0.3+i*0.15))
        if i<2 and kk>0: poly(x,[(xs[i]-60,ny),(xs[i+1]+60,ny)],GOLD if cur>i else SOFT,0.9 if cur>i else 0.35,fill=False,lw=6,close=False)
        if kk>0: num_node(x,xs[i],ny,50,i+1,1,i==cur,kk)
    kb=back(seg(t,s5-0.3,s5+0.2)); kbo=eio(seg(t,s6-0.2,s6+0.2))
    if kb>0 and kbo<1: box(x,0,80,200*kb*(1-kbo),190*kb*(1-kbo),1,glow=0.5,t=t)
    # step 1: define your box + resources
    k=seg(t,s6-0.2,s6+0.3); ko=eio(seg(t,s7-0.2,s7+0.25))
    if k>0 and ko<1:
        a=eo(k)*(1-ko); d=eo(seg(t,s6,s6+1.0))
        pts=[(-260,-170),(260,-170),(260,300),(-260,300),(-260,-170)]; pp=partial(pts,d)
        for j in range(len(pp)-1):
            p0,p1=pp[j],pp[j+1]; n=max(1,int(math.dist(p0,p1)/30))
            for q in range(0,n,2): poly(x,[(lerp(p0[0],p1[0],q/n),lerp(p0[1],p1[1],q/n)),(lerp(p0[0],p1[0],(q+1)/n),lerp(p0[1],p1[1],(q+1)/n))],GOLD,a,fill=False,lw=6,close=False)
        for j,ic in enumerate((0xe8b5,0xe1a1,0xf233,0xe0f0)):
            kk=back(seg(t,wt(6,len(P[6]['words'])-2)-0.2+j*0.15,wt(6,len(P[6]['words'])-2)+0.2+j*0.15))
            icon_disc(x,ic,-120+(j%2)*240,-10+(j//2)*190,64,kk,a,gold=(j==3))
    # step 2: rearrange tiles -> idea
    k=seg(t,s7-0.2,s7+0.3); ko=eio(seg(t,s8-0.2,s8+0.25))
    if k>0 and ko<1:
        a=eo(k)*(1-ko); sh=eio(seg(t,wt(7,3)-0.2,wt(7,3)+0.9))
        order=[5,2,7,0,8,3,6,1,4]
        for j in range(9):
            src=order[j]; sx,sy=(src%3-1)*150,(src//3-1)*150+60; dx,dy=(j%3-1)*150,(j//3-1)*150+60
            px,py=lerp(sx,dx,sh),lerp(sy,dy,sh)-60*math.sin(sh*math.pi)*(1 if j%2 else -1)
            gold=(j==4 and sh>=1)
            rrect(x,px-62,py-62,124,124,18,GOLD if gold else CARD,a); rrect(x,px-62,py-62,124,124,18,WHITE,a*0.4,fill=False,lw=3)
            if gold: radial(x,px,py,200,GOLD,0.4*a); blit(x,mi(0xe0f0,80,(14,32,80)),px,py,back(seg(t,wt(7,3)+0.9,wt(7,3)+1.3)),a)
            else: blit(x,mi(0xe87b,50,(150,170,215)),px,py,1,a*0.7)
    # step 3: what I have vs what I lack
    k=seg(t,s8-0.2,s8+0.3)
    if k>0:
        a=eo(k)
        rrect(x,40,-200,240,440,26,CARD,a); rrect(x,40,-200,240,440,26,GOLD,a,fill=False,lw=4)
        blit(x,sprite("ما أملك",36,(246,196,72),b'Black'),160,-240,1,a)
        for j,ic in enumerate((0xe8b5,0xe1a1,0xf233)):
            kk=back(seg(t,wt(8,len(P[8]['words'])-4)+j*0.15,wt(8,len(P[8]['words'])-4)+0.35+j*0.15)); icon_disc(x,ic,160,-110+j*120,48,kk,a,gold=True)
        dim=seg(t,wt(8,len(P[8]['words'])-1)-0.1,wt(8,len(P[8]['words'])-1)+0.4)
        rrect(x,-280,-200,240,440,26,SOFT,a*0.4,fill=False,lw=3)
        blit(x,sprite("ما ينقصني",34,(200,210,235),b'Bold'),-160,-240,1,a*0.8)
        for j in range(3): circle(x,-160,-110+j*120,48,SOFT,a*0.25*(1-dim*0.6),fill=False,lw=3)
        if dim>0: ban(x,-160,20,140,dim,a*0.8,lw=9)
        kp=back(seg(t,ph(8)['end']-0.4,ph(8)['end']+0.2))
        if kp>0:
            glowline(x,[(160,250),(160,330)],GOLD2,a,6); poly(x,star_pts(160,370,46*kp,rot=-math.pi/2+t*0.4),GOLD2,a)
def scene_5(x,t):
    s9,s10=ph(9)['start'],ph(10)['start']
    kn=back(seg(t,s9-0.3,s9+0.2)); kno=eio(seg(t,s10-0.15,s10+0.3))
    if kn>0 and kno<1:
        a=1-kno; x.save(); x.translate(0,20); x.rotate(-0.05); x.scale(S_(kn),S_(kn))
        rrect(x,-170,-160,340,320,16,(0,0,0),0.25*a); rrect(x,-170,-170,340,320,16,GOLD2,a)
        for i in range(3): rrect(x,-120,-60+i*60,240-i*50,14,7,NAVY2,0.5*a)
        x.restore()
        kp=eo(seg(t,s9,s9+0.35)); py=lerp(-480,-150,kp)
        circle(x,0,py,34,(0.85,0.2,0.25),a); circle(x,-10,py-10,10,WHITE,0.6*a); poly(x,[(0,py+30),(0,py+70)],SOFT,a,fill=False,lw=6,close=False)
    k=seg(t,s10-0.2,s10+0.3)
    if k>0:
        a=eo(k); run=eo(seg(t,wt(10,3)-0.3,wt(10,3)+0.5)); inside=eio(seg(t,wt(10,len(P[10]['words'])-2)-0.2,wt(10,len(P[10]['words'])-2)+0.6))
        box(x,0,150,220,200,a,glow=inside,t=t)
        if inside<1:
            px=lerp(0,-330,run)*(1-inside); person(x,px,lerp(110,90,run),1.3,SOFT,a*(1-inside))
            if run>0.6: ban(x,-300,40,120,seg(t,wt(10,3)+0.3,wt(10,3)+0.8),a*(1-inside),lw=9)
        if inside>0:
            person(x,0,lerp(-120,120,inside),1.4,GOLD,a*inside)
            for j in range(8):
                ang=t*0.8+j*0.785; r=170+15*math.sin(t*3+j); poly(x,star_pts(r*math.cos(ang),60+r*0.6*math.sin(ang),12*inside,rot=t),GOLD2,a)
def scene_6(x,t):
    s11=ph(11)['start']
    k=back(seg(t,s11-0.4,s11+0.2))
    if k<=0: return
    fin=eio(seg(t,VEND+0.1,VEND+0.8))
    rise=eio(seg(t,wt(11,len(P[11]['words'])-4)-0.2,wt(11,len(P[11]['words'])-1)+0.4))
    x.save(); x.translate(0,lerp(0,-150,fin)); s=lerp(1,0.7,fin); x.scale(s,s)
    box(x,0,200,240*k,210*k,1,glow=0.3+0.7*rise,t=t)
    icon_disc(x,0xe0f0,0,lerp(200,-150,rise),lerp(40,110,rise),back(seg(t,s11,s11+0.5)),1,gold=True,glow=rise)
    x.restore()
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        l1=sprite("لا تهرب من الصندوق…",68,(255,255,255),b'Black'); l2=sprite("أبدِع فيه",96,(246,196,72),b'Black',glow=(245,180,40))
        e2=eo(seg(t,VEND+0.7,VEND+1.4))
        blit(x,l1,0,250+(1-fin)*30,1,a); blit(x,l2,0,370+(1-e2)*30,1+0.03*math.sin(t*3),a*e2)
        glowline(x,[(-240*e2,450),(240*e2,450)],GOLD,a*0.7,4)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5,scene_6]
