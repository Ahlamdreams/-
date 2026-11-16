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
def scene_1(x,t):
    s0,s1,s2=[ph(i)['start'] for i in range(3)]
    k=back(seg(t,s0-0.45,s0+0.15)); ko=eio(seg(t,s2-0.2,s2+0.3))
    if k>0 and ko<1:
        a=1-ko
        book(x,-60,-150,0.9*k,a); icon_disc(x,0xebcc,190,-170,70,k,a,gold=True)
        v1=0.42*eo(seg(t,s0,s0+0.8)); v2=0.88*eo(seg(t,s0+0.2,s0+1.0))
        bar(x,60,"المنهج",v1,WHITE,a,k); bar(x,150,"الوقت",v2,GOLD,a,k)
        kh=back(seg(t,s1-0.15,s1+0.3))
        if kh>0:
            clock(x,0,300,90*kh,t*6,a); circle(x,0,300,100*kh,GOLD,a,fill=False,lw=5)
    k=seg(t,s2-0.2,s2+0.3)
    if k>0:
        a=eo(k)
        for i in range(5):
            kk=back(seg(t,wt(2,1)-0.1+i*0.12,wt(2,1)+0.3+i*0.12))
            num_node(x,(2-i)*150,-40,62,i+1,a,True,kk)
            if kk>0: blit(x,mi([0xe55b,0xe8e5,0xe1c4,0xf233,0xe425][i],52,(255,255,255)),(2-i)*150,110,kk,a)
        kg=back(seg(t,wt(2,2)-0.1,wt(2,2)+0.4))
        if kg>0: blit(x,sprite("أفكار عملية",54,(246,196,72),b'Black',glow=(245,180,40)),0,240,kg,a)
def scene_2(x,t):
    s3,s4=ph(3)['start'],ph(4)['start']
    top5(x,t,0,eo(seg(t,s3-0.4,s3)))
    k=seg(t,s3-0.2,s3+0.3)
    if k<=0: return
    a=eo(k); dist=eio(seg(t,P[4]['words'][-1]['start']-0.5,ph(4)['end']+0.3))
    kc=eo(seg(t,s4,s4+0.6))
    for w in range(4):
        cx=270-w*180
        if kc>0:
            rrect(x,cx-75,-40,150,360,20,CARD,a*kc*0.8); rrect(x,cx-75,-40,150,360,20,GOLD,a*kc,fill=False,lw=3)
            blit(x,sprite("أسبوع "+'١٢٣٤'[w],28,(246,196,72),b'Bold'),cx,-75,kc,a)
    for j in range(12):
        kk=back(seg(t,s3+0.1+j*0.06,s3+0.35+j*0.06))
        if kk<=0: continue
        gx,gy=(j%4)*170-255,(j//4)*110-200+60
        w=j%4; tx,ty=270-w*180,10+(j//4)*100
        px,py=lerp(-gx,tx,dist),lerp(gy,ty,dist)
        rrect(x,px-62*kk,py-38*kk,124*kk,76*kk,14,GOLD if j%4==0 and dist>0.9 else WHITE,a)
        blit(x,mi(0xea19,40,(14,32,80)),px,py,kk,a)
def scene_3(x,t):
    s5,s6=ph(5)['start'],ph(6)['start']
    top5(x,t,1)
    k=seg(t,s5-0.2,s5+0.3); ko=eio(seg(t,s6+0.1,s6+0.5))
    if k>0 and ko<1:
        a=eo(k)*(1-ko); srt=eio(seg(t,P[5]['words'][-1]['start']-0.2,P[5]['words'][-1]['start']+0.7))
        L=[230,420,160,330,280]; order=sorted(range(5),key=lambda i:-L[i])
        for i in range(5):
            r0=i; r1=order.index(i); y=lerp(-200+r0*95,-200+r1*95,srt); w=L[i]
            gold=r1==0 and srt>0.9
            rrect(x,240-w,y-30,w,60,18,GOLD if gold else (WHITE if r1<2 and srt>0.9 else CARD),a); 
            if gold: blit(x,sprite("أساسي",30,(14,32,80),b'Black'),240-w/2,y-2,1,a)
    k=seg(t,s6-0.1,s6+0.4)
    if k>0:
        a=eo(k); m=eio(seg(t,P[6]['words'][-3]['start']-0.2,P[6]['words'][-1]['start']+0.2))
        icon_disc(x,0xe2e6,0,-150,80,back(k),a,gold=True,glow=0.6); blit(x,sprite("الأساسيات",38,(246,196,72),b'Black'),0,-40,back(k),a)
        for sg in (1,-1):
            cx=lerp(sg*170,0,m)
            rrect(x,cx-110,90,220,150,18,GOLD if m>=1 else WHITE,a*(1 if m<1 or sg>0 else 0))
            blit(x,mi(0xea19,60,(14,32,80)),cx,150,1,a*(1 if m<1 or sg>0 else 0))
        if m>=1:
            blit(x,sprite("درس واحد",34,(14,32,80),b'Black'),0,215,back(seg(t,ph(6)['end']-0.6,ph(6)['end'])),a)
            for j in range(6):
                ang=t+j*1.05; poly(x,star_pts(170*math.cos(ang),165+100*math.sin(ang),9,rot=t),GOLD2,a)
def scene_4(x,t):
    s7,s8=ph(7)['start'],ph(8)['start']
    top5(x,t,2)
    k=back(seg(t,s7-0.3,s7+0.2)); ko=eio(seg(t,s8-0.1,s8+0.3))
    if k>0 and ko<1:
        fl=seg(t,P[7]['words'][-2]['start'],P[7]['words'][-2]['start']+0.8); sx=abs(math.cos(fl*math.pi))
        x.save(); x.translate(0,20); x.scale(S_(sx),1); x.translate(0,-20)
        if fl<0.5: board(x,0,20,1.2*k,1-ko)
        else: phone(x,0,30,1.0*k,1-ko)
        x.restore()
        if fl>=0.5: blit(x,mi(0xe1c4,110,(246,196,72)),0,20,sx,1-ko)
    k=seg(t,s8-0.1,s8+0.4)
    if k>0:
        a=eo(k)
        k1=back(seg(t,P[8]['words'][2]['start']-0.2,P[8]['words'][2]['start']+0.3))
        icon_disc(x,0xe88a,200,-170,56,k1,a); phone(x,200,40,0.6*k1,a); blit(x,mi(0xe1c4,60,(246,196,72)),200,40,k1,a)
        blit(x,sprite("قبل الحصة",34,(255,255,255),b'Bold'),200,190,k1,a)
        k2=back(seg(t,P[8]['words'][-1]['start']-0.3,P[8]['words'][-1]['start']+0.3))
        if k2>0:
            glowline(x,[(90,40),(-80,40)],GOLD,a,6); arrow_shape(x,-60,40,-90,40,GOLD,a,lw=1,head=24)
            for j in range(4):
                px=-280+(j%2)*120; py=-60+(j//2)*150; person(x,px,py+40,0.85*k2,GOLD if j%2==0 else WHITE,a)
                icon_disc(x,[0xea3d,0xe0f0,0xe913,0xe2e6][j],px+50,py-55,26,k2,a,gold=True)
            blit(x,sprite("الحصة للتطبيق",36,(246,196,72),b'Black'),-220,230,k2,a)
def scene_5(x,t):
    s9,s10=ph(9)['start'],ph(10)['start']
    top5(x,t,3)
    k=seg(t,s9-0.2,s9+0.3)
    if k<=0: return
    a=eo(k); asm=eio(seg(t,P[10]['words'][-2]['start']-0.3,ph(10)['end']+0.2))
    corners=[(230,-180),(-230,-180),(230,230),(-230,230)]; cols=[GOLD,(0.35,0.55,0.95),WHITE,GOLD2]
    pieces=[(-60,-60),(60,-60),(-60,60),(60,60)]
    for i,(cx,cy) in enumerate(corners):
        kk=back(seg(t,s9+0.1+i*0.15,s9+0.5+i*0.15))
        if kk<=0: continue
        for q,dx in enumerate((-45,0,45)): person(x,cx+dx,cy+60-(12 if q==1 else 0),0.6*kk,SOFT,a*(1-0.5*asm))
        pc=back(seg(t,s10-0.1+i*0.12,s10+0.3+i*0.12))
        if pc>0:
            px,py=lerp(cx,pieces[i][0],asm),lerp(cy-70,pieces[i][1]+40,asm)
            rrect(x,px-56*pc,py-56*pc,112*pc,112*pc,14,cols[i],a); blit(x,mi(0xe87b,50,(14,32,80)),px,py,pc,a)
    if asm>=1:
        kc=seg(t,ph(10)['end']+0.2,ph(10)['end']+0.6); radial(x,0,40,260,GOLD,0.4*kc*a)
        check(x,0,40,3*back(kc),eo(kc),NAVY1,a,lw=14)
def scene_6(x,t):
    s11,s12=ph(11)['start'],ph(12)['start']
    top5(x,t,4)
    k=back(seg(t,s11-0.3,s11+0.2))
    if k<=0: return
    shr=eio(seg(t,P[12]['words'][2]['start']-0.2,P[12]['words'][3]['start']+0.4))
    x.save(); x.translate(lerp(0,170,shr),lerp(40,-40,shr)); s=lerp(1.25,0.55,shr)*k; x.scale(S_(s),S_(s))
    if shr<0.95: paper(x,0,0,1,1)
    else:
        rrect(x,-120,-80,240,160,18,WHITE,1); rrect(x,-120,-80,240,46,18,GOLD,1)
        for j in range(2): rrect(x,-90,-10+j*40,180-j*50,14,7,SOFT,1)
    x.restore()
    if shr>=0.95: blit(x,sprite("بطاقة خروج",36,(246,196,72),b'Black'),170,60,1,1)
    kt=seg(t,P[12]['words'][-2]['start']-0.2,ph(12)['end']+1.0)
    if kt>0:
        a=eo(min(1,kt*4)); cx,cy=-200,-20
        circle(x,cx,cy,110,CARD,a); x.new_path(); x.arc(cx,cy,110,-math.pi/2,-math.pi/2+2*math.pi*max(0.001,1-kt)); x.set_line_width(14); rgba(GOLD,a,x); x.stroke()
        rem=int((1-kt)*120); sp=sprite(f"{rem//60}:{rem%60:02d}",56,(255,255,255),b'Black',rtl=False); blit(x,sp,cx,cy-4,1,a)
        blit(x,sprite("دقيقتان",32,(255,255,255),b'Bold'),cx,cy+150,1,a)
def scene_7(x,t):
    s13,s14,s15,s16=[ph(i)['start'] for i in (13,14,15,16)]
    pinnote(x,t,s13,s14)
    k=seg(t,s14-0.2,s14+0.3); ko=eio(seg(t,s15-0.15,s15+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        for j in range(7):
            ang=-0.6+j*0.2+0.15*math.sin(t*6+j); x.save(); x.translate(0,120); x.rotate(ang)
            rrect(x,-6,-260,160,240,10,WHITE if j%2 else (0.9,0.92,0.97),a); x.restore()
        ban(x,0,0,260,seg(t,P[14]['words'][3]['start']-0.1,P[14]['words'][3]['start']+0.5),a,lw=13)
    k=seg(t,s15-0.15,s15+0.4); ko=eio(seg(t,s16-0.15,s16+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko); icon_disc(x,0xea4a,0,-40,140,back(k),a,gold=True,glow=1)
        kc=back(seg(t,P[15]['words'][-2]['start']-0.1,P[15]['words'][-2]['start']+0.4))
        if kc>0: circle(x,120,80,46*kc,WHITE,a); check(x,120,78,1.3*kc,1,NAVY1,a,lw=8)
    fin=eio(seg(t,VEND+0.1,VEND+0.8))
    k=seg(t,s16-0.15,s16+0.4)
    if k>0 and fin<1:
        a=eo(k)*(1-fin)
        icon_disc(x,0xe0bf,0,-60,120,back(k),a,gold=True,glow=1)
        for j,(ic,px,py) in enumerate(((0xe87d,-230,-200),(0xe8dc,230,-180),(0xe0f0,-250,90),(0xe0bf,240,110))):
            kk=back(seg(t,s16+0.3+j*0.2,s16+0.7+j*0.2)); icon_disc(x,ic,px,py+8*math.sin(t*3+j),46,kk,a,gold=(j%2==1))
        kd=seg(t,P[16]['words'][-1]['start']-0.3,VEND+1)
        if kd>0:
            for j in range(3): blit(x,mi(0xe5db,70,(246,196,72)),0,170+j*55+10*math.sin(t*5),1,a*((math.sin(t*5-j)+1)/2))
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        for j,(w,c) in enumerate((("خطّط…",(255,255,255)),("ركّز…",(255,255,255)),("أنجِز",(246,196,72)))):
            e=eo(seg(t,VEND+0.2+j*0.45,VEND+0.8+j*0.45))
            sp=sprite(w,110 if j==2 else 84,c,b'Black',glow=(245,180,40) if j==2 else None); blit(x,sp,0,-180+j*170+(1-e)*40,1+0.03*math.sin(t*3)*(j==2),a*e)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5,scene_6,scene_7]
