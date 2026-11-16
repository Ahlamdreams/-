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

PARTS=[("الدور",0xe7fd,"أنت معلّم علوم للصف الخامس",GOLD),("المهمة",0xe2e6,"اكتب ١٠ أسئلة عن دورة الماء",(0.45,0.65,1.0)),("المستويات",0xe8e5,"من التذكّر إلى التحليل والتطبيق",GOLD2),("الشكل",0xe6b1,"اختيار من متعدد مع الإجابة",(0.85,0.9,1.0))]
def tag(x,label,col,cx,cy,k,a,fs=28):
    sp=sprite(label,fs,(14,32,80),b'Black'); w=sp['adv']+34; h=fs+20
    x.save(); x.translate(cx,cy); x.scale(S_(k),S_(k)); rrect(x,-w/2,-h/2,w,h,h/2,col,a); blit(x,sp,0,-2,1,a); x.restore(); return w
def typed(x,sp,rx,cy,prog,a):
    """reveal sprite right-to-left, right edge at rx"""
    if prog<=0: return
    w=sp['adv']; x.save(); x.new_path(); x.rectangle(rx-w*prog-6,cy-40,w*prog+12,80); x.clip(); blit(x,sp,rx-w/2,cy,1,a); x.restore()
    if prog<1: rrect(x,rx-w*prog-10,cy-20,4,40,2,GOLD2,a*(0.5+0.5*math.sin(t_now[0]*12)))
t_now=[0]
def window(x,a,t):
    rrect(x,-410,-330,820,700,36,(0,0,0),0.25*a); rrect(x,-410,-340,820,700,36,(0.08,0.15,0.34),a)
    rrect(x,-410,-340,820,90,36,(0.12,0.24,0.52),a); rrect(x,-410,-290,820,40,0,(0.12,0.24,0.52),a)
    rrect(x,-410,-340,820,700,36,GOLD,a,fill=False,lw=3)
    icon_disc(x,0xf06c,350,-295,30,1,a,gold=True)
    blit(x,sprite("مساعد الذكاء الاصطناعي",30,(255,255,255),b'Bold'),180,-297,1,a)
    for j in range(3): circle(x,-370+j*28,-295,8,(0.4,0.5,0.75),a)
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    k=back(seg(t,s0-0.45,s0+0.15)); ko=eio(seg(t,s1-0.15,s1+0.35))
    if k>0 and ko<1:
        a=1-ko; cx,cy=0,-60
        circle(x,cx,cy,170*k,CARD,a); x.new_path(); x.arc(cx,cy,170*k,-math.pi/2,-math.pi/2+2*math.pi*max(0.001,1-seg(t,s0,ph(0)['end']))); x.set_line_width(16); rgba(GOLD,a,x); x.stroke()
        sec=int(60*(1-seg(t,s0,ph(0)['end']))); blit(x,sprite(f"{sec}",110,(255,255,255),b'Black',rtl=False),cx,cy-8,k,a)
        blit(x,sprite("ثانية",34,(246,196,72),b'Bold'),cx,cy+80,k,a)
        for j in range(6):
            kk=back(seg(t,wt(0,1)+j*0.12,wt(0,1)+0.35+j*0.12)); ang=-math.pi*0.95+j*math.pi*0.38
            if kk>0:
                px,py=cx+300*math.cos(ang)*0.95,cy+260*math.sin(ang)*0.9+40
                rrect(x,px-50*kk,py-34*kk,100*kk,68*kk,14,WHITE if j%2 else GOLD,a)
                blit(x,sprite("؟",40,(14,32,80),b'Black'),px,py-2,kk,a)
    k=seg(t,s1-0.2,s1+0.4)
    if k>0:
        a=eo(k); window(x,a*0.5,t)
        icon_disc(x,0xf06c,0,-30,130,back(k),a,gold=True,glow=1)
        kq=back(seg(t,P[1]['words'][-1]['start']-0.1,P[1]['words'][-1]['start']+0.35))
        if kq>0: rrect(x,-300,160,600,80,40,WHITE,a*kq); blit(x,sprite("كيف تطلب؟",40,(14,32,80),b'Black'),0,198,kq,a)
def scene_2(x,t):
    s2=ph(2)['start']
    k=seg(t,s2-0.2,s2+0.3)
    if k<=0: return
    a=eo(k)
    for i,(lab,ic,_,col) in enumerate(PARTS):
        kk=back(seg(t,s2+0.3+i*0.3,s2+0.65+i*0.3)); y=-290+i*130
        if kk<=0: continue
        x.save(); x.translate(0,y); x.scale(S_(kk),S_(kk))
        rrect(x,-230,-46,460,92,46,CARD,a); rrect(x,-230,-46,460,92,46,col,a,fill=False,lw=4)
        icon_disc(x,ic,180,0,34,1,a,gold=True); blit(x,sprite(lab,40,(255,255,255),b'Black'),-20,-3,1,a)
        x.restore()
        if i<3: blit(x,sprite("+",50,(246,196,72),b'Black',rtl=False),0,y+66,kk,a)
    ke=back(seg(t,s2+1.6,s2+2.0))
    if ke>0:
        blit(x,sprite("=",60,(246,196,72),b'Black',rtl=False),0,250,ke,a)
        x.save(); x.translate(0,340); x.scale(S_(ke),S_(ke)); radial(x,0,0,260,GOLD,0.4*a)
        rrect(x,-210,-48,420,96,48,GOLD,a); blit(x,sprite("طلب ذكي",44,(14,32,80),b'Black'),20,-3,1,a); blit(x,mi(0xe65f,46,(14,32,80)),-150,0,1,a); x.restore()
def scene_3(x,t):
    t_now[0]=t
    s3,s4,s5,s6,s7=[ph(i)['start'] for i in (3,4,5,6,7)]
    k=back(seg(t,s3-0.4,s3+0.1))
    if k<=0: return
    x.save(); x.scale(S_(k),S_(k)); a=1
    window(x,a,t)
    sent=eio(seg(t,s7+0.1,s7+0.6))
    # composer
    starts=[s3,s4,s5,s6]
    nlines=sum(1 for s in starts if t>=s-0.2)
    if sent<1:
        top=lerp(220-max(1,nlines)*95,-210,sent); bot=330
        rrect(x,-380,top-20,760,bot-top+10,26,(0.05,0.1,0.25),a*(1-sent))
        for i,st in enumerate(starts):
            kk=back(seg(t,st-0.2,st+0.2))
            if kk<=0: continue
            lab,ic,txt,col=PARTS[i]; y=top+30+i*95
            tag(x,lab,col,290,y,kk,a*(1-sent))
            pr=seg(t,P[3+i]['words'][1]['start'],ph(3+i)['end']-0.1)
            sp=sprite(txt,32,(255,255,255),b'Bold'); typed(x,sp,200,y,eo(pr),a*(1-sent))
        # send button
        pulse=seg(t,s7-0.1,s7+0.3)
        circle(x,-330,290,36*(1+0.2*math.sin(pulse*math.pi)),GOLD,a*(1-sent)); blit(x,mi(0xe163,40,(14,32,80)),-330,290,1,a*(1-sent))
    if sent>0:
        # user bubble
        rrect(x,-40,-230,400,80,24,GOLD,a*sent); blit(x,sprite("الدور + المهمة + المستويات + الشكل",24,(14,32,80),b'Bold'),160,-192,1,a*sent)
        lv=[("تذكّر",WHITE),("فهم",(0.6,0.75,1.0)),("تطبيق",GOLD2),("تحليل",GOLD)]
        kr=seg(t,s7+0.6,ph(7)['end']+0.2)
        if kr<0.15:
            for j in range(3): circle(x,-320+j*30,-100,9,WHITE,a*(0.4+0.6*abs(math.sin(t*6+j))))
        else:
            rrect(x,-380,-140,560,460,24,(0.12,0.24,0.52),a)
            for j,(lab,col) in enumerate(lv):
                kk=back(seg(kr,0.15+j*0.2,0.45+j*0.2))
                if kk<=0: continue
                y=-90+j*105
                x.save(); x.translate(0,y); x.scale(S_(kk),1)
                tag(x,lab,col,110,0,1,a,fs=26)
                rrect(x,-350,-14,380,12,6,SOFT,a); rrect(x,-350,6,300,12,6,(0.45,0.55,0.8),a)
                x.restore()
            kl=seg(t,P[7]['words'][-1]['start']-0.2,P[7]['words'][-1]['start']+0.6)
            if kl>0:
                for j in range(4):
                    yy=-90+j*105; poly(x,[(220,yy+40),(240,yy-55)],GOLD,a*kl,fill=False,lw=5,close=False)
                arrow_shape(x,300,320,300,lerp(320,-120,eo(kl)),GOLD2,a,lw=7,head=28)
    x.restore()
def scene_4(x,t):
    s8=ph(8)['start']
    k=seg(t,s8-0.25,s8+0.3)
    if k<=0: return
    a=eo(k)
    for j in range(4):
        kk=back(seg(t,s8+j*0.12,s8+0.35+j*0.12)); y=-230+j*140
        rrect(x,-320,y-50,640,100,20,WHITE,a*0.95*kk)
        rrect(x,-160,y-26,400,14,7,SOFT,a*kk); rrect(x,-160,y+6,300,14,7,(0.75,0.8,0.9),a*kk)
        ck=seg(t,P[8]['words'][1]['start']+j*0.3,P[8]['words'][1]['start']+0.3+j*0.3)
        circle(x,-260,y,30*kk,GOLD if ck>0.5 else (0.85,0.88,0.95),a); check(x,-260,y-2,0.8,eo(ck),NAVY1,a,lw=6)
        if j==2:
            ke=seg(t,P[8]['words'][2]['start'],P[8]['words'][2]['start']+0.8)
            if ke>0: blit(x,mi(0xe3c9,56,(246,196,72)),lerp(260,120,eo(ke)),y,1,a)
    u=seg(t,s8,ph(8)['end']); my=lerp(-260,240,eio(u))
    magnifier(x,180+40*math.sin(t*2),my,0.6,a)
    kx=back(seg(t,P[8]['words'][-2]['start']-0.1,P[8]['words'][-2]['start']+0.4))
    if kx>0: icon_disc(x,0xe7fd,0,-360,52,kx,a,gold=True,glow=0.8); blit(x,sprite("أنت الخبير",34,(246,196,72),b'Black'),150,-360,kx,a)
def scene_5(x,t):
    s9,s10=ph(9)['start'],ph(10)['start']
    k=seg(t,s9-0.2,s9+0.3); ko=eio(seg(t,s10-0.15,s10+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        person(x,120,150,2.0,GOLD,a); icon_disc(x,0xf06c,-170,60,70,back(k),a)
        kc=back(seg(t,P[9]['words'][-3]['start']-0.1,P[9]['words'][-3]['start']+0.4))
        if kc>0:
            clock(x,120,-190,80*kc,t*0.8,a); circle(x,120,-190,92*kc,GOLD,a,fill=False,lw=5)
            for j in range(3): person(x,-250+j*90,260,0.6*kc,WHITE,a); circle(x,-250+j*90,170,10*kc,(0.95,0.4,0.5),a)
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
        blit(x,sprite("معادلة الطلب الذكي",62,(246,196,72),b'Black',glow=(245,180,40)),0,-340,1,a)
        for i,(lab,ic,_,col) in enumerate(PARTS):
            e=back(seg(t,VEND+0.3+i*0.25,VEND+0.7+i*0.25)); y=-200+i*125
            if e<=0: continue
            x.save(); x.translate(0,y); x.scale(S_(e),S_(e))
            rrect(x,-230,-46,460,92,46,CARD,a); rrect(x,-230,-46,460,92,46,col,a,fill=False,lw=4)
            icon_disc(x,ic,180,0,34,1,a,gold=True); blit(x,sprite(lab,40,(255,255,255),b'Black'),-20,-3,1,a); x.restore()
            if i<3: blit(x,sprite("+",44,(246,196,72),b'Black',rtl=False),0,y+62,e,a)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
def OVERLAY(x,t): pass
