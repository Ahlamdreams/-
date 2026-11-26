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

LV=[("داعم",(0.35,0.78,0.45),"خطوات + أمثلة محلولة",0xe896),("أساسي",(0.98,0.8,0.25),"تمارين مباشرة",0xe2e6),("متقدّم",(0.35,0.6,1.0),"مسائل تتحدّى التفكير",0xe0f0)]
def sheet(x,cx,cy,s,col,lab,a,rot=0,lines=4,icon=None):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.rotate(rot); x.scale(S_(s),S_(s))
    rrect(x,-110,-140,220,280,16,(0,0,0),0.25*a); rrect(x,-110,-150,220,280,16,WHITE,a)
    rrect(x,-110,-150,220,56,16,col,a); rrect(x,-110,-110,220,16,0,col,a)
    if lab: blit(x,sprite(lab,30,(14,32,80),b'Black'),0,-124,1,a)
    for j in range(lines): rrect(x,-80,-60+j*40,lerp(160,90,(j%3)/2),12,6,(0.78,0.82,0.9),a)
    if icon: blit(x,mi(icon,40,(14,32,80)),70,95,1,a)
    x.restore()
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    k=seg(t,s0-0.3,s0+0.3); ko=eio(seg(t,s1-0.15,s1+0.35))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        for i,(lab,col,_,_) in enumerate(LV):
            kk=back(seg(t,s0+0.2+i*0.2,s0+0.6+i*0.2)); px=(1-i)*230; h=[90,170,250][i]
            if kk<=0: continue
            rrect(x,px-80,320-h*kk,160,h*kk,14,CARD,a); rrect(x,px-80,320-h*kk,160,10,5,col,a)
            person(x,px,320-h*kk-10,1.1*kk,col,a)
        kw=back(seg(t,P[0]['words'][-3]['start']-0.1,P[0]['words'][-3]['start']+0.4))
        if kw>0:
            sheet(x,0,-210,0.85*kw,(0.7,0.74,0.85),"ورقة واحدة",a)
            for i in range(3):
                px=(1-i)*230; poly(x,[(0,-80),(px,320-[90,170,250][i]-130)],SOFT,a*0.5*kw,fill=False,lw=3,close=False)
            kq=back(seg(t,ph(0)['end']-0.5,ph(0)['end']))
            if kq>0: blit(x,sprite("؟",120,(246,196,72),b'Black',glow=(245,190,60)),230,-260,kq,a)
    k=seg(t,s1-0.2,s1+0.4)
    if k>0:
        a=eo(k)
        icon_disc(x,0xf06c,0,-240,90,back(k),a,gold=True,glow=1)
        for i,(lab,col,_,_) in enumerate(LV):
            kk=back(seg(t,P[1]['words'][2]['start']-0.2+i*0.15,P[1]['words'][2]['start']+0.2+i*0.15)); px=(1-i)*230
            if kk>0:
                glowline(x,partial([(0,-150),(px,0)],eo(kk)),GOLD,a*0.6,3)
                sheet(x,px,110,0.75*kk,col,lab,a,rot=(1-i)*0.05)
        kt=back(seg(t,P[1]['words'][-1]['start']-0.1,P[1]['words'][-1]['start']+0.4))
        if kt>0:
            clock(x,260,-260,56*kt,t*4,a); circle(x,260,-260,64*kt,GOLD,a,fill=False,lw=4)
def scene_2(x,t):
    t_now[0]=t
    s2=ph(2)['start']; s8=ph(8)['start']
    k=back(seg(t,s2-0.4,s2+0.1))
    if k<=0: return
    x.save(); x.scale(S_(k),S_(k)); a=1
    window(x,a,t)
    kc=back(seg(t,s2,s2+0.4))
    if kc>0:
        rrect(x,-380,-240,760,520,26,(0.05,0.1,0.25),a)
        kk=eo(seg(t,P[2]['words'][0]['start'],P[2]['words'][0]['start']+0.4))
        blit(x,mi(0xe14d,40,(246,196,72)),-340,-205,kk,a)
    rows=[(3,None,"أنت معلّم رياضيات للصف الرابع",WHITE,-180),(4,None,"اكتب ورقة عمل عن الكسور بثلاثة مستويات",WHITE,-110)]
    for i,(pi,_,txt,col,y) in enumerate(rows):
        if t<ph(pi)['start']-0.2: continue
        pr=seg(t,ph(pi)['start'],ph(pi)['end']-0.1); sp=sprite(txt,32,(255,255,255),b'Bold'); typed(x,sp,340,y,eo(pr),a)
    for i,(lab,col,desc,ic) in enumerate(LV):
        pi=5+i
        if t<ph(pi)['start']-0.2: continue
        y=-20+i*95; kk=back(seg(t,ph(pi)['start']-0.2,ph(pi)['start']+0.2))
        tag(x,lab,col,290,y,kk,a)
        pr=seg(t,P[pi]['words'][2]['start'] if len(P[pi]['words'])>2 else ph(pi)['start'],ph(pi)['end']-0.1)
        sp=sprite(desc,32,(255,255,255),b'Bold'); typed(x,sp,200,y,eo(pr),a)
    pulse=seg(t,s8-0.1,s8+0.3)
    circle(x,-330,330,34*(1+0.25*math.sin(pulse*math.pi)),GOLD,a); blit(x,mi(0xe163,38,(14,32,80)),-330,330,1,a)
    x.restore()
def scene_3(x,t):
    s8=ph(8)['start']
    k=seg(t,s8-0.2,s8+0.3)
    if k<=0: return
    a=eo(k)
    icon_disc(x,0xf06c,0,-300,70,back(k),a,gold=True,glow=1)
    for i,(lab,col,desc,ic) in enumerate(LV):
        kk=eio(seg(t,P[8]['words'][-2]['start']-0.3+i*0.2,P[8]['words'][-2]['start']+0.4+i*0.2)); px=(1-i)*240
        if kk<=0: continue
        cx=lerp(0,px,kk); cy=lerp(-300,40,kk)
        sheet(x,cx,cy,lerp(0.2,1.0,kk),col,lab,a,rot=(1-i)*0.06*kk,icon=ic)
        if kk>=1: blit(x,sprite(desc,26,(255,255,255),b'Bold'),px,230,1,a)
    kc=back(seg(t,ph(8)['end']-0.3,ph(8)['end']+0.2))
    if kc>0:
        for i in range(3): circle(x,(1-i)*240+80,-80,22*kc,GOLD,a); check(x,(1-i)*240+80,-81,0.6,1,NAVY1,a,lw=5)
def scene_4(x,t):
    s9,s10=ph(9)['start'],ph(10)['start']
    k=seg(t,s9-0.2,s9+0.3); ko=eio(seg(t,s10-0.15,s10+0.35))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        for i,(lab,col,_,ic) in enumerate(LV):
            kk=back(seg(t,P[9]['words'][2]['start']+i*0.35,P[9]['words'][2]['start']+0.4+i*0.35)); px=(1-i)*240
            if kk<=0: continue
            for q,dx in enumerate((-55,0,55)): person(x,px+dx,230-(14 if q==1 else 0),0.7,WHITE,a)
            fly=eio(seg(t,P[9]['words'][2]['start']+i*0.35,P[9]['words'][2]['start']+0.6+i*0.35))
            sheet(x,px,lerp(-260,30,fly),0.55*kk,col,None,a,lines=3)
        kh=back(seg(t,P[9]['words'][-3]['start']-0.1,P[9]['words'][-3]['start']+0.4))
        if kh>0:
            rrect(x,-300,-330,600,90,45,WHITE,a*kh)
            blit(x,sprite("لا أحد يشعر بالفرق",36,(14,32,80),b'Black'),40,-287,kh,a); blit(x,mi(0xe87d,44,(230,70,90)),-240,-285,kh,a)
    k=seg(t,s10-0.2,s10+0.3)
    if k>0:
        a=eo(k)
        for i,(lab,col,_,ic) in enumerate(LV):
            sheet(x,(1-i)*230,-40,0.85,col,lab,a)
            ck=seg(t,s10+0.4+i*0.4,s10+0.8+i*0.4)
            circle(x,(1-i)*230+75,-170,22*back(ck),GOLD,a); check(x,(1-i)*230+75,-171,0.6,eo(ck),NAVY1,a,lw=5)
        u=seg(t,s10,ph(10)['end']); magnifier(x,lerp(260,-260,eio(u)),40,0.65,a)
        ke=seg(t,P[10]['words'][1]['start'],P[10]['words'][1]['start']+0.7)
        if ke>0: blit(x,mi(0xe3c9,56,(246,196,72)),lerp(0,-40,eo(ke)),-60,1,a)
        kx=back(seg(t,P[10]['words'][-2]['start']-0.1,P[10]['words'][-2]['start']+0.4))
        if kx>0: icon_disc(x,0xe7fd,0,280,52,kx,a,gold=True,glow=0.8); blit(x,sprite("أنت الأعرف بطلابك",32,(246,196,72),b'Black'),0,370,kx,a)
def scene_5(x,t):
    s11,s12=ph(11)['start'],ph(12)['start']
    k=seg(t,s11-0.2,s11+0.3); ko=eio(seg(t,s12-0.15,s12+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        for j in range(6):
            g=back(seg(t,s11+j*0.1,s11+0.4+j*0.1),1.2); bx=330-j*120; h=(j+1)*80*g
            rrect(x,bx-56,330-h,112,h,10,GOLD if j==5 else CARD,a); rrect(x,bx-56,330-h,112,10,5,GOLD2,a)
        climb=eio(seg(t,P[11]['words'][-3]['start'],ph(11)['end']+0.3))
        for i,(lab,col,_,_) in enumerate(LV):
            st=i; en=st+3; pos=lerp(st,en,climb); jj=int(min(5,pos)); f=pos-jj
            x0=330-jj*120; x1=330-min(5,jj+1)*120; y0=330-(jj+1)*80; y1=330-(min(5,jj+1)+1)*80
            person(x,lerp(x0,x1,f),lerp(y0,y1,f)-10-math.sin(f*math.pi)*30,0.8,col,a)
    fin=eio(seg(t,VEND+0.1,VEND+0.7))
    k=seg(t,s12-0.15,s12+0.4)
    if k>0 and fin<1:
        a=eo(k)*(1-fin)
        icon_disc(x,0xe866,0,-80,120,back(k),a,gold=True,glow=1); blit(x,sprite("احفظ",44,(246,196,72),b'Black'),0,80,back(k),a)
        for j,(ic,px,py) in enumerate(((0xe0bf,-240,-200),(0xe87d,240,-190),(0xe0b7,-230,120),(0xe8dc,250,140))):
            kk=back(seg(t,s12+0.5+j*0.2,s12+0.9+j*0.2)); icon_disc(x,ic,px,py+8*math.sin(t*3+j),44,kk,a,gold=(j%2==1))
        kd=seg(t,P[12]['words'][-1]['start']-0.4,VEND+1)
        if kd>0:
            for j in range(3): blit(x,mi(0xe5db,64,(246,196,72)),0,190+j*50+8*math.sin(t*5),1,a*((math.sin(t*5-j)+1)/2))
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        for i,(lab,col,_,ic) in enumerate(LV):
            e=back(seg(t,VEND+0.2+i*0.2,VEND+0.6+i*0.2)); sheet(x,(1-i)*200,-200,0.7*e,col,lab,a,rot=(1-i)*0.08)
        e1=eo(seg(t,VEND+0.7,VEND+1.3)); e2=eo(seg(t,VEND+1.1,VEND+1.7))
        blit(x,sprite("درس واحد… ٣ مستويات",72,(255,255,255),b'Black'),0,70+(1-e1)*30,1,a*e1)
        blit(x,sprite("في دقيقة واحدة",92,(246,196,72),b'Black',glow=(245,180,40)),0,200+(1-e2)*30,1+0.03*math.sin(t*3),a*e2)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
def OVERLAY(x,t): pass
