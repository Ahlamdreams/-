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

def cwin(x,a,t,x0=-410,y0=-340,w=820,h=700):
    rrect(x,x0,y0+10,w,h,36,(0,0,0),0.25*a); rrect(x,x0,y0,w,h,36,(0.08,0.15,0.34),a)
    rrect(x,x0,y0,w,90,36,(0.12,0.24,0.52),a); rrect(x,x0,y0+50,w,40,0,(0.12,0.24,0.52),a)
    rrect(x,x0,y0,w,h,36,GOLD,a,fill=False,lw=3)
    icon_disc(x,0xe65f,x0+w-60,y0+45,30,1,a,gold=True)
    blit(x,sprite("Claude",36,(255,255,255),b'Black',rtl=False),x0+w-150,y0+43,1,a)
    for j in range(3): circle(x,x0+40+j*28,y0+45,8,(0.4,0.5,0.75),a)
def game(x,cx,cy,s,a,t,vivid=0,timer=0,ans=1,score=30):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    g=cairo.LinearGradient(0,-300,0,300)
    c1=mix((0.1,0.2,0.45),(0.95,0.35,0.55),vivid); c2=mix((0.06,0.12,0.3),(0.98,0.6,0.2),vivid)
    g.add_color_stop_rgba(0,c1[0],c1[1],c1[2],a); g.add_color_stop_rgba(1,c2[0],c2[1],c2[2],a)
    x.new_path(); x.rectangle(-230,-300,460,600); x.set_source(g)
    x.save(); rrect(x,-230,-300,460,600,34,WHITE,0); x.new_path()
    r=34; x.arc(230-r,-300+r,r,-math.pi/2,0); x.arc(230-r,300-r,r,0,math.pi/2); x.arc(-230+r,300-r,r,math.pi/2,math.pi); x.arc(-230+r,-300+r,r,math.pi,1.5*math.pi); x.close_path(); x.clip()
    x.set_source(g); x.paint(); x.restore()
    rrect(x,-230,-300,460,600,34,WHITE,a*0.8,fill=False,lw=5)
    blit(x,mi(0xe838,40,(246,196,72)),150,-250,1,a)
    blit(x,sprite(''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(score)),36,(255,255,255),b'Black'),95,-252,1,a)
    blit(x,sprite("سؤال ٣ من ١٠",26,(220,230,255),b'Bold'),-100,-252,1,a)
    if timer>0:
        circle(x,-170,-180,34*timer,WHITE,a*0.25,fill=False,lw=6)
        x.new_path(); x.arc(-170,-180,34*timer,-math.pi/2,-math.pi/2+2*math.pi*(1-((t*0.2)%1))); x.set_line_width(6); rgba(GOLD2,a,x); x.stroke()
    rrect(x,-190,-150,380,120,22,WHITE,a)
    blit(x,sprite("١/٤ + ١/٤ = ؟",50,(14,32,80),b'Black'),0,-92,1,a)
    opts=["١/٢","٢/٨","١/٨","٢/٤"]
    for j,o in enumerate(opts):
        r_,q=j//2,j%2; px=100-q*200; py=20+r_*110
        ok=(j==0); sel=ans>0.5 and ok
        rrect(x,px-85,py-44,170,88,20,GOLD if sel else (1,1,1),a*(1 if sel else 0.22)); rrect(x,px-85,py-44,170,88,20,WHITE,a*0.7,fill=False,lw=3)
        blit(x,sprite(o,40,(14,32,80) if sel else (255,255,255),b'Black'),px,py-3,1,a)
    if ans>0.5:
        e=back(cl((ans-0.5)*2)); x.save(); x.translate(0,250); x.scale(S_(e),S_(e))
        rrect(x,-140,-36,280,72,36,WHITE,a); blit(x,sprite("أحسنت!",36,(14,32,80),b'Black'),20,-3,1,a); blit(x,mi(0xea65,40,(246,196,72)),-90,0,1,a); x.restore()
    x.restore()
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    k=back(seg(t,s0-0.4,s0+0.2)); ko=eio(seg(t,s1-0.15,s1+0.35))
    if k>0 and ko<1:
        a=1-ko
        icon_disc(x,0xea28,0,-120,150,k,a,gold=True,glow=1)
        kc=back(seg(t,P[0]['words'][-1]['start']-0.3,P[0]['words'][-1]['start']+0.2))
        if kc>0:
            icon_disc(x,0xe86f,0,200,80,kc,a); ban(x,0,200,110,seg(t,P[0]['words'][-1]['start'],P[0]['words'][-1]['start']+0.5),a,lw=10)
            for j in range(5):
                y=150+j*22; rrect(x,-330,y,lerp(60,140,(j%3)/2),10,5,SOFT,a*0.5*kc); rrect(x,200,y,lerp(70,130,(j%2)),10,5,SOFT,a*0.5*kc)
    k=seg(t,s1-0.2,s1+0.4)
    if k>0:
        a=eo(k); e=back(k)
        x.save(); x.scale(S_(e),S_(e)); radial(x,0,-60,320,GOLD,0.4*a)
        rrect(x,-230,-150,460,180,90,WHITE,a); blit(x,sprite("Claude",96,(14,32,80),b'Black',rtl=False),30,-62,1,a); blit(x,mi(0xe65f,80,(246,196,72)),-150,-60,1,a); x.restore()
        p=seg(t,P[1]['words'][-2]['start'],ph(1)['end']+0.3); cx,cy=0,220
        circle(x,cx,cy,80,CARD,a); x.new_path(); x.arc(cx,cy,80,-math.pi/2,-math.pi/2+2*math.pi*max(0.001,1-p)); x.set_line_width(12); rgba(GOLD,a,x); x.stroke()
        blit(x,sprite(str(int(60*(1-p))),54,(255,255,255),b'Black',rtl=False),cx,cy-4,1,a)
def scene_2(x,t):
    t_now[0]=t
    s2=ph(2)['start']; s5=ph(5)['start']
    k=back(seg(t,s2-0.4,s2+0.1))
    if k<=0: return
    x.save(); x.scale(S_(k),S_(k)); a=1
    cwin(x,a,t)
    rrect(x,-380,-220,760,500,26,(0.05,0.1,0.25),a)
    lines=[(3,"اصنع لي لعبة تفاعلية",GOLD,"لعبة"),(3,"للصف الخامس عن الكسور",(0.6,0.75,1.0),"الدرس"),(4,"١٠ أسئلة + عدّاد للنقاط",GOLD2,"المحتوى"),(4,"رسالة تشجيع لكل إجابة صحيحة",(0.85,0.9,1.0),"التحفيز")]
    for i,(pi,txt,col,lab) in enumerate(lines):
        st=ph(pi)['start']+(0 if i%2==0 else (ph(pi)['end']-ph(pi)['start'])*0.5); en=st+(ph(pi)['end']-ph(pi)['start'])*0.5
        if t<st-0.2: continue
        y=-150+i*100; kk=back(seg(t,st-0.2,st+0.2))
        tag(x,lab,col,300,y,kk,a)
        typed(x,sprite(txt,32,(255,255,255),b'Bold'),210,y,eo(seg(t,st,en)),a)
    kw=back(seg(t,P[2]['words'][-1]['start']-0.1,P[2]['words'][-1]['start']+0.3))
    if kw>0 and t<ph(3)['start']: blit(x,mi(0xe3c9,70,(246,196,72)),0,40,kw,a)
    pulse=seg(t,s5-0.1,s5+0.3)
    circle(x,-330,330,34*(1+0.25*math.sin(pulse*math.pi)),GOLD,a); blit(x,mi(0xe163,38,(14,32,80)),-330,330,1,a)
    x.restore()
def scene_3(x,t):
    s5,s6=ph(5)['start'],ph(6)['start']
    k=seg(t,s5-0.2,s5+0.3)
    if k<=0: return
    a=eo(k)
    build=seg(t,P[5]['words'][2]['start']-0.4,P[5]['words'][2]['start']+1.2)
    if build<1:
        rrect(x,-230,-300,460,600,34,(0.05,0.1,0.25),a*(1-build))
        for j in range(14):
            y=-270+((j*44+t*160)%560); w=[260,180,320,140,240][j%5]
            rrect(x,-200,y,w,14,7,[GOLD,(0.6,0.75,1.0),WHITE,GOLD2][j%4],a*0.7*(1-build))
        blit(x,sprite("Claude يبني اللعبة…",36,(255,255,255),b'Bold'),0,340,1,a*(1-build))
    if build>0:
        vivid=eio(seg(t,P[6]['words'][2]['start'],P[6]['words'][2]['start']+1.0))
        tim=back(seg(t,P[6]['words'][-3]['start']-0.1,P[6]['words'][-3]['start']+0.4))
        ans=seg(t,ph(5)['end']-0.8,ph(5)['end']+0.4)
        game(x,0,0,lerp(0.7,1.0,eo(build)),a*eo(build),t,vivid=vivid,timer=tim,ans=ans,score=int(30+10*(ans>0.5)))
        if build>=1 and t<s6: 
            for j in range(8):
                ang=t+j*0.785; poly(x,star_pts(300*math.cos(ang),340*math.sin(ang),10,rot=t),GOLD2,a)
    km=back(seg(t,s6-0.2,s6+0.3)); kmo=eio(seg(t,ph(6)['end']-0.3,ph(6)['end']+0.2))
    if km>0 and kmo<1:
        rrect(x,-330,-400,660,90,45,GOLD,a*(1-kmo)); blit(x,sprite("ألوان أكثر حيوية + مؤقّت لكل سؤال",30,(14,32,80),b'Black'),20,-357,km,a*(1-kmo))
        blit(x,mi(0xe40a,44,(14,32,80)),-280,-355,km,a*(1-kmo))
def scene_4(x,t):
    s7,s8=ph(7)['start'],ph(8)['start']
    k=seg(t,s7-0.2,s7+0.3); ko=eio(seg(t,s8-0.15,s8+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        game(x,0,-120,0.55,a,t,vivid=1,timer=1,ans=0)
        kl=back(seg(t,P[7]['words'][1]['start']-0.1,P[7]['words'][1]['start']+0.4))
        if kl>0:
            rrect(x,-250,110,500,80,40,WHITE,a*kl); blit(x,mi(0xe157,44,(246,196,72)),-190,150,kl,a); blit(x,sprite("claude.ai/…",32,(14,32,80),b'Bold',rtl=False),30,148,kl,a)
        for j in range(3):
            kk=back(seg(t,P[7]['words'][3]['start']+j*0.2,P[7]['words'][3]['start']+0.4+j*0.2)); px=220-j*140
            if kk>0:
                glowline(x,partial([(0,190),(px,250)],eo(kk)),GOLD,a*0.6,3)
                rrect(x,px-38,250,76,120,14,WHITE,a*kk); rrect(x,px-30,262,60,90,8,mix((0.95,0.35,0.55),(0.98,0.6,0.2),j/2),a*kk)
        kt=back(seg(t,P[7]['words'][-2]['start']-0.1,P[7]['words'][-2]['start']+0.4))
        if kt>0: icon_disc(x,0xe333,-250,320,62,kt,a,gold=True,glow=0.5)
    k=seg(t,s8-0.2,s8+0.3)
    if k>0:
        a=eo(k); ans=seg(t,P[8]['words'][1]['start'],P[8]['words'][1]['start']+1.0)
        game(x,0,-40,0.85,a,t,vivid=1,timer=1,ans=ans,score=40)
        magnifier(x,lerp(200,-120,eio(seg(t,P[8]['words'][-3]['start'],ph(8)['end']))),-130,0.6,a)
        kc=back(seg(t,ph(8)['end']-0.4,ph(8)['end']+0.1))
        if kc>0: circle(x,200,-330,40*kc,GOLD,a); check(x,200,-332,1.1,1,NAVY1,a,lw=7)
def scene_5(x,t):
    s9,s10=ph(9)['start'],ph(10)['start']
    k=seg(t,s9-0.2,s9+0.3); ko=eio(seg(t,s10-0.15,s10+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        person(x,250,-100,1.4,GOLD,a); icon_disc(x,0xe0f0,250,-290,50,back(k),a,gold=True,glow=0.6)
        kk=eo(seg(t,P[9]['words'][2]['start'],P[9]['words'][2]['start']+0.8))
        glowline(x,partial([(170,-200),(0,-230),(-150,-120)],kk),GOLD2,a,6)
        game(x,-170,40,0.45*back(kk),a,t,vivid=1,timer=0,ans=1,score=50)
        kp=seg(t,P[9]['words'][-2]['start'],ph(9)['end']+0.3)
        for j in range(4):
            e=back(seg(kp,j*0.15,j*0.15+0.4)); px=-300+j*170
            if e>0: person(x,px,330,0.8*e,WHITE,a); poly(x,star_pts(px,210,12*e,rot=t+j),GOLD2,a)
    fin=eio(seg(t,VEND+0.1,VEND+0.7))
    k=seg(t,s10-0.15,s10+0.4)
    if k>0 and fin<1:
        a=eo(k)*(1-fin)
        icon_disc(x,0xe866,0,-80,120,back(k),a,gold=True,glow=1); blit(x,sprite("احفظ",44,(246,196,72),b'Black'),0,80,back(k),a)
        for j,(ic,px,py) in enumerate(((0xea28,-240,-200),(0xe87d,240,-190),(0xe0b7,-230,120),(0xe8dc,250,140))):
            kk=back(seg(t,s10+0.5+j*0.2,s10+0.9+j*0.2)); icon_disc(x,ic,px,py+8*math.sin(t*3+j),44,kk,a,gold=(j%2==1))
        kd=seg(t,P[10]['words'][-1]['start']-0.4,VEND+1)
        if kd>0:
            for j in range(3): blit(x,mi(0xe5db,64,(246,196,72)),0,190+j*50+8*math.sin(t*5),1,a*((math.sin(t*5-j)+1)/2))
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        e0=back(seg(t,VEND+0.1,VEND+0.6))
        icon_disc(x,0xe0f0,200,-230,70,e0,a,gold=True,glow=0.6); blit(x,sprite("+",60,(246,196,72),b'Black',rtl=False),80,-230,e0,a)
        rrect(x,-130,-275,200,90,45,WHITE,a*e0); blit(x,sprite("Claude",40,(14,32,80),b'Black',rtl=False),-30,-232,e0,a)
        e1=eo(seg(t,VEND+0.6,VEND+1.2)); e2=eo(seg(t,VEND+1.0,VEND+1.6))
        blit(x,sprite("فكرتك + Claude",74,(255,255,255),b'Black'),0,-40+(1-e1)*30,1,a*e1)
        blit(x,sprite("= لعبة يعيشها طلابك",80,(246,196,72),b'Black',glow=(245,180,40)),0,100+(1-e2)*30,1+0.03*math.sin(t*3),a*e2)
        blit(x,sprite("بدون برمجة",40,(200,210,235),b'Bold'),0,210,1,a*e2)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
def OVERLAY(x,t): pass
