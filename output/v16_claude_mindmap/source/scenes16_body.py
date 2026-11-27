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

BR=[("التبخّر",0xe430,(1.0,0.62,0.25),"الشمس تسخّن الماء",(170,-190)),("التكاثف",0xe2bd,(0.62,0.78,1.0),"تتكوّن الغيوم",(-170,-190)),("الهطول",0xe798,(0.35,0.6,1.0),"مطر وثلج",(-170,190)),("التجمّع",0xe176,(0.3,0.8,0.75),"أنهار وبحار",(170,190))]
def bookpage(x,cx,cy,s,a,rot=0,dense=1):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.rotate(rot); x.scale(S_(s),S_(s))
    rrect(x,-170,-230,340,460,14,(0,0,0),0.25*a); rrect(x,-170,-240,340,460,14,WHITE,a)
    blit(x,sprite("دورة الماء",34,(14,32,80),b'Black'),0,-200,1,a)
    for j in range(int(14*dense)):
        y=-160+j*26; w=[280,250,270,230,260][j%5]
        rrect(x,140-w,y,w,10,5,(0.72,0.76,0.86),a)
    x.restore()
def mindmap(x,cx,cy,t,a,kc,kb,kl,sc=1):
    x.save(); x.translate(cx,cy); x.scale(S_(sc),S_(sc))
    for i,(lab,ic,col,leaf,(bx,by)) in enumerate(BR):
        k=kb[i]
        if k<=0: continue
        ex,ey=bx*eo(k),by*eo(k)
        x.new_path(); x.move_to(0,0); x.curve_to(bx*0.2,by*0.7,bx*0.6,by*0.3,ex,ey); x.set_line_width(10); rgba(col,a,x); x.set_line_cap(cairo.LINE_CAP_ROUND); x.stroke()
        if k>=0.95:
            e=back(cl((k-0.95)*20))
            x.save(); x.translate(bx,by); x.scale(S_(e),S_(e))
            rrect(x,-95,-42,190,84,42,col,a); blit(x,sprite(lab,32,(14,32,80),b'Black'),18,-3,1,a); blit(x,mi(ic,36,(14,32,80)),-58,0,1,a); x.restore()
            kk=kl[i]
            if kk>0:
                lx,ly=bx*1.0,by+(90 if by>0 else -90)
                poly(x,[(bx,by+(42 if by>0 else -42)),(lx,ly+(-26 if by>0 else 26))],col,a*0.7,fill=False,lw=4,close=False)
                rrect(x,lx-110*kk,ly-26,220*kk,52,26,WHITE,a); blit(x,sprite(leaf,24,(14,32,80),b'Bold'),lx,ly-2,kk,a)
    if kc>0:
        e=back(kc); x.save(); x.scale(S_(e),S_(e)); radial(x,0,0,200,GOLD,0.5*a)
        circle(x,0,0,92,GOLD,a); circle(x,0,0,92,WHITE,a,fill=False,lw=5); blit(x,sprite("دورة الماء",32,(14,32,80),b'Black'),0,-3,1,a); x.restore()
    x.restore()
def scene_1(x,t):
    s0,s1=ph(0)['start'],ph(1)['start']
    k=back(seg(t,s0-0.4,s0+0.2)); m=eio(seg(t,P[1]['words'][-3]['start']-0.3,P[1]['words'][-3]['start']+0.7))
    if k>0 and m<1:
        bookpage(x,0,-60,1.25*k*(1-m*0.5),1-m,rot=-0.03)
        kq=back(seg(t,P[0]['words'][-2]['start']-0.1,P[0]['words'][-2]['start']+0.4))
        if kq>0 and m<1:
            person(x,250,330,1.2*kq,SOFT,1-m)
            for j in range(3): blit(x,sprite("؟",60+j*12,(246,196,72),b'Black'),200+j*45,150-j*35,kq,(1-m))
    k=seg(t,s1-0.2,s1+0.3)
    if k>0:
        a=eo(k)
        x.save(); x.translate(0,-330); rrect(x,-140,-45,280,90,45,WHITE,a); blit(x,sprite("Claude",44,(14,32,80),b'Black',rtl=False),25,-3,1,a); blit(x,mi(0xe65f,44,(246,196,72)),-90,0,1,a); x.restore()
        if m>0:
            kb=[eo(seg(m,0.2+i*0.15,0.5+i*0.15)) for i in range(4)]
            mindmap(x,0,20,t,a,m,kb,[0,0,0,0],sc=0.9)
        kt=back(seg(t,P[1]['words'][-1]['start']-0.1,P[1]['words'][-1]['start']+0.4))
        if kt>0: clock(x,280,-330,44*kt,t*5,a)
def scene_2(x,t):
    t_now[0]=t
    s2,s3=ph(2)['start'],ph(3)['start']
    k=seg(t,s2-0.2,s2+0.3); ko=eio(seg(t,s3-0.2,s3+0.2))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        bookpage(x,0,40,1.0,a)
        snap=seg(t,P[2]['words'][1]['start']-0.1,P[2]['words'][1]['start']+0.4)
        x.save(); x.translate(lerp(380,0,eio(seg(t,s2,s2+0.6))),40)
        rrect(x,-200,-300,400,600,40,(0.05,0.08,0.15),a*0.35); rrect(x,-200,-300,400,600,40,WHITE,a,fill=False,lw=8)
        for (px,py,sx,sy) in ((-150,-250,1,1),(150,-250,-1,1),(-150,250,1,-1),(150,250,-1,-1)):
            poly(x,[(px,py+sy*40),(px,py),(px+sx*40,py)],GOLD,a,fill=False,lw=6,close=False)
        circle(x,0,340,36,WHITE,a); circle(x,0,340,28,GOLD,a)
        x.restore()
        if 0<snap<1: rgba(WHITE,0.8*(1-snap)*a,x); x.paint() if False else rrect(x,-540,-700,1080,1400,0,WHITE,0.7*(1-snap)*a)
        up=eio(seg(t,P[2]['words'][-2]['start']-0.2,ph(2)['end']+0.2))
        if up>0:
            icon_disc(x,0xe2c6,0,-330,60,back(up),a,gold=True,glow=0.6)
            rrect(x,-140,-430,280,70,35,WHITE,a*up); blit(x,sprite("Claude",34,(14,32,80),b'Black',rtl=False),0,-397,up,a)
    k=back(seg(t,s3-0.3,s3+0.1))
    if k>0:
        x.save(); x.scale(S_(k),S_(k)); a=1
        cwin(x,a,t)
        rrect(x,-380,-220,760,500,26,(0.05,0.1,0.25),a)
        x.save(); x.translate(240,-130); x.rotate(0.05); x.scale(0.35,0.35); bookpage(x,0,0,1,a); x.restore()
        blit(x,mi(0xe3f4,40,(246,196,72)),110,-200,1,a)
        lines=[("حوّل هذه الصفحة إلى",0.0),("خريطة ذهنية ملوّنة",0.33),("لطلاب الصف الخامس",0.66)]
        d=ph(3)['end']-P[3]['words'][2]['start'] if len(P[3]['words'])>2 else 2
        for i,(txt,f) in enumerate(lines):
            st=P[3]['words'][2]['start']+d*f; en=st+d*0.33
            if t<st-0.1: continue
            typed(x,sprite(txt,34,(255,255,255) if i!=1 else (246,196,72),b'Bold'),330,20+i*80,eo(seg(t,st,en)),a)
        pulse=seg(t,ph(4)['start']-0.1,ph(4)['start']+0.3)
        circle(x,-330,330,34*(1+0.25*math.sin(pulse*math.pi)),GOLD,a); blit(x,mi(0xe163,38,(14,32,80)),-330,330,1,a)
        x.restore()
def scene_3(x,t):
    s4,s5=ph(4)['start'],ph(5)['start']
    k=seg(t,s4-0.2,s4+0.3)
    if k<=0: return
    a=eo(k)
    kc=seg(t,P[4]['words'][4]['start']-0.3,P[4]['words'][4]['start']+0.3) if len(P[4]['words'])>4 else seg(t,s4+0.5,s4+1)
    br0=P[4]['words'][-2]['start']-0.6
    kb=[seg(t,br0+i*0.3,br0+0.6+i*0.3) for i in range(4)]
    kl=[eo(seg(t,ph(4)['end']-0.2+i*0.15,ph(4)['end']+0.2+i*0.15)) for i in range(4)]
    sh=eio(seg(t,s5,s5+0.6))
    mindmap(x,0,lerp(0,-120,sh),t,a,kc,kb,kl,sc=lerp(1.0,0.75,sh))
    if sh>0:
        qs=["ما سبب تبخّر الماء؟","كيف تتكوّن الغيوم؟","أين يتجمّع الماء؟"]
        for j,q in enumerate(qs):
            kk=back(seg(t,P[5]['words'][-3]['start']+j*0.2,P[5]['words'][-3]['start']+0.4+j*0.2)); y=230+j*80
            if kk<=0: continue
            rrect(x,-320,y-32,640*kk,64,18,WHITE,a); circle(x,280,y,22,GOLD,a)
            blit(x,sprite('٠١٢٣'[j+1],26,(14,32,80),b'Black'),280,y-2,kk,a); blit(x,sprite(q,28,(14,32,80),b'Bold'),20,y-2,kk,a)
def scene_4(x,t):
    s6=ph(6)['start']
    k=seg(t,s6-0.2,s6+0.3)
    if k<=0: return
    a=eo(k)
    bookpage(x,190,-60,0.8,a,rot=0.03)
    mindmap(x,-190,-60,t,a,1,[1,1,1,1],[0,0,0,0],sc=0.5)
    u=seg(t,s6+0.2,ph(6)['end']); magnifier(x,lerp(190,-190,eio(u)),-60+60*math.sin(u*math.pi*2),0.6,a)
    for i in range(4):
        ck=seg(t,s6+0.6+i*0.3,s6+0.9+i*0.3)
        if ck>0:
            px,py=-190+BR[i][4][0]*0.5+40,-60+BR[i][4][1]*0.5-20
            circle(x,px,py,16*back(ck),GOLD,a); check(x,px,py-1,0.45,1,NAVY1,a,lw=4)
    kk=back(seg(t,ph(6)['end']-0.4,ph(6)['end']+0.1))
    if kk>0: blit(x,sprite("مطابقة ✓",40,(246,196,72),b'Black'),0,290,kk,a)
def scene_5(x,t):
    s7,s8=ph(7)['start'],ph(8)['start']
    k=seg(t,s7-0.2,s7+0.3); ko=eio(seg(t,s8-0.15,s8+0.3))
    if k>0 and ko<1:
        a=eo(k)*(1-ko)
        bookpage(x,200,-40,0.75,a*0.85,rot=0.05)
        ka=eo(seg(t,s7+0.6,s7+1.2)); glowline(x,partial([(80,-40),(-60,-40)],ka),GOLD,a,7)
        if ka>0.9: arrow_shape(x,-40,-40,-70,-40,GOLD,a,lw=1,head=26)
        km=seg(t,P[7]['words'][-3]['start']-0.3,P[7]['words'][-3]['start']+0.6)
        mindmap(x,-230,-40,t,a,km,[km]*4,[0,0,0,0],sc=0.42)
        kh=back(seg(t,P[7]['words'][-1]['start']-0.2,P[7]['words'][-1]['start']+0.3))
        if kh>0: person(x,-230,330,1.1*kh,GOLD,a); icon_disc(x,0xe0f0,-230,170,40,kh,a,gold=True,glow=0.6)
    fin=eio(seg(t,VEND+0.1,VEND+0.7))
    k=seg(t,s8-0.15,s8+0.4)
    if k>0 and fin<1:
        a=eo(k)*(1-fin)
        icon_disc(x,0xe866,0,-80,120,back(k),a,gold=True,glow=1); blit(x,sprite("احفظ",44,(246,196,72),b'Black'),0,80,back(k),a)
        for j,(ic,px,py) in enumerate(((0xe97a,-240,-200),(0xe87d,240,-190),(0xe0b7,-230,120),(0xe8dc,250,140))):
            kk=back(seg(t,s8+0.5+j*0.2,s8+0.9+j*0.2)); icon_disc(x,ic,px,py+8*math.sin(t*3+j),44,kk,a,gold=(j%2==1))
        kd=seg(t,P[8]['words'][-1]['start']-0.4,VEND+1)
        if kd>0:
            for j in range(3): blit(x,mi(0xe5db,64,(246,196,72)),0,190+j*50+8*math.sin(t*5),1,a*((math.sin(t*5-j)+1)/2))
    if fin>0:
        a=fin*(1-eio(seg(t,OUTRO-0.3,OUTRO+0.2)))
        e0=back(seg(t,VEND+0.1,VEND+0.6))
        bookpage(x,190,-220,0.45*e0,a,rot=0.05); mindmap(x,-190,-220,t,a,e0,[e0]*4,[0,0,0,0],sc=0.38)
        glowline(x,[(100,-220),(-80,-220)],GOLD,a*e0,6)
        e1=eo(seg(t,VEND+0.6,VEND+1.2)); e2=eo(seg(t,VEND+1.0,VEND+1.6))
        blit(x,sprite("صوّر الصفحة…",74,(255,255,255),b'Black'),0,60+(1-e1)*30,1,a*e1)
        blit(x,sprite("تحصل على خريطة ذهنية",76,(246,196,72),b'Black',glow=(245,180,40)),0,190+(1-e2)*30,1+0.03*math.sin(t*3),a*e2)
        blit(x,sprite("مع Claude",44,(200,210,235),b'Bold'),0,300,1,a*e2)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
def OVERLAY(x,t): pass
