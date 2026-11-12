def cloud(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s)); x.new_path()
    for (px,py,r) in ((-110,10,70),(-40,-45,88),(55,-38,80),(125,12,62),(0,20,80)): x.arc(px,py,r,0,2*math.pi); x.new_sub_path()
    rgba(WHITE,a,x); x.fill(); rrect(x,-175,10,350,72,36,WHITE,a)
    circle(x,-150,130,18,WHITE,a); circle(x,-185,175,10,WHITE,a); x.restore()
def nums(n): return ''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(n))
def card_row(x,x0,y,w,a,k,gold=False):
    rrect(x,x0,y-26,w,52,14,GOLD if gold else WHITE,a*(0.95 if gold else 0.9))
    rrect(x,x0+w-46,y-14,28,28,7,NAVY2,a*0.8,fill=False,lw=4)
    rrect(x,x0+20,y-7,(w-90)*k,14,7,NAVY2 if gold else SOFT,a)
def diamond(x,cx,cy,r,a,t):
    radial(x,cx,cy,r*2.3,GOLD,0.45*a)
    top=[(cx-r,cy-r*0.35),(cx-r*0.55,cy-r*0.85),(cx+r*0.55,cy-r*0.85),(cx+r,cy-r*0.35),(cx,cy+r)]
    poly(x,top,GOLD,a); poly(x,[(cx-r,cy-r*0.35),(cx+r,cy-r*0.35),(cx,cy+r)],GOLD2,a*0.9)
    poly(x,[(cx-r*0.55,cy-r*0.85),(cx-r*0.2,cy-r*0.35),(cx,cy-r*0.85),(cx+r*0.2,cy-r*0.35),(cx+r*0.55,cy-r*0.85)],WHITE,a*0.35)
    poly(x,top,WHITE,a,fill=False,lw=4)
    for i in range(6):
        ang=t*0.9+i*1.05; rr=r*1.6+12*math.sin(t*3+i); poly(x,star_pts(cx+rr*math.cos(ang),cy+rr*math.sin(ang),9,rot=t),GOLD2,a*0.85)
def scene_1(x,t):
    s0,s1,s2,s3=[ph(i)['start'] for i in (0,1,2,3)]
    # thought cloud with ?
    kc=back(seg(t,s0-0.45,s0+0.15)); kco=eio(seg(t,s1-0.15,s1+0.35))
    if kc>0 and kco<1:
        s=kc*lerp(1.2,0.5,kco); cy=lerp(-20,-300,kco)
        radial(x,0,cy,300*s,GOLD,0.18); cloud(x,0,cy,s,1-kco)
        q=sprite("؟",150,(246,196,72),b'Black'); blit(x,q,0,cy-10*s,s*(1+0.06*math.sin(t*5)),1-kco)
    # busy gears + clock
    kg=back(seg(t,s1-0.25,s1+0.3)); kgo=eio(seg(t,s3-0.15,s3+0.35))
    up=eio(seg(t,s2-0.1,s2+0.4))
    if kg>0 and kgo<1:
        a=1-kgo; x.save(); x.translate(0,lerp(0,-170,up)); s=lerp(1,0.62,up); x.scale(s,s)
        sp=1+3*seg(t,wt(1,1),wt(1,1)+0.6)
        gear(x,-150,-20,95*kg,t*1.2*sp,GOLD,a); gear(x,10,-110,62*kg,-t*1.9*sp+0.2,WHITE,a,teeth=8); gear(x,40,60,48*kg,-t*2.4*sp,SOFT,a,teeth=8)
        clock(x,190,-10,92*kg,t*2.5*sp,a)
        x.restore()
    # aspiration star not reached
    km=seg(t,s2-0.1,s2+0.3); kmo=eio(seg(t,s3-0.15,s3+0.3))
    if km>0 and kmo<1:
        a=eo(km)*(1-kmo); sx,sy=230,170
        radial(x,sx,sy,170,GOLD,0.35*a); poly(x,star_pts(sx,sy,70*back(km),rot=-math.pi/2+t*0.3),GOLD2,a)
        kp=seg(t,wt(2,4)-0.3,ph(2)['end'])
        pts=[(lerp(-300,150,u),lerp(330,190,u)-120*math.sin(u*math.pi)) for u in np.linspace(0,eo(kp),30)]
        for j in range(0,len(pts)-1,2): poly(x,[pts[j],pts[j+1]],WHITE,a*0.8,fill=False,lw=5,close=False)
        if kp>=1:
            kx=seg(t,ph(2)['end'],ph(2)['end']+0.3); d=26
            glowline(x,partial([(150-d,190-d),(150+d,190+d)],kx),GOLD,a,9); glowline(x,partial([(150+d,190-d),(150-d,190+d)],kx),GOLD,a,9)
    # busy != achievement
    kn=seg(t,s3-0.2,s3+0.3)
    if kn>0:
        a=eo(kn)
        k1=back(seg(t,wt(3,1)-0.1,wt(3,1)+0.35)); k2=back(seg(t,wt(3,5)-0.15,wt(3,5)+0.35)); ke=back(seg(t,wt(3,3)-0.1,wt(3,3)+0.3))
        if k1>0:
            circle(x,220,0,140*k1,CARD,a); circle(x,220,0,140*k1,WHITE,a*0.7,fill=False,lw=4)
            gear(x,220,-20,60*k1,t*3,SOFT,a); sp=sprite("انشغال",44,(255,255,255),b'Bold'); blit(x,sp,220,80,k1,a)
        if k2>0:
            radial(x,-220,0,300*k2,GOLD,0.35*a); circle(x,-220,0,140*k2,GOLD,a)
            check(x,-220,-25,2.2*k2,eo(seg(t,wt(3,5),wt(3,5)+0.5)),NAVY1,a,lw=12)
            sp=sprite("إنجاز",46,(14,32,80),b'Black'); blit(x,sp,-220,80,k2,a)
        if ke>0:
            poly(x,[(-30,-25),(30,-25)],GOLD2,a,fill=False,lw=12,close=False); poly(x,[(-30,25),(30,25)],GOLD2,a,fill=False,lw=12,close=False)
            glowline(x,[(26,-60*ke),(-26,60*ke)],GOLD2,a,10)
def scene_2(x,t):
    s4,s5,s6,s7,s8=[ph(i)['start'] for i in (4,5,6,7,8)]
    out=eio(seg(t,s7-0.15,s7+0.35))
    # schedule board
    kb=back(seg(t,s4-0.35,s4+0.2))
    if kb>0 and out<1:
        a=1-out; x.save(); x.translate(0,lerp(0,-80,out)); x.scale(S_(kb),S_(kb))
        rrect(x,-330,-330,660,620,30,(0,0,0),0.25*a); rrect(x,-330,-340,660,620,30,(0.09,0.17,0.38),a); rrect(x,-330,-340,660,80,30,GOLD,a)
        days=["أحد","اثنين","ثلاثاء","أربعاء","خميس"]
        for i,d in enumerate(days):
            sp=sprite(d,30,(14,32,80),b'Bold'); blit(x,sp,260-i*130,-302,1,a)
        # tasks (s4)
        for j in range(15):
            c,r=j%5,j//5; k=back(seg(t,s4+0.15+j*0.07,s4+0.4+j*0.07))
            if k>0: rrect(x,205-c*130,-240+r*58,110*k,46,10,WHITE,a*0.9); rrect(x,215-c*130,-224+r*58,70*k,12,6,SOFT,a)
        # meetings (s5)
        for j in range(10):
            c,r=j%5,j//5; k=back(seg(t,s5+0.1+j*0.08,s5+0.35+j*0.08))
            if k>0:
                bx,by=205-c*130,-66+r*80
                rrect(x,bx,by,110*k,68,10,GOLD,a*0.9)
                for q in range(3): circle(x,bx+25+q*30,by+26,9*k,NAVY1,a); 
        # flowing work (s6)
        kf=seg(t,s6-0.1,s6+0.3)
        if kf>0:
            for j in range(8):
                u=((t-s6)*0.6+j/8)%1; px=lerp(300,-300,u); al=a*eo(kf)*math.sin(u*math.pi)
                rrect(x,px-50,120,100,110,14,WHITE,al*0.9); rrect(x,px-35,140,70,10,5,SOFT,al); rrect(x,px-35,160,50,10,5,SOFT,al)
        x.restore()
    # magnifier over dots -> one makes the difference
    kq=seg(t,s7-0.2,s7+0.3)
    if kq>0:
        a=eo(kq)
        for r in range(5):
            for c in range(7):
                j=r*7+c; bx,by=-300+c*100,-200+r*100
                special=(r==2 and c==4)
                k=seg(t,wt(8,2)-0.1,wt(8,2)+0.5) if special else 0
                if special and k>0:
                    rr=lerp(18,70,back(k)); radial(x,bx,by,rr*3,GOLD,0.5*a); circle(x,bx,by,rr,GOLD,a)
                    check(x,bx,by-4,1.6*back(k),eo(seg(t,wt(8,3)-0.1,wt(8,3)+0.4)),NAVY1,a,lw=9)
                else: circle(x,bx,by,18,SOFT,a*(0.75-0.45*seg(t,wt(8,2),wt(8,2)+0.5)))
        km=back(seg(t,wt(7,2)-0.15,wt(7,2)+0.3)); kmo=eio(seg(t,wt(8,2)-0.2,wt(8,2)+0.2))
        if km>0 and kmo<1:
            u=seg(t,wt(7,2),wt(8,2)); mx=lerp(-250,100,eio(u))+40*math.sin(t*3); my=lerp(150,0,eio(u))
            magnifier(x,mx,my,km*(1-kmo)*0.9,1)
def scene_3(x,t):
    s9,s10,s11=[ph(i)['start'] for i in (9,10,11)]
    # clock + priority bars
    kc=back(seg(t,s9-0.35,s9+0.2)); kco=eio(seg(t,s10-0.15,s10+0.35))
    if kc>0 and kco<1:
        a=1-kco; cx=170
        radial(x,cx,-30,280,GOLD,0.25*a); circle(x,cx,-30,150*kc,CARD,a)
        calm=eio(seg(t,wt(9,4)-0.2,wt(9,4)+0.6)); clock(x,cx,-30,150*kc,lerp(t*3,0.0,calm)+math.pi/2*0,a)
        circle(x,cx,-30,150*kc,GOLD,a,fill=False,lw=6)
        for i in range(3):
            k=back(seg(t,wt(9,5)-0.1+i*0.2,wt(9,5)+0.3+i*0.2))
            if k>0:
                y=-150+i*120; w=(260-i*60)*k
                rrect(x,-160-w,y-38,w,76,20,GOLD if i==0 else CARD,a); rrect(x,-160-w,y-38,w,76,20,WHITE,a*0.5,fill=False,lw=3)
                num_node(x,-160-w-10,y,34,i+1,a,i==0,k)
    # person asking 3 questions
    kp=back(seg(t,s10-0.25,s10+0.25)); kpo=eio(seg(t,s11-0.15,s11+0.35))
    if kp>0 and kpo<1:
        a=1-kpo; person(x,0,150,2.4*kp,WHITE,a)
        for i in range(3):
            k=back(seg(t,wt(10,4)-0.1+i*0.18,wt(10,4)+0.3+i*0.18))
            if k>0:
                bx,by=(1-i)*200,-170-(40 if i==1 else 0)
                circle(x,bx,by,62*k,GOLD if i==1 else CARD,a); circle(x,bx,by,62*k,WHITE,a*0.7,fill=False,lw=3)
                q=sprite("؟",70,(14,32,80) if i==1 else (246,196,72),b'Black'); blit(x,q,bx,by-4,k,a)
    # sunrise: start your day
    ks=seg(t,s11-0.2,s11+0.3)
    if ks>0:
        a=eo(ks); rise=eo(seg(t,s11-0.1,ph(11)['end']+0.3)); cy=330-260*rise; R=120
        x.save(); x.new_path(); x.rectangle(-600,-600,1200,930); x.clip()
        radial(x,0,cy,420,GOLD,0.5*a)
        for k in range(20):
            ang=k/20*2*math.pi+t*0.2; r1=R+24; r2=R+80+30*math.sin(t*2+k)
            poly(x,[(r1*math.cos(ang),cy+r1*math.sin(ang)),(r2*math.cos(ang),cy+r2*math.sin(ang))],GOLD2,a*0.7,fill=False,lw=5,close=False)
        circle(x,0,cy,R,GOLD2,a); x.restore()
        glowline(x,[(-420,330),(420,330)],GOLD,a,4)
def scene_4(x,t):
    s12,s13,s14=[ph(i)['start'] for i in (12,13,14)]
    cur=0 if t<s13-0.1 else 1 if t<s14-0.1 else 2
    xs=[220,0,-220]; ny=-310; kin=eo(seg(t,s12-0.4,s12))
    pr=[seg(t,s13-0.3,s13+0.2),seg(t,s14-0.3,s14+0.2)]
    for i in range(2):
        poly(x,[(xs[i]-60,ny),(xs[i+1]+60,ny)],SOFT,0.35*kin,fill=False,lw=6,close=False)
        if pr[i]>0: glowline(x,partial([(xs[i]-60,ny),(xs[i+1]+60,ny)],eo(pr[i])),GOLD,kin,6)
    for i in range(3): num_node(x,xs[i],ny,52,i+1,kin,i==cur,1)
    # 1. biggest impact task
    k1=eo(seg(t,s12-0.2,s12+0.3)); k1o=eio(seg(t,s13-0.2,s13+0.25))
    if k1>0 and k1o<1:
        a=k1*(1-k1o); base=300; hs=[150,220,380,180]; hi=2
        for i,h in enumerate(hs):
            g=back(seg(t,wt(12,2)-0.2+i*0.12,wt(12,2)+0.25+i*0.12),1.2)
            win=seg(t,wt(12,5)-0.1,wt(12,5)+0.4)
            if g>0:
                bx=210-i*140; hh=h*g
                c=GOLD if (i==hi and win>0) else CARD
                rrect(x,bx-55,base-hh,110,hh,14,c,a*(1 if i==hi or win==0 else 1-0.5*win)); rrect(x,bx-55,base-hh,110,hh,14,WHITE,a*0.4,fill=False,lw=3)
                if i==hi and win>0:
                    radial(x,bx,base-hh-60,150,GOLD,0.5*a*win); poly(x,star_pts(bx,base-hh-60,46*back(win),rot=-math.pi/2+t*0.4),GOLD2,a)
        glowline(x,[(-400,base),(300,base)],WHITE,a,4)
    # 2. delegate / drop
    k2=eo(seg(t,s13-0.2,s13+0.3)); k2o=eio(seg(t,s14-0.2,s14+0.25))
    if k2>0 and k2o<1:
        a=k2*(1-k2o)
        for i in range(3): card_row(x,-170,-120+i*90,340,a,0.7,gold=(i==0))
        kd=eio(seg(t,wt(13,5)-0.2,wt(13,5)+0.5))   # تفويضه -> to person
        person(x,-300,230,1.3,WHITE,a); 
        card_row(x,lerp(-170,-380,kd),lerp(60,110,kd),lerp(340,160,kd),a*(1-0.3*kd),0.6)
        kt=eio(seg(t,wt(13,7)-0.2,wt(13,7)+0.5))   # الاستغناء -> trash
        bx,by=280,250
        rrect(x,bx-55,by-60,110,130,14,SOFT,a,fill=False,lw=6); rrect(x,bx-70,by-82,140,16,6,SOFT,a)
        for q in (-25,0,25): poly(x,[(bx+q,by-35),(bx+q,by+45)],SOFT,a,fill=False,lw=4,close=False)
        card_row(x,lerp(-170,bx-60,kt),lerp(150,by-110,kt)+60*math.sin(kt*math.pi)*-1,lerp(340,120,kt),a*(1-kt*0.9),0.5)
    # 3. real result: ring to 100 + check
    k3=eo(seg(t,s14-0.2,s14+0.3))
    if k3>0:
        a=k3; kp=eio(seg(t,wt(14,3),ph(14)['end']+0.2)); cy=40
        circle(x,0,cy,190,CARD,a); circle(x,0,cy,190,WHITE,a*0.2,fill=False,lw=26)
        x.new_path(); x.arc(0,cy,190,-math.pi/2,-math.pi/2+2*math.pi*max(0.001,kp)); x.set_line_width(26); x.set_line_cap(cairo.LINE_CAP_ROUND); rgba(GOLD,a,x); x.stroke()
        if kp<1:
            sp=sprite(nums(int(kp*100))+'٪',90,(255,255,255),b'Black'); blit(x,sp,0,cy-4,1,a)
        else:
            kc=seg(t,ph(14)['end']+0.2,ph(14)['end']+0.7); radial(x,0,cy,300,GOLD,0.4*a*kc)
            check(x,0,cy-10,4.2*back(kc),eo(kc),GOLD2,a,lw=16)
def scene_5(x,t):
    s15,s16,s17,s18,s19=[ph(i)['start'] for i in (15,16,17,18,19)]
    kn=back(seg(t,s15-0.3,s15+0.2)); kno=eio(seg(t,s16-0.15,s16+0.3))
    if kn>0 and kno<1:
        a=1-kno; x.save(); x.translate(0,20); x.rotate(-0.05); x.scale(S_(kn),S_(kn))
        rrect(x,-170,-160,340,320,16,(0,0,0),0.25*a); rrect(x,-170,-170,340,320,16,GOLD2,a)
        for i in range(3): rrect(x,-120,-60+i*60,240-i*50,14,7,NAVY2,0.5*a)
        x.restore()
        kp=eo(seg(t,s15,s15+0.35)); py=lerp(-480,-150,kp)
        circle(x,0,py,34,(0.85,0.2,0.25),a); circle(x,-10,py-10,10,WHITE,0.6*a); poly(x,[(0,py+30),(0,py+70)],SOFT,a,fill=False,lw=6,close=False)
    # many checked tasks -> one diamond of value
    kg=seg(t,s16-0.2,s16+0.2); col=eio(seg(t,s17-0.1,s17+0.6))
    if kg>0 and t<s18+0.3:
        n=0
        for r in range(5):
            for c in range(7):
                j=r*7+c; k=back(seg(t,s16+j*0.04,s16+0.25+j*0.04))
                if k<=0: continue
                n+=1; bx,by=-300+c*100,-230+r*100; px,py=lerp(bx,0,col),lerp(by,0,col)
                dimk=eio(seg(t,wt(16,4)-0.1,wt(16,4)+0.4))
                al=(1-0.5*dimk)*(1-col)
                rrect(x,px-30*k,py-30*k,60*k,60*k,12,CARD,al); check(x,px,py,0.9*k,1,WHITE,al,lw=5)
        if col<1: sp=sprite(nums(n),96,(255,255,255),b'Black'); blit(x,sp,0,-340,1-col,1-col)
    kd=seg(t,s17-0.1,s17+0.8); kdo=eio(seg(t,s18-0.1,s18+0.4))
    if kd>0 and kdo<1:
        r=lerp(10,170,back(seg(t,wt(17,3)-0.15,wt(17,3)+0.6),1.2)); diamond(x,0,0,r,1-kdo,t)
        sp=sprite("النتائج",44,(255,255,255),b'Black'); blit(x,sp,0,r+70,seg(t,wt(17,3),wt(17,3)+0.4),1-kdo)
    # busy gears slowing & fading
    kb=seg(t,s18-0.2,s18+0.3); kbo=eio(seg(t,s19-0.1,s19+0.4))
    if kb>0 and kbo<1:
        a=eo(kb)*(1-kbo)*lerp(1,0.4,seg(t,wt(18,1),ph(18)['end']))
        slow=lerp(3,0.2,eio(seg(t,s18,ph(18)['end']+0.2)))
        gear(x,-110,0,110,t*slow,SOFT,a); gear(x,90,-80,70,-t*slow*1.6,SOFT,a,teeth=8); gear(x,100,110,56,-t*slow*2,SOFT,a,teeth=8)
        ban(x,0,0,250,seg(t,wt(18,1)-0.1,wt(18,1)+0.5),a,lw=13)
    # achieve what matters
    kz=seg(t,s19-0.15,s19+0.4)
    if kz>0:
        e=back(kz,1.3); radial(x,0,-10,420*e,GOLD,0.5)
        for i in range(16):
            ang=i*math.pi/8+t*0.25; r1=230*e; r2=r1+70+25*math.sin(t*3+i)
            poly(x,[(r1*math.cos(ang),-10+r1*math.sin(ang)),(r2*math.cos(ang),-10+r2*math.sin(ang))],GOLD2,0.75,fill=False,lw=6,close=False)
        circle(x,0,-10,200*e,GOLD,1); circle(x,0,-10,200*e,WHITE,1,fill=False,lw=6)
        check(x,0,-20,4.6*e,eo(seg(t,wt(19,1)-0.1,wt(19,1)+0.5)),NAVY1,1,lw=18)
        kst=seg(t,wt(19,3)-0.1,wt(19,3)+0.9)
        for i in range(12):
            a2=i*math.pi/6+0.2; r=230+200*eo(kst)
            poly(x,star_pts(r*math.cos(a2),-10+r*math.sin(a2),14*(1-kst*0.7),rot=t),GOLD2,1-kst)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
