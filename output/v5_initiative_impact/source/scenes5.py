# ---------------- scenes (v5: initiatives & impact) ----------------
def ph(i): return P[i]
def wt(i,j): return P[i]['words'][j]['start']
def scene_span(si):
    a=P[TL.SCENES[si][1]]['start']-0.55
    b=P[TL.SCENES[si+1][1]]['start']-0.55 if si+1<len(TL.SCENES) else OUTRO
    return a,b
def S_(v): return max(1e-3,v)
def bulb(x,cx,cy,s,lit,a,t=0):
    if a<=0 or s<=0.01: return
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    if lit>0:
        radial(x,0,-40,260,GOLD,0.45*lit*a)
        for i in range(10):
            ang=-math.pi/2+(i-4.5)*0.32; r0=150; r1=150+40*lit+8*math.sin(t*6+i)
            poly(x,[(r0*math.cos(ang),-40+r0*math.sin(ang)),(r1*math.cos(ang),-40+r1*math.sin(ang))],GOLD2,a*lit,fill=False,lw=7,close=False)
    x.new_path(); x.arc(0,-40,105,math.radians(140),math.radians(40)+2*math.pi*0); 
    x.new_path(); x.arc(0,-40,105,math.radians(135),math.radians(405)); x.line_to(48,90); x.line_to(-48,90); x.close_path()
    path=x.copy_path(); rgba(mix(CARD,GOLD2,lit),a,x); x.fill()
    x.append_path(path); x.set_line_width(8); rgba(WHITE,a,x); x.stroke()
    for k in range(3): rrect(x,-50+k*4,98+k*22,100-k*8,16,8,WHITE,a)
    poly(x,[(-25,60),(-12,0),(0,30),(12,0),(25,60)],NAVY1 if lit>0.5 else WHITE,a*0.8,fill=False,lw=5,close=False)
    x.restore()
def check(x,cx,cy,s,k,c,a,lw=8):
    if k>0: glowline(x,partial([(cx-22*s,cy),(cx-6*s,cy+16*s),(cx+24*s,cy-18*s)],k),c,a,lw)
def magnifier(x,cx,cy,s,a):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    circle(x,0,0,80,CARD,a*0.8); circle(x,0,0,80,WHITE,a,fill=False,lw=12)
    poly(x,[(58,58),(130,130)],GOLD,a,fill=False,lw=24,close=False)
    q=sprite("؟",110,(246,196,72),b'Black'); blit(x,q,0,4,1,a)
    x.restore()
def trophy(x,cx,cy,s,a,t):
    x.save(); x.translate(cx,cy); x.scale(S_(s),S_(s))
    radial(x,0,-20,240,GOLD,0.35*a)
    x.new_path(); x.move_to(-90,-120); x.line_to(90,-120); x.curve_to(90,10,40,50,0,55); x.curve_to(-40,50,-90,10,-90,-120); x.close_path(); rgba(GOLD,a,x); x.fill()
    for sg in (-1,1):
        x.new_path(); x.arc(sg*92,-70,38,math.pi/2 if sg>0 else -math.pi/2, -math.pi/2 if sg>0 else math.pi/2) if False else None
        x.new_path(); x.arc_negative(sg*90,-75,36,math.pi*0.5,math.pi*1.5) if sg<0 else x.arc(sg*90,-75,36,-math.pi*0.5,math.pi*0.5)
        x.set_line_width(12); rgba(GOLD,a,x); x.stroke()
    rrect(x,-14,50,28,50,6,GOLD,a); rrect(x,-60,96,120,26,10,GOLD2,a); rrect(x,-75,118,150,22,8,WHITE,a)
    poly(x,star_pts(0,-50,32,rot=-math.pi/2),WHITE,a)
    for i in range(5):
        ang=t*0.8+i*1.26; r=170+10*math.sin(t*3+i); poly(x,star_pts(r*math.cos(ang),-30+r*0.6*math.sin(ang),10,rot=t+i),GOLD2,a*0.8)
    x.restore()
def tile(x,cx,cy,w,h,a,active):
    rrect(x,cx-w/2,cy-h/2+8,w,h,28,(0,0,0),0.25*a); rrect(x,cx-w/2,cy-h/2,w,h,28,CARD,a)
    rrect(x,cx-w/2,cy-h/2,w,h,28,GOLD if active else (0.35,0.47,0.75),a,fill=False,lw=5 if active else 2.5)
def num_node(x,cx,cy,r,n,a,active,k=1):
    if k<=0: return
    rr=r*k*(1.18 if active else 1)
    if active: radial(x,cx,cy,rr*2.6,GOLD,0.35*a)
    circle(x,cx,cy,rr,GOLD if active else CARD,a); circle(x,cx,cy,rr,WHITE if not active else GOLD2,a,fill=False,lw=4)
    sp=sprite('٠١٢٣٤٥'[n],int(r*1.1),(14,32,80) if active else (255,255,255),b'Black'); blit(x,sp,cx,cy-r*0.08,k*(1.18 if active else 1),a)
def scene_1(x,t):
    s0,s1,s2=ph(0)['start'],ph(1)['start'],ph(2)['start']
    up=eio(seg(t,s1-0.1,s1+0.5))
    kb=back(seg(t,s0-0.45,s0+0.15)); lit=eo(seg(t,wt(0,2)-0.1,wt(0,2)+0.5))
    if kb>0:
        cx,cy,s=lerp(0,0,up),lerp(-20,-250,up),kb*lerp(1.15,0.55,up)
        bulb(x,cx,cy,s,lit,1,t)
        for i in range(4):
            k=back(seg(t,wt(0,3)-0.1+i*0.1,wt(0,3)+0.3+i*0.1))*(1-up)
            ang=t*1.2+i*math.pi/2; circle(x,cx+250*math.cos(ang),cy-40+120*math.sin(ang),16*k,WHITE if i%2 else GOLD,1)
    # ripple impact
    kr=seg(t,s1-0.2,s1+0.3); kro=eio(seg(t,s2-0.1,s2+0.4))
    if kr>0 and kro<1:
        a=1-kro; dy=lerp(-200,120,eo(seg(t,s1-0.2,s1+0.25)))
        if t<s1+0.25: circle(x,0,dy,18,GOLD2,a)
        for k in range(5):
            pr=cl((t-s1-0.25)*0.7-k*0.18)%1 if t>s1+0.25 else 0
            if pr>0:
                x.save(); x.translate(0,120); x.scale(1,0.3); circle(x,0,0,30+pr*360,GOLD if k%2==0 else WHITE,(1-pr)*a,fill=False,lw=8); x.restore()
        kq=back(seg(t,wt(1,1)-0.1,wt(1,1)+0.35))
        if kq>0: q=sprite("؟",200,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,270,-40,kq*(1+0.05*math.sin(t*5)),a)
    # gauge (ليس بالضرورة)
    kg=back(seg(t,s2-0.2,s2+0.35))
    if kg>0:
        cx,cy,R=0,230,250
        for i in range(30):
            a0=math.pi+i*math.pi/30; a1=a0+math.pi/30-0.03; c=mix(WHITE,GOLD,i/29)
            x.new_path(); x.arc(cx,cy,R*kg,a0,a1); x.set_line_width(30); rgba(c,0.9,x); x.stroke()
        sw=seg(t,s2,ph(2)['end']+0.4); ang=math.pi*1.5+math.sin(sw*14)*0.9*(1-sw)
        L=R*0.82*kg; poly(x,[(cx,cy),(cx+L*math.cos(ang),cy+L*math.sin(ang))],GOLD2,1,fill=False,lw=10,close=False)
        circle(x,cx,cy,24*kg,GOLD,1); circle(x,cx,cy,10*kg,NAVY1,1)
        ke=back(seg(t,wt(2,1),wt(2,1)+0.4))
        if ke>0: sp=sprite("!",150,(246,196,72),b'Black',glow=(245,190,60),rtl=False); blit(x,sp,0,cy-330,ke,1)
def scene_2(x,t):
    s3,s4,s5,s6,s7=[ph(i)['start'] for i in (3,4,5,6,7)]
    dim=eio(seg(t,s6-0.1,s6+0.4)); shift=-60*dim
    xs=[250,0,-250]; ty=-200+shift
    act=[s3<=t<s4,s4<=t<s5,s5<=t<s6]
    for i,(cx,st) in enumerate(zip(xs,(s3,s4,s5))):
        k=back(seg(t,st-0.3,st+0.2))
        if k<=0: continue
        a=1-0.45*dim
        x.save(); x.translate(cx,ty); x.scale(S_(k),S_(k))
        tile(x,0,0,210,210,a,act[i])
        if i==0: bulb(x,0,10,0.55,eo(seg(t,wt(3,2),wt(3,2)+0.4)),a,t)
        elif i==1:
            kd=eo(seg(t,s4,wt(4,1)+0.5)); pts=arcp(0,0,70,-math.pi/2,-math.pi/2+2*math.pi*kd,60)
            glowline(x,pts,GOLD,a,7)
            ks=back(seg(t,wt(4,1),wt(4,1)+0.4)); 
            if ks>0: poly(x,star_pts(0,0,42*ks,rot=-math.pi/2),WHITE,a)
            if kd<1: pe=pts[-1]; poly(x,[(pe[0],pe[1]),(pe[0]+22,pe[1]-40)],WHITE,a,fill=False,lw=8,close=False)
        else:
            for r in range(3):
                for c in range(3):
                    j=r*3+c; kk=seg(t,s5+0.1+j*0.09,s5+0.35+j*0.09)
                    bx,by=-66+c*66,-62+r*62; rrect(x,bx-26,by-24,52,48,10,WHITE,a*0.15+a*0.1*kk)
                    check(x,bx,by,0.9,eo(kk),GOLD,a,lw=6)
        x.restore()
    # magnifier (السؤال الأهم)
    km=back(seg(t,wt(6,2)-0.15,wt(6,2)+0.3)); kmo=eio(seg(t,s7-0.1,s7+0.3))
    if km>0 and kmo<1:
        magnifier(x,140*math.sin(t*1.8),130,km*(1-kmo),1)
    # before / after (ما الذي تغيّر فعلًا؟)
    kc=seg(t,s7-0.2,s7+0.3)
    if kc>0:
        a=eo(kc); base=330
        glowline(x,[(-280,base),(280,base)],WHITE,a,4)
        h1=200*back(seg(t,s7,s7+0.5)); h2=200*back(seg(t,wt(7,2)-0.1,wt(7,2)+0.45))
        rrect(x,-200,base-h1,140,h1,14,SOFT,a); rrect(x,60,base-h2,140,h2,14,GOLD,a)
        for lab,lx in (("قبل",-130),("بعد",130)):
            sp=sprite(lab,44,(255,255,255),b'Bold'); blit(x,sp,lx,base+40,1,a)
        if h2>190:
            kq=back(seg(t,wt(7,3)-0.1,wt(7,3)+0.35))
            for j in range(0,10): poly(x,[(-130+j*26,base-215),(-118+j*26,base-215)],WHITE,a*0.8,fill=False,lw=4,close=False)
            q=sprite("؟",170,(246,196,72),b'Black',glow=(245,190,60)); blit(x,q,0,base-330,kq*(1+0.05*math.sin(t*5)),a)
            eq=sprite("=",90,(255,255,255),b'Black',rtl=False); blit(x,eq,0,base-110,kq,a)
def scene_3(x,t):
    s8,s9,s10=[ph(i)['start'] for i in (8,9,10)]
    kt=back(seg(t,s8-0.35,s8+0.25)); up=eio(seg(t,s9-0.1,s9+0.45)); kto=eio(seg(t,s10-0.1,s10+0.4))
    if kt>0 and kto<1:
        trophy(x,lerp(0,250,up),lerp(-20,-260,up),kt*lerp(1.1,0.55,up)*(1-kto),1,t)
    # checklist card crossed (ماذا سننفذ؟)
    kc=back(seg(t,s9-0.2,s9+0.3)); kco=eo(seg(t,s10-0.1,s10+0.5))
    if kc>0 and kco<1:
        x.save(); x.translate(lerp(-40,-700,kco),70); x.rotate(-0.06+0.25*kco); x.scale(S_(kc),S_(kc))
        rrect(x,-190,-190,380,380,30,(0,0,0),0.25); rrect(x,-190,-200,380,380,30,WHITE,0.96)
        rrect(x,-190,-200,380,70,30,GOLD,1)
        for i in range(4):
            y=-90+i*70; rrect(x,110,y-18,36,36,8,NAVY2,0.8,fill=False,lw=4); rrect(x,-150,y-8,240-i*30,16,8,SOFT,0.9)
            check(x,128,y,0.7,eo(seg(t,s9+0.2+i*0.12,s9+0.45+i*0.12)),NAVY2,1,lw=6)
        kx=seg(t,wt(9,1)+0.05,wt(9,1)+0.5)
        if kx>0:
            glowline(x,partial([(-170,-180),(170,160)],eo(kx)),GOLD,1,18); glowline(x,partial([(170,-180),(-170,160)],eo(cl(kx*1.4-0.3))),GOLD,1,18)
        x.restore()
    # three questions
    for i in range(3):
        k=back(seg(t,wt(10,2)-0.1+i*0.18,wt(10,2)+0.3+i*0.18))
        if k<=0: continue
        cx=(1-i)*230; cy=60
        if i<2:
            kl=eo(seg(t,wt(10,2)+0.2+i*0.18,wt(10,2)+0.5+i*0.18))
            if kl>0: glowline(x,partial([(cx-70,cy),(cx-160,cy)],kl),GOLD,1,5)
        num_node(x,cx,cy,72,i+1,1,False,k)
        q=sprite("؟",64,(246,196,72),b'Black'); blit(x,q,cx,cy-120,k*(1+0.08*math.sin(t*4+i)),1)
def scene_4(x,t):
    s11,s12,s13=[ph(i)['start'] for i in (11,12,13)]
    cur=0 if t<s12-0.1 else 1 if t<s13-0.1 else 2
    xs=[220,0,-220]; ny=-310
    kin=eo(seg(t,s11-0.4,s11))
    pr=[seg(t,s12-0.3,s12+0.2),seg(t,s13-0.3,s13+0.2)]
    for i in range(2):
        poly(x,[(xs[i]-60,ny),(xs[i+1]+60,ny)],SOFT,0.35*kin,fill=False,lw=6,close=False)
        if pr[i]>0: glowline(x,partial([(xs[i]-60,ny),(xs[i+1]+60,ny)],eo(pr[i])),GOLD,kin,6)
    for i in range(3): num_node(x,xs[i],ny,52,i+1,kin,i==cur,1)
    # 1. problem -> puzzle
    k1=eo(seg(t,s11-0.2,s11+0.35)); k1o=eio(seg(t,s12-0.2,s12+0.25))
    if k1>0 and k1o<1:
        a=k1*(1-k1o); sz=150; ox,oy=-sz,-40
        for (c,r) in ((0,0),(1,0),(0,1)):
            rrect(x,ox+c*sz+6,oy+r*sz+6,sz-12,sz-12,18,CARD,a); rrect(x,ox+c*sz+6,oy+r*sz+6,sz-12,sz-12,18,WHITE,a*0.7,fill=False,lw=3)
        kw=back(seg(t,wt(11,1)-0.1,wt(11,1)+0.35))
        if kw>0:
            tri=[(ox+sz*0.5,oy+sz*0.15),(ox+sz*0.88,oy+sz*0.8),(ox+sz*0.12,oy+sz*0.8)]
            poly(x,[(lerp(ox+sz/2,p[0],kw),lerp(oy+sz/2,p[1],kw)) for p in tri],GOLD,a)
            sp=sprite("!",70,(14,32,80),b'Black',rtl=False); blit(x,sp,ox+sz/2,oy+sz*0.58,kw,a)
        # empty slot then piece slides in
        sx,sy=ox+sz+6,oy+sz+6
        for j in range(12):
            u=j/12; per=4*(sz-12); d=u*per; L=sz-12
            p=(sx+d,sy) if d<L else (sx+L,sy+d-L) if d<2*L else (sx+L-(d-2*L),sy+L) if d<3*L else (sx,sy+L-(d-3*L))
            circle(x,p[0],p[1],4,GOLD2,a*0.8)
        ks=eio(seg(t,wt(11,3)-0.3,ph(11)['end']+0.1))
        if ks>0:
            px=lerp(sx+360,sx,ks); rrect(x,px,sy,sz-12,sz-12,18,GOLD,a)
            if ks>=1: check(x,sx+(sz-12)/2,sy+(sz-12)/2,1.4,eo(seg(t,ph(11)['end']+0.1,ph(11)['end']+0.4)),NAVY1,a,lw=9)
    # 2. change: A -> B
    k2=eo(seg(t,s12-0.2,s12+0.35)); k2o=eio(seg(t,s13-0.2,s13+0.25))
    if k2>0 and k2o<1:
        a=k2*(1-k2o)
        circle(x,230,90,70,SOFT,a*0.9); sp=sprite("قبل",40,(14,32,80),b'Bold'); blit(x,sp,230,86,1,a)
        ka=seg(t,wt(12,1)-0.1,wt(12,1)+0.7)
        pts=[(lerp(160,-130,u),90-200*math.sin(u*math.pi)) for u in np.linspace(0,1,40)]
        pp=partial(pts,eo(ka)); glowline(x,pp,GOLD2,a,8)
        if len(pp)>1 and ka>0: arrow_shape(x,pp[-2][0],pp[-2][1],pp[-1][0],pp[-1][1],GOLD2,a,lw=1,head=34)
        kb=back(seg(t,wt(12,4)-0.15,wt(12,4)+0.4))
        if kb>0:
            radial(x,-230,90,240*kb,GOLD,0.45*a); circle(x,-230,90,105*kb,GOLD,a)
            sp=sprite("بعد",46,(14,32,80),b'Black'); blit(x,sp,-230,86,kb,a)
            for i in range(8):
                ang=i*math.pi/4+t*0.6; poly(x,star_pts(-230+150*kb*math.cos(ang),90+150*kb*math.sin(ang),10*kb,rot=t),GOLD2,a)
    # 3. measure: bars + ruler + counter
    k3=eo(seg(t,s13-0.2,s13+0.35))
    if k3>0:
        a=k3; base=300
        poly(x,[(300,base),(300,-140)],WHITE,a,fill=False,lw=5,close=False)
        for j in range(0,11):
            y=base-j*42; poly(x,[(300,y),(300-(30 if j%5==0 else 16),y)],WHITE,a,fill=False,lw=4,close=False)
        hs=[150,250,360]
        for i,h in enumerate(hs):
            g=back(seg(t,wt(13,2)-0.1+i*0.2,wt(13,2)+0.4+i*0.2),1.2)
            if g>0: rrect(x,-240+i*150,base-h*g,110,h*g,14,GOLD if i==2 else CARD,a); rrect(x,-240+i*150,base-h*g,110,h*g,14,WHITE,a*0.4,fill=False,lw=3)
        kp=eio(seg(t,wt(13,2),ph(13)['end']+0.3)); v=int(kp*85)
        sp=sprite(''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(v))+'٪',96,(246,196,72),b'Black',glow=(245,190,60)); blit(x,sp,-100,-120,1,a)
        glowline(x,partial([(-185,base-170),(-35,base-270),(115,base-380)],eo(seg(t,wt(13,3),ph(13)['end']+0.4))),GOLD2,a,6)
def scene_5(x,t):
    s14,s15,s16,s17,s18=[ph(i)['start'] for i in (14,15,16,17,18)]
    # pin + note (تذكّر)
    kn=back(seg(t,s14-0.3,s14+0.2)); kno=eio(seg(t,s15-0.15,s15+0.3))
    if kn>0 and kno<1:
        a=1-kno; x.save(); x.translate(0,20); x.rotate(-0.05); x.scale(S_(kn),S_(kn))
        rrect(x,-170,-160,340,320,16,(0,0,0),0.25*a); rrect(x,-170,-170,340,320,16,GOLD2,a)
        for i in range(3): rrect(x,-120,-60+i*60,240-i*50,14,7,NAVY2,0.5*a)
        x.restore()
        kp=eo(seg(t,s14,s14+0.35)); py=lerp(-480,-150,kp)+(-14*math.sin(seg(t,s14+0.35,s14+0.7)*math.pi))
        circle(x,0,py,34,(0.85,0.2,0.25),a); circle(x,-10,py-10,10,WHITE,0.6*a); poly(x,[(0,py+30),(0,py+70)],SOFT,a,fill=False,lw=6,close=False)
    # many small initiatives (عدد المبادرات)
    kg=seg(t,s15-0.2,s15+0.2); col=eio(seg(t,s16-0.1,s16+0.6))
    if kg>0 and t<s17:
        n=0
        for r in range(5):
            for c in range(7):
                j=r*7+c; k=back(seg(t,s15+j*0.035,s15+0.25+j*0.035))
                if k<=0: continue
                n+=1; bx,by=-300+c*100,-230+r*100
                dimk=eio(seg(t,wt(15,4)-0.1,wt(15,4)+0.4))
                px,py=lerp(bx,0,col),lerp(by,0,col)
                rrect(x,px-30*k,py-30*k,60*k,60*k,12,mix(SOFT,WHITE,0.3),(1-0.55*dimk)*(1-col))
        if col<1:
            sp=sprite(''.join('٠١٢٣٤٥٦٧٨٩'[int(c)] for c in str(n)),96,(255,255,255),b'Black'); blit(x,sp,0,-340,1-col,1-col)
    # one big impact (حجم الأثر)
    kb=seg(t,s16-0.1,s16+1.0); kbo=eio(seg(t,s17-0.1,s17+0.4))
    if kb>0 and kbo<1:
        a=1-kbo; r=lerp(20,230,back(seg(t,wt(16,2)-0.1,wt(16,2)+0.7),1.2))
        for k in range(4):
            pr=((t-s16)*0.6+k/4)%1; circle(x,0,0,r+pr*260,GOLD,(1-pr)*0.6*a,fill=False,lw=6)
        radial(x,0,0,r*1.8,GOLD,0.4*a); circle(x,0,0,r,GOLD,a)
        sp=sprite("الأثر",int(max(20,r*0.55)),(14,32,80),b'Black'); blit(x,sp,0,-6,1,a*seg(t,wt(16,3),wt(16,3)+0.3))
    # activity that ends (نشاط ينتهي)
    kl=seg(t,s17-0.1,wt(17,5)+0.1); klo=eio(seg(t,s18+0.6,s18+1.1))
    if kl>0 and klo<1:
        a=1-klo; y=80
        glowline(x,[(380,y),(lerp(380,-120,eo(kl)),y)],WHITE,a*0.9,8)
        circle(x,380,y,16,GOLD,a)
        if kl>=1:
            ke=back(seg(t,wt(17,5),wt(17,5)+0.35)); rrect(x,-132,y-60*ke,24,120*ke,8,GOLD,a)
            sp=sprite("نهاية",42,(255,255,255),b'Bold'); blit(x,sp,-120,y+100,ke,a*0.85)
    # infinity loop (يبدأ ويستمر)
    ki=seg(t,s18-0.1,s18+1.4)
    if ki>0:
        pts=[(260*math.sin(u),-120+150*math.sin(u)*math.cos(u)) for u in np.linspace(0,2*math.pi,160)]
        glowline(x,pts,GOLD,1,14,k=eo(ki),close=True)
        if ki>=1:
            for k in range(3):
                u=((t-s18)*0.35+k/3)%1; i=int(u*159); p=pts[i]
                radial(x,p[0],p[1],60,GOLD2,0.7); circle(x,p[0],p[1],12,WHITE,1)
        kpl=back(seg(t,wt(18,3)-0.1,wt(18,3)+0.4))
        if kpl>0:
            for k in range(3):
                pr=((t-wt(18,3))*0.5+k/3)%1; circle(x,0,-120,60+pr*420,GOLD,(1-pr)*0.4*kpl,fill=False,lw=4)
SCENE_FN=[scene_1,scene_2,scene_3,scene_4,scene_5]
