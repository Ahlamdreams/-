# Patch a render.py with varied scene transitions (iris, diagonal wipe, push, logo tiles).
import sys
p=sys.argv[1]; s=open(p).read()
old_loop=s[s.index("    for si in range(len(SCENE_FN)):\n        a,b=scene_span(si)"):s.index("    transition_ring(x,t); OVERLAY(x,t)")]
new_loop='''    draw_scenes(x,t)
'''
s=s.replace(old_loop,new_loop)
helper='''
TRANS=['iris','wipe','push','tiles']
def scene_pat(x,si,t,sc=1.0,dx=0.0):
    x.push_group(); x.save(); x.translate(VC[0]+dx,VC[1]); x.scale(sc,sc); SCENE_FN[si](x,t); x.restore(); return x.pop_group()
def draw_scenes(x,t):
    N=len(SCENE_FN)
    for i in range(1,N):
        B=scene_span(i)[0]+0.3; t0,t1=B-0.35,B+0.45
        if t0<=t<=t1:
            p=(t-t0)/(t1-t0); e=eio(p); kind=TRANS[(i-1)%len(TRANS)]
            pA=scene_pat(x,i-1,t,lerp(1,1.08,e) if kind!='push' else 1, -W*e if kind=='push' else 0)
            pB=scene_pat(x,i,t,lerp(0.92,1,e) if kind!='push' else 1, W*(1-e) if kind=='push' else 0)
            if kind=='iris':
                R=lerp(0,1500,eo(p)); cx,cy=VC
                x.save(); x.new_path(); x.rectangle(0,0,W,H); x.arc_negative(cx,cy,max(R,0.1),0,-2*math.pi); x.set_fill_rule(cairo.FILL_RULE_EVEN_ODD); x.clip(); x.set_source(pA); x.paint(); x.restore()
                x.save(); x.new_path(); x.arc(cx,cy,max(R,0.1),0,2*math.pi); x.clip(); x.set_source(pB); x.paint(); x.restore()
                if 0<p<1:
                    radial(x,cx,cy,R+60,GOLD,0.0)
                    circle(x,cx,cy,R,GOLD,0.9*(1-p*0.6),fill=False,lw=14*(1-p)+3); circle(x,cx,cy,R*0.86,WHITE,0.35*(1-p),fill=False,lw=3)
            elif kind=='wipe':
                f=lerp(-500,W+H*0.45+500,e)   # band moves right->left along x'=(W-x)+0.45y
                def region(lo,hi):
                    x.new_path()
                    pts=[]
                    for (yy) in (0,H):
                        pts.append((W-(lo-0.45*yy),yy))
                    pts2=[(W-(hi-0.45*yy),yy) for yy in (H,0)]
                    q=pts+pts2; x.move_to(*q[0])
                    for pt in q[1:]: x.line_to(*pt)
                    x.close_path()
                x.save(); region(f,99999); x.clip(); x.set_source(pA); x.paint(); x.restore()
                x.save(); region(-99999,f); x.clip(); x.set_source(pB); x.paint(); x.restore()
                if 0<p<1:
                    for wdt,al in ((120,0.18),(60,0.35),(18,0.95)):
                        x.save(); region(f-wdt/2,f+wdt/2); rgba(GOLD if wdt>18 else GOLD2,al,x); x.fill(); x.restore()
                    x.save(); region(f-wdt-4,f-wdt+2); rgba(WHITE,0.7,x); x.fill(); x.restore()
            elif kind=='push':
                x.set_source(pA); x.paint(); x.set_source(pB); x.paint()
                if 0.05<p<0.95:
                    k=math.sin(p*math.pi)
                    for j in range(9):
                        yy=260+j*95+20*math.sin(j*3.1); ln=lerp(120,420,(j%3)/2)*k; xx=lerp(W+200,-400,e)+((j*137)%300)
                        rrect(x,xx,yy,ln,6,3,GOLD if j%2 else WHITE,0.55*k)
            else:  # tiles
                cs=108; cols=int(W/cs)+1; rows=int(H/cs)+1; cx,cy=VC; dmax=math.hypot(W,H)*0.6
                flipped=[]; tiles=[]
                for r in range(rows):
                    for c in range(cols):
                        tx,ty=c*cs,r*cs; d=math.hypot(tx+cs/2-cx,ty+cs/2-cy)/dmax
                        tp=cl(p*1.7-d*0.7); tiles.append((tx,ty,tp,(r+c)%3))
                        if tp>=0.5: flipped.append((tx,ty))
                x.set_source(pA); x.paint()
                if flipped:
                    x.save(); x.new_path()
                    for (tx,ty) in flipped: x.rectangle(tx,ty,cs+0.5,cs+0.5)
                    x.clip(); x.set_source(pB); x.paint(); x.restore()
                for (tx,ty,tp,k) in tiles:
                    sz=math.sin(tp*math.pi)*cs*0.92
                    if sz>1: rrect(x,tx+cs/2-sz/2,ty+cs/2-sz/2,sz,sz,sz*0.18,[GOLD,(0.16,0.3,0.62),GOLD2][k],0.95)
            return
    # single scene (no boundary transition)
    for si in range(N):
        a,b=scene_span(si)
        if si==0:
            if t>scene_span(1)[0]+0.3 if N>1 else False: continue
            zin=eo(seg(t,a-0.1,a+0.55)); s=lerp(0.85,1,zin); al=zin
        elif si==N-1:
            if t<a+0.3: continue
            zout=eio(seg(t,b-0.15,b+0.45)); s=lerp(1,1.35,zout); al=1-zout
        else:
            if not (a+0.3<=t<=scene_span(si+1)[0]+0.3): continue
            s=1; al=1
        if si==0 and N>1 and t>scene_span(1)[0]+0.3: continue
        if al<=0: continue
        pat=scene_pat(x,si,t,s); x.set_source(pat); x.paint_with_alpha(al)
'''
s=s.replace("def transition_ring(x,t):",helper+"def transition_ring(x,t):",1)
# keep only the outro ring
s=s.replace("""    for si in range(1,len(TL.SCENES)):
        b=scene_span(si)[0]+0.4; k=seg(t,b-0.25,b+0.45)
        if 0<k<1: circle(x,VC[0],VC[1],60+900*eo(k),GOLD,0.55*(1-k),fill=False,lw=10*(1-k)+2)
""","")
open(p,'w').write(s); print('patched')
