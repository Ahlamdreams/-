import numpy as np, wave, sys, os
sys.path.insert(0,os.path.dirname(os.path.abspath(__file__)))
import render as R, timeline as TL
R.P=TL.compute(); P=R.P
SR=44100; N=int(SR*R.DUR); L=np.zeros(N); rng=np.random.default_rng(2)
def add(sig,t,g):
    i=int(t*SR); n=min(len(sig),N-i)
    if 0<=i and n>0: L[i:i+n]+=sig[:n]*g
def pop(f=900):
    n=int(0.12*SR); tt=np.arange(n)/SR; return np.sin(2*np.pi*np.cumsum(f*np.exp(-tt/0.04)+200)/SR)*np.exp(-tt/0.03)
def whoosh(d=0.7):
    n=int(d*SR); x=rng.standard_normal(n); y=np.convolve(x,np.ones(12)/12,'same'); tt=np.arange(n)/SR
    return y*np.exp(-((tt-d*0.6)/(d*0.25))**2)
def ding(f=1320):
    n=int(1.2*SR); tt=np.arange(n)/SR; return (np.sin(2*np.pi*f*tt)+0.4*np.sin(2*np.pi*f*2.01*tt))*np.exp(-tt/0.35)
wt=lambda i,j:P[i]['words'][j]['start']
for si in range(1,5): add(whoosh(),R.scene_span(si)[0]+0.4-0.45,0.5)
add(whoosh(0.9),R.OUTRO-0.5,0.6)
for i,p in enumerate(P): add(pop(700+40*(i%5)),p['start']-0.05,0.10)
pe=lambda i:P[i]['end']
lw=lambda i:P[i]['words'][-1]['start']
for t,f in ((lw(0),1319),(lw(1),1175),(lw(2),1568),(lw(3),1568),(lw(4),1568),(lw(6),1760),(P[8]['words'][3]['start'],1319),(lw(9),1568),(lw(10),1568),(P[11]['words'][-3]['start'],1568),(P[13]['start'],1175),(P[14]['start'],1319),(P[15]['words'][3]['start'],1760),(P[17]['words'][-1]['start'],988),(P[18]['start'],1568),(P[19]['words'][2]['start'],988),(lw(20),1760),(R.VEND+0.2,1319),(R.VEND+1.0,1568),(R.OUTRO+0.2,1047)):
    add(ding(f),t,0.07)
add(pop(1100),R.OUTRO+1.5,0.1)
L=np.clip(L,-1,1); w=wave.open(sys.argv[1],'wb'); w.setnchannels(1); w.setsampwidth(2); w.setframerate(SR); w.writeframes((L*32767).astype(np.int16).tobytes()); w.close()
print('sfx ok',R.DUR,[round(R.scene_span(i)[0]+0.4,2) for i in range(5)])
