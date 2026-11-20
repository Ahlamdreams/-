import numpy as np, wave
sr=44100; D=54.0; N=int(sr*D); out=np.zeros((N,2)); rs=np.random.RandomState(3)
def add(t0,s,pan=0,amp=1):
    i=int(t0*sr)
    if i>=N: return
    e=min(N,i+len(s)); s=s[:e-i]*amp; out[i:e,0]+=s*(1-pan); out[i:e,1]+=s*(1+pan)
def oud(freq,dur=1.6,bright=.5):
    n=int(dur*sr); p=int(sr/freq); buf=rs.uniform(-1,1,p)
    # soften attack (plectrum)
    for _ in range(2): buf=(buf+np.roll(buf,1))/2
    y=np.zeros(n); 
    for i in range(n):
        y[i]=buf[i%p]; buf[i%p]=.996*(.5*(buf[i%p]+buf[(i+1)%p]))*(1-bright*.0)+0
    body=np.convolve(y,np.exp(-np.arange(60)/12)/6,'same')
    return (y*.7+body*.5)*np.minimum(1,np.arange(n)/40)
def dum():
    t=np.arange(int(.5*sr))/sr; f=55+70*np.exp(-t*25)
    return np.sin(2*np.pi*np.cumsum(f)/sr)*np.exp(-t*7)+rs.uniform(-1,1,len(t))*np.exp(-t*60)*.15
def tak():
    t=np.arange(int(.18*sr))/sr; n=rs.uniform(-1,1,len(t))
    n=n-np.concatenate([[0],n[:-1]])*.6
    return (n*.6+np.sin(2*np.pi*380*t)*.5)*np.exp(-t*38)
def clap():
    t=np.arange(int(.22*sr))/sr; n=rs.uniform(-1,1,len(t))
    env=sum(np.exp(-np.clip(t-d,0,None)*90)*(t>=d) for d in (0,.012,.024))*.6+np.exp(-t*25)*.5
    hp=n-np.concatenate([[0],n[:-1]])
    return hp*env*.5
def ney(freq,dur):
    t=np.arange(int(dur*sr))/sr; vib=1+.006*np.sin(2*np.pi*5.2*t)*np.minimum(1,t/.6)
    ph=2*np.pi*np.cumsum(freq*vib)/sr
    s=np.sin(ph)+.25*np.sin(2*ph)+.1*np.sin(3*ph)+rs.uniform(-1,1,len(t))*.06
    return s*np.minimum(1,t/.25)*np.minimum(1,(dur-t)/.3)
# Hijaz on D: D Eb F# G A Bb C D
def hz(m): return 440*2**((m-69)/12)
H=[62,63,66,67,69,70,72,74]
bpm=104; b=60/bpm
# rhythm (Omani-style 4/4 cycle, 2 bars): D=dum T=tak C=clap
pat=[('D',0),('T',1),('T',1.5),('D',2),('D',2.5),('T',3),('T',3.5),
     ('D',4),('T',5),('D',6),('T',6.5),('T',7),('T',7.5)]
claps=[1,3,5,7, 3.5,7.5]
cyc=8*b; nC=int(D/cyc)+1
for c in range(nC):
    t0=c*cyc
    if t0<1.5: continue
    for k,o in pat: add(t0+o*b, dum() if k=='D' else tak(), pan=(-.15 if k=='D' else .25), amp=.55 if k=='D' else .28)
    for o in claps: add(t0+o*b+rs.uniform(0,.01),clap(),pan=rs.uniform(-.4,.4),amp=.22)
# oud melody (phrases in Hijaz), call & answer
phr=[[(0,7),(.5,6),(1,5),(1.5,4),(2,3),(3,2),(3.5,1),(4,0),(6,4),(6.5,3),(7,2)],
     [(0,4),(.5,5),(1,4),(1.5,3),(2,2),(2.5,3),(3,4),(4,3),(4.5,2),(5,1),(5.5,2),(6,0)],
     [(0,0),(.5,2),(1,3),(1.5,4),(2,5),(3,4),(3.5,5),(4,7),(5,6),(5.5,5),(6,4),(7,3)],
     [(0,3),(.5,2),(1,1),(1.5,2),(2,3),(3,2),(3.5,1),(4,0),(5,0),(6,-1),(7,0)]]
for c in range(nC):
    t0=c*cyc; ph=phr[c%4]
    for o,d in ph:
        m=62 if d==-1 else H[d]
        if d==-1: m=60
        add(t0+o*b,oud(hz(m-12 if c%8>=4 else m)),pan=-.2,amp=.30)
    # bass drone on 1 and 5
    add(t0,oud(hz(50),2.5),amp=.35); add(t0+4*b,oud(hz(57),2.5),amp=.28)
# ney intro & outro
add(0.0,ney(hz(74),2.2),pan=.2,amp=.18); add(1.2,ney(hz(72),1.4),pan=.2,amp=.12)
add(49.6,ney(hz(69),1.6),pan=.2,amp=.16); add(51.0,ney(hz(74),2.6),pan=.2,amp=.18)
d=int(.12*sr); out[d:]+=.18*out[:-d][:, ::-1]
out/=np.abs(out).max()
t=np.arange(N)/sr
g=np.interp(t,[0,2.0,2.6,49.4,50.4,52.5,54],[1,1,.22,.22,.85,.85,0])
music=out*g[:,None]*.6
vw=wave.open('voice.wav');vo=np.frombuffer(vw.readframes(vw.getnframes()),'<i2').reshape(-1,2)/32768.;vw.close()
vo=vo/np.abs(vo).max()*.92
mix=music.copy(); i=int(2.4*sr); e=min(N,i+len(vo)); mix[i:e]+=vo[:e-i]
mix/=max(1,np.abs(mix).max()/.97)
w=wave.open('mix.wav','wb');w.setnchannels(2);w.setsampwidth(2);w.setframerate(sr);w.writeframes((mix*32767).astype('<i2').tobytes());w.close()
m2=out*np.interp(t,[0,52.5,54],[1,1,0])[:,None]*.9
w=wave.open('music_only.wav','wb');w.setnchannels(2);w.setsampwidth(2);w.setframerate(sr);w.writeframes((m2*32767).astype('<i2').tobytes());w.close()
print('ok')
