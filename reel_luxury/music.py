import numpy as np, wave
sr=44100; D=54.0; N=int(sr*D); out=np.zeros((N,2))
def f(m): return 440*2**((m-69)/12)
def add(i,s,pan=0):
    if i>=N: return
    e=min(N,i+len(s)); s=s[:e-i]; out[i:e,0]+=s*(1-pan); out[i:e,1]+=s*(1+pan)
def piano(t0,m,dur,amp,pan=0):
    t=np.arange(int(dur*sr))/sr; fr=f(m)
    s=(np.sin(2*np.pi*fr*t)+.4*np.sin(4*np.pi*fr*t)*np.exp(-t*2)+.15*np.sin(6*np.pi*fr*t)*np.exp(-t*4))*np.exp(-t*1.6)*np.minimum(1,t*200)
    add(int(t0*sr),s*amp,pan)
def pad(t0,ms,dur,amp):
    t=np.arange(int(dur*sr))/sr; env=np.minimum(1,t/1.2)*np.minimum(1,(dur-t)/1.2)
    s=sum(np.sin(2*np.pi*f(m)*t+np.sin(2*np.pi*.2*t))+.5*np.sin(2*np.pi*f(m)*1.004*t)+.3*np.sin(2*np.pi*f(m)*2*t) for m in ms)*env*amp
    add(int(t0*sr),s)
bpm=72; beat=60/bpm; bar=4*beat
# D minor-ish elegant: Dm Bb F C  / Gm Bb F A
prog=[(50,[62,65,69]),(46,[58,62,65]),(53,[60,65,69]),(48,[60,64,67]),(55,[58,62,67]),(46,[58,62,65]),(53,[60,65,69]),(45,[61,64,69])]
for b in range(int(D/bar)+1):
    root,ch=prog[b%8]; t0=b*bar
    pad(t0,[m-12 for m in ch],bar+1.2,.03)
    piano(t0,root-12,3.5,.18)
    arp=[ch[0],ch[1],ch[2],ch[1]+12,ch[2],ch[1]]
    for k,m in enumerate(arp): piano(t0+k*beat*2/3,m,2.5,.06,pan=(-.4 if k%2 else .4))
    if b%2==1: piano(t0+2*beat,ch[2]+12,3,.07,pan=.2)
for ts in [0.3]:
    for j,m in enumerate([86,89,93,98]): piano(ts+j*.09,m,2.5,.05,pan=(j-1.5)/3)
d=int(.33*sr); out[d:]+=.3*out[:-d][:, ::-1]
out/=np.abs(out).max()
# ducking + mix voice
t=np.arange(N)/sr
g=np.interp(t,[0,2.0,2.6,49.6,50.6,51.5,54],[1,1,.2,.2,.75,.75,0])
music=out*g[:,None]*.55
vw=wave.open('voice.wav');vo=np.frombuffer(vw.readframes(vw.getnframes()),'<i2').reshape(-1,2)/32768.;vw.close()
vo=vo/np.abs(vo).max()*.92
mix=music.copy(); i=int(2.4*sr); e=min(N,i+len(vo)); mix[i:e]+=vo[:e-i]
mix/=max(1,np.abs(mix).max()/.97)
w=wave.open('mix.wav','wb');w.setnchannels(2);w.setsampwidth(2);w.setframerate(sr);w.writeframes((mix*32767).astype('<i2').tobytes());w.close()
print('ok')
