import numpy as np, wave
sr=44100; D=37.0; N=int(sr*D); out=np.zeros((N,2))
def f(m): return 440*2**((m-69)/12)
def add(i,s,pan=0):
    if i>=N: return
    e=min(N,i+len(s)); s=s[:e-i]; out[i:e,0]+=s*(1-pan); out[i:e,1]+=s*(1+pan)
def pluck(t0,m,dur,amp,pan=0,bright=1.0):
    t=np.arange(int(dur*sr))/sr; fr=f(m)
    s=(np.sin(2*np.pi*fr*t)+.35*bright*np.sin(4*np.pi*fr*t)+.12*bright*np.sin(6*np.pi*fr*t))*np.exp(-t*3.2)*np.minimum(1,t*300)
    add(int(t0*sr),s*amp,pan)
def pad(t0,ms,dur,amp):
    t=np.arange(int(dur*sr))/sr; env=np.minimum(1,t/.6)*np.minimum(1,(dur-t)/.6)
    s=sum(np.sin(2*np.pi*f(m)*t)+.3*np.sin(2*np.pi*f(m)*1.003*t) for m in ms)*env*amp
    add(int(t0*sr),s)
bpm=96; beat=60/bpm; bar=4*beat
prog=[(60,[60,64,67]),(55,[55,59,62]),(57,[57,60,64]),(53,[53,57,60])]
mel=[[72,76,79,76,74,72,74,76],[74,71,67,71,74,76,74,71],[72,69,76,72,69,72,74,76],[77,76,74,72,74,72,69,67]]
for b in range(int(D/bar)+1):
    root,ch=prog[b%4]; t0=b*bar
    pad(t0,[m-12 for m in ch],bar+.3,.035)
    pluck(t0,root-24,1.5,.22); pluck(t0+2*beat,root-17,1.2,.15)
    arp=ch+[ch[1]+12]
    for k in range(8): pluck(t0+k*beat/2,arp[k%4]+(12 if k%2 else 0),0.9,.07,pan=-.3)
    if b>=2:
        for k,m in enumerate(mel[b%4]): pluck(t0+k*beat/2,m,1.0,.11 if k%2==0 else .08,pan=.3,bright=.5)
for ts in [0.2,6.2,12.4,18.8,25.6,31.6]:
    for j,m in enumerate([84,88,91,96]): pluck(ts+j*.06,m,1.5,.06,pan=(j-1.5)/3,bright=.2)
d=int(.28*sr); out[d:]+=.25*out[:-d][:, ::-1]
fade=np.ones(N); fl=int(2.5*sr); fade[-fl:]=np.linspace(1,0,fl)
out*=fade[:,None]; out/=np.abs(out).max()*1.12
w=wave.open('music.wav','wb');w.setnchannels(2);w.setsampwidth(2);w.setframerate(sr);w.writeframes((out*32767).astype('<i2').tobytes());w.close()
print('ok')
