# Procedural cinematic/inspirational score (A minor -> C major), 72 bpm
import numpy as np, wave, sys
SR=44100; DUR=float(sys.argv[2]) if len(sys.argv)>2 else 37.0
N=int(SR*DUR); t=np.arange(N)/SR
L=np.zeros(N); R=np.zeros(N)
rng=np.random.default_rng(7)
BPM=72; beat=60/BPM; bar=4*beat
def hz(m): return 440*2**((m-69)/12)
# chords as midi notes
A,F,C,G=[57,60,64],[53,57,60],[48,52,55],[55,59,62]
prog=[A,F,C,G, A,F,C,G, F,G,C,C,C]
def add(sig,start,pan=0.0,gain=1.0):
    i=int(start*SR); n=min(len(sig),N-i)
    if n<=0 or i<0: return
    L[i:i+n]+=sig[:n]*gain*np.sqrt(0.5*(1-pan)); R[i:i+n]+=sig[:n]*gain*np.sqrt(0.5*(1+pan))
def pad_note(m,dur,amp):
    n=int(dur*SR); tt=np.arange(n)/SR; s=np.zeros(n)
    for det in (-0.0025,0,0.0028):
        f=hz(m)*(1+det); ph=rng.uniform(0,6.28)
        for h in range(1,7): s+=np.sin(2*np.pi*f*h*tt+ph*h)/h**1.6
    env=np.minimum(1,tt/1.2)*np.minimum(1,(dur-tt)/1.4).clip(0,1)
    return s*env*amp/3
def piano(m,amp,dur=2.2):
    n=int(dur*SR); tt=np.arange(n)/SR; f=hz(m); s=np.zeros(n)
    for h,a,d in [(1,1,1.6),(2,.45,.9),(3,.2,.5),(4,.12,.35),(5,.05,.25)]:
        s+=a*np.sin(2*np.pi*f*h*tt*(1+0.0004*h*h))*np.exp(-tt/d)
    s*=np.minimum(1,tt/0.004)
    return s*amp
# pads
for i,ch in enumerate(prog):
    st=i*bar-0.6
    if st>DUR: break
    d=bar+1.4 if i<len(prog)-3 else 3*bar+1.4
    if i>=len(prog)-2: continue
    for j,m in enumerate(ch+[ch[0]-12]):
        add(pad_note(m,d,0.06 if i>0 else 0.045),max(0,st),pan=(-0.5+j*0.33))
# piano arpeggios (eighths)
pat=[0,2,1,2,0+12,2,1,2]
for i,ch in enumerate(prog[:11]):
    notes=[ch[0]+12,ch[1]+12,ch[2]+12]
    for k in range(8):
        tt0=i*bar+k*beat/2
        if tt0>DUR-3: break
        idx=pat[k]; m=notes[idx%12 if idx<3 else 0]+(12 if idx>=12 else 0)
        amp=0.07*(1.0 if k%2==0 else 0.7)*(0.7 if i<2 else 1.0)
        add(piano(m,amp),tt0,pan=0.3*np.sin(k))
    if i>=6:  # high counter melody in climax
        mel=[ch[2]+24,ch[1]+24,ch[0]+24]
        for k,m in enumerate(mel): add(piano(m,0.035,3),i*bar+k*beat*1.33,pan=-0.3)
# final sustained C major chord
end=11*bar
for m in [48,55,60,64,67,72,76]: add(pad_note(m,DUR-end+0.5,0.05),end-0.3,pan=rng.uniform(-.5,.5))
for m in [60,64,67,72]: add(piano(m,0.08,5),end+0.05*([60,64,67,72].index(m)))
# sub bass from bar 4
for i,ch in enumerate(prog[:11]):
    if i<4: continue
    n=int(bar*SR); tt=np.arange(n)/SR
    s=np.sin(2*np.pi*hz(ch[0]-24)*tt)*np.minimum(1,tt/0.05)*np.minimum(1,(bar-tt)/0.1)
    add(s*0.12,i*bar)
# soft kick on beats 1,3 from bar 4
def kick():
    n=int(0.5*SR); tt=np.arange(n)/SR
    f=45+80*np.exp(-tt/0.04); ph=2*np.pi*np.cumsum(f)/SR
    return np.sin(ph)*np.exp(-tt/0.18)
k=kick()
for i in range(4,11):
    for b in (0,2):
        add(k,i*bar+b*beat,gain=0.22 if i<7 else 0.3)
# scene whooshes & riser/impact
def noise_swell(dur,peak):
    n=int(dur*SR); x=rng.standard_normal(n)
    # crude bandpass by differencing + smoothing
    y=np.convolve(x,np.ones(30)/30,'same'); y=x-np.convolve(x,np.ones(6)/6,'same')+y*0.5
    tt=np.arange(n)/SR; env=np.exp(-((tt-dur*peak)/(dur*0.22))**2)
    return y*env
scenes=[float(x) for x in sys.argv[3].split(',')] if len(sys.argv)>3 else []
for s in scenes:
    add(noise_swell(1.2,0.7)*0.05,s-0.85,pan=-0.4); add(noise_swell(1.2,0.7)*0.05,s-0.85,pan=0.4)
    add(piano(84,0.03,3),s); add(piano(91,0.02,3),s+0.08)
if len(scenes)>=4:
    s4=scenes[3]
    n=int(2.5*SR); tt=np.arange(n)/SR; x=rng.standard_normal(n)
    x=x-np.convolve(x,np.ones(4)/4,'same')
    riser=x*(tt/2.5)**3*0.12 + np.sin(2*np.pi*np.cumsum(200+600*(tt/2.5)**2)/SR)*(tt/2.5)**3*0.03
    add(riser,s4-2.5,pan=-0.2); add(riser[::-1][::-1],s4-2.5,pan=0.2)
    n=int(3*SR); tt=np.arange(n)/SR
    boom=np.sin(2*np.pi*np.cumsum(30+60*np.exp(-tt/0.08))/SR)*np.exp(-tt/0.9)
    add(boom*0.35,s4)
# reverb (FFT convolution with decaying noise IR)
def reverb(x,seed):
    r=np.random.default_rng(seed); n=int(2.8*SR); tt=np.arange(n)/SR
    ir=r.standard_normal(n)*np.exp(-tt/0.75); ir[:int(0.02*SR)]*=np.linspace(0,1,int(0.02*SR))
    ir/=np.sqrt(np.sum(ir**2))
    M=len(x)+n; F=1<<int(np.ceil(np.log2(M)))
    return np.fft.irfft(np.fft.rfft(x,F)*np.fft.rfft(ir,F),F)[:len(x)]
L2=0.65*L+0.5*reverb(L,1); R2=0.65*R+0.5*reverb(R,2)
# fade in/out
fade=np.minimum(1,t/1.0)*np.clip((DUR-t)/2.5,0,1)
st=np.stack([L2*fade,R2*fade],1); st/=np.max(np.abs(st))*1.12
w=wave.open(sys.argv[1],'wb'); w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR)
w.writeframes((st*32767).astype(np.int16).tobytes()); w.close()
print("ok",DUR)
