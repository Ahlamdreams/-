import wave,numpy as np
sr=44100;D=42.0;N=int(sr*D)
def rd(p):
    w=wave.open(p);x=np.frombuffer(w.readframes(w.getnframes()),'<i2').reshape(-1,2)/32768.;w.close();return x
m=rd('music.wav')[:N]; m=np.vstack([m,np.zeros((N-len(m),2))]) if len(m)<N else m
v=rd('voice.wav');v=v/np.abs(v).max()*.92
t=np.arange(N)/sr
g=np.interp(t,[0,0.8,1.3,36.8,37.6,40.3,42],[1,1,.28,.28,.95,.95,0])
m=m/np.abs(m).max()*.9*g[:,None]
mix=m.copy();i=int(1.2*sr);e=min(N,i+len(v));mix[i:e]+=v[:e-i]
mix/=max(1,np.abs(mix).max()/.97)
w=wave.open('mix.wav','wb');w.setnchannels(2);w.setsampwidth(2);w.setframerate(sr);w.writeframes((mix*32767).astype('<i2').tobytes());w.close()
