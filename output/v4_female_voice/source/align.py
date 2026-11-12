import numpy as np, wave, sys, json, re
sys.path.insert(0,sys.argv[2]); import timeline as TL
w=wave.open(sys.argv[1]); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
e=np.array([np.sqrt(np.mean(a[i:i+320]**2)) for i in range(0,len(a)-320,160)])
th=max(np.percentile(e,30)*1.5,0.008); v=e>th
# voiced segments separated by gaps >= 0.12s
segs=[];st=None;last=None
for i,x in enumerate(v):
    if x:
        if st is None: st=i
        elif i-last>12: segs.append((st,last)); st=i
        last=i
segs.append((st,last)); segs=[(s/100,(e_+1)/100) for s,e_ in segs]
gaps=[segs[i+1][0]-segs[i][1] for i in range(len(segs)-1)]
print(len(segs),'segments'); 
L=[len(re.sub(r'[ً-ْ\s]','',p[0])) for p in TL.PHRASES]; m=len(L); n=len(segs)
tot=sum(s[1]-s[0] for s in segs); span=segs[-1][1]-segs[0][0]
# DP: phrase k covers segs[i..j]; cost = squared error of phrase duration vs expected (by letters), minus bonus for big gap at boundary
exp=[span*l/sum(L) for l in L]
INF=1e9; dp=np.full((m+1,n+1),INF); bk=np.zeros((m+1,n+1),int); dp[0,0]=0
for k in range(1,m+1):
    for j in range(1,n+1):
        for i in range(k-1,j):
            if dp[k-1,i]>=INF: continue
            d=segs[j-1][1]-segs[i][0]; c=(d-exp[k-1])**2/max(exp[k-1],0.3)
            if j<n: c-=1.2*min(gaps[j-1],0.8)
            if dp[k-1,i]+c<dp[k,j]: dp[k,j]=dp[k-1,i]+c; bk[k,j]=i
out=[];j=n
for k in range(m,0,-1):
    i=bk[k,j]; out.append((round(segs[i][0],2),round(segs[j-1][1],2))); j=i
out=out[::-1]
for p,(s,e_) in zip(TL.PHRASES,out): print(f"{s:6.2f} {e_:6.2f}  {p[0]}")
json.dump(out,open('phrase_times.json','w'))
