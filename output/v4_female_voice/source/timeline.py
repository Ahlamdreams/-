"""Transcript, phrase/word timings (voice-activity aligned) and scene plan."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0            # intro before the voice starts (s)
VOICE_LEN=31.68    # full recording length (s) - nothing is cut
FPS=30
# phrase text, start, end (seconds in the original recording), highlighted words
PHRASES=[
 ("لا تنتظر",0.96,2.43,["تنتظر"]),
 ("أن يمنحك الطريق فرصة",2.58,3.42,["الطريق","فرصة"]),
 ("اصنع فرصتك بنفسك",3.93,5.61,["فرصتك","بنفسك"]),
 ("لا تخشَ البدايات المختلفة",6.27,8.13,["البدايات","المختلفة"]),
 ("ولا تتراجع",8.52,9.57,["تتراجع"]),
 ("لأن الطريق لم تتضح ملامحه بعد",9.87,12.69,["الطريق","تتضح"]),
 ("فبعض الأبواب",13.32,14.28,["الأبواب"]),
 ("لا تُفتح",14.55,14.97,["تُفتح"]),
 ("إلا لمن يجرؤ على المحاولة",15.15,17.52,["يجرؤ","المحاولة"]),
 ("وبعض الأحلام",17.97,19.02,["الأحلام"]),
 ("لا تتحقق",19.29,20.22,["تتحقق"]),
 ("إلا لمن يؤمن بها",20.37,21.90,["يؤمن"]),
 ("ويعمل لأجلها",22.20,23.58,["ويعمل"]),
 ("كن صاحب المبادرة",24.21,25.62,["المبادرة"]),
 ("لا منتظر الفرصة",25.95,27.21,["الفرصة"]),
 ("فالحياة تتسع",27.81,28.44,["فالحياة","تتسع"]),
 ("لمن يملك شجاعة السعي",28.77,31.30,["شجاعة","السعي"]),
]
SCENES=[("الفرصة",0,2),("البدايات",3,5),("الأبواب",6,8),("الأحلام",9,12),("المبادرة",13,16)]
def _letters(w): return len(re.sub(r'[ً-ْ]','',w))
def compute(wav=os.path.join(HERE,'..','v4','takeA.wav')):
    w=wave.open(wav); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
    sr=16000; hop=160; e=np.array([np.sqrt(np.mean(a[i:i+320]**2)) for i in range(0,len(a)-320,hop)])
    th=np.percentile(e,30)*1.5; voiced=(e>th).astype(float)
    out=[]
    for (txt,t0,t1,hl) in PHRASES:
        i0,i1=int(t0*100),int(t1*100); v=voiced[i0:i1]; cum=np.concatenate([[0],np.cumsum(v)]); cum/=max(cum[-1],1)
        words=txt.split(); wt=np.array([_letters(x)+1.5 for x in words],float); edges=np.concatenate([[0],np.cumsum(wt)])/wt.sum()
        ws=[]
        for j,wd in enumerate(words):
            s=t0+np.searchsorted(cum,edges[j]+1e-6)/100; en=t0+np.searchsorted(cum,edges[j+1]-1e-6)/100
            ws.append(dict(word=wd,start=round(float(s)+OFF,3),end=round(float(en)+OFF,3),key=wd in hl))
        out.append(dict(text=txt,start=round(t0+OFF,3),end=round(t1+OFF,3),words=ws))
    return out
import json as _j
_T=_j.load(open(os.path.join(HERE,'phrase_times.json')))
PHRASES=[(p[0],_T[i][0],_T[i][1],p[3]) for i,p in enumerate(PHRASES)]
if __name__=='__main__':
    P=compute(); json.dump(dict(offset=OFF,phrases=P,scenes=SCENES),open(os.path.join(HERE,'timings.json'),'w'),ensure_ascii=False,indent=1)
    for i,p in enumerate(P): print(i,p['start'],p['end'],' | '.join(f"{w['word']}@{w['start']}" for w in p['words']))
