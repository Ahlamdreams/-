"""Transcript, phrase/word timings and scene plan (initiatives & impact)."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=49.71
FPS=30
PHRASES=[
 ("هل كل مبادرة ننفذها",0,0,["مبادرة"]),
 ("تصنع أثرًا؟",0,0,["أثرًا؟"]),
 ("ليس بالضرورة!",0,0,["بالضرورة!"]),
 ("فقد نمتلك فكرة جميلة",0,0,["فكرة"]),
 ("ونصمم شعارًا جذابًا",0,0,["شعارًا"]),
 ("وننفذ أنشطة متعددة",0,0,["أنشطة"]),
 ("لكن يبقى السؤال الأهم",0,0,["السؤال","الأهم"]),
 ("ما الذي تغيّر فعلًا؟",0,0,["تغيّر"]),
 ("المبادرة الناجحة لا تبدأ بالسؤال",0,0,["الناجحة"]),
 ("ماذا سننفذ؟",0,0,["سننفذ؟"]),
 ("بل تبدأ بثلاثة أسئلة",0,0,["بثلاثة","أسئلة"]),
 ("ما المشكلة التي نريد حلّها؟",0,0,["المشكلة"]),
 ("ما التغيير الذي نسعى إلى تحقيقه؟",0,0,["التغيير"]),
 ("وكيف سنقيس هذا التغيير؟",0,0,["سنقيس"]),
 ("تذكّر",0,0,["تذكّر"]),
 ("الإنجاز ليس في عدد المبادرات",0,0,["الإنجاز","عدد"]),
 ("بل في حجم الأثر الذي تتركه",0,0,["حجم","الأثر"]),
 ("فلا تجعل مبادرتك مجرد نشاط ينتهي",0,0,["ينتهي"]),
 ("بل أثرًا يبدأ ويستمر!",0,0,["أثرًا","ويستمر!"]),
]
SCENES=[("الأثر",0,2),("الأنشطة",3,7),("البداية",8,10),("ثلاثة أسئلة",11,13),("الخلاصة",14,18)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!]','',w))
def compute(wav=os.path.join(HERE,'..','v5','voice.wav')):
    w=wave.open(wav); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
    e=np.array([np.sqrt(np.mean(a[i:i+320]**2)) for i in range(0,len(a)-320,160)])
    th=max(np.percentile(e,30)*1.5,0.008); voiced=(e>th).astype(float)
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
if __name__=='__main__':
    P=compute(); json.dump(dict(offset=OFF,phrases=P,scenes=SCENES),open(os.path.join(HERE,'timings.json'),'w'),ensure_ascii=False,indent=1)
    for i,p in enumerate(P): print(i,p['start'],p['end'],' | '.join(f"{w['word']}@{w['start']}" for w in p['words']))
