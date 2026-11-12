"""Transcript, phrase/word timings and scene plan (initiatives & impact)."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=61.231020
FPS=30
PHRASES=[
 ("هل فكرت يومًا",0,0,["فكرت"]),
 ("لماذا ننشغل كثيرًا",0,0,["ننشغل"]),
 ("ومع ذلك لا نحقق دائمًا ما نطمح إليه؟",0,0,["نطمح"]),
 ("لأن الانشغال لا يعني بالضرورة الإنجاز!",0,0,["الانشغال","الإنجاز!"]),
 ("قد تمتلئ جداولنا بالمهام",0,0,["بالمهام"]),
 ("وتزدحم أيامنا بالاجتماعات",0,0,["بالاجتماعات"]),
 ("وتتوالى الأعمال",0,0,["الأعمال"]),
 ("لكن يبقى السؤال",0,0,["السؤال"]),
 ("ما الذي أحدث فرقًا حقيقيًا؟",0,0,["فرقًا","حقيقيًا؟"]),
 ("هنا تبدأ الإدارة الواعية للوقت والأولويات",0,0,["الواعية","والأولويات"]),
 ("جرّب أن تسأل نفسك ثلاثة أسئلة",0,0,["ثلاثة"]),
 ("قبل أن تبدأ يومك",0,0,["يومك"]),
 ("أولًا: ما المهمة التي ستصنع الأثر الأكبر؟",0,0,["الأثر","الأكبر؟"]),
 ("ثانيًا: ما العمل الذي يمكن تفويضه أو الاستغناء عنه؟",0,0,["تفويضه","الاستغناء"]),
 ("ثالثًا: كيف سأعرف أنني حققت نتيجة فعلية؟",0,0,["نتيجة"]),
 ("تذكّر",0,0,["تذكّر"]),
 ("ليست قيمة يومك في عدد المهام التي أنجزتها",0,0,["عدد"]),
 ("بل في قيمة النتائج التي صنعتها",0,0,["النتائج"]),
 ("فلا تنشغل فقط",0,0,["تنشغل"]),
 ("بل أنجز ما يستحق!",0,0,["أنجز","يستحق!"]),
]
SCENES=[("الانشغال",0,3),("الزحام",4,8),("الأولويات",9,11),("ثلاثة أسئلة",12,14),("الخلاصة",15,19)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!]','',w))
def compute(wav=os.path.join(HERE,'..','v6','voice.wav')):
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
