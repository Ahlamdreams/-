"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=55.654853
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("عندك ثلاثون طالبًا، ولا يوجد جهاز واحد في الصف؟",0,0,["جهاز"],"٣٠ طالبًا… ولا جهاز واحد في الصف؟"),
 ("وتريد أن تعرف من فهم الدرس، خلال ثلاثين ثانية؟",0,0,["فهم…","ثانية؟"],"تريد أن تعرف من فهم… خلال ٣٠ ثانية؟"),
 ("جرّب Plickers",0,0,["Plickers"],"جرّب Plickers"),
 ("كل طالب يحصل على بطاقة ورقية مطبوعة، لها شكل خاص، وأربعة أضلاع: أ، ب، ج، د",0,0,["بطاقة","ورقية"],"بطاقة ورقية لكل طالب… لها ٤ أضلاع: أ ب ج د"),
 ("أولًا: أضف سؤالك في الموقع، واعرضه على الشاشة",0,0,["سؤالك…"],"١. أضف سؤالك… واعرضه على الشاشة"),
 ("ثانيًا: يرفع كل طالب بطاقته، ويدير الحرف الذي يمثل إجابته إلى الأعلى",0,0,["بطاقته…","للأعلى"],"٢. يرفع الطالب بطاقته… وحرف إجابته للأعلى"),
 ("ثالثًا: امسح الصف بكاميرا جوالك، فتُقرأ كل الإجابات في ثوانٍ",0,0,["امسح","ثوانٍ"],"٣. امسح الصف بجوالك… فتُقرأ الإجابات في ثوانٍ"),
 ("والنتيجة؟ تظهر لك نسبة الإجابات الصحيحة، واسم كل طالب وإجابته",0,0,["النتيجة:"],"النتيجة: نسبة الصواب… واسم كل طالب وإجابته"),
 ("فتعرف فورًا: هل تنتقل إلى الدرس التالي، أم تعيد الشرح؟",0,0,["تنتقل","تعيد"],"تنتقل للدرس التالي… أم تعيد الشرح؟"),
 ("تقويم فوري، بلا أجهزة للطلاب، وبلا أوراق تصحيح",0,0,["فوري…"],"تقويم فوري… بلا أجهزة… بلا تصحيح"),
 ("احفظ الريلز، وجرّب Plickers في حصتك القادمة، وأخبرنا كيف تفاعل طلابك",0,0,["احفظ","تفاعل"],"احفظ الريلز… وأخبرنا كيف تفاعل طلابك"),
]
SCENES=[("المشكلة",0,1),("الأداة",2,3),("الخطوات",4,6),("النتيجة",7,8),("الخلاصة",9,10)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v13','voice.wav')):
    w=wave.open(wav); a=np.frombuffer(w.readframes(w.getnframes()),np.int16).astype(np.float32)/32768
    e=np.array([np.sqrt(np.mean(a[i:i+320]**2)) for i in range(0,len(a)-320,160)])
    th=max(np.percentile(e,30)*1.5,0.008); voiced=(e>th).astype(float)
    out=[]
    for (txt,t0,t1,hl,disp) in PHRASES:
        i0,i1=int(t0*100),int(t1*100); v=voiced[i0:i1]; cum=np.concatenate([[0],np.cumsum(v)]); cum/=max(cum[-1],1)
        words=disp.split(); wt=np.array([_letters(x)+1.5 for x in words],float); edges=np.concatenate([[0],np.cumsum(wt)])/wt.sum()
        span=0.8 if disp!=txt else 1.0
        ws=[]
        for j,wd in enumerate(words):
            s=t0+np.searchsorted(cum,edges[j]*span+1e-6)/100; en=t0+np.searchsorted(cum,edges[j+1]*span-1e-6)/100
            ws.append(dict(word=wd,start=round(float(s)+OFF,3),end=round(float(en)+OFF,3),key=wd in hl))
        out.append(dict(text=txt,caption=disp,start=round(t0+OFF,3),end=round(t1+OFF,3),words=ws))
    return out
if __name__=='__main__':
    P=compute(); json.dump(dict(offset=OFF,phrases=P,scenes=SCENES),open(os.path.join(HERE,'timings.json'),'w'),ensure_ascii=False,indent=1)
    for i,p in enumerate(P): print(i,p['start'],p['end'],' | '.join(f"{w['word']}@{w['start']}" for w in p['words']))
