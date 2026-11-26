"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=55.398594
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("في صفك طلاب بمستويات مختلفة، وورقة عمل واحدة، لا تناسبهم جميعًا؟",0,0,["مختلفة…"],"مستويات مختلفة… وورقة عمل واحدة؟"),
 ("بالذكاء الاصطناعي، تصنع ورقة عمل بثلاثة مستويات، في دقيقة واحدة",0,0,["بثلاثة","دقيقة"],"ورقة عمل بثلاثة مستويات… في دقيقة"),
 ("انسخ هذا الطلب",0,0,["الطلب"],"انسخ هذا الطلب"),
 ("أنت معلّم رياضيات للصف الرابع",0,0,["معلّم"],"أنت معلّم رياضيات للصف الرابع"),
 ("اكتب ورقة عمل عن الكسور، بثلاثة مستويات",0,0,["الكسور،"],"اكتب ورقة عمل عن الكسور، بثلاثة مستويات"),
 ("مستوى داعم، بخطوات وأمثلة محلولة",0,0,["داعم،"],"مستوى داعم، بخطوات وأمثلة محلولة"),
 ("ومستوى أساسي، بتمارين مباشرة",0,0,["أساسي،"],"ومستوى أساسي، بتمارين مباشرة"),
 ("ومستوى متقدّم، بمسائل حياتية تتحدّى التفكير",0,0,["متقدّم،"],"ومستوى متقدّم، بمسائل تتحدّى التفكير"),
 ("اضغط إرسال، وستحصل على ثلاث أوراق جاهزة",0,0,["جاهزة"],"اضغط إرسال… ثلاث أوراق جاهزة"),
 ("وزّعها بألوان مختلفة: الأخضر، والأصفر، والأزرق، دون أن يشعر أي طالب بالفرق",0,0,["بألوان"],"وزّعها بألوان مختلفة… دون أن يشعر أحد بالفرق"),
 ("راجعها، وعدّل ما يناسب طلابك، فأنت الأعرف بهم",0,0,["راجعها"],"راجعها وعدّلها… فأنت الأعرف بطلابك"),
 ("هكذا يتعلّم كل طالب، من حيث يقف، إلى حيث يمكن أن يصل",0,0,["يقف","يصل"],"كل طالب يتعلّم… من حيث يقف إلى حيث يصل"),
 ("احفظ الريلز، وأخبرنا: في أي مادة ستجرّبها؟",0,0,["احفظ","مادة"],"احفظ الريلز… وأخبرنا: في أي مادة ستجرّبها؟"),
]
SCENES=[("المشكلة",0,1),("الطلب",2,7),("النتيجة",8,8),("التوزيع",9,10),("الخلاصة",11,12)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v14','voice.wav')):
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
