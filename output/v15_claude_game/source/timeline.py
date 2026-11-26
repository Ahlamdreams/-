"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=56.280159
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("هل تعلم أنك تستطيع صناعة لعبة تعليمية تفاعلية لطلابك، دون أن تكتب سطرًا واحدًا من البرمجة؟",0,0,["لعبة","برمجة"],"لعبة تعليمية تفاعلية… بدون سطر برمجة واحد؟"),
 ("مع Claude، الأمر يحتاج دقيقة واحدة",0,0,["Claude","دقيقة"],"مع Claude يكفي دقيقة واحدة"),
 ("افتح Claude، واكتب طلبك بوضوح",0,0,["بوضوح"],"افتح Claude واكتب طلبك بوضوح"),
 ("اصنع لي لعبة تفاعلية لطلاب الصف الخامس عن الكسور",0,0,["لعبة"],"لعبة تفاعلية للصف الخامس عن الكسور"),
 ("فيها عشرة أسئلة، وعدّاد للنقاط، ورسالة تشجيع عند كل إجابة صحيحة",0,0,["نقاط…","تشجيع"],"١٠ أسئلة… عدّاد نقاط… رسالة تشجيع"),
 ("اضغط إرسال، وشاهد Claude يبني اللعبة أمامك، جاهزة للعب مباشرة",0,0,["جاهزة"],"اضغط إرسال… واللعبة جاهزة أمامك"),
 ("تريد تعديلًا؟ اكتب: اجعل الألوان أكثر حيوية، وأضف مؤقّتًا لكل سؤال",0,0,["تعديل؟","مؤقّت"],"تعديل؟ ألوان أكثر حيوية… ومؤقّت لكل سؤال"),
 ("وعندما تنتهي، انشر الرابط وشاركه مع طلابك، أو اعرضها على شاشة الصف",0,0,["الرابط"],"انشر الرابط لطلابك… أو اعرضها على شاشة الصف"),
 ("ولا تنسَ: جرّبها بنفسك أولًا، وتأكد من صحة الأسئلة",0,0,["جرّبها","صحة"],"جرّبها أولًا… وتأكد من صحة الأسئلة"),
 ("Claude لا يصنع الدرس بدلًا منك، بل يحوّل فكرتك إلى تجربة يعيشها طلابك",0,0,["فكرتك…","تجربة"],"Claude يحوّل فكرتك… إلى تجربة يعيشها طلابك"),
 ("احفظ الريلز، وأخبرنا: ما اللعبة التي ستصنعها أولًا؟",0,0,["احفظ","اللعبة"],"احفظ الريلز… وأخبرنا: ما اللعبة التي ستصنعها أولًا؟"),
]
SCENES=[("الفكرة",0,1),("الطلب",2,4),("اللعبة",5,6),("المشاركة",7,8),("الخلاصة",9,10)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v15','voice.wav')):
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
