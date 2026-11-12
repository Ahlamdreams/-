"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=68.432971
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("هل نستخدم التقنية لتغيير التعلّم",0,0,["التعلّم"],"هل نستخدم التقنية لتغيير التعلّم"),
 ("أم لتغيير شكل العرض فقط؟",0,0,["العرض"],"أم لتغيير شكل العرض فقط؟"),
 ("قد نستبدل السبورة بشاشة",0,0,["الشاشة"],"من السبورة إلى الشاشة"),
 ("والكتاب بعرض تفاعلي",0,0,["تفاعلي"],"من الكتاب إلى عرض تفاعلي"),
 ("والورقة بتطبيق رقمي",0,0,["رقمي"],"من الورقة إلى تطبيق رقمي"),
 ("لكن، هل تغيّر التعلّم فعلًا؟",0,0,["التعلّم"],"لكن… هل تغيّر التعلّم فعلًا؟"),
 ("فليست كل حصة تستخدم التقنية حصةً تُحدث فرقًا!",0,0,["فرقًا!"],"ليست كل حصة تقنية… تُحدث فرقًا!"),
 ("التقنية لا تصبح مؤثرة لمجرد حضورها في الصف",0,0,["حضورها"],"التقنية لا تؤثر بمجرد حضورها"),
 ("بل عندما تغيّر ما يفعله المتعلّم",0,0,["يفعله"],"بل عندما تغيّر ما يفعله المتعلّم"),
 ("حين ينتقل من المشاهدة إلى المشاركة",0,0,["المشاركة"],"من المشاهدة إلى المشاركة"),
 ("ومن تلقّي المعلومة إلى تحليلها واكتشافها",0,0,["التحليل","والاكتشاف"],"من التلقّي إلى التحليل والاكتشاف"),
 ("ومن انتظار التصحيح إلى الحصول على تغذية راجعة تساعده على التقدّم",0,0,["تغذية","راجعة"],"من انتظار التصحيح إلى تغذية راجعة تدفع للتقدّم"),
 ("وقبل أن تختار أداتك الرقمية، اسأل نفسك ثلاثة أسئلة",0,0,["ثلاثة","أسئلة"],"قبل اختيار أداتك الرقمية… ثلاثة أسئلة"),
 ("ماذا سيفعل المتعلّم بهذه الأداة؟",0,0,["سيفعل"],"ماذا سيفعل المتعلّم بهذه الأداة؟"),
 ("كيف سأعرف أنه تعلّم؟",0,0,["تعلّم؟"],"كيف سأعرف أنه تعلّم؟"),
 ("وما الذي ستضيفه التقنية، ولم يكن ممكنًا بالفاعلية نفسها دونها؟",0,0,["ستضيفه"],"ما الذي ستضيفه التقنية… ولم يكن ممكنًا دونها؟"),
 ("تذكّر",0,0,["تذكّر"],"تذكّر"),
 ("التحوّل الرقمي في التعليم لا يبدأ بتغيير الشاشة",0,0,["الشاشة"],"التحوّل الرقمي لا يبدأ بتغيير الشاشة"),
 ("بل بتغيير تجربة التعلّم",0,0,["تجربة","التعلّم"],"بل بتغيير تجربة التعلّم"),
 ("فالتقنية ليست غاية نصل إليها",0,0,["غاية"],"التقنية ليست غاية نصل إليها"),
 ("بل وسيلة نصنع بها تعلّمًا أعمق، وأثرًا أبقى",0,0,["وسيلة","أعمق","أبقى"],"بل وسيلة لتعلّم أعمق وأثر أبقى"),
]
SCENES=[("السؤال",0,1),("الاستبدال",2,5),("الفرق",6,8),("التحوّل",9,11),("ثلاثة أسئلة",12,15),("تذكّر",16,18),("الأثر",19,20)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v8','voice.wav')):
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
