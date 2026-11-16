"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=55.109388
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("هل تحتاج أسئلة متنوعة لدرسك، في دقيقة واحدة؟",0,0,["دقيقة"],"أسئلة متنوعة لدرسك… في دقيقة واحدة؟"),
 ("الذكاء الاصطناعي يساعدك، بشرط أن تعرف كيف تطلب",0,0,["تطلب"],"الذكاء الاصطناعي يساعدك… إذا عرفت كيف تطلب"),
 ("إليك المعادلة في أربع خطوات",0,0,["المعادلة"],"إليك المعادلة في أربع خطوات"),
 ("أولًا: حدّد الدور، أنت معلّم علوم للصف الخامس",0,0,["الدور"],"١. حدّد الدور"),
 ("ثانيًا: حدّد المهمة، اكتب عشرة أسئلة عن دورة الماء",0,0,["المهمة"],"٢. حدّد المهمة"),
 ("ثالثًا: حدّد المستويات، من التذكّر إلى التحليل والتطبيق",0,0,["المستويات"],"٣. حدّد المستويات"),
 ("رابعًا: حدّد الشكل، اختيار من متعدد، مع الإجابة الصحيحة",0,0,["الشكل"],"٤. حدّد الشكل"),
 ("اضغط إرسال، وخلال ثوانٍ تحصل على أسئلة جاهزة ومتدرّجة",0,0,["جاهزة","ومتدرّجة"],"اضغط إرسال… أسئلة جاهزة ومتدرّجة خلال ثوانٍ"),
 ("لكن تذكّر، راجعها دائمًا قبل استخدامها، فأنت الخبير بطلابك",0,0,["راجعها","الخبير"],"راجعها دائمًا… فأنت الخبير بطلابك"),
 ("الذكاء الاصطناعي لا يستبدل المعلّم، بل يمنحه وقتًا أكثر لما يهمّ",0,0,["يستبدل","وقتًا"],"لا يستبدل المعلّم… بل يمنحه وقتًا لما يهمّ"),
 ("احفظ الريلز، وجرّب المعادلة في درسك القادم، وأخبرنا بالنتيجة في التعليقات",0,0,["احفظ","وأخبرنا"],"احفظ الريلز… جرّب المعادلة وأخبرنا بالنتيجة"),
]
SCENES=[("السؤال",0,1),("المعادلة",2,2),("التطبيق",3,7),("المراجعة",8,8),("الخلاصة",9,10)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v12','voice.wav')):
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
