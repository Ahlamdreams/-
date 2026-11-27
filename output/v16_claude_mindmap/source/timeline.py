"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=48.338798
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("صفحة كاملة من الكتاب، وطلابك لا يعرفون من أين يبدؤون؟",0,0,["يبدؤون؟"],"صفحة كاملة… وطلابك لا يعرفون من أين يبدؤون؟"),
 ("مع Claude، تتحوّل الصفحة إلى خريطة ذهنية واضحة، في ثوانٍ",0,0,["Claude","خريطة","ذهنية"],"مع Claude تتحوّل الصفحة إلى خريطة ذهنية في ثوانٍ"),
 ("أولًا: صوّر الصفحة بجوالك، وارفعها إلى Claude",0,0,["صوّر"],"١. صوّر الصفحة… وارفعها إلى Claude"),
 ("ثانيًا: اكتب طلبك: حوّل هذه الصفحة إلى خريطة ذهنية ملوّنة، لطلاب الصف الخامس",0,0,["ملوّنة"],"٢. حوّل الصفحة إلى خريطة ذهنية ملوّنة"),
 ("ثالثًا: اضغط إرسال، وشاهد الفكرة الرئيسية في المنتصف، والأفكار الفرعية تتفرّع حولها",0,0,["الرئيسية","والفروع"],"٣. الفكرة الرئيسية في المنتصف… والفروع حولها"),
 ("تريد المزيد؟ اكتب: وأضف ثلاثة أسئلة للمراجعة",0,0,["أسئلة"],"وأضف ٣ أسئلة للمراجعة"),
 ("وتذكّر: راجع الخريطة دائمًا مع الكتاب الأصلي",0,0,["راجع"],"راجع الخريطة دائمًا مع الكتاب الأصلي"),
 ("صفحة مزدحمة بالكلمات، تحوّلت إلى صورة يفهمها الطالب ويتذكّرها",0,0,["يفهمها","ويتذكّرها"],"صفحة مزدحمة… صارت صورة يفهمها الطالب ويتذكّرها"),
 ("احفظ الريلز، وأخبرنا: أي درس ستحوّله أولًا؟",0,0,["احفظ","درس"],"احفظ الريلز… وأخبرنا: أي درس ستحوّله أولًا؟"),
]
SCENES=[("المشكلة",0,1),("الطلب",2,3),("الخريطة",4,5),("المراجعة",6,6),("الخلاصة",7,8)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v16','voice.wav')):
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
