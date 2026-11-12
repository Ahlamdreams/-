"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=69.907324
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("هل سبق أن حضرت برنامجًا تدريبيًا",0,0,["برنامجًا","تدريبيًا"],"هل سبق أن حضرت برنامجًا تدريبيًا"),
 ("خرجت منه بأفكار رائعة وحماس كبير",0,0,["رائعة","وحماس"],"أفكار رائعة وحماس كبير"),
 ("ثم عدت إلى عملك",0,0,["عملك"],"ثم عدت إلى عملك"),
 ("ولم يتغير شيء؟",0,0,["يتغير"],"ولم يتغير شيء؟"),
 ("لماذا يحدث ذلك؟",0,0,["لماذا"],"لماذا يحدث ذلك؟"),
 ("لأن المعرفة وحدها لا تصنع التغيير!",0,0,["المعرفة","التغيير!"],"لأن المعرفة وحدها لا تصنع التغيير!"),
 ("قد يكون البرنامج ثريًا بالمعلومات، ومميزًا في أسلوب تقديمه",0,0,["ثري","ومميز"],"برنامج ثري بالمعلومات… ومميز في تقديمه"),
 ("لكن الأثر الحقيقي لا يُقاس بما تعلّمناه داخل القاعة، بل بما نطبّقه خارجها",0,0,["نطبّقه","خارجها"],"الأثر لا يُقاس بما نتعلّمه داخل القاعة… بل بما نطبّقه خارجها"),
 ("فالمشكلة أحيانًا ليست في جودة التدريب، بل في الفجوة بين التعلّم والتطبيق",0,0,["الفجوة"],"المشكلة ليست في جودة التدريب… بل في الفجوة بين التعلّم والتطبيق"),
 ("كيف نعبر هذه الفجوة؟",0,0,["نعبر","الفجوة؟"],"كيف نعبر هذه الفجوة؟"),
 ("بثلاث خطوات بسيطة",0,0,["بثلاث","خطوات"],"بثلاث خطوات بسيطة"),
 ("أولًا: حوّل ما تعلمته إلى ممارسة محددة، لا مجرد فكرة جميلة",0,0,["ممارسة","محددة"],"أولًا: حوّل الفكرة إلى ممارسة محددة"),
 ("ثانيًا: ابدأ بتطبيق صغير، وحدد متى وأين ستنفذه",0,0,["صغير…","متى","وأين"],"ثانيًا: ابدأ بتطبيق صغير… وحدد متى وأين"),
 ("ثالثًا: تابع الأثر، ما الذي تغيّر فعلًا بعد التطبيق؟",0,0,["الأثر…","تغيّر"],"ثالثًا: تابع الأثر… ما الذي تغيّر فعلًا؟"),
 ("وتذكّر",0,0,["تذكّر"],"تذكّر"),
 ("الشهادة قد تُثبت حضورك، والمعرفة قد تُثري تفكيرك",0,0,["الشهادة","والمعرفة"],"الشهادة تُثبت حضورك… والمعرفة تُثري تفكيرك"),
 ("لكن الممارسة وحدها تكشف أثر تعلّمك",0,0,["الممارسة","أثر"],"لكن الممارسة وحدها تكشف أثر تعلّمك"),
 ("فلا تجعل التدريب محطة تنتهي بانتهاء البرنامج",0,0,["محطة","تنتهي"],"لا تجعل التدريب محطة تنتهي"),
 ("بل نقطة انطلاق لتغيير يبدأ ويستمر",0,0,["انطلاق","ويستمر"],"بل نقطة انطلاق لتغيير يبدأ… ويستمر"),
]
SCENES=[("التجربة",0,3),("السبب",4,7),("الفجوة",8,9),("ثلاث خطوات",10,13),("تذكّر",14,16),("الانطلاق",17,18)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v9','voice.wav')):
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
