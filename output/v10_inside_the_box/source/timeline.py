"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=50.0
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("نسمع دائمًا: فكّر خارج الصندوق",0,0,["خارج"],"نسمع دائمًا: فكّر خارج الصندوق"),
 ("لكن، ماذا لو كان الإبداع الحقيقي داخل الصندوق؟",0,0,["الإبداع","داخل"],"ماذا لو كان الإبداع… داخل الصندوق؟"),
 ("الصندوق ليس سجنًا، بل هو حدودك الحقيقية: وقتك، ومواردك، وفريقك",0,0,["سجنًا…","وقتك،","مواردك،","فريقك"],"الصندوق ليس سجنًا… بل حدودك: وقتك، مواردك، فريقك"),
 ("وحين تفهم حدودك جيدًا، تبدأ في اكتشاف ما بداخلها من إمكانات",0,0,["تكتشف","إمكانات"],"حين تفهم حدودك… تكتشف ما بداخلها من إمكانات"),
 ("فالقيود لا تقتل الإبداع، بل توجّهه وتشحذه",0,0,["توجّهه","وتشحذه"],"القيود لا تقتل الإبداع… بل توجّهه وتشحذه"),
 ("جرّب أن تفكر داخل الصندوق بثلاث خطوات",0,0,["داخل","بثلاث"],"فكّر داخل الصندوق بثلاث خطوات"),
 ("أولًا: حدّد صندوقك، ما الموارد التي بين يديك الآن؟",0,0,["صندوقك…","مواردك"],"أولًا: حدّد صندوقك… ما مواردك الآن؟"),
 ("ثانيًا: أعد ترتيب ما بداخله، فقد تكون الفكرة موجودة، لكنها في غير مكانها",0,0,["ترتيب","موجودة"],"ثانيًا: أعد ترتيب ما بداخله… فالفكرة قد تكون موجودة"),
 ("ثالثًا: اسأل: ماذا يمكنني أن أصنع بما أملك، لا بما ينقصني؟",0,0,["أملك…","ينقصني؟"],"ثالثًا: ماذا أصنع بما أملك… لا بما ينقصني؟"),
 ("تذكّر",0,0,["تذكّر"],"تذكّر"),
 ("ليس الإبداع أن تهرب من الصندوق، بل أن تُبدع فيه",0,0,["تهرب","تُبدع"],"ليس الإبداع أن تهرب من الصندوق… بل أن تُبدع فيه"),
 ("فأحيانًا، أعظم الأفكار كانت بالداخل طوال الوقت",0,0,["أعظم","بالداخل"],"أحيانًا… أعظم الأفكار كانت بالداخل طوال الوقت"),
]
SCENES=[("المقولة",0,1),("الصندوق",2,3),("القيود",4,4),("ثلاث خطوات",5,8),("تذكّر",9,10),("الفكرة",11,11)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v10','voice.wav')):
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
