"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=64.99
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("هل القائد الناجح هو الأكثر حضورًا",0,0,["حضورًا"],"هل القائد الناجح هو الأكثر حضورًا"),
 ("أم الأكثر تمكينًا؟",0,0,["تمكينًا؟"],"أم الأكثر تمكينًا؟"),
 ("قد يظن البعض أن القائد المتميز هو من يكون حاضرًا في كل تفصيلة",0,0,["تفصيلة"],"حاضر في كل تفصيلة"),
 ("متابعًا لكل صغيرة وكبيرة",0,0,["متابعة"],"متابعة كل صغيرة وكبيرة"),
 ("ممسكًا بكل الخيوط بيده",0,0,["الخيوط"],"كل الخيوط بيده"),
 ("لكن الحقيقة الأعمق تقول",0,0,["الأعمق"],"الحقيقة الأعمق"),
 ("إن القائد الحقيقي لا يُقاس فقط بما ينجزه بنفسه",0,0,["بنفسه"],"لا يُقاس بما ينجزه بنفسه"),
 ("بل بما يصنعه في الآخرين",0,0,["الآخرين"],"بل بما يصنعه في الآخرين"),
 ("القائد الأكثر حضورًا قد ينجز كثيرًا",0,0,["حضورًا"],"الأكثر حضورًا قد ينجز كثيرًا"),
 ("لكن القائد الأكثر تمكينًا يبني من حوله أشخاصًا قادرين على الإنجاز",0,0,["تمكينًا","الإنجاز"],"الأكثر تمكينًا يبني أشخاصًا قادرين على الإنجاز"),
 ("وقادرين على اتخاذ القرار",0,0,["القرار"],"اتخاذ القرار"),
 ("وقادرين على تحمّل المسؤولية بثقة",0,0,["المسؤولية","بثقة"],"تحمّل المسؤولية بثقة"),
 ("فالقيادة ليست أن تبقى أنت في الواجهة دائمًا",0,0,["الواجهة"],"القيادة ليست البقاء في الواجهة"),
 ("بل أن تصنع فريقًا يعرف الطريق",0,0,["فريقًا","الطريق"],"بل تصنع فريقًا يعرف الطريق"),
 ("ويملك الأدوات",0,0,["الأدوات"],"ويملك الأدوات"),
 ("ويواصل الأثر",0,0,["الأثر"],"ويواصل الأثر"),
 ("اسأل نفسك",0,0,["نفسك"],"اسأل نفسك"),
 ("هل نجاحي مرتبط بوجودي فقط؟",0,0,["بوجودي"],"هل نجاحي مرتبط بوجودي فقط؟"),
 ("أم أنني بنيت من حولي من يستطيع أن ينجح أيضًا؟",0,0,["ينجح"],"أم بنيتُ من حولي من ينجح أيضًا؟"),
 ("فالقائد الناجح ليس فقط من يُدير العمل",0,0,["يُدير"],"ليس فقط من يُدير العمل"),
 ("بل من يُمكّن الآخرين ليقودوا الأثر من بعده",0,0,["يُمكّن","الأثر"],"بل من يُمكّن الآخرين ليقودوا الأثر"),
]
SCENES=[("السؤال",0,1),("الحضور",2,4),("التحوّل",5,7),("التمكين",8,11),("الفريق",12,15),("تأمّل",16,18),("الأثر",19,20)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v7','voice.wav')):
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
