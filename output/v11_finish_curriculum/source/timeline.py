"""Leadership reel: spoken phrases, on-screen keyword captions, timings, scenes."""
import numpy as np, wave, json, os, re
HERE=os.path.dirname(os.path.abspath(__file__))
OFF=1.0
VOICE_LEN=58.467823
FPS=30
# (spoken text, start, end, gold keywords, on-screen caption)
PHRASES=[
 ("متأخر في المنهج",0,0,["المنهج؟"],"متأخر في المنهج؟"),
 ("والوقت يداهمك؟",0,0,["يداهمك؟"],"والوقت يداهمك؟"),
 ("إليك خمس أفكار عملية، تساعدك على إنهاء منهجك في الوقت المحدد",0,0,["خمس","أفكار"],"خمس أفكار عملية لإنهاء منهجك في الوقت"),
 ("أولًا: ارسم خريطة المنهج",0,0,["خريطة"],"١. ارسم خريطة المنهج"),
 ("اكتب الدروس المتبقية، وعدد الأسابيع المتبقية، ووزّعها بالتساوي",0,0,["وزّع"],"وزّع الدروس المتبقية على الأسابيع"),
 ("ثانيًا: رتّب حسب الأهمية",0,0,["الأهمية"],"٢. رتّب حسب الأهمية"),
 ("ركّز على المفاهيم الأساسية، واجمع الدروس المتشابهة في درس واحد",0,0,["الأساسيات…","المتشابه"],"ركّز على الأساسيات… واجمع المتشابه"),
 ("ثالثًا: اقلب الصف",0,0,["اقلب"],"٣. اقلب الصف"),
 ("اجعل الشرح في مقطع قصير قبل الحصة، ووقت الحصة للتطبيق",0,0,["قبل","للتطبيق"],"الشرح قبل الحصة… والحصة للتطبيق"),
 ("رابعًا: وزّع المهام على مجموعات",0,0,["مجموعات"],"٤. وزّع المهام على مجموعات"),
 ("كل مجموعة تتعلّم جزءًا، وتشرحه لزملائها",0,0,["تشرحه"],"كل مجموعة تتعلّم جزءًا… وتشرحه"),
 ("خامسًا: اختصر التقويم",0,0,["اختصر"],"٥. اختصر التقويم"),
 ("بدل الاختبار الطويل، استخدم بطاقة خروج في آخر دقيقتين",0,0,["بطاقة","خروج"],"بدل الاختبار الطويل… بطاقة خروج في دقيقتين"),
 ("تذكّر",0,0,["تذكّر"],"تذكّر"),
 ("إنهاء المنهج ليس سباقًا لتغطية الصفحات",0,0,["سباقًا"],"إنهاء المنهج ليس سباقًا لتغطية الصفحات"),
 ("بل لضمان أن يتعلّم الطالب ما يهمّ فعلًا",0,0,["يتعلّم","يهمّ"],"بل لضمان أن يتعلّم الطالب ما يهمّ فعلًا"),
 ("شاركنا في التعليقات: ما الفكرة التي جرّبتها ونجحت معك؟",0,0,["شاركنا"],"شاركنا في التعليقات: ما الفكرة التي نجحت معك؟"),
]
SCENES=[("المشكلة",0,2),("الخريطة",3,4),("الأولويات",5,6),("الصف المقلوب",7,8),("التعاون",9,10),("التقويم",11,12),("الخلاصة",13,16)]
_tp=os.path.join(HERE,'phrase_times.json')
if os.path.exists(_tp):
    _T=json.load(open(_tp)); PHRASES=[(p[0],_T[i][0],_T[i][1],p[3],p[4]) for i,p in enumerate(PHRASES)]
def _letters(w): return len(re.sub(r'[ً-ْ؟!:…]','',w))
def compute(wav=os.path.join(HERE,'..','v11','voice.wav')):
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
