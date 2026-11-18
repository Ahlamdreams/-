W,H=794,1123
o=[];A=o.append
def pill(x,y,t,fs=15,pad=14,fill='#fff8ee',col='#3b2a1f',stroke='#d9c3a3'):
    w=len(t)*fs*0.56+pad*2;h=fs+14
    A(f'<g filter="url(#sh)"><rect x="{x-w/2}" y="{y-h/2}" width="{w}" height="{h}" rx="{h/2}" fill="{fill}" stroke="{stroke}" stroke-width="1.2"/></g><text x="{x}" y="{y+fs*0.36}" class="pl" font-size="{fs}" fill="{col}">{t}</text>')
def plant(x,y,pink=False):
    A(f'<circle cx="{x}" cy="{y}" r="7" fill="#6aa84f"/><circle cx="{x-2}" cy="{y-2}" r="3.5" fill="{"#f28cb1" if pink else "#93c86b"}"/><circle cx="{x+3}" cy="{y+1}" r="3" fill="{"#f7b3c9" if pink else "#7dbb55"}"/>')
def vtable(cx,cy,name,side,w=120,h=36,color='#efe1f5',sub=None):
    cxo=cx+w/2+12 if side=='r' else cx-w/2-12
    A(f'<rect x="{cxo-9}" y="{cy-10}" width="18" height="20" rx="4" fill="#8a6e9e"/>')
    bx=cxo+6 if side=='r' else cxo-12
    A(f'<rect x="{bx}" y="{cy-13}" width="6" height="26" rx="3" fill="#5d4470"/>')
    A(f'<rect x="{cx-w/2}" y="{cy-h/2}" width="{w}" height="{h}" rx="5" fill="{color}" stroke="#a987bb" stroke-width="1.4" filter="url(#sh2)"/>')
    A(f'<rect x="{cx-w/2+2}" y="{cy-h/2+2}" width="{w-4}" height="7" rx="3" fill="#fff" opacity=".8"/>')
    A(f'<text x="{cx}" y="{cy+7}" class="tn">{name}</text>')
    if sub:
        tx=cx+w/2+8
        A(f'<text x="{tx}" y="{cy+5}" class="lb" font-size="13" fill="#fff" style="text-anchor:end">{sub}</text>')
def htable(cx,cy,name,w=84,h=36):
    A(f'<rect x="{cx-10}" y="{cy-h/2-24}" width="20" height="18" rx="4" fill="#8a6e9e"/><rect x="{cx-13}" y="{cy-h/2-29}" width="26" height="6" rx="3" fill="#5d4470"/>')
    A(f'<rect x="{cx-w/2}" y="{cy-h/2}" width="{w}" height="{h}" rx="5" fill="#efe1f5" stroke="#a987bb" stroke-width="1.4" filter="url(#sh2)"/>')
    A(f'<rect x="{cx-w/2+2}" y="{cy-h/2+2}" width="{w-4}" height="7" rx="3" fill="#fff" opacity=".8"/>')
    A(f'<text x="{cx}" y="{cy+7}" class="tn">{name}</text>')
def vstrip(x,y1,y2,t,rot,w=34,fs=16.5):
    cy=(y1+y2)/2
    A(f'<g filter="url(#sh)"><rect x="{x-w/2}" y="{y1}" width="{w}" height="{y2-y1}" rx="{w/2}" fill="#fff8ee" stroke="#c4a265" stroke-width="1.3"/></g>')
    A(f'<text transform="translate({x} {cy}) rotate({rot})" x="0" y="{fs*0.36}" class="pl" font-size="{fs}" fill="#4a2a5c">{t}</text>')
def hstrip(x1,x2,y,t,h=32):
    A(f'<g filter="url(#sh)"><rect x="{x1}" y="{y-h/2}" width="{x2-x1}" height="{h}" rx="{h/2}" fill="#fff8ee" stroke="#c4a265" stroke-width="1.3"/></g><text x="{(x1+x2)/2}" y="{y+6}" class="pl" font-size="16.5" fill="#4a2a5c">{t}</text>')

A(f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}">
<defs>
<filter id="sh" x="-20%" y="-30%" width="140%" height="170%"><feDropShadow dx="0" dy="1.5" stdDeviation="1.6" flood-opacity=".22"/></filter>
<filter id="sh2" x="-20%" y="-30%" width="140%" height="170%"><feDropShadow dx="0" dy="1.4" stdDeviation="1.2" flood-opacity=".25"/></filter>
<pattern id="grass" width="10" height="10" patternUnits="userSpaceOnUse"><rect width="10" height="10" fill="#4e9a3a"/><path d="M0 10L10 0" stroke="#58a542" stroke-width="2.4"/></pattern>
<pattern id="pave" width="16" height="16" patternUnits="userSpaceOnUse"><rect width="16" height="16" fill="#f6e8dc"/><path d="M0 0H16M0 0V16" stroke="#ead6c6" stroke-width=".8"/></pattern>
</defs>
<style>.pl{{font-family:Tajawal;font-weight:800;text-anchor:middle}}.tn{{font-family:Tajawal;font-weight:800;font-size:18px;fill:#4a2a5c;text-anchor:middle}}.lb{{font-family:Tajawal;font-weight:700;text-anchor:middle}}</style>
<rect width="{W}" height="{H}" fill="#f3e2d6"/>''')
# header
A('<g filter="url(#sh)"><rect x="14" y="10" width="766" height="70" rx="20" fill="#fff8ee" stroke="#c4a265" stroke-width="1.5"/></g>')
A('<image href="school-logo-t.png" x="26" y="15" width="78" height="60" preserveAspectRatio="xMidYMid meet"/>')
A('<text x="420" y="42" class="pl" font-size="22" fill="#4a2a5c">خريطة توزيع طاولات المعلمات في ساحة المدرسة</text>')
A('<text x="420" y="67" class="pl" font-size="16" fill="#8a6e9e">مدرسة أم مالك الأنصارية للبنات (5-6)</text>')
A('<g transform="translate(752 45)"><circle r="17" fill="#fff" stroke="#3b2a1f" stroke-width="1.2"/><path d="M0 -12 L5 6 L0 2 L-5 6Z" fill="#3b2a1f"/><text y="-20" class="pl" font-size="0">N</text></g>')
# field
A('<rect x="12" y="90" width="770" height="806" rx="14" fill="url(#pave)"/>')
A('<rect x="20" y="96" width="754" height="794" rx="12" fill="url(#grass)"/>')
A('<rect x="28" y="104" width="738" height="778" rx="8" fill="none" stroke="#fff" stroke-width="2.2" stroke-dasharray="8 6" opacity=".9"/>')
ar=['وفاء','آمنة','شمسة','بشرى','هدى','إيمان']
en=['شيخة','سوسن','مريم','ثريا','أسماء','فاطمة']
ma=['ضنينة','ولاء','رفيدة','رقية','بلقيس']
isl=['موزة','عائشة']
so=['ريهام','سحر','عائشة']
sc=['نورة','أمل','خديجة']
sk=[('جوهرة','فنون'),('إيمان','رياضة'),('حليمة','موسيقى'),('رويدا','تقنية'),('حسناء','تقنية')]
# top row
hstrip(286,736,128,'الرياضيات')
for x,n in zip([690,600,510,420,330],ma): htable(x,196,n)
hstrip(66,244,128,'التربية الإسلامية')
for x,n in zip([200,110],isl): htable(x,196,n)
# direction pills
pill(632,250,'جهة اليمين',17,fill='#0b6b3a',col='#fff',stroke='#0b6b3a')
pill(162,250,'جهة اليسار',17,fill='#b5121b',col='#fff',stroke='#b5121b')
RX,LX=632,162
eny=[528,482,436,390,344,298]
for y,n in zip(eny,en): vtable(RX,y,n,'r')
vstrip(748,278,548,'اللغة الإنجليزية',90)
ary=[838,792,746,700,654,608]
for y,n in zip(ary,ar): vtable(RX,y,n,'r')
vstrip(748,588,858,'اللغة العربية',90)
for y,n in zip([298,344,390],so): vtable(LX,y,n,'l')
vstrip(46,268,420,'الدراسات الاجتماعية',-90,fs=15)
for y,n in zip([462,508,554],sc): vtable(LX,y,n,'l')
vstrip(46,440,576,'العلوم',-90)
for y,(n,sb) in zip([630,676,722,768,814],sk): vtable(LX,y,n,'l',color='#e3f1fb' if sb=='تقنية' else '#efe1f5',sub=sb)
vstrip(46,608,836,'المهارات وتقنية المعلومات',-90)
pill(397,520,'ساحة المدرسة',24,pad=24)
# flag
A('<ellipse cx="397" cy="862" rx="28" ry="9" fill="#e2c8b6"/><line x1="397" y1="862" x2="397" y2="812" stroke="#555" stroke-width="3"/>')
A('<g transform="translate(398 812)"><rect width="40" height="9" fill="#fff" stroke="#ddd" stroke-width=".5"/><rect y="9" width="40" height="9" fill="#c8102e"/><rect y="18" width="40" height="9" fill="#007a3d"/><rect width="12" height="27" fill="#c8102e"/></g>')
plant(362,862,True);plant(432,862,True)
pill(397,775,'سارية العلم',14)
# reception
A('<g filter="url(#sh)"><rect x="14" y="910" width="766" height="200" rx="16" fill="#efe0d2" stroke="#c9b49a" stroke-width="1.5"/></g>')
A('<rect x="26" y="934" width="492" height="164" rx="10" fill="#fbf4ea" stroke="#c9b49a" stroke-width="2"/>')
A('<rect x="528" y="934" width="240" height="164" rx="10" fill="#f7ecdf" stroke="#c9b49a" stroke-width="2"/>')
pill(397,912,'الاستقبال',19,pad=26,fill='#4a2a5c',col='#fff',stroke='#4a2a5c')
pill(272,958,'استقبال مجلس أولياء الأمور',16)
for cx in (150,394):
    for y in (1000,1050):
        A(f'<rect x="{cx-70}" y="{y-14}" width="140" height="28" rx="4" fill="#efe1f5" stroke="#b99ac7"/>')
        for dx in (-40,0,40): A(f'<rect x="{cx+dx-8}" y="{y-10}" width="16" height="10" rx="1.5" fill="#3d4f7a"/>')
        for dx in (-30,30): A(f'<rect x="{cx+dx-8}" y="{y+16}" width="16" height="9" rx="3" fill="#e9e3dc" stroke="#cfc4b8" stroke-width=".8"/>')
plant(272,1005,True);plant(272,1050,True);plant(40,950);plant(504,1085,True)
pill(648,958,'الأخصائية الاجتماعية',16)
A('<rect x="598" y="990" width="100" height="36" rx="4" fill="#d9b48a" stroke="#b78f63"/><rect x="638" y="996" width="20" height="12" rx="1" fill="#3d4f7a"/>')
A('<rect x="638" y="1031" width="22" height="15" rx="4" fill="#8a6e9e"/>')
A('<rect x="572" y="1060" width="62" height="22" rx="9" fill="#c7b2d6"/><rect x="662" y="1060" width="62" height="22" rx="9" fill="#c7b2d6"/><circle cx="648" cy="1071" r="8" fill="#e6d1b3"/>')
# entrance
A('<path d="M372 1098 h50 v14 h-50z" fill="#d9c3a3"/>')
pill(397,1108,'المدخل الرئيسي',13,pad=12)
A('</svg>')
svg='\n'.join(o)
html='''<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><title>خريطة توزيع المعلمات</title>
<link href="https://fonts.googleapis.com/css2?family=Tajawal:wght@500;700;800&display=swap" rel="stylesheet">
<style>@page{size:A4;margin:0}*{margin:0;padding:0}body{background:#f3e2d6}svg{display:block;width:210mm;height:297mm}</style></head><body>'''+svg+'</body></html>'
open('/home/user/-/map/map.html','w').write(html)
