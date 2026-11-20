"""Original patriotic march (30 s, 120 BPM, Bb major) + Omani-style frame drums + SFX.
Bar = 2 s. Synced to the reel's scenes."""
import numpy as np
from scipy.signal import butter, sosfilt, fftconvolve
import wave, sys

SR = 44100
DUR = 40.0
N = int(SR * DUR)
BEAT = 0.5
L = np.zeros(N); R = np.zeros(N)
rng = np.random.default_rng(7)

def hz(m): return 440.0 * 2 ** ((m - 69) / 12)

def lp(x, fc, order=2):
    return sosfilt(butter(order, fc, 'low', fs=SR, output='sos'), x)

def hp(x, fc, order=2):
    return sosfilt(butter(order, fc, 'high', fs=SR, output='sos'), x)

def bp(x, lo, hi, order=2):
    return sosfilt(butter(order, [lo, hi], 'band', fs=SR, output='sos'), x)

def add(sig, t0, gain=1.0, pan=0.0):
    i = int(t0 * SR)
    if i >= N: return
    sig = sig[: N - i]
    l = np.cos((pan + 1) * np.pi / 4); r = np.sin((pan + 1) * np.pi / 4)
    L[i:i + len(sig)] += sig * gain * l * 1.414
    R[i:i + len(sig)] += sig * gain * r * 1.414

def env(n, a, d, s, rl):
    """ADSR in seconds; n samples total (sustain portion fills)."""
    A, D, Rl = int(a * SR), int(d * SR), int(rl * SR)
    S_ = max(n - A - D - Rl, 0)
    e = np.concatenate([np.linspace(0, 1, A, False), np.linspace(1, s, D, False),
                        np.full(S_, s), np.linspace(s, 0, Rl)])
    return e[:n] if len(e) >= n else np.pad(e, (0, n - len(e)))

def brass(m, dur, bright=3000, vib=True):
    n = int((dur + 0.25) * SR); t = np.arange(n) / SR
    f = hz(m) * (1 + (0.004 * np.sin(2 * np.pi * 5.5 * t) * np.clip(t - 0.25, 0, 1) if vib else 0))
    ph = 2 * np.pi * np.cumsum(f) / SR
    x = np.zeros(n)
    for k in range(1, 14):
        if hz(m) * k > 12000: break
        x += np.sin(k * ph + 0.3 * k) / k ** 1.05
    e = env(n, 0.035, 0.12, 0.75, 0.22)
    # brightness follows envelope (brass "blat")
    x = lp(x * e, bright) * 0.6 + lp(x * e, bright * 0.4) * 0.4
    return x

def strings(m, dur):
    n = int((dur + 0.6) * SR); t = np.arange(n) / SR
    out = np.zeros(n)
    for det in (-0.08, 0.0, 0.07):
        f = hz(m + det) * (1 + 0.003 * np.sin(2 * np.pi * 5 * t + det * 40))
        ph = 2 * np.pi * np.cumsum(f) / SR
        saw = 2 * ((ph / (2 * np.pi)) % 1) - 1
        out += saw
    return lp(out * env(n, 0.25, 0.2, 0.85, 0.6), 2500) * 0.33

def pluck(m, dur=0.8):  # oud-like
    n = int(dur * SR); t = np.arange(n) / SR
    f = hz(m)
    x = sum(np.sin(2 * np.pi * f * k * t) * np.exp(-t * (4 + 3 * k)) / k for k in range(1, 9))
    x += 0.3 * lp(rng.standard_normal(n) * np.exp(-t * 80), 3000)
    return x * 0.5

def timpani(m, dur=1.2):
    n = int(dur * SR); t = np.arange(n) / SR
    f = hz(m) * (1 + 0.06 * np.exp(-t * 30))
    ph = 2 * np.pi * np.cumsum(f) / SR
    x = np.sin(ph) * np.exp(-t * 3.2) + 0.4 * np.sin(1.5 * ph) * np.exp(-t * 6)
    x += 0.5 * lp(rng.standard_normal(n), 800) * np.exp(-t * 40)
    return x

def bass_drum(dur=0.6):
    n = int(dur * SR); t = np.arange(n) / SR
    f = 50 + 90 * np.exp(-t * 35)
    return np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t * 7) * 1.2

def snare(dur=0.3, g=1.0):
    n = int(dur * SR); t = np.arange(n) / SR
    nz = bp(rng.standard_normal(n), 1500, 9000) * np.exp(-t * 22)
    body = np.sin(2 * np.pi * 190 * t) * np.exp(-t * 30)
    return (nz * 0.8 + body * 0.5) * g

def doum(dur=0.5):  # low frame-drum (Omani "ranna"/kasir)
    n = int(dur * SR); t = np.arange(n) / SR
    f = 85 + 40 * np.exp(-t * 25)
    return (np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t * 9)
            + 0.3 * lp(rng.standard_normal(n), 400) * np.exp(-t * 50))

def tak(dur=0.2):  # high slap
    n = int(dur * SR); t = np.arange(n) / SR
    return (bp(rng.standard_normal(n), 2000, 6000) * np.exp(-t * 45) * 0.7
            + np.sin(2 * np.pi * 420 * t) * np.exp(-t * 40) * 0.4)

def clap():
    n = int(0.25 * SR); t = np.arange(n) / SR
    x = np.zeros(n)
    for d in (0, 0.011, 0.022):
        i = int(d * SR); m = n - i
        x[i:] += bp(rng.standard_normal(m), 900, 4000) * np.exp(-np.arange(m) / SR * 30)
    return x * 0.6

def cymbal(dur=3.0):
    n = int(dur * SR); t = np.arange(n) / SR
    return hp(rng.standard_normal(n), 5000) * np.exp(-t * 1.4) * 0.5

def roll(dur, m=46):  # timpani roll crescendo
    n = int(dur * SR); t = np.arange(n) / SR
    x = np.zeros(n)
    for k in np.arange(0, dur, 0.055):
        h = timpani(m, 0.4) * (0.15 + 0.85 * (k / dur) ** 2)
        i = int(k * SR); x[i:i + len(h)] += h[: n - i]
    return x

def whoosh(dur=0.9):
    n = int(dur * SR); t = np.arange(n) / SR
    nz = rng.standard_normal(n)
    out = np.zeros(n); seg = 512
    for s in range(0, n, seg):
        p = s / n; fc = 300 + 5000 * np.sin(np.pi * p)
        out[s:s + seg] = bp(nz[max(0, s - 2048):s + seg], fc * 0.7, fc * 1.3)[-len(nz[s:s + seg]):]
    return out * np.sin(np.pi * t / dur) ** 2 * 0.9

def chime(m=94):
    n = int(1.5 * SR); t = np.arange(n) / SR
    return sum(np.sin(2 * np.pi * hz(m) * r * t) * np.exp(-t * (3 + 2 * i)) * a
               for i, (r, a) in enumerate([(1, 1), (2.76, .4), (5.4, .2)])) * 0.35

# ---------------- composition (40 s) ----------------
T = 70
I, IV, V, vi = [0, 4, 7], [5, 9, 12], [7, 11, 14], [9, 12, 16]
# S1 0-6: roll + fanfare
add(roll(2.0), 0.0, 0.35)
for o in I: add(strings(T - 12 + o, 3.8), 0.0, 0.18)
fanfare = [(2.0, -5, 0.33), (2.33, -5, 0.17), (2.5, 0, 0.5), (3.0, 4, 0.5), (3.5, 7, 0.5),
           (4.0, 12, 1.4)]
for (t0, s, d) in fanfare:
    add(brass(T + s, d), t0, 0.30, -0.15)
    add(brass(T + s - 12, d, 1800), t0, 0.16, 0.2)
add(timpani(46), 2.0, 0.5); add(timpani(41), 4.0, 0.55); add(cymbal(2.5), 4.0, 0.22)
for o in I: add(strings(T - 12 + o, 2.0), 4.0, 0.15)
# S2 6-16: calm oud arpeggios + pads + soft frame drum (voice section)
prog = [I, vi, IV, V, I]
for k, ch in enumerate(prog):
    t0 = 6.0 + 2 * k
    for o in ch: add(strings(T - 12 + o, 1.95), t0, 0.11)
    add(brass(T - 24 + ch[0], 1.8, 700, vib=False), t0, 0.16)
    arp = [ch[0], ch[1], ch[2], ch[1] + 12 if False else ch[0] + 12, ch[2], ch[1], ch[2], ch[0] + 12]
    for i, n in enumerate(arp):
        add(pluck(T - 12 + n), t0 + i * 0.25, 0.16, 0.35 if i % 2 else -0.35)
    if k >= 1:
        add(doum(), t0, 0.32, -0.2); add(doum(), t0 + 1.0, 0.28, -0.2)
        add(tak(), t0 + 0.5, 0.15, 0.3); add(tak(), t0 + 1.5, 0.15, 0.3)
add(roll(1.0, 41), 15.0, 0.35)
# S3/S4 16-32: march theme (8 bars)
theme = [
    [(0, 1.5), (2, .5), (4, 1), (7, 1)], [(9, 1.5), (7, .5), (5, 1), (9, 1)],
    [(7, 1.5), (5, .5), (4, 1), (2, 1)], [(0, 2), (4, 1), (7, 1)],
    [(9, 1.5), (7, .5), (9, 1), (12, 1)], [(12, 1.5), (9, .5), (5, 1), (9, 1)],
    [(7, 1), (9, .5), (7, .5), (5, 1), (2, 1)], [(0, 3), (None, 1)]]
tchords = [I, IV, V, I, vi, IV, V, I]
for bi, bar in enumerate(theme):
    t0 = 16.0 + bi * 2.0; cur = t0
    for s_, d in bar:
        if s_ is not None:
            dd = d * BEAT
            add(brass(T + s_, dd * 0.95), cur, 0.26, -0.1)
            if bi >= 4: add(brass(T + s_ - 12, dd * 0.95, 1600), cur, 0.13, 0.25)
            else: add(pluck(T + s_), cur, 0.11, 0.45)
        cur += d * BEAT
    ch = tchords[bi]; root = T - 24 + ch[0]
    for b in range(4):
        tb = t0 + b * BEAT
        add(brass(root if b % 2 == 0 else root + 7, 0.4, 900, vib=False), tb, 0.28)
        if b % 2 == 1:
            for o in ch: add(brass(T - 12 + o, 0.22, 2200, vib=False), tb, 0.07, 0.3)
        if b % 2 == 0: add(bass_drum(), tb, 0.5)
        else: add(snare(), tb, 0.3, 0.1)
        add(snare(0.12, 0.25), tb + 0.375, 0.22, 0.1)
    if bi >= 2:
        for k, p_ in enumerate(['D', 't', 't', 'D', 't', 'D', 't', 't']):
            tk = t0 + k * 0.25
            if p_ == 'D': add(doum(), tk, 0.4, -0.3)
            else: add(tak(), tk, 0.24, 0.35)
        add(clap(), t0 + 0.5, 0.3); add(clap(), t0 + 1.5, 0.3)
    for o in ch: add(strings(T - 12 + o, 1.95), t0, 0.07)
add(cymbal(2.0), 16.0, 0.25); add(cymbal(2.0), 24.0, 0.2)
add(roll(1.0, 41), 31.0, 0.4)
# S5 32-40: climax + final chord
climax = [[(12, 1.5), (14, .5), (16, 1), (12, 1)], [(14, 1.5), (12, .5), (11, 1), (14, 1)],
          [(12, 1.5), (9, .5), (5, 1), (9, 1)], [(7, 1), (9, .5), (7, .5), (5, 1), (2, 1)]]
cch = [I, V, IV, V]
for bi, bar in enumerate(climax):
    t0 = 32.0 + bi * 2.0; cur = t0
    for s_, d in bar:
        add(brass(T + s_, d * BEAT * 0.95), cur, 0.28, -0.1)
        add(brass(T + s_ - 12, d * BEAT * 0.95, 1600), cur, 0.14, 0.25)
        cur += d * BEAT
    ch = cch[bi]
    for o in ch: add(strings(T - 12 + o, 1.95), t0, 0.12)
    add(brass(T - 24 + ch[0], 1.8, 800, vib=False), t0, 0.2)
    add(timpani(46 if bi % 2 == 0 else 41), t0, 0.4); add(doum(), t0, 0.35); add(doum(), t0 + 1.0, 0.3)
    add(tak(), t0 + 0.5, 0.2); add(tak(), t0 + 1.5, 0.2)
add(cymbal(3.0), 32.0, 0.35)
for o in I + [12]: add(brass(T + o - 12, 2.0, 2600), 38.0 - 0.0, 0.0)  # placeholder silent
# final chord at 38 (under the last words)
for o in I: add(strings(T - 12 + o, 2.0), 38.0, 0.2)
add(brass(T, 1.9), 38.0, 0.22); add(brass(T - 12, 1.9, 1800), 38.0, 0.16)
add(timpani(46), 38.0, 0.5); add(cymbal(3.0), 38.0, 0.25)
# SFX
for tw in (5.6, 15.6, 21.6, 31.6): add(whoosh(0.9), tw, 0.3)
for tc in (22.8, 24.6, 26.4, 28.2, 30.0): add(chime(94), tc + 1.2, 0.4, 0.2)
for i in range(6): add(chime(89 + 2 * i), 16.4 + i * 0.35, 0.2, -0.2)
add(chime(98), 12.0, 0.35)
for k in range(18):
    tf = 32.6 + k * 0.38 + rng.uniform(0, 0.15)
    n = int(0.6 * SR); t = np.arange(n) / SR
    add(hp(rng.standard_normal(n), 1500) * np.exp(-t * 18) * (rng.random(n) > 0.6), tf, 0.14, rng.uniform(-0.7, 0.7))

# ---------------- mix ----------------
def reverb(x, rt=1.6):
    n = int(rt * SR); t = np.arange(n) / SR
    ir = rng.standard_normal(n) * np.exp(-t * 6.9 / rt)
    ir = lp(ir, 5000); ir /= np.sqrt(np.sum(ir ** 2))
    return fftconvolve(x, ir)[:len(x)]

L2 = L + 0.22 * reverb(L); R2 = R + 0.22 * reverb(R)
st = np.stack([L2, R2], 1)
st = hp(st.T, 30).T
# gentle compression via tanh
st = st / np.max(np.abs(st)) * 1.4
st = np.tanh(st)
# fade out last 0.8 s
fade = np.ones(N); fade[-int(0.8 * SR):] = np.linspace(1, 0, int(0.8 * SR))
st *= fade[:, None]
st = st / np.max(np.abs(st)) * 0.92
pcm = (st * 32767).astype(np.int16)
with wave.open(sys.argv[1], 'wb') as w:
    w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR); w.writeframes(pcm.tobytes())
print('ok')
