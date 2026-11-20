"""Original patriotic march (30 s, 120 BPM, Bb major) + Omani-style frame drums + SFX.
Bar = 2 s. Synced to the reel's scenes."""
import numpy as np
from scipy.signal import butter, sosfilt, fftconvolve
import wave, sys

SR = 44100
DUR = 30.0
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

# ---------------- composition ----------------
T = 70  # Bb4 tonic
# chords per bar (root offsets & triad) bars 1..15
I, IV, V, vi = [0, 4, 7], [5, 9, 12], [7, 11, 14], [9, 12, 16]
chords = [I, I, V, I, I, IV, V, I, vi, IV, V, I, I, V, I]

# Intro: timpani roll + strings swell
add(roll(2.0), 0.0, 0.35)
for c in (I,):
    for o in c:
        add(strings(T - 12 + o, 3.6), 0.0, 0.18)
# fanfare bar 1-4 (brass)
fan = [(0.0, 7, .25), (0.25, 7, .25), (0.5, 7, .25), (1.0, 12, 1.0),  # call at 2s
       ]
fanfare = [  # (time, semis, dur)
    (2.0, -5, 0.33), (2.33, -5, 0.17), (2.5, 0, 0.5), (3.0, 4, 0.5), (3.5, 7, 0.5),
    (4.0, 12, 1.25), (5.25, 11, 0.25), (5.5, 9, 0.25), (5.75, 11, 0.25),
    (6.0, 7, 0.5), (6.5, 4, 0.5), (7.0, 7, 0.5), (7.5, 2, 0.5),
]
for (t0, s, d) in fanfare:
    add(brass(T + s, d), t0, 0.32, -0.15)
    add(brass(T + s - 12, d, 1800), t0, 0.18, 0.2)
add(timpani(46), 2.0, 0.5); add(timpani(41), 4.0, 0.55); add(cymbal(2.5), 4.0, 0.25)
add(timpani(46), 6.0, 0.4); add(timpani(41), 7.0, 0.4)
for o in V: add(strings(T - 12 + o, 1.9), 6.0, 0.15)
for o in I: add(strings(T - 12 + o, 1.9), 4.0, 0.15)

# Main theme bars 5..12 (8..24 s), climax bars 13..14 (24..28), final 15
theme = [
    [(0, 1.5), (2, .5), (4, 1), (7, 1)],
    [(9, 1.5), (7, .5), (5, 1), (9, 1)],
    [(7, 1.5), (5, .5), (4, 1), (2, 1)],
    [(0, 2), (4, 1), (7, 1)],
    [(9, 1.5), (7, .5), (9, 1), (12, 1)],
    [(12, 1.5), (9, .5), (5, 1), (9, 1)],
    [(7, 1), (9, .5), (7, .5), (5, 1), (2, 1)],
    [(0, 3), (None, 1)],
    [(12, 1.5), (14, .5), (16, 1), (12, 1)],
    [(14, 1.5), (12, .5), (11, 1), (14, 1)],
]
for bi, bar in enumerate(theme):
    t0 = 8.0 + bi * 2.0; cur = t0
    for s, d in bar:
        if s is not None:
            dd = d * BEAT
            g = 0.36 if bi >= 8 else 0.30
            add(brass(T + s, dd * 0.95), cur, g, -0.1)
            if bi >= 4:  # doubled an octave lower in 2nd half
                add(brass(T + s - 12, dd * 0.95, 1600), cur, 0.16, 0.25)
            if bi in (0, 1, 2, 3):  # oud-like doubling in first half
                add(pluck(T + s), cur, 0.12, 0.45)
        cur += d * BEAT

# final chord bar 15 (28-30)
for o in I + [12]:
    add(brass(T + o - 12, 1.9, 2600, vib=True), 28.0, 0.17)
add(brass(T + 12, 1.9), 28.0, 0.3)
for o in I: add(strings(T - 12 + o, 1.8), 28.0, 0.2)
add(timpani(46), 28.0, 0.6); add(cymbal(3.0), 28.0, 0.35); add(bass_drum(), 28.0, 0.6)

# accompaniment: march pattern bars 5-14 (8-28 s)
for bar in range(4, 14):
    t0 = bar * 2.0; ch = chords[bar]
    root = T - 24 + ch[0]
    for b in range(4):
        tb = t0 + b * BEAT
        # bass oom-pah
        bn = root if b % 2 == 0 else root + 7
        add(brass(bn, 0.4, 900, vib=False), tb, 0.30)
        # brass stabs on off-beats
        if b % 2 == 1:
            for o in ch: add(brass(T - 12 + o, 0.22, 2200, vib=False), tb, 0.08, 0.3)
        # drums
        if b % 2 == 0: add(bass_drum(), tb, 0.55)
        else: add(snare(), tb, 0.35, 0.1)
        add(snare(0.12, 0.25), tb + 0.375, 0.25, 0.1)
    # Omani frame-drum pattern (eighths with accents), from bar 7
    if bar >= 6:
        pat = ['D', 't', 't', 'D', 't', 'D', 't', 't']
        for k, p in enumerate(pat):
            tk = t0 + k * 0.25
            if p == 'D': add(doum(), tk, 0.45, -0.3)
            else: add(tak(), tk, 0.28, 0.35)
        add(clap(), t0 + 0.5, 0.35, 0.0); add(clap(), t0 + 1.5, 0.35, 0.0)
    # strings pad
    for o in ch: add(strings(T - 12 + o, 1.95), t0, 0.07)
    if bar in (8, 12): add(cymbal(2.0), t0, 0.22)

# timpani roll into climax
add(roll(1.0, 41), 23.0, 0.4)
add(cymbal(3.0), 24.0, 0.35)

# SFX synced to visuals
for tw in (3.6, 7.6, 19.6, 23.6):
    add(whoosh(0.9), tw, 0.35, 0.0)
for tc in (10.0, 12.0, 14.0, 16.0, 18.0):  # counters finish
    add(chime(94), tc, 0.5, 0.2)
for tp in np.arange(20.4, 23.0, 0.35)[:6]:  # heritage icons pop
    add(chime(89 + 2 * int((tp - 20.4) / 0.35)), tp, 0.25, -0.2)

# fireworks crackles in finale
for k in range(14):
    tf = 24.6 + k * 0.38 + rng.uniform(0, 0.15)
    n = int(0.6 * SR); t = np.arange(n) / SR
    pop = hp(rng.standard_normal(n), 1500) * np.exp(-t * 18) * (rng.random(n) > 0.6)
    add(pop, tf, 0.18, rng.uniform(-0.7, 0.7))

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
