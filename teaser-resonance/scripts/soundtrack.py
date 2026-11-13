#!/usr/bin/env python3
"""Original resonant soundtrack (music + SFX + voiceover placement), driven by cues.js.

  python3 scripts/soundtrack.py   ->  assets/soundtrack.wav (+ .m4a)
The voice recording is split at its natural pause and each phrase is placed on its cue,
untouched (no pitch / speed change).
"""
import json, re, subprocess, wave
from pathlib import Path
import numpy as np

ROOT = Path(__file__).resolve().parent.parent
SR = 48000
src = (ROOT / "cues.js").read_text(encoding="utf-8")
C = json.loads(src[src.index("{"):src.rindex("}") + 1])
D = C["duration"]
N = int(D * SR)
rng = np.random.default_rng(2264)
t_all = np.arange(N) / SR


def buf(sec): return np.zeros(int(sec * SR))
def tt(n): return np.arange(n) / SR


def add(dst, sig, at, gain=1.0):
    i = int(at * SR)
    if i >= len(dst): return
    j = min(len(dst), i + len(sig))
    dst[i:j] += sig[: j - i] * gain


def lowpass(x, cutoff):
    X = np.fft.rfft(x); f = np.fft.rfftfreq(len(x), 1 / SR)
    X *= 1 / (1 + (f / cutoff) ** 4)
    return np.fft.irfft(X, len(x))


def bandpass(x, lo, hi):
    X = np.fft.rfft(x); f = np.fft.rfftfreq(len(x), 1 / SR)
    X *= 1 / (1 + (lo / np.maximum(f, 1)) ** 4) / (1 + (f / hi) ** 4)
    return np.fft.irfft(X, len(x))


def reverb(x, sec=2.2, mix=.3):
    n = int(sec * SR)
    ir = rng.standard_normal(n) * np.exp(-np.arange(n) / SR * 3.2)
    ir = lowpass(ir, 5000); ir /= np.sqrt(np.sum(ir ** 2))
    L = 1 << int(np.ceil(np.log2(len(x) + n)))
    wet = np.fft.irfft(np.fft.rfft(x, L) * np.fft.rfft(ir, L), L)[: len(x)]
    return x * (1 - mix) + wet * mix * 1.4


# ---------- instruments ----------
def kick(gain=1.0, f0=95, f1=42, dec=.35):
    n = int(.6 * SR); t = tt(n)
    f = f1 + (f0 - f1) * np.exp(-t * 18)
    ph = 2 * np.pi * np.cumsum(f) / SR
    return np.sin(ph) * np.exp(-t / dec) * gain


def whoosh(dur=.75):
    n = int(dur * SR); t = tt(n)
    env = np.sin(np.pi * np.clip(t / dur, 0, 1)) ** 2
    return bandpass(rng.standard_normal(n), 500, 4500) * env * .5


def chime(f=880, dur=2.0):
    n = int(dur * SR); t = tt(n)
    s = sum(a * np.sin(2 * np.pi * f * m * t) * np.exp(-t * d) for m, a, d in [(1, 1, 2.2), (2.76, .45, 3.5), (5.4, .22, 5), (8.9, .1, 7)])
    return s * np.minimum(1, t / .004) * .35


def pad(freqs, dur, bright=.3):
    n = int(dur * SR); t = tt(n); s = np.zeros(n)
    for k, f in enumerate(freqs):
        for det in (-.7, .7):
            ph = rng.uniform(0, 6.28)
            s += (np.sin(2 * np.pi * (f + det) * t + ph) + bright * np.sin(4 * np.pi * (f + det) * t + ph)) / (1 + k * .3)
    return s / len(freqs)


def saw_chord(freqs, dur, cutoff=1400):
    n = int(dur * SR); t = tt(n); s = np.zeros(n)
    for f in freqs:
        for det in (-1.2, 0, 1.3):
            s += 2 * (((f + det) * t) % 1) - 1
    return lowpass(s / (3 * len(freqs)), cutoff)


def impact(chord, gain=1.0):
    n = int(2.6 * SR); t = tt(n)
    boom = kick(1.0, 70, 32, .9)
    crack = bandpass(rng.standard_normal(int(.35 * SR)), 1500, 9000) * np.exp(-tt(int(.35 * SR)) * 14)
    brass = saw_chord(chord, 2.6, 1600) * np.exp(-t * 1.4) * np.minimum(1, t / .02)
    s = np.zeros(n); s[: len(boom)] += boom * 1.3; s[: len(crack)] += crack * .6; s += brass * .9
    return s * gain


def riser(dur):
    n = int(dur * SR); t = tt(n)
    noise = rng.standard_normal(n); out = np.zeros(n); step = int(.05 * SR)
    for i in range(0, n, step):  # sweep the band upward in short blocks
        p = i / n; seg = noise[i:i + step]
        out[i:i + step] = bandpass(np.pad(seg, (0, step - len(seg))), 300 + 3000 * p, 900 + 7000 * p)[: len(seg)]
    tone = np.sin(2 * np.pi * np.cumsum(110 + 330 * (t / dur) ** 2) / SR)
    return (out * .45 + tone * .25) * (t / dur) ** 2


# ---------- music bus ----------
def echo(x, delay=.32, fb=.45, n=5):
    y = x.copy()
    for k in range(1, n + 1):
        d = int(delay * k * SR); y[d:] += x[:len(x) - d] * fb ** k
    return y


def pluck(f, dur=1.2):
    n = int(dur * SR); t = tt(n)
    s = np.sin(2 * np.pi * f * t) + .4 * np.sin(4 * np.pi * f * t) + .15 * np.sin(6 * np.pi * f * t)
    return s * np.exp(-t * 4) * np.minimum(1, t / .003) * .3


mus = np.zeros(N)
cut, hit1, hit2, end = C["cut"], C["hit1"], C["hit2"], C["end"]
# act 1: airy pad (A minor -> F -> C -> G), pulses, resonant plucks echoing on the beats
prog = [[110.0, 164.81, 220.0, 261.63], [87.31, 174.61, 220.0, 261.63], [130.81, 196.0, 261.63, 329.63], [98.0, 196.0, 246.94, 293.66]]
seg = cut / 4
for k, ch in enumerate(prog):
    p = pad(ch, seg + 1.0, .3); e = np.minimum(1, tt(len(p)) / .8) * np.minimum(1, (len(p) / SR - tt(len(p))) / 1.0)
    add(mus, lowpass(p * e, 1200 + 400 * k), k * seg, .35 + .1 * k)
notes = [440.0, 523.25, 659.25, 587.33, 523.25, 659.25, 783.99, 659.25]
plk = np.zeros(N)
for i, bt in enumerate(C["beats"]):
    add(mus, kick(.35 + .5 * i / len(C["beats"])), bt)
    add(plk, pluck(notes[i % len(notes)]), bt, .55)
mus += echo(plk, .3, .5, 5)
a_, b_ = C["riser"]; add(mus, riser(b_ - a_), a_, .7)
add(mus, impact([110.0, 164.81, 220.0, 261.63]), hit1, .75)
sus = pad([55.0, 110.0, 164.81], cut - hit1, .2); add(mus, lowpass(sus * np.linspace(.5, .9, len(sus)), 700), hit1, .5)
add(mus, riser(.45), cut - .45, .4)
# act 2: big reveal + resonant echo booms on every ring
big = impact([73.42, 110.0, 146.83, 220.0]); add(mus, echo(np.pad(big, (0, SR)), .38, .42, 4), hit2, .85)
boom = np.zeros(N)
for k, rt in enumerate(C["rings"]):
    add(boom, kick(.9, 70, 34, .7), rt); add(boom, chime([880.0, 1046.5, 1174.66, 1318.5][k], 2.2), rt, .5)
mus += echo(boom, .36, .45, 5)
dr = pad([36.71, 73.42, 110.0], end - hit2, .15); de = np.minimum(1, tt(len(dr)) / .4) * np.minimum(1, (len(dr) / SR - tt(len(dr))) / .6)
add(mus, lowpass(dr * de, 900), hit2, .5)
# act 3: elegant closing
cl = pad([146.83, 185.0, 220.0, 293.66], D - end, .35)
ce = np.minimum(1, tt(len(cl)) / 1.2) * np.minimum(1, (len(cl) / SR - tt(len(cl))) / 1.6) * .55
add(mus, lowpass(cl * ce, 2400), end)
for k, nt in enumerate(C["names"]):
    add(mus, chime([587.33, 739.99, 880.0, 1174.66][k % 4], 2.5), nt, .5)

# ---------- SFX bus ----------
sfx = np.zeros(N)
for w in C["whoosh"]: add(sfx, whoosh(), w - .25, .5)
for k, c in enumerate(C["chime"]): add(sfx, chime([1318.5, 1567.98, 1760.0][k % 3], 1.8), c, .45)
for k, c in enumerate(C["cards"]):
    add(sfx, whoosh(.45), c - .3, .45); add(sfx, chime([1046.5, 1318.5, 1567.98][k], 1.4), c, .5)
n = int(.8 * SR); add(sfx, bandpass(rng.standard_normal(n), 4000, 14000) * np.sin(np.pi * tt(n) / .8) ** 2 * .25, C["shimmer"])

music = reverb(mus + sfx, 2.6, .3)
music[int(cut * SR):int(hit2 * SR)] *= np.linspace(1, 0, int(hit2 * SR) - int(cut * SR)) ** 3   # breath before the reveal
fo = int(.5 * SR); music[-fo:] *= np.linspace(1, 0, fo)


# ---------- voice: split at its pause, place each phrase on its cue ----------
def load_voice(p):
    raw = subprocess.run(["ffmpeg", "-v", "error", "-i", str(p), "-f", "f32le", "-ac", "1", "-ar", str(SR), "-"], capture_output=True, check=True).stdout
    return np.frombuffer(raw, np.float32).astype(np.float64)


vpath = ROOT / "assets/voice.mp3"
voice = np.zeros(N)
if vpath.exists():
    v = load_voice(vpath)
    log = subprocess.run(["ffmpeg", "-i", str(vpath), "-af", "silencedetect=noise=-38dB:d=0.3", "-f", "null", "-"], capture_output=True, text=True).stderr
    st = [float(x) for x in re.findall(r"silence_start: ([\d.]+)", log)]
    en = [float(x) for x in re.findall(r"silence_end: ([\d.]+)", log)]
    sil = [(s, e) for s, e in zip(st, en) if s > 0.05 and e < len(v) / SR - 0.05]
    gap = max(sil, key=lambda x: x[1] - x[0]) if sil else (len(v) / SR / 2,) * 2
    head = en[0] if st and st[0] == 0 else 0
    segs = [(max(0, head - .05), gap[0] + .08), (gap[1] - .08, len(v) / SR)]
    for (s, e), at in zip(segs, C["voice"]):
        add(voice, v[int(s * SR):int(e * SR)], at)
    duck = np.convolve((np.abs(voice) > 1e-3).astype(float), np.ones(int(.25 * SR)) / int(.25 * SR), "same")
    music *= 1 - .45 * np.clip(duck * 3, 0, 1)
    print("voice phrases placed at", C["voice"], "from", [tuple(round(x, 2) for x in s) for s in segs])

mixed = music / (np.max(np.abs(music)) + 1e-9) * .55 + voice * 1.0
mixed = np.tanh(mixed * 1.1) / np.tanh(1.1)
mixed /= max(1.0, np.max(np.abs(mixed)) / .97)
stereo = np.stack([mixed, mixed], 1)
out = ROOT / "assets/soundtrack.wav"
with wave.open(str(out), "wb") as w:
    w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR)
    w.writeframes((stereo * 32767).astype(np.int16).tobytes())
subprocess.run(["ffmpeg", "-y", "-v", "error", "-i", str(out), "-c:a", "aac", "-b:a", "192k", str(out.with_suffix(".m4a"))], check=True)
print("wrote", out.relative_to(ROOT), f"{D:.1f}s")
