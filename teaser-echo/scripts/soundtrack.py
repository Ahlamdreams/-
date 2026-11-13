#!/usr/bin/env python3
"""Original cinematic soundtrack (music + SFX + voiceover placement), driven by cues.js.

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
mus = np.zeros(N)
# act 1 (0 → cut): dark pad building + heartbeat pulses accelerating
cut = C["cut"]
p1 = pad([55.0, 82.41, 130.81, 164.81], cut, .25)  # A minor colour
env = np.clip(t_all[: len(p1)] / 1.5, 0, 1) * (0.18 + 0.5 * (t_all[: len(p1)] / cut) ** 1.6)
mus[: len(p1)] += lowpass(p1 * env, 900) * .9
for i, b in enumerate(C["beats"]):
    add(mus, kick(.45 + .45 * i / len(C["beats"])), b)
    if i % 2: add(mus, kick(.18, 120, 60, .12), b + .28)
a, b_ = C["riser"]; add(mus, riser(b_ - a), a, .7)
# act 2: silence, then golden-dot heartbeats with a tense drone
dot, hit1, hit2, stop, end = C["dot"], C["hit1"], C["hit2"], C["stop"], C["end"]
dr = pad([36.71, 55.0, 73.42], hit1 - dot, .1) * np.linspace(.15, .6, int((hit1 - dot) * SR))
add(mus, lowpass(dr, 500), dot)
for k, pt in enumerate(C["pulse4"]):
    add(mus, kick(.8 + .05 * k, 80, 38, .3), pt)
add(mus, riser(hit1 - dot - .05), dot, .55)
add(mus, impact([73.42, 110.0, 146.83, 174.61]), hit1, .9)          # D minor hit
add(mus, impact([87.31, 130.81, 174.61, 220.0], ), hit2, .8)      # F major lift
# hard stop
mus[int(stop * SR):int(end * SR)] = 0
# act 3: elegant closing (D major pad + bell arpeggio on each name)
cl = pad([146.83, 185.0, 220.0, 293.66], D - end, .35)
ce = np.minimum(1, tt(len(cl)) / 1.2) * np.minimum(1, (len(cl) / SR - tt(len(cl))) / 1.8) * .55
add(mus, lowpass(cl * ce, 2400), end)
for k, nt in enumerate(C["names"]):
    add(mus, chime([587.33, 739.99, 880.0, 1174.66][k % 4], 2.5), nt, .55)
add(mus, chime(293.66, 3), end, .5)

# ---------- SFX bus ----------
sfx = np.zeros(N)
for w in C["whoosh"]: add(sfx, whoosh(), w - .25, .55)
for k, c in enumerate(C["chime"]): add(sfx, chime([1318.5, 1567.98, 1760.0][k % 3], 1.8), c, .5)
sz = C["sizzle"]; n = int(.9 * SR)
add(sfx, bandpass(rng.standard_normal(n), 3000, 12000) * np.exp(-tt(n) * 4) * .4, sz)
for k in range(8): add(sfx, chime(2000 + 300 * k, .6), sz + k * .07, .15)
add(sfx, chime(1046.5, 1.4), dot, .4)
for pt in C["pulse4"]: add(sfx, chime(1975.5, .4), pt, .12)
add(sfx, whoosh(.25)[::-1], stop - .22, .6)

music = reverb(mus + sfx, 2.4, .28)
music[int(cut * SR):int(dot * SR)] *= 0.0           # the half-second of silence
music[int(stop * SR):int(end * SR)] *= 0.0          # abrupt disappearance


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
