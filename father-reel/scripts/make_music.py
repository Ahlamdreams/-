#!/usr/bin/env python3
"""Generate a very soft ambient pad (original, royalty-free) to sit under the voice.

python3 scripts/make_music.py --seconds 90 --out assets/music.wav
ffmpeg -i assets/music.wav -c:a aac -b:a 160k assets/music.m4a   # smaller file used by config.js
"""
import argparse, wave
import numpy as np

SR = 44100
# slow, warm progression (Hz): Dmaj9 → Bm7 → Gmaj7 → Asus2
CHORDS = [
    [146.83, 220.00, 277.18, 329.63, 369.99],
    [123.47, 185.00, 246.94, 293.66, 369.99],
    [98.00, 196.00, 246.94, 293.66, 369.99],
    [110.00, 164.81, 220.00, 246.94, 329.63],
]


def pad(freqs, n, rng):
    t = np.arange(n) / SR
    out = np.zeros((n, 2))
    for i, f in enumerate(freqs):
        for ch, det in enumerate((-0.6, 0.6)):  # slight detune → stereo width
            ff = f + det
            ph = rng.uniform(0, 2 * np.pi)
            tone = (np.sin(2 * np.pi * ff * t + ph)
                    + 0.25 * np.sin(2 * np.pi * 2 * ff * t + ph)
                    + 0.08 * np.sin(2 * np.pi * 3 * ff * t + ph))
            trem = 0.85 + 0.15 * np.sin(2 * np.pi * (0.07 + 0.02 * i) * t + ph)
            out[:, ch] += tone * trem / (1 + i * 0.35)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--seconds", type=float, default=90)
    ap.add_argument("--out", default="assets/music.wav")
    a = ap.parse_args()
    rng = np.random.default_rng(7)
    total = int(a.seconds * SR)
    seg, fade = int(9.0 * SR), int(3.0 * SR)
    mix = np.zeros((total + seg, 2))
    pos, k = 0, 0
    while pos < total:
        n = seg + fade
        p = pad(CHORDS[k % len(CHORDS)], n, rng)
        env = np.ones(n)
        env[:fade] = np.sin(np.linspace(0, np.pi / 2, fade)) ** 2
        env[-fade:] = np.cos(np.linspace(0, np.pi / 2, fade)) ** 2
        end = min(pos + n, len(mix))
        mix[pos:end] += (p * env[:, None])[: end - pos]
        pos += seg
        k += 1
    mix = mix[:total]
    # gentle one-pole low-pass for warmth
    alpha = 0.08
    for ch in range(2):
        x = mix[:, ch]
        y = np.empty_like(x)
        acc = 0.0
        # vectorised enough for ~90s audio
        for i in range(0, len(x), 4096):
            blk = x[i:i + 4096]
            out = np.empty_like(blk)
            for j, v in enumerate(blk):
                acc += alpha * (v - acc)
                out[j] = acc
            y[i:i + 4096] = out
        mix[:, ch] = y
    # master fades
    fi, fo = int(3 * SR), int(4 * SR)
    mix[:fi] *= np.linspace(0, 1, fi)[:, None]
    mix[-fo:] *= np.linspace(1, 0, fo)[:, None]
    mix = mix / np.max(np.abs(mix)) * 0.5
    data = (mix * 32767).astype(np.int16)
    with wave.open(a.out, "wb") as w:
        w.setnchannels(2); w.setsampwidth(2); w.setframerate(SR)
        w.writeframes(data.tobytes())
    print(f"wrote {a.out} ({a.seconds:.1f}s)")


if __name__ == "__main__":
    main()
