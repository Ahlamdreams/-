#!/usr/bin/env python3
"""Offline word timing without a transcription service.

Splits the recording into speech segments at its pauses (ffmpeg silencedetect),
splits script.txt into clauses at punctuation, matches clauses to segments with
dynamic programming (by length), then spreads words inside each segment by letter count.

  python3 scripts/align_silence.py assets/voice.mp3 > words.json
"""
import json, math, re, subprocess, sys
sys.path.insert(0, __import__("os").path.dirname(__file__))
from build_timeline import parse_script, norm

audio = sys.argv[1]
noise = sys.argv[2] if len(sys.argv) > 2 else "-35dB"
log = subprocess.run(["ffmpeg", "-i", audio, "-af", f"silencedetect=noise={noise}:d=0.3", "-f", "null", "-"],
                     capture_output=True, text=True).stderr
dur = float(subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "csv=p=0", audio],
                           capture_output=True, text=True).stdout)
starts = [float(x) for x in re.findall(r"silence_start: ([\d.]+)", log)]
ends = [float(x) for x in re.findall(r"silence_end: ([\d.]+)", log)]
# speech segments = gaps between silences
segs, cur = [], 0.0
for s, e in zip(starts, ends):
    if s - cur > 0.15:
        segs.append((cur, s))
    cur = e
if dur - cur > 0.3:
    segs.append((cur, dur))

phr = parse_script()
words = [w["w"] for p in phr for w in p["words"]]
ends_of_phrase, acc = set(), 0
for p in phr:
    acc += len(p["words"]); ends_of_phrase.add(acc - 1)
chunks, buf = [], []
for i, w in enumerate(words):
    buf.append(i)
    if re.search(r"[…،.,]$", w) or i in ends_of_phrase:
        chunks.append(buf); buf = []
if buf:
    chunks.append(buf)
clen = [sum(len(norm(words[i])) + 1.5 for i in c) for c in chunks]
rate = sum(e - s for s, e in segs) / sum(clen)

# DP: each step pairs a run of segments with a run of clauses (one side has length 1)
INF = float("inf")
n, m = len(chunks), len(segs)
pre = [0]
for L in clen:
    pre.append(pre[-1] + L)
dp = [[INF] * (n + 1) for _ in range(m + 1)]
bk = [[None] * (n + 1) for _ in range(m + 1)]
dp[0][0] = 0
for j in range(1, m + 1):
    for i in range(1, n + 1):
        moves = [(1, b) for b in range(1, 7)] + [(a, 1) for a in range(2, 5)]
        for a, b in moves:
            j0, i0 = j - a, i - b
            if j0 < 0 or i0 < 0 or dp[j0][i0] == INF:
                continue
            sd = segs[j - 1][1] - segs[j0][0]
            c = dp[j0][i0] + math.log(sd / ((pre[i] - pre[i0]) * rate)) ** 2 + (0.15 * (a - 1))
            if c < dp[j][i]:
                dp[j][i], bk[j][i] = c, (j0, i0)
if dp[m][n] == INF:
    sys.exit(f"cannot align {n} clauses to {m} speech segments")
out, j, i, assign = [], m, n, []
while j > 0:
    j0, i0 = bk[j][i]
    assign.append((i0, i, segs[j0:j])); j, i = j0, i0
for k, i, sg in reversed(assign):
    idx = [w for c in chunks[k:i] for w in c]
    if len(sg) == 1 or i - k > 1:
        s, e = sg[0][0], sg[-1][1]
        lens = [len(norm(words[w])) + 1.5 for w in idx]
        t = s
        for w, L in zip(idx, lens):
            d = (e - s) * L / sum(lens)
            out.append({"word": words[w], "start": round(t, 3), "end": round(t + d, 3)}); t += d
    else:  # one clause spoken across several segments: spread over speech time only
        speech = sum(e - s for s, e in sg)
        lens = [len(norm(words[w])) + 1.5 for w in idx]
        pos = 0.0
        def at(x):
            for s, e in sg:
                if x <= e - s: return s + x
                x -= e - s
            return sg[-1][1]
        for w, L in zip(idx, lens):
            d = speech * L / sum(lens)
            out.append({"word": words[w], "start": round(at(pos), 3), "end": round(at(pos + d), 3)}); pos += d
print(f"{n} clauses → {m} speech segments", file=sys.stderr)
json.dump({"words": out}, sys.stdout, ensure_ascii=False, indent=1)
