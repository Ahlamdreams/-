#!/usr/bin/env python3
"""Build timeline.js (word-level timings) from script.txt.

Without --words: timings are ESTIMATED from a calm narration pace (draft mode).
With --words: timings come from a transcription of the real recording
(whisper / ElevenLabs-style JSON with per-word start/end), aligned to the script.

Usage:
  python3 scripts/build_timeline.py                       # estimated draft
  python3 scripts/build_timeline.py --audio assets/voice.mp3 --words words.json
"""
import argparse, difflib, json, re, subprocess, sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DIAC = re.compile(r"[ؐ-ًؚ-ٰٟۖ-ۭـ]")


def norm(w):
    w = DIAC.sub("", w)
    w = re.sub(r"[^ء-ي0-9a-zA-Z]", "", w)
    w = re.sub("[إأآٱ]", "ا", w)
    return w.replace("ى", "ي").replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")


def parse_script():
    phrases = []
    for line in (ROOT / "script.txt").read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        m = re.match(r"\[([\w-]+)\]\s*(.+)", line)
        if not m:
            sys.exit(f"bad line: {line}")
        words = []
        for tok in m.group(2).split():
            key = tok.startswith("*") and "*" in tok[1:]
            words.append({"w": tok.replace("*", ""), "k": key})
        phrases.append({"scene": m.group(1), "words": words})
    return phrases


def estimate(phrases, start=1.4):
    t = start
    for p in phrases:
        p["start"] = t
        for w in p["words"]:
            letters = len(norm(w["w"]))
            d = 0.20 + 0.062 * letters + (0.08 if w["k"] else 0)
            w["t"], w["e"] = round(t, 3), round(t + d, 3)
            t += d
            if "…" in w["w"]:
                t += 0.85
            elif w["w"].endswith(("،", ",")):
                t += 0.35
            elif w["w"].endswith("."):
                t += 0.7
        p["end"] = p["words"][-1]["e"]
        t += 0.35
    return phrases


def load_words(path):
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(data, dict):
        if "words" in data:
            data = data["words"]
        elif "segments" in data:
            data = [w for s in data["segments"] for w in s.get("words", [])]
    out = []
    for w in data:
        if w.get("type") and w["type"] != "word":
            continue
        txt = w.get("word") or w.get("text") or ""
        if norm(txt):
            out.append({"n": norm(txt), "s": float(w["start"]), "e": float(w["end"])})
    return out


def align(phrases, rec):
    flat = [w for p in phrases for w in p["words"]]
    a = [norm(w["w"]) for w in flat]
    b = [r["n"] for r in rec]
    sm = difflib.SequenceMatcher(None, a, b, autojunk=False)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal" or (tag == "replace" and i2 - i1 == j2 - j1):
            for k in range(i2 - i1):
                flat[i1 + k]["t"], flat[i1 + k]["e"] = rec[j1 + k]["s"], rec[j1 + k]["e"]
        elif tag == "replace" and j2 > j1:  # spread unequal block over its recorded span
            s, e = rec[j1]["s"], rec[j2 - 1]["e"]
            n = i2 - i1
            for k in range(n):
                flat[i1 + k]["t"] = s + (e - s) * k / n
                flat[i1 + k]["e"] = s + (e - s) * (k + 1) / n
    # interpolate anything still missing
    known = [i for i, w in enumerate(flat) if "t" in w]
    if not known:
        sys.exit("alignment failed: no words matched")
    matched = sum(1 for w in flat if "t" in w)
    for i, w in enumerate(flat):
        if "t" in w:
            continue
        prev = max([k for k in known if k < i], default=None)
        nxt = min([k for k in known if k > i], default=None)
        s = flat[prev]["e"] if prev is not None else 0.0
        e = flat[nxt]["t"] if nxt is not None else s + 0.5 * (i - (prev or 0))
        gap = (i - (prev if prev is not None else -1)) / ((nxt if nxt is not None else i + 1) - (prev if prev is not None else -1))
        w["t"] = s + (e - s) * gap
        w["e"] = w["t"] + 0.3
    for w in flat:
        w["t"], w["e"] = round(w["t"], 3), round(w["e"], 3)
    for p in phrases:
        p["start"], p["end"] = p["words"][0]["t"], p["words"][-1]["e"]
    print(f"aligned {matched}/{len(flat)} script words to the recording")
    return phrases


def duration(audio):
    out = subprocess.run(["ffprobe", "-v", "error", "-show_entries", "format=duration",
                          "-of", "csv=p=0", str(audio)], capture_output=True, text=True, check=True)
    return float(out.stdout.strip())


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--audio", help="original voice recording (kept untouched)")
    ap.add_argument("--words", help="word-timestamp JSON of the recording")
    ap.add_argument("--outro", type=float, default=5.0, help="seconds for the end card")
    ap.add_argument("--match-audio", action="store_true",
                    help="never extend past the recording: end card overlaps its tail")
    args = ap.parse_args()

    phrases = parse_script()
    phrases = align(phrases, load_words(args.words)) if args.words else estimate(phrases)
    last = phrases[-1]["end"]
    if args.audio:
        adur = duration(args.audio)
        total = adur if args.match_audio else max(adur, last + 0.9 + args.outro)
        outro = min(last + 0.9, total - 3.5) if args.match_audio else last + 0.9
    else:
        adur, outro = None, last + 0.9
        total = outro + args.outro
    tl = {
        "estimated": not args.words,
        "audio": args.audio and str(Path(args.audio).as_posix()),
        "audioDuration": adur,
        "duration": round(total, 3),
        "outro": round(outro, 3),
        "phrases": phrases,
    }
    (ROOT / "timeline.js").write_text(
        "// generated by scripts/build_timeline.py — do not edit by hand\nwindow.TIMELINE = "
        + json.dumps(tl, ensure_ascii=False, indent=1) + ";\n", encoding="utf-8")
    print(f"timeline.js: {len(phrases)} phrases, duration {total:.1f}s, outro at {outro:.1f}s"
          + (" (ESTIMATED timings)" if not args.words else ""))


if __name__ == "__main__":
    main()
