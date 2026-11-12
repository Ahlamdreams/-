#!/usr/bin/env python3
"""Word-level transcription of the original recording (Arabic) → words.json.

  pip install faster-whisper
  python3 scripts/transcribe.py assets/voice.mp3 > words.json

Any other tool that yields per-word start/end (e.g. ElevenLabs Scribe JSON) works too:
build_timeline.py accepts {"words":[{"text"|"word", "start", "end"}]} or whisper segments.
"""
import json, sys
from faster_whisper import WhisperModel

model = WhisperModel(sys.argv[2] if len(sys.argv) > 2 else "large-v3", compute_type="int8")
segments, _ = model.transcribe(sys.argv[1], language="ar", word_timestamps=True, vad_filter=False)
words = [{"word": w.word.strip(), "start": round(w.start, 3), "end": round(w.end, 3)}
         for s in segments for w in s.words]
json.dump({"words": words}, sys.stdout, ensure_ascii=False, indent=1)
