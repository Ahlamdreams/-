#!/usr/bin/env node
/*
 * Render index.html frame-by-frame to a 1080x1920 MP4 and mux the ORIGINAL voice
 * recording (bit-identical samples, only re-encoded to AAC) + optional soft music.
 *
 *   node scripts/render.js                  # 30fps, out/father-reel.mp4
 *   node scripts/render.js --fps 30 --workers 3 --music-volume 0.10 --no-music
 */
const { chromium } = require("playwright");
const { spawn, execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");

const ROOT = path.resolve(__dirname, "..");
const arg = (k, d) => { const i = process.argv.indexOf("--" + k); return i > 0 ? process.argv[i + 1] : d; };
const flag = k => process.argv.includes("--" + k);
const FPS = +arg("fps", 30), WORKERS = +arg("workers", 3), MUSIC_VOL = +arg("music-volume", 0.10);
const OUT = path.resolve(ROOT, arg("out", "out/father-reel.mp4"));
const TMP = path.join(ROOT, "out", ".parts");
fs.mkdirSync(TMP, { recursive: true });

function readTimeline() {
  const src = fs.readFileSync(path.join(ROOT, "timeline.js"), "utf8");
  return JSON.parse(src.slice(src.indexOf("{"), src.lastIndexOf("}") + 1));
}
function readConfig() {
  const src = fs.readFileSync(path.join(ROOT, "config.js"), "utf8");
  const m = k => (src.match(new RegExp(k + ':\\s*"([^"]*)"')) || [])[1];
  return { music: m("music") };
}

async function renderPart(browser, idx, f0, f1) {
  const page = await browser.newPage({ viewport: { width: 1080, height: 1920 }, deviceScaleFactor: 1 });
  page.on("pageerror", e => console.error("page error:", e.message));
  await page.goto("file://" + path.join(ROOT, "index.html") + "?render");
  await page.evaluate(() => window.READY);
  const file = path.join(TMP, `part${idx}.mp4`);
  const ff = spawn("ffmpeg", ["-y", "-loglevel", "error", "-f", "image2pipe", "-framerate", String(FPS), "-c:v", "png", "-i", "-",
    "-c:v", "libx264", "-preset", "medium", "-crf", "17", "-pix_fmt", "yuv420p", "-r", String(FPS), file], { stdio: ["pipe", "inherit", "inherit"] });
  for (let f = f0; f < f1; f++) {
    await page.evaluate(t => window.seekTo(t), f / FPS);
    const buf = await page.screenshot({ type: "png" });
    if (!ff.stdin.write(buf)) await new Promise(r => ff.stdin.once("drain", r));
    if (idx === 0 && f % (FPS * 5) === 0) console.log(`  frame ${f}/${f1} (worker 0)`);
  }
  ff.stdin.end();
  await new Promise((res, rej) => ff.on("close", c => c === 0 ? res() : rej(new Error("ffmpeg " + c))));
  await page.close();
  return file;
}

(async () => {
  const tl = readTimeline(), cfg = readConfig();
  const total = Math.round(tl.duration * FPS);
  console.log(`Rendering ${total} frames @${FPS}fps (${tl.duration}s) with ${WORKERS} workers…`);
  const browser = await chromium.launch();
  const per = Math.ceil(total / WORKERS);
  const parts = await Promise.all([...Array(WORKERS).keys()].map(i => renderPart(browser, i, i * per, Math.min(total, (i + 1) * per))));
  await browser.close();

  const list = path.join(TMP, "list.txt");
  fs.writeFileSync(list, parts.map(p => `file '${p}'`).join("\n"));
  const silent = path.join(TMP, "video.mp4");
  execFileSync("ffmpeg", ["-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", list, "-c", "copy", silent]);

  // audio: original voice untouched (volume 1.0, no filters) + optional quiet music bed
  const inputs = ["-i", silent], filters = [];
  let n = 1, voiceIdx = null, musicIdx = null;
  if (tl.audio) { inputs.push("-i", path.join(ROOT, tl.audio)); voiceIdx = n++; }
  const musicPath = cfg.music && path.join(ROOT, cfg.music);
  if (!flag("no-music") && musicPath && fs.existsSync(musicPath)) { inputs.push("-stream_loop", "-1", "-i", musicPath); musicIdx = n++; }
  const D = tl.duration.toFixed(3);
  let map = [];
  if (musicIdx !== null) filters.push(`[${musicIdx}:a]atrim=0:${D},volume=${MUSIC_VOL},afade=t=in:d=2,afade=t=out:st=${(tl.duration - 3).toFixed(2)}:d=3[m]`);
  if (voiceIdx !== null && musicIdx !== null) { filters.push(`[${voiceIdx}:a][m]amix=inputs=2:normalize=0:duration=longest[a]`); map = ["-map", "[a]"]; }
  else if (voiceIdx !== null) map = ["-map", `${voiceIdx}:a`];
  else if (musicIdx !== null) map = ["-map", "[m]"];
  const args = ["-y", "-loglevel", "error", ...inputs];
  if (filters.length) args.push("-filter_complex", filters.join(";"));
  args.push("-map", "0:v", ...map, "-c:v", "copy");
  if (map.length) args.push("-c:a", "aac", "-b:a", "192k", "-ar", "48000");
  args.push("-t", D, "-movflags", "+faststart", OUT);
  execFileSync("ffmpeg", args, { stdio: "inherit" });
  parts.concat([silent, list]).forEach(f => fs.unlinkSync(f));
  console.log("✔ " + path.relative(ROOT, OUT));
})().catch(e => { console.error(e); process.exit(1); });
