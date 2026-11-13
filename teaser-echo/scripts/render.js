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
const FPS = +arg("fps", 30), WORKERS = +arg("workers", 3), _unused = 0;
const OUT = path.resolve(ROOT, arg("out", "out/teaser-echo.mp4"));
const TMP = path.join(ROOT, "out", ".parts");
fs.mkdirSync(TMP, { recursive: true });

function readCues() {
  const src = fs.readFileSync(path.join(ROOT, "cues.js"), "utf8");
  return JSON.parse(src.slice(src.indexOf("{"), src.lastIndexOf("}") + 1));
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
  const cues = readCues(), dur = cues.duration;
  const total = Math.round(dur * FPS);
  console.log(`Rendering ${total} frames @${FPS}fps (${dur}s) with ${WORKERS} workers…`);
  const browser = await chromium.launch();
  const per = Math.ceil(total / WORKERS);
  const parts = await Promise.all([...Array(WORKERS).keys()].map(i => renderPart(browser, i, i * per, Math.min(total, (i + 1) * per))));
  await browser.close();
  const list = path.join(TMP, "list.txt");
  fs.writeFileSync(list, parts.map(p => `file '${p}'`).join("\n"));
  const silent = path.join(TMP, "video.mp4");
  execFileSync("ffmpeg", ["-y", "-loglevel", "error", "-f", "concat", "-safe", "0", "-i", list, "-c", "copy", silent]);
  execFileSync("ffmpeg", ["-y", "-loglevel", "error", "-i", silent, "-i", path.join(ROOT, "assets/soundtrack.wav"),
    "-map", "0:v", "-map", "1:a", "-c:v", "copy", "-c:a", "aac", "-b:a", "192k", "-ar", "48000", "-t", dur.toFixed(3), "-movflags", "+faststart", OUT], { stdio: "inherit" });
  parts.concat([silent, list]).forEach(f => fs.unlinkSync(f));
  console.log("✔ " + path.relative(ROOT, OUT));
})().catch(e => { console.error(e); process.exit(1); });
