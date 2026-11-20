// usage: node render.js stills 2,6,14,22,27   |  node render.js video out.mp4
const { chromium } = require('playwright');
const { spawn } = require('child_process');
const path = require('path');
(async () => {
  const mode = process.argv[2];
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1080, height: 1920 }, deviceScaleFactor: 1 });
  await page.goto('file://' + path.join(__dirname, 'reel.html') + '#render');
  await page.waitForFunction(() => window.READY === true);
  await page.waitForTimeout(300);
  if (mode === 'stills') {
    for (const t of process.argv[3].split(',').map(Number)) {
      await page.evaluate(t => render(t), t);
      await page.screenshot({ path: path.join(__dirname, `still_${t}.jpg`), type: 'jpeg', quality: 80 });
    }
  } else {
    const FPS = 30, DUR = 40, out = process.argv[3];
    const ff = spawn('ffmpeg', ['-y', '-f', 'image2pipe', '-framerate', String(FPS), '-c:v', 'mjpeg', '-i', '-',
      '-i', path.join(__dirname, 'mix.wav'),
      '-c:v', 'libx264', '-preset', 'slow', '-crf', '18', '-pix_fmt', 'yuv420p', '-profile:v', 'high',
      '-c:a', 'aac', '-b:a', '192k', '-movflags', '+faststart', '-shortest', out], { stdio: ['pipe', 'inherit', 'inherit'] });
    for (let f = 0; f < FPS * DUR; f++) {
      await page.evaluate(t => render(t), f / FPS);
      const buf = await page.screenshot({ type: 'jpeg', quality: 95 });
      if (!ff.stdin.write(buf)) await new Promise(r => ff.stdin.once('drain', r));
      if (f % 150 === 0) console.error('frame', f);
    }
    ff.stdin.end();
    await new Promise(r => ff.on('close', r));
  }
  await browser.close();
})();
