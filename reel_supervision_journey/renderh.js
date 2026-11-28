const {chromium}=require('playwright');const {spawn}=require('child_process');
const only=process.argv.slice(2).map(Number);
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1080,height:1920}});
p.on('pageerror',e=>console.log('ERR',e.message));
await p.goto('file://'+__dirname+'/hybrid.html');await p.waitForFunction('window.READY');await p.waitForTimeout(800);
if(only.length){for(const t of only){await p.evaluate(t=>seek(t),t);await p.screenshot({path:`ov_${t}.png`,omitBackground:true});
  await new Promise(r=>spawn('ffmpeg',['-v','error','-y','-ss',''+t,'-i','bg2.mp4','-i',`ov_${t}.png`,'-filter_complex','[0][1]overlay','-frames:v','1',`tr_${t}.jpg`]).on('close',r))}await b.close();return}
const fps=30,n=Math.round(44*fps);
const ff=spawn('ffmpeg',['-y','-i','bg2.mp4','-f','image2pipe','-framerate',''+fps,'-c:v','png','-i','-','-i','mix.wav','-filter_complex','[0:v][1:v]overlay=format=auto,format=yuv420p[v]','-map','[v]','-map','2:a','-c:v','libx264','-crf','18','-preset','medium','-c:a','aac','-b:a','192k','-shortest','-movflags','+faststart','hybrid.mp4'],{stdio:['pipe','ignore','inherit']});
for(let i=0;i<n;i++){await p.evaluate(t=>seek(t),i/fps);const buf=await p.screenshot({type:'png',omitBackground:true});
 if(!ff.stdin.write(buf))await new Promise(r=>ff.stdin.once('drain',r));}
ff.stdin.end();await new Promise(r=>ff.on('close',r));await b.close();console.log('done')})();
