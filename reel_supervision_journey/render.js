const {chromium}=require('playwright');const {spawn}=require('child_process');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1080,height:1920}});
p.on('pageerror',e=>console.log('ERR',e.message));
await p.goto('file://'+__dirname+'/trip.html');await p.waitForFunction('window.READY');await p.waitForTimeout(800);
const fps=30,D=await p.evaluate('DURATION'),n=Math.round(D*fps);
const ff=spawn('ffmpeg',['-y','-f','image2pipe','-framerate',''+fps,'-i','-','-i','mix.wav','-c:v','libx264','-preset','medium','-crf','18','-pix_fmt','yuv420p','-c:a','aac','-b:a','192k','-shortest','-movflags','+faststart','trip.mp4'],{stdio:['pipe','ignore','inherit']});
for(let i=0;i<n;i++){await p.evaluate(t=>seek(t),i/fps);const buf=await p.screenshot({type:'jpeg',quality:95});
 if(!ff.stdin.write(buf))await new Promise(r=>ff.stdin.once('drain',r)); if(i%150==0)console.log(i,'/',n);}
ff.stdin.end();await new Promise(r=>ff.on('close',r));await b.close();console.log('done')})();
