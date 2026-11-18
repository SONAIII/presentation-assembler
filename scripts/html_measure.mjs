#!/usr/bin/env node
/* Обмер одного слайда HTML-дека → boxes/texts/images/svgs/tables JSON (координаты в px канвы 1920×1080).
 * Использование: node measure.mjs <deck.html> <slideIndex0> <out.json>
 * Инжект измерителя + headless Chrome --dump-dom.
 * seq — сквозной DOM-порядок отрисовки: билдер восстанавливает по нему z-order.
 * <table> уходит СТРУКТУРОЙ (tables): сетка/ячейки/границы/раны — билдер соберёт
 * нативный редактируемый a:tbl; потомки таблицы в общий обход не попадают,
 * КРОМЕ элементов с собственным фоном (chip/mark в ячейке) — они остаются
 * оверлеями (box + text поверх таблицы), их текст из ранов ячейки исключён.
 * data-ppt-role снимается у текстов и svg (role) — явная семантика для
 * переноса готового HTML: title / body / eyebrow / decor (custom-html.md). */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { pathToFileURL } from 'node:url';
import { execFileSync } from 'node:child_process';

// Путь к Chrome. Питон-вызыватель (lib/html2pptx.py) проставляет CHROME_PATH сам;
// в standalone-запуске повторяем порядок резолва lib/chrome.py: env → playwright →
// /opt/pw-browsers → штатные пути (фидбек «перевёрстка Байк» п.10, облако/Linux).
const PW_BINARIES = [
  'chrome-linux/chrome',
  'chrome-mac/Chromium.app/Contents/MacOS/Chromium',
  'chrome-win/chrome.exe',
];
function scanPwRoot(root) {
  if (!root) return null;
  let bases = [];
  try { bases = fs.readdirSync(root).filter(n => n.startsWith('chromium-'))
    .sort().reverse().map(n => path.join(root, n)); } catch { /* нет корня */ }
  for (const base of bases)
    for (const rel of PW_BINARIES) {
      const p = path.join(base, rel);
      if (fs.existsSync(p)) return p;
    }
  for (const rel of [...PW_BINARIES, 'chromium', 'chrome']) {
    const p = path.join(root, rel);
    if (fs.existsSync(p)) return p;
  }
  return null;
}
function findChrome() {
  if (process.env.CHROME_PATH && fs.existsSync(process.env.CHROME_PATH))
    return process.env.CHROME_PATH;
  const home = process.env.HOME || '';
  return scanPwRoot(process.env.PLAYWRIGHT_BROWSERS_PATH)
    || scanPwRoot('/opt/pw-browsers')
    // дефолтные кэши playwright (0.21.0, зеркало lib/chrome.py)
    || scanPwRoot(path.join(home, 'Library/Caches/ms-playwright'))
    || scanPwRoot(path.join(home, '.cache/ms-playwright'))
    || scanPwRoot(path.join(process.env.LOCALAPPDATA || '', 'ms-playwright'))
    || [
      '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
      '/Applications/Chromium.app/Contents/MacOS/Chromium',
      '/usr/bin/google-chrome', '/usr/bin/chromium-browser', '/usr/bin/chromium',
      // win-пути Chrome (0.21.0); форки (ЯБраузер/Edge) в авто-подбор не входят —
      // headless ЯБ виснет (живая проба 2026-07-23, зеркало lib/chrome.py)
      'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
      'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
    ].find(p => fs.existsSync(p));
}
const CHROME = findChrome();
if (!CHROME) { console.error('Chrome/Chromium не найден'); process.exit(2); }

const file = path.resolve(process.argv[2]);
const slideIdx = Number(process.argv[3] || 0);
const outJson = path.resolve(process.argv[4] || 'measure.json');
let html = fs.readFileSync(file, 'utf8');
html = html.replace(/\bsrc=(['"])(?![a-z][a-z0-9+.-]*:|\/|#)(.*?)\1/gi,
  (whole, quote, value) => `src=${quote}${pathToFileURL(path.resolve(path.dirname(file), value)).href}${quote}`);

function assertSafeHtml(value) {
  const forbiddenTag = /<(script|iframe|object|embed|link|base|form|input|button|textarea|select|video|audio|source|track|foreignObject)\b/i;
  if (forbiddenTag.test(value)) throw new Error('custom HTML contains active/forbidden tag');
  if (/\son[a-z]+\s*=/i.test(value)) throw new Error('custom HTML contains event handler');
  if (/(?:src|href|action|formaction|poster|data)\s*=\s*["']?\s*(?:https?:|ftp:|wss?:|javascript:|\/\/)/i.test(value))
    throw new Error('custom HTML contains external/active URL');
  if (/@import\b/i.test(value) || /url\(\s*["']?\s*(?:https?:|ftp:|wss?:|javascript:|\/\/)/i.test(value))
    throw new Error('custom HTML CSS contains external URL');
}
assertSafeHtml(html);

const MEASURE = `
<script>
window.addEventListener('load',()=>{ (document.fonts?document.fonts.ready:Promise.resolve()).then(()=>setTimeout(()=>{
  const deck=document.getElementById('deck');
  deck.style.transform='none';deck.style.position='static';
  const slides=[...deck.querySelectorAll('.slide')];
  slides.forEach((s,i)=>{s.style.display=(i===${slideIdx})?'flex':'none';});
  const root=slides[${slideIdx}];
  const rootBox=root.getBoundingClientRect();
  const rel=b=>({x:+(b.left-rootBox.left).toFixed(2),y:+(b.top-rootBox.top).toFixed(2),w:+b.width.toFixed(2),h:+b.height.toFixed(2)});
  const parseC=c=>{if(!c||c==='transparent')return null;const m=c.match(/[\\d.]+/g);if(!m)return null;const v=m.map(Number);if(v.length===4&&v[3]===0)return null;return v;};
  const isBlock=cs=>/block|flex|grid|table|list-item/.test(cs.display);
  // рамка-клип: ближайший предок с overflow hidden/clip (для радиуса вуалей и кропа фото)
  const clipAncestor=el=>{let e=el.parentElement;while(e&&e!==root){const c=getComputedStyle(e);
    if(/hidden|clip/.test(c.overflow))return e;e=e.parentElement;}return null;};
  const range=document.createRange();
  const effOp=e0=>{let o=1,e=e0;while(e&&e.getAttribute){o*=parseFloat(getComputedStyle(e).opacity);e=e.parentElement;}return +o.toFixed(3);};
  const ownBg=el=>{const c=getComputedStyle(el);
    return !!(parseC(c.backgroundColor)||(c.backgroundImage&&c.backgroundImage.startsWith('linear-gradient')));};
  // оверлей в таблице: ближайший предок-или-сам с СОБСТВЕННЫМ фоном (до <table>)
  const overlayRoot=el=>{let e=el;while(e&&e.tagName!=='TABLE'){if(ownBg(e))return e;e=e.parentElement;}return null;};

  // текст-нода → ран (стиль от styleEl)
  const pushRun=(runs,n,styleEl)=>{
    const t=n.textContent;if(!t.trim())return;
    const scs=getComputedStyle(styleEl);
    range.selectNodeContents(n);
    let u=null;
    for(const r of range.getClientRects())if(r.width>0.5)u=u?{left:Math.min(u.left,r.left),top:Math.min(u.top,r.top),right:Math.max(u.right,r.right),bottom:Math.max(u.bottom,r.bottom)}:{left:r.left,top:r.top,right:r.right,bottom:r.bottom};
    runs.push({text:t.replace(/\\s+/g,' '),family:scs.fontFamily.split(',')[0].replace(/"/g,''),
      sizePx:parseFloat(scs.fontSize),weight:+scs.fontWeight,color:parseC(scs.color),
      italic:scs.fontStyle==='italic',effOp:effOp(styleEl),
      rect:u?rel({left:u.left,top:u.top,width:u.right-u.left,height:u.bottom-u.top}):null});
  };
  // строки: client rects текстовых узлов поддерева, группировка по y;
  // walkInline=true — только инлайн-потомки (обычные блоки), false — всё
  // поддерево (ячейка таблицы); skipOverlay — не заходить в оверлеи
  const lineRectsOf=(el,{deep=false,skipOverlay=false}={})=>{
    const lr=[];
    (function rects(node){
      for(const n of node.childNodes){
        if(n.nodeType===3&&n.textContent.trim()){
          range.selectNodeContents(n);
          for(const r of range.getClientRects())if(r.width>0.5)lr.push(r);
        } else if(n.nodeType===1){
          if(skipOverlay&&ownBg(n))continue;
          if(!deep&&isBlock(getComputedStyle(n)))continue;
          rects(n);
        }
      }
    })(el);
    return lr;
  };
  const groupLines=lr=>{
    lr.sort((a,b)=>a.top-b.top||a.left-b.left);
    const lines=[];
    for(const r of lr){
      const L=lines[lines.length-1];
      if(L&&Math.abs(r.top-L.top)<3){L.left=Math.min(L.left,r.left);L.right=Math.max(L.right,r.right);L.bottom=Math.max(L.bottom,r.bottom);}
      else lines.push({top:r.top,left:r.left,right:r.right,bottom:r.bottom});
    }
    return lines.map(L=>rel({left:L.left,top:L.top,width:L.right-L.left,height:L.bottom-L.top}));
  };
  const sideOf=(cs,p)=>{const w=parseFloat(cs['border'+p+'Width'])||0;return w>0?{w,c:parseC(cs['border'+p+'Color'])}:null;};
  // border-radius: computed может быть в % ("50%") или парой "150px 100px" — резолвим в px по боксу
  const parseRad=(cs,b)=>{const v=String(cs.borderTopLeftRadius).trim().split(/\\s+/);
    const one=(t,len)=>t.endsWith('%')?(parseFloat(t)||0)*len/100:(parseFloat(t)||0);
    let rx=one(v[0],b.width),ry=one(v[1]||v[0],b.height);
    // кламп отрисовки (радиусы больше бокса браузер сжимает пропорционально):
    // computed хранит заявленное (999px), считаем ИСПОЛЬЗУЕМОЕ
    const f=Math.min(1,rx?b.width/(2*rx):1,ry?b.height/(2*ry):1);
    return {rx:+(rx*f).toFixed(2),ry:+(ry*f).toFixed(2)};};

  // -------- table → структура для нативного a:tbl --------
  const captureTable=(el,seq)=>{
    const b=el.getBoundingClientRect();
    const rowEls=[...el.rows];
    const occupied={};
    const cells=[];
    let nCols=0;
    rowEls.forEach((tr,ri)=>{
      let ci=0;
      for(const td of tr.cells){
        while(occupied[ri+','+ci])ci++;
        const rs=td.rowSpan||1,cs2=td.colSpan||1;
        for(let dr=0;dr<rs;dr++)for(let dc=0;dc<cs2;dc++)occupied[(ri+dr)+','+(ci+dc)]=true;
        const tcs=getComputedStyle(td);
        // раны ячейки: всё поддерево БЕЗ оверлеев (у них свой box+text поверх);
        // блочные дети (div в ячейке) входят ранами с br на границах блоков
        const runs=[];
        (function collect(node){
          for(const n of node.childNodes){
            if(n.nodeType===3){pushRun(runs,n,node.nodeType===1?node:td);}
            else if(n.nodeType===1){
              if(n.tagName==='BR'){runs.push({br:true});continue;}
              const ncs=getComputedStyle(n);
              if(ncs.display==='none')continue;
              if(ownBg(n))continue;              // оверлей (chip/mark) — не в ячейку
              if(String(n.tagName).toLowerCase()==='svg')continue;
              if(isBlock(ncs)&&runs.filter(r=>!r.br).length)runs.push({br:true});
              collect(n);
            }
          }
        })(td);
        const lines=groupLines(lineRectsOf(td,{deep:true,skipOverlay:true}));
        cells.push({r:ri,c:ci,rowSpan:rs,colSpan:cs2,bbox:rel(td.getBoundingClientRect()),
          bg:parseC(tcs.backgroundColor),
          sides:{t:sideOf(tcs,'Top'),r:sideOf(tcs,'Right'),b:sideOf(tcs,'Bottom'),l:sideOf(tcs,'Left')},
          padT:parseFloat(tcs.paddingTop)||0,padR:parseFloat(tcs.paddingRight)||0,
          padB:parseFloat(tcs.paddingBottom)||0,padL:parseFloat(tcs.paddingLeft)||0,
          align:tcs.textAlign,valign:tcs.verticalAlign,
          lineHeightPx:parseFloat(tcs.lineHeight)||null,
          runs:runs.filter((r,i,a)=>!(r.br&&(i===0||i===a.length-1||(a[i+1]&&a[i+1].br)))),
          nLines:lines.length,lines});
        ci+=cs2;
      }
      nCols=Math.max(nCols,ci);
    });
    // ширины колонок: ячейка colSpan=1 в столбце; фолбэк — равные доли остатка
    const colW=[];
    for(let c=0;c<nCols;c++){
      const cell=cells.find(x=>x.c===c&&x.colSpan===1);
      colW.push(cell?cell.bbox.w:null);
    }
    const known=colW.filter(w=>w!==null).reduce((a,w)=>a+w,0),miss=colW.filter(w=>w===null).length;
    for(let c=0;c<nCols;c++)if(colW[c]===null)colW[c]=Math.max(10,(b.width-known)/miss);
    const rowH=rowEls.map(tr=>+tr.getBoundingClientRect().height.toFixed(2));
    return {seq,bbox:rel(b),cls:String(el.className).slice(0,60),
      nRows:rowEls.length,nCols,colW:colW.map(w=>+w.toFixed(2)),rowH,cells};
  };

  const boxes=[],texts=[],images=[],svgs=[],tables=[],lines2=[],charts=[];let seq=0;
  const all=[root,...root.querySelectorAll('*')];
  for(const el of all){
    if(el.tagName==='SCRIPT'||el.tagName==='STYLE')continue;
    // svg: корневой захватываем целиком (растеризация), внутренности пропускаем
    if(el.parentElement&&el.parentElement.closest('svg'))continue;
    // data-chart: НАТИВНЫЙ PowerPoint Chart (0.35.0) — бокс + JSON-спека;
    // потомки = заглушка для браузерного превью, в .pptx не конвертируются
    if(el.parentElement&&el.parentElement.closest('[data-chart]'))continue;
    if(el.hasAttribute&&el.hasAttribute('data-chart')){
      const b=el.getBoundingClientRect();
      let spec=null,err=null;
      try{spec=JSON.parse(el.getAttribute('data-chart'));}
      catch(e){err='data-chart не парсится как JSON: '+e.message;}
      charts.push({seq:seq++,bbox:rel(b),chart:spec,chartError:err,
        role:el.getAttribute('data-ppt-role')||null,
        cls:String(el.className).slice(0,60)});
      continue;
    }
    if(String(el.tagName).toLowerCase()==='svg'){
      const b=el.getBoundingClientRect();
      if(b.width>=1&&b.height>=1)svgs.push({seq:seq++,bbox:rel(b),svg:el.outerHTML,
        role:el.getAttribute('data-ppt-role')||null,
        cls:String(el.parentElement?el.parentElement.className:'').slice(0,60)});
      continue;
    }
    const cs=getComputedStyle(el);
    if(cs.visibility==='hidden'||cs.display==='none')continue;
    const b=el.getBoundingClientRect();
    if(b.width<1||b.height<1)continue;
    // -------- data-line: НАТИВНАЯ линия/стрелка (cxnSp), не box --------
    // Толщина = меньшая сторона бокса, направление = большая; цвет — background.
    // Токены в значении: arrow | arrow-start | arrow-both | dashed.
    if(el.hasAttribute('data-line')&&!el.closest('table')){
      lines2.push({seq:seq++,bbox:rel(b),color:parseC(cs.backgroundColor)||[35,15,10],
        spec:el.getAttribute('data-line')||'',cls:String(el.className).slice(0,60)});
      continue;
    }
    // -------- table: структурой, потомки — только оверлеи --------
    if(el.tagName==='TABLE'){tables.push(captureTable(el,seq++));continue;}
    const tAnc=el.closest('table');
    const inTable=tAnc&&tAnc!==el;
    let ovl=null;
    if(inTable){
      ovl=overlayRoot(el);
      if(!ovl)continue;              // обычное содержимое ячейки — уже в tables
    }
    // -------- img: рамка = клип-предок (грамматика: фото заполняет рамку) --------
    if(el.tagName==='IMG'){
      const fr=clipAncestor(el);
      const frb=fr?fr.getBoundingClientRect():b;
      images.push({seq:seq++,frame:rel(frb),img:rel(b),src:el.currentSrc,
        natW:el.naturalWidth,natH:el.naturalHeight,fit:cs.objectFit,
        radius:fr?parseRad(getComputedStyle(fr),frb).rx:parseRad(cs,b).rx,
        effOp:+cs.opacity,cls:String(el.className).slice(0,60)});
      continue;
    }
    // -------- box (фон / рамки по сторонам / градиент) --------
    const bg=parseC(cs.backgroundColor);
    const grad=cs.backgroundImage&&cs.backgroundImage.startsWith('linear-gradient')?cs.backgroundImage:null;
    const S={t:sideOf(cs,'Top'),r:sideOf(cs,'Right'),b:sideOf(cs,'Bottom'),l:sideOf(cs,'Left')};
    const sw=x=>x?x.w:0;
    const uniform=sw(S.t)===sw(S.r)&&sw(S.t)===sw(S.b)&&sw(S.t)===sw(S.l);
    if(el!==root&&(bg||grad||S.t||S.r||S.b||S.l)){
      let {rx:radius,ry:radiusY}=parseRad(cs,b);
      if(radius<0.5){const fr=clipAncestor(el);           // вуаль inset:0 в скруглённой рамке
        if(fr){const frb=fr.getBoundingClientRect();
          if(Math.abs(frb.left-b.left)<2&&Math.abs(frb.top-b.top)<2&&Math.abs(frb.width-b.width)<2&&Math.abs(frb.height-b.height)<2)
            ({rx:radius,ry:radiusY}=parseRad(getComputedStyle(fr),frb));}}
      boxes.push({seq:seq++,bbox:rel(b),bg,grad,radius,radiusY,
        borderW:uniform?sw(S.t):0,borderC:uniform&&S.t?S.t.c:null,
        sides:uniform?null:{t:S.t,r:S.r,b:S.b,l:S.l},
        opacity:+cs.opacity,cls:String(el.className).slice(0,60)});
    }
    // -------- text (блочные контейнеры; инлайны собираются ранами;
    //           инлайн-оверлей в таблице (mark) эмитит текст принудительно) --------
    if(!isBlock(cs)&&!(inTable&&ovl===el))continue;
    const runs=[];
    (function collect(node,styleEl){
      for(const n of node.childNodes){
        if(n.nodeType===3){pushRun(runs,n,styleEl);}
        else if(n.nodeType===1){
          if(n.tagName==='BR'){runs.push({br:true});continue;}
          const ncs=getComputedStyle(n);
          if(ncs.display==='none')continue;
          if(isBlock(ncs))continue;           // блочный ребёнок эмитит свой item
          if(String(n.tagName).toLowerCase()==='svg')continue;
          collect(n,n);                        // инлайн — раны его стилем
        }
      }
    })(el,el);
    if(!runs.filter(r=>!r.br).length)continue;
    const lines=groupLines(lineRectsOf(el));
    if(!lines.length)continue;
    // <li> — семантика списка для нативных буллитов (O2 хэндоффа Ultima
    // 2026-08-11): вид ul/ol + порядковый номер среди li-сиблингов (продолжение
    // нумерации <ol>: каждый li конвертируется своим текстбоксом)
    const liKind=el.tagName==='LI'?((el.parentElement&&el.parentElement.tagName==='OL')?'ol':'ul'):null;
    texts.push({seq:seq++,bbox:rel(b),runs,align:cs.textAlign,lineHeightPx:parseFloat(cs.lineHeight)||null,
      nLines:lines.length,lines,
      li:liKind,liIdx:liKind?[...el.parentElement.children].filter(c=>c.tagName==='LI').indexOf(el):0,
      padL:parseFloat(cs.paddingLeft)||0,padR:parseFloat(cs.paddingRight)||0,padT:parseFloat(cs.paddingTop)||0,
      role:el.getAttribute('data-ppt-role')||null,
      cls:String(el.className).slice(0,60)});
  }
  const rootCs=getComputedStyle(root);
  const out={slide:{w:rootBox.width,h:rootBox.height,bg:parseC(rootCs.backgroundColor)},boxes,texts,images,svgs,tables,lines:lines2,charts};
  document.body.innerHTML='<pre id=R>'+btoa(unescape(encodeURIComponent(JSON.stringify(out))))+'</pre>';
},300));});
</script>`;

const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'presentation-assembler-html-'));
const measureFile = path.join(tempDir, 'measure.html');
// Зонд — перед </body>, а во фрагменте без обёртки документа — просто в конец
// (браузер сам достроит implied body; O6.3 хэндоффа Ultima 2026-08-11: раньше
// replace молча не срабатывал и обмер падал «измерение не вернулось»)
fs.writeFileSync(measureFile, html.includes('</body>')
  ? html.replace('</body>', '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; img-src file: data:; style-src \'unsafe-inline\'; font-src file: data:; script-src \'unsafe-inline\'; connect-src \'none\'; object-src \'none\'; frame-src \'none\'; base-uri \'none\'; form-action \'none\'">' + MEASURE + '\n</body>')
  : '<meta http-equiv="Content-Security-Policy" content="default-src \'none\'; img-src file: data:; style-src \'unsafe-inline\'; font-src file: data:; script-src \'unsafe-inline\'; connect-src \'none\'; object-src \'none\'; frame-src \'none\'; base-uri \'none\'; form-action \'none\'">' + html + '\n' + MEASURE);
// Флаги — эмпирика живых проб 2026-07-23/24 (0.21.0): СВОЙ --user-data-dir в
// headless=new НЕ давать — с ним виснут и скриншот-режим, и dump-dom с
// --virtual-time-budget (Chrome сам делает эфемерный профиль, он стабилен);
// mock-keychain/basic гасят невидимый Keychain-диалог macOS на свежем профиле;
// disable-background-networking и соседи глушат GoogleUpdater-возню между
// последовательными запусками. Зеркало — html_preview.py и html2pptx.py.
let dom;
try {
  const securityFlags = ['--disable-background-networking','--disable-component-update',
    '--disable-sync','--metrics-recording-only','--host-resolver-rules=MAP * ~NOTFOUND'];
  if (process.env.PRESENTATION_ASSEMBLER_ALLOW_UNSANDBOXED_CHROME === '1') securityFlags.push('--no-sandbox');
  dom = execFileSync(CHROME, ['--headless=new','--disable-gpu',
    '--no-first-run','--no-default-browser-check',
    '--use-mock-keychain','--password-store=basic',
    ...securityFlags,
    '--virtual-time-budget=8000','--window-size=1920,1200','--dump-dom','file://'+measureFile],
    {encoding:'utf8', maxBuffer:1e8, timeout:120000});
} finally {
  fs.rmSync(tempDir, {recursive:true, force:true});
}
const m = dom.match(/<pre id="?R"?>([A-Za-z0-9+/=]+)<\/pre>/);
if(!m){ console.error('измерение не вернулось'); process.exit(2); }
const data = JSON.parse(Buffer.from(m[1],'base64').toString('utf8'));
fs.writeFileSync(outJson, JSON.stringify(data,null,1));
console.log(`ok: slide ${slideIdx}: boxes=${data.boxes.length} texts=${data.texts.length} images=${data.images.length} svgs=${data.svgs.length} tables=${data.tables.length} lines=${(data.lines||[]).length} charts=${(data.charts||[]).length} → ${outJson}`);
