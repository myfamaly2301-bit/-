<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Сверка табелей</title>
<style>
:root{box-sizing:border-box;padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px);color-scheme:light;
--bg:#f6f7f9;--fg:#1c2330;--mut:#667085;--card:#fff;--bd:#d9dee7;--ac:#2456d6;--bad:#b42318;--badbg:#fde4e1;--ok:#067647;--okbg:#dcf5e7;--warn:#fff4d6}
html{scroll-padding-top:env(safe-area-inset-top,0px)}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){color-scheme:dark;--bg:#12151b;--fg:#e6e9ef;--mut:#98a2b3;--card:#1b2029;--bd:#323a48;--ac:#7aa2ff;--bad:#ff8f83;--badbg:#4a1f1b;--ok:#6ee0a0;--okbg:#14382a;--warn:#4a3a12}}
:root[data-theme="dark"]{color-scheme:dark;--bg:#12151b;--fg:#e6e9ef;--mut:#98a2b3;--card:#1b2029;--bd:#323a48;--ac:#7aa2ff;--bad:#ff8f83;--badbg:#4a1f1b;--ok:#6ee0a0;--okbg:#14382a;--warn:#4a3a12}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--fg);font:15px/1.45 -apple-system,"Segoe UI",Roboto,Arial,sans-serif}
main{max-width:1100px;margin:0 auto;padding:20px 16px 60px}
h1{font-size:22px;margin:0 0 4px}.mut{color:var(--mut);font-size:13px}
.zones{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin:16px 0}
@media(max-width:700px){.zones{grid-template-columns:1fr}}
.zone{background:var(--card);border:2px dashed var(--bd);border-radius:12px;padding:14px;min-height:120px}
.zone.drag{border-color:var(--ac)}.zone.ok{border-style:solid}
.zone b{display:block;margin-bottom:6px}.fn{word-break:break-all;margin:4px 0 8px;font-size:13px}
button,select,input[type=search]{font:inherit;color:var(--fg);background:var(--card);border:1px solid var(--bd);border-radius:8px;padding:6px 10px}
button{cursor:pointer}select{max-width:100%}
.tag{display:inline-block;font-size:12px;border-radius:6px;padding:1px 7px;background:var(--okbg);color:var(--ok);margin-left:6px}
.opts{display:flex;flex-wrap:wrap;gap:14px;align-items:center;margin:8px 0 14px}
.chips{display:flex;flex-wrap:wrap;gap:10px;margin:10px 0}
.chip{background:var(--card);border:1px solid var(--bd);border-radius:10px;padding:8px 14px}.chip b{font-size:20px;display:block}
.chip.bad b{color:var(--bad)}.chip.good b{color:var(--ok)}
details.emp{background:var(--card);border:1px solid var(--bd);border-radius:10px;margin:6px 0}
details.emp>summary{display:grid;grid-template-columns:2.2fr 1.6fr .7fr .7fr .6fr .8fr;gap:8px;padding:9px 12px;cursor:pointer;align-items:center;list-style:none}
details.emp>summary::-webkit-details-marker{display:none}
.head{display:grid;grid-template-columns:2.2fr 1.6fr .7fr .7fr .6fr .8fr;gap:8px;padding:0 12px;font-size:12px;color:var(--mut)}
@media(max-width:760px){.head{display:none}details.emp>summary{grid-template-columns:1fr 1fr}}
.n{text-align:right;font-variant-numeric:tabular-nums}.pos{color:var(--mut);font-size:13px}
.d{color:var(--bad);font-weight:600}.z{color:var(--ok)}
.grid{overflow-x:auto;padding:4px 12px 12px}
table{border-collapse:collapse;font-size:12.5px}
th,td{border:1px solid var(--bd);padding:3px 6px;text-align:center;min-width:30px}th{background:var(--bg);font-weight:500}
td.l,th.l{text-align:left;white-space:nowrap}td.x{background:var(--badbg);color:var(--bad);font-weight:700}
.warn{background:var(--warn);border-radius:6px;padding:1px 6px;font-size:12px}
.msg{padding:10px 14px;border-radius:10px;background:var(--warn);margin:10px 0}
.only li{margin:2px 0}
</style>
</head>
<body>
<main>
<h1>Сверка табелей</h1>
<div class="mut">Загрузите два файла Excel — внутренний табель и/или табель из 1С. Формат определяется автоматически, файлы обрабатываются прямо в браузере и никуда не отправляются.</div>
<div class="zones" id="zones"></div>
<div class="opts">
 <label><input type="checkbox" id="onlyDiff" checked> только с отличиями</label>
 <label title="В 1С отпуск ставится и на выходные, во внутреннем табеле там пусто"><input type="checkbox" id="otEq"> отпуск (ОТ) на выходных не считать отличием</label>
 <input type="search" id="q" placeholder="Поиск по ФИО…">
</div>
<div id="msg"></div><div id="sum"></div><div id="out"></div>
<p class="mut">Правила сравнения: Т→Б (больничный), НН и НБ→НВ (неявка), НП, В и пусто → «нет работы». У дней с кодами ОТ/Б/БС/НВ часы не учитываются. Составные отметки сменных (Я/Я/Н/Н, часы 2/2/2/5) складываются в 11 ч. Несколько строк одного сотрудника в 1С (разные должности) объединяются по табельному номеру.</p>
</main>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<script>
/* ===== ЛОГИКА (работает на массивах строк) ===== */
const KZ={'ә':'а','і':'и','ң':'н','ғ':'г','ү':'у','ұ':'у','қ':'к','ө':'о','һ':'х','ё':'е'};
const norm=s=>String(s==null?'':s).toLowerCase().replace(/[әіңғүұқөһё]/g,c=>KZ[c]).replace(/\u00a0/g,' ');
const CODE={'ОТ':'ОТ','Б':'Б','Т':'Б','БС':'БС','НВ':'НВ','НБ':'НВ','НН':'НВ'};
const Z={h:0,c:''};
const num=v=>{const n=parseFloat(String(v==null?'':v).replace(',','.'));return isNaN(n)?0:n};
const fmt=n=>(Math.round(n*100)/100).toString().replace('.',',');
const ident=n=>{const t=norm(n).replace(/\(.*?\)/g,' ').split(/[\s.]+/).filter(Boolean);return{sn:t[0]||'',ini:t.slice(1).map(x=>x[0])}};
const isOneC=R=>R.slice(0,25).some(r=>(r||[]).some(v=>typeof v==='string'&&v.includes('Отметки о явках')));
function periodOf(R){for(const r of R.slice(0,15))for(const v of r||[]){const m=/^01\.(\d\d)\.(\d{4})$/.exec(String(v==null?'':v).trim());if(m)return{m:+m[1],y:+m[2]}}return null}

function cellInt(v){
  if(typeof v==='number')return{h:v,c:''};
  const s=String(v==null?'':v).trim().toUpperCase(); if(!s)return Z;
  if(/^[\d.,]+$/.test(s))return{h:num(s),c:''};
  return{h:0,c:CODE[s]||''};
}
function parseInternal(R){
  const out=[];let map=null,nm=1,ps=2;
  for(let r=0;r<R.length;r++){
    const row=R[r]||[];
    if(String(row[0]).trim()==='№'&&String(row[1]).includes('Ф.И.О')){
      map={};let last=0;
      for(let c=3;c<row.length;c++){const n=Number(row[c]);if(row[c]!=null&&Number.isInteger(n)&&n>=1&&n<=31&&!map[c]){map[c]=n;last=c}}
      for(let k=1;k<=3;k++)(R[r+k]||[]).forEach((v,c)=>{const n=Number(v);if(v!=null&&c>2&&c<=last+4&&Number.isInteger(n)&&n>=1&&n<=31&&!Object.values(map).includes(n)&&typeof v!=='string')map[c]=n});
      continue;
    }
    if(!map||typeof row[0]!=='number')continue;
    const name=String(row[nm]==null?'':row[nm]).replace(/\u00a0/g,' ').replace(/\s+/g,' ').trim();
    if(!name||/ваканси/i.test(name))continue;
    const days={};
    for(const c in map){const v=cellInt(row[c]);if(v.h||v.c)days[map[c]]=v}
    out.push({name,pos:String(row[ps]||'').trim(),days,...ident(name)});
  }
  return{emps:out,period:null};
}
function cell1c(code,hours){
  const c=CODE[String(code==null?'':code).trim().split('/')[0].trim().toUpperCase()]||'';
  const h=c?0:String(hours==null?'':hours).split('/').reduce((s,x)=>s+num(x),0);
  return(h||c)?{h,c}:Z;
}
function parse1c(R){
  const find=v=>R.findIndex(r=>r&&String(r[8])===v);
  const r1=find('1'),r2=find('16');const m1={},m2={};
  for(let c=8;c<45;c++){
    const a=String(R[r1][c]==null?'':R[r1][c]).trim(),b=String(R[r2][c]==null?'':R[r2][c]).trim();
    if(/^\d+$/.test(a)&&+a<=15)m1[c]=+a; if(/^\d+$/.test(b)&&+b>=16&&+b<=31)m2[c]=+b;
  }
  const by=new Map();
  for(let r=0;r<R.length-3;r++){
    const row=R[r]||[],C=row[2];
    if(!/^\d+$/.test(String(row[1]==null?'':row[1]).trim())||typeof C!=='string'||!/\(.+\)/s.test(C)||!row[4])continue;
    const name=C.replace(/\n?\(.*\)\s*$/s,'').trim(),pos=(/\((.*)\)\s*$/s.exec(C)||[])[1]||'';
    const tn=String(row[4]).trim(),days={};
    const put=(m,cr,hr)=>{for(const c in m){const v=cell1c(R[cr][c],R[hr][c]);if(v!==Z)days[m[c]]=v}};
    put(m1,r,r+1);put(m2,r+2,r+3);
    let e=by.get(tn);
    if(!e){by.set(tn,e={name,pos,tn,days:{},...ident(name)})}
    else if(!e.pos.includes(pos))e.pos+='; '+pos;
    for(const d in days){const x=e.days[d];e.days[d]=x?{h:x.h+days[d].h,c:x.c||days[d].c}:days[d]}
  }
  return{emps:[...by.values()],period:periodOf(R)};
}
function lev(a,b){const d=[...Array(b.length+1).keys()];for(let i=1;i<=a.length;i++){let p=d[0];d[0]=i;for(let j=1;j<=b.length;j++){const t=d[j];d[j]=Math.min(d[j]+1,d[j-1]+1,p+(a[i-1]===b[j-1]?0:1));p=t}}return d[b.length]}
const KW=['водит','грузч','кладов','техни','учетч','админ','операт','старш','сменн'];
function posScore(a,b){const x=norm(a.pos),y=norm(b.pos);return KW.filter(k=>x.includes(k)&&y.includes(k)).length-0.5*KW.filter(k=>x.includes(k)!==y.includes(k)).length}
function match(A,B){
  const pairs=[],ua=new Set(),ub=new Set();
  for(const pass of[0,1])for(const a of A){
    if(ua.has(a))continue;
    const cand=B.filter(b=>{if(ub.has(b))return false;
      const ok=pass?(a.sn.length>3&&lev(a.sn,b.sn)<=1):a.sn===b.sn;
      return ok&&Array.from({length:Math.min(a.ini.length,b.ini.length)},(_,i)=>a.ini[i]===b.ini[i]).every(Boolean)});
    if(!cand.length)continue;
    const b=cand.map((x,i)=>[x,posScore(a,x),i]).sort((p,q)=>q[1]-p[1]||p[2]-q[2])[0][0];
    ua.add(a);ub.add(b);pairs.push({a,b,approx:!!pass});
  }
  return{pairs,onlyA:A.filter(x=>!ua.has(x)),onlyB:B.filter(x=>!ub.has(x))};
}
function same(x,y,ot){
  if(x.c===y.c&&Math.abs(x.h-y.h)<1e-6)return true;
  return ot&&((x.c==='ОТ'&&!y.c&&!y.h)||(y.c==='ОТ'&&!x.c&&!x.h));
}
function compare(A,B,ot){
  const m=match(A,B);
  const res=m.pairs.map(({a,b,approx})=>{
    const d=[];let ta=0,tb=0;
    for(let k=1;k<=31;k++){const x=a.days[k]||Z,y=b.days[k]||Z;ta+=x.h;tb+=y.h;if(!same(x,y,ot))d.push(k)}
    return{a,b,approx,ta,tb,d};
  });
  res.sort((p,q)=>q.d.length-p.d.length);
  return{res,onlyA:m.onlyA,onlyB:m.onlyB};
}
const total=e=>Object.values(e.days).reduce((s,v)=>s+v.h,0);

/* ===== ИНТЕРФЕЙС ===== */
if(typeof document!=='undefined'){
const $=id=>document.getElementById(id);
const esc=s=>String(s==null?'':s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const S=[null,null],MON=['янв','фев','мар','апр','май','июн','июл','авг','сен','окт','ноя','дек'];
const rowsOf=ws=>{const r=XLSX.utils.decode_range(ws['!ref']||'A1');return XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null,range:XLSX.utils.encode_range({s:{r:0,c:0},e:r.e})})};
function analyze(s){
  const R=rowsOf(s.wb.Sheets[s.sheet]);
  s.fmt=isOneC(R)?'1c':'int';
  try{s.data=s.fmt==='1c'?parse1c(R):parseInternal(R);s.err=null}catch(e){s.data={emps:[],period:null};s.err='Не удалось разобрать лист «'+s.sheet+'»: '+e.message}
}
function pick(){
  S.forEach(s=>{if(s&&!s.manual){s.sheet=s.wb.SheetNames[s.wb.SheetNames.length-1];analyze(s)}});
  const per=S.map(s=>s&&s.data&&s.data.period).find(Boolean);
  if(per)S.forEach(s=>{
    if(!s||s.manual||s.wb.SheetNames.length<2)return;
    const ok=s.wb.SheetNames.filter(n=>{const l=n.toLowerCase(),y=/\d{4}/.exec(l);return l.startsWith(MON[per.m-1])&&(!y||+y[0]===per.y)});
    if(ok.length){s.sheet=ok[ok.length-1];analyze(s)}
  });
}
function zones(){
  $('zones').innerHTML=[0,1].map(i=>{const s=S[i];
    return `<div class="zone ${s?'ok':''}" data-i="${i}"><b>${i?'Второй':'Первый'} табель${s?`<span class="tag">${s.fmt==='1c'?'формат 1С':'внутренний'}</span>`:''}</b>
    <div class="fn">${s?esc(s.name):'Перетащите файл сюда или нажмите кнопку'}</div>
    ${s&&s.wb.SheetNames.length>1?`<div class="mut">Лист:</div><select data-sh="${i}">${s.wb.SheetNames.map(n=>`<option${n===s.sheet?' selected':''}>${esc(n)}</option>`).join('')}</select> `:''}
    <button data-up="${i}">${s?'Заменить файл':'Выбрать файл'}</button><input type="file" accept=".xlsx,.xlsm,.xls" hidden data-f="${i}"></div>`}).join('');
}
function load(i,file){
  const fr=new FileReader();
  fr.onload=e=>{try{S[i]={name:file.name,wb:XLSX.read(e.target.result,{type:'array'}),manual:false};pick();refresh()}catch(err){S[i]=null;zones();$('msg').innerHTML='<div class="msg">Не удалось прочитать файл: '+esc(err.message)+'</div>'}};
  fr.readAsArrayBuffer(file);
}
function refresh(){zones();render()}
let CMP=null;
function render(){
  $('msg').innerHTML='';$('sum').innerHTML='';$('out').innerHTML='';
  const bad=S.find(s=>s&&s.err);if(bad)$('msg').innerHTML='<div class="msg">'+esc(bad.err)+'</div>';
  if(!S[0]||!S[1]){$('msg').innerHTML+='<div class="msg">Загрузите оба файла, чтобы увидеть отличия.</div>';return}
  const [A,B]=[S[0].data.emps,S[1].data.emps];
  if(!A.length||!B.length){$('msg').innerHTML+='<div class="msg">На выбранном листе не найдено сотрудников. Проверьте выбранный лист.</div>';return}
  const [la,lb]=S[0].fmt!==S[1].fmt?[S[0].fmt==='1c'?'1С':'Табель',S[1].fmt==='1c'?'1С':'Табель']:['Файл 1','Файл 2'];
  CMP=compare(A,B,$('otEq').checked);
  const {res,onlyA,onlyB}=CMP,withD=res.filter(r=>r.d.length);
  const maxD=Math.max(28,...res.flatMap(r=>[...Object.keys(r.a.days),...Object.keys(r.b.days)].map(Number)));
  const per=S.map(s=>s.data.period).find(Boolean),dim=per?new Date(per.y,per.m,0).getDate():maxD;
  $('sum').innerHTML=`<div class="chips">
   <div class="chip">Сопоставлено<b>${res.length}</b></div>
   <div class="chip ${withD.length?'bad':'good'}">С отличиями<b>${withD.length}</b></div>
   <div class="chip good">Совпадают полностью<b>${res.length-withD.length}</b></div>
   <div class="chip ${onlyA.length?'bad':''}">Только в «${la}»<b>${onlyA.length}</b></div>
   <div class="chip ${onlyB.length?'bad':''}">Только в «${lb}»<b>${onlyB.length}</b></div></div>`;
  const q=norm($('q').value).trim(),od=$('onlyDiff').checked;
  const rows=res.filter(r=>(!od||r.d.length)&&(!q||norm(r.a.name+' '+r.b.name).includes(q)));
  const val=v=>v.c||(v.h?fmt(v.h):'');
  let h=`<div class="head"><div>Сотрудник</div><div>Должность</div><div class="n">Часы «${la}»</div><div class="n">Часы «${lb}»</div><div class="n">Разница</div><div class="n">Дней с отличиями</div></div>`;
  for(const r of rows){
    const dl=r.ta-r.tb,days=Array.from({length:dim},(_,i)=>i+1);
    h+=`<details class="emp"><summary><div><b>${esc(r.a.name)}</b>${r.approx?' <span class="warn">≈ по фамилии: '+esc(r.b.name)+'</span>':''}</div><div class="pos">${esc(r.a.pos)}</div>
    <div class="n">${fmt(r.ta)}</div><div class="n">${fmt(r.tb)}</div><div class="n ${dl?'d':'z'}">${dl>0?'+':''}${fmt(dl)}</div><div class="n ${r.d.length?'d':'z'}">${r.d.length||'—'}</div></summary>
    <div class="grid"><table><tr><th class="l">День</th>${days.map(d=>`<th>${d}</th>`).join('')}</tr>
    ${[[la,r.a],[lb,r.b]].map(([l,e])=>`<tr><td class="l">${esc(l)}</td>${days.map(d=>`<td class="${r.d.includes(d)?'x':''}">${esc(val(e.days[d]||Z))}</td>`).join('')}</tr>`).join('')}</table></div></details>`;
  }
  if(!rows.length)h+='<div class="msg">'+(withD.length||!od?'Ничего не найдено.':'Отличий не обнаружено 🎉')+'</div>';
  const lst=(t,arr)=>arr.length?`<details class="emp only"><summary style="display:block"><b>${t} (${arr.length})</b></summary><div class="grid"><ul>${arr.map(e=>`<li>${esc(e.name)} <span class="pos">— ${esc(e.pos)}, ${fmt(total(e))} ч</span></li>`).join('')}</ul></div></details>`:'';
  $('out').innerHTML=h+'<p></p>'+lst(`Есть только в «${la}»`,onlyA)+lst(`Есть только в «${lb}»`,onlyB);
}
document.addEventListener('click',e=>{const b=e.target.closest('[data-up]');if(b)document.querySelector(`[data-f="${b.dataset.up}"]`).click()});
document.addEventListener('change',e=>{
  const t=e.target;
  if(t.dataset.f!=null&&t.files[0])load(+t.dataset.f,t.files[0]);
  if(t.dataset.sh!=null){const s=S[+t.dataset.sh];s.sheet=t.value;s.manual=true;analyze(s);refresh()}
});
['dragover','dragleave','drop'].forEach(ev=>document.addEventListener(ev,e=>{
  const z=e.target.closest&&e.target.closest('.zone');if(!z)return;e.preventDefault();
  z.classList.toggle('drag',ev==='dragover');
  if(ev==='drop'&&e.dataTransfer.files[0])load(+z.dataset.i,e.dataTransfer.files[0]);
}));
$('onlyDiff').onchange=$('otEq').onchange=$('q').oninput=()=>{if(S[0]&&S[1])render()};
refresh();
}
</script>
</body>
</html>
