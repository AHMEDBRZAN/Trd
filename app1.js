var WB=null;
var ITEMS=[];
var HEADERS=[];
var MAP={};
var RAWROWS=[];
var sheetName='';
var SIM=[];
var MSTACK=[];
var CURRENT_REOPEN=null;
var CUR_SCREEN=null;
var ALIASES=load('aliasMap',{});
var IMGS=load('imgMap',{});
var scanner=null;
var lastScan=0;
function $(id){return document.getElementById(id)}
var FIELDS={name:'اسم المادة',code:'الرمز',barcode:'الباركود',price:'سعر المبيع',qty:'الكمية',unit:'الوحدة',group:'المجموعة/بطاقة'};
function load(k,d){
 try{
  var v=JSON.parse(localStorage.getItem(k));
  if(v==null){return d}
  return v;
 }catch(e){return d}
}
function saveLS(k,v){
 try{localStorage.setItem(k,JSON.stringify(v))}catch(e){toast('تعذر الحفظ المحلي',1)}
}
function normAr(s){
 var t=String(s==null?'':s);
 t=t.replace(/[\u064B-\u0652\u0640]/g,'');
 t=t.replace(/[أإآ]/g,'ا');
 t=t.replace(/ة/g,'ه');
 t=t.replace(/[ىئ]/g,'ي');
 t=t.replace(/ؤ/g,'و');
 t=t.toLowerCase();
 t=t.replace(/\s+/g,' ').trim();
 return t;
}
function digits(s){return String(s==null?'':s).replace(/\D/g,'')}
function num(s){
 var n=parseFloat(String(s==null?'':s).replace(/[^\d.-]/g,''));
 if(isFinite(n)){return n}
 return 0;
}
function fmt(n){return Math.round(n).toLocaleString('en-US')}
function fsize(n){
 if(n>1048576){return (n/1048576).toFixed(1)+' MB'}
 return Math.max(1,Math.round(n/1024))+' KB';
}
function nextFrame(){return new Promise(function(r){setTimeout(r,30)})}
function toast(m,err){
 var t=$('toast');
 t.textContent=m;
 t.classList.add('show');
 if(err){t.classList.add('err')}
 setTimeout(function(){t.classList.remove('show');t.classList.remove('err')},2600);
}
function setLoad(on,txt,step){
 $('loader').hidden=!on;
 if(txt){$('loadTxt').textContent=txt}
 if(step!==undefined){$('loadStep').textContent=step}
}
function log(m,cls){
 $('log').hidden=false;
 var li=document.createElement('li');
 if(cls){li.className=cls}
 li.textContent='• '+m;
 $('logList').appendChild(li);
}
function lockScroll(v){
 document.body.style.overflow=v?'hidden':'';
}
function beep(){
 try{
  var AC=window.AudioContext||window.webkitAudioContext;
  var a=new AC();
  var o=a.createOscillator();
  var g=a.createGain();
  o.connect(g);
  g.connect(a.destination);
  o.frequency.value=880;
  g.gain.value=0.08;
  o.start();
  setTimeout(function(){o.stop();a.close()},120);
 }catch(e){}
}
function loadScript(src){
 return new Promise(function(res,rej){
  var s=document.createElement('script');
  s.src=src;
  s.onload=res;
  s.onerror=rej;
  document.head.appendChild(s);
 });
}
function ensureLib(name,urls){
 if(window[name]){return Promise.resolve(true)}
 var p=Promise.resolve(false);
 urls.forEach(function(u){
  p=p.then(function(ok){
   if(ok){return true}
   return loadScript(u).then(function(){
    return !!window[name];
   }).catch(function(){return false});
  });
 });
 return p;
}
function bindUpload(){
 var pairs=[['fileInput2','lblOpen2']];
 pairs.forEach(function(pair){
  var inp=$(pair[0]);
  var lbl=$(pair[1]);
  if(!inp||!lbl){return}
  lbl.addEventListener('click',function(){
   log('ضغطة زر اختيار الملف سُجّلت — يُفتح المنتقي…');
  });
  inp.addEventListener('change',onFile);
 });
}
function onFile(e){
 var f=e.target.files&&e.target.files[0];
 e.target.value='';
 if(!f){
  toast('⚠ المنتقي أُغلق بدون ملف',1);
  return;
 }
 log('✔ وصل الملف: '+f.name+' ('+fsize(f.size)+')','ok');
 handleFile(f);
}
function handleFile(f){
 if(!/\.(xlsx|xls|csv)$/i.test(f.name)){
  toast('⚠ صيغة غير مدعومة — اختر ملف Excel',1);
  return;
 }
 setLoad(true,'جارٍ قراءة الملف…','الخطوة 1/3: تحميل '+f.name);
 var urls=['https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js','https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];
 ensureLib('XLSX',urls).then(function(ok){
  if(!ok){
   setLoad(false);
   toast('✖ تعذر تحميل مكتبة Excel — تحقق من الإنترنت',1);
   return;
  }
  var rd=new FileReader();
  rd.onerror=function(){
   setLoad(false);
   toast('✖ تعذرت قراءة الملف',1);
  };
  rd.onload=function(){processBuf(rd.result,f)};
  rd.readAsArrayBuffer(f);
 });
}
function processBuf(buf,f){
 setLoad(true,'جارٍ فك ضغط الملف…','الخطوة 2/3: قراءة الأوراق');
 nextFrame().then(function(){
  WB=XLSX.read(buf,{type:'array'});
  log('الأوراق المكتشفة: '+WB.SheetNames.join('، '),'ok');
  if(!WB.SheetNames.length){
   setLoad(false);
   toast('✖ الملف بلا أوراق',1);
   return;
  }
  setLoad(true,'جارٍ فهم الأعمدة…','الخطوة 3/3: بناء الفهارس');
  return nextFrame().then(function(){
   useSheet(WB.SheetNames[0]);
   $('dropZone').hidden=true;
   showWork();
   showFileStat(f);
   setLoad(false);
   window.scrollTo(0,0);
   log('✔ اكتمل التحميل — الأصناف: '+ITEMS.length,'ok');
   toast('✔ اكتمل التحميل: '+fmt(ITEMS.length)+' صنف جاهز');
   if(window.markLoaded){window.markLoaded(f.name)}
   if(!ITEMS.length){
    toast('⚠ لم تُقرأ أصناف — افتح 🔎 الفهم الآلي وصحّح الأعمدة',1);
   }
  });
 }).catch(function(err){
  setLoad(false);
  log('خطأ أثناء المعالجة: '+err.message,'er');
  toast('✖ ملف تالف أو غير مدعوم',1);
 });
}
function showFileStat(f){
 $('fileStat').hidden=false;
 var h='✅ <b>'+f.name+'</b> ';
 h=h+'<span class="mut">('+fsize(f.size)+')</span> • ';
 h=h+fmt(ITEMS.length)+' صنف • ';
 h=h+'<b style="color:var(--ac2)">مكتمل ✔</b> ';
 h=h+'<label class="btn ghost" style="cursor:pointer">📂 رفع ملف آخر<input type="file" id="fileInput3" accept=".xlsx,.xls,.csv" class="hid"></label> ';
 h=h+'<button class="btn ghost" id="btnMap">🔎 الفهم الآلي</button>';
 $('fileStat').innerHTML=h;
 var fi3=$('fileInput3');
 if(fi3){fi3.addEventListener('change',onFile)}
 var bm=$('btnMap');
 if(bm){bm.onclick=showMapModal}
}
/* ====== الفهم الآلي داخل نافذة ====== */
function showMapModal(){
 if(CUR_SCREEN!=='map'){MSTACK.push(CURRENT_REOPEN)}
 CURRENT_REOPEN=function(){showMapModal()};
 CUR_SCREEN='map';
 lockScroll(true);
 $('mTitle').textContent='🔎 الفهم الآلي للملف';
 var h='<div class="chips" id="mSheetPick">';
 if(WB){
  WB.SheetNames.forEach(function(s){
   h=h+'<span class="chip'+(s===sheetName?' warn':'')+'" data-sh="'+s+'">📄 '+s+'</span>';
  });
 }
 h=h+'</div><div class="grid map" id="mMapGrid">';
 Object.keys(FIELDS).forEach(function(fkey){
  var opts='<option value="-1">— غير مرتبط —</option>';
  HEADERS.forEach(function(hh,i){
   var sel='';
   if(MAP[fkey]===i){sel=' selected'}
   opts=opts+'<option value="'+i+'"'+sel+'>'+hh+'</option>';
  });
  h=h+'<label class="mapitem"><span>'+FIELDS[fkey]+'</span><select data-f="'+fkey+'">'+opts+'</select></label>';
 });
 h=h+'</div><div class="row" style="margin-top:10px"><button class="btn" id="mMapSave">💾 حفظ وتطبيق</button></div>';
 $('mBody').innerHTML=h;
 $('modal').hidden=false;
 $('mBody').querySelectorAll('[data-sh]').forEach(function(c){
  c.onclick=function(){
   useSheet(c.getAttribute('data-sh'));
   showMapModal();
  };
 });
 $('mMapSave').onclick=function(){
  $('mMapGrid').querySelectorAll('select').forEach(function(s){
   MAP[s.getAttribute('data-f')]=Number(s.value);
  });
  buildItems();
  renderAll();
  MSTACK.pop();
  CURRENT_REOPEN=null;
  CUR_SCREEN=null;
  $('modal').hidden=true;
  lockScroll(false);
  toast('✔ طُبّق الفهم الآلي');
 };
}
function useSheet(name){
 sheetName=name;
 var rows=XLSX.utils.sheet_to_json(WB.Sheets[name],{header:1,defval:''});
 if(!rows.length){
  toast('الورقة '+name+' فارغة',1);
  return;
 }
 var hi=-1;
 var max=Math.min(rows.length,25);
 for(var i=0;i<max;i=i+1){
  var line=rows[i].map(function(c){return String(c)}).join('|');
  if(/اسم الماد[ةه]|باركود|سعر المبيع|الرمز/.test(line)){
   hi=i;
   break;
  }
 }
 if(hi<0){hi=0}
 HEADERS=rows[hi].map(function(c,idx){
  var t=String(c).trim();
  if(t){return t}
  return 'عمود '+(idx+1);
 });
 RAWROWS=rows.slice(hi+1);
 autoMap();
 buildItems();
 renderAll();
}
function autoMap(){
 MAP={name:-1,code:-1,barcode:-1,price:-1,qty:-1,unit:-1,group:-1};
 var used={};
 function find(t){
  for(var i=0;i<HEADERS.length;i=i+1){
   if(used[i]){continue}
   var h=normAr(HEADERS[i]);
   if(t(h)){
    used[i]=1;
    return i;
   }
  }
  return -1;
 }
 MAP.name=find(function(h){return h.indexOf('اسم الماده')>=0||h.indexOf('اسم المادة')>=0});
 if(MAP.name<0){
  MAP.name=find(function(h){return h.indexOf('اسم')>=0});
 }
 MAP.barcode=find(function(h){return h.indexOf('باركود')>=0||h.indexOf('barcode')>=0});
 MAP.code=find(function(h){return h.indexOf('الرمز')>=0||h.indexOf('رمز')>=0||h.indexOf('كود')>=0});
 MAP.price=find(function(h){return h.indexOf('سعر')>=0});
 MAP.qty=find(function(h){return h.indexOf('كميه')>=0||h.indexOf('qty')>=0});
 MAP.unit=find(function(h){return h.indexOf('وحده')>=0});
 MAP.group=find(function(h){return h.indexOf('بطاقه')>=0||h.indexOf('مجموعه')>=0||h.indexOf('فئه')>=0||h.indexOf('تصنيف')>=0});
 log('ربط الأعمدة: اسم='+(MAP.name+1)+' رمز='+(MAP.code+1)+' باركود='+(MAP.barcode+1)+' سعر='+(MAP.price+1)+' كمية='+(MAP.qty+1));
}
function buildItems(){
 ITEMS=[];
 RAWROWS.forEach(function(r){
  function g(i){
   if(i>=0){return r[i]}
   return '';
  }
  var name=String(g(MAP.name)).trim();
  if(!name){return}
  var code=String(g(MAP.code)).trim();
  ITEMS.push({
   name:name,
   code:code,
   barcode:String(g(MAP.barcode)).trim(),
   price:num(g(MAP.price)),
   qty:num(g(MAP.qty)),
   unit:String(g(MAP.unit)).trim(),
   group:String(g(MAP.group)).trim(),
   aliases:ALIASES[code]||[]
  });
 });
}
function showWork(){
 var ids=['dash','tools','btnSaveSession'];
 ids.forEach(function(i){$(i).hidden=false});
}
/* ====== التشابه مع فهم المقاسات ====== */
var UNIT_WORDS={'انج':1,'انش':1,'بوصه':1,'بوصة':1,'لتر':1,'ل':1,'مل':1,'كيلو':1,'كغم':1,'غم':1,'سم':1,'مم':1,'متر':1,'l':1,'kg':1,'g':1,'ltr':1};
function sizeSig(s){
 var toks=s.split(' ');
 var sig=[];
 for(var i=0;i<toks.length;i=i+1){
  var t=toks[i];
  if(/^[0-9٠-٩]+([.,][0-9٠-٩]+)?$/.test(t)){
   var u=toks[i+1]||'';
   if(UNIT_WORDS[u]){sig.push(t+u)}else{sig.push(t)}
  }
 }
 return sig.join('|');
}
function stripSize(s){
 return s.split(' ').filter(function(t){
  if(/^[0-9٠-٩]+([.,][0-9٠-٩]+)?$/.test(t)){return false}
  if(UNIT_WORDS[t]){return false}
  if(/^[0-9٠-٩]+([.,][0-9٠-٩]+)?[a-zأ-ي]{1,3}$/.test(t)){return false}
  return t.length>0;
 }).join(' ');
}
function simNames(a,b){
 if(a===b){return 1}
 if(sizeSig(a)!==sizeSig(b)){return 0}
 if(stripSize(a)===stripSize(b)){return 0}
 var sa={};
 var sb={};
 a.split(' ').forEach(function(t){if(t.length>1){sa[t]=1}});
 b.split(' ').forEach(function(t){if(t.length>1){sb[t]=1}});
 var inter=0;
 Object.keys(sa).forEach(function(t){
  if(sb[t]){inter=inter+1}
 });
 var na=Object.keys(sa).length;
 var nb=Object.keys(sb).length;
 if(na+nb===0){return 0}
 return 2*inter/(na+nb);
}
function computeSimOn(arr){
 var pairs=[];
 var nn=arr.map(function(it){return normAr(it.name)});
 var byFirst={};
 nn.forEach(function(s,idx){
  var f=s.split(' ')[0]||s;
  byFirst[f]=byFirst[f]||[];
  byFirst[f].push(idx);
 });
 Object.keys(byFirst).forEach(function(k){
  var a2=byFirst[k];
  for(var i=0;i<a2.length;i=i+1){
   for(var j=i+1;j<a2.length;j=j+1){
    if(pairs.length>=200){return}
    var s=simNames(nn[a2[i]],nn[a2[j]]);
    if(s>=0.6){pairs.push({a:a2[i],b:a2[j],s:s})}
   }
  }
 });
 pairs.sort(function(x,y){return y.s-x.s});
 return pairs;
}
function computeSim(){SIM=computeSimOn(ITEMS)}
function openSimList(arr,pairs,title){
 if(CUR_SCREEN!=='sim'){MSTACK.push(CURRENT_REOPEN)}
 CURRENT_REOPEN=function(){openSimList(arr,pairs,title)};
 CUR_SCREEN='sim';
 lockScroll(true);
 $('mTitle').textContent=title+' ('+pairs.length+')';
 var rows=pairs.map(function(p){
  var a=arr[p.a];
  var b=arr[p.b];
  return {
   'المادة الأولى':a.name,'رمز الأولى':a.code,'سعر الأولى':a.price,'كمية الأولى':a.qty,
   'المادة الثانية':b.name,'رمز الثانية':b.code,'سعر الثانية':b.price,'كمية الثانية':b.qty,
   'نسبة التشابه %':Math.round(p.s*100)
  };
 });
 window.CURLIST={title:title,rows:rows};
 var h='<div class="row" style="margin-bottom:8px"><button class="btn ghost" id="btnExpList">⬇️ تصدير Excel</button></div>';
 h=h+'<div class="list">';
 if(!pairs.length){
  h=h+'<p class="mut">لا توجد أصناف متشابهة ✔</p>';
 }
 pairs.forEach(function(p){
  var a=arr[p.a];
  var b=arr[p.b];
  h=h+'<div class="item" style="flex-direction:column;align-items:stretch;gap:8px">';
  h=h+'<span class="sim small">👥 نسبة التشابه: '+Math.round(p.s*100)+'%</span>';
  h=h+'<div class="row" style="justify-content:space-between"><div><b>'+a.name+'</b><div class="mut small">السعر: '+fmt(a.price)+' • الكمية: '+a.qty+' • رمز: '+a.code+'</div></div><button class="btn ghost" data-open="'+a.code+'">عرض</button></div>';
  h=h+'<div class="row" style="justify-content:space-between"><div><b>'+b.name+'</b><div class="mut small">السعر: '+fmt(b.price)+' • الكمية: '+b.qty+' • رمز: '+b.code+'</div></div><button class="btn ghost" data-open="'+b.code+'">عرض</button></div>';
  h=h+'</div>';
 });
 h=h+'</div>';
 $('mBody').innerHTML=h;
 $('modal').hidden=false;
 $('mBody').querySelectorAll('[data-open]').forEach(function(btn){
  btn.onclick=function(){openModal(btn.getAttribute('data-open'))};
 });
 var be=$('btnExpList');
 if(be){be.onclick=function(){if(window.exportItemsList){window.exportItemsList()}};}
}
/* ====== نافذة البطاقة (المجموعة) ====== */
function showGroupModal(g){
 if(CUR_SCREEN!=='group'){MSTACK.push(CURRENT_REOPEN)}
 CURRENT_REOPEN=function(){showGroupModal(g)};
 CUR_SCREEN='group';
 lockScroll(true);
 var sub=ITEMS.filter(function(i){return i.group===g});
 var neg=sub.filter(function(i){return i.qty<0});
 var zero=sub.filter(function(i){return i.qty===0});
 var noBar=sub.filter(function(i){return digits(i.barcode).length<4});
 var noPrice=sub.filter(function(i){return !i.price});
 var pairs=computeSimOn(sub);
 $('mTitle').textContent='🗂 بطاقة: '+g+' ('+sub.length+')';
 var h='<div class="chips">';
 h=h+'<span class="chip" data-g="all">كل المواد: '+sub.length+'</span>';
 h=h+'<span class="chip bad" data-g="neg">سالبة: '+neg.length+'</span>';
 h=h+'<span class="chip warn" data-g="zero">صفر: '+zero.length+'</span>';
 h=h+'<span class="chip warn" data-g="nobar">بدون باركود: '+noBar.length+'</span>';
 h=h+'<span class="chip warn" data-g="noprice">بدون سعر: '+noPrice.length+'</span>';
 h=h+'<span class="chip warn" data-g="sim">متشابهة: '+pairs.length+'</span>';
 h=h+'</div><div class="list">'+sub.map(itemCard).join('')+'</div>';
 $('mBody').innerHTML=h;
 $('modal').hidden=false;
 bindItems($('mBody'));
 $('mBody').querySelectorAll('[data-g]').forEach(function(c){
  c.onclick=function(){
   var key=c.getAttribute('data-g');
   if(key==='all'){listModal('مواد '+g,sub,'group');return}
   if(key==='neg'){listModal('سالبة '+g,neg,'neg');return}
   if(key==='zero'){listModal('صفر '+g,zero,'zero');return}
   if(key==='nobar'){listModal('بدون باركود '+g,noBar,'nobar');return}
   if(key==='noprice'){listModal('بدون سعر '+g,noPrice,'noprice');return}
   if(key==='sim'){openSimList(sub,pairs,'متشابهة '+g);return}
  };
 });
}
function renderAll(){
 analyze();
 $('results').innerHTML='';
}
function analyze(){
 var tot=ITEMS.length;
 var withBar=ITEMS.filter(function(i){return digits(i.barcode).length>=4});
 var neg=ITEMS.filter(function(i){return i.qty<0});
 var zero=ITEMS.filter(function(i){return i.qty===0});
 var noBar=ITEMS.filter(function(i){return digits(i.barcode).length<4});
 var noPrice=ITEMS.filter(function(i){return !i.price});
 var value=ITEMS.reduce(function(s,i){return s+(i.qty>0?i.qty*i.price:0)},0);
 var units=ITEMS.reduce(function(s,i){return s+(i.qty>0?i.qty:0)},0);
 var seen={};
 var dups=[];
 ITEMS.forEach(function(i){
  var d=digits(i.barcode);
  if(d.length>=4){
   seen[d]=(seen[d]||0)+1;
   if(seen[d]===2){dups.push(i)}
  }
 });
 computeSim();
 var st='';
 st=st+'<div class="stat"><b>'+fmt(tot)+'</b><span>إجمالي الأصناف</span></div>';
 st=st+'<div class="stat"><b>'+fmt(withBar.length)+'</b><span>بباركود</span></div>';
 st=st+'<div class="stat"><b>'+fmt(units)+'</b><span>القطع الموجبة</span></div>';
 st=st+'<div class="stat"><b>'+fmt(value)+'</b><span>قيمة المخزون</span></div>';
 $('stats').innerHTML=st;
 var al='';
 al=al+'<span class="chip bad" data-l="neg">⚠ كميات سالبة: '+neg.length+'</span>';
 al=al+'<span class="chip warn" data-l="zero">كميات صفر: '+zero.length+'</span>';
 al=al+'<span class="chip warn" data-l="nobar">بدون باركود: '+noBar.length+'</span>';
 al=al+'<span class="chip warn" data-l="noprice">بدون سعر: '+noPrice.length+'</span>';
 al=al+'<span class="chip bad" data-l="dups">باركود مكرر: '+dups.length+'</span>';
 al=al+'<span class="chip warn" data-l="sim">👥 أصناف متشابهة: '+SIM.length+'</span>';
 $('alerts').innerHTML=al;
 var g={};
 ITEMS.forEach(function(i){g[i.group]=(g[i.group]||0)+1});
 var gh='';
 Object.keys(g).forEach(function(k){
  gh=gh+'<span class="chip" data-gr="'+k+'">🗂 '+(k||'بدون مجموعة')+': '+g[k]+'</span>';
 });
 $('groups').innerHTML=gh;
 $('alerts').querySelectorAll('.chip').forEach(function(c){
  c.onclick=function(){
   var key=c.getAttribute('data-l');
   if(key==='sim'){openSimList(ITEMS,SIM,'الأصناف المتشابهة');return}
   var L={neg:neg,zero:zero,nobar:noBar,noprice:noPrice,dups:dups}[key];
   listModal('قائمة الأصناف',L,key);
  };
 });
 $('groups').querySelectorAll('[data-gr]').forEach(function(c){
  c.onclick=function(){showGroupModal(c.getAttribute('data-gr'))};
 });
 var rows='<table><tr><th>العمود</th><th>الدور</th><th>معبأ</th><th>أرقام</th><th>عينة</th></tr>';
 HEADERS.forEach(function(h,i){
  var f=0;
  var n=0;
  var smp='';
  RAWROWS.forEach(function(r){
   var v=String(r[i]).trim();
   if(v){
    f=f+1;
    if(!smp&&v.length<30){smp=v}
    if(isFinite(parseFloat(v))){n=n+1}
   }
  });
  var role=null;
  Object.keys(MAP).forEach(function(k){
   if(MAP[k]===i){role=k}
  });
  rows=rows+'<tr><td>'+h+'</td><td>'+(role?FIELDS[role]:'—')+'</td><td>'+f+'/'+RAWROWS.length+'</td><td>'+n+'</td><td>'+smp+'</td></tr>';
 });
 $('colInsight').innerHTML=rows+'</table>';
}
bindUpload();
