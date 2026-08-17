var WB=null;
var ITEMS=[];
var HEADERS=[];
var MAP={};
var RAWROWS=[];
var sheetName='';
var ALIASES=load('aliasMap',{});
var IMGS=load('imgMap',{});
var scanner=null;
var lastScan=0;
var vStream=null;
var curSig=null;
var camOpen=false;
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
   log('تنزيل مكتبة ناقصة من: '+u);
   return loadScript(u).then(function(){
    return !!window[name];
   }).catch(function(){
    log('فشل التنزيل من: '+u,'er');
    return false;
   });
  });
 });
 return p;
}
function bindUpload(){
 var pairs=[['fileInput','lblOpen'],['fileInput2','lblOpen2']];
 pairs.forEach(function(pair){
  var inp=$(pair[0]);
  var lbl=$(pair[1]);
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
  log('لم يصل أي ملف — إن تكرر ذلك افتح الصفحة في متصفح مباشرة','er');
  return;
 }
 log('✔ وصل الملف: '+f.name+' ('+fsize(f.size)+')','ok');
 handleFile(f);
}
function handleFile(f){
 if(!/\.(xlsx|xls|csv)$/i.test(f.name)){
  toast('⚠ صيغة غير مدعومة — اختر ملف Excel',1);
  log('امتداد مرفوض: '+f.name,'er');
  return;
 }
 setLoad(true,'جارٍ قراءة الملف…','الخطوة 1/3: تحميل '+f.name);
 var urls=['https://unpkg.com/xlsx@0.18.5/dist/xlsx.full.min.js','https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'];
 ensureLib('XLSX',urls).then(function(ok){
  if(!ok){
   setLoad(false);
   toast('✖ تعذر تحميل مكتبة Excel — تحقق من الإنترنت',1);
   log('فشل تحميل مكتبة XLSX من كل المصادر','er');
   return;
  }
  var rd=new FileReader();
  rd.onerror=function(){
   setLoad(false);
   log('فشل FileReader في قراءة الملف','er');
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
   var btns='';
   WB.SheetNames.forEach(function(s){
    btns=btns+'<button class="chip" data-s="'+s+'">📄 '+s+'</button>';
   });
   $('sheetPick').innerHTML=btns;
   $('sheetPick').querySelectorAll('button').forEach(function(b){
    b.onclick=function(){useSheet(b.getAttribute('data-s'))};
   });
   useSheet(WB.SheetNames[0]);
   $('dropZone').hidden=true;
   showWork();
   showFileStat(f);
   setLoad(false);
   window.scrollTo(0,0);
   log('✔ اكتمل التحميل — الأصناف: '+ITEMS.length,'ok');
   toast('✔ اكتمل التحميل: '+fmt(ITEMS.length)+' صنف جاهز');
   if(!ITEMS.length){
    toast('⚠ لم تُقرأ أصناف — صحّح عمود اسم المادة',1);
    log('لا أصناف — راجع ربط الأعمدة','er');
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
 h=h+WB.SheetNames.length+' ورقة • ';
 h=h+fmt(RAWROWS.length)+' صف • ';
 h=h+fmt(ITEMS.length)+' صنف • ';
 h=h+'الحالة: <b style="color:var(--ac2)">مكتمل ✔</b>';
 $('fileStat').innerHTML=h;
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
 log('ورقة '+name+': العناوين بالصف '+(hi+1)+' — '+RAWROWS.length+' صف بيانات');
 autoMap();
 buildItems();
 renderAll();
 $('fileMeta').textContent=' — ورقة '+name+' — '+RAWROWS.length+' صف — '+HEADERS.length+' عمود';
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
 var ids=['mapSec','dash','tools','btnExport','btnSaveSession'];
 ids.forEach(function(i){$(i).hidden=false});
}
function renderMapping(){
 var html='';
 Object.keys(FIELDS).forEach(function(f){
  var opts='<option value="-1">— غير مرتبط —</option>';
  HEADERS.forEach(function(h,i){
   var sel='';
   if(MAP[f]===i){sel=' selected'}
   opts=opts+'<option value="'+i+'"'+sel+'>'+h+'</option>';
  });
  html=html+'<label class="mapitem"><span>'+FIELDS[f]+'</span><select data-f="'+f+'">'+opts+'</select></label>';
 });
 $('mapGrid').innerHTML=html;
 $('mapGrid').querySelectorAll('select').forEach(function(s){
  s.onchange=function(e){
   MAP[e.target.getAttribute('data-f')]=Number(e.target.value);
   buildItems();
   renderAll();
  };
 });
}
function renderAll(){
 renderMapping();
 analyze();
 renderTable();
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
 $('alerts').innerHTML=al;
 var g={};
 ITEMS.forEach(function(i){g[i.group]=(g[i.group]||0)+1});
 var gh='';
 Object.keys(g).forEach(function(k){
  gh=gh+'<span class="chip">🗂 '+(k||'بدون مجموعة')+': '+g[k]+'</span>';
 });
 $('groups').innerHTML=gh;
 $('alerts').querySelectorAll('.chip').forEach(function(c){
  c.onclick=function(){
   var key=c.getAttribute('data-l');
   var L={neg:neg,zero:zero,nobar:noBar,noprice:noPrice,dups:dups}[key];
   listModal('قائمة الأصناف',L);
  };
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
function renderTable(){
 var h='<table><tr>';
 HEADERS.forEach(function(x){h=h+'<th>'+x+'</th>'});
 h=h+'</tr>';
 RAWROWS.slice(0,200).forEach(function(r){
  h=h+'<tr>';
  r.forEach(function(c){h=h+'<td>'+String(c)+'</td>'});
  h=h+'</tr>';
 });
 $('tableWrap').innerHTML=h+'</table>';
}
bindUpload();
