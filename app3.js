/* ====== إصلاح CSS للجداول والأيقونات ====== */
(function(){
 var st=document.createElement('style');
 st.textContent='table{background:#fff;color:#000}th,td{border:1px solid #999;color:#000}th{background:#e8e8e8;color:#000}'
 +'.item{flex-wrap:wrap}.item b{font-size:13px}.item small{font-size:11px}';
 document.head.appendChild(st);
})();

/* إخفاء العناصر غير المطلوبة من الواجهة الثابتة */
(function(){
 var lo=$('lblOpen'); if(lo){lo.style.display='none'}
 ['tTable','tVision'].forEach(function(t){
  var b=document.querySelector('.tab[data-t="'+t+'"]');
  if(b&&b.parentNode){b.parentNode.removeChild(b)}
  var p=$(t);
  if(p&&p.parentNode){p.parentNode.removeChild(p)}
 });
 var badges=document.querySelectorAll('header .badge');
 if(badges.length){badges[0].textContent='الإصدار 7.3'}
})();

/* تحسين الكيبورد وإخفاؤه عند النقر خارج الحقول */
(function(){
 var mb=$('manualBar');
 if(mb){mb.setAttribute('inputmode','numeric')}
 document.addEventListener('click',function(e){
  var t=e.target;
  if(t.tagName!=='INPUT'&&t.tagName!=='TEXTAREA'&&t.tagName!=='SELECT'){
   var a=document.activeElement;
   if(a&&(a.tagName==='INPUT'||a.tagName==='TEXTAREA')){a.blur()}
  }
 },true);
})();

/* ====== نظام التصدير والإصلاح الجذري للترتيب ====== */
function exportRowsFile(title,rows){
 var ws=XLSX.utils.json_to_sheet(rows);
 var wb=XLSX.utils.book_new();
 wb.Workbook={Views:[{RTL:true}]};
 XLSX.utils.book_append_sheet(wb,ws,'القائمة');
 XLSX.writeFile(wb,title+'.xlsx');
 toast('تم التصدير ✔');
}

window.exportItemsList=function(){
 if(window.CURLIST&&window.CURLIST.rows&&window.CURLIST.rows.length){
  exportRowsFile(window.CURLIST.title,window.CURLIST.rows);
 }else{
  toast('لا توجد بيانات للتصدير',1);
 }
};

/* إعدادات الأعمدة الافتراضية */
var DEFAULT_EXPC=[
 {k:'seq',on:1,type:'std'},
 {k:'code',on:1,type:'std'},
 {k:'name',on:1,type:'std'},
 {k:'unit',on:1,type:'std'},
 {k:'price',on:1,type:'std'},
 {k:'group',on:1,type:'std'},
 {k:'qty',on:1,type:'std'},
 {k:'barcode',on:0,type:'std'},
 {k:'aliases',on:0,type:'std'}
];

/* تحميل الإعدادات مع ضمان عدم استخدام كاش قديم خاطئ */
var EXPC=load('expConf',null);
if(!EXPC || !Array.isArray(EXPC) || EXPC.length===0){
 EXPC=JSON.parse(JSON.stringify(DEFAULT_EXPC));
 saveLS('expConf',EXPC);
}

var COLT={seq:'تسلسل',code:'الرمز',name:'اسم المادة',unit:'الوحدة',price:'سعر المبيع',group:'المجموعة',qty:'الكمية',barcode:'الباركود',aliases:'أسماء بديلة'};

function itemRowByOrder(it,idx){
 var o={};
 EXPC.forEach(function(c){
  if(!c.on){return}
  if(c.k==='sep'){o[c.label||'']=''; return}
  if(c.k==='custom'){o[c.label||'']=''; return}
  
  var key=c.label||COLT[c.k]||c.k;
  if(c.k==='seq'){o[key]=idx+1}
  else if(c.k==='code'){o[key]=it.code}
  else if(c.k==='name'){o[key]=it.name}
  else if(c.k==='unit'){o[key]=it.unit}
  else if(c.k==='price'){o[key]=it.price}
  else if(c.k==='group'){o[key]=it.group}
  else if(c.k==='qty'){o[key]=it.qty}
  else if(c.k==='barcode'){o[key]=it.barcode}
  else if(c.k==='aliases'){o[key]=(it.aliases||[]).join('، ')}
 });
 return o;
}

function doExport(){
 if(!ITEMS.length){
  toast('لا توجد بيانات — ارفع ملفًا أولًا',1);
  return;
 }
 var rows=ITEMS.map(function(it,idx){return itemRowByOrder(it,idx)});
 exportRowsFile('الأصناف',rows);
}

function showExportModal(){
 if(CUR_SCREEN!=='exp'){MSTACK.push(CURRENT_REOPEN)}
 CURRENT_REOPEN=function(){showExportModal()};
 CUR_SCREEN='exp';
 lockScroll(true);
 $('mTitle').textContent='⚙ ترتيب أعمدة التصدير (يمين→يسار)';
 
 var h='<div class="list">';
 EXPC.forEach(function(c,i){
  var isSep=(c.k==='sep');
  var isCustom=(c.k==='custom');
  var label=isSep?'--- فاصل ---':(isCustom?(c.label||'اسم مخصص'):(COLT[c.k]||c.k));
  
  h+='<div class="item" style="gap:6px">';
  
  /* زر الإظهار/الإخفاء */
  h+='<button class="btn ghost" data-tg="'+i+'" style="min-width:40px">'+(c.on?'✔':'')+'</button>';
  
  /* حقل الاسم المخصص أو الفاصل */
  if(isSep){
   h+='<span style="flex:1;color:var(--mut);text-align:center;border-bottom:1px dashed var(--bd)">فاصل بصري</span>';
  }else if(isCustom){
   h+='<input type="text" class="inp" data-clbl="'+i+'" value="'+(c.label||'')+'" placeholder="اكتب العنوان هنا..." style="flex:1;font-size:12px;padding:6px">';
  }else{
   h+='<span style="flex:1;font-weight:700">'+label+'</span>';
  }
  
  /* أزرار التحكم */
  h+='<button class="btn ghost" data-up="'+i+'">⬆</button>';
  h+='<button class="btn ghost" data-dn="'+i+'">⬇</button>';
  h+='</div>';
 });
 
 h+='</div>';
 h+='<div class="row" style="margin-top:10px;gap:6px;flex-wrap:wrap">';
 h+='<button class="btn ghost" id="addSepBtn">➕ إضافة سطر فارغ</button>';
 h+='<button class="btn ghost" id="addCustBtn">➕ إضافة اسم مخصص</button>';
 h+='</div>';
 h+='<div class="row" style="margin-top:8px"><button class="btn" id="expSave">💾 حفظ الترتيب والتصدير</button></div>';
 
 $('mBody').innerHTML=h;
 $('modal').hidden=false;
 
 /* ربط الأحداث الجديدة */
 $('expSave').onclick=function(){
  /* تحديث الأسماء المخصصة قبل الحفظ */
  $('mBody').querySelectorAll('[data-clbl]').forEach(function(inp){
   var idx=Number(inp.getAttribute('data-clbl'));
   EXPC[idx].label=inp.value.trim();
  });
  saveLS('expConf',EXPC);
  doExport();
 };
 
 $('addSepBtn').onclick=function(){
  EXPC.push({k:'sep',on:1,label:''});
  showExportModal();
 };
 
 $('addCustBtn').onclick=function(){
  EXPC.push({k:'custom',on:1,label:'عنوان جديد'});
  showExportModal();
 };
 
 /* أزرار التحريك والتبديل */
 $('mBody').querySelectorAll('[data-up]').forEach(function(b){
  b.onclick=function(){
   var i=Number(b.getAttribute('data-up'));
   if(i>0){var t=EXPC[i-1];EXPC[i-1]=EXPC[i];EXPC[i]=t;showExportModal()}
  };
 });
 $('mBody').querySelectorAll('[data-dn]').forEach(function(b){
  b.onclick=function(){
   var i=Number(b.getAttribute('data-dn'));
   if(i<EXPC.length-1){var t=EXPC[i+1];EXPC[i+1]=EXPC[i];EXPC[i]=t;showExportModal()}
  };
 });
 $('mBody').querySelectorAll('[data-tg]').forEach(function(b){
  b.onclick=function(){
   var i=Number(b.getAttribute('data-tg'));
   EXPC[i].on=EXPC[i].on?0:1;
   showExportModal();
  };
 });
}

/* إضافة زرّي التصدير داخل بطاقة التحليل */
(function(){
 var dash=$('dash');
 if(!dash){return}
 var row=document.createElement('div');
 row.className='row';
 row.style.marginBottom='8px';
 row.innerHTML='<button class="btn" id="btnExpDirect">⬇️ تصدير Excel</button><button class="btn ghost" id="btnExpConf">⚙ ترتيب الأعمدة</button>';
 var stats=dash.querySelector('.stats');
 dash.insertBefore(row,stats);
 row.querySelector('#btnExpDirect').onclick=doExport;
 row.querySelector('#btnExpConf').onclick=showExportModal;
})();

/* ====== مدير الجلسات (نفس النسخة السابقة) ====== */
var SESSIONS=load('sessions',{});
var CURSID=null;
var CURFILE='';
window.markLoaded=function(fname){
 CURFILE=fname||CURFILE;
 CURSID=null;
 updateSessionBtn();
};
function updateSessionBtn(){
 var n=Object.keys(SESSIONS).length;
 var b=$('btnResume');
 if(n>0){
  b.hidden=false;
  b.textContent='⏳ الجلسات ('+n+')';
 }else{
  b.hidden=true;
 }
}
function saveCurrentSession(){
 if(!ITEMS.length){
  toast('لا توجد بيانات لحفظها — ارفع ملفًا أولًا',1);
  return;
 }
 var id=CURSID||String(Date.now());
 var name=(SESSIONS[id]&&SESSIONS[id].name)?SESSIONS[id].name:(CURFILE||('جلسة '+new Date().toLocaleString()));
 SESSIONS[id]={name:name,t:Date.now(),ITEMS:ITEMS,HEADERS:HEADERS,MAP:MAP,RAWROWS:RAWROWS};
 try{
  localStorage.setItem('sessions',JSON.stringify(SESSIONS));
  CURSID=id;
  updateSessionBtn();
  toast('💾 حُفظت الجلسة: '+name);
 }catch(e){
  delete SESSIONS[id];
  toast('مساحة التخزين ممتلئة — احذف جلسة قديمة أولًا',1);
 }
}
$('btnSaveSession').onclick=saveCurrentSession;
function restoreSession(id){
 var s=SESSIONS[id];
 if(!s){toast('الجلسة غير موجودة',1);return}
 ITEMS=s.ITEMS||[]; HEADERS=s.HEADERS||[]; MAP=s.MAP||{}; RAWROWS=s.RAWROWS||[];
 CURSID=id; CURFILE=s.name;
 $('dropZone').hidden=true; showWork(); renderAll();
 $('modal').hidden=true; MSTACK=[]; CURRENT_REOPEN=null; CUR_SCREEN=null; lockScroll(false);
 toast('✔ فُتحت الجلسة: '+s.name);
}
function showSessionsModal(){
 if(CUR_SCREEN!=='sess'){MSTACK.push(CURRENT_REOPEN)}
 CURRENT_REOPEN=function(){showSessionsModal()}; CUR_SCREEN='sess'; lockScroll(true);
 var ids=Object.keys(SESSIONS).sort(function(a,b){return (SESSIONS[b].t||0)-(SESSIONS[a].t||0)});
 $('mTitle').textContent='الجلسات المحفوظة ('+ids.length+')';
 var h='<div class="list">';
 if(!ids.length){h+='<p class="mut">لا جلسات محفوظة بعد.</p>'}
 ids.forEach(function(id){
  var s=SESSIONS[id]; var d=new Date(s.t||0);
  h+='<div class="item"><div style="flex:1;min-width:150px"><b>'+s.name+'</b><div class="mut small">'+(s.ITEMS?s.ITEMS.length:0)+' صنف • '+d.toLocaleString();
  if(id===CURSID){h+=' • <span class="sim">مفتوحة الآن</span>'}
  h+='</div></div><button class="btn" data-sopen="'+id+'">فتح</button><button class="btn ghost" data-sren="'+id+'">✏️</button><button class="btn ghost" data-sdel="'+id+'">🗑</button></div>';
 });
 h+='</div>';
 $('mBody').innerHTML=h; $('modal').hidden=false;
 $('mBody').querySelectorAll('[data-sopen]').forEach(function(b){b.onclick=function(){restoreSession(b.getAttribute('data-sopen'))}});
 $('mBody').querySelectorAll('[data-sren]').forEach(function(b){
  b.onclick=function(){
   var id=b.getAttribute('data-sren');
   var nn=prompt('التسمية الجديدة:',SESSIONS[id].name);
   if(nn&&nn.trim()){SESSIONS[id].name=nn.trim();saveLS('sessions',SESSIONS);showSessionsModal();toast('تمت إعادة التسمية ✔')}
  };
 });
 $('mBody').querySelectorAll('[data-sdel]').forEach(function(b){
  b.onclick=function(){
   var id=b.getAttribute('data-sdel');
   if(confirm('حذف "'+SESSIONS[id].name+'"؟')){
    delete SESSIONS[id]; if(CURSID===id)CURSID=null;
    saveLS('sessions',SESSIONS); updateSessionBtn(); showSessionsModal(); toast('حُذفت 🗑');
   }
  };
 });
}
$('btnResume').onclick=showSessionsModal;
updateSessionBtn();

/* ====== البيانات التجريبية وتهيئة النظام ====== */
$('btnDemo').onclick=function(){
 HEADERS=['#','الرمز','اسم المادة','الوحدة','سعر المبيع','بطاقة مجموعة','الكمية','الباركود'];
 RAWROWS=[
 [1,21002,'جام جارح الملون','قطعة',11000,'سوفيات',74.9,';31490134;'],
 [2,1814,'منشف بهجت 240م م','قطعة',5000,'شركة تابكو',83,';6945646512491;'],
 [5,187005,'ايك ضد البكتريا 2.5L مكس 1','قطعة',30000,'فلوري',2,';57992075;'],
 [11,187011,'ايك ضد البكتريا ابيض 2.5L','قطعة',30000,'فلوري',20,';14255943;'],
 [12,187012,'اجرة ستوتة','قطعة',10000,'خدمات',-7,';87198621;'],
 [20,187014,'اس بي ار 5 لتر','قطعة',20000,'فلوري',21,';81757;'],
 [27,1868,'اساس حديد رصاصي 0.75 L','قطعة',8000,'فلوري',47,';88065;'],
 [33,1756,'اسمير لون سيدار','قطعة',2000,'سوفيات',77,';22525;']];
 autoMap(); buildItems(); $('dropZone').hidden=true; showWork();
 $('fileStat').hidden=false;
 $('fileStat').innerHTML='✅ <b>بيانات تجريبية</b> • '+RAWROWS.length+' صف • '+ITEMS.length+' صنف • الحالة: <b style="color:var(--ac2)">مكتمل ✔</b>';
 renderAll(); toast('بيانات تجريبية ✔');
};
document.querySelectorAll('.tab').forEach(function(t){
 t.onclick=function(){
  document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on')});
  document.querySelectorAll('.tabp').forEach(function(x){x.classList.remove('on')});
  t.classList.add('on'); var p=$(t.getAttribute('data-t')); if(p)p.classList.add('on');
 };
});
['dragover','drop'].forEach(function(ev){document.addEventListener(ev,function(e){e.preventDefault()})});
document.addEventListener('drop',function(e){var f=e.dataTransfer.files[0];if(f)handleFile(f)});
(function(){
 var missing=[];
 ['handleFile','autoMap','analyze','showMapModal','showGroupModal'].forEach(function(f){
  if(typeof window[f]!=='function'){missing.push('app1.js')}
 });
 ['openModal','search','listModal','findByBarcode'].forEach(function(f){
  if(typeof window[f]!=='function'){missing.push('app2.js')}
 });
 var b=$('sysState');
 if(missing.length){
  b.textContent='⚠ ملف تالف: '+missing.join('، ');
  b.style.color='var(--bad)';
  b.style.borderColor='var(--bad)';
 }else{
  window.APP_READY=true;
  b.textContent='✔ جاهز 7.3';
  b.style.color='var(--ac2)';
  b.style.borderColor='var(--ac2)';
  log('النظام اكتمل تشغيله','ok');
 }
})();
