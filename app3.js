(function(){
 var st=document.createElement('style');
 st.textContent='table{background:#fff;color:#000}th,td{border:1px solid #999;color:#000}th{background:#e8e8e8;color:#000}';
 document.head.appendChild(st);
})();
function exportRowsFile(title,rows){
 var ws=XLSX.utils.json_to_sheet(rows);
 var wb=XLSX.utils.book_new();
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
$('btnExport').onclick=function(){
 var data=ITEMS.map(function(it){
  return {
   'الرمز':it.code,
   'اسم المادة':it.name,
   'الباركود':it.barcode,
   'سعر المبيع':it.price,
   'الكمية':it.qty,
   'الوحدة':it.unit,
   'المجموعة':it.group,
   'أسماء بديلة':(it.aliases||[]).join('، ')
  };
 });
 exportRowsFile('الأصناف',data);
};
/* ====== مدير الجلسات (بدون مدة انتهاء) ====== */
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
 if(!s){
  toast('الجلسة غير موجودة',1);
  return;
 }
 ITEMS=s.ITEMS||[];
 HEADERS=s.HEADERS||[];
 MAP=s.MAP||{};
 RAWROWS=s.RAWROWS||[];
 CURSID=id;
 CURFILE=s.name;
 $('dropZone').hidden=true;
 showWork();
 renderAll();
 $('modal').hidden=true;
 toast('✔ فُتحت الجلسة: '+s.name);
}
function showSessionsModal(){
 var ids=Object.keys(SESSIONS).sort(function(a,b){
  return (SESSIONS[b].t||0)-(SESSIONS[a].t||0);
 });
 $('mTitle').textContent='الجلسات المحفوظة ('+ids.length+')';
 var h='<div class="row" style="margin-bottom:8px"><button class="btn" id="btnSaveNow">💾 حفظ الجلسة الحالية</button></div>';
 h=h+'<div class="list">';
 if(!ids.length){
  h=h+'<p class="mut">لا جلسات محفوظة بعد — ارفع ملفًا ثم اضغط "حفظ الجلسة".</p>';
 }
 ids.forEach(function(id){
  var s=SESSIONS[id];
  var d=new Date(s.t||0);
  h=h+'<div class="item" style="flex-wrap:wrap">';
  h=h+'<div style="flex:1;min-width:150px"><b>'+s.name+'</b>';
  h=h+'<div class="mut small">'+(s.ITEMS?s.ITEMS.length:0)+' صنف • '+d.toLocaleString();
  if(id===CURSID){h=h+' • <span class="sim">مفتوحة الآن</span>'}
  h=h+'</div></div>';
  h=h+'<button class="btn" data-sopen="'+id+'">فتح</button>';
  h=h+'<button class="btn ghost" data-sren="'+id+'">✏️</button>';
  h=h+'<button class="btn ghost" data-sdel="'+id+'">🗑</button>';
  h=h+'</div>';
 });
 h=h+'</div>';
 $('mBody').innerHTML=h;
 $('modal').hidden=false;
 $('btnSaveNow').onclick=function(){
  saveCurrentSession();
  showSessionsModal();
 };
 $('mBody').querySelectorAll('[data-sopen]').forEach(function(b){
  b.onclick=function(){restoreSession(b.getAttribute('data-sopen'))};
 });
 $('mBody').querySelectorAll('[data-sren]').forEach(function(b){
  b.onclick=function(){
   var id=b.getAttribute('data-sren');
   var nn=prompt('التسمية الجديدة للجلسة:',SESSIONS[id].name);
   if(nn&&nn.trim()){
    SESSIONS[id].name=nn.trim();
    saveLS('sessions',SESSIONS);
    showSessionsModal();
    toast('تمت إعادة التسمية ✔');
   }
  };
 });
 $('mBody').querySelectorAll('[data-sdel]').forEach(function(b){
  b.onclick=function(){
   var id=b.getAttribute('data-sdel');
   if(confirm('حذف الجلسة "'+SESSIONS[id].name+'" نهائيًا؟')){
    delete SESSIONS[id];
    if(CURSID===id){CURSID=null}
    saveLS('sessions',SESSIONS);
    updateSessionBtn();
    showSessionsModal();
    toast('حُذفت الجلسة 🗑');
   }
  };
 });
}
$('btnResume').onclick=showSessionsModal;
updateSessionBtn();
/* ====== بقية النظام ====== */
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
 autoMap();
 buildItems();
 $('dropZone').hidden=true;
 showWork();
 $('fileStat').hidden=false;
 $('fileStat').innerHTML='✅ <b>بيانات تجريبية</b> • '+RAWROWS.length+' صف • '+ITEMS.length+' صنف • الحالة: <b style="color:var(--ac2)">مكتمل ✔</b>';
 renderAll();
 toast('بيانات تجريبية ✔');
};
document.querySelectorAll('.tab').forEach(function(t){
 t.onclick=function(){
  document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on')});
  document.querySelectorAll('.tabp').forEach(function(x){x.classList.remove('on')});
  t.classList.add('on');
  $(t.getAttribute('data-t')).classList.add('on');
 };
});
['dragover','drop'].forEach(function(ev){
 document.addEventListener(ev,function(e){e.preventDefault()});
});
document.addEventListener('drop',function(e){
 var f=e.dataTransfer.files[0];
 if(f){handleFile(f)}
});
(function(){
 var missing=[];
 ['handleFile','autoMap','analyze','renderAll'].forEach(function(f){
  if(typeof window[f]!=='function'){missing.push('app1.js')}
 });
 ['openModal','search','capture','findByBarcode'].forEach(function(f){
  if(typeof window[f]!=='function'){missing.push('app2.js')}
 });
 var b=$('sysState');
 if(missing.length){
  b.textContent='⚠ ملف تالف/ناقص: '+missing.join('، ')+' — أعد لصقه';
  b.style.color='var(--bad)';
  b.style.borderColor='var(--bad)';
 }else{
  window.APP_READY=true;
  b.textContent='✔ جاهز';
  b.style.color='var(--ac2)';
  b.style.borderColor='var(--ac2)';
  log('النظام اكتمل تشغيله وجاهز للعمل','ok');
 }
})();
