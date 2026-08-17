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
 var ws=XLSX.utils.json_to_sheet(data);
 var wb=XLSX.utils.book_new();
 XLSX.utils.book_append_sheet(wb,ws,'الأصناف');
 XLSX.writeFile(wb,'items_export.xlsx');
 toast('تم التصدير ✔');
};
$('btnSaveSession').onclick=function(){
 try{
  localStorage.setItem('session',JSON.stringify({ITEMS:ITEMS,HEADERS:HEADERS,MAP:MAP,RAWROWS:RAWROWS}));
  toast('حُفظت الجلسة ✔');
 }catch(e){
  toast('الملف كبير على الحفظ المحلي',1);
 }
};
if(localStorage.getItem('session')){
 $('btnResume').hidden=false;
}
$('btnResume').onclick=function(){
 var s=JSON.parse(localStorage.getItem('session'));
 ITEMS=s.ITEMS;
 HEADERS=s.HEADERS;
 MAP=s.MAP;
 RAWROWS=s.RAWROWS;
 $('dropZone').hidden=true;
 showWork();
 renderAll();
 log('استُعيدت الجلسة السابقة','ok');
 toast('تمت المتابعة ✔');
};
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
