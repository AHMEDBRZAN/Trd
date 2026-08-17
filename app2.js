var CURLIST_ARR=[];
var CURLIST_TITLE='';
var CURLIST_KEY='';
var CURLIST_SORT='def';
function search(q){
 q=normAr(q);
 var qd=digits(q);
 var out=[];
 ITEMS.forEach(function(it){
  var s=0;
  var nn=normAr(it.name);
  if(qd.length>=4&&digits(it.barcode).indexOf(qd)>=0){s=s+100}
  if(qd&&digits(it.code)===qd){s=s+95}
  if(nn===q){s=s+80}
  else if(nn.indexOf(q)>=0){s=s+60}
  it.aliases.forEach(function(a){
   var an=normAr(a);
   if(an&&(an.indexOf(q)>=0||q.indexOf(an)>=0)){s=s+70}
  });
  if(!s&&q.length>2){
   var hit=0;
   q.split(' ').forEach(function(t){
    if(t.length>1&&nn.indexOf(t)>=0){hit=hit+1}
   });
   if(hit){s=20+hit*10}
  }
  if(s>0){out.push([s,it])}
 });
 out.sort(function(a,b){return b[0]-a[0]});
 return out.slice(0,50).map(function(x){return x[1]});
}
$('q').addEventListener('input',function(e){
 var r=search(e.target.value);
 $('results').innerHTML=e.target.value?(r.map(itemCard).join('')||'<p class="mut">لا نتائج</p>'):'';
 bindItems($('results'));
});
function itemCard(it){
 var al=it.aliases.length?' • أيضًا: '+it.aliases.join('، '):'';
 return '<div class="item" data-code="'+it.code+'"><div style="flex:1;min-width:120px"><b>'+it.name+'</b><small>'+(it.group||'')+al+'</small></div><span class="badge">💰 '+fmt(it.price)+'</span><span class="badge">🔢 '+it.qty+'</span></div>';
}
function bindItems(root){
 root.querySelectorAll('.item').forEach(function(el){
  el.onclick=function(){openModal(el.getAttribute('data-code'))};
 });
}
function findByBarcode(q){
 var d=digits(q);
 if(!d){return null}
 for(var i=0;i<ITEMS.length;i=i+1){
  var it=ITEMS[i];
  var bs=String(it.barcode).split(/[,;|\s]+/).map(digits).filter(function(x){return x.length>=4});
  if(digits(it.barcode)===d||bs.indexOf(d)>=0){return it}
  if(d.length>=6){
   var tail=false;
   bs.forEach(function(b){
    if(b.indexOf(d)===b.length-d.length){tail=true}
   });
   if(tail){return it}
  }
 }
 return null;
}
function showScan(it,q){
 $('scanResult').innerHTML=it?itemCard(it):'<div class="item"><div class="ph">❓</div><div><b>باركود غير موجود</b><small>'+q+'</small></div></div>';
 bindItems($('scanResult'));
}
$('manualBar').addEventListener('keydown',function(e){
 if(e.key==='Enter'){
  var it=findByBarcode(e.target.value);
  showScan(it,e.target.value);
  if(it){beep()}
 }
});
function stopScan(){
 if(scanner){
  scanner.stop().catch(function(){});
  scanner.clear();
  scanner=null;
 }
 $('qr-reader').hidden=true;
 $('btnScanStart').hidden=false;
 $('btnScanStop').hidden=true;
}
$('btnScanStop').onclick=stopScan;
$('btnScanStart').onclick=function(){
 $('qr-reader').hidden=false;
 $('btnScanStart').hidden=true;
 $('btnScanStop').hidden=false;
 ensureLib('Html5Qrcode',['https://cdn.jsdelivr.net/npm/html5-qrcode@2.3.8/html5-qrcode.min.js']).then(function(ok){
  if(!ok){throw new Error('no-lib')}
  scanner=new Html5Qrcode('qr-reader');
  return scanner.start({facingMode:'environment'},{fps:10,qrbox:{width:260,height:150}},function(txt){
   var now=Date.now();
   if(now-lastScan<1500){return}
   lastScan=now;
   beep();
   if(navigator.vibrate){navigator.vibrate(150)}
   showScan(findByBarcode(txt),txt);
  },function(){});
 }).catch(function(){
  toast('✖ تعذر تشغيل كاميرا الباركود',1);
  stopScan();
 });
};
/* ====== بطاقة المادة (بدون صور) ====== */
function openModal(code){
 var it=null;
 ITEMS.forEach(function(i){if(i.code===code){it=i}});
 if(!it){return}
 if(CUR_SCREEN!=='item'){MSTACK.push(CURRENT_REOPEN)}
 CURRENT_REOPEN=null;
 CUR_SCREEN='item';
 lockScroll(true);
 $('mTitle').textContent=it.name;
 var chips='<span class="chip">الرمز: '+it.code+'</span>';
 chips=chips+'<span class="chip">الباركود: '+(it.barcode||'—')+'</span>';
 chips=chips+'<span class="chip">السعر: '+fmt(it.price)+'</span>';
 chips=chips+'<span class="chip">الكمية: '+it.qty+'</span>';
 chips=chips+'<span class="chip">الوحدة: '+(it.unit||'—')+'</span>';
 chips=chips+'<span class="chip">المجموعة: '+(it.group||'—')+'</span>';
 var als='<span class="mut small">لا يوجد</span>';
 if(it.aliases.length){
  als=it.aliases.map(function(a,i){
   return '<span class="chip">'+a+' <i data-del="'+i+'" style="cursor:pointer">✖</i></span>';
  }).join('');
 }
 $('mBody').innerHTML='<div class="chips">'+chips+'</div><h4>🏷 الأسماء البديلة</h4><div class="chips" id="mAl">'+als+'</div><div class="row" style="margin-top:8px"><input id="mAlIn" class="inp grow" placeholder="أضف اسمًا شعبيًا…"/><button class="btn" id="mAlAdd">＋</button></div>';
 $('modal').hidden=false;
 $('mAlAdd').onclick=function(){
  var v=$('mAlIn').value.trim();
  if(!v){return}
  it.aliases.push(v);
  ALIASES[code]=it.aliases;
  saveLS('aliasMap',ALIASES);
  openModal(code);
  toast('أُضيف الاسم ✔');
 };
 $('mAl').querySelectorAll('[data-del]').forEach(function(x){
  x.onclick=function(){
   it.aliases.splice(Number(x.getAttribute('data-del')),1);
   ALIASES[code]=it.aliases;
   saveLS('aliasMap',ALIASES);
   openModal(code);
  };
 });
}
$('mClose').onclick=function(){
 var prev=MSTACK.pop();
 if(prev){
  CURRENT_REOPEN=prev;
  prev();
 }else{
  CURRENT_REOPEN=null;
  CUR_SCREEN=null;
  $('modal').hidden=true;
  lockScroll(false);
 }
};
/* ====== القوائم مع الفرز ====== */
function listModal(title,arr,key){
 if(CUR_SCREEN!=='list'){MSTACK.push(CURRENT_REOPEN)}
 CURLIST_TITLE=title;
 CURLIST_ARR=arr;
 CURLIST_KEY=key||'';
 CURLIST_SORT=(key==='neg')?'qtya':'def';
 CURRENT_REOPEN=function(){listModal(CURLIST_TITLE,CURLIST_ARR,CURLIST_KEY)};
 CUR_SCREEN='list';
 renderListModal();
}
function sortOptions(){
 var o=[['def','الترتيب الافتراضي'],['qtyd','الكمية: من الأكثر'],['qtya','الكمية: من الأقل'],['name','الاسم'],['code','الرمز']];
 if(CURLIST_KEY!=='noprice'){
  o.push(['priced','السعر: من الأعلى']);
  o.push(['pricea','السعر: من الأدنى']);
 }
 return o;
}
function renderListModal(){
 lockScroll(true);
 var arr=CURLIST_ARR.slice();
 if(CURLIST_SORT==='qtyd'){arr.sort(function(a,b){return b.qty-a.qty})}
 if(CURLIST_SORT==='qtya'){arr.sort(function(a,b){return a.qty-b.qty})}
 if(CURLIST_SORT==='priced'){arr.sort(function(a,b){return b.price-a.price})}
 if(CURLIST_SORT==='pricea'){arr.sort(function(a,b){return a.price-b.price})}
 if(CURLIST_SORT==='name'){arr.sort(function(a,b){return normAr(a.name).localeCompare(normAr(b.name),'ar')})}
 if(CURLIST_SORT==='code'){arr.sort(function(a,b){return (Number(a.code)||0)-(Number(b.code)||0)})}
 var rows=arr.map(function(it){
  return {
   'الرمز':it.code,
   'اسم المادة':it.name,
   'الباركود':it.barcode,
   'سعر المبيع':it.price,
   'الكمية':it.qty,
   'الوحدة':it.unit,
   'المجموعة':it.group
  };
 });
 window.CURLIST={title:CURLIST_TITLE,rows:rows};
 $('mTitle').textContent=CURLIST_TITLE+' ('+arr.length+')';
 var h='<div class="row" style="margin-bottom:8px">';
 h=h+'<button class="btn ghost" id="btnExpList">⬇️ تصدير Excel</button>';
 h=h+'<select id="sortSel" class="inp" style="width:auto;flex:1">';
 sortOptions().forEach(function(o){
  h=h+'<option value="'+o[0]+'"'+(CURLIST_SORT===o[0]?' selected':'')+'>'+o[1]+'</option>';
 });
 h=h+'</select></div>';
 h=h+'<div class="list">'+(arr.map(itemCard).join('')||'<p class="mut">لا يوجد</p>')+'</div>';
 $('mBody').innerHTML=h;
 $('modal').hidden=false;
 bindItems($('mBody'));
 $('btnExpList').onclick=function(){if(window.exportItemsList){window.exportItemsList()}};
 $('sortSel').onchange=function(e){
  CURLIST_SORT=e.target.value;
  renderListModal();
 };
}
