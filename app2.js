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
 var th=(IMGS[it.code]||[])[0];
 var img=th?'<img src="'+th.thumb+'"/>':'<div class="ph">📦</div>';
 var al=it.aliases.length?' • أيضًا: '+it.aliases.join('، '):'';
 return '<div class="item" data-code="'+it.code+'">'+img+'<div style="flex:1"><b>'+it.name+'</b><small>'+(it.group||'')+al+'</small></div><span class="badge">💰 '+fmt(it.price)+'</span><span class="badge">🔢 '+it.qty+'</span></div>';
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
 stopVision();
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
function sigFrom(src){
 var c=$('vCanvas');
 c.width=c.height=32;
 var x=c.getContext('2d');
 x.drawImage(src,0,0,32,32);
 var d=x.getImageData(0,0,32,32).data;
 var sig=[];
 for(var i=0;i<d.length;i=i+4){
  sig.push(Math.round(0.299*d[i]+0.587*d[i+1]+0.114*d[i+2]));
 }
 var t=document.createElement('canvas');
 t.width=t.height=96;
 t.getContext('2d').drawImage(src,0,0,96,96);
 return {sig:sig,thumb:t.toDataURL('image/jpeg',0.6)};
}
function matchSig(sig){
 var out=[];
 Object.keys(IMGS).forEach(function(code){
  IMGS[code].forEach(function(im){
   var s=0;
   for(var i=0;i<sig.length;i=i+1){
    s=s+Math.abs(sig[i]-im.sig[i]);
   }
   out.push({code:code,sim:1-(s/sig.length)/255,thumb:im.thumb});
  });
 });
 out.sort(function(a,b){return b.sim-a.sim});
 return out.slice(0,5);
}
function capture(src){
 var r=sigFrom(src);
 curSig=r;
 $('vThumb').src=r.thumb;
 $('vThumbWrap').hidden=false;
 var m=matchSig(r.sig);
 $('vResults').innerHTML=m.length?m.map(function(x){
  var it=null;
  ITEMS.forEach(function(i){if(i.code===x.code){it=i}});
  return '<div class="item"><img src="'+x.thumb+'"/><div style="flex:1"><b>'+(it?it.name:'صنف محذوف')+'</b><small>رمز: '+x.code+'</small></div><span class="sim">'+Math.round(x.sim*100)+'%</span><button class="btn ghost" data-teach="'+x.code+'">تعليم ✔</button></div>';
 }).join(''):'<p class="mut">لا صور مرجعية بعد — علِّم النظام من بطاقة المادة.</p>';
 $('vResults').querySelectorAll('[data-teach]').forEach(function(b){
  b.onclick=function(){
   var code=b.getAttribute('data-teach');
   IMGS[code]=IMGS[code]||[];
   IMGS[code].push({sig:curSig.sig,thumb:curSig.thumb});
   saveLS('imgMap',IMGS);
   toast('تم تعليم النظام ✔');
  };
 });
 bindItems($('vResults'));
}
function stopVision(){
 if(vStream){
  vStream.getTracks().forEach(function(t){t.stop()});
  vStream=null;
 }
 $('vVideo').hidden=true;
 $('btnStopCam').hidden=true;
 $('btnCam').textContent='▶ فتح الكاميرا';
 camOpen=false;
}
$('btnStopCam').onclick=stopVision;
$('btnCam').onclick=function(){
 if(camOpen){
  capture($('vVideo'));
  return;
 }
 stopScan();
 navigator.mediaDevices.getUserMedia({video:{facingMode:'environment'}}).then(function(st){
  vStream=st;
  var v=$('vVideo');
  v.srcObject=vStream;
  v.hidden=false;
  $('btnStopCam').hidden=false;
  $('btnCam').textContent='📸 التقاط صورة';
  camOpen=true;
 }).catch(function(){
  toast('✖ تعذر فتح الكاميرا — استخدم اختيار صورة',1);
 });
};
$('visionFile').addEventListener('change',function(e){
 var f=e.target.files[0];
 if(!f){return}
 var img=new Image();
 img.onload=function(){capture(img)};
 img.src=URL.createObjectURL(f);
});
function openModal(code){
 var it=null;
 ITEMS.forEach(function(i){if(i.code===code){it=i}});
 if(!it){return}
 $('mTitle').textContent=it.name;
 var imgs=IMGS[code]||[];
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
 var ims='<span class="mut small">لا صور بعد</span>';
 if(imgs.length){
  ims=imgs.map(function(im,i){
   return '<img src="'+im.thumb+'" style="width:64px;height:64px;border-radius:8px" data-rm="'+i+'"/>';
  }).join('');
 }
 $('mBody').innerHTML='<div class="chips">'+chips+'</div><h4>🏷 الأسماء البديلة</h4><div class="chips" id="mAl">'+als+'</div><div class="row" style="margin-top:8px"><input id="mAlIn" class="inp grow" placeholder="أضف اسمًا شعبيًا…"/><button class="btn" id="mAlAdd">＋</button></div><h4>🖼 الصور المرجعية</h4><div class="chips">'+ims+'</div><label class="btn ghost" style="display:inline-block;margin-top:8px">＋ إضافة صورة مرجعية<input type="file" id="mImg" accept="image/*" class="hid"></label>';
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
 $('mBody').querySelectorAll('[data-rm]').forEach(function(im){
  im.onclick=function(){
   imgs.splice(Number(im.getAttribute('data-rm')),1);
   IMGS[code]=imgs;
   saveLS('imgMap',IMGS);
   openModal(code);
  };
 });
 $('mImg').addEventListener('change',function(e){
  var f=e.target.files[0];
  if(!f){return}
  var img=new Image();
  img.onload=function(){
   IMGS[code]=IMGS[code]||[];
   IMGS[code].push(sigFrom(img));
   saveLS('imgMap',IMGS);
   openModal(code);
   toast('أُضيفت الصورة ✔');
  };
  img.src=URL.createObjectURL(f);
 });
}
$('mClose').onclick=function(){$('modal').hidden=true};
$('modal').onclick=function(e){
 if(e.target.id==='modal'){$('modal').hidden=true}
};
function listModal(title,arr){
 $('mTitle').textContent=title+' ('+arr.length+')';
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
 window.CURLIST={title:title,rows:rows};
 var h='<div class="row" style="margin-bottom:8px"><button class="btn ghost" id="btnExpList">⬇️ تصدير هذه القائمة Excel</button></div>';
 h=h+'<div class="list">'+(arr.map(itemCard).join('')||'<p class="mut">لا يوجد</p>')+'</div>';
 $('mBody').innerHTML=h;
 $('modal').hidden=false;
 bindItems($('mBody'));
 var be=$('btnExpList');
 if(be){be.onclick=function(){if(window.exportItemsList){window.exportItemsList()}};}
}
