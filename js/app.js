(function(){
'use strict';
var $=function(i){return document.getElementById(i)};
var EXT={'image/jpeg':'jpg','image/webp':'webp','image/png':'png'};
var WORKER='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
function fmt(b){return b>1048576?(b/1048576).toFixed(2)+' MB':(b/1024).toFixed(1)+' KB'}
function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]})}
function base(n){return n.replace(/\.[^.]+$/,'')}
function save(blob,fn){var u=URL.createObjectURL(blob),a=document.createElement('a');a.href=u;a.download=fn;document.body.appendChild(a);a.click();a.remove();setTimeout(function(){URL.revokeObjectURL(u)},4000)}
function toast(m,err){var t=$('toast');t.textContent=m;t.className='show'+(err?' err':'');clearTimeout(toast.t);toast.t=setTimeout(function(){t.className=''},4000)}
function show(id,on){$(id).classList.toggle('hidden',!on)}
function isImg(f){return f.type.indexOf('image/')===0}
function loadImg(f){return new Promise(function(res,rej){var u=URL.createObjectURL(f),im=new Image();im.onload=function(){res({file:f,img:im,url:u})};im.onerror=function(){URL.revokeObjectURL(u);rej()};im.src=u})}
function dz(zone,input,cb){
  zone.onclick=function(){input.click()};
  zone.ondragover=function(e){e.preventDefault();zone.classList.add('over')};
  zone.ondragleave=function(){zone.classList.remove('over')};
  zone.ondrop=function(e){e.preventDefault();zone.classList.remove('over');cb([].slice.call(e.dataTransfer.files))};
  input.onchange=function(){cb([].slice.call(input.files));input.value=''};
}
function need(lib,name){if(!window[lib]){toast(name+' did not load. Please check your internet.',1);return false}return true}

/* Theme */
function setTheme(t){document.documentElement.setAttribute('data-theme',t);try{localStorage.setItem('theme',t)}catch(e){}$('theme').textContent=t==='dark'?'\u2600 Light mode':'\uD83C\uDF19 Dark mode'}
setTheme(document.documentElement.getAttribute('data-theme')||'light');
$('theme').onclick=function(){setTheme(document.documentElement.getAttribute('data-theme')==='dark'?'light':'dark')};

/* Navigation: home grid + one panel per tool */
function route(){
  var t=(location.hash||'').slice(1),p=t&&$('t-'+t);
  document.querySelectorAll('.panel').forEach(function(x){x.classList.toggle('hidden',x!==p)});
  show('home',!p);show('hero',!p);window.scrollTo(0,0);
}
document.querySelectorAll('.tool[data-t]').forEach(function(b){b.onclick=function(){location.hash=b.dataset.t}});
document.querySelectorAll('.back').forEach(function(b){b.onclick=function(){location.hash='';route()}});
window.onhashchange=route;route();
$('search').oninput=function(){var q=this.value.toLowerCase();document.querySelectorAll('.tool').forEach(function(c){c.classList.toggle('hidden',c.textContent.toLowerCase().indexOf(q)<0)})};

/* Simple file list with optional reorder */
function rows(el,arr,redraw,thumb,order){
  el.innerHTML=arr.map(function(it,i){
    return '<div class="item">'+(thumb?'<img src="'+it.url+'" alt="">':'<div class="ico">'+esc(it.file.name.split('.').pop().slice(0,4).toUpperCase())+'</div>')+
    '<div class="info"><div class="n">'+esc(it.file.name)+'</div><div class="m">'+fmt(it.file.size)+(it.note?' \u00b7 '+esc(it.note):'')+'</div></div>'+
    (order?'<button class="sec sm" data-a="u" data-i="'+i+'" title="Move up">\u2191</button><button class="sec sm" data-a="d" data-i="'+i+'" title="Move down">\u2193</button>':'')+
    '<button class="sec sm" data-a="x" data-i="'+i+'" title="Remove">\u2715</button></div>';
  }).join('');
  el.onclick=function(e){
    var b=e.target.closest('button');if(!b)return;
    var i=+b.dataset.i,a=b.dataset.a,t;
    if(a==='x')arr.splice(i,1);
    if(a==='u'&&i>0){t=arr[i];arr[i]=arr[i-1];arr[i-1]=t}
    if(a==='d'&&i<arr.length-1){t=arr[i];arr[i]=arr[i+1];arr[i+1]=t}
    redraw();
  };
}
function busy(btn,on,txt){btn.disabled=on;btn.classList.toggle('busy',on);if(on){btn._t=btn.textContent;btn.textContent=txt}else if(btn._t){btn.textContent=btn._t}}

/* ---------- 1. Resize images ---------- */
var R=[],rTimer=null,rRun=0;
function rLabels(){
  var q=+$('rQ').value;
  $('rQv').textContent=q;$('rQn').textContent=q<35?'(low)':q<70?'(medium)':'(high)';
  var tgOn=$('rTg').checked;show('rTgBox',tgOn);
  $('rQ').disabled=$('rFmt').value==='image/png'||tgOn;
  $('rPv').textContent=$('rPct').value;
  show('rPxBox',$('rMode').value==='px');show('rPctBox',$('rMode').value==='pct');
}
function rGo(){rLabels();clearTimeout(rTimer);rTimer=setTimeout(rRender,200)}
function rDims(it){
  var m=$('rMode').value,nw=it.img.naturalWidth,nh=it.img.naturalHeight,w,h;
  if(m==='orig'){w=nw;h=nh}
  else if(m==='pct'){var s=+$('rPct').value/100;w=nw*s;h=nh*s}
  else{w=+$('rW').value||nw;h=+$('rH').value||nh;if($('rLock').checked&&R.length>1)h=w*nh/nw}
  return [Math.max(1,Math.round(w)),Math.max(1,Math.round(h))];
}
function rEnc(it,w,h,type,q){
  return new Promise(function(res){
    var c=document.createElement('canvas');c.width=w;c.height=h;
    var x=c.getContext('2d');
    if(type==='image/jpeg'){x.fillStyle='#fff';x.fillRect(0,0,w,h)}
    x.imageSmoothingQuality='high';x.drawImage(it.img,0,0,w,h);
    c.toBlob(res,type,q);
  });
}
async function rOne(it){
  var d=rDims(it),w=d[0],h=d[1],type=$('rFmt').value,b=null,missed=false;
  if(w*h>120000000)return;
  var tg=$('rTg').checked?Math.round(+$('rTv').value*+$('rTu').value):0;
  if(!(tg>0)){b=await rEnc(it,w,h,type,+$('rQ').value/100)}
  else{
    var png=type==='image/png',last=null;
    for(var s=1;s>0.08;s*=0.85){
      var ww=Math.max(1,Math.round(w*s)),hh=Math.max(1,Math.round(h*s));
      if(png){
        var t=await rEnc(it,ww,hh,type,1);last={b:t,w:ww,h:hh};
        if(t.size<=tg){b=t;w=ww;h=hh;break}
        continue;
      }
      var lo=0.05,hi=0.95,t0=await rEnc(it,ww,hh,type,lo);last={b:t0,w:ww,h:hh};
      if(t0.size>tg)continue;
      b=t0;
      for(var i=0;i<7;i++){var mid=(lo+hi)/2,tb=await rEnc(it,ww,hh,type,mid);if(tb.size<=tg){b=tb;lo=mid}else hi=mid}
      w=ww;h=hh;break;
    }
    if(!b&&last){b=last.b;w=last.w;h=last.h;missed=true}
  }
  if(b){
    if(it.out)URL.revokeObjectURL(it.out);
    it.blob=b;it.out=URL.createObjectURL(b);it.w=w;it.h=h;it.ext=EXT[type];it.missed=missed;
  }
}
function rRender(){
  var id=++rRun,p=Promise.resolve();
  R.forEach(function(it){p=p.then(function(){return id===rRun?rOne(it):0})});
  p.then(function(){if(id===rRun)rDraw()});
}
function rName(it){return base(it.file.name)+'-'+it.w+'x'+it.h+'.'+it.ext}
function rDraw(){
  var tin=0,tout=0,h='';
  R.forEach(function(it,i){
    if(!it.blob)return;
    tin+=it.file.size;tout+=it.blob.size;
    var d=(1-it.blob.size/it.file.size)*100;
    h+='<div class="item"><img src="'+it.out+'" alt=""><div class="info"><div class="n">'+esc(it.file.name)+'</div><div class="m">'+fmt(it.file.size)+' ('+it.img.naturalWidth+'x'+it.img.naturalHeight+') \u2192 <b>'+fmt(it.blob.size)+'</b> ('+it.w+'x'+it.h+') <span class="'+(d>=0?'good':'badn')+'">'+(d>=0?'-':'+')+Math.abs(d).toFixed(0)+'%</span>'+(it.missed?' <span class="badn">Target not reached</span>':'')+'</div></div>'+
      '<button class="sm" data-i="'+i+'">Download</button><button class="sec sm" data-r="'+i+'">\u2715</button></div>';
  });
  $('rList').innerHTML=h;
  var d2=tin?(1-tout/tin)*100:0;
  $('rSum').innerHTML='<b>'+R.length+'</b> image'+(R.length>1?'s':'')+' \u00b7 '+fmt(tin)+' \u2192 <b>'+fmt(tout)+'</b> <span class="'+(d2>=0?'good':'badn')+'">('+(d2>=0?'-':'+')+Math.abs(d2).toFixed(0)+'%)</span>';
  $('rAll').textContent=R.length>1?'Download all (ZIP)':'Download';
}
function rReset(){
  R.forEach(function(i){URL.revokeObjectURL(i.url);if(i.out)URL.revokeObjectURL(i.out)});
  R=[];$('rW').value='';$('rH').value='';$('rList').innerHTML='';show('rTool',false);show('rDrop',true);
}
dz($('rDrop'),$('rFile'),rAdd);
$('rAdd').onclick=function(){$('rFile').click()};
function rAdd(fl){
  Promise.all(fl.filter(isImg).map(function(f){return loadImg(f).catch(function(){return null})})).then(function(a){
    a=a.filter(Boolean);if(!a.length)return;
    R=R.concat(a);
    if(!$('rW').value){$('rW').value=R[0].img.naturalWidth;$('rH').value=R[0].img.naturalHeight}
    show('rTool',true);rGo();
  });
}
$('rList').onclick=function(e){
  var b=e.target.closest('button');if(!b)return;
  if(b.dataset.i!==undefined){var it=R[+b.dataset.i];save(it.blob,rName(it))}
  if(b.dataset.r!==undefined){var r=R.splice(+b.dataset.r,1)[0];URL.revokeObjectURL(r.url);if(r.out)URL.revokeObjectURL(r.out);if(!R.length)rReset();else rDraw()}
};
$('rAll').onclick=function(){
  var ok=R.filter(function(i){return i.blob});if(!ok.length)return;
  if(ok.length===1){save(ok[0].blob,rName(ok[0]));return}
  if(!window.JSZip){ok.forEach(function(it,k){setTimeout(function(){save(it.blob,rName(it))},k*400)});return}
  var z=new JSZip(),used={};
  ok.forEach(function(it){var n=rName(it),k=n,c=1;while(used[k]){c++;k=n.replace(/(\.[^.]+)$/,'-'+c+'$1')}used[k]=1;z.file(k,it.blob)});
  busy($('rAll'),true,'Zipping...');
  z.generateAsync({type:'blob'}).then(function(b){save(b,'resized-images.zip');busy($('rAll'),false)});
};
$('rClear').onclick=rReset;
$('rW').oninput=function(){if($('rLock').checked&&R[0])$('rH').value=Math.round($('rW').value*R[0].img.naturalHeight/R[0].img.naturalWidth);rGo()};
$('rH').oninput=function(){if($('rLock').checked&&R[0])$('rW').value=Math.round($('rH').value*R[0].img.naturalWidth/R[0].img.naturalHeight);rGo()};
['rQ','rPct'].forEach(function(i){$(i).oninput=rGo});
['rFmt','rMode','rLock','rTg','rTu'].forEach(function(i){$(i).onchange=rGo});
$('rTv').oninput=rGo;
document.querySelectorAll('.chip').forEach(function(b){b.onclick=function(){$('rTv').value=b.dataset.v;$('rTu').value=b.dataset.u;$('rTg').checked=true;rGo()}});
rLabels();

/* ---------- 2. Image to PDF ---------- */
var P=[];
function pDraw(){
  show('pTool',P.length>0);
  $('pSum').innerHTML='<b>'+P.length+'</b> page'+(P.length>1?'s':'');
  rows($('pList'),P,pDraw,true,true);
}
dz($('pDrop'),$('pFile'),function(fl){
  Promise.all(fl.filter(isImg).map(function(f){return loadImg(f).catch(function(){return null})})).then(function(a){
    P=P.concat(a.filter(Boolean));pDraw();
  });
});
$('pAdd').onclick=function(){$('pFile').click()};
$('pGo').onclick=function(){
  if(!P.length||!need('jspdf','PDF library'))return;
  var size=$('pSize').value,mar=+$('pMargin').value,q=+$('pQ').value,doc=null,J=window.jspdf.jsPDF;
  busy($('pGo'),true,'Working...');
  setTimeout(function(){
    try{
      P.forEach(function(it){
        var k=Math.min(1,2600/Math.max(it.img.naturalWidth,it.img.naturalHeight));
        var c=document.createElement('canvas');c.width=Math.round(it.img.naturalWidth*k);c.height=Math.round(it.img.naturalHeight*k);
        var x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(it.img,0,0,c.width,c.height);
        var iw=c.width,ih=c.height,pw,ph;
        if(size==='fit'){pw=iw*0.75;ph=ih*0.75}
        else{var a=size==='a4'?[595.28,841.89]:[612,792],land=iw>ih;pw=land?a[1]:a[0];ph=land?a[0]:a[1]}
        var m=size==='fit'?0:mar,s=Math.min((pw-2*m)/iw,(ph-2*m)/ih),w=iw*s,h=ih*s,o=pw>ph?'l':'p';
        if(!doc)doc=new J({unit:'pt',format:[pw,ph],orientation:o});else doc.addPage([pw,ph],o);
        doc.addImage(c.toDataURL('image/jpeg',q),'JPEG',(pw-w)/2,(ph-h)/2,w,h);
      });
      doc.save(P.length===1?base(P[0].file.name)+'.pdf':'images.pdf');toast('PDF created');
    }catch(e){toast('Could not create the PDF.',1)}
    busy($('pGo'),false);
  },50);
};

/* 3. Doc to PDF lives in js/doc-pdf.js */

/* ---------- 4. PDF to Word ---------- */
var W=[];
function wDraw(){show('wTool',W.length>0);$('wSum').innerHTML='<b>'+W.length+'</b> file'+(W.length>1?'s':'');rows($('wList'),W,wDraw,false,false)}
dz($('wDrop'),$('wFile'),function(fl){
  var ok=fl.filter(function(f){return /\.pdf$/i.test(f.name)||f.type==='application/pdf'});
  if(ok.length<fl.length)toast('Only PDF files are supported here.',1);
  W=W.concat(ok.map(function(f){return {file:f}}));wDraw();
});
$('wAdd').onclick=function(){$('wFile').click()};
/* the converter itself lives in js/pdf-to-word.js */
window.JU={$:$,fmt:fmt,esc:esc,base:base,save:save,toast:toast,show:show,dz:dz,rows:rows,busy:busy,need:need,loadImg:loadImg,getW:function(){return W},redrawW:function(){wDraw()}};
})();
