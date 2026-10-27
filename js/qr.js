/* QR Code Generator tab. Uses js/vendor/qrcode.js (MIT, Kazuhiko Arase). */
(function(){
'use strict';
var U=window.JU;if(!U||!document.getElementById('qText'))return;
var $=U.$,cur=null;
if(window.qrcode&&qrcode.stringToBytesFuncs&&qrcode.stringToBytesFuncs['UTF-8'])qrcode.stringToBytes=qrcode.stringToBytesFuncs['UTF-8'];

function lum(hex){
  var r=parseInt(hex.substr(1,2),16)/255,g=parseInt(hex.substr(3,2),16)/255,b=parseInt(hex.substr(5,2),16)/255;
  function f(v){return v<=.03928?v/12.92:Math.pow((v+.055)/1.055,2.4)}
  return .2126*f(r)+.7152*f(g)+.0722*f(b);
}
function fixLink(t){
  /* "example.com/page" -> "https://example.com/page" */
  if(/^[a-z][a-z0-9+.-]*:/i.test(t))return t;
  if(/^(www\.)?[a-z0-9-]+(\.[a-z0-9-]+)+(:\d+)?(\/\S*)?$/i.test(t))return 'https://'+t;
  return t;
}
function setReady(on){
  $('qPng').disabled=!on;$('qSvg').disabled=!on;
  U.show('qCanvas',on);U.show('qEmpty',!on);
}
function build(){
  var raw=$('qText').value.trim();cur=null;$('qMsg').textContent='';
  if(!raw){setReady(false);return}
  if(!window.qrcode){setReady(false);$('qMsg').textContent='The QR library did not load. Check that js/vendor/qrcode.js is in your folder.';return}
  var data=fixLink(raw),qr;
  try{qr=qrcode(0,$('qEc').value);qr.addData(data);qr.make()}
  catch(e){setReady(false);$('qMsg').textContent='This is too long for the chosen error correction. Use a shorter text or a lower level (L or M).';return}
  var n=qr.getModuleCount(),m=+$('qMargin').value,tot=n+2*m,size=+$('qSize').value,cell=Math.max(1,Math.floor(size/tot));
  var fg=$('qFg').value,bg=$('qBg').value,c=$('qCanvas');
  c.width=c.height=cell*tot;
  var x=c.getContext('2d');x.fillStyle=bg;x.fillRect(0,0,c.width,c.height);x.fillStyle=fg;
  for(var r=0;r<n;r++)for(var k=0;k<n;k++)if(qr.isDark(r,k))x.fillRect((k+m)*cell,(r+m)*cell,cell,cell);
  cur={qr:qr,n:n,m:m,fg:fg,bg:bg,data:data};
  setReady(true);
  var note=data!==raw?'Link saved as: '+data:'';
  if(Math.abs(lum(fg)-lum(bg))<.3)note+=(note?' \u00b7 ':'')+'Colors are too close. Some scanners may not read this QR code.';
  $('qMsg').textContent=note;
}
function svg(){
  var q=cur,t=q.n+2*q.m,d='';
  for(var r=0;r<q.n;r++){
    var k=0;
    while(k<q.n){
      if(q.qr.isDark(r,k)){var s=k;while(k<q.n&&q.qr.isDark(r,k))k++;d+='M'+(s+q.m)+' '+(r+q.m)+'h'+(k-s)+'v1h-'+(k-s)+'z'}
      else k++;
    }
  }
  return '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 '+t+' '+t+'" width="'+t*10+'" height="'+t*10+'" shape-rendering="crispEdges"><rect width="'+t+'" height="'+t+'" fill="'+q.bg+'"/><path d="'+d+'" fill="'+q.fg+'"/></svg>';
}
function fname(){
  var s=(cur.data||'qr').replace(/^[a-z]+:\/\//i,'').replace(/[^a-z0-9]+/gi,'-').replace(/^-+|-+$/g,'').slice(0,30);
  return 'qr-'+(s||'code');
}
['qText'].forEach(function(i){$(i).oninput=build});
['qSize','qEc','qMargin','qFg','qBg'].forEach(function(i){$(i).oninput=build;$(i).onchange=build});
$('qPng').onclick=function(){if(!cur)return;$('qCanvas').toBlob(function(b){U.save(b,fname()+'.png')},'image/png')};
$('qSvg').onclick=function(){if(!cur)return;U.save(new Blob([svg()],{type:'image/svg+xml'}),fname()+'.svg')};
setReady(false);
})();
