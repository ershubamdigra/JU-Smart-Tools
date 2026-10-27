(function(){
'use strict';
var U=window.JU,$=U.$,toast=U.toast,fmt=U.fmt,esc=U.esc,base=U.base,save=U.save,show=U.show,busy=U.busy;
var WORKER='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
function lib(){return U.need('PDFLib','PDF editor')}
function isPdf(f){return /\.pdf$/i.test(f.name)||f.type==='application/pdf'}
async function bytesOf(f){return new Uint8Array(await f.arrayBuffer())}
async function openDoc(bytes){return PDFLib.PDFDocument.load(bytes,{ignoreEncryption:true})}
function pdfBlob(b){return new Blob([b],{type:'application/pdf'})}
async function pdfjs(){
  var L=window.pdfjsLib;
  if(!L.GlobalWorkerOptions.workerSrc){var t=await (await fetch(WORKER)).text();L.GlobalWorkerOptions.workerSrc=URL.createObjectURL(new Blob([t],{type:'text/javascript'}))}
  return L;
}
/* "1-3, 5" -> [[0,1,2],[4]] (0-based). Bad input returns null. */
function ranges(str,n){
  var out=[],parts=String(str).split(',');
  for(var i=0;i<parts.length;i++){
    var s=parts[i].trim();if(!s)continue;
    var m=s.match(/^(\d+)(?:\s*-\s*(\d+))?$/);if(!m)return null;
    var a=+m[1],b=m[2]?+m[2]:a;if(a<1||b<a||b>n)return null;
    var r=[];for(var k=a;k<=b;k++)r.push(k-1);out.push(r);
  }
  return out.length?out:null;
}
function flat(r){var s={};r.forEach(function(x){x.forEach(function(i){s[i]=1})});return Object.keys(s).map(Number).sort(function(a,b){return a-b})}

/* ---------- Compression (keeps text) ---------- */
var LV=[{q:.82,d:2400},{q:.7,d:1800},{q:.55,d:1400},{q:.42,d:1100},{q:.3,d:850},{q:.2,d:650},{q:.12,d:480}];
var RL=[{s:1.4,q:.6},{s:1.1,q:.5},{s:.9,q:.4},{s:.7,q:.35},{s:.55,q:.3}];
async function shrink(bytes,lv){
  var PL=PDFLib,N=PL.PDFName,doc=await openDoc(bytes),ctx=doc.context,list=[];
  ctx.enumerateIndirectObjects().forEach(function(p){
    var o=p[1];if(!(o instanceof PL.PDFRawStream))return;
    var d=o.dict;
    if(d.get(N.of('Subtype'))!==N.of('Image')||String(d.get(N.of('Filter')))!=='/DCTDecode')return;
    var im=d.get(N.of('ImageMask'));if(im&&String(im)==='true')return;
    list.push(p);
  });
  for(var i=0;i<list.length;i++){
    var ref=list[i][0],o=list[i][1],url=URL.createObjectURL(new Blob([o.contents],{type:'image/jpeg'}));
    try{
      var img=await new Promise(function(r,j){var t=new Image();t.onload=function(){r(t)};t.onerror=j;t.src=url});
      var k=Math.min(1,lv.d/Math.max(img.naturalWidth,img.naturalHeight)),w=Math.max(1,Math.round(img.naturalWidth*k)),h=Math.max(1,Math.round(img.naturalHeight*k));
      var c=document.createElement('canvas');c.width=w;c.height=h;
      var x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,w,h);x.drawImage(img,0,0,w,h);
      var nb=await new Promise(function(r){c.toBlob(r,'image/jpeg',lv.q)});
      if(nb&&nb.size<o.contents.length){
        var e={Type:'XObject',Subtype:'Image',Width:w,Height:h,ColorSpace:'DeviceRGB',BitsPerComponent:8,Filter:'DCTDecode'},sm=o.dict.get(N.of('SMask'));
        if(sm)e.SMask=sm;
        ctx.assign(ref,ctx.stream(new Uint8Array(await nb.arrayBuffer()),e));
      }
    }catch(err){}
    URL.revokeObjectURL(url);
  }
  return doc.save({useObjectStreams:true});
}
async function rasterize(bytes,s,q){
  var L=await pdfjs(),pdf=await L.getDocument({data:bytes.slice()}).promise,out=await PDFLib.PDFDocument.create();
  for(var p=1;p<=pdf.numPages;p++){
    var pg=await pdf.getPage(p),vp=pg.getViewport({scale:s}),one=pg.getViewport({scale:1});
    var c=document.createElement('canvas');c.width=Math.round(vp.width);c.height=Math.round(vp.height);
    await pg.render({canvasContext:c.getContext('2d'),viewport:vp}).promise;
    var jb=await new Promise(function(r){c.toBlob(r,'image/jpeg',q)});
    var emb=await out.embedJpg(new Uint8Array(await jb.arrayBuffer()));
    out.addPage([one.width,one.height]).drawImage(emb,{x:0,y:0,width:one.width,height:one.height});
  }
  return out.save();
}
/* target = bytes or 0. lvl = start level when there is no target. */
async function squeeze(bytes,target,lvl,pics,btn){
  if(target&&bytes.length<=target)return {bytes:bytes,tag:'',met:true,already:true};
  var best=bytes,tag='';
  var i0=target?0:lvl,i1=target?LV.length:lvl+1;
  for(var i=i0;i<i1;i++){
    if(btn)btn.textContent='Compressing... '+(i-i0+1);
    var r=await shrink(bytes,LV[i]);
    if(r.length<best.length){best=r;tag='text'}
    if(target&&best.length<=target)break;
  }
  if(target&&best.length>target&&pics){
    for(var j=0;j<RL.length;j++){
      if(btn)btn.textContent='Pictures mode... '+(j+1);
      var r2=await rasterize(bytes,RL[j].s,RL[j].q);
      if(r2.length<best.length){best=r2;tag='pictures'}
      if(best.length<=target)break;
    }
  }
  return {bytes:best,tag:tag,met:!target||best.length<=target};
}

/* ---------- One-PDF tools: shared picker ---------- */
function single(p,after){
  var st={file:null,bytes:null,doc:null,n:0};
  function reset(){st.file=null;show(p+'Tool',false);show(p+'Drop',true)}
  U.dz($(p+'Drop'),$(p+'File'),async function(fl){
    var f=fl.filter(isPdf)[0];if(!f)return toast('Please choose a PDF file.',1);
    if(!lib())return;
    try{
      st.file=f;st.bytes=await bytesOf(f);st.doc=await openDoc(st.bytes);st.n=st.doc.getPageCount();
      $(p+'Info').innerHTML='<b>'+esc(f.name)+'</b> \u00b7 '+st.n+' page'+(st.n>1?'s':'')+' \u00b7 '+fmt(f.size);
      show(p+'Drop',false);show(p+'Tool',true);if(after)after(st);
    }catch(e){st.file=null;toast('Could not open this PDF. It may be damaged or locked.',1)}
  });
  $(p+'Clear').onclick=reset;
  return st;
}

/* ---------- Merge ---------- */
var M=[];
function mDraw(){
  show('mTool',M.length>0);
  $('mSum').innerHTML='<b>'+M.length+'</b> file'+(M.length>1?'s':'');
  U.rows($('mList'),M,mDraw,false,true);
}
U.dz($('mDrop'),$('mFile'),async function(fl){
  if(!lib())return;
  for(var i=0;i<fl.length;i++){
    if(!isPdf(fl[i]))continue;
    try{var d=await openDoc(await bytesOf(fl[i]));M.push({file:fl[i],note:d.getPageCount()+' pages'})}
    catch(e){toast('Could not open '+fl[i].name,1)}
  }
  mDraw();
});
$('mAdd').onclick=function(){$('mFile').click()};
$('mGo').onclick=async function(){
  if(M.length<1||!lib())return;
  var btn=$('mGo');busy(btn,true,'Merging...');
  try{
    var out=await PDFLib.PDFDocument.create();
    for(var i=0;i<M.length;i++){
      var src=await openDoc(await bytesOf(M[i].file));
      (await out.copyPages(src,src.getPageIndices())).forEach(function(p){out.addPage(p)});
    }
    var rem=$('mRem').value.trim();
    if(rem){
      var r=ranges(rem,out.getPageCount());
      if(!r){toast('Check the pages to remove. Example: 2, 5-7',1);busy(btn,false);return}
      var del=flat(r);if(del.length>=out.getPageCount()){toast('You cannot remove every page.',1);busy(btn,false);return}
      del.reverse().forEach(function(i){out.removePage(i)});
    }
    var bytes=await out.save(),lvl=+$('mComp').value;
    if(lvl>=0)bytes=(await squeeze(bytes,0,lvl,false,btn)).bytes;
    save(pdfBlob(bytes),'merged.pdf');toast('Merged PDF ready ('+fmt(bytes.length)+')');
  }catch(e){toast('Could not merge these files.',1)}
  busy(btn,false);
};

/* ---------- Split ---------- */
var S=single('s');
$('sMode').onchange=function(){show('sRangeBox',$('sMode').value!=='each')};
$('sGo').onclick=async function(){
  if(!S.file)return;
  var btn=$('sGo'),mode=$('sMode').value,groups;
  if(mode==='each'){groups=S.doc.getPageIndices().map(function(i){return [i]})}
  else{groups=ranges($('sRange').value,S.n);if(!groups)return toast('Check the pages. Example: 1-3, 5, 8-10 (the PDF has '+S.n+' pages).',1)}
  busy(btn,true,'Splitting...');
  try{
    async function make(idx){var o=await PDFLib.PDFDocument.create();(await o.copyPages(S.doc,idx)).forEach(function(p){o.addPage(p)});return o.save()}
    var nm=base(S.file.name);
    if(mode==='extract'){save(pdfBlob(await make(flat(groups))),nm+'-extract.pdf')}
    else if(groups.length===1){save(pdfBlob(await make(groups[0])),nm+'-part1.pdf')}
    else{
      if(!U.need('JSZip','ZIP library')){busy(btn,false);return}
      var z=new JSZip();
      for(var i=0;i<groups.length;i++){z.file(nm+'-'+(mode==='each'?'page':'part')+(i+1)+'.pdf',await make(groups[i]))}
      save(await z.generateAsync({type:'blob'}),nm+'-split.zip');
    }
    toast('Done');
  }catch(e){toast('Could not split this PDF.',1)}
  busy(btn,false);
};

/* ---------- Compress / resize ---------- */
var C=single('c',function(){$('cRes').textContent=''});
$('cMode').onchange=function(){var t=$('cMode').value==='target';show('cTargetBox',t);show('cLevelBox',!t)};
$('cGo').onclick=async function(){
  if(!C.file)return;
  var btn=$('cGo'),t=$('cMode').value==='target',target=t?Math.round(+$('cSize').value*+$('cUnit').value):0;
  if(t&&!(target>0))return toast('Enter a target size.',1);
  busy(btn,true,'Compressing...');
  try{
    var r=await squeeze(C.bytes,target,+$('cLevel').value,$('cPics').checked,btn),old=C.bytes.length,d=(1-r.bytes.length/old)*100;
    if(r.already){$('cRes').innerHTML='Your PDF is <b>'+fmt(old)+'</b>. That is already under the target, so nothing was changed.'}
    else if(!r.tag){$('cRes').innerHTML='This PDF could not be made smaller. It may have no large pictures, or it is already compressed.'}
    else{
      save(pdfBlob(r.bytes),base(C.file.name)+'-compressed.pdf');
      $('cRes').innerHTML=fmt(old)+' \u2192 <b>'+fmt(r.bytes.length)+'</b> <span class="good">(-'+d.toFixed(0)+'%)</span>. '+
        (r.tag==='pictures'?'Pictures mode was used, so text is not selectable.':'Text is still selectable.')+
        (r.met?'':' <span class="badn">The target was not reached.</span> Try "pictures mode" or a bigger target.');
    }
  }catch(e){toast('Could not compress this PDF.',1)}
  busy(btn,false);
};

/* ---------- Multiple pages per sheet ---------- */
var NU=single('n');
$('nGo').onclick=async function(){
  if(!NU.file)return;
  var btn=$('nGo'),per=+$('nPer').value,m=+$('nMar').value,gap=8;
  var cfg={2:[2,1,true],4:[2,2,false],6:[3,2,true],9:[3,3,false]}[per],cols=cfg[0],rows=cfg[1];
  var W=cfg[2]?841.89:595.28,H=cfg[2]?595.28:841.89;
  busy(btn,true,'Working...');
  try{
    var out=await PDFLib.PDFDocument.create(),emb=await out.embedPdf(NU.doc,NU.doc.getPageIndices());
    var cw=(W-2*m-(cols-1)*gap)/cols,ch=(H-2*m-(rows-1)*gap)/rows,gray=PDFLib.rgb(.75,.75,.75);
    for(var i=0;i<emb.length;i+=per){
      var pg=out.addPage([W,H]);
      for(var k=0;k<per&&i+k<emb.length;k++){
        var e=emb[i+k],col=k%cols,row=Math.floor(k/cols),s=Math.min(cw/e.width,ch/e.height),w=e.width*s,h=e.height*s;
        var x=m+col*(cw+gap)+(cw-w)/2,y=H-m-row*(ch+gap)-ch+(ch-h)/2;
        pg.drawPage(e,{x:x,y:y,width:w,height:h});
        pg.drawRectangle({x:x,y:y,width:w,height:h,borderColor:gray,borderWidth:.5});
      }
    }
    save(pdfBlob(await out.save()),base(NU.file.name)+'-'+per+'-per-sheet.pdf');toast('Done');
  }catch(e){toast('Could not create the sheets.',1)}
  busy(btn,false);
};

/* ---------- Add image to PDF ---------- */
var A=single('a');
$('aW').oninput=function(){$('aWv').textContent=$('aW').value};
$('aO').oninput=function(){$('aOv').textContent=$('aO').value};
$('aGo').onclick=async function(){
  if(!A.file)return;
  var f=$('aImg').files[0];if(!f)return toast('Please choose an image first.',1);
  var btn=$('aGo');busy(btn,true,'Working...');
  try{
    var li=await U.loadImg(f),im=li.img,k=Math.min(1,2000/Math.max(im.naturalWidth,im.naturalHeight));
    var c=document.createElement('canvas');c.width=Math.round(im.naturalWidth*k);c.height=Math.round(im.naturalHeight*k);
    c.getContext('2d').drawImage(im,0,0,c.width,c.height);URL.revokeObjectURL(li.url);
    var pb=await new Promise(function(r){c.toBlob(r,'image/png')});
    var doc=await openDoc(A.bytes),emb=await doc.embedPng(new Uint8Array(await pb.arrayBuffer())),op=+$('aO').value/100;
    if($('aNew').checked){
      var W=595.28,H=841.89,s=Math.min((W-60)/emb.width,(H-60)/emb.height),w=emb.width*s,h=emb.height*s;
      doc.addPage([W,H]).drawImage(emb,{x:(W-w)/2,y:(H-h)/2,width:w,height:h,opacity:op});
    }else{
      var txt=$('aPages').value.trim().toLowerCase(),idx;
      if(!txt||txt==='all')idx=doc.getPageIndices();
      else{var r=ranges(txt,A.n);if(!r){busy(btn,false);return toast('Check the pages. Example: 1, 3-5',1)}idx=flat(r)}
      var pos=$('aPos').value,pct=+$('aW').value/100,pages=doc.getPages();
      idx.forEach(function(i){
        var pg=pages[i],sz=pg.getSize(),w=sz.width*pct,h=w*emb.height/emb.width,mg=20,x,y;
        x=pos[1]==='l'?mg:pos[1]==='r'?sz.width-w-mg:(sz.width-w)/2;
        y=pos[0]==='t'?sz.height-h-mg:pos[0]==='b'?mg:(sz.height-h)/2;
        pg.drawImage(emb,{x:x,y:y,width:w,height:h,opacity:op});
      });
    }
    save(pdfBlob(await doc.save()),base(A.file.name)+'-with-image.pdf');toast('Done');
  }catch(e){toast('Could not add the image.',1)}
  busy(btn,false);
};

/* ---------- Sign PDF ---------- */
var sig=null,gm='draw',fx=.75,fy=.88;
var G=single('g',async function(st){
  try{
    var L=await pdfjs();G.pj=await L.getDocument({data:st.bytes.slice()}).promise;G.cur=1;
    if(sig)gRender();
  }catch(e){toast('Could not preview this PDF.',1)}
});
$('gClear').addEventListener('click',function(){show('gPlace',false)});
async function gRender(){
  if(!G.pj)return;
  var pg=await G.pj.getPage(G.cur),v1=pg.getViewport({scale:1}),s=Math.min(1.4,760/v1.width),vp=pg.getViewport({scale:s}),c=$('gPrev');
  c.width=Math.round(vp.width);c.height=Math.round(vp.height);
  await pg.render({canvasContext:c.getContext('2d'),viewport:vp}).promise;
  $('gPg').textContent='Page '+G.cur+' of '+G.n;gPos();
}
function gPos(){var im=$('gSig');im.style.width=$('gSize').value+'%';im.style.left=(fx*100)+'%';im.style.top=(fy*100)+'%';im.classList.toggle('hidden',!sig);$('gSv').textContent=$('gSize').value}
$('gWrap').onclick=function(e){if(!sig)return;var r=this.getBoundingClientRect();fx=(e.clientX-r.left)/r.width;fy=(e.clientY-r.top)/r.height;gPos()};
$('gSize').oninput=gPos;
$('gPrevB').onclick=function(){if(G.cur>1){G.cur--;gRender()}};
$('gNextB').onclick=function(){if(G.cur<G.n){G.cur++;gRender()}};
document.querySelectorAll('.gtab').forEach(function(b){b.onclick=function(){
  gm=b.dataset.m;
  document.querySelectorAll('.gtab').forEach(function(x){x.classList.toggle('on',x===b)});
  show('gDraw',gm==='draw');show('gType',gm==='type');show('gUp',gm==='up');
}});
var pad=$('gPad'),pc=pad.getContext('2d'),drawing=false;
pc.lineWidth=3;pc.lineCap='round';pc.lineJoin='round';
function pp(e){var r=pad.getBoundingClientRect();return [(e.clientX-r.left)*pad.width/r.width,(e.clientY-r.top)*pad.height/r.height]}
pad.onpointerdown=function(e){drawing=true;pad.setPointerCapture(e.pointerId);pc.strokeStyle=$('gCol').value;var p=pp(e);pc.beginPath();pc.moveTo(p[0],p[1]);pc.lineTo(p[0]+.1,p[1]+.1);pc.stroke()};
pad.onpointermove=function(e){if(!drawing)return;var p=pp(e);pc.lineTo(p[0],p[1]);pc.stroke()};
pad.onpointerup=pad.onpointercancel=function(){drawing=false};
$('gClr').onclick=function(){pc.clearRect(0,0,pad.width,pad.height)};
function trim(c){
  var w=c.width,h=c.height,d=c.getContext('2d').getImageData(0,0,w,h).data,x0=w,y0=h,x1=-1,y1=-1;
  for(var y=0;y<h;y++)for(var i=0;i<w;i++){if(d[(y*w+i)*4+3]>10){if(i<x0)x0=i;if(i>x1)x1=i;if(y<y0)y0=y;if(y>y1)y1=y}}
  if(x1<0)return null;
  var o=document.createElement('canvas'),m=4,cw=x1-x0+1,ch=y1-y0+1;
  o.width=cw+2*m;o.height=ch+2*m;o.getContext('2d').drawImage(c,x0,y0,cw,ch,m,m,cw,ch);return o;
}
$('gUse').onclick=async function(){
  var c,x;
  if(gm==='draw')c=pad;
  else if(gm==='type'){
    var t=$('gTxt').value.trim();if(!t)return toast('Type your name first.',1);
    var f='72px '+$('gFont').value;c=document.createElement('canvas');x=c.getContext('2d');x.font=f;
    c.width=Math.ceil(x.measureText(t).width)+40;c.height=140;
    x=c.getContext('2d');x.font=f;x.fillStyle=$('gCol').value;x.textBaseline='middle';x.fillText(t,20,70);
  }else{
    var fl=$('gImg').files[0];if(!fl)return toast('Choose a signature image first.',1);
    var li=await U.loadImg(fl),k=Math.min(1,900/li.img.naturalWidth);
    c=document.createElement('canvas');c.width=Math.round(li.img.naturalWidth*k);c.height=Math.round(li.img.naturalHeight*k);
    x=c.getContext('2d');x.drawImage(li.img,0,0,c.width,c.height);URL.revokeObjectURL(li.url);
    if($('gRm').checked){var id=x.getImageData(0,0,c.width,c.height),d=id.data;for(var i=0;i<d.length;i+=4){if(d[i]>225&&d[i+1]>225&&d[i+2]>225)d[i+3]=0}x.putImageData(id,0,0)}
  }
  var t2=trim(c);if(!t2)return toast('Please draw, type or choose a signature first.',1);
  sig=t2;$('gSig').src=t2.toDataURL('image/png');show('gPlace',true);gRender();
};
$('gGo').onclick=async function(){
  if(!G.file)return;if(!sig)return toast('Create a signature first.',1);
  var btn=$('gGo');busy(btn,true,'Signing...');
  try{
    var doc=await openDoc(G.bytes),pb=await new Promise(function(r){sig.toBlob(r,'image/png')});
    var emb=await doc.embedPng(new Uint8Array(await pb.arrayBuffer())),font=await doc.embedFont(PDFLib.StandardFonts.Helvetica);
    var pages=doc.getPages(),ap=$('gApply').value,idx=ap==='all'?doc.getPageIndices():ap==='last'?[pages.length-1]:[G.cur-1];
    idx.forEach(function(i){
      var pg=pages[i],sz=pg.getSize(),w=sz.width*$('gSize').value/100,h=w*emb.height/emb.width;
      var x=Math.min(Math.max(0,fx*sz.width-w/2),sz.width-w),y=Math.min(Math.max(14,sz.height-fy*sz.height-h/2),sz.height-h);
      pg.drawImage(emb,{x:x,y:y,width:w,height:h});
      if($('gDate').checked)pg.drawText(new Date().toLocaleDateString(),{x:x,y:y-12,size:10,font:font,color:PDFLib.rgb(.1,.1,.1)});
    });
    save(pdfBlob(await doc.save()),base(G.file.name)+'-signed.pdf');toast('Signed PDF ready');
  }catch(e){toast('Could not sign this PDF.',1)}
  busy(btn,false);
};

/* ---------- Fill PDF form ---------- */
var FF=single('f',function(st){
  var PL=PDFLib,html='',n=0;
  try{
    st.fs=st.doc.getForm().getFields();
    st.fs.forEach(function(f,i){
      var inp='';
      if(f instanceof PL.PDFTextField)inp='<input type="text" data-i="'+i+'" value="'+esc(f.getText()||'')+'">';
      else if(f instanceof PL.PDFCheckBox)inp='<input type="checkbox" data-i="'+i+'"'+(f.isChecked()?' checked':'')+'>';
      else if(f instanceof PL.PDFDropdown||f instanceof PL.PDFRadioGroup||f instanceof PL.PDFOptionList){
        var s=f.getSelected(),cur=Array.isArray(s)?s[0]:s;
        inp='<select data-i="'+i+'"><option value="">(none)</option>'+f.getOptions().map(function(o){return '<option value="'+esc(o)+'"'+(o===cur?' selected':'')+'>'+esc(o)+'</option>'}).join('')+'</select>';
      }else return;
      n++;html+='<div class="frow"><label>'+esc(f.getName())+'</label>'+inp+'</div>';
    });
  }catch(e){n=0}
  $('fFields').innerHTML=n?html:'<p class="note">No fillable fields were found in this PDF. It may be a flat (printed) form.</p>';
  show('fGo',n>0);show('fFlatBox',n>0);
});
$('fGo').onclick=async function(){
  if(!FF.file)return;
  var btn=$('fGo');busy(btn,true,'Filling...');
  try{
    var doc=await openDoc(FF.bytes),form=doc.getForm(),fs=form.getFields(),PL=PDFLib;
    document.querySelectorAll('#fFields [data-i]').forEach(function(el){
      var f=fs[+el.dataset.i];
      if(f instanceof PL.PDFTextField)f.setText(el.value);
      else if(f instanceof PL.PDFCheckBox){if(el.checked)f.check();else f.uncheck()}
      else if(el.value)f.select(el.value);
    });
    if($('fFlat').checked)form.flatten();
    save(pdfBlob(await doc.save()),base(FF.file.name)+'-filled.pdf');toast('Filled PDF ready');
  }catch(e){toast('Could not fill the form. Check for unsupported characters.',1)}
  busy(btn,false);
};
})();
