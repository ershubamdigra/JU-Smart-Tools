/* Doc to PDF (Word .docx, .txt, .html).
   - .docx is drawn with docx-preview, which keeps the Word page size, margins, fonts, tables, headers and footers.
   - Pages that run past the page height are split cleanly: paragraphs and table rows are never cut in half.
   - Every page is drawn in a hidden, isolated frame, so nothing from this website can leak into the PDF.
   - An invisible text layer and real links are added, so the PDF text can be searched, selected and copied. */
(function(){
'use strict';
var U=window.JU;if(!U||!document.getElementById('dDrop'))return;
var $=U.$,D=[],fr=null,PX=0.75; /* CSS px -> PDF points */
var SIZES={a4:[794,1123],letter:[816,1056]};

function dDraw(){U.show('dTool',D.length>0);$('dSum').innerHTML='<b>'+D.length+'</b> file'+(D.length>1?'s':'');U.rows($('dList'),D,dDraw,false,false)}
U.dz($('dDrop'),$('dFile'),function(fl){
  var ok=fl.filter(function(f){return /\.(docx|txt|html?)$/i.test(f.name)});
  if(ok.length<fl.length)U.toast(fl.some(function(f){return /\.doc$/i.test(f.name)})?'Old .doc files are not supported. Open it in Word and save as .docx first.':'Only .docx, .txt and .html files are supported.',1);
  D=D.concat(ok.map(function(f){return {file:f}}));dDraw();
});
$('dAdd').onclick=function(){$('dFile').click()};

/* ---------- hidden isolated frame ---------- */
function frame(){
  if(fr&&fr.isConnected)return fr;
  fr=document.createElement('iframe');
  fr.setAttribute('aria-hidden','true');fr.tabIndex=-1;
  fr.style.cssText='position:fixed;left:-30000px;top:0;width:1400px;height:1600px;border:0;pointer-events:none';
  document.body.appendChild(fr);
  var d=fr.contentDocument;d.open();
  d.write('<!doctype html><html><head><meta charset="utf-8"></head><body style="margin:0;background:#fff"></body></html>');d.close();
  return fr;
}
function dropFrame(){if(fr&&fr.parentNode)fr.parentNode.removeChild(fr);fr=null}
function tick(){return new Promise(function(r){setTimeout(r,30)})}
async function settle(d){
  try{if(d.fonts&&d.fonts.ready)await d.fonts.ready}catch(e){}
  var imgs=[].slice.call(d.images||[]);
  await Promise.all(imgs.map(function(im){
    if(im.complete)return 0;
    return new Promise(function(r){im.onload=im.onerror=r;setTimeout(r,4000)});
  }));
  await tick();
}

/* ---------- Word pages that are taller than the paper: split them cleanly ---------- */
function paperH(sec){var w=sec.ownerDocument.defaultView;return parseFloat(w.getComputedStyle(sec).minHeight)||sec.offsetHeight}
function articleOf(sec){for(var i=0;i<sec.children.length;i++)if(sec.children[i].tagName==='ARTICLE')return sec.children[i];return null}
function newPageAfter(prev){
  var s=prev.cloneNode(false);
  [].slice.call(prev.children).forEach(function(ch){
    if(ch.tagName==='ARTICLE')s.appendChild(ch.cloneNode(false));
    else if(ch.tagName==='HEADER'||ch.tagName==='FOOTER')s.appendChild(ch.cloneNode(true));
  });
  prev.parentNode.insertBefore(s,prev.nextSibling);
  return s;
}
function splitSection(sec){
  var H=paperH(sec),art=articleOf(sec);
  if(!art||sec.offsetHeight<=H+1)return [sec];
  var queue=[].slice.call(art.children),pages=[sec],cur=sec,curArt=art,over=function(){return cur.offsetHeight>H+1};
  while(art.firstChild)art.removeChild(art.firstChild);
  var guard=0;
  while(queue.length&&guard++<20000){
    var b=queue.shift();
    curArt.appendChild(b);
    if(!over())continue;
    var tbl=b.tagName==='TABLE'?b:(b.querySelector&&b.children.length===1&&b.firstElementChild.tagName==='TABLE'?b.firstElementChild:null);
    if(tbl&&tbl.rows.length>2){
      var moved=[];
      while(over()&&tbl.rows.length>1){moved.unshift(tbl.rows[tbl.rows.length-1]);tbl.rows[tbl.rows.length-1].parentNode.removeChild(moved[0])}
      /* keep at least a header row and one more row together; otherwise move the whole table down */
      if(!over()&&moved.length&&tbl.rows.length>=3){
        var t2=tbl.cloneNode(false),cg=tbl.querySelector('colgroup');
        if(cg)t2.appendChild(cg.cloneNode(true));
        var tb=document.createElement('tbody');moved.forEach(function(r){tb.appendChild(r)});t2.appendChild(tb);
        queue.unshift(t2);
        cur=newPageAfter(cur);curArt=articleOf(cur);pages.push(cur);continue;
      }
      var tb0=tbl.tBodies[0]||tbl;moved.forEach(function(r){tb0.appendChild(r)});
    }
    if(curArt.children.length>1){
      curArt.removeChild(b);
      cur=newPageAfter(cur);curArt=articleOf(cur);pages.push(cur);
      queue.unshift(b); /* try again on the fresh page (it may still need row splitting) */
    }
  }
  /* drop a trailing page that has no visible content */
  return pages.filter(function(p,i){var a=articleOf(p);return i===0||(a&&a.textContent.trim().length>0)||(a&&a.querySelector('img,svg,table,canvas'))});
}

/* ---------- Plain text and HTML: flow into pages without cutting lines ---------- */
function clean(html){
  var doc=new DOMParser().parseFromString(html,'text/html');
  doc.querySelectorAll('script,iframe,object,embed,base,meta[http-equiv],link[rel~="import"],form input[type=hidden]').forEach(function(n){n.remove()});
  doc.querySelectorAll('*').forEach(function(n){[].slice.call(n.attributes).forEach(function(a){if(/^on/i.test(a.name)||/^\s*javascript:/i.test(a.value))n.removeAttribute(a.name)})});
  var css='';doc.querySelectorAll('style').forEach(function(s){css+=s.textContent+'\n'});
  var links='';doc.querySelectorAll('link[rel~="stylesheet"][href]').forEach(function(l){links+=l.outerHTML});
  return {css:css,links:links,body:doc.body?doc.body.innerHTML:''};
}
function atomsOf(flow){
  var d=flow.ownerDocument,top=flow.getBoundingClientRect().top,out=[];
  function add(r){if(r.height>0.5)out.push({t:r.top-top,b:r.bottom-top})}
  var tw=d.createTreeWalker(flow,NodeFilter.SHOW_TEXT,null),n,rg=d.createRange();
  while((n=tw.nextNode())){
    if(!n.nodeValue.trim())continue;
    rg.selectNodeContents(n);[].slice.call(rg.getClientRects()).forEach(add);
  }
  flow.querySelectorAll('img,svg,canvas,video,hr,tr').forEach(function(e){add(e.getBoundingClientRect())});
  return out;
}
function cutAt(y0,Hc,atoms){
  var lim=y0+Hc,cand=[];
  atoms.forEach(function(a){if(a.b>y0+1&&a.b<=lim+.5)cand.push(a.b)});
  cand.sort(function(x,y){return y-x});
  for(var i=0;i<cand.length;i++){
    var c=cand[i];
    if(!atoms.some(function(a){return a.t<c-1&&a.b>c+1}))return c;
  }
  return lim;
}
function flowPages(d,kind,text,pw,ph,m){
  var body=d.body,Wc=pw-2*m,Hc=ph-2*m,css,links='',inner;
  if(kind==='txt'){
    inner='<pre style="margin:0;white-space:pre-wrap;word-wrap:break-word;font:13px/1.5 \'Courier New\',Consolas,monospace;color:#000"></pre>';
  }else{
    var c=clean(text);css=c.css;links=c.links;inner='<div class="ju-html" style="font:15px/1.55 Arial,Helvetica,sans-serif;color:#000">'+c.body+'</div>';
  }
  d.head.insertAdjacentHTML('beforeend',(css?'<style>'+css+'</style>':'')+links+'<style>img{max-width:100%;height:auto;max-height:'+Hc+'px}table{border-collapse:collapse;max-width:100%}</style>');
  body.innerHTML='<div class="pgbox" style="position:relative;width:'+pw+'px;height:'+ph+'px;padding:'+m+'px;box-sizing:border-box;background:#fff;overflow:hidden"><div class="clip" style="position:relative;overflow:hidden;height:'+Hc+'px"><div class="flow" style="position:absolute;left:0;top:0;width:'+Wc+'px">'+inner+'</div></div></div>';
  var box=body.firstChild,clip=box.firstChild,flow=clip.firstChild;
  if(kind==='txt')flow.firstChild.textContent=text.replace(/\r\n?/g,'\n').replace(/\t/g,'    ');
  return {box:box,clip:clip,flow:flow,Wc:Wc,Hc:Hc,pw:pw,ph:ph};
}
function planFlow(F){
  var atoms=atomsOf(F.flow),H=F.flow.getBoundingClientRect().height,y0=0,plan=[],guard=0;
  /* start of the first real line */
  while(y0<H-1&&guard++<5000){
    var nxt=atoms.filter(function(a){return a.b>y0+1}).sort(function(a,b){return a.t-b.t})[0];
    if(nxt&&nxt.t>y0)y0=Math.max(y0,nxt.t-3);
    var c=(H-y0<=F.Hc+.5)?H:cutAt(y0,F.Hc,atoms);
    if(c<=y0+1)c=y0+F.Hc;
    plan.push({y0:y0,h:Math.min(c,H)-y0});
    y0=c;
  }
  if(!plan.length)plan.push({y0:0,h:1});
  return plan;
}

/* ---------- invisible text layer + links ---------- */
function textLayer(pdf,root,clipEl,stat){
  var d=root.ownerDocument,w=d.defaultView,pr=root.getBoundingClientRect(),cr=(clipEl||root).getBoundingClientRect();
  function inside(rc){return rc.width>0&&rc.height>0&&rc.top>=cr.top-2&&rc.bottom<=cr.bottom+2&&rc.left>=cr.left-2&&rc.right<=cr.right+2}
  var tw=d.createTreeWalker(root,NodeFilter.SHOW_TEXT,null),n,rg=d.createRange();
  pdf.setFont('helvetica','normal');
  while((n=tw.nextNode())){
    var t=n.nodeValue;if(!t.trim())continue;
    var el=n.parentElement;if(!el)continue;
    var cs=w.getComputedStyle(el);if(cs.visibility==='hidden'||cs.display==='none')continue;
    var fs=Math.max(1,parseFloat(cs.fontSize)*PX),re=/\S+/g,m;
    while((m=re.exec(t))){
      rg.setStart(n,m.index);rg.setEnd(n,m.index+m[0].length);
      var rc=rg.getClientRects()[0];if(!rc||!inside(rc))continue;
      var txt=m[0].replace(/[^\x20-\x7E\xA0-\xFF]/g,'');
      if(txt.length<m[0].length)stat.lost+=m[0].length-txt.length;
      if(!txt)continue;
      pdf.setFontSize(fs);
      var tw0=pdf.getTextWidth(txt),want=rc.width*PX,cs2=txt.length>1?(want-tw0)/(txt.length-1):0;
      try{pdf.text(txt+' ',(rc.left-pr.left)*PX,(rc.top-pr.top+rc.height*.8)*PX,{renderingMode:'invisible',charSpace:Math.abs(cs2)<fs?cs2:0})}catch(e){}
    }
  }
  root.querySelectorAll('a[href]').forEach(function(a){
    var h=a.getAttribute('href')||'';if(!/^(https?:|mailto:|tel:)/i.test(h))return;
    [].slice.call(a.getClientRects()).forEach(function(rc){
      if(!inside(rc))return;
      try{pdf.link((rc.left-pr.left)*PX,(rc.top-pr.top)*PX,rc.width*PX,rc.height*PX,{url:h})}catch(e){}
    });
  });
}

/* ---------- main conversion of one file ---------- */
async function shot(el,scale){
  return html2canvas(el,{scale:scale,backgroundColor:'#ffffff',logging:false,useCORS:true,imageTimeout:6000});
}
async function convert(it,opt,prog){
  var f=it.file,nm=f.name.toLowerCase(),F=frame(),d=F.contentDocument,J=window.jspdf.jsPDF,pdf=null,total=0,stat={lost:0};
  d.body.innerHTML='';d.head.querySelectorAll('style,link').forEach(function(x){x.remove()});
  var list=[]; /* {el, clip, w, h, prep} */
  if(/\.docx$/.test(nm)&&window.docx){
    await window.docx.renderAsync(await f.arrayBuffer(),d.body,d.head,{className:'docx',inWrapper:false,ignoreWidth:false,ignoreHeight:false,breakPages:true,renderHeaders:true,renderFooters:true,renderFootnotes:true,renderEndnotes:true,useBase64URL:true,trimXmlDeclaration:true});
    await settle(d);
    var secs=[].slice.call(d.body.querySelectorAll('section.docx'));
    if(!secs.length)throw new Error('empty');
    var all=[];secs.forEach(function(s){all=all.concat(splitSection(s))});
    var mw=0;all.forEach(function(s){mw=Math.max(mw,s.offsetWidth)});F.style.width=(mw+40)+'px';
    await settle(d);
    all.forEach(function(s){list.push({el:s,clip:s,w:s.offsetWidth,h:s.offsetHeight})});
  }else{
    var kind,text;
    if(/\.docx$/.test(nm)){ /* docx-preview did not load: basic fallback */
      if(!window.mammoth)throw new Error('lib');
      kind='html';text='<div>'+(await mammoth.convertToHtml({arrayBuffer:await f.arrayBuffer()})).value+'</div>';
    }else{kind=/\.txt$/.test(nm)?'txt':'html';text=await f.text()}
    var sz=SIZES[opt.size]||SIZES.a4,m=opt.margin;
    var FL=flowPages(d,kind,text,sz[0],sz[1],m);F.style.width=(sz[0]+40)+'px';
    await settle(d);
    var plan=planFlow(FL);
    plan.forEach(function(p){list.push({el:FL.box,clip:FL.clip,w:sz[0],h:sz[1],prep:function(){FL.flow.style.top=(-p.y0)+'px';FL.clip.style.height=Math.max(1,p.h)+'px'}})});
  }
  total=list.length;
  for(var i=0;i<list.length;i++){
    var pg=list[i];prog(i+1,total);
    if(pg.prep){pg.prep();await tick()}
    var wpt=pg.w*PX,hpt=pg.h*PX,or=wpt>hpt?'l':'p',c=await shot(pg.el,opt.scale);
    if(!pdf)pdf=new J({unit:'pt',format:[wpt,hpt],orientation:or,compress:true});else pdf.addPage([wpt,hpt],or);
    var jp=c.toDataURL('image/jpeg',opt.q),pn=c.toDataURL('image/png'),usePng=pn.length<jp.length*0.9; /* flat text pages are smaller as PNG; photo pages as JPEG */
    pdf.addImage(usePng?pn:jp,usePng?'PNG':'JPEG',0,0,wpt,hpt,undefined,'FAST');
    c.width=c.height=0;
    textLayer(pdf,pg.el,pg.clip,stat);
  }
  if(!pdf)throw new Error('nothing');
  var blob=pdf.output('blob');
  U.save(blob,U.base(f.name)+'.pdf');
  return {pages:total,size:blob.size,lost:stat.lost};
}

$('dGo').onclick=async function(){
  if(!D.length||!U.need('jspdf','PDF library')||!U.need('html2canvas','Page renderer'))return;
  var btn=$('dGo'),hi=$('dQ').value==='high',opt={size:$('dSize').value,margin:+$('dMargin').value,scale:hi?2.5:1.5,q:hi?.92:.82},okc=0;
  U.busy(btn,true,'Converting...');
  for(var i=0;i<D.length;i++){
    var it=D[i];
    try{
      var r=await convert(it,opt,function(p,t){btn.textContent='File '+(i+1)+'/'+D.length+' \u00b7 page '+p+'/'+t});
      it.note='Done: '+r.pages+' page'+(r.pages>1?'s':'')+' \u00b7 '+U.fmt(r.size)+(r.lost?' \u00b7 some non-English letters are not searchable':'');okc++;
    }catch(e){
      it.note='Failed';U.toast('Could not convert '+it.file.name+'. The file may be damaged or password protected.',1);
    }
    dDraw();
  }
  dropFrame();U.busy(btn,false);
  if(okc)U.toast(okc===1?'PDF is ready':okc+' PDFs are ready');
};
})();
