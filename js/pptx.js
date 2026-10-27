/* PowerPoint (.pptx) reader: turns each slide into HTML so it can be saved as PDF. */
(function(){
'use strict';
var EMU=9525,MAP={tx1:'dk1',bg1:'lt1',tx2:'dk2',bg2:'lt2'};
function K(n,name){var o=[];if(!n)return o;for(var c=n.firstChild;c;c=c.nextSibling)if(c.nodeName===name)o.push(c);return o}
function k1(n,name){return K(n,name)[0]||null}
function P(n){for(var i=1;i<arguments.length&&n;i++)n=k1(n,arguments[i]);return n||null}
function A(n,a,d){if(!n)return d;var v=n.getAttribute(a);return v===null||v===''?d:v}
function px(e){return Math.round(+e/EMU*100)/100}
function esc(s){return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}
function xml(s){return new DOMParser().parseFromString(s,'application/xml').documentElement}
function firstEl(n){for(var c=n.firstChild;c;c=c.nextSibling)if(c.nodeType===1)return c;return null}
function resolve(dir,t){var p=dir.split('/');t.split('/').forEach(function(s){if(s==='..')p.pop();else if(s!=='.'&&s!=='')p.push(s)});return p.join('/')}
async function text(zip,p){var f=zip.file(p);return f?f.async('string'):null}
async function relsOf(zip,part){
  var i=part.lastIndexOf('/'),d=part.slice(0,i),n=part.slice(i+1),s=await text(zip,d+'/_rels/'+n+'.rels'),m={};
  if(!s)return m;
  K(xml(s),'Relationship').forEach(function(e){
    var t=e.getAttribute('Target')||'';
    m[e.getAttribute('Id')]={t:e.getAttribute('Type')||'',p:e.getAttribute('TargetMode')==='External'?null:(t[0]==='/'?t.slice(1):resolve(d,t))};
  });
  return m;
}
function themeOf(root){
  var th={c:{},maj:'Calibri',min:'Calibri'};if(!root)return th;
  var cs=root.getElementsByTagName('a:clrScheme')[0];
  if(cs)for(var c=cs.firstChild;c;c=c.nextSibling){
    if(c.nodeType!==1)continue;var v=firstEl(c);
    if(v)th.c[c.nodeName.slice(2)]=v.nodeName==='a:sysClr'?A(v,'lastClr','000000'):A(v,'val','000000');
  }
  var mj=root.getElementsByTagName('a:majorFont')[0],mn=root.getElementsByTagName('a:minorFont')[0];
  th.maj=A(P(mj,'a:latin'),'typeface','Calibri');th.min=A(P(mn,'a:latin'),'typeface','Calibri');
  return th;
}
function font(tf,th){if(!tf)return null;if(tf==='+mj-lt')return th.maj;if(tf[0]==='+')return th.min;return tf}
function hsl(r,g,b,m,o){
  r/=255;g/=255;b/=255;
  var mx=Math.max(r,g,b),mn=Math.min(r,g,b),l=(mx+mn)/2,h=0,s=0,d=mx-mn;
  if(d){s=l>.5?d/(2-mx-mn):d/(mx+mn);h=mx===r?(g-b)/d+(g<b?6:0):mx===g?(b-r)/d+2:(r-g)/d+4;h/=6}
  l=Math.min(1,Math.max(0,l*m+o));
  function f(p,q,t){if(t<0)t+=1;if(t>1)t-=1;return t<1/6?p+(q-p)*6*t:t<.5?q:t<2/3?p+(q-p)*(2/3-t)*6:p}
  var R,G,B;
  if(!s){R=G=B=l}else{var q=l<.5?l*(1+s):l+s-l*s,p=2*l-q;R=f(p,q,h+1/3);G=f(p,q,h);B=f(p,q,h-1/3)}
  return [R,G,B].map(function(v){return Math.round(v*255)});
}
/* n = any node that has a colour child (solidFill, gs, fontRef ...) */
function color(n,th){
  if(!n)return null;
  var c=null;
  for(var x=n.firstChild;x;x=x.nextSibling){if(x.nodeType===1&&/^a:(srgbClr|schemeClr|sysClr|prstClr)$/.test(x.nodeName)){c=x;break}}
  if(!c)return null;
  var hex='000000',nm=c.nodeName;
  if(nm==='a:srgbClr')hex=A(c,'val','000000');
  else if(nm==='a:sysClr')hex=A(c,'lastClr','000000');
  else if(nm==='a:schemeClr'){var v=A(c,'val','tx1');hex=th.c[MAP[v]||v]||'000000'}
  var rgb=[parseInt(hex.substr(0,2),16),parseInt(hex.substr(2,2),16),parseInt(hex.substr(4,2),16)];
  var lm=k1(c,'a:lumMod'),lo=k1(c,'a:lumOff'),al=k1(c,'a:alpha');
  if(lm||lo)rgb=hsl(rgb[0],rgb[1],rgb[2],lm?+A(lm,'val',100000)/100000:1,lo?+A(lo,'val',0)/100000:0);
  if(al&&+A(al,'val',100000)<100000)return 'rgba('+rgb.join(',')+','+(+A(al,'val')/100000)+')';
  return 'rgb('+rgb.join(',')+')';
}
function fillCss(pr,th){
  if(!pr)return null;
  if(k1(pr,'a:noFill'))return 'none';
  var s=k1(pr,'a:solidFill');if(s)return color(s,th);
  var g=k1(pr,'a:gradFill'),gs=g&&P(g,'a:gsLst'),f=K(gs,'a:gs');
  if(f.length)return 'linear-gradient(135deg,'+color(f[0],th)+','+color(f[f.length-1],th)+')';
  return null;
}
function lineCss(pr,th){
  var ln=pr&&k1(pr,'a:ln');if(!ln||k1(ln,'a:noFill'))return '';
  var c=color(k1(ln,'a:solidFill'),th);if(!c)return '';
  var d=P(ln,'a:prstDash');
  return 'border:'+Math.max(1,px(A(ln,'w',12700)))+'px '+(d&&/dash|dot/i.test(A(d,'val',''))?'dashed':'solid')+' '+c+';';
}
function xf(pr,tag){
  var x=pr&&k1(pr,tag||'a:xfrm');if(!x)return null;
  var o=k1(x,'a:off'),e=k1(x,'a:ext');if(!o||!e)return null;
  return {x:+o.getAttribute('x'),y:+o.getAttribute('y'),w:+e.getAttribute('cx'),h:+e.getAttribute('cy'),rot:+A(x,'rot',0)/60000,fh:A(x,'flipH','0')==='1',fv:A(x,'flipV','0')==='1'};
}
function box(p,T){return {X:T.ax+(p.x-T.cx)*T.sx,Y:T.ay+(p.y-T.cy)*T.sy,W:p.w*T.sx,H:p.h*T.sy}}
function posCss(b,p,noFlip){
  var tr='';if(p.rot)tr+='rotate('+p.rot+'deg) ';
  if(!noFlip&&(p.fh||p.fv))tr+='scale('+(p.fh?-1:1)+','+(p.fv?-1:1)+')';
  return 'position:absolute;left:'+px(b.X)+'px;top:'+px(b.Y)+'px;width:'+px(b.W)+'px;height:'+px(b.H)+'px;box-sizing:border-box;'+(tr?'transform:'+tr+';':'');
}
function geo(pr,b){
  var g=pr&&k1(pr,'a:prstGeom');if(!g)return '';var p=A(g,'prst','rect');
  if(p==='ellipse')return 'border-radius:50%;';
  if(p==='roundRect'){var gd=P(g,'a:avLst','a:gd'),v=gd?+A(gd,'fmla','val 16667').split(' ')[1]:16667;return 'border-radius:'+px(Math.min(b.W,b.H)*v/100000)+'px;'}
  if(p==='triangle')return 'clip-path:polygon(50% 0,100% 100%,0 100%);';
  if(p==='diamond')return 'clip-path:polygon(50% 0,100% 50%,50% 100%,0 50%);';
  return '';
}
function phKind(ph){var t=A(ph,'type','body');return /title/i.test(t)?'title':/^(dt|ftr|sldNum)$/.test(t)?'other':'body'}
function matchPh(tree,ph,useIdx){
  if(!tree)return null;
  var idx=A(ph,'idx',null),kind=phKind(ph),first=null,sps=K(tree,'p:sp');
  for(var i=0;i<sps.length;i++){
    var p=P(sps[i],'p:nvSpPr','p:nvPr','p:ph');if(!p)continue;
    if(useIdx&&idx!==null&&A(p,'idx',null)===idx)return sps[i];
    if(!first&&phKind(p)===kind&&(kind!=='other'||A(p,'type','')===A(ph,'type','')))first=sps[i];
  }
  return first;
}
/* text defaults from layout, then master, for one paragraph level */
function dstyle(ph,lvl,ctx){
  var tag='a:lvl'+(lvl+1)+'pPr',nodes=[];
  if(ph){
    [ctx.layout&&matchPh(ctx.layout.tree,ph,true),ctx.master&&matchPh(ctx.master.tree,ph,false)].forEach(function(sp){
      var l=sp&&P(sp,'p:txBody','a:lstStyle',tag);if(l)nodes.push(l);
    });
  }
  var sn=ph?{title:'p:titleStyle',body:'p:bodyStyle',other:'p:otherStyle'}[phKind(ph)]:'p:otherStyle',ts=ctx.master&&P(ctx.master.root,'p:txStyles',sn,tag);
  if(ts)nodes.push(ts);
  var r={};
  nodes.forEach(function(l){
    var d=k1(l,'a:defRPr');
    if(r.algn===undefined&&A(l,'algn',null))r.algn=A(l,'algn');
    if(r.marL===undefined&&A(l,'marL',null)!==null)r.marL=+A(l,'marL');
    if(r.ind===undefined&&A(l,'indent',null)!==null)r.ind=+A(l,'indent');
    if(r.bu===undefined){var bc=k1(l,'a:buChar');if(k1(l,'a:buNone'))r.bu='none';else if(bc)r.bu=A(bc,'char','\u2022');else if(k1(l,'a:buAutoNum'))r.bu='#'}
    if(r.lsp===undefined){var ls=P(l,'a:lnSpc','a:spcPct');if(ls)r.lsp=+A(ls,'val',100000)/100000}
    if(r.sb===undefined){var sb=P(l,'a:spcBef','a:spcPts');if(sb)r.sb=+A(sb,'val',0)/100}
    if(d){
      if(r.sz===undefined&&A(d,'sz',null))r.sz=+A(d,'sz');
      if(!r.color){var c=color(k1(d,'a:solidFill'),ctx.th);if(c)r.color=c}
      if(!r.font){var f=font(A(P(d,'a:latin'),'typeface',null),ctx.th);if(f)r.font=f}
    }
  });
  return r;
}
function parasHtml(tb,o){
  var th=o.ctx.th,out='',cnt={};
  K(tb,'a:p').forEach(function(p){
    var pr=k1(p,'a:pPr'),lvl=+A(pr,'lvl',0),d=dstyle(o.ph,lvl,o.ctx);
    var algn=A(pr,'algn',d.algn||'l'),ta={l:'left',ctr:'center',r:'right',just:'justify'}[algn]||'left';
    var bu=k1(pr,'a:buNone')?'none':k1(pr,'a:buChar')?A(k1(pr,'a:buChar'),'char','\u2022'):k1(pr,'a:buAutoNum')?'#':(d.bu!==undefined?d.bu:'none');
    var marL=A(pr,'marL',null)!==null?+A(pr,'marL'):(d.marL!==undefined?d.marL:(bu!=='none'?(lvl+1)*342900:0));
    var ind=A(pr,'indent',null)!==null?+A(pr,'indent'):(d.ind!==undefined?d.ind:(bu!=='none'?-285750:0));
    var sb=P(pr,'a:spcBef','a:spcPts'),sa=P(pr,'a:spcAft','a:spcPts'),ls=P(pr,'a:lnSpc','a:spcPct');
    var lh=1.2*(ls?+A(ls,'val',100000)/100000:(d.lsp||1)),mt=sb?+A(sb,'val',0)/100:(d.sb||0),mb=sa?+A(sa,'val',0)/100:0;
    var runs='',lastSz=null;
    for(var c=p.firstChild;c;c=c.nextSibling){
      if(c.nodeName==='a:br'){runs+='<br>';continue}
      if(c.nodeName!=='a:r'&&c.nodeName!=='a:fld')continue;
      var rp=k1(c,'a:rPr'),t=k1(c,'a:t'),tx=t?t.textContent:'';
      var sz=(+A(rp,'sz',d.sz||o.defSz||1800))*o.scale,col=color(k1(rp,'a:solidFill'),th)||o.defColor||d.color||'#000';
      var fam=font(A(P(rp,'a:latin'),'typeface',null),th)||d.font||th.min;
      var st='font-size:'+Math.round(sz*1.3333)/100+'px;color:'+col+';font-family:\''+fam+'\',Arial,sans-serif;';
      if(A(rp,'b','0')==='1')st+='font-weight:bold;';
      if(A(rp,'i','0')==='1')st+='font-style:italic;';
      var u=A(rp,'u','none');if(u!=='none')st+='text-decoration:underline;';
      var bl=+A(rp,'baseline',0);if(bl>0)st+='vertical-align:super;font-size:70%;';else if(bl<0)st+='vertical-align:sub;font-size:70%;';
      runs+='<span style="'+st+'">'+esc(tx)+'</span>';lastSz=sz;
    }
    if(!runs){var es=+A(k1(p,'a:endParaRPr'),'sz',d.sz||o.defSz||1800)*o.scale;runs='<span style="font-size:'+Math.round(es*1.3333)/100+'px">&nbsp;</span>'}
    var label='';
    if(bu!=='none'&&runs.indexOf('&nbsp;</span>')<0){
      if(bu==='#'){cnt[lvl]=(cnt[lvl]||0)+1;label=cnt[lvl]+'.'}else{label=bu;cnt[lvl]=0}
      label='<span style="display:inline-block;width:'+(ind<0?px(-ind):0)+'px;text-indent:0;font-size:'+Math.round((lastSz||(d.sz||o.defSz||1800)*o.scale)*1.3333)/100+'px">'+esc(label)+'</span>';
    }else cnt[lvl]=0;
    out+='<div style="text-align:'+ta+';padding-left:'+px(marL)+'px;text-indent:'+px(ind)+'px;margin-top:'+Math.round(mt*1.3333*100)/100+'px;margin-bottom:'+Math.round(mb*1.3333*100)/100+'px;line-height:'+lh+'">'+label+runs+'</div>';
  });
  return out;
}
function textBox(sp,ph,ctx,defColor){
  var tb=k1(sp,'p:txBody');if(!tb||!tb.textContent.trim())return '';
  var bps=[k1(tb,'a:bodyPr')];
  if(ph){
    var l=matchPh(ctx.layout&&ctx.layout.tree,ph,true),m=matchPh(ctx.master&&ctx.master.tree,ph,false);
    bps.push(l&&P(l,'p:txBody','a:bodyPr'));bps.push(m&&P(m,'p:txBody','a:bodyPr'));
  }
  function bp(a,d){for(var i=0;i<bps.length;i++){var v=A(bps[i],a,null);if(v!==null)return v}return d}
  var jc={t:'flex-start',ctr:'center',b:'flex-end'}[bp('anchor','t')]||'flex-start';
  var na=bps[0]&&P(bps[0],'a:normAutofit'),fs=na&&A(na,'fontScale',null)?+A(na,'fontScale')/100000:1;
  var inner=parasHtml(tb,{ph:ph,ctx:ctx,scale:fs,defColor:defColor});
  return '<div style="position:absolute;left:0;top:0;width:100%;height:100%;box-sizing:border-box;padding:'+px(bp('tIns',45720))+'px '+px(bp('rIns',91440))+'px '+px(bp('bIns',45720))+'px '+px(bp('lIns',91440))+'px;display:flex;flex-direction:column;justify-content:'+jc+';'+(bp('wrap','square')==='none'?'white-space:nowrap;':'')+'">'+inner+'</div>';
}
async function shape(sp,ctx,T,skipPh){
  var ph=P(sp,'p:nvSpPr','p:nvPr','p:ph');if(ph&&skipPh)return '';
  var pr=k1(sp,'p:spPr'),pos=xf(pr);
  if(!pos&&ph){
    var l=matchPh(ctx.layout&&ctx.layout.tree,ph,true);pos=l&&xf(k1(l,'p:spPr'));
    if(!pos){var m=matchPh(ctx.master&&ctx.master.tree,ph,false);pos=m&&xf(k1(m,'p:spPr'))}
  }
  if(!pos)return '';
  var b=box(pos,T),st=k1(sp,'p:style'),fill=fillCss(pr,ctx.th),ln=lineCss(pr,ctx.th);
  if(fill===null&&st){var fr=k1(st,'a:fillRef');if(fr&&A(fr,'idx','0')!=='0')fill=color(fr,ctx.th)}
  if(!ln&&st&&pr&&!k1(pr,'a:ln')){var lr=k1(st,'a:lnRef');if(lr&&A(lr,'idx','0')!=='0'){var lc=color(lr,ctx.th);if(lc)ln='border:1px solid '+lc+';'}}
  var fr2=P(sp,'p:style','a:fontRef'),defCol=fr2?color(fr2,ctx.th):null,hasTx=!!(k1(sp,'p:txBody')&&k1(sp,'p:txBody').textContent.trim());
  return '<div style="'+posCss(b,pos,hasTx)+(fill&&fill!=='none'?'background:'+fill+';':'')+ln+geo(pr,b)+'">'+textBox(sp,ph,ctx,defCol)+'</div>';
}
function mimeOf(p){return {png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',gif:'image/gif',svg:'image/svg+xml'}[p.split('.').pop().toLowerCase()]}
async function pic(p,ctx,T){
  var bf=k1(p,'p:blipFill'),bl=bf&&k1(bf,'a:blip'),id=bl&&bl.getAttribute('r:embed'),rel=id&&ctx.rels[id],pr=k1(p,'p:spPr'),pos=xf(pr);
  if(!pos||!rel||!rel.p)return '';
  var mime=mimeOf(rel.p),f=ctx.zip.file(rel.p);
  if(!mime||!f){ctx.stats.skip++;return ''}
  var data=await f.async('base64'),b=box(pos,T),sr=k1(bf,'a:srcRect');
  var l=+A(sr,'l',0)/1e5,t=+A(sr,'t',0)/1e5,r=+A(sr,'r',0)/1e5,bb=+A(sr,'b',0)/1e5,iw=1-l-r,ih=1-t-bb;
  if(iw<.01||ih<.01){l=t=r=bb=0;iw=ih=1}
  return '<div style="'+posCss(b,pos)+'overflow:hidden;'+geo(pr,b)+'"><img src="data:'+mime+';base64,'+data+'" style="position:absolute;width:'+(100/iw)+'%;height:'+(100/ih)+'%;left:-'+(l/iw*100)+'%;top:-'+(t/ih*100)+'%"></div>';
}
function cxn(n,ctx,T){
  var pr=k1(n,'p:spPr'),pos=xf(pr);if(!pos)return '';
  var b=box(pos,T),ln=pr&&k1(pr,'a:ln'),c=color(ln&&k1(ln,'a:solidFill'),ctx.th)||'#444',w=Math.max(1,px(A(ln,'w',12700))),W=px(b.W),H=px(b.H);
  return '<svg style="'+posCss(b,pos)+'overflow:visible" width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'"><line x1="0" y1="0" x2="'+W+'" y2="'+H+'" stroke="'+c+'" stroke-width="'+w+'"/></svg>';
}
async function group(g,ctx,T,skipPh){
  var pr=k1(g,'p:grpSpPr'),x=pr&&k1(pr,'a:xfrm');if(!x)return '';
  var o=k1(x,'a:off'),e=k1(x,'a:ext'),co=k1(x,'a:chOff'),ce=k1(x,'a:chExt');if(!o||!e||!co||!ce)return '';
  var ccx=+ce.getAttribute('cx')||1,ccy=+ce.getAttribute('cy')||1;
  return tree(g,ctx,{ax:T.ax+(+o.getAttribute('x')-T.cx)*T.sx,ay:T.ay+(+o.getAttribute('y')-T.cy)*T.sy,sx:+e.getAttribute('cx')*T.sx/ccx,sy:+e.getAttribute('cy')*T.sy/ccy,cx:+co.getAttribute('x'),cy:+co.getAttribute('y')},skipPh);
}
async function frame(gf,ctx,T){
  var pos=xf(gf,'p:xfrm');if(!pos)return '';
  var b=box(pos,T),tbl=gf.getElementsByTagName('a:tbl')[0];
  if(!tbl){ctx.stats.skip++;return '<div style="'+posCss(b,pos)+'border:2px dashed #bbb;color:#888;display:flex;align-items:center;justify-content:center;font:16px Arial,sans-serif">Chart or diagram (not supported)</div>'}
  var cols=K(P(tbl,'a:tblGrid'),'a:gridCol'),tp=k1(tbl,'a:tblPr'),head=A(tp,'firstRow','0')==='1',band=A(tp,'bandRow','0')==='1',acc='rgb('+[0,2,4].map(function(i){return parseInt((ctx.th.c.accent1||'4472C4').substr(i,2),16)}).join(',')+')';
  var h='<table style="position:absolute;left:'+px(b.X)+'px;top:'+px(b.Y)+'px;width:'+px(b.W)+'px;border-collapse:collapse;table-layout:fixed"><colgroup>'+cols.map(function(c){return '<col style="width:'+px(+A(c,'w',914400)*T.sx)+'px">'}).join('')+'</colgroup>';
  K(tbl,'a:tr').forEach(function(tr,ri){
    h+='<tr style="height:'+px(+A(tr,'h',370840)*T.sy)+'px">';
    K(tr,'a:tc').forEach(function(tc){
      if(A(tc,'hMerge','0')==='1'||A(tc,'vMerge','0')==='1')return;
      var cp=k1(tc,'a:tcPr'),bg=fillCss(cp,ctx.th),hd=head&&ri===0;
      if(!bg||bg==='none')bg=hd?acc:(band&&ri%2===1?'#eef1f8':'transparent');
      var inner=parasHtml(k1(tc,'a:txBody'),{ph:null,ctx:ctx,scale:1,defColor:hd?'#fff':'#000'});
      h+='<td colspan="'+A(tc,'gridSpan',1)+'" rowspan="'+A(tc,'rowSpan',1)+'" style="background:'+bg+';border:1px solid #cfd4e0;padding:4px 9px;vertical-align:middle;overflow:hidden'+(hd?';font-weight:bold':'')+'">'+inner+'</td>';
    });
    h+='</tr>';
  });
  return h+'</table>';
}
async function tree(tr,ctx,T,skipPh){
  var out='';
  for(var n=tr.firstChild;n;n=n.nextSibling){
    if(n.nodeType!==1)continue;
    try{
      var nm=n.nodeName;
      if(nm==='p:sp')out+=await shape(n,ctx,T,skipPh);
      else if(nm==='p:pic')out+=await pic(n,ctx,T);
      else if(nm==='p:cxnSp')out+=cxn(n,ctx,T);
      else if(nm==='p:grpSp')out+=await group(n,ctx,T,skipPh);
      else if(nm==='p:graphicFrame')out+=await frame(n,ctx,T);
    }catch(e){ctx.stats.err++}
  }
  return out;
}
async function bgOf(root,ctx,rels){
  var b=P(root,'p:cSld','p:bg');if(!b)return null;
  var pr=k1(b,'p:bgPr');
  if(pr){
    var f=fillCss(pr,ctx.th);if(f&&f!=='none')return 'background:'+f+';';
    var bl=k1(pr,'a:blipFill'),bb=bl&&k1(bl,'a:blip'),id=bb&&bb.getAttribute('r:embed'),rel=id&&rels[id];
    if(rel&&rel.p&&mimeOf(rel.p)&&ctx.zip.file(rel.p))return 'background:url(data:'+mimeOf(rel.p)+';base64,'+await ctx.zip.file(rel.p).async('base64')+') center/cover no-repeat;';
  }
  var ref=k1(b,'p:bgRef'),c=ref&&color(ref,ctx.th);
  return c?'background:'+c+';':null;
}
async function parse(buf){
  var zip=await JSZip.loadAsync(buf),ps=await text(zip,'ppt/presentation.xml');
  if(!ps)throw new Error('Not a PowerPoint file');
  var pr=xml(ps),sz=P(pr,'p:sldSz'),W=+A(sz,'cx',9144000),H=+A(sz,'cy',6858000),prel=await relsOf(zip,'ppt/presentation.xml');
  var cache={},stats={skip:0,err:0,hidden:0},paths=[];
  async function part(p){
    if(cache[p])return cache[p];var s=await text(zip,p);if(!s)return null;
    var root=xml(s);return cache[p]={root:root,tree:P(root,'p:cSld','p:spTree'),rels:await relsOf(zip,p)};
  }
  function relOf(rels,re){for(var k in rels)if(re.test(rels[k].t))return rels[k];return null}
  K(P(pr,'p:sldIdLst'),'p:sldId').forEach(function(s){var r=prel[s.getAttribute('r:id')];if(r&&r.p)paths.push(r.p)});
  var live=[];
  for(var i=0;i<paths.length;i++){var s0=await part(paths[i]);if(!s0)continue;if(A(s0.root,'show','1')==='0')stats.hidden++;else live.push(paths[i])}
  var T0={ax:0,ay:0,sx:1,sy:1,cx:0,cy:0};
  return {count:live.length,w:px(W),h:px(H),stats:stats,
    render:async function(i){
      var sl=await part(live[i]),lr=relOf(sl.rels,/slideLayout$/),layout=lr&&lr.p?await part(lr.p):null;
      var mr=layout&&relOf(layout.rels,/slideMaster$/),master=mr&&mr.p?await part(mr.p):null;
      var tr=master&&relOf(master.rels,/theme$/),tx=tr&&tr.p?await part(tr.p):null;
      var base={zip:zip,th:themeOf(tx&&tx.root),layout:layout,master:master,stats:stats};
      var cs=Object.assign({},base,{rels:sl.rels}),cl=Object.assign({},base,{rels:layout?layout.rels:{}}),cm=Object.assign({},base,{rels:master?master.rels:{}});
      var bg=(await bgOf(sl.root,cs,sl.rels))||(layout&&await bgOf(layout.root,cl,layout.rels))||(master&&await bgOf(master.root,cm,master.rels))||'background:#fff;';
      var html='';
      if(A(sl.root,'showMasterSp','1')!=='0'){
        if(master&&(!layout||A(layout.root,'showMasterSp','1')!=='0'))html+=await tree(master.tree,cm,T0,true);
        if(layout)html+=await tree(layout.tree,cl,T0,true);
      }
      html+=await tree(sl.tree,cs,T0,false);
      return '<div style="position:relative;overflow:hidden;width:'+px(W)+'px;height:'+px(H)+'px;'+bg+'">'+html+'</div>';
    }};
}
window.PPTX={parse:parse};

/* ---------- Page: PowerPoint to PDF ---------- */
var U=window.JU;if(!U||!document.getElementById('oDrop'))return;
var $=U.$,O=null;
U.dz($('oDrop'),$('oFile'),function(fl){
  var f=fl[0];if(!f)return;
  if(/\.ppt$/i.test(f.name))return U.toast('Old .ppt files are not supported. Open it in PowerPoint and save as .pptx first.',1);
  if(!/\.pptx$/i.test(f.name))return U.toast('Please choose a .pptx file.',1);
  O=f;$('oInfo').innerHTML='<b>'+U.esc(f.name)+'</b> \u00b7 '+U.fmt(f.size);$('oRes').textContent='';U.show('oDrop',false);U.show('oTool',true);
});
$('oClear').onclick=function(){O=null;U.show('oTool',false);U.show('oDrop',true)};
$('oGo').onclick=async function(){
  if(!O||!U.need('JSZip','ZIP library')||!U.need('html2canvas','Page renderer')||!U.need('jspdf','PDF library'))return;
  var btn=$('oGo');U.busy(btn,true,'Reading...');
  try{
    var deck=await parse(await O.arrayBuffer());
    if(!deck.count){U.toast('No slides found in this file.',1);U.busy(btn,false);return}
    var box=$('slideBox'),J=window.jspdf.jsPDF,wp=deck.w*.75,hp=deck.h*.75,or=wp>=hp?'l':'p',doc=null;
    box.style.width=deck.w+'px';box.style.height=deck.h+'px';
    for(var i=0;i<deck.count;i++){
      btn.textContent='Slide '+(i+1)+' of '+deck.count;
      box.innerHTML=await deck.render(i);
      await new Promise(function(r){setTimeout(r,60)});
      var c=await html2canvas(box,{scale:Math.min(2,2200/deck.w),backgroundColor:'#ffffff',logging:false,ignoreElements:function(e){return e.id==='prog'||e.id==='toast'||e.classList&&(e.classList.contains('top')||e.classList.contains('bg'))}});
      if(!doc)doc=new J({unit:'pt',format:[wp,hp],orientation:or});else doc.addPage([wp,hp],or);
      doc.addImage(c.toDataURL('image/jpeg',0.92),'JPEG',0,0,wp,hp);
      box.innerHTML='';
    }
    doc.save(O.name.replace(/\.[^.]+$/,'')+'.pdf');
    var st=deck.stats,notes=[];
    if(st.skip)notes.push(st.skip+' chart/diagram/picture item'+(st.skip>1?'s were':' was')+' not drawn');
    if(st.hidden)notes.push(st.hidden+' hidden slide'+(st.hidden>1?'s were':' was')+' skipped');
    $('oRes').textContent='Done: '+deck.count+' slide'+(deck.count>1?'s':'')+' saved.'+(notes.length?' Note: '+notes.join('; ')+'.':'');
  }catch(e){U.toast('Could not convert this file. It may be damaged or password protected.',1)}
  $('slideBox').innerHTML='';U.busy(btn,false);
};
})();
