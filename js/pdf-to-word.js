/* PDF to Word (.docx)
   Reads the text, fonts and drawing commands of every PDF page with pdf.js and rebuilds the layout as a real,
   editable Word file: headings, paragraphs, bullet and numbered lists, tables (with borders and shading),
   bold / italic / code text, highlighted text, links and pictures. Runs fully in the browser.
   The core (convert) has no browser dependency, so it can also be tested in Node. */
(function(root){
'use strict';

/* ---------------------------------------------------------------- small helpers */
function esc(s){return String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g,'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function h2(n){n=Math.max(0,Math.min(255,Math.round(n)));return (n<16?'0':'')+n.toString(16)}
function hexOf(rgb){return (h2(rgb[0])+h2(rgb[1])+h2(rgb[2])).toUpperCase()}
function lumOf(rgb){return (0.2126*rgb[0]+0.7152*rgb[1]+0.0722*rgb[2])/255}
function clamp(v,a,b){return Math.max(a,Math.min(b,v))}
function pct(a,p){if(!a.length)return 0;a=a.slice().sort(function(x,y){return x-y});return a[Math.min(a.length-1,Math.max(0,Math.floor(p*a.length)))]}
function r05(n){return Math.round(n*2)/2}
var MONO=/mono|courier|consolas|menlo|lucida\s?console|source\s?code|fira\s?code|cascadia|inconsolata|typewriter/i;

/* letter-spaced text such as "i n d e x . h t m l" -> "index.html" */
function despace(str){
  if(str.indexOf(' ')<0)return str;
  var tk=str.split(' '),out=[],i=0;
  var core=str.trim().split(' ');
  if(core.length>=2&&core.every(function(t){return t.length===1})&&!core.every(function(t){return /\d/.test(t)})&&!/^[AI]$/.test(core[0]+''))return str.replace(core.join(' '),core.join(''));
  while(i<tk.length){
    var j=i;while(j<tk.length&&tk[j].length===1)j++;
    if(j-i>=4&&!tk.slice(i,j).every(function(t){return /\d/.test(t)})){out.push(tk.slice(i,j).join(''));i=j}
    else{out.push(tk[i]);i++}
  }
  return out.join(' ');
}

/* ---------------------------------------------------------------- fonts */
function famName(real,generic,mono){
  var n=String(real||'').replace(/^[A-Z]{6}\+/,'').split(/[-,]/)[0].replace(/(PSMT|MT|PS)$/,'');
  if(!n||/^g_d\d|^f\d+$/i.test(n)){return mono?'Consolas':(generic==='serif'?'Times New Roman':'Arial')}
  if(/helvetica|^arial/i.test(n))return 'Arial';
  if(/^times/i.test(n))return 'Times New Roman';
  if(/^courier/i.test(n))return 'Courier New';
  if(/^symbol|dingbat|wingding/i.test(n))return mono?'Consolas':'Arial';
  return n.replace(/([a-z])([A-Z])/g,'$1 $2');
}
function fontInfo(page,id,styles,cache){
  if(cache[id])return cache[id];
  var real='',f=null,gen=(styles[id]&&styles[id].fontFamily)||'';
  try{f=page.commonObjs.get(id);real=(f&&f.name)||''}catch(e){}
  var low=(real||'').toLowerCase();
  var o={
    b:/bold|black|heavy|semibold|demi/.test(low)||!!(f&&(f.bold||f.black)),
    i:/italic|oblique/.test(low)||!!(f&&f.italic),
    m:MONO.test(real)||gen==='monospace'
  };
  o.f=famName(real,gen,o.m);
  return cache[id]=o;
}

/* ---------------------------------------------------------------- graphics: lines, filled boxes, images */
function mul(a,b){return [a[0]*b[0]+a[1]*b[2],a[0]*b[1]+a[1]*b[3],a[2]*b[0]+a[3]*b[2],a[2]*b[1]+a[3]*b[3],a[4]*b[0]+a[5]*b[2]+b[4],a[4]*b[1]+a[5]*b[3]+b[5]]}
function ap(m,x,y){return [x*m[0]+y*m[2]+m[4],x*m[1]+y*m[3]+m[5]]}
function rgbArg(a){
  if(!a)return null;
  if(typeof a[0]==='string'){var s=a[0].replace('#','');return [parseInt(s.substr(0,2),16),parseInt(s.substr(2,2),16),parseInt(s.substr(4,2),16)]}
  return [a[0],a[1],a[2]];
}
function scanOps(P,ol,ox,oy){
  var O=P.OPS,fn=ol.fnArray,ar=ol.argsArray,ctm=[1,0,0,1,0,0],stack=[],fill=[0,0,0];
  var segs=[],fills=[],imgs=[],tm=0,visN=0,invN=0;
  function addSeg(x0,y0,x1,y1){
    if(Math.abs(y1-y0)<1.2&&Math.abs(x1-x0)>6)segs.push({h:1,x0:Math.min(x0,x1),x1:Math.max(x0,x1),y:(y0+y1)/2});
    else if(Math.abs(x1-x0)<1.2&&Math.abs(y1-y0)>6)segs.push({h:0,y0:Math.min(y0,y1),y1:Math.max(y0,y1),x:(x0+x1)/2});
  }
  for(var i=0;i<fn.length;i++){
    var op=fn[i],a=ar[i];
    if(op===O.save)stack.push([ctm.slice(),fill.slice()]);
    else if(op===O.restore){var s=stack.pop();if(s){ctm=s[0];fill=s[1]}}
    else if(op===O.transform)ctm=mul(a,ctm);
    else if(op===O.setFillRGBColor){var c=rgbArg(a);if(c)fill=c}
    else if(op===O.setTextRenderingMode)tm=a[0];
    else if(op===O.showText||op===O.showSpacedText||op===O.nextLineShowText||op===O.nextLineSetSpacingShowText){if(tm===3)invN++;else visN++}
    else if(op===O.constructPath){
      var nx=fn[i+1],doFill=(nx===O.fill||nx===O.eoFill||nx===O.fillStroke||nx===O.eoFillStroke||nx===O.closeFillStroke||nx===O.closeEOFillStroke);
      var doStroke=(nx===O.stroke||nx===O.closeStroke||nx===O.fillStroke||nx===O.eoFillStroke||nx===O.closeFillStroke||nx===O.closeEOFillStroke);
      if(!doFill&&!doStroke)continue;
      var ops=a[0],co=a[1],k=0,cur=null,start=null;
      for(var j=0;j<ops.length;j++){
        var p=ops[j];
        if(p===O.moveTo){cur=ap(ctm,co[k],co[k+1]);start=cur;k+=2}
        else if(p===O.lineTo){var q=ap(ctm,co[k],co[k+1]);k+=2;if(cur&&doStroke)addSeg(cur[0]-ox,cur[1]-oy,q[0]-ox,q[1]-oy);cur=q}
        else if(p===O.curveTo)k+=6;
        else if(p===O.curveTo2||p===O.curveTo3)k+=4;
        else if(p===O.closePath){if(cur&&start&&doStroke)addSeg(cur[0]-ox,cur[1]-oy,start[0]-ox,start[1]-oy);cur=start}
        else if(p===O.rectangle){
          var x=co[k],y=co[k+1],w=co[k+2],h=co[k+3];k+=4;
          var p1=ap(ctm,x,y),p2=ap(ctm,x+w,y+h),X0=Math.min(p1[0],p2[0])-ox,X1=Math.max(p1[0],p2[0])-ox,Y0=Math.min(p1[1],p2[1])-oy,Y1=Math.max(p1[1],p2[1])-oy;
          if(doFill){
            if(X1-X0<1.8||Y1-Y0<1.8){addSeg(X0,Y0,X1,Y0);addSeg(X0,Y1,X1,Y1);addSeg(X0,Y0,X0,Y1);addSeg(X1,Y0,X1,Y1);addSeg((X0+X1)/2,Y0,(X0+X1)/2,Y1);addSeg(X0,(Y0+Y1)/2,X1,(Y0+Y1)/2)}
            else fills.push({x0:X0,x1:X1,y0:Y0,y1:Y1,rgb:fill.slice()});
          }
          if(doStroke){addSeg(X0,Y0,X1,Y0);addSeg(X0,Y1,X1,Y1);addSeg(X0,Y0,X0,Y1);addSeg(X1,Y0,X1,Y1)}
          cur=null;
        }
      }
    }
    else if(op===O.paintImageXObject||op===O.paintInlineImageXObject||op===O.paintJpegXObject){
      var c1=ap(ctm,0,0),c2=ap(ctm,1,1);
      imgs.push({x0:Math.min(c1[0],c2[0])-ox,x1:Math.max(c1[0],c2[0])-ox,y0:Math.min(c1[1],c2[1])-oy,y1:Math.max(c1[1],c2[1])-oy});
    }
  }
  return {segs:segs,fills:fills,imgs:imgs,inv:invN>visN};
}

/* ---------------------------------------------------------------- one PDF page -> lines */
var BULLET=/^([\u2022\u25CF\u25E6\u25AA\u25AB\u2023\u2043\u25A0\u25A1\u2219\u00B7\u25CB\u25C6\u25C7\u27A2\u27A4\u2713\uF0B7\uF0A7\uF0D8])\s*/;
var DASHB=/^([\u2013\u2014\-\*])\s+(?=\S)/;
var NUMB=/^(\d{1,3})([.)])\s+(?=\S)/;
var LETB=/^\(?([a-z])\)\s+(?=\S)/;
function markerOf(text){
  var m=BULLET.exec(text)||DASHB.exec(text);
  if(m)return {kind:'b',len:m[0].length,ch:m[1]};
  m=NUMB.exec(text);if(m)return {kind:'n',len:m[0].length,num:+m[1],sep:m[2]};
  m=LETB.exec(text);if(m)return {kind:'n',len:m[0].length,num:m[1].charCodeAt(0)-96,sep:')',letter:1};
  return null;
}
function plain(items){return items.map(function(i){return i.s}).join('')}


/* ---------------------------------------------------------------- pages that are a picture + invisible text (OCR / "print to PDF" style)
   Font names, ruling lines and bullets are not in the PDF data there, so they are read from the page picture. */
function rasterStep(R,W,H,lines,items,segsOut){
  var s=R.s,PW=R.w,PH=R.h,d=R.data,n=PW*PH;
  var gray=new Uint8Array(n),dark=new Uint8Array(n),i;
  for(i=0;i<n;i++){var o=i*4,l=0.2126*d[o]+0.7152*d[o+1]+0.0722*d[o+2];gray[i]=l;dark[i]=l<130?1:0}
  function px(x){return Math.max(0,Math.min(PW-1,Math.round(x*s)))}
  function py(y){return Math.max(0,Math.min(PH-1,Math.round((H-y)*s)))}
  function rgbAt(x,y){var o=(py(y)*PW+px(x))*4;return [d[o],d[o+1],d[o+2]]}
  R.rgbAt=rgbAt;
  /* ruling lines */
  var minLen=Math.round(20*s),hr=[],vr=[],x,y,run;
  for(y=0;y<PH;y++){run=0;for(x=0;x<=PW;x++){if(x<PW&&dark[y*PW+x])run++;else{if(run>=minLen)hr.push({c:y,a:x-run,b:x});run=0}}}
  for(x=0;x<PW;x++){run=0;for(y=0;y<=PH;y++){if(y<PH&&dark[y*PW+x])run++;else{if(run>=minLen)vr.push({c:x,a:y-run,b:y});run=0}}}
  function merge(rs){
    var out=[];
    rs.forEach(function(r){
      for(var k=0;k<out.length;k++){
        var o=out[k];if(Math.abs(o.c-r.c)>3)continue;
        var ov=Math.min(o.b,r.b)-Math.max(o.a,r.a);
        if(ov>0.7*Math.min(o.b-o.a,r.b-r.a)){o.a=Math.min(o.a,r.a);o.b=Math.max(o.b,r.b);o.c=(o.c+r.c)/2;return}
      }
      out.push({c:r.c,a:r.a,b:r.b});
    });
    return out;
  }
  merge(hr).forEach(function(r){segsOut.push({h:1,x0:r.a/s,x1:r.b/s,y:H-(r.c+0.5)/s})});
  merge(vr).forEach(function(r){segsOut.push({h:0,y0:H-r.b/s,y1:H-r.a/s,x:(r.c+0.5)/s})});
  /* body size of this page */
  var cnt={},best=0,body=11;
  lines.forEach(function(L){var k=r05(L.size);cnt[k]=(cnt[k]||0)+L.text.length;if(cnt[k]>best){best=cnt[k];body=k}});
  /* ink density of each word group, highlighted (code) boxes */
  function darkIn(x0,y0,x1,y1){
    var a=px(x0),b=py(y1),c=px(x1),e=py(y0),k=0,t=0,u,v;
    for(v=b;v<=e;v++)for(u=a;u<=c;u++){t++;k+=dark[v*PW+u]}
    return t?k/t:0;
  }
  function stroke(x0,y0,x1,y1,size){
    var a=px(x0),c=px(x1),v0=py(y1),v1=py(y0),tot=0,cnt=0,u,v,run;
    for(v=v0;v<=v1;v++){run=0;for(u=a;u<=c+1;u++){if(u<=c&&gray[v*PW+u]<140)run++;else{if(run){tot+=run;cnt++}run=0}}}
    return cnt>=3?(tot/cnt)/size:null;
  }
  var dens=[];
  items.forEach(function(it){
    if(it.ws)return;
    var len=it.s.trim().length;
    /* a light grey box behind the text = highlighted (code) text */
    if(it.w>3){
      var ua=px(it.x),ub=px(it.x+it.w),va=py(it.y+it.size*0.86),vb2=py(it.y-it.size*0.2),tn=0,wh=0,sr=0,sg=0,sb=0,u2,v2;
      for(v2=va;v2<=vb2;v2++)for(u2=ua;u2<=ub;u2++){
        var gg=gray[v2*PW+u2];
        if(gg>=236&&gg<=249){tn++;var oo=(v2*PW+u2)*4;sr+=d[oo];sg+=d[oo+1];sb+=d[oo+2]}
        else if(gg>=251)wh++;
      }
      if(tn+wh>12&&tn/(tn+wh)>0.7){it.sh=hexOf([sr/tn,sg/tn,sb/tn]);it.mr=true;it.f0=it.f;it.m=true;it.f='Consolas'}
    }
  });
  /* bold: measure ink word by word, then split text groups where the weight changes */
  function wOf(ch){return /[iljt.,:;'!|()\[\]\s]/.test(ch)?0.55:/[mwMW@]/.test(ch)?1.5:/[A-Z0-9]/.test(ch)?1.1:1}
  var wd=[],A_LEFT=(function(){var xs0=lines.filter(function(L){return L.text.length>25}).map(function(L){return L.x0});return pct(xs0.length?xs0:lines.map(function(L){return L.x0}),0.05)})();
  items.forEach(function(it){
    it.words=null;
    if(it.ws||it.m)return;
    var str=it.s,tot=0,k;for(k=0;k<str.length;k++)tot+=wOf(str[k]);
    var ws=[],re=/\S+/g,mm,acc=0,last=0;
    while((mm=re.exec(str))){
      var a0=0;for(k=last;k<mm.index;k++)a0+=wOf(str[k]);acc+=a0;
      var aw=0;for(k=mm.index;k<mm.index+mm[0].length;k++)aw+=wOf(str[k]);
      var wx0=it.x+it.w*(acc/tot),wx1=it.x+it.w*((acc+aw)/tot);
      var pad=(wx1-wx0)*0.08;
      var dn2=mm[0].length>=3?stroke(wx0+pad,it.y+it.size*0.05,wx1-pad,it.y+it.size*0.5,it.size):null;
      ws.push({s:mm[0],i0:mm.index,x0:wx0,x1:wx1,dn:dn2});
      acc+=aw;last=mm.index+mm[0].length;
    }
    it.words=ws;
    if(Math.abs(it.size-body)<=0.8)ws.forEach(function(w){if(w.dn!==null&&w.s.length>=4)wd.push(w.dn)});
  });
  var med=pct(wd,0.5)||0;
  items.forEach(function(it){
    if(!it.words)return;
    var ws=it.words;
    ws.forEach(function(w){w.b=(w.dn!==null&&med)?w.dn>med*1.28:null});
    ws.forEach(function(w,ix){
      if(w.b&&w.dn<med*1.5){
        var l1=ix>0&&ws[ix-1].b,r1=ix<ws.length-1&&ws[ix+1].b;
        if(!l1&&!r1&&ws.length>1&&!/[:.]$/.test(w.s))w.b=false;
      }
    });
    ws.forEach(function(w,ix){if(w.b&&ix>0&&ix<=3&&/:$/.test(w.s))for(var z=0;z<ix;z++)ws[z].b=true});
    for(var i2=0;i2<ws.length;i2++)if(ws[i2].b===null){
      var l=i2>0?ws[i2-1].b:null,r=i2<ws.length-1?ws[i2+1].b:null;
      if(l&&/[.:]$/.test(ws[i2-1].s)){ws[i2].b=false;continue}
      ws[i2].b=(l!==null&&l===r)?l:(l!==null?l:(r!==null?r:false));
    }
  });
  /* words too short to measure follow the line's weight */
  lines.forEach(function(L){
    var out2=[];
    L.items.forEach(function(it){
      if(!it.words||!it.words.length){out2.push(it);return}
      var ws=it.words;
      if(ws.every(function(w){return w.b===ws[0].b})){it.b=ws[0].b;out2.push(it);return}
      /* split into groups of equal weight */
      var g0=0;
      for(var q=1;q<=ws.length;q++){
        if(q===ws.length||ws[q].b!==ws[g0].b){
          var st=ws[g0].i0,en=ws[q-1].i0+ws[q-1].s.length;
          if(q<ws.length)en=ws[q].i0;
          var piece=Object.assign({},it);
          piece.s=it.s.slice(st,en);piece.x=ws[g0].x0;piece.w=(q<ws.length?ws[q].x0:it.x+it.w)-ws[g0].x0;piece.b=ws[g0].b;piece.words=null;
          out2.push(piece);g0=q;
        }
      }
    });
    L.items=out2;
  });
  /* bullets and numbers drawn in front of a line */
  var xs=lines.filter(function(L){return L.text.length>25}).map(function(L){return L.x0}),left=pct(xs.length?xs:lines.map(function(L){return L.x0}),0.05);
  lines.forEach(function(L){
    if(L.x0<left+10||Math.abs(L.size-body)>1)return;
    var x0=Math.max(left-2,L.x0-36),x1=L.x0-3,y0=L.y-L.size*0.25,y1=L.y+L.size*0.85;
    if(x1-x0<6)return;
    var a=px(x0),c=px(x1),b=py(y1),e=py(y0),minx=1e9,maxx=-1,miny=1e9,maxy=-1,k=0,u,v;
    for(v=b;v<=e;v++)for(u=a;u<=c;u++)if(dark[v*PW+u]){k++;if(u<minx)minx=u;if(u>maxx)maxx=u;if(v<miny)miny=v;if(v>maxy)maxy=v}
    if(k<3)return;
    var bw=(maxx-minx+1)/s,bh=(maxy-miny+1)/s;
    if(bh>1.25*L.size||bw>1.6*L.size)return;
    if(bw<=0.55*L.size&&bh<=0.55*L.size)L.rm={kind:'b',x:minx/s};
    else if(bh>=0.5*L.size)L.rm={kind:'n',x:minx/s};
  });
  lines.forEach(function(L){var real=L.items.filter(function(i){return !i.ws});L.bold=real.length>0&&real.every(function(i){return i.b})});
}

async function readPage(P,page,no,opts,fc){
  var vp=page.getViewport({scale:1}),vb=page.view,ox=vb[0],oy=vb[1],W=vp.width,H=vp.height;
  var ol=await page.getOperatorList(),g=scanOps(P,ol,ox,oy);
  var tc=await page.getTextContent(),styles=tc.styles||{},items=[];
  tc.items.forEach(function(it){
    if(!it.str)return;
    var t=it.transform;if(!t||Math.abs(t[1])>Math.abs(t[0])*0.5)return;
    var size=Math.hypot(t[2],t[3])||it.height||10;if(size<2.5)return;
    var fi=fontInfo(page,it.fontName,styles,fc);
    items.push({s:despace(it.str),x:t[4]-ox,y:t[5]-oy,w:it.width,size:size,b:fi.b,i:fi.i,m:fi.m,f:fi.f,ws:!it.str.trim()});
  });
  /* links */
  var links=[];
  try{(await page.getAnnotations()).forEach(function(a){
    if(a.subtype==='Link'&&a.url&&/^(https?:|mailto:|tel:)/i.test(a.url)&&a.rect)links.push({url:a.url,x0:a.rect[0]-ox,y0:a.rect[1]-oy,x1:a.rect[2]-ox,y1:a.rect[3]-oy});
  })}catch(e){}
  /* highlight boxes behind text (inline code, marker pen) */
  items.forEach(function(it){
    var cx=it.x+it.w/2,cy=it.y+it.size*0.3;
    for(var k=0;k<g.fills.length;k++){
      var f=g.fills[k],fh=f.y1-f.y0,fw=f.x1-f.x0;
      if(cx<f.x0||cx>f.x1||cy<f.y0||cy>f.y1)continue;
      if(fh>2*it.size+6||fw>W*0.6)continue;
      if(f.rgb[0]>=250&&f.rgb[1]>=250&&f.rgb[2]>=250)continue;
      if(lumOf(f.rgb)<0.45)continue;
      it.sh=hexOf(f.rgb);break;
    }
    for(var l=0;l<links.length;l++){var a=links[l];if(cx>=a.x0&&cx<=a.x1&&cy>=a.y0-1&&cy<=a.y1+1){it.lnk=a.url;break}}
  });
  /* group into lines */
  items.sort(function(a,b){return b.y-a.y||a.x-b.x});
  var lines=[],cur=null;
  items.forEach(function(it){
    if(cur&&Math.abs(cur.y-it.y)<=0.4*Math.max(cur.size,it.size))cur.items.push(it);
    else{cur={y:it.y,size:it.size,items:[it],page:no};lines.push(cur)}
  });
  var out=[];
  lines.forEach(function(L){
    L.items.sort(function(a,b){return a.x-b.x});
    var real=L.items.filter(function(i){return !i.ws});if(!real.length)return;
    var chars={},best=0;
    real.forEach(function(i){var k=r05(i.size);chars[k]=(chars[k]||0)+i.s.trim().length;if(chars[k]>best){best=chars[k];L.size=i.size}});
    L.y=real.reduce(function(m,i){return Math.max(m,i.y)},-1e9)>L.y?L.y:L.y;
    L.x0=real[0].x;L.x1=Math.max.apply(null,real.map(function(i){return i.x+i.w}));
    L.text=plain(L.items).replace(/\s+/g,' ').trim();
    L.bold=real.every(function(i){return i.b});
    /* cells = pieces of the line separated by a wide gap */
    var cells=[],cc=null,prevEnd=null;
    L.items.forEach(function(i){
      if(i.ws){if(cc)cc.items.push(i);return}
      if(!cc||i.x-prevEnd>Math.max(1.8*L.size,16)){cc={x0:i.x,x1:i.x+i.w,items:[i]};cells.push(cc)}
      else{cc.items.push(i);cc.x1=i.x+i.w}
      prevEnd=i.x+i.w;
    });
    L.cells=cells;L.pageW=W;L.pageH=H;
    out.push(L);
  });
  /* picture + invisible text pages */
  var R=null,segsAll=g.segs;
  if(g.inv&&opts.raster){
    try{R=await opts.raster(page);if(R){segsAll=[];rasterStep(R,W,H,out,items,segsAll)}}catch(e){R=null;segsAll=g.segs}
  }
  /* pictures */
  var pics=[];
  g.imgs.forEach(function(b){
    var w=b.x1-b.x0,h=b.y1-b.y0;
    if(w<14||h<14||w*h>0.7*W*H)return;
    for(var i=0;i<pics.length;i++){var p=pics[i],ix=Math.min(p.x1,b.x1)-Math.max(p.x0,b.x0),iy=Math.min(p.y1,b.y1)-Math.max(p.y0,b.y0);if(ix>0&&iy>0&&ix*iy>0.6*Math.min(w*h,(p.x1-p.x0)*(p.y1-p.y0)))return}
    pics.push(b);
  });
  var picEls=[];
  if(opts.crop){
    for(var i=0;i<pics.length;i++){
      try{
        var c=await opts.crop(page,pics[i],{ox:ox,oy:oy});
        if(c&&c.bytes)picEls.push({kind:'img',page:no,top:pics[i].y1,bottom:pics[i].y0,x0:pics[i].x0,x1:pics[i].x1,w:pics[i].x1-pics[i].x0,h:pics[i].y1-pics[i].y0,bytes:c.bytes});
      }catch(e){}
    }
    /* text drawn inside a picture is already part of the picture */
    var keep=[];
    out.forEach(function(L){
      var cx=(L.x0+L.x1)/2,inside=picEls.some(function(p){return cx>=p.x0&&cx<=p.x1&&L.y>=p.bottom&&L.y<=p.top});
      if(!inside)keep.push(L);
    });
    out=keep;
  }
  return {no:no,w:W,h:H,lines:out,pics:picEls,segs:segsAll,fills:g.fills,R:R,inv:g.inv};
}

/* ---------------------------------------------------------------- runs (styled pieces of text) */
function sk(r){return [r.b?1:0,r.i?1:0,r.m?1:0,r.size?Math.round(r.size*2):'',r.f||'',r.sh||'',r.lnk||''].join('|')}
function pushRun(runs,t,st){
  if(!t)return;
  var last=runs[runs.length-1];
  if(last&&sk(last)===sk(st))last.t+=t;
  else runs.push({t:t,b:st.b,i:st.i,m:st.m,size:st.size,f:st.f,sh:st.sh,lnk:st.lnk});
}
var NOSP_BEFORE=/^[)\]}.,;:!?%\u201D\u2019'"]/,NOSP_AFTER=/[(\[{\u201C\u2018]$/;
function buildRuns(items,o){
  o=o||{};
  var runs=[],prev=null,pend=false,plainSt={b:false,i:false,m:false};
  items.forEach(function(it){
    if(it.ws){if(prev)pend=true;return}
    var t=it.s.replace(/\s+/g,' '),lead=t[0]===' ';
    if(prev){
      var gap=it.x-(prev.x+prev.w),ptxt=prev.s,need=false;
      if(pend||gap>0.15*it.size)need=true;
      var hasSp=/ $/.test(ptxt)||lead;
      if(need&&!hasSp){
        var rawNext=t.replace(/^ /,'');
        if(NOSP_BEFORE.test(rawNext)||NOSP_AFTER.test(ptxt.trim())||/(^|\s)["']$/.test(ptxt))need=false;
      }
      if(need&&!hasSp){
        var sp=it.m?(prev&&!prev.m?prevStyle(prev):plainSt):stOf(it,o);
        pushRun(runs,' ',sp);
      }
    }
    pushRun(runs,t,stOf(it,o));
    prev=it;pend=false;
  });
  /* trim ends */
  while(runs.length&&!runs[0].t.trim())runs.shift();
  while(runs.length&&!runs[runs.length-1].t.trim())runs.pop();
  if(runs.length){runs[0].t=runs[0].t.replace(/^\s+/,'');runs[runs.length-1].t=runs[runs.length-1].t.replace(/\s+$/,'')}
  return runs;
}
function prevStyle(p){return {b:p.b,i:p.i,m:p.m,f:p.f,lnk:p.lnk}}
function stOf(it,o){return {b:it.b,i:it.i,m:it.m,f:it.f,sh:o.noShade?null:it.sh,lnk:o.noLink?null:it.lnk,sz:null,size:it.size}}

/* ---------------------------------------------------------------- tables */
function mapCols(L,cols){
  var res=[];
  for(var i=0;i<L.cells.length;i++){
    var c=L.cells[i],k=-1;
    for(var j=0;j<cols.length;j++){if(Math.abs(c.x0-cols[j].x0)<=6||(Math.abs(c.x1-cols[j].x1)<=6&&c.x0>=cols[j].x0-6)){k=j;break}}
    if(k<0)return null;
    if(res.length&&res[res.length-1].k>=k)return null;
    var lim=k<cols.length-1?cols[k+1].x0-2:1e9;
    if(c.x1>lim)return null;
    res.push({k:k,cell:c});
  }
  return res;
}
function candRow(L){
  if(L.cells.length<2)return false;
  var first=plain(L.cells[0].items).trim();
  if(markerOf(first+' x'))return false;
  return true;
}
function makeRow(L,map,n){
  var cells=[];for(var i=0;i<n;i++)cells.push({items:[],x0:null,x1:null});
  map.forEach(function(m){cells[m.k]={items:m.cell.items.slice(),x0:m.cell.x0,x1:m.cell.x1}});
  return {y:L.y,size:L.size,cells:cells};
}
function findTables(P,pg){
  var lines=pg.lines,els=[],used={},i=0;
  while(i<lines.length){
    var L=lines[i];
    if(!candRow(L)){i++;continue}
    var cols=L.cells.map(function(c){return {x0:c.x0,x1:c.x1}});
    var rows=[makeRow(L,L.cells.map(function(c,k){return {k:k,cell:c}}),cols.length)],lastY=L.y,j=i+1,pitch=0,lastRowY=L.y;
    while(j<lines.length){
      var N=lines[j],gap=lastY-N.y;
      if(gap>3.2*Math.max(L.size,N.size)||gap<=0)break;
      var map=mapCols(N,cols);if(!map)break;
      var full=map.length===cols.length;
      if(rows.length<2&&!full)break;
      var rowPitch=rows.length>=2?(rows[rows.length-1].y-rows[0].y)/(rows.length-1):0;
      var isCont=rows.length>=2&&!full&&(lastRowY-N.y)<0.8*rowPitch&&(lastY-N.y)<0.8*rowPitch;
      if(rows.length>=2&&!full&&map.length===1&&(lastY-N.y)<0.8*rowPitch)isCont=true;
      if(isCont){
        var row=rows[rows.length-1];
        map.forEach(function(m){var c=row.cells[m.k];c.items=c.items.concat([{s:' ',ws:true,x:0,w:0}],m.cell.items);if(c.x0===null){c.x0=m.cell.x0}c.x1=Math.max(c.x1||0,m.cell.x1)});
      }else{
        rows.push(makeRow(N,map,cols.length));lastRowY=N.y;
        if(full)cols.forEach(function(c,k){c.x0=Math.min(c.x0,N.cells[k].x0);c.x1=Math.max(c.x1,N.cells[k].x1)});
      }
      lastY=N.y;j++;
    }
    if(rows.length>=2){
      for(var k=i;k<j;k++)used[k]=1;
      els.push(tableEl(pg,rows,cols,lines[i],lines[j-1]));
      i=j;
    }else i++;
  }
  return {els:els,used:used};
}
function tableEl(pg,rows,cols,first,last){
  var n=cols.length,top=first.y+first.size*0.95,bot=last.y-last.size*0.35;
  var minX=cols[0].x0,maxX=cols[n-1].x1;
  /* ruling lines */
  var hs=pg.segs.filter(function(s){return s.h&&s.y>=bot-14&&s.y<=top+14&&s.x1>minX-14&&s.x0<maxX+14});
  var vs=pg.segs.filter(function(s){return !s.h&&s.y1>bot-14&&s.y0<top+14});
  var bordered=hs.length>=2,L=minX,R=null,bounds=null;
  if(bordered){
    L=Math.min.apply(null,hs.map(function(s){return s.x0}));R=Math.max.apply(null,hs.map(function(s){return s.x1}));
    var xs=[];
    vs.filter(function(s){return s.x>=L-3&&s.x<=R+3}).forEach(function(s){
      for(var q=0;q<xs.length;q++)if(Math.abs(xs[q]-s.x)<2)return;
      xs.push(s.x);
    });
    xs.sort(function(a,b){return a-b});
    if(xs.length===n+1)bounds=xs;
  }
  if(!bounds){
    bounds=[];for(var k=0;k<n;k++)bounds.push(bordered&&k===0?L:cols[k].x0-(bordered?0:0));
    if(bordered){bounds[0]=L;for(var k2=1;k2<n;k2++)bounds[k2]=cols[k2].x0-(cols[k2].x0-(cols[k2-1].x1))/2}
    bounds.push(bordered?R:Math.max(maxX+6,pg.w-0));
    if(!bordered)bounds[n]=Math.max(maxX+6,bounds[n]);
    if(!bordered)for(var k3=0;k3<n;k3++)bounds[k3]=cols[k3].x0-4;
  }
  var shade=rows.map(function(r){
    return r.cells.map(function(c,ci){
      var cx=(c.x0===null?bounds[ci]+6:c.x0+2),cy=r.y+r.size*0.3,best=null,ba=1e12;
      if(pg.R&&pg.R.rgbAt){
        var sx=(c.x0===null?bounds[ci]+5:c.x0-4),px=pg.R.rgbAt(sx,cy),mn=Math.min(px[0],px[1],px[2]);
        return (mn<246&&lumOf(px)>0.45)?hexOf(px):null;
      }
      pg.fills.forEach(function(f){
        if(cx<f.x0||cx>f.x1||cy<f.y0||cy>f.y1)return;
        if(f.rgb[0]>=250&&f.rgb[1]>=250&&f.rgb[2]>=250)return;
        var a=(f.x1-f.x0)*(f.y1-f.y0);if(a>0.6*pg.w*pg.h)return;
        if(a<ba){ba=a;best=f}
      });
      return best?hexOf(best.rgb):null;
    });
  });
  return {kind:'table',page:pg.no,top:top,bottom:bot,rows:rows,bounds:bounds,bordered:bordered,shade:shade};
}

/* ---------------------------------------------------------------- build the document */
function median(a){return pct(a,0.5)}
function analyse(pages){
  var sizes={},all=[],x0s=[],x1s=[];
  pages.forEach(function(pg){pg.lines.forEach(function(L){
    all.push(L);var k=r05(L.size);sizes[k]=(sizes[k]||0)+L.text.length;
    if(L.text.length>25){x0s.push(L.x0);x1s.push(L.x1)}
  })});
  var body=12,best=0;Object.keys(sizes).forEach(function(k){if(sizes[k]>best){best=sizes[k];body=+k}});
  return {body:body,all:all,left:pct(x0s.length?x0s:all.map(function(l){return l.x0}),0.03),right:pct(x1s.length?x1s:all.map(function(l){return l.x1}),0.92)};
}
function dropRepeats(pages){
  if(pages.length<3)return;
  var cnt={};
  pages.forEach(function(pg){
    var seen={};
    pg.lines.forEach(function(L){
      var edge=L.y>pg.h*0.91||L.y<pg.h*0.09;if(!edge)return;
      var sig=(L.y>pg.h*0.5?'T':'B')+L.text.toLowerCase().replace(/\d+/g,'#');
      if(!seen[sig]){seen[sig]=1;cnt[sig]=(cnt[sig]||0)+1}
    });
  });
  pages.forEach(function(pg){
    pg.lines=pg.lines.filter(function(L){
      var edge=L.y>pg.h*0.91||L.y<pg.h*0.09;if(!edge)return true;
      var sig=(L.y>pg.h*0.5?'T':'B')+L.text.toLowerCase().replace(/\d+/g,'#');
      return !(cnt[sig]>=3&&cnt[sig]>=pages.length*0.6);
    });
  });
}

function build(JSZip,pages,opts,stats){
  dropRepeats(pages);
  var A=analyse(pages),body=A.body;
  if(!A.all.length)return null;
  var textW=A.right-A.left;

  /* heading sizes */
  var cand=[];
  A.all.forEach(function(L){if((L.size>=body*1.12||(L.size>=body*1.06&&L.bold))&&L.text.length<=160&&!markerOf(L.text)&&!L.rm)cand.push(L)});
  var lv=[];
  cand.map(function(L){return r05(L.size)}).sort(function(a,b){return b-a}).forEach(function(s){
    if(!lv.length||lv[lv.length-1].size-s>0.9)lv.push({size:s,n:0});
    lv[lv.length-1].n++;
  });
  var hasTitle=false;
  if(lv.length){
    var tl=cand.filter(function(L){return r05(L.size)>=lv[0].size-0.9});
    if(tl.length<=3&&tl.every(function(L){return L.page===1})&&(lv.length===1||lv[0].size>=lv[1].size*1.15)){hasTitle=true}
  }
  function levelOf(L){
    for(var i=0;i<lv.length;i++)if(r05(L.size)>=lv[i].size-0.9&&(L.size>=body*1.12||(L.size>=body*1.06&&L.bold))&&!markerOf(L.text)&&!L.rm)return i;
    return -1;
  }
  var hLevels=lv.length-(hasTitle?1:0);
  var lowHead=Math.min(6,hLevels+1);

  /* pages -> element stream */
  var stream=[];
  pages.forEach(function(pg){
    var t=findTables(P_,pg),els=[];
    pg.lines.forEach(function(L,i){if(!t.used[i])els.push({kind:'line',top:L.y+L.size*0.9,L:L,page:pg.no})});
    t.els.forEach(function(e){els.push(e)});
    pg.pics.forEach(function(e){els.push(e)});
    els.sort(function(a,b){return b.top-a.top});
    els.forEach(function(e){e.pg=pg;stream.push(e)});
    stats.tables+=t.els.length;
  });

  /* style table for headings */
  var hs={};
  function runSz(r,base){return r}
  /* paragraphs */
  var blocks=[],cur=null,prevEl=null;
  function flush(){if(cur){blocks.push(cur);cur=null}}
  function lineRuns(L,items){
    return buildRuns(items||L.items,{});
  }
  function addRuns(p,runs){
    if(!runs.length)return;
    if(p.runs.length){
      var last=p.runs[p.runs.length-1],first=runs[0];
      if(/[A-Za-z\u00C0-\u024F]-$/.test(last.t)&&/^[a-z\u00DF-\u00FF]/.test(first.t)&&!last.m){last.t=last.t.slice(0,-1)}
      else pushRun(p.runs,' ',/^\s/.test(first.t)||last.m?{b:false,i:false,m:false}:{b:first.b,i:first.i,m:first.m,f:first.f,lnk:(last.lnk&&first.lnk===last.lnk)?last.lnk:null});
    }
    runs.forEach(function(r){pushRun(p.runs,r.t,r)});
  }
  stream.forEach(function(e){
    if(e.kind!=='line'){flush();blocks.push(e);prevEl=e;return}
    var L=e.L,m=markerOf(L.text),hl=levelOf(L),isH=hl>=0,boldHead=false;
    if(!isH&&!m&&!L.rm&&L.bold&&L.text.length<=80&&!/[.,;:]$/.test(L.text)&&L.size>=body*0.99&&L.cells.length===1)boldHead=true;
    var pl=cur&&cur.lastLine,sameP=!!pl&&pl.page===L.page,gap=pl?(pl.y-L.y):0;
    /* heading */
    if(isH||boldHead){
      var level=isH?(hasTitle&&hl===0?0:hl-(hasTitle?1:0)+1):lowHead;
      if(cur&&cur.t==='h'&&cur.level===level&&sameP&&gap<1.6*L.size){addRuns(cur,lineRuns(L));cur.lastLine=L;return}
      flush();
      cur={t:'h',level:level,runs:[],lastLine:L,left:L.x0,lines:[L]};addRuns(cur,lineRuns(L));
      return;
    }
    /* list item */
    if(!m&&L.rm)m={kind:L.rm.kind,len:0,x:L.rm.x,raster:1};
    if(m){
      flush();
      var prevB=blocks[blocks.length-1];
      if(m.raster&&m.kind==='n')m.num=(prevB&&prevB.t==='li'&&prevB.kind==='n'&&prevB.raster&&prevB.lastLine.page===L.page&&(prevB.lastLine.y-L.y)<4*L.size)?prevB.num+1:1;
      if(m.raster&&m.kind==='n'&&!m.sep)m.sep='.';
      var items=L.items.slice(),need=m.len,mEnd=null;
      /* strip marker characters from the front of the line */
      var out=[];
      items.forEach(function(it){
        if(need<=0){out.push(it);return}
        var t=it.s;
        if(t.length<=need){need-=t.length;return}
        var cut=need,nit=Object.assign({},it);nit.s=t.slice(cut);nit.x=it.x+it.w*(cut/t.length);nit.w=it.w*(1-cut/t.length);need=0;out.push(nit);
      });
      var first=out.filter(function(i){return !i.ws})[0];
      cur={t:'li',kind:m.kind,num:m.num,sep:m.sep,letter:m.letter,raster:m.raster,markerX:m.x!==undefined?m.x:L.x0,textX:first?first.x:L.x0+18,runs:[],lastLine:L,lines:[L]};
      addRuns(cur,buildRuns(out,{}));
      return;
    }
    /* plain text line */
    if(cur&&cur.t==='li'){
      var okc=L.x0>cur.markerX+3&&gap<1.7*L.size&&gap>0&&sameP;
      var crossPage=pl&&pl.page!==L.page&&L.x0>cur.markerX+3&&!/[.!?]$/.test(pl.text);
      if(okc||crossPage){addRuns(cur,lineRuns(L));cur.lastLine=L;cur.lines.push(L);return}
      flush();
    }
    if(cur&&cur.t==='p'){
      var prev=cur.lastLine,join=false;
      var shortPrev=prev.x1<A.right-0.22*textW;
      if(sameP){
        join=gap>0&&gap<=1.75*Math.max(prev.size,L.size)&&Math.abs(prev.size-L.size)<1.2&&L.x0<=cur.left+8&&L.x0>=cur.left-8&&!shortPrev;
      }else{
        join=Math.abs(prev.size-L.size)<1.2&&Math.abs(L.x0-cur.left)<8&&(prev.x1>=A.right-0.12*textW||(!/[.!?:;)"\u201D]$/.test(prev.text)&&/^[a-z]/.test(L.text)));
      }
      if(join){addRuns(cur,lineRuns(L));cur.lastLine=L;cur.lines.push(L);return}
      flush();
    }
    flush();
    cur={t:'p',runs:[],lastLine:L,left:L.x0,lines:[L],size:L.size};addRuns(cur,lineRuns(L));
  });
  flush();

  /* ---------- XML ---------- */
  var rels=[],media=[],relId=100,linkIds={};
  function linkRel(u){if(linkIds[u])return linkIds[u];var id='rId'+(++relId);linkIds[u]=id;rels.push('<Relationship Id="'+id+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'+esc(u)+'" TargetMode="External"/>');return id}
  var defFont=(function(){var c={},bf='Arial',mx=0;A.all.forEach(function(L){L.items.forEach(function(i){if(i.ws||i.m)return;c[i.f]=(c[i.f]||0)+i.s.length;if(c[i.f]>mx){mx=c[i.f];bf=i.f}})});return bf})();
  var bodyHalf=Math.round(body*2);

  function rPr(r,base){
    var x='';
    if(r.lnk)x+='<w:rStyle w:val="Hyperlink"/>';
    var f=r.m?(r.f||'Consolas'):(r.f&&r.f!==defFont?r.f:null);
    if(f)x+='<w:rFonts w:ascii="'+esc(f)+'" w:hAnsi="'+esc(f)+'" w:cs="'+esc(f)+'"/>';
    if(base.heading){}else if(r.b)x+='<w:b/><w:bCs/>';
    if(r.i)x+='<w:i/><w:iCs/>';
    var sz=null;
    if(!base.heading&&r.size&&Math.abs(Math.round(r.size*2)-bodyHalf)>=2)sz=Math.round(r.size*2);
    if(sz)x+='<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/>';
    if(r.sh)x+='<w:shd w:val="clear" w:color="auto" w:fill="'+r.sh+'"/>';
    return x?'<w:rPr>'+x+'</w:rPr>':'';
  }
  function runsXml(runs,base){
    var out='';
    runs.forEach(function(r){
      var t='<w:r>'+rPr(r,base)+'<w:t xml:space="preserve">'+esc(r.t)+'</w:t></w:r>';
      if(r.lnk)t='<w:hyperlink r:id="'+linkRel(r.lnk)+'" w:history="1">'+t+'</w:hyperlink>';
      out+=t;
    });
    return out;
  }
  /* run sizes: remember each run's own font size for non-body text */
  function annotate(runs,L){}
  var numXml=[],numId=1,bulletNum=1,lastListKind=null,curNum=null;
  var abs=1; /* abstractNum 1 bullets, 2 numbers '.', 3 numbers ')' */
  numXml.push('<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>');
  var nextNum=2;
  function newNum(absId,start){
    var id=nextNum++;
    numXml.push('<w:num w:numId="'+id+'"><w:abstractNumId w:val="'+absId+'"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="'+start+'"/></w:lvlOverride></w:num>');
    return id;
  }
  var LM=clamp(A.left,28,150),RM=clamp(pages[0].w-A.right,28,150);
  var tops=[],bots=[];pages.forEach(function(pg){if(pg.lines.length){tops.push(pg.h-Math.max.apply(null,pg.lines.map(function(l){return l.y+l.size})));bots.push(Math.min.apply(null,pg.lines.map(function(l){return l.y})))}});
  var TM=clamp(median(tops)-6,28,150),BM=clamp(median(bots)-6,28,150);
  var contentW=pages[0].w-LM-RM;
  var body_xml='',lastBlockWasList=null,imgN=0,docPr=0,lastNum=null;

  var hSizes={};
  function hStyleSize(level){
    if(hSizes[level])return hSizes[level];
    return null;
  }
  /* collect heading sizes + boldness per level for the style sheet */
  var styleInfo={};
  blocks.forEach(function(b){
    if(b.t!=='h')return;
    var s=styleInfo[b.level]||(styleInfo[b.level]={size:0,bold:0,chars:0});
    var L=b.lines[0],n=b.runs.reduce(function(a,r){return a+r.t.length},0),bn=b.runs.reduce(function(a,r){return a+(r.b?r.t.length:0)},0);
    s.size=Math.max(s.size,L.size);s.bold+=bn;s.chars+=n;
  });
  Object.keys(styleInfo).forEach(function(k){styleInfo[k].isBold=styleInfo[k].bold>=styleInfo[k].chars*0.5});

  function pXml(b){
    var ppr='',base={size:b.size||body};
    if(b.pageBreak)ppr+='<w:pageBreakBefore/>';
    var lines=b.lines,L0=lines[0],jc=null,ind='';
    function centered(L){return Math.abs((L.x0+L.x1)/2-(A.left+A.right)/2)<8}
    if(b.t==='h'){
      var si=styleInfo[b.level];base={heading:true,bold:si.isBold};
      ppr='<w:pStyle w:val="'+(b.level===0?'Title':'Heading'+b.level)+'"/>'+ppr;
    }else if(b.t==='li'){
      var num;
      if(b.kind==='b')num=1;
      else{
        var sameRun=lastListKind==='n:'+b.sep&&curNum&&lastNum!==null&&b.num===lastNum+1;
        if(!sameRun)curNum=newNum(b.sep===')'?3:2,b.num||1);
        num=curNum;lastNum=b.num||1;
      }
      lastListKind=b.kind==='b'?'b':'n:'+b.sep;
      var left=Math.max(0,b.textX-LM),hang=clamp(b.textX-b.markerX,10,48);
      if(b.textX-b.markerX<6)hang=18;
      ppr+='<w:numPr><w:ilvl w:val="0"/><w:numId w:val="'+num+'"/></w:numPr><w:spacing w:after="60"/><w:ind w:left="'+Math.round((Math.max(left,hang))*20)+'" w:hanging="'+Math.round(hang*20)+'"/>';
    }else{
      var multi=lines.length>1;
      if(!multi&&centered(L0)&&L0.x0-A.left>20&&A.right-L0.x1>20)jc='center';
      else if(multi&&lines.every(centered)&&lines.some(function(l){return Math.abs(l.x0-lines[0].x0)>3}))jc='center';
      else if(!multi&&L0.x1>A.right-3&&L0.x0-A.left>0.35*textW)jc='right';
      var off=L0.x0-A.left;
      if(!jc&&off>10)ind='<w:ind w:left="'+Math.round(off*20)+'"/>';
      if(jc)ppr+='<w:jc w:val="'+jc+'"/>';else ppr+=ind;
    }
    if(b.t!=='li'){lastListKind=null;lastNum=null}
    if(b.t==='h'&&jc===null&&lines.length===1&&centered(L0)&&L0.x0-A.left>20&&A.right-L0.x1>20)ppr+='<w:jc w:val="center"/>';
    /* give every run its own point size when it differs from the paragraph */
    var runs=b.runs;
    return '<w:p>'+(ppr?'<w:pPr>'+ppr+'</w:pPr>':'')+runsXml(runs,base)+'</w:p>';
  }

  function cellRunsXml(items,size,shaded){
    if(shaded)items=items.map(function(i){return i.mr?Object.assign({},i,{m:false,f:i.f0,sh:null,mr:false}):i});
    var rs=buildRuns(items,{noShade:true});
    return runsXml(rs,{size:size});
  }
  function tableXml(t){
    var n=t.bounds.length-1,W=[],tot=0;
    for(var i=0;i<n;i++){var w=Math.max(300,Math.round((t.bounds[i+1]-t.bounds[i])*20));W.push(w);tot+=w}
    var maxTot=Math.round(contentW*20)+200;
    if(tot>maxTot){var k=maxTot/tot;W=W.map(function(w){return Math.round(w*k)});tot=W.reduce(function(a,b){return a+b},0)}
    var ind=Math.round((t.bounds[0]-LM)*20);if(ind<40)ind=0;
    var bd='<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
    var x='<w:tbl><w:tblPr><w:tblW w:w="'+tot+'" w:type="dxa"/>'+(ind?'<w:tblInd w:w="'+ind+'" w:type="dxa"/>':'')+(t.bordered?'<w:tblBorders>'+bd+'</w:tblBorders>':'')+'<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="50" w:type="dxa"/><w:left w:w="110" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>'+W.map(function(w){return '<w:gridCol w:w="'+w+'"/>'}).join('')+'</w:tblGrid>';
    t.rows.forEach(function(r,ri){
      var head=ri===0&&(t.shade[0].some(Boolean)||r.cells.every(function(c){return !c.items.length||c.items.filter(function(i){return !i.ws}).every(function(i){return i.b})}));
      x+='<w:tr><w:trPr><w:cantSplit/>'+(head?'<w:tblHeader/>':'')+'</w:trPr>';
      r.cells.forEach(function(c,ci){
        var sh=t.shade[ri][ci];
        x+='<w:tc><w:tcPr><w:tcW w:w="'+W[ci]+'" w:type="dxa"/>'+(sh?'<w:shd w:val="clear" w:color="auto" w:fill="'+sh+'"/>':'')+'</w:tcPr><w:p><w:pPr><w:spacing w:after="0"/></w:pPr>'+cellRunsXml(c.items,r.size,!!sh)+'</w:p></w:tc>';
      });
      x+='</w:tr>';
    });
    return x+'</w:tbl><w:p><w:pPr><w:spacing w:after="60"/></w:pPr></w:p>';
  }
  function imgXml(e){
    var id=++imgN,rid='rId'+(200+id),w=Math.min(e.w,contentW),h=e.h*(w/e.w),cx=Math.round(w*12700),cy=Math.round(h*12700);
    media.push({name:'image'+id+'.png',bytes:e.bytes});
    rels.push('<Relationship Id="'+rid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image'+id+'.png"/>');
    var ctr=Math.abs((e.x0+e.x1)/2-(A.left+A.right)/2)<10&&e.x0-A.left>12;
    var dp=++docPr;
    return '<w:p><w:pPr><w:spacing w:after="120"/>'+(ctr?'<w:jc w:val="center"/>':'')+'</w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="'+cx+'" cy="'+cy+'"/><wp:docPr id="'+dp+'" name="Picture '+dp+'"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="'+dp+'" name="image'+id+'.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="'+rid+'"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
  }

  /* page breaks only where the PDF clearly ended a page early */
  var pageStartEl={};
  pages.forEach(function(pg,i){
    if(i===0)return;
    var prev=pages[i-1],lowest=prev.lines.length?Math.min.apply(null,prev.lines.map(function(l){return l.y})):null;
    if(lowest!==null&&lowest>prev.h*0.30&&lowest-BM>prev.h*0.28)pageStartEl[pg.no]=1;
  });
  var seenPage={};
  blocks.forEach(function(b){
    var pn=b.t?(b.lines[0].page):b.page;
    if(pageStartEl[pn]&&!seenPage[pn]){seenPage[pn]=1;if(b.t==='p'||b.t==='h')b.pageBreak=true}
    seenPage[pn]=1;
  });

  blocks.forEach(function(b){
    if(b.kind==='table'){body_xml+=tableXml(b);lastListKind=null;return}
    if(b.kind==='img'){body_xml+=imgXml(b);lastListKind=null;return}
    body_xml+=pXml(b);
  });

  var pw=pages[0].w,ph=pages[0].h,land=pw>ph;
  var sect='<w:sectPr><w:pgSz w:w="'+Math.round(pw*20)+'" w:h="'+Math.round(ph*20)+'"'+(land?' w:orient="landscape"':'')+'/><w:pgMar w:top="'+Math.round(TM*20)+'" w:right="'+Math.round(RM*20)+'" w:bottom="'+Math.round(BM*20)+'" w:left="'+Math.round(LM*20)+'" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr>';
  var NS='xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';
  var docXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document '+NS+'><w:body>'+body_xml+sect+'</w:body></w:document>';

  /* styles */
  var st='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="'+esc(defFont)+'" w:hAnsi="'+esc(defFont)+'" w:eastAsia="'+esc(defFont)+'" w:cs="'+esc(defFont)+'"/><w:sz w:val="'+bodyHalf+'"/><w:szCs w:val="'+bodyHalf+'"/><w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>';
  st+='<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>';
  st+='<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>';
  st+='<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>';
  Object.keys(styleInfo).forEach(function(k){
    k=+k;var s=styleInfo[k],id=k===0?'Title':'Heading'+k,sz=Math.round(s.size*2);
    st+='<w:style w:type="paragraph" w:styleId="'+id+'"><w:name w:val="'+(k===0?'Title':'heading '+k)+'"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="'+(k===0?0:Math.round(s.size*14))+'" w:after="'+Math.round(s.size*6)+'"/>'+(k===0?'':'<w:outlineLvl w:val="'+(k-1)+'"/>')+'</w:pPr><w:rPr>'+(s.isBold?'<w:b/><w:bCs/>':'')+'<w:sz w:val="'+sz+'"/><w:szCs w:val="'+sz+'"/></w:rPr></w:style>';
  });
  /* body-size bold lines use the lowest heading level */
  if(blocks.some(function(b){return b.t==='h'&&b.level===lowHead})&&!styleInfo[lowHead]){
    st+='<w:style w:type="paragraph" w:styleId="Heading'+lowHead+'"><w:name w:val="heading '+lowHead+'"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="'+(lowHead-1)+'"/></w:pPr><w:rPr><w:b/><w:bCs/></w:rPr></w:style>';
  }
  st+='</w:styles>';

  /* numbering */
  function lvls(fmt,text){
    var s='';
    for(var i=0;i<9;i++){
      var f=fmt==='bullet'?'bullet':(i%3===0?'decimal':i%3===1?'lowerLetter':'lowerRoman');
      var tx=fmt==='bullet'?(['\u2022','\u25E6','\u25AA'][i%3]):('%'+(i+1)+text);
      s+='<w:lvl w:ilvl="'+i+'"><w:start w:val="1"/><w:numFmt w:val="'+f+'"/><w:lvlText w:val="'+tx+'"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="'+(720+360*i)+'" w:hanging="360"/></w:pPr></w:lvl>';
    }
    return s;
  }
  var nm='<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+
    '<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>'+lvls('bullet')+'</w:abstractNum>'+
    '<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>'+lvls('num','.')+'</w:abstractNum>'+
    '<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="hybridMultilevel"/>'+lvls('num',')')+'</w:abstractNum>'+
    numXml.join('')+'</w:numbering>';

  var zip=new JSZip();
  zip.file('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>');
  zip.file('_rels/.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>');
  zip.file('word/document.xml',docXml);
  zip.file('word/styles.xml',st);
  zip.file('word/numbering.xml',nm);
  zip.file('word/_rels/document.xml.rels','<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>'+rels.join('')+'</Relationships>');
  media.forEach(function(m){zip.file('word/media/'+m.name,m.bytes)});
  stats.images=media.length;
  stats.paragraphs=blocks.filter(function(b){return b.t}).length;
  return zip;
}

var P_=null;
async function convert(P,JSZip,data,opts){
  opts=opts||{};P_=P;
  var pdf=await P.getDocument({data:data}).promise,N=pdf.numPages,pages=[],fc={},stats={pages:N,tables:0,images:0,paragraphs:0};
  for(var p=1;p<=N;p++){
    var pg=await pdf.getPage(p);
    pages.push(await readPage(P,pg,p,opts,fc));
    if(opts.onProgress)opts.onProgress(p,N);
    pg.cleanup&&pg.cleanup();
  }
  var zip=build(JSZip,pages,opts,stats);
  return {zip:zip,stats:stats};
}

var API={convert:convert};
if(typeof module!=='undefined'&&module.exports)module.exports=API;
else root.PDF2DOCX=API;

/* ---------------------------------------------------------------- page wiring (browser only) */
if(typeof document==='undefined'||!root.JU||!document.getElementById('wGo'))return;
var U=root.JU,$=U.$,WORKER='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js',shot={n:0,c:null,s:1};
async function setup(){
  var L=root.pdfjsLib;
  if(!L.GlobalWorkerOptions.workerSrc){
    var t=await (await fetch(WORKER)).text();
    L.GlobalWorkerOptions.workerSrc=URL.createObjectURL(new Blob([t],{type:'text/javascript'}));
  }
  return L;
}
async function renderPage(page){
  if(shot.n===page.pageNumber&&shot.c)return shot;
  var base=page.getViewport({scale:1}),s=Math.min(2,3200/Math.max(base.width,base.height));
  var vp=page.getViewport({scale:s}),c=document.createElement('canvas');c.width=Math.round(vp.width);c.height=Math.round(vp.height);
  var x=c.getContext('2d',{willReadFrequently:true});x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);
  await page.render({canvasContext:x,viewport:vp}).promise;
  shot={n:page.pageNumber,c:c,s:s,vp:vp};
  return shot;
}
async function raster(page){
  var r=await renderPage(page),c=r.c,id=c.getContext('2d',{willReadFrequently:true}).getImageData(0,0,c.width,c.height);
  return {w:c.width,h:c.height,data:id.data,s:r.s};
}
async function crop(page,b,o){
  var r=await renderPage(page),vp2=r.vp;
  var p1=vp2.convertToViewportPoint(b.x0+o.ox,b.y1+o.oy),p2=vp2.convertToViewportPoint(b.x1+o.ox,b.y0+o.oy);
  var sx=Math.max(0,Math.floor(Math.min(p1[0],p2[0]))),sy=Math.max(0,Math.floor(Math.min(p1[1],p2[1])));
  var sw=Math.min(r.c.width-sx,Math.ceil(Math.abs(p2[0]-p1[0]))),sh=Math.min(r.c.height-sy,Math.ceil(Math.abs(p2[1]-p1[1])));
  if(sw<4||sh<4)return null;
  var cc=document.createElement('canvas');cc.width=sw;cc.height=sh;
  cc.getContext('2d').drawImage(r.c,sx,sy,sw,sh,0,0,sw,sh);
  var blob=await new Promise(function(res){cc.toBlob(res,'image/png')});
  return {bytes:new Uint8Array(await blob.arrayBuffer())};
}
$('wGo').onclick=async function(){
  var W=U.getW();
  if(!W.length||!U.need('pdfjsLib','PDF reader')||!U.need('JSZip','ZIP library'))return;
  var btn=$('wGo'),done=0;
  U.busy(btn,true,'Converting...');
  for(var i=0;i<W.length;i++){
    var it=W[i];
    try{
      var L=await setup();shot={n:0,c:null,s:1};
      var res=await convert(L,root.JSZip,new Uint8Array(await it.file.arrayBuffer()),{
        crop:crop,raster:raster,
        onProgress:function(p,n){btn.textContent='File '+(i+1)+'/'+W.length+' \u00b7 page '+p+'/'+n}
      });
      if(!res.zip){U.toast('No text found in '+it.file.name+'. It may be a scanned PDF.',1);it.note='No text found';continue}
      btn.textContent='Saving...';
      var blob=await res.zip.generateAsync({type:'blob',mimeType:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',compression:'DEFLATE'});
      U.save(blob,U.base(it.file.name)+'.docx');
      var s=res.stats;
      it.note='Done: '+s.pages+' page'+(s.pages>1?'s':'')+(s.tables?' \u00b7 '+s.tables+' table'+(s.tables>1?'s':''):'')+(s.images?' \u00b7 '+s.images+' picture'+(s.images>1?'s':''):'');
      done++;
    }catch(e){
      if(root.console)console.error(e);
      it.note='Failed';U.toast('Could not read '+it.file.name+'. It may be damaged or password protected.',1);
    }
    root.JU.redrawW&&root.JU.redrawW();
  }
  shot={n:0,c:null,s:1};
  U.busy(btn,false);
  if(done)U.toast(done===1?'Word file is ready':done+' Word files are ready');
};
})(typeof window!=='undefined'?window:globalThis);
