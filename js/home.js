/* Home page extras: category chips and card entrance animation. */
(function(){
  var chips=document.querySelectorAll('.fchip'),tools=document.querySelectorAll('.tool');
  chips.forEach(function(b){b.onclick=function(){
    chips.forEach(function(x){x.classList.toggle('on',x===b)});
    var c=b.dataset.c;
    tools.forEach(function(t){t.classList.toggle('cat-off',c!=='all'&&t.dataset.cat!==c)});
  }});
  tools.forEach(function(t,i){t.style.animationDelay=(i*45)+'ms'});
})();

/* ---- Motion and polish ---- */
(function(){
  document.documentElement.classList.add('js');
  var reduce=window.matchMedia&&window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var prog=document.getElementById('prog'),bar=document.querySelector('.top');
  function onScroll(){
    var h=document.documentElement,max=h.scrollHeight-h.clientHeight;
    if(prog)prog.style.width=(max>0?h.scrollTop/max*100:0)+'%';
    if(bar)bar.classList.toggle('scrolled',h.scrollTop>24);
  }
  window.addEventListener('scroll',onScroll,{passive:true});onScroll();
  document.querySelectorAll('.tool').forEach(function(t){
    t.addEventListener('mousemove',function(e){var r=t.getBoundingClientRect();t.style.setProperty('--mx',(e.clientX-r.left)+'px');t.style.setProperty('--my',(e.clientY-r.top)+'px')});
  });
  document.addEventListener('mousedown',function(e){
    var b=e.target.closest&&e.target.closest('button');
    if(!b||b.classList.contains('tool')||b.disabled||reduce)return;
    var r=b.getBoundingClientRect(),s=Math.max(r.width,r.height),sp=document.createElement('span');
    sp.className='ripple';sp.style.width=sp.style.height=s+'px';
    sp.style.left=(e.clientX-r.left-s/2)+'px';sp.style.top=(e.clientY-r.top-s/2)+'px';
    b.appendChild(sp);setTimeout(function(){sp.remove()},650);
  });
  function count(el){
    var to=+el.dataset.to,suf=el.dataset.suf||'';
    if(!to||reduce){el.textContent=(to||0)+suf;return}
    var t0=null;
    function step(t){if(!t0)t0=t;var p=Math.min(1,(t-t0)/1300),e=1-Math.pow(1-p,3);el.textContent=Math.round(to*e)+suf;if(p<1)requestAnimationFrame(step)}
    requestAnimationFrame(step);
  }
  setTimeout(function(){document.querySelectorAll('.stat b[data-to]').forEach(count)},550);
  var rv=document.querySelectorAll('.reveal');
  if('IntersectionObserver' in window&&!reduce){
    var io=new IntersectionObserver(function(es){es.forEach(function(x){if(x.isIntersecting){x.target.classList.add('in');io.unobserve(x.target)}})},{threshold:.15});
    rv.forEach(function(x){io.observe(x)});
  }else rv.forEach(function(x){x.classList.add('in')});
})();
