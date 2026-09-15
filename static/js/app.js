(function(){
  function parseWibDate(value){
    if(!value) return NaN;
    var normalized=String(value).replace(' ','T');
    var parts=normalized.split(/[-T:]/).map(Number);
    if(parts.length<5) return NaN;
    // WIB = UTC+7; convert manually so server/client timezone does not matter.
    return Date.UTC(parts[0],parts[1]-1,parts[2],parts[3]-7,parts[4],parts[5]||0);
  }
  function initTimer(root){
    var timer=root && root.querySelector ? root.querySelector('#custom-timer') : null;
    if(!timer || timer.dataset.timerBound==='1') return;
    timer.dataset.timerBound='1';
    var start=parseWibDate(timer.dataset.startWib);
    var duration=Number(timer.dataset.durationSeconds||0);
    var value=timer.querySelector('.timer-value');
    if(!Number.isFinite(start)||!duration||!value) return;
    function tick(){
      var remaining=Math.max(0, Math.ceil((duration*1000-(Date.now()-start))/1000));
      var m=Math.floor(remaining/60), s=remaining%60;
      value.textContent=String(m).padStart(2,'0')+':'+String(s).padStart(2,'0');
      if(remaining<=0){
        clearInterval(timer._interval);
        var form=document.getElementById('page-form');
        if(!form) return;
        var data=new FormData(form); data.set('__button__','__timeout__');
        fetch('/quiz/timeout',{method:'POST',headers:{'HX-Request':'true'},body:new URLSearchParams(data)})
          .then(function(r){return r.text()})
          .then(function(html){
            var app=document.getElementById('app'); if(app) app.innerHTML=html;
            if(window.htmx) htmx.process(app);
          });
      }
    }
    tick(); timer._interval=setInterval(tick,1000);
  }
  document.addEventListener('DOMContentLoaded',function(){initTimer(document);});
  document.body.addEventListener('htmx:afterSwap',function(e){initTimer(e.target);});
  document.body.addEventListener('htmx:beforeRequest',function(e){
    var b=e.detail && e.detail.elt;
    if(b && b.classList) b.classList.add('is-disabled');
  });
})();

window.addEventListener('load',function(){
  document.addEventListener('contextmenu',function(e){e.preventDefault();});
  document.addEventListener('copy',function(e){e.preventDefault();});
  document.addEventListener('keydown',function(e){if(e.ctrlKey&&['c','u','s','a'].includes(e.key.toLowerCase()))e.preventDefault();});
  var wakeLock=null;
  async function requestWakeLock(){if(!('wakeLock' in navigator))return;try{wakeLock=await navigator.wakeLock.request('screen')}catch(e){}}
  requestWakeLock();
  document.addEventListener('visibilitychange',function(){if(document.visibilityState==='visible')requestWakeLock();});
});
