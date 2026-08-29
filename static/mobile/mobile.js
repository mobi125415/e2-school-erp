(function(){
 const drawer=document.getElementById('drawer'),backdrop=document.getElementById('backdrop'),menu=document.getElementById('menuBtn'),more=document.getElementById('moreBtn');
 function close(){drawer.classList.remove('open');backdrop.classList.remove('show')}
 function open(){drawer.classList.add('open');backdrop.classList.add('show')}
 if(menu) menu.addEventListener('click',open); if(backdrop) backdrop.addEventListener('click',close); if(more) more.addEventListener('click',function(e){e.preventDefault();open()});
 document.querySelectorAll('.drawer a').forEach(a=>a.addEventListener('click',close));
 let deferred=null,card=document.getElementById('installCard'),btn=document.getElementById('installBtn');
 window.addEventListener('beforeinstallprompt',e=>{e.preventDefault();deferred=e;card.style.display='flex';});
 if(btn) btn.addEventListener('click',async()=>{if(deferred){deferred.prompt();await deferred.userChoice;deferred=null}else{alert('iPhone: Safari > Share > Add to Home Screen\nAndroid: Chrome > Menu > Add to Home screen');}});
 if('serviceWorker' in navigator) navigator.serviceWorker.register('/static/mobile/sw.js').catch(()=>{});
})();
