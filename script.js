(function(){
  const buttons = Array.from(document.querySelectorAll('[data-tab]'));
  const sections = Array.from(document.querySelectorAll('.section'));

  function setActive(tabId, pushHash){
    buttons.forEach(btn=>{
      const isActive = btn.dataset.tab === tabId;
      btn.setAttribute('aria-selected', isActive ? 'true' : 'false');
      btn.classList.toggle('active', isActive);
    });
    sections.forEach(sec=>{
      const isActive = sec.id === tabId;
      sec.classList.toggle('active', isActive);
    });

    if(pushHash){
      history.replaceState(null, '', '#' + tabId);
    }
    // focus the section heading for accessibility
    const activeSection = document.getElementById(tabId);
    const heading = activeSection && activeSection.querySelector('h2, h1');
    if(heading) heading.setAttribute('tabindex','-1'), heading.focus({preventScroll:true});
    window.scrollTo({top:0, behavior:'smooth'});
  }

  function normalizeHash(){
    const hash = (location.hash || '').replace('#','').trim();
    if(!hash) return buttons[0]?.dataset.tab;
    const exists = sections.some(s=>s.id===hash);
    return exists ? hash : buttons[0]?.dataset.tab;
  }

  buttons.forEach(btn=>{
    btn.addEventListener('click', ()=>{
      setActive(btn.dataset.tab, true);
    });
  });

  window.addEventListener('hashchange', ()=>{
    setActive(normalizeHash(), false);
  });

  // init
  setActive(normalizeHash(), false);
})();
