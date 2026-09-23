(function () {
  // Smooth scroll for same-page anchors (with sticky nav offset)
  document.addEventListener('click', function (e) {
    var a = e.target && e.target.closest ? e.target.closest('a[href^="#"]') : null;
    if (!a) return;
    var href = a.getAttribute('href');
    if (!href || href === '#') return;

    var el = document.querySelector(href);
    if (!el) return;

    e.preventDefault();
    var nav = document.querySelector('.site-nav');
    var offset = nav ? nav.offsetHeight + 24 : 0;
    var top = el.getBoundingClientRect().top + window.pageYOffset - offset;
    window.scrollTo({ top: top, behavior: 'smooth' });

    if (nav) {
      nav.classList.remove('is-open');
      var toggle = nav.querySelector('.fk-nav-toggle');
      if (toggle) toggle.setAttribute('aria-expanded', 'false');
    }
  });

  // Fahim-style mobile nav toggle
  var siteNav = document.getElementById('siteNav');
  var menuBtn = siteNav ? siteNav.querySelector('.fk-nav-toggle') : null;
  if (siteNav && menuBtn) {
    menuBtn.addEventListener('click', function () {
      var open = siteNav.classList.toggle('is-open');
      menuBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    });
  }

  // Light / dark theme switcher (default: light)
  function getTheme() {
    return document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
  }

  function setTheme(theme) {
    var next = theme === 'dark' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', next);
    try { localStorage.setItem('theme', next); } catch (e) {}
    var meta = document.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute('content', next === 'dark' ? '#0F0F0F' : '#f6f7f6');
    document.querySelectorAll('.fk-theme-toggle').forEach(function (btn) {
      btn.setAttribute('aria-label', next === 'dark' ? 'Switch to light mode' : 'Switch to dark mode');
    });
  }

  document.querySelectorAll('.fk-theme-toggle').forEach(function (btn) {
    btn.addEventListener('click', function () {
      setTheme(getTheme() === 'dark' ? 'light' : 'dark');
    });
  });
  setTheme(getTheme());

  // Scroll reveal
  var reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  var nodes = document.querySelectorAll('.reveal');
  if (reduceMotion || !('IntersectionObserver' in window)) {
    nodes.forEach(function (n) { n.classList.add('is-in'); });
  } else {
    var io = new IntersectionObserver(
      function (entries) {
        entries.forEach(function (entry) {
          if (!entry.isIntersecting) return;
          entry.target.classList.add('is-in');
          io.unobserve(entry.target);
        });
      },
      { threshold: 0.12, rootMargin: '0px 0px -6% 0px' }
    );
    nodes.forEach(function (n) { io.observe(n); });
  }

  // Skills tabs (legacy support if present)
  document.querySelectorAll('[data-skill-tabs]').forEach(function (root) {
    var tabs = root.querySelectorAll('[data-skill-tab]');
    var panels = root.querySelectorAll('[data-skill-panel]');

    tabs.forEach(function (tab) {
      tab.addEventListener('click', function () {
        var id = tab.getAttribute('data-skill-tab');
        tabs.forEach(function (t) {
          var on = t === tab;
          t.classList.toggle('is-active', on);
          t.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        panels.forEach(function (panel) {
          var on = panel.getAttribute('data-skill-panel') === id;
          panel.classList.toggle('is-active', on);
          if (on) panel.removeAttribute('hidden');
          else panel.setAttribute('hidden', '');
        });
      });
    });
  });
})();
