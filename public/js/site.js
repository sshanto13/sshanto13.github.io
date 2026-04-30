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
    var offset = nav ? nav.offsetHeight + 12 : 0;
    var top = el.getBoundingClientRect().top + window.pageYOffset - offset;
    window.scrollTo({ top: top, behavior: 'smooth' });
  });
})();

