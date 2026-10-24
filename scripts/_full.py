import subprocess, time, sys, os, glob
from playwright.sync_api import sync_playwright
srv = subprocess.Popen([sys.executable, '-m', 'http.server', '4998', '-d', 'dist'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(2)
pages = ['/'] + ['/' + os.path.relpath(p, 'dist').replace('\\', '/').rsplit('index.html', 1)[0]
                 for p in glob.glob('dist/**/index.html', recursive=True) if os.path.relpath(p, 'dist') != 'index.html']
print('pages:', pages)
OVER = """
() => {
  const vw = document.documentElement.clientWidth, bad = [];
  document.querySelectorAll('body *').forEach(el => {
    const r = el.getBoundingClientRect();
    if (!r.width || !r.height) return;
    const cs = getComputedStyle(el);
    if (cs.position === 'fixed' || cs.visibility === 'hidden' || cs.display === 'none') return;
    if (r.right > vw + 1 || r.left < -1) {
      let p = el.parentElement, sc = false;
      while (p && p !== document.body) { const o = getComputedStyle(p).overflowX; if (['auto','scroll','hidden','clip'].includes(o)) { sc = true; break; } p = p.parentElement; }
      if (!sc) bad.push(el.tagName + '.' + String(el.className).slice(0, 30) + ' R' + Math.round(r.right));
    }
  });
  const imgs = [...document.images].filter(i => i.complete && i.naturalWidth === 0 && i.src).map(i => i.src.slice(-50));
  return {sw: document.documentElement.scrollWidth - vw, bad: bad.slice(0, 5), brokenImgs: imgs};
}
"""
sizes = [(320, 640), (375, 667), (414, 896), (768, 1024), (1024, 768), (1366, 768), (1920, 1080), (667, 375), (812, 375)]
problems = 0
try:
    with sync_playwright() as p:
        b = p.chromium.launch(channel='msedge')
        for scheme in ('light', 'dark'):
            for w, h in sizes:
                ctx = b.new_context(viewport={'width': w, 'height': h}, color_scheme=scheme, has_touch=w < 1000, is_mobile=w < 900)
                for path in pages:
                    pg = ctx.new_page()
                    pg.goto('http://localhost:4998' + path, wait_until='networkidle')
                    pg.evaluate("document.querySelectorAll('img').forEach(i=>i.loading='eager')")
                    H = pg.evaluate('document.documentElement.scrollHeight'); y = 0
                    while y < H:
                        pg.evaluate(f'window.scrollTo(0,{y})'); pg.wait_for_timeout(60); y += max(h // 2, 200)
                    pg.wait_for_timeout(300)
                    r = pg.evaluate(OVER)
                    if r['sw'] > 0 or r['bad'] or r['brokenImgs']:
                        problems += 1
                        print(f'PROBLEM {scheme} {w}x{h} {path}: {r}')
                    pg.close()
                ctx.close()
            print(scheme, 'done')
        # mobile nav toggle
        ctx = b.new_context(viewport={'width': 375, 'height': 667}, has_touch=True, is_mobile=True)
        pg = ctx.new_page(); pg.goto('http://localhost:4998/', wait_until='networkidle')
        t = pg.locator('.fk-nav-toggle')
        print('nav toggle visible:', t.is_visible())
        t.click(); pg.wait_for_timeout(500)
        links = pg.locator('.fk-nav a, nav a')
        vis = pg.evaluate("[...document.querySelectorAll('nav a')].filter(a=>a.getBoundingClientRect().width>0).map(a=>[a.innerText.trim(), Math.round(a.getBoundingClientRect().left), Math.round(a.getBoundingClientRect().right)])")
        print('nav links after open:', vis)
        pg.screenshot(path='_prev/_nav.png')
        b.close()
    print('TOTAL PROBLEMS', problems)
finally:
    srv.terminate()
