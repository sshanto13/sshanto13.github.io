from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    b = p.chromium.launch(channel='msedge')
    for w, h in ((320, 640), (375, 667), (414, 896), (768, 1024), (667, 375)):
        ctx = b.new_context(viewport={'width': w, 'height': h}, has_touch=True, is_mobile=True)
        pg = ctx.new_page()
        pg.goto('http://localhost:4321/', wait_until='networkidle')
        pg.wait_for_timeout(1500)
        r = pg.evaluate("""() => {
          const q = s => { const e = document.querySelector(s); const r = e.getBoundingClientRect(); return {top: Math.round(r.top + scrollY), bottom: Math.round(r.bottom + scrollY), left: Math.round(r.left), right: Math.round(r.right)} };
          return {social: q('.profile .social'), frame: q('.profile-frame'), kicker: q('.hero-kicker'), card: q('.hero-card'), profile: q('.profile'),
                  socialPos: getComputedStyle(document.querySelector('.profile .social')).position}
        }""")
        print(w, r)
        pg.screenshot(path=f'_prev/_hero_{w}.png', full_page=False)
        ctx.close()
    b.close()
