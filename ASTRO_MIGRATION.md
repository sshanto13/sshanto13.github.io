# Astro migration plan (GitHub Pages)

This document describes how we’ll migrate your current HTML/CSS/JS portfolio (Bootstrap theme style, single `index.html`) to an Astro-based static site and deploy it to GitHub Pages.

## Goals

- Keep **your current design and content**, but improve structure and maintainability.
- Produce a **fast, SEO-friendly static site** (ideal for GitHub Pages).
- Make projects easy to update via **data-driven content** (Markdown/JSON).
- Set up **automatic deployment** to GitHub Pages on every push to `main`.

## Why Astro is a good fit for you

- **Static-first**: Astro generates HTML at build time, which is perfect for GitHub Pages.
- **Performance**: minimal JS by default; you can add interactivity only where needed.
- **Engineer-friendly**: clean component/layout structure, TypeScript support, content collections.
- **Incremental**: we can migrate your current site gradually without a full redesign.

## Current repo snapshot (what we’ll migrate)

- `index.html`: one-page portfolio with sections + many modal blocks.
- `css/suman-prorfolio.min.css` and `scss/`: theme styling (SCSS build exists via gulp).
- `js/freelancer.js`, `js/contact_me.js` and vendor scripts (jQuery + Bootstrap + popup plugin).
- `mail/contact_me.php`: PHP contact handler (not compatible with GitHub Pages runtime).

## Target approach (recommended)

### 1) Astro site structure

We’ll create a modern structure while keeping your visual identity:

- **Layouts**: base page shell (head/meta, header/footer)
- **Components**: Navbar, Hero, ProjectCard, ProjectModal, Footer, etc.
- **Content**: projects stored as Markdown (or JSON) and rendered automatically

Suggested folder layout:

```
/
  src/
    layouts/
      BaseLayout.astro
    components/
      Navbar.astro
      Hero.astro
      Projects.astro
      ProjectCard.astro
      Footer.astro
    content/
      projects/
        terbb.md
        paikaree.md
        ...
    pages/
      index.astro
  public/
    img/
    vendor/ (optional; ideally we remove old vendor libs later)
  astro.config.mjs
  package.json
```

### 2) Styling strategy (pick one, we can decide during implementation)

- **Option A (recommended)**: Tailwind CSS for quick iteration + consistent UI.
- **Option B**: keep **SCSS** and port your existing styles (Astro supports SCSS).

We can start by keeping your existing CSS (fastest) and refactor later.

### 3) JavaScript strategy

Astro doesn’t require a JS framework. For your site:

- Keep the page mostly static.
- Replace jQuery behaviors with small vanilla JS where needed (smooth scroll, modal, etc.)
- If you want a modern modal/gallery, we can use a lightweight library or build a simple accessible modal.

## GitHub Pages deployment (Astro)

We’ll deploy the built static site from `dist/` to GitHub Pages using GitHub Actions.

Key points:

- GitHub Pages sites often run under a subpath: `/<repo-name>/`
- Astro must be configured with the correct `base` in `astro.config.mjs`

We will set:

- `site`: `https://<your-username>.github.io`
- `base`: `/<repo-name>/` (example: `/sshanto13.github.io/` or `/portfolio/`)

Then add a workflow that:

- installs dependencies
- runs `npm run build`
- publishes `dist/` to GitHub Pages

## Contact form on GitHub Pages (important)

Your current contact flow posts to `mail/contact_me.php`. GitHub Pages cannot run PHP.

Recommended replacements (static-friendly):

- **Formspree**: simplest drop-in (HTML form posts to Formspree)
- **Netlify Forms**: if you deploy via Netlify (not required)
- **Cloudflare Workers / Pages Functions**: if you want your own endpoint

We’ll implement Formspree by default because it’s the fastest and works on Pages.

## Migration steps (implementation plan)

### Phase 0 — Prep (fast)

- Add Astro project scaffolding.
- Confirm GitHub Pages target:
  - repo name (important for `base`)
  - branch used for Pages (Actions-based recommended)

### Phase 1 — “Same site, Astro structure”

- Convert `index.html` → `src/pages/index.astro`
- Create `BaseLayout.astro` and shared components.
- Move `img/` into `public/img/` (Astro serves `public/` as-is).
- Keep existing CSS initially (so the look stays identical).

### Phase 2 — Data-driven projects

- Extract project entries into `src/content/projects/*.md`
- Render the projects list from content.
- Replace repeated modal HTML with a single reusable component.

### Phase 3 — Quality upgrades (quick wins)

- Add proper meta tags (description, OG/Twitter), favicon, canonical URL.
- Add `loading="lazy"` to non-critical images, define width/height where possible.
- Fix typos and improve project descriptions (remove all lorem ipsum).

### Phase 4 — Optional modernization

- Replace Bootstrap 4 + jQuery with modern CSS/components (if you want).
- Update icons (Font Awesome 6 or inline SVGs).
- Add blog (optional) using MDX.

## Acceptance checklist (what “done” looks like)

- `npm run dev` runs locally.
- `npm run build` produces a static `dist/` site.
- GitHub Actions deploys automatically to GitHub Pages.
- Site works at the GitHub Pages URL (correct base path).
- Projects are easy to update by editing Markdown files.
- Contact form works (Formspree or chosen alternative).

## What I need from you before I start (quick)

When you give the “GO”, please share:

- GitHub username (or confirm it’s `sshanto13`)
- Repo name you’ll publish from:
  - if it’s **`<user>.github.io`** (root site) OR
  - a **project repo** (subpath)
- Whether you want **Tailwind** or **keep SCSS first**
- Which contact option you want (I recommend **Formspree**)

