# Portfolio (Astro)

This portfolio is now built with **Astro** and deployed to **GitHub Pages**.

## Development

Install dependencies:

```bash
npm install
```

Run locally:

```bash
npm run dev
```

Build a static site into `dist/`:

```bash
npm run build
```

Preview the production build:

```bash
npm run preview
```

## Deployment (GitHub Pages)

This repo includes a GitHub Actions workflow at `.github/workflows/deploy.yml` which builds and deploys `dist/` to GitHub Pages on every push to `main`.

If your repo is a **project site**, the workflow automatically builds with `BASE_PATH=/<repo>/`.
If your repo is a **user site** like `<username>.github.io`, it builds with `BASE_PATH=/`.
