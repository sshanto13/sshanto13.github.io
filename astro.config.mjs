import { defineConfig } from 'astro/config';

export default defineConfig({
  // `site` is used for generating absolute URLs (RSS/sitemap/etc.).
  // You can change this later if you publish under a custom domain.
  site: 'https://sshanto13.github.io',

  // GitHub Pages can be hosted at "/" (user/organization site)
  // or at "/<repo>/" (project site). We compute this in CI and pass
  // it as BASE_PATH so the built site has correct asset URLs.
  base: process.env.BASE_PATH ?? '/',
  outDir: process.env.ASTRO_OUT_DIR ?? 'dist',
});

