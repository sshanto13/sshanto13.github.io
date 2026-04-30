import { chromium } from 'playwright';
import fs from 'node:fs/promises';
import path from 'node:path';

const targets = [
  { url: 'https://buzfi.com', out: 'buzfi.png' },
  { url: 'https://destinique.com', out: 'destinique.png' },
  { url: 'https://paimart.painest.com/', out: 'paimart.png' },
];

const outDir = path.join(process.cwd(), 'public', 'img', 'screenshots');

const viewport = { width: 1440, height: 900 };

const browser = await chromium.launch();
const context = await browser.newContext({ viewport, deviceScaleFactor: 2 });
const page = await context.newPage();

await fs.mkdir(outDir, { recursive: true });

for (const t of targets) {
  try {
    await page.goto(t.url, { waitUntil: 'networkidle', timeout: 90_000 });
    await page.waitForTimeout(1200);
    await page.screenshot({
      path: path.join(outDir, t.out),
      fullPage: true,
    });
    // eslint-disable-next-line no-console
    console.log(`Saved ${t.out}`);
  } catch (err) {
    // eslint-disable-next-line no-console
    console.error(`Failed ${t.url}`, err);
    process.exitCode = 1;
  }
}

await browser.close();

