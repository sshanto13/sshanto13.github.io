import { chromium } from 'playwright';
import path from 'node:path';
import fs from 'node:fs/promises';

const url = 'https://paimart.painest.com/';
const outDir = path.join(process.cwd(), 'public', 'img', 'screenshots');
const outPath = path.join(outDir, 'paimart.png');

await fs.mkdir(outDir, { recursive: true });

const browser = await chromium.launch({
  args: [
    '--disable-blink-features=AutomationControlled',
    '--no-sandbox',
    '--disable-dev-shm-usage',
  ],
});

const context = await browser.newContext({
  viewport: { width: 1440, height: 900 },
  deviceScaleFactor: 2,
  userAgent:
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
  locale: 'en-US',
});

await context.addInitScript(() => {
  // Hide webdriver flag (best-effort)
  // @ts-ignore
  Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
});

const page = await context.newPage();

await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 90_000 });
await page.waitForTimeout(2500);

// Try to wait for something meaningful on the homepage.
// If blocked by a verification page, this will likely time out and we still take a screenshot.
try {
  await page.waitForSelector('text=New Arrivals', { timeout: 15_000 });
} catch {}

await page.screenshot({ path: outPath, fullPage: true });
console.log('Saved paimart.png');

await browser.close();

