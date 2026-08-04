import { createRequire } from 'module';
const require = createRequire('/opt/node22/lib/node_modules/');
const { chromium } = require('playwright');
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';
import { mkdirSync } from 'fs';
const dir = dirname(fileURLToPath(import.meta.url));
const out = join(dir, 'png'); mkdirSync(out, { recursive: true });
const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' });
const pg = await b.newPage({ viewport: { width: 1280, height: 720 }, deviceScaleFactor: 2 });
await pg.goto('file://' + join(dir, 'deck.html'), { waitUntil: 'networkidle' });
await pg.waitForTimeout(300);
const slides = await pg.$$('.slide');
console.log('slides:', slides.length);
for (let i = 0; i < slides.length; i++) {
  const n = String(i + 1).padStart(2, '0');
  await slides[i].screenshot({ path: join(out, `slide-${n}.png`) });
}
await b.close();
console.log('done ->', out);
