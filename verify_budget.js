const { chromium } = require('playwright');
const SCREENSHOTS = process.env.SCREENSHOT_DIR;
const GAS_URL     = process.env.GAS_URL;
const APP_URL     = 'https://meteorkz.github.io/expense-tracker/';

const click = (page, sel) => page.evaluate(s => {
  const el = document.querySelector(s);
  if (!el) throw new Error('Not found: ' + s);
  el.click();
}, sel);

const clickNth = (page, sel, n) => page.evaluate(({s,n}) => {
  const els = document.querySelectorAll(s);
  if (!els[n]) throw new Error(`Not found: ${s}[${n}]`);
  els[n].click();
}, {s: sel, n});

const fill = (page, sel, val) => page.evaluate(({s,v}) => {
  const el = document.querySelector(s);
  el.value = v;
  el.dispatchEvent(new Event('input', {bubbles:true}));
}, {s: sel, v: String(val)});

async function addBudget(page, catIdx, amount) {
  await click(page, '#btn-add-budget');
  await page.waitForSelector('#bm-cat-grid .cat-btn', { timeout: 5000 });
  await page.waitForTimeout(300);
  await clickNth(page, '#bm-cat-grid .cat-btn', catIdx);
  await fill(page, '#bm-amt', amount);
  await click(page, '#bm-save-btn');
  await page.waitForTimeout(3500);
}

(async () => {
  const browser = await chromium.launch({ headless: true });
  const ctx  = await browser.newContext({ viewport: { width: 430, height: 932 } });
  const page = await ctx.newPage();

  // ── Load ───────────────────────────────────────────────────────
  await page.goto(APP_URL, { waitUntil: 'networkidle' });
  const hasSetup = await page.$('#setup-url');
  if (hasSetup) {
    await fill(page, '#setup-url', GAS_URL);
    await click(page, 'button.btn-primary');
    await page.waitForTimeout(4000);
  }
  await page.screenshot({ path: `${SCREENSHOTS}/01_main.png` });
  console.log('STEP 1: App loaded ✅');

  // ── Budget view ────────────────────────────────────────────────
  await click(page, '#nav-budget');
  await page.waitForTimeout(1500);
  await page.screenshot({ path: `${SCREENSHOTS}/02_budget_empty.png` });
  console.log('STEP 2: Budget view opened');

  // ── Add 3 monthly budgets ─────────────────────────────────────
  await addBudget(page, 0, 5000);  // อาหาร
  console.log('STEP 3: อาหาร ฿5,000 ✅');
  await addBudget(page, 1, 2000);  // เดินทาง
  console.log('STEP 4: เดินทาง ฿2,000 ✅');
  await addBudget(page, 3, 8000);  // ที่พัก
  await page.screenshot({ path: `${SCREENSHOTS}/03_budgets_set.png` });
  console.log('STEP 5: ที่พัก ฿8,000 ✅');

  const total = await page.$eval('#bs-total', el => el.textContent.trim());
  console.log(`  งบรวม: ${total}`);

  // ── Add expense to see progress bar ───────────────────────────
  await click(page, '#nav-main');
  await page.waitForTimeout(500);
  await click(page, '#fab');
  await page.waitForSelector('#cat-grid .cat-btn', { timeout: 3000 });
  await page.waitForTimeout(300);
  await clickNth(page, '#cat-grid .cat-btn', 0); // food
  await fill(page, '#amt-in', 3800);
  await fill(page, '#desc-in', 'ค่าอาหารทั้งเดือน');
  await click(page, '#save-btn');
  await page.waitForTimeout(4000);
  console.log('STEP 6: Added expense ฿3,800 ✅');

  // ── Back to budget — check progress ───────────────────────────
  await click(page, '#nav-budget');
  await page.waitForTimeout(800);
  await page.screenshot({ path: `${SCREENSHOTS}/04_progress.png` });

  const pcts  = await page.$$eval('.bc-pct',  els => els.map(e => e.textContent.trim()));
  const cards = await page.$$('.budget-card');
  console.log(`STEP 7: ${cards.length} cards — ${pcts.join(' | ')}`);

  // ── Edit existing budget ───────────────────────────────────────
  await clickNth(page, '.btn-set-budget', 0);
  await page.waitForSelector('#bm-amt', { timeout: 3000 });
  await page.waitForTimeout(400);
  const pre = await page.$eval('#bm-amt', el => el.value);
  console.log(`STEP 8: Edit pre-filled: ฿${pre}`);
  await fill(page, '#bm-amt', 6000);
  await click(page, '#bm-save-btn');
  await page.waitForTimeout(3000);
  await page.screenshot({ path: `${SCREENSHOTS}/05_after_edit.png` });
  const newPct = await page.$eval('.bc-pct', el => el.textContent.trim());
  console.log(`  อาหาร after edit ฿6,000: ${newPct} ✅`);

  // ── Yearly budget ──────────────────────────────────────────────
  await click(page, '#btype-yr');
  await page.waitForTimeout(400);
  await addBudget(page, 0, 72000); // food yearly
  await page.screenshot({ path: `${SCREENSHOTS}/06_yearly.png` });
  const yearTotal = await page.$eval('#bs-total', el => el.textContent.trim());
  console.log(`STEP 9: Yearly ฿72,000 set — งบรวม: ${yearTotal} ✅`);

  // ── Final monthly view ─────────────────────────────────────────
  await click(page, '#btype-mon');
  await page.waitForTimeout(400);
  await page.screenshot({ path: `${SCREENSHOTS}/07_final.png` });

  await browser.close();
  console.log(`\nScreenshots: ${SCREENSHOTS}`);
  console.log('ALL DONE ✅');
})().catch(e => { console.error('ERROR:', e.message); process.exit(1); });
