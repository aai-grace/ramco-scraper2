require("dotenv").config();
const { chromium } = require("playwright");
const { mkdirSync, writeFileSync, existsSync, readFileSync, readdirSync, rmSync } = require("fs");
const { execFileSync } = require("child_process");
const { tmpdir } = require("os");
const path = require("path");
const readline = require("readline");

// Default OU shown after login. If OU_INSTANCE is empty, null, or equals
// this string, the Change Context step is skipped.
// (Note the double space between "HO" and "OPERATING" — matches the Ramco label.)
const DEFAULT_OU_INSTANCE = "AAI WORLDWIDE LOGISTICS, INC. HO  OPERATING UNIT";

// Map full OU Instance label → short folder name used under
// output/Customer Receipts/. Unknown OUs fall back to a sanitized
// PascalCase version of the label (e.g. "FOO BAR INC." → "FooBarInc").
const OU_FOLDER_MAP = {
  "AAI WORLDWIDE LOGISTICS, INC. HO  OPERATING UNIT": "AAIHO",
  "A2Z LOGISTICS INC.": "A2ZLogisticsInc",
};

function deriveOuFolder(ou) {
  const raw = (ou || DEFAULT_OU_INSTANCE).trim();
  const normalized = raw.replace(/\s+/g, ' ').toUpperCase();
  for (const [key, val] of Object.entries(OU_FOLDER_MAP)) {
    if (key.replace(/\s+/g, ' ').toUpperCase() === normalized) return val;
  }
  return raw
    .replace(/[,.]/g, '')
    .replace(/\s+/g, ' ')
    .split(' ')
    .filter(Boolean)
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join('');
}

// ─── CONFIG (loaded from .env) ─────────────────────────────────────────────
const CONFIG = {
  url: process.env.RAMCO_URL || "https://aai.ramcoes.com/RVW/extui/vwrt/LaunchPanel.htm",
  username: process.env.RAMCO_USERNAME || "",
  password: process.env.RAMCO_PASSWORD || "",
  ouInstance: (process.env.OU_INSTANCE_CUSTOMER_RECEIPT || "").trim(),
  headless: false,
  slowMo: 50,
  timeout: 30000,

  // View Receipt search filters (dd/mm/yyyy — same format shown in the Ramco UI).
  // Ramco labels these "Receipt Date From / To" on the Select Receipt screen
  // (not "Transaction Date" like the Journal Voucher / Sundry screens).
  receiptDateFrom: process.env.RECEIPT_DATE_FROM_CUSTOMER_RECEIPT || "01/04/2026",
  receiptDateTo: process.env.RECEIPT_DATE_TO_CUSTOMER_RECEIPT || "23/04/2026",

  instrumentDateFrom: (process.env.INSTRUMENT_DATE_FROM_CUSTOMER_RECEIPT || "").trim(),
  instrumentDateTo: (process.env.INSTRUMENT_DATE_TO_CUSTOMER_RECEIPT || "").trim(),

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // receipt number and begins processing from there. Blank = start at row 1.
  resumeFromReceipt: (process.env.RESUME_FROM_CUSTOMER_RECEIPT || "").trim(),

  // Single-receipt mode: if set, the Excel download/list is bypassed entirely
  // and only this exact Receipt No. is processed via Direct Entry. Takes
  // precedence over RESUME_FROM_CUSTOMER_RECEIPT.
  onlyVoucher: (process.env.ONLY_CUSTOMER_RECEIPT || "").trim(),
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Customer Receipts/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Customer Receipts/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Customer Receipt workflow:
// Receivable Management → Customer Receipt → View Receipt
const MENU_PATHS = [
  ["Receivable Management", "Customer Receipt", "View Receipt"],
];

// ─── HELPERS ───────────────────────────────────────────────────────────────
function waitForEnter(msg = "👉 Press ENTER to close the browser...") {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => {
    rl.question("\n" + msg + "\n", () => {
      rl.close();
      resolve();
    });
  });
}

async function askWithCountdown(query, timeoutSec) {
  return new Promise((resolve) => {
    let secondsLeft = timeoutSec;
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });

    const updatePrompt = () => {
      readline.clearLine(process.stdout, 0);
      readline.cursorTo(process.stdout, 0);
      process.stdout.write(`${query} [Auto-Y in ${secondsLeft}s]: `);
    };

    updatePrompt();

    const interval = setInterval(() => {
      secondsLeft--;
      updatePrompt();
      if (secondsLeft <= 0) {
        clearInterval(interval);
        rl.close();
        process.stdout.write("\n");
        resolve("y");
      }
    }, 1000);

    rl.question("", (ans) => {
      clearInterval(interval);
      rl.close();
      resolve(ans || "y");
    });
  });
}

async function extractTableData(frame) {
  return await frame.evaluate(() => {
    const tables = Array.from(document.querySelectorAll("table"));
    return tables.map((table) => {
      const rows = Array.from(table.querySelectorAll("tr"));
      return rows.map((row) =>
        Array.from(row.querySelectorAll("th, td")).map((cell) => cell.innerText.trim())
      );
    });
  });
}

// ─── LOGIN ─────────────────────────────────────────────────────────────────
async function login(page) {
  console.log("📂 Navigating to Ramco VirtualWorks...");
  await page.goto(CONFIG.url, { waitUntil: "domcontentloaded" });

  console.log("  ⏳ Waiting for app shell to initialize...");
  await page.waitForTimeout(4000);

  console.log("  🔍 Scanning frames for login form...");
  let loginFrame = null;

  for (let attempt = 0; attempt < 12; attempt++) {
    const frames = page.frames();
    console.log(`  [attempt ${attempt + 1}] frames: ${frames.length}`);
    for (const frame of frames) {
      const url = frame.url();
      if (url === CONFIG.url || url === "about:blank" || url.includes("rvw_blank")) continue;
      const inputs = await frame.$$("input").catch(() => []);
      if (inputs.length > 0) {
        console.log(`  ✔ Login frame: name="${frame.name()}" url="${url}"`);
        loginFrame = frame;
        break;
      }
    }
    if (loginFrame) break;
    await page.waitForTimeout(2000);
  }

  if (!loginFrame) {
    for (const frame of page.frames()) {
      if (frame.name() === "ilboinnerframe") {
        const inputs = await frame.$$("input").catch(() => []);
        if (inputs.length > 0) loginFrame = frame;
      }
    }
  }

  if (!loginFrame) {
    throw new Error("❌ Could not find login frame.");
  }

  const userSel = "#ide_username";
  const pwSel = "#ide_password";

  await loginFrame.waitForSelector(userSel, { state: "attached", timeout: CONFIG.timeout });
  await loginFrame.waitForSelector(pwSel, { state: "attached", timeout: CONFIG.timeout });

  console.log("  ✏ Typing username...");
  await loginFrame.focus(userSel);
  await loginFrame.fill(userSel, "");
  await loginFrame.type(userSel, CONFIG.username, { delay: 30 });
  await page.waitForTimeout(200);

  console.log("  ✏ Typing password...");
  await loginFrame.focus(pwSel);
  await loginFrame.fill(pwSel, "");
  await loginFrame.type(pwSel, CONFIG.password, { delay: 30 });
  await page.waitForTimeout(300);

  console.log("  🖱 Pressing Enter to submit...");
  await loginFrame.press(pwSel, "Enter");
  await page.waitForTimeout(1000);

  const stillOnLogin = page.frames().some(f => f.url().includes("security/login"));
  if (stillOnLogin) {
    console.log("  ⚠ Still on login — trying JS anchor click...");
    await loginFrame.evaluate((pwSel) => {
      const anchors = Array.from(document.querySelectorAll("a"));
      const btn = anchors.find(
        a => a.offsetParent !== null && !a.id.includes("forgot") && !/forgot/i.test(a.innerText)
      );
      if (btn) { btn.click(); return; }
      const form = document.querySelector("form");
      if (form) { form.submit(); return; }
      const pw = document.querySelector(pwSel);
      if (pw) ["keydown", "keypress", "keyup"].forEach(t =>
        pw.dispatchEvent(new KeyboardEvent(t, { key: "Enter", keyCode: 13, bubbles: true }))
      );
    }, pwSel);
    await page.waitForTimeout(2000);
  }

  console.log("✅ Login complete.");
}

// ─── CHANGE CONTEXT (User Menu → Change Context → SwitchContexts) ──────────
async function changeContext(page, ouInstance) {
  const target = (ouInstance || "").trim();

  if (!target || target === DEFAULT_OU_INSTANCE) {
    console.log(`\n🔄 OU Instance unchanged (using default: "${DEFAULT_OU_INSTANCE}").`);
    return;
  }

  console.log(`\n🔄 Switching OU Instance → "${target}"...`);

  // Wait for the user toolbar to be ready (it appears after login).
  await page.waitForSelector('#userinfo', { timeout: CONFIG.timeout });
  await page.waitForTimeout(800);

  // 1) Open the User menu.
  console.log('  🖱 Opening User Menu (#userinfo)...');
  await page.locator('#userinfo').first().click();
  await page.waitForTimeout(500);

  // 2) Click "Change Context".
  console.log('  🖱 Clicking "Change Context"...');
  const changeCtxItem = page.locator(
    '.x-menu-item-text:has-text("Change Context"), [qtip="Change Context"], a:has-text("Change Context")'
  ).first();
  await changeCtxItem.waitFor({ state: 'visible', timeout: CONFIG.timeout });
  await changeCtxItem.click();

  // 3) Wait for the SwitchContexts dialog to appear. It can render in the
  //    top page OR inside one of the frames — find whichever scope contains
  //    both the title and an input near an "OU Instance" text node.
  console.log('  ⏳ Waiting for SwitchContexts dialog...');
  await page.waitForTimeout(1500);

  const scopes = [page, ...page.frames()];
  let scope = null;
  let ouInputHandle = null;

  for (let attempt = 0; attempt < 10 && !ouInputHandle; attempt++) {
    for (const s of scopes) {
      try {
        const handle = await s.evaluateHandle(() => {
          // Walk every element and find one whose own (non-descendant) text
          // is exactly or starts with "OU Instance". ExtJS renders labels as
          // <div>/<td> with plain text — not <label> elements.
          const all = document.querySelectorAll('div, td, span, label');
          let labelEl = null;
          for (const el of all) {
            const own = Array.from(el.childNodes)
              .filter(n => n.nodeType === Node.TEXT_NODE)
              .map(n => n.textContent.trim())
              .join(' ')
              .trim();
            if (own === 'OU Instance' || /^OU\s+Instance$/i.test(own)) {
              labelEl = el;
              break;
            }
          }
          if (!labelEl) return null;

          // From the label, walk up to the row container that also holds
          // the input, then return the first input inside it.
          let container = labelEl.parentElement;
          for (let i = 0; i < 6 && container; i++) {
            const inp = container.querySelector('input[type="text"], input:not([type])');
            if (inp && !labelEl.contains(inp)) return inp;
            container = container.parentElement;
          }
          // Last-ditch: scan following siblings.
          let next = labelEl.nextElementSibling;
          while (next) {
            if (next.tagName === 'INPUT') return next;
            const inp = next.querySelector && next.querySelector('input');
            if (inp) return inp;
            next = next.nextElementSibling;
          }
          return null;
        });

        const el = handle.asElement();
        if (el) {
          const visible = await el.isVisible().catch(() => false);
          if (visible) {
            scope = s;
            ouInputHandle = el;
            break;
          }
        }
      } catch (_) { /* frame might be detached — try the next one */ }
    }
    if (!ouInputHandle) await page.waitForTimeout(500);
  }

  if (!ouInputHandle) {
    throw new Error("❌ Could not locate the OU Instance input in the SwitchContexts dialog.");
  }

  // 4) Fill OU Instance: focus, clear, type, and pick the matching dropdown
  //    item (or fall back to pressing Enter).
  console.log('  ✏ Filling OU Instance...');
  await ouInputHandle.scrollIntoViewIfNeeded().catch(() => { });
  await ouInputHandle.click({ clickCount: 3 }).catch(async () => { await ouInputHandle.focus(); });
  await page.keyboard.press('Backspace');
  await page.keyboard.type(target, { delay: 30 });
  await page.waitForTimeout(900);

  const dropdownItem = scope.locator(
    `.x-combo-list-item:has-text("${target}"), li:has-text("${target}")`
  ).first();
  if (await dropdownItem.isVisible({ timeout: 1500 }).catch(() => false)) {
    await dropdownItem.click();
  } else {
    await page.keyboard.press('Enter');
  }
  await page.waitForTimeout(500);

  // 5) Click OK to apply the context switch.
  console.log('  🖱 Clicking OK...');
  const okBtn = scope.locator(
    'button:has-text("OK"), .x-btn-text:has-text("OK"), input[type="button"][value="OK"]'
  ).first();
  if (await okBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
    await okBtn.click();
  } else {
    await page.keyboard.press('Enter');
  }

  // 6) Wait for the toolbar to reflect the new OU.
  console.log('  ⏳ Waiting for context to switch...');
  await page.waitForTimeout(4500);
  await page.waitForSelector('#userinfo', { timeout: CONFIG.timeout }).catch(() => { });

  const userInfoText = await page.locator('#userinfo').first().innerText().catch(() => "");
  if (userInfoText && userInfoText.toUpperCase().includes(target.toUpperCase())) {
    console.log(`✅ OU Instance switched to "${target}".`);
  } else {
    console.warn(`  ⚠ Could not verify OU switch from toolbar text: "${userInfoText.replace(/\s+/g, ' ').trim()}"`);
  }
}

// ─── GO TO LANDING PAGE ────────────────────────────────────────────────────
async function goToLandingPage(page) {
  console.log("\n🏠 Navigating to landing/home page...");
  await page.waitForTimeout(3000);

  const homeSelectors = [
    "#tb-homepage button",
    "button.x-icon-uhome",
    'div[qtip="Go to Home Page"]',
    'div.x-tool-uhome',
    '#home button',
  ];

  for (const sel of homeSelectors) {
    try {
      const el = page.locator(sel).first();
      if (await el.isVisible({ timeout: 2000 }).catch(() => false)) {
        await el.click();
        console.log(`  ✔ Clicked home: ${sel}`);
        await page.waitForSelector('#tb-bpc-menu', { timeout: 8000 }).catch(() => { });
        break;
      }
    } catch (_) { }
  }

  await page.waitForSelector('#tb-bpc-menu', { timeout: 10000 }).catch(() => { });
}

// ─── NAVIGATE MENU ─────────────────────────────────────────────────────────
async function navigateMenu(page, menuPath) {
  console.log(`\n📋 Navigating: ${menuPath.join(" → ")}`);

  const isMenuCollapsed = await page.evaluate(() => {
    const tp = document.getElementById('ilbo-tp');
    const collapsedBar = document.getElementById('ilbo-tp-xcollapsed');
    return !tp || tp.classList.contains('x-panel-collapsed') || (collapsedBar && window.getComputedStyle(collapsedBar).display !== 'none');
  });

  if (isMenuCollapsed) {
    console.log("  📂 Menu is collapsed. Expanding...");

    const expandSelectors = [
      '#tb-bpc-menu button',
      '#ext-gen309',
      '.x-tool-expand-west',
    ];

    for (const sel of expandSelectors) {
      const btn = page.locator(sel).first();
      if (await btn.isVisible({ timeout: 2000 }).catch(() => false)) {
        await btn.click();
        await page.waitForTimeout(2000);
        break;
      }
    }
  }

  for (let i = 0; i < menuPath.length; i++) {
    const label = menuPath[i];
    console.log(`  [${i + 1}/${menuPath.length}] Looking for "${label}"...`);

    let clicked = false;

    const selectors = [
      `.x-menu-item-text:has-text("${label}")`,
      `.x-menu-item:has-text("${label}")`,
      `.x-tree-node-text:has-text("${label}")`,
      `.x-tree-node-anchor:has-text("${label}")`,
      `span:has-text("${label}")`,
      `a:has-text("${label}")`,
      `div:has-text("${label}")`,
    ];

    for (const sel of selectors) {
      try {
        const el = page.locator(sel).filter({ hasText: new RegExp(`^\\s*${label}\\s*$`, 'i') }).first();

        if (await el.isVisible({ timeout: 3000 }).catch(() => false)) {
          await el.scrollIntoViewIfNeeded().catch(() => { });

          if (i < menuPath.length - 1) {
            console.log(`    🖱 Hovering "${label}"...`);
            await el.hover({ force: true }).catch(() => { });
            await page.waitForTimeout(400);
          }

          console.log(`    🖱 Clicking animated "${label}"...`);
          await el.click();
          await page.waitForTimeout(1000);
          console.log(`    ✔ Clicked "${label}" via: ${sel}`);
          clicked = true;
          break;
        }
      } catch (_) { }
    }

    if (!clicked) {
      console.warn(`    ⚠ Exact match for "${label}" failed. Trying generic visibility...`);
      try {
        const looseEl = page.getByText(label).filter({ visible: true }).first();
        if (await looseEl.isVisible({ timeout: 2000 }).catch(() => false)) {
          if (i < menuPath.length - 1) {
            await looseEl.hover().catch(() => { });
            await page.waitForTimeout(1000);
          }
          await looseEl.click();
          await page.waitForTimeout(2500);
          console.log(`    ✔ Clicked "${label}" via visibility filter`);
          clicked = true;
        }
      } catch (_) { }
    }

    if (!clicked) {
      console.error(`    ❌ FAILED to find "${label}"`);

      if (i === menuPath.length - 1) {
        console.log("    🔍 Attempting search box fallback for final item...");
        await page.fill('#bpcsearch', label);
        await page.waitForTimeout(500);
        await page.press('#bpcsearch', 'Enter');

        await page.waitForTimeout(4000);
        console.log("    ✔ Search submitted.");
        clicked = true;
      }
    }
  }

  console.log("  ⏳ Waiting for page to load...");
  await page.waitForLoadState("networkidle").catch(() => { });
  await page.waitForTimeout(3000);

  const screenshotPath = `screenshots/result-${menuPath.join("-").replace(/[\s/]+/g, "_")}.png`;

  return { menuPath, url: page.url(), screenshot: screenshotPath };
}

async function getFrame(page) {
  const frame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('ilrt/rvw_blank.aspx') || f.url().includes('ilbo'));
  if (!frame) {
    return page.mainFrame().childFrames()[0];
  }
  return frame;
}

/**
 * Orchestrates the Search workflow for Sundry Payment "Select Voucher" screen.
 * (See Image 2 of Sundry Expense Invoice_Voucher.docx — filters by
 *  Pay Date From/To and Request Date From/To, then clicks Search.)
 */
async function performVoucherSearch(page) {
  console.log("\n🔍 Preparing to search vouchers inside frame...");
  const frame = await getFrame(page);

  if (!frame) {
    throw new Error("Could not find ilboinnerframe");
  }

  try {
    // 1. Clear "User ID" field (same pattern as invoice search)
    console.log('  ⌨ Attempting to clear "User ID" field...');

    const userIdField = frame.locator('input').filter({ hasText: "" }).and(
      frame.locator('xpath=//label[contains(text(), "User ID")]/following::input[1] | //td[contains(text(), "User ID")]/following::input[1] | //span[contains(text(), "User ID")]/following::input[1]')
    ).first();

    if (await userIdField.isVisible({ timeout: 5000 }).catch(() => false)) {
      console.log('  ✅ User ID field found via label.');
      await userIdField.click();
      await page.keyboard.down('Control');
      await page.keyboard.press('a');
      await page.keyboard.up('Control');
      await page.keyboard.press('Backspace');
      await userIdField.fill("");
      console.log('  ✅ User ID cleared.');
    } else {
      console.warn('  ⚠️ Could not identify exact "User ID" input via label. Trying value-based search...');
      const fallbackField = frame.locator('input[value*="CVL"]').first();
      if (await fallbackField.isVisible({ timeout: 2000 }).catch(() => false)) {
        await fallbackField.click();
        await page.keyboard.down('Control');
        await page.keyboard.press('a');
        await page.keyboard.up('Control');
        await page.keyboard.press('Backspace');
        console.log('  ✅ User ID cleared (value fallback).');
      } else {
        console.warn('  ⚠️ Attempting generic keyboard clear as last resort...');
        await frame.locator('input').first().focus().catch(() => { });
        await page.keyboard.down('Control');
        await page.keyboard.press('a');
        await page.keyboard.up('Control');
        await page.keyboard.press('Backspace');
      }
    }

    // 2. Fill "Pay Date From/To" and "Request Date From/To" filters
    //    Each row in the UI is: [<Label>From] [<input From>] [label "To"] [<input To>]
    //    so following::input[1] = From field, following::input[2] = To field.
    const fillDateField = async (labelText, position, value) => {
      const field = frame.locator(
        `xpath=//label[contains(text(), "${labelText}")]/following::input[${position}] | ` +
        `//td[contains(text(), "${labelText}")]/following::input[${position}] | ` +
        `//span[contains(text(), "${labelText}")]/following::input[${position}]`
      ).first();

      if (await field.isVisible({ timeout: 3000 }).catch(() => false)) {
        await field.click();
        await page.keyboard.down('Control');
        await page.keyboard.press('a');
        await page.keyboard.up('Control');
        await page.keyboard.press('Delete');
        await field.fill(value);
        await page.keyboard.press('Tab'); // commit the value so the grid filter picks it up
        return true;
      }
      console.warn(`  ⚠️ Date field not found for label "${labelText}" (position ${position}).`);
      return false;
    };

    console.log(`  📅 Filling Receipt Date: ${CONFIG.receiptDateFrom} → ${CONFIG.receiptDateTo}`);
    await fillDateField("Receipt Date From", 1, CONFIG.receiptDateFrom);
    await fillDateField("Receipt Date From", 2, CONFIG.receiptDateTo);
    console.log('  ✅ Receipt Date range set.');

    console.log(`  📅 Filling Instrument Date: ${CONFIG.instrumentDateFrom} → ${CONFIG.instrumentDateTo}`);
    await fillDateField("Instrument Date From", 1, CONFIG.instrumentDateFrom);
    await fillDateField("Instrument Date From", 2, CONFIG.instrumentDateTo);
    console.log('  ✅ Intrument Date range set.');

    // 3. Click the Search button
    console.log('  🖱 Clicking "Search" button...');
    const searchBtn = frame.locator('button, .x-btn-text').filter({ hasText: /^Search$/i }).last();

    if (await searchBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await searchBtn.click();
      console.log("  🚀 Search triggered.");

      await frame.locator('.x-grid3-row').first().waitFor({ state: "visible", timeout: 15000 }).catch(() => {
        console.log("  ⏳ Still waiting for grid rows...");
      });

      const closeBtn = frame.locator('button').filter({ hasText: /^Close$/i }).first();
      if (await closeBtn.isVisible({ timeout: 1500 }).catch(() => false)) {
        console.warn("  ⚠ Search message popup detected. Closing it...");
        await closeBtn.click();
        await page.waitForTimeout(500);
      }
    } else {
      console.error("  ❌ Search button not found in frame.");
    }

    console.log("  📊 Extracting Grid metadata...");
    const stats = await frame.evaluate(() => {
      const pagingText = document.querySelector('.x-paging-info')?.innerText || "";
      const match = pagingText.match(/(\d+)\s*$/);
      const total = match ? parseInt(match[1]) : 0;

      const visibleRows = document.querySelectorAll('.x-grid3-row').length;

      return { total, visibleRows };
    });

    console.log(`  ✔ Grid detected: Total ${stats.total} rows, ${stats.visibleRows} on current page.`);
    return stats;

  } catch (error) {
    console.error(`  ❌ Search automation failed: ${error.message}`);
    return { total: 0, visibleRows: 0 };
  }
}


/**
 * Receipt-number text pattern, e.g. "000409568".
 * Customer Receipts use a zero-padded purely-numeric ID (typically 9 digits)
 * — different from the "02XXNN-NNNNNN" format used by vouchers / invoices.
 * Allow 6+ digits to be safe across OUs.
 */
const VOUCHER_NO_PATTERN = /^\d{6,}$/;

/**
 * Enumerates every Receipt No. currently rendered in the Select Receipt
 * search-results grid. ExtJS grids render "hyperlinks" as styled <span>/<div>
 * elements, not <a> tags — so we scan grid-cell text and match the
 * receipt-number pattern instead of relying on the anchor tag.
 *
 * Returns an array of unique receipt numbers (e.g. ["02CR26-000658", "02CR26-000660"]).
 */
async function getVoucherNumbers(page) {
  console.log("\n🔢 Collecting receipt numbers from the grid...");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  if (!actualFrame) {
    console.error("  ❌ Could not find ilboinnerframe.");
    return [];
  }

  // The Select Voucher grid is a real paginated ExtJS grid
  // (toolbar: |< < 1 > >|, "Displaying 1 - 5 of 7"). Rows beyond the
  // current page are NOT in the DOM, so we must step through the pager
  // and scan each page in turn.
  const total = await actualFrame.evaluate(() => {
    const txt = document.querySelector('.x-paging-info')?.innerText || "";
    const m = txt.match(/(\d+)\s*$/);
    return m ? parseInt(m[1], 10) : 0;
  });
  console.log(`  📊 Paging toolbar reports total: ${total || "unknown"}`);

  const seen = new Set();
  const out = [];

  const scanCurrentPage = async () => {
    const batch = await actualFrame.evaluate((patternSrc) => {
      const pattern = new RegExp(patternSrc);
      const found = [];
      document.querySelectorAll('.x-grid3-row').forEach(row => {
        const cells = row.querySelectorAll('.x-grid3-cell-inner, .x-grid3-cell, td, a, span, div');
        for (const cell of cells) {
          const txt = (cell.innerText || cell.textContent || "").trim();
          if (pattern.test(txt)) { found.push(txt); break; }
        }
      });
      return found;
    }, VOUCHER_NO_PATTERN.source);

    let added = 0;
    for (const v of batch) {
      if (!seen.has(v)) { seen.add(v); out.push(v); added++; }
    }
    return added;
  };


  // Jump to the first page so collection starts deterministically at row 1.
  await actualFrame.evaluate(() => {
    const firstBtn = document.querySelector('.x-tbar-page-first');
    if (firstBtn && !firstBtn.classList.contains('x-item-disabled')) firstBtn.click();
  });
  await page.waitForTimeout(500);

  // Iterate pages. The pager "Next" button gets `x-item-disabled` on the
  // last page, which is our termination signal.
  let safety = 50;
  let pageNum = 1;
  while (safety-- > 0) {
    const added = await scanCurrentPage();
    console.log(`    📄 Page ${pageNum}: +${added} new voucher(s) (running total ${out.length}${total ? `/${total}` : ""})`);

    if (total > 0 && out.length >= total) break;

    const pagingBefore = await actualFrame.evaluate(() =>
      document.querySelector('.x-paging-info')?.innerText || ""
    );

    const clickedNext = await actualFrame.evaluate(() => {
      const nextBtn = document.querySelector('.x-tbar-page-next');
      if (!nextBtn) return false;
      if (nextBtn.classList.contains('x-item-disabled')) return false;
      nextBtn.click();
      return true;
    });
    if (!clickedNext) break;

    // Wait for the paging info text to change (confirms the page actually advanced).
    try {
      await actualFrame.waitForFunction(
        (prev) => (document.querySelector('.x-paging-info')?.innerText || "") !== prev,
        pagingBefore,
        { timeout: 5000 }
      );
    } catch (_) {
      // No change → grid didn't advance; bail out.
      break;
    }
    await page.waitForTimeout(300);
    pageNum++;
  }

  // Return to page 1 so the downstream row-click logic starts in a known state.
  await actualFrame.evaluate(() => {
    const firstBtn = document.querySelector('.x-tbar-page-first');
    if (firstBtn && !firstBtn.classList.contains('x-item-disabled')) firstBtn.click();
  });
  await page.waitForTimeout(500);

  if (total && out.length < total) {
    console.warn(`  ⚠ Collected ${out.length} of ${total} voucher(s).`);
  }
  console.log(`  ✔ Collected ${out.length} voucher(s)${total ? ` (of ${total} total)` : ""}: ${out.join(", ")}`);
  return out;
}

/**
 * Scans every frame for a voucher number belonging to the currently-open F2
 * Direct Entry popup. Priority (high → low):
 *   1) document.activeElement — the F2 popup input is focused immediately
 *      after F2 is pressed, so this is the most reliable source.
 *   2) Inputs / text inside visible ExtJS dialogs (.x-window, .x-dlg, etc.).
 *   3) Any visible <input>/<textarea> on the page (last resort — will match
 *      the Direct Entry "Voucher No." field even when stale).
 *
 * Without priorities 1 and 2, row 2+ would read the stale Direct Entry field
 * (still holding row 1's voucher) and we'd open the same voucher every time.
 */
async function extractVoucherFromPopup(page) {
  const patternSrc = VOUCHER_NO_PATTERN.source.replace(/^\^|\$$/g, ''); // match anywhere

  for (const f of page.frames()) {
    try {
      const val = await f.evaluate((src) => {
        const pattern = new RegExp(src);
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== 'none' && s.visibility !== 'hidden';
        };

        // 1) Focused input — F2 popup field is focused right after F2
        const active = document.activeElement;
        if (active && (active.tagName === 'INPUT' || active.tagName === 'TEXTAREA') && isVisible(active)) {
          const v = (active.value || '').trim();
          const m = v.match(pattern);
          if (m) return m[0];
        }

        // 2) Inputs / text inside visible ExtJS dialog containers
        const popups = document.querySelectorAll('.x-window, .x-dlg, [role="dialog"], .x-layer');
        for (const popup of popups) {
          if (!isVisible(popup)) continue;
          for (const inp of popup.querySelectorAll('input, textarea')) {
            if (!isVisible(inp)) continue;
            const v = (inp.value || '').trim();
            const m = v.match(pattern);
            if (m) return m[0];
          }
          const text = (popup.innerText || '').trim();
          const m = text.match(pattern);
          if (m) return m[0];
        }

        // 3) Fallback — any visible input (may be stale Direct Entry value)
        for (const inp of document.querySelectorAll('input, textarea')) {
          if (!isVisible(inp)) continue;
          const v = (inp.value || '').trim();
          const m = v.match(pattern);
          if (m) return m[0];
        }

        return null;
      }, patternSrc);

      if (val) return val;
    } catch (_) { }
  }
  return null;
}

/**
 * Opens a voucher's detail screen using the F2 Direct Entry shortcut, per the
 * user's annotated SOP:
 *   1. Press F2 (row already highlighted post-search) → popup shows voucher number
 *   2. Copy voucher number from popup → Escape the popup
 *   3. Paste into "Voucher No." Direct Entry field
 *   4. Click "View Voucher" to trigger navigation
 *
 * Row 1 is already highlighted after the search, so no row click is needed.
 * For rows 2+, we click the row at the given index to highlight it first.
 *
 * Returns { ok: boolean, voucherNo: string|null }.
 */
async function openVoucherViaF2(page, rowIndex, knownVoucherNo = null, forceDirectEntry = false) {
  console.log(`\n🖱 Opening voucher via F2 — row ${rowIndex + 1}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  // If the target row isn't on the current grid page (e.g. row 6 when page 1
  // only holds rows 1-5) we can't click it or F2 it. Fall back to pasting
  // the already-collected voucher number directly into Direct Entry.
  //
  // `forceDirectEntry=true` bypasses the grid-click + F2 path even when the
  // row index technically falls inside rowsInDom. Callers set this when the
  // voucher number was already confirmed externally (e.g. the ArrowDown
  // continuation) and a second F2 would risk reading stale data — also
  // important because ExtJS grids with locked columns double the `.x-grid3-row`
  // count so the plain index comparison can't distinguish on-page vs off-page.
  const rowsInDom = await frame.locator('.x-grid3-row').count().catch(() => 0);
  const rowOnCurrentPage = !forceDirectEntry && rowIndex < rowsInDom;

  let voucherNo;

  if (!rowOnCurrentPage) {
    if (!knownVoucherNo) {
      console.error(`  ❌ Row ${rowIndex + 1} not on current grid page and no known voucher number to fall back on.`);
      return { ok: false, voucherNo: null };
    }
    if (forceDirectEntry) {
      console.log(`  ⏭ Using Direct Entry (forced) with known voucher ${knownVoucherNo}.`);
    } else {
      console.log(`  ⏭ Row ${rowIndex + 1} not on current grid page — bypassing F2 and using known voucher ${knownVoucherNo}.`);
    }
    voucherNo = knownVoucherNo;
  } else {
    // 1. Ensure the target row is highlighted (skip for row 1)
    if (rowIndex === 0) {
      console.log(`  ℹ Row 1 already highlighted — skipping click.`);
    } else {
      console.log(`  👆 Highlighting row ${rowIndex + 1}...`);
      let highlighted = false;
      for (const sel of ['.x-grid3-row', '.x-grid-row', 'tr[role="row"]']) {
        try {
          const row = frame.locator(sel).nth(rowIndex);
          if (await row.isVisible({ timeout: 2000 }).catch(() => false)) {
            await row.click({ force: true });
            highlighted = true;
            break;
          }
        } catch (_) { }
      }
      if (!highlighted) {
        console.warn(`  ⚠ Could not highlight row ${rowIndex + 1}.`);
        return { ok: false, voucherNo: null };
      }
      await page.waitForTimeout(500);

      // Click alone doesn't move keyboard focus into the grid row on row 2+.
      // One Tab press shifts focus so F2 actually triggers the Direct Entry popup.
      console.log(`  ⌨ Pressing Tab to move keyboard focus into the grid...`);
      await page.keyboard.press('Tab');
      await page.waitForTimeout(300);
    }

    // 2. Press F2 → popup opens with voucher number pre-filled
    console.log(`  ⌨ Pressing F2...`);
    await page.keyboard.press('F2');
    await page.waitForTimeout(1200);

    // 3. Copy voucher number from popup, then dismiss the popup
    voucherNo = await extractVoucherFromPopup(page);
    if (!voucherNo) {
      // Popup extraction failed — but if we already know the voucher from the
      // grid scan, we can still proceed by pasting it directly.
      if (knownVoucherNo) {
        console.warn(`  ⚠ Popup read failed; falling back to known voucher ${knownVoucherNo}.`);
        voucherNo = knownVoucherNo;
      } else {
        console.error(`  ❌ Could not read voucher number from F2 popup.`);
        await page.keyboard.press('Escape').catch(() => { });
        return { ok: false, voucherNo: null };
      }
    }
    console.log(`  📋 Copied voucher number: ${voucherNo}`);

    await page.keyboard.press('Escape');
    await page.waitForTimeout(500);
  }

  // 4. Paste into "Receipt No." Direct Entry field (first match — excludes
  //    the "Receipt No. From" search-criteria field because of exact-text match).
  console.log(`  ⌨ Pasting into Direct Entry "Receipt No." field...`);
  const voucherInput = frame
    .locator('xpath=//*[normalize-space(text())="Receipt No."]/following::input[1]')
    .first();
  try {
    await voucherInput.waitFor({ state: "visible", timeout: 5000 });
    await voucherInput.click();
    await voucherInput.fill("");
    await voucherInput.fill(voucherNo);
    console.log(`  ✔ Pasted ${voucherNo}.`);
  } catch (e) {
    console.error(`  ❌ Failed to paste into Direct Entry: ${e.message}`);
    return { ok: false, voucherNo };
  }

  // 5. Trigger the search — click the "View Receipt" hyperlink.
  //    Strategy A (primary): JS-dispatch a real mouse click on the innermost
  //    visible element whose *direct* text node is exactly "View Receipt".
  //    Strategy B (fallback): Playwright getByText with exact match.
  console.log(`  🖱 Clicking "View Receipt"...`);
  let clicked = false;

  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  if (actualFrame) {
    try {
      clicked = await actualFrame.evaluate(() => {
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== 'none' && s.visibility !== 'hidden';
        };
        for (const el of document.querySelectorAll('*')) {
          if (!isVisible(el)) continue;
          // Only element whose *own* text nodes equal "View Receipt"
          const directText = Array.from(el.childNodes)
            .filter(n => n.nodeType === Node.TEXT_NODE)
            .map(n => n.textContent.trim())
            .join('');
          if (directText !== 'View Receipt') continue;
          const rect = el.getBoundingClientRect();
          const x = rect.left + rect.width / 2;
          const y = rect.top + rect.height / 2;
          ['mousedown', 'mouseup', 'click'].forEach(t => {
            el.dispatchEvent(new MouseEvent(t, {
              bubbles: true, cancelable: true, view: window,
              clientX: x, clientY: y, button: 0,
            }));
          });
          return true;
        }
        return false;
      });
      if (clicked) console.log(`  ✔ Clicked "View Receipt" via JS dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      const viewLink = frame.getByText('View Receipt', { exact: true }).first();
      await viewLink.waitFor({ state: "visible", timeout: 5000 });
      await viewLink.click({ force: true });
      clicked = true;
      console.log(`  ✔ Clicked "View Receipt" via locator.`);
    } catch (e) {
      console.error(`  ❌ Failed to click View Receipt: ${e.message}`);
      return { ok: false, voucherNo };
    }
  }

  // Poll for one of two outcomes after the View Receipt click:
  //   (a) detail page attaches → break and continue normally
  //   (b) "Receipt Number does not exist" popup (#message-panel ExtJS window)
  //        appears → close it, retry once, then skip if it still fails.
  //
  // The popup renders as a top-level ExtJS window (<div id="message-panel">)
  // which may live in the main page OR in the ilboinnerframe — so we check
  // both contexts on each poll tick.
  const findErrorPopup = async () => {
    for (const root of [page, frame]) {
      const panel = root.locator('#message-panel').first();
      const visible = await panel.isVisible({ timeout: 0 }).catch(() => false);
      if (!visible) continue;
      const text = await panel.innerText().catch(() => "");
      if (/does not exist/i.test(text)) return panel;
    }
    return null;
  };

  const isDetailVisible = async () => {
    // "Receipt Information" is a section header unique to the View Receipt
    // detail page — it never appears on Select Receipt, so its presence is
    // a reliable signal the detail screen has attached.
    return await frame.locator('text="Receipt Information"').first()
      .isVisible({ timeout: 0 })
      .catch(() => false);
  };

  const dismissPopup = async (panel) => {
    // Primary: Close button inside the panel (HTML shows <button class="x-btn-text">Close</button>).
    const closeBtn = panel.locator('button').filter({ hasText: /^Close$/i }).first();
    await closeBtn.click({ force: true }).catch(() => { });
    await page.waitForTimeout(300);
    // Backup: top-right X icon (.x-tool-close), in case Close didn't register.
    if (await panel.isVisible({ timeout: 0 }).catch(() => false)) {
      await panel.locator('.x-tool-close').first().click({ force: true }).catch(() => { });
      await page.waitForTimeout(300);
    }
    // Final fallback: Escape key.
    if (await panel.isVisible({ timeout: 0 }).catch(() => false)) {
      await page.keyboard.press("Escape");
      await page.waitForTimeout(300);
    }
  };

  let detailLoaded = false;
  for (let attempt = 1; attempt <= 2; attempt++) {
    const deadline = Date.now() + 10000; // 10s window per attempt
    let popup = null;
    while (Date.now() < deadline) {
      if (await isDetailVisible()) { detailLoaded = true; break; }
      popup = await findErrorPopup();
      if (popup) break;
      await page.waitForTimeout(300);
    }

    if (detailLoaded) break;

    if (popup) {
      console.warn(`  ⚠ "Receipt Number does not exist" popup detected (attempt ${attempt}/2). Closing...`);
      await dismissPopup(popup);

      if (attempt === 2) {
        console.error(`  ❌ Receipt ${voucherNo} does not exist after retry — skipping.`);
        return { ok: false, voucherNo };
      }

      console.log(`  🔄 Retrying: re-paste + click "View Receipt"...`);
      try {
        await voucherInput.click();
        await voucherInput.fill("");
        await voucherInput.fill(voucherNo);
        await frame.getByText('View Receipt', { exact: true }).first().click({ force: true });
      } catch (e) {
        console.error(`  ❌ Retry failed: ${e.message}`);
        return { ok: false, voucherNo };
      }
      continue;
    }

    // Neither popup nor detail within 10s — fall through to the final wait.
    console.log(`  ⏳ No popup or detail within 10s on attempt ${attempt}, falling through...`);
    break;
  }

  // Final readiness signal (no-op if detail already attached above).
  if (!detailLoaded) {
    await frame.locator('text="Receipt Information"').first()
      .waitFor({ state: "attached", timeout: 8000 })
      .catch(() => console.log("  ⏳ Detail screen still loading..."));
  }
  await page.waitForTimeout(300);
  return { ok: true, voucherNo };
}

/**
 * Counts data rows currently rendered in the Select Voucher grid (DOM-based).
 * Fallback when getVoucherNumbers can't extract IDs — lets us still iterate
 * by row index and rely on F2 to surface the voucher number.
 */
async function getGridRowCount(page) {
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  if (!actualFrame) return 0;
  return await actualFrame.evaluate(() => {
    const a = document.querySelectorAll('.x-grid3-row').length;
    const b = document.querySelectorAll('.x-grid-row').length;
    const c = document.querySelectorAll('tr[role="row"]').length;
    return Math.max(a, b, c);
  });
}

/**
 * Legacy direct-cell click — kept as fallback / for debugging. The main loop
 * now prefers openVoucherViaF2.
 */
async function clickVoucherByNumber(page, voucherNo) {
  console.log(`\n🖱 Clicking voucher: ${voucherNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  let clicked = false;

  // Strategy 1 — JS click inside the grid-cell matching the voucher text
  if (actualFrame) {
    try {
      clicked = await actualFrame.evaluate((vno) => {
        const rows = document.querySelectorAll('.x-grid3-row');
        for (const row of rows) {
          const cells = row.querySelectorAll('.x-grid3-cell-inner, .x-grid3-cell, td, a, span, div');
          for (const cell of cells) {
            const txt = (cell.innerText || cell.textContent || "").trim();
            if (txt === vno) {
              const target = cell.querySelector('a') || cell.closest('a') || cell;
              // Dispatch a proper mouse click — plain .click() on a <div> is a no-op
              // for some ExtJS handlers, but works for anchors and cell inners.
              const rect = target.getBoundingClientRect();
              const x = rect.left + rect.width / 2;
              const y = rect.top + rect.height / 2;
              ['mousedown', 'mouseup', 'click'].forEach(type => {
                target.dispatchEvent(new MouseEvent(type, {
                  bubbles: true, cancelable: true, view: window,
                  clientX: x, clientY: y, button: 0,
                }));
              });
              return true;
            }
          }
        }
        return false;
      }, voucherNo);

      if (clicked) console.log(`  ✔ Clicked voucher ${voucherNo} via grid-cell dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS cell-click failed: ${e.message}`);
    }
  }

  // Strategy 2 — Playwright locator fallback (works if it's a real <a>)
  if (!clicked) {
    console.log("  ℹ Falling back to Playwright locator click...");
    const esc = voucherNo.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
    const link = frame
      .locator('a, span, div')
      .filter({ hasText: new RegExp(`^\\s*${esc}\\s*$`) })
      .first();
    try {
      await link.waitFor({ state: "visible", timeout: 5000 });
      await link.scrollIntoViewIfNeeded().catch(() => { });
      await link.click({ force: true, timeout: 5000 });
      clicked = true;
      console.log(`  ✔ Clicked voucher ${voucherNo} via locator.`);
    } catch (e) {
      console.error(`  ❌ Locator click failed: ${e.message}`);
    }
  }

  if (!clicked) return false;

  // Readiness signal for the View Receipt screen
  await frame.locator('text="Receipt Information"').first()
    .waitFor({ state: "attached", timeout: 15000 })
    .catch(() => console.log("  ⏳ Detail screen still loading..."));
  await page.waitForTimeout(1000);
  return true;
}

/**
 * Selects the first receipt row (Stable Phase 1)
 */
async function selectFirstRow(page) {
  console.log("\n✅ Selecting first row...");
  const frame = page.frameLocator("#ilboinnerframe");

  console.log('  ⌨ Focusing "User ID" field to start keyboard sequence...');
  const userIdField = frame.locator('xpath=//label[contains(text(), "User ID")]/following::input[1] | //td[contains(text(), "User ID")]/following::input[1] | //span[contains(text(), "User ID")]/following::input[1]').first();

  await userIdField.waitFor({ state: "attached", timeout: 10000 });
  await userIdField.focus();
  await page.waitForTimeout(1000);

  console.log("  ⌨ Pressing Tab 5 times...");
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Tab");
    await page.waitForTimeout(150);
  }

  console.log("  ⌨ Pressing ArrowLeft and SpaceBar...");
  await page.keyboard.press("ArrowLeft");
  await page.waitForTimeout(200);
  await page.keyboard.press(" ");
  console.log("  ✔ Row 1 selected.");
  await page.waitForTimeout(500);
}

/**
 * Selects the NEXT row using dynamic keyboard sequences
 */
async function selectNextRow(page, processedCount) {
  console.log(`\n✅ Selecting row ${processedCount}...`);

  if (processedCount >= 6) {
    // Row 6+: grid has typically scrolled, so plain ArrowDown + Space is
    // unreliable. Use ArrowDown to move the selector, F2 to engage the row
    // (Ramco's standard "enter cell" key), then SpaceBar to tick the checkbox.
    console.log(`  ⌨ [Row ${processedCount}] Pressing ArrowDown → F2 → SpaceBar...`);
    await page.keyboard.press("ArrowDown");
    await page.waitForTimeout(300);
    await page.keyboard.press("F2");
    await page.waitForTimeout(300);
    await page.keyboard.press(" ");
    await page.waitForTimeout(200);

    // Fallback: if the checkbox did not tick, retry with an ArrowLeft+Space
    // sequence (mirrors the row-1 pattern which is known-stable).
    const frame = page.frameLocator("#ilboinnerframe");
    const selectedRowCount = await frame.locator('.x-grid3-row-selected').count().catch(() => 0);
    if (selectedRowCount === 0) {
      console.log(`  ⚠ Row ${processedCount} not engaged after F2. Retrying ArrowLeft + SpaceBar...`);
      await page.keyboard.press("ArrowLeft");
      await page.waitForTimeout(200);
      await page.keyboard.press(" ");
    }
  } else {
    console.log(`  ⌨ [Row ${processedCount}] Pressing Tab, ArrowRight, and SpaceBar...`);
    await page.keyboard.press("Tab");
    await page.waitForTimeout(300);
    await page.keyboard.press("ArrowRight");
    await page.waitForTimeout(300);
    await page.keyboard.press(" ");
  }

  console.log(`  ✔ Row ${processedCount} selected.`);
  await page.waitForTimeout(500);
}

/**
 * Clicks the "View Receipt" link/button on the Select Receipt screen.
 * Extracts the Receipt No. (e.g. 02CR26-000540) from the loaded detail screen.
 */
async function triggerViewVoucher(page) {
  console.log('  🖱 Clicking "View Receipt" link...');
  const frameLocator = page.frameLocator("#ilboinnerframe");
  const viewBtn = frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /^View Receipt$/i }).first();

  const tryClick = async (locator, name) => {
    try {
      if (await locator.count() > 0) {
        await locator.scrollIntoViewIfNeeded().catch(() => { });
        await locator.click({ force: true, timeout: 5000 });
        return true;
      }
    } catch (e) {
      console.warn(`    ⚠ Click on ${name} failed: ${e.message}`);
    }
    return false;
  };

  let clicked = await tryClick(viewBtn, "Exact Link");
  if (!clicked) {
    const looseBtn = frameLocator.getByText('View Receipt', { exact: false }).first();
    clicked = await tryClick(looseBtn, "Flexible Link");
  }
  if (!clicked) {
    await frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /View Receipt/i }).first().evaluate(el => el.click()).catch(e => console.error("    ⚠ JS Click failed: " + e.message));
  }

  // Wait for the View Receipt detail screen to appear ("Receipt Information" header).
  await frameLocator.locator('text="Receipt Information"').first().waitFor({ state: "attached", timeout: 15000 }).catch(() => {
    console.log("  ⏳ Detail screen still loading...");
  });
  await page.waitForTimeout(1000);

  // Extract Receipt No. from the detail screen
  let voucherNo = "Unknown";
  try {
    const frame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

    if (frame) {
      voucherNo = await frame.evaluate(() => {
        const elements = Array.from(document.querySelectorAll('label, td, span, div, b'));
        const labelEl = elements.find(el => {
          const text = el.innerText.trim();
          return text === 'Receipt No.' || text === 'Receipt No';
        });

        if (labelEl) {
          const td = labelEl.closest('td');
          if (td && td.nextElementSibling) {
            const val = td.nextElementSibling.innerText.trim();
            if (val) return val;
          }

          if (labelEl.tagName === 'LABEL' && labelEl.getAttribute('for')) {
            const input = document.getElementById(labelEl.getAttribute('for'));
            if (input) return input.value || input.innerText;
          }

          const next = labelEl.nextElementSibling;
          if (next && next.innerText.trim()) return next.innerText.trim();
        }

        for (const el of elements) {
          const match = el.innerText.match(/Receipt No\.\s*[:\-]?\s*([A-Z0-9\-_]+)/i);
          if (match && match[1]) return match[1];
        }

        return "Unknown";
      });
    } else {
      console.warn("    ⚠ Could not find frame 'ilboinnerframe' for extraction.");
    }
  } catch (err) {
    console.warn("    ⚠ Failed to extract Receipt No: " + err.message);
  }

  return voucherNo;
}

/**
 * Scrapes the header sections of the View Receipt screen (Receipt
 * Information, Customer Information, Bank Information) and writes the
 * result to `<exportDir>/Header Info - <receiptNo>.txt` as pretty-printed
 * JSON.
 *
 * Must run BEFORE downloadVoucherDetail's Tab + Shift+E sequence, since
 * that keyboard sequence moves focus and can shift the DOM layout.
 */
async function saveVoucherInfoJson(page, voucherNo, exportDir) {
  console.log("  📋 Extracting Receipt / Customer / Bank Information → JSON...");
  try {
    const actualFrame = page.frames().find(
      (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
    );
    if (!actualFrame) {
      console.warn("    ⚠ Could not locate ilboinnerframe — skipping JSON dump.");
      return;
    }

    // Schema mirrors the reference "Header Info" JSON: snake_case section
    // keys, snake_case field keys, values typed (numbers where numeric,
    // null when blank). Each entry pairs the *visible* label on the View
    // Receipt screen with the JSON key we want in the output. Labels that
    // do not exist on a given Receipt Type simply scrape as null and are
    // safe to leave in the schema.
    const schema = {
      receipt_information: [
        ["Receipt No.", "receipt_no"],
        ["Receipt Type", "receipt_type"],
        ["Status", "status"],
        ["Receipt Date", "receipt_date"],
        ["Finance Book", "finance_book"],
        ["Receipt Amount", "receipt_amount"],
        ["Receipt Currency", "receipt_currency"],
      ],
      customer_information: [
        ["Customer Code", "customer_code"],
        ["Customer Name", "customer_name"],
        ["Customer PAN", "customer_pan"],
        ["Address", "address"],
      ],
      bank_information: [
        ["Bank / Cash Code", "bank_cash_code"],
        ["Pay Mode", "pay_mode"],
        ["Instrument No.", "instrument_no"],
        ["Instrument Date", "instrument_date"],
        ["Drawee Bank", "drawee_bank"],
        ["Drawee Branch", "drawee_branch"],
      ],
      reference_details: [
        ["Reference Document No.", "reference_document_no"],
        ["Remarks", "remarks"],
      ],
    };

    // Field keys whose value should be serialized as a Number, not a string.
    const numericKeys = new Set(["receipt_amount"]);

    const allLabels = Object.values(schema).flatMap((pairs) => pairs.map(([lbl]) => lbl));

    const valuesMap = await actualFrame.evaluate((labels) => {
      const isVisible = (el) => {
        if (!el) return false;
        if (el.tagName !== "BODY" && el.offsetParent === null) return false;
        const s = window.getComputedStyle(el);
        if (s.display === "none" || s.visibility === "hidden") return false;
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      };

      const getDirectText = (el) =>
        Array.from(el.childNodes)
          .filter((n) => n.nodeType === Node.TEXT_NODE)
          .map((n) => n.textContent.trim())
          .join("");

      const findLabelEl = (lbl) => {
        // Prefer the most specific (leaf) visible element whose direct text
        // matches exactly — so its bounding rect hugs the label text, not a
        // whole panel.
        let best = null;
        let bestArea = Infinity;
        for (const el of document.querySelectorAll("*")) {
          if (!isVisible(el)) continue;
          if (getDirectText(el) !== lbl) continue;
          const r = el.getBoundingClientRect();
          const area = r.width * r.height;
          if (area < bestArea) { best = el; bestArea = area; }
        }
        return best;
      };

      const readValueFromEl = (el) => {
        if (!el) return "";
        if (el.tagName === "INPUT" || el.tagName === "TEXTAREA" || el.tagName === "SELECT") {
          return (el.value || "").trim();
        }
        return (el.textContent || "").trim();
      };

      const extractFromCell = (cell) => {
        if (!cell) return "";
        const inp = cell.querySelector("input, textarea, select");
        if (inp) return (inp.value || "").trim();
        const disp = cell.querySelector(".x-form-display-field");
        if (disp) return (disp.textContent || "").trim();
        return (cell.textContent || "").trim();
      };

      // Pre-collect all candidate value elements once — used by the
      // spatial-proximity strategy below. We include hidden ExtJS inputs
      // too because their value is often held in the <input> even when the
      // visible pixels come from a wrapper <div>. Plain <td> cells are
      // included because some fields (e.g. Finance Book, Receipt Amount)
      // are rendered as table-cell text rather than ExtJS form inputs.
      const valueCandidates = Array.from(
        document.querySelectorAll(
          'input, textarea, select, .x-form-display-field, .x-form-field, div[class*="x-form"], td'
        )
      ).filter((el) => {
        // Skip the hidden Ramco helpers that aren't actually displayed
        // and skip any candidate that itself *is* a label (matches one of
        // our target label strings exactly).
        if (!isVisible(el) && el.tagName !== "INPUT") return false;
        const t = getDirectText(el);
        if (t && labels.includes(t)) return false;
        return true;
      });

      const map = {};
      for (const lbl of labels) {
        const labelEl = findLabelEl(lbl);
        if (!labelEl) { map[lbl] = ""; continue; }
        const labelRect = labelEl.getBoundingClientRect();

        let value = "";

        // Strategy 1 (primary for Ramco): spatial proximity. The value
        // element is visually to the RIGHT of the label on the same
        // horizontal band. Pick the closest candidate that:
        //   (a) starts to the right of the label (left edge ≥ label's right)
        //   (b) vertically overlaps the label's box
        //   (c) isn't farther than ~500px (next column would overshoot)
        //   (d) isn't another field's label for a different row
        let best = null;
        let bestDist = Infinity;
        for (const cand of valueCandidates) {
          const r = cand.getBoundingClientRect();
          if (r.width === 0 || r.height === 0) continue;
          if (r.left < labelRect.right - 2) continue;
          const vOverlap =
            Math.min(labelRect.bottom, r.bottom) - Math.max(labelRect.top, r.top);
          if (vOverlap < Math.min(labelRect.height, r.height) * 0.4) continue;
          const dist = r.left - labelRect.right;
          if (dist < 0 || dist > 500) continue;
          if (dist < bestDist) { best = cand; bestDist = dist; }
        }
        if (best) {
          value = readValueFromEl(best);
          // If the nearest candidate is an empty wrapper <div>, drill in.
          if (!value) {
            const inner = best.querySelector("input, textarea, select, .x-form-display-field");
            if (inner) value = readValueFromEl(inner);
          }
        }

        // Strategy 2: the label sits in a <td>; the value lives in the next <td>.
        if (!value) {
          const labelCell = labelEl.closest("td");
          if (labelCell && labelCell.nextElementSibling) {
            value = extractFromCell(labelCell.nextElementSibling);
          }
        }

        // Strategy 3: ExtJS form-item container that wraps label + field.
        if (!value) {
          const item = labelEl.closest(".x-form-item, .x-form-element");
          if (item) value = extractFromCell(item);
        }

        // Strategy 4: walk a few following siblings until we find a value.
        if (!value) {
          let sib = labelEl.nextElementSibling;
          for (let i = 0; i < 5 && sib && !value; i++, sib = sib.nextElementSibling) {
            value = extractFromCell(sib);
          }
        }

        // Guard: if the value we picked is literally another label name
        // (happens when the layout wraps oddly), discard it.
        if (value && labels.includes(value)) value = "";

        map[lbl] = value;
      }
      return map;
    }, allLabels);

    // Normalize a raw scraped string into the typed JSON value we want.
    // Empty / missing → null. Numeric-typed keys parse the digits out of
    // amount strings like "25,540.00" → 25540.00. Everything else stays
    // as a trimmed string.
    const normalize = (raw, key) => {
      const s = (raw ?? "").toString().trim();
      if (!s) return null;
      if (numericKeys.has(key)) {
        const cleaned = s.replace(/,/g, "");
        const n = Number(cleaned);
        return Number.isFinite(n) ? n : null;
      }
      return s;
    };

    const structured = {};
    for (const [section, pairs] of Object.entries(schema)) {
      structured[section] = {};
      for (const [label, key] of pairs) {
        structured[section][key] = normalize(valuesMap[label], key);
      }
    }
    // Guarantee the receipt number is populated from the caller-provided
    // value, so this field is never null even if the scrape missed it.
    structured.receipt_information.receipt_no = voucherNo;

    // Extract the bottom totals (Transaction Amount / Base Amount /
    // Parallel Base Amount, each with Debit Total + Credit Total).
    // The labels "Debit Total" / "Credit Total" appear three times on the
    // page, so we resolve them by spatial proximity — find each section
    // header, then grab the closest "Debit/Credit Total" label below it
    // whose left edge aligns with the header's column.
    const totalsRaw = await actualFrame.evaluate(() => {
      const isVisible = (el) => {
        if (!el) return false;
        if (el.tagName !== "BODY" && el.offsetParent === null) return false;
        const s = window.getComputedStyle(el);
        if (s.display === "none" || s.visibility === "hidden") return false;
        const r = el.getBoundingClientRect();
        return r.width > 0 && r.height > 0;
      };
      const getDirectText = (el) =>
        Array.from(el.childNodes)
          .filter((n) => n.nodeType === Node.TEXT_NODE)
          .map((n) => n.textContent.trim())
          .join("");
      const findHeader = (lbl) => {
        let best = null;
        let bestArea = Infinity;
        for (const el of document.querySelectorAll("*")) {
          if (!isVisible(el)) continue;
          if (getDirectText(el) !== lbl) continue;
          const r = el.getBoundingClientRect();
          const area = r.width * r.height;
          if (area < bestArea) { best = el; bestArea = area; }
        }
        return best;
      };
      const findBelowAndAligned = (label, headerRect) => {
        let best = null;
        let bestDist = Infinity;
        for (const el of document.querySelectorAll("*")) {
          if (!isVisible(el)) continue;
          if (getDirectText(el) !== label) continue;
          const r = el.getBoundingClientRect();
          if (r.top < headerRect.bottom - 5) continue;
          if (Math.abs(r.left - headerRect.left) > 300) continue;
          const dist = r.top - headerRect.bottom;
          if (dist < bestDist) { best = el; bestDist = dist; }
        }
        return best;
      };
      const readValueRightOf = (labelEl) => {
        if (!labelEl) return "";
        const labelRect = labelEl.getBoundingClientRect();
        const candidates = document.querySelectorAll(
          'input, textarea, select, .x-form-display-field, .x-form-field, div[class*="x-form"]'
        );
        let best = null;
        let bestDist = Infinity;
        for (const cand of candidates) {
          if (cand === labelEl || labelEl.contains(cand) || cand.contains(labelEl)) continue;
          const r = cand.getBoundingClientRect();
          if (r.width === 0 || r.height === 0) continue;
          if (r.left < labelRect.right - 2) continue;
          const vOverlap = Math.min(labelRect.bottom, r.bottom) - Math.max(labelRect.top, r.top);
          if (vOverlap < Math.min(labelRect.height, r.height) * 0.4) continue;
          const dist = r.left - labelRect.right;
          if (dist > 500) continue;
          if (dist < bestDist) { best = cand; bestDist = dist; }
        }
        if (!best) return "";
        if (best.tagName === "INPUT" || best.tagName === "TEXTAREA" || best.tagName === "SELECT") {
          return (best.value || "").trim();
        }
        return (best.textContent || "").trim();
      };

      const sections = [
        ["Transaction Amount", "transaction_amount"],
        ["Base Amount", "base_amount"],
        ["Parallel Base Amount", "parallel_base_amount"],
      ];
      const out = {};
      for (const [hdrLbl, hdrKey] of sections) {
        out[hdrKey] = { debit_total: "", credit_total: "" };
        const headerEl = findHeader(hdrLbl);
        if (!headerEl) continue;
        const headerRect = headerEl.getBoundingClientRect();
        const debitEl = findBelowAndAligned("Debit Total", headerRect);
        const creditEl = findBelowAndAligned("Credit Total", headerRect);
        out[hdrKey].debit_total = readValueRightOf(debitEl);
        out[hdrKey].credit_total = readValueRightOf(creditEl);
      }
      return out;
    });

    const parseNum = (raw) => {
      const s = (raw ?? "").toString().trim();
      if (!s) return null;
      const n = Number(s.replace(/,/g, ""));
      return Number.isFinite(n) ? n : null;
    };
    structured.totals = {};
    for (const sec of ["transaction_amount", "base_amount", "parallel_base_amount"]) {
      structured.totals[sec] = {
        debit_total: parseNum(totalsRaw[sec] && totalsRaw[sec].debit_total),
        credit_total: parseNum(totalsRaw[sec] && totalsRaw[sec].credit_total),
      };
    }

    const outPath = `${exportDir}/Header Info - ${voucherNo}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save voucher info JSON: ${error.message}`);
  }
}

/**
 * Returns true if the current page (View Receipt or Accounting Information
 * sub-screen) has at least one data row in its main grid. False when the
 * grid shows an empty-state marker like "Found no rows to display!!!" or
 * "[No records to display]" in the paging toolbar.
 *
 * Some receipts have no Debit Document entries; some have no Accounting
 * lines. Without this check, Tab+Shift+E would either export an empty
 * file or hang waiting for a download event that never fires.
 */
async function pageHasGridData(page) {
  const actualFrame = page.frames().find(
    (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
  );
  if (!actualFrame) return false;

  return await actualFrame.evaluate(() => {
    const text = document.body.innerText || "";
    const emptyMarkers = [
      "Found no rows to display",
      "No records to display",
      "No rows to display",
    ];
    for (const marker of emptyMarkers) {
      if (text.includes(marker)) return false;
    }
    return document.querySelectorAll(".x-grid3-row").length > 0;
  });
}

/**
 * Clicks the "Accounting Information" hyperlink on the View Receipt screen.
 * It sits at the bottom-left of the page (between "View Tax Calculation"
 * and "Collaborate") and opens a sub-screen with the debit/credit
 * accounting entries — that's the grid we want to export.
 *
 * Uses the same JS-dispatch trick as the "View Receipt" click because
 * Ramco often renders "hyperlinks" as styled <span>/<div> elements rather
 * than real <a> tags, which Playwright's plain .click() handles unreliably.
 */
async function clickAccountingInformation(page) {
  console.log('  🖱 Clicking "Accounting Information" hyperlink...');
  const frameLocator = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(
    (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
  );

  let clicked = false;
  if (actualFrame) {
    try {
      clicked = await actualFrame.evaluate(() => {
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== "none" && s.visibility !== "hidden";
        };
        // The View Receipt page can carry the text "Accounting Information"
        // in more than one place (action hyperlink + potential section
        // header). Pick the SMALLEST visible element whose direct text
        // exactly matches — that's the leaf hyperlink, not a wrapping
        // section/legend.
        let best = null;
        let bestArea = Infinity;
        for (const el of document.querySelectorAll("*")) {
          if (!isVisible(el)) continue;
          const directText = Array.from(el.childNodes)
            .filter((n) => n.nodeType === Node.TEXT_NODE)
            .map((n) => n.textContent.trim())
            .join("");
          if (directText !== "Accounting Information") continue;
          const r = el.getBoundingClientRect();
          const area = r.width * r.height;
          if (area > 0 && area < bestArea) {
            best = el;
            bestArea = area;
          }
        }
        if (!best) return false;
        const r = best.getBoundingClientRect();
        const x = r.left + r.width / 2;
        const y = r.top + r.height / 2;
        ["mousedown", "mouseup", "click"].forEach((t) => {
          best.dispatchEvent(
            new MouseEvent(t, {
              bubbles: true,
              cancelable: true,
              view: window,
              clientX: x,
              clientY: y,
              button: 0,
            })
          );
        });
        return true;
      });
      if (clicked) console.log('  ✔ Clicked "Accounting Information" via JS dispatch.');
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      // Use .last() so the hyperlink (rendered last in DOM order, at the
      // bottom of the page) wins over any earlier collapsed section header.
      const link = frameLocator.getByText("Accounting Information", { exact: true }).last();
      await link.waitFor({ state: "visible", timeout: 5000 });
      await link.click({ force: true });
      clicked = true;
      console.log('  ✔ Clicked "Accounting Information" via locator fallback.');
    } catch (e) {
      console.error(`  ❌ Failed to click Accounting Information: ${e.message}`);
      return false;
    }
  }

  // The sub-screen takes ~1-2s to attach in Ramco; give it a moment so
  // the keyboard focus settles on the new grid before Shift+E fires.
  await page.waitForTimeout(2500);
  return true;
}

/**
 * Downloads the Accounting Information Excel from the View Receipt screen.
 *
 * Flow: dump header JSON → click "Accounting Information" hyperlink → the
 * accounting sub-screen attaches → Tab + Shift+E to export the grid →
 * Go back to View Receipt (so the main loop's goBackFromDetail returns
 * to Select Receipt as expected).
 */
async function downloadVoucherDetail(page, voucherNo = "Unknown") {
  console.log(`\n📥 Preparing to download Accounting Information Excel (Receipt: ${voucherNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });

    await page.waitForTimeout(500);

    // Resolve the export folder up front — the JSON info dump and the Excel
    // download both land here, so they must agree on the date/receipt path.
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Customer Receipts/${OU_FOLDER}/${dateFolder}/${voucherNo}`;
    mkdirSync(exportDir, { recursive: true });

    // Dump JSON from the View Receipt page BEFORE navigating away — once
    // we click "Accounting Information" the header fields disappear.
    await saveVoucherInfoJson(page, voucherNo, exportDir);

    // Step 1 — export the "Debit Document Information" grid that's on the
    // View Receipt page itself. Per the SOP this needs only ONE Tab to
    // park focus on the grid (vs. five for the Accounting Information
    // sub-screen) before Shift+E fires.
    if (!(await pageHasGridData(page))) {
      console.log("  ⏭ Debit Document Information has no rows — skipping export.");
    } else {
      const MAX_ATTEMPTS = 2;
      let lastErr = null;
      let saved = false;
      for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
        try {
          if (attempt > 1) {
            console.log(`  🔄 Retrying Debit Document export (attempt ${attempt}/${MAX_ATTEMPTS})...`);
            await page.waitForTimeout(800);
          }

          console.log("  ⌨ Pressing Tab once (Debit Document grid)...");
          await page.keyboard.press("Tab");
          await page.waitForTimeout(150);

          console.log("  ⌨ Triggering Excel Export (Shift + E)...");
          const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
          await page.keyboard.down('Shift');
          await page.keyboard.press('E');
          await page.keyboard.up('Shift');

          const download = await downloadPromise;
          const suggestedName = download.suggestedFilename();
          console.log(`  🔍 Server suggested name: "${suggestedName}"`);

          const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
          const fileName = `${exportDir}/Debit Document Info - (${voucherNo}).${ext}`;

          await download.saveAs(fileName);
          console.log(`  ✅ Debit Document Info download successful!`);
          console.log(`  📁 Saved → ${fileName}`);
          saved = true;
          break;
        } catch (e) {
          lastErr = e;
          console.warn(`  ⚠ Debit Document attempt ${attempt}/${MAX_ATTEMPTS} failed: ${e.message}`);
        }
      }
      if (!saved) {
        console.error(`  ❌ Debit Document export failed after ${MAX_ATTEMPTS} attempts: ${lastErr ? lastErr.message : 'unknown'}`);
      }
    }

    // Step 2 — open the Accounting Information sub-screen. Without this,
    // the next Shift+E would just re-export the Debit Document grid we
    // already saved above.
    const opened = await clickAccountingInformation(page);
    if (!opened) {
      console.error("  ❌ Could not open Accounting Information — skipping Excel export.");
      return;
    }

    // Some receipts (e.g. fully-allocated against another voucher) land on
    // an Accounting Information screen with no journal lines. Skip the
    // export but still pop back to View Receipt so the main loop continues.
    if (!(await pageHasGridData(page))) {
      console.log("  ⏭ Accounting Information has no rows — skipping export.");
      console.log("  🔙 Returning from Accounting Information → View Receipt...");
      await goBackFromDetail(page, { waitFor: 'receipt-info' });
      return;
    }

    // Retry the Tab + Shift+E + waitForEvent('download') sequence up to 2
    // times. Transient failures here are usually one of: focus drifted off
    // the grid before Shift+E, server briefly slow to render the file, or
    // a one-off network blip. A second attempt almost always succeeds and
    // saves the row instead of skipping it.
    const MAX_ATTEMPTS = 2;
    let lastErr = null;
    let saved = false;
    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
      try {
        if (attempt > 1) {
          console.log(`  🔄 Retrying Excel export (attempt ${attempt}/${MAX_ATTEMPTS})...`);
          await page.waitForTimeout(800);
        }

        console.log("  ⌨ Pressing Tab once (Accounting Information grid)...");
        await page.keyboard.press("Tab");
        await page.waitForTimeout(150);

        console.log("  ⌨ Triggering Excel Export (Shift + E)...");
        const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
        await page.keyboard.down('Shift');
        await page.keyboard.press('E');
        await page.keyboard.up('Shift');

        const download = await downloadPromise;
        const suggestedName = download.suggestedFilename();
        console.log(`  🔍 Server suggested name: "${suggestedName}"`);

        const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
        const fileName = `${exportDir}/Accounting Info - (${voucherNo}).${ext}`;

        await download.saveAs(fileName);
        console.log(`  ✅ Excel download successful!`);
        console.log(`  📁 Saved → ${fileName}`);
        saved = true;
        break;
      } catch (e) {
        lastErr = e;
        console.warn(`  ⚠ Excel attempt ${attempt}/${MAX_ATTEMPTS} failed: ${e.message}`);
      }
    }
    if (!saved) {
      console.error(`  ❌ Excel download failed after ${MAX_ATTEMPTS} attempts: ${lastErr ? lastErr.message : 'unknown'}`);
    }

    // Always pop back to View Receipt — even on Excel failure — so the
    // main loop's goBackFromDetail correctly lands on Select Receipt.
    console.log("  🔙 Returning from Accounting Information → View Receipt...");
    await goBackFromDetail(page, { waitFor: 'receipt-info' });

  } catch (error) {
    console.error(`  ❌ Excel download failed: ${error.message}`);
  }
}

/**
 * Clicks the "Go back" button. Caller picks the readiness target so we
 * don't waste time waiting for an element the destination screen doesn't
 * have:
 *   - 'grid'         → Select Receipt search results (.x-grid3-row)
 *   - 'receipt-info' → View Receipt header ("Receipt Information")
 *
 * The wait timeout is short on purpose. If the target doesn't appear we
 * still proceed; the next action's own waits will catch any real lag.
 */
async function goBackFromDetail(page, options = {}) {
  const { waitFor = 'grid', timeout = 4000 } = options;
  console.log("\n🔙 Navigating back to previous screen...");

  try {
    const backBtn = page.locator('#back button, button[qtip="Go back"], .x-tool-back').first();

    await backBtn.waitFor({ state: "attached", timeout: 10000 });
    await backBtn.scrollIntoViewIfNeeded().catch(() => { });

    try {
      await backBtn.click({ force: true, timeout: 5000 });
      console.log("  ✔ Clicked 'Go back' (forced).");
    } catch (e) {
      console.warn("  ⚠ Direct 'Go back' click failed. Trying flexible search or JS dispatch...");
      const altBtn = page.locator('button, a, .x-btn-text').filter({ hasText: /Go back/i }).first();
      await altBtn.evaluate(el => el.click()).catch(err => console.error("    ❌ JS Click failed: " + err.message));
    }

    const frame = page.frameLocator("#ilboinnerframe");
    const target = waitFor === 'receipt-info'
      ? frame.locator('text="Receipt Information"').first()
      : frame.locator('.x-grid3-row').first();
    await target.waitFor({ state: "visible", timeout }).catch(() => { });
    await page.waitForTimeout(250);

  } catch (error) {
    console.error(`  ❌ Failed to navigate back: ${error.message}`);
  }
}

async function signOut(page) {
  console.log("\n🔒 Attempting to Sign Out...");

  try {
    console.log('  🖱 Clicking User Menu (#userinfo)...');
    const userMenu = page.locator('#userinfo, button:has-text("Cris V. Ladines")').first();

    if (await userMenu.isVisible({ timeout: 5000 }).catch(() => false)) {
      await userMenu.click();
      await page.waitForTimeout(500);
    } else {
      console.warn("    ⚠ User menu not found, trying direct signout click...");
    }

    console.log('  🖱 Clicking Signout (#signout)...');
    const signoutBtn = page.locator('#signout, [qtip="Signout"], .x-menu-item-text:text("Signout")').first();

    if (await signoutBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await signoutBtn.click();
      console.log("  ✔ Signout button clicked.");
    } else {
      await page.evaluate(() => {
        const btn = document.getElementById('signout') ||
          Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('Signout'));
        if (btn) btn.click();
      }).catch(() => { });
    }

    await page.waitForTimeout(5000);

  } catch (error) {
    console.error(`  ❌ Sign out failed: ${error.message}`);
  }
}

function formatRamcoDate(rawDate) {
  if (!rawDate) return "UnknownDate";
  const cleanDate = rawDate.replace(/[^\d/]/g, "");
  const parts = cleanDate.split("/");
  if (parts.length !== 3) return rawDate.replace(/\//g, "-");

  const d = parseInt(parts[0]);
  const m = parseInt(parts[1]);
  const y = parts[2];

  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  if (m < 1 || m > 12) return rawDate.replace(/\//g, "-");

  return `${months[m - 1]} ${d} ${y}`;
}

/**
 * Downloads the initial Select Receipt search results grid.
 * Uses "Receipt Date From" as the filename anchor.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const dateField = frame.locator('xpath=//label[contains(text(), "Receipt Date From")]/following::input[1] | //td[contains(text(), "Receipt Date From")]/following::input[1]').first();
    const rawDate = await dateField.inputValue();
    console.log(`  📅 Detected Receipt Date From: ${rawDate}`);

    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Customer Receipts/${OU_FOLDER}/${dateFolder}`;
    mkdirSync(exportDir, { recursive: true });

    console.log("  ⌨ Focusing grid for initial download...");
    const searchBtn = frame.locator('button, .x-btn-text').filter({ hasText: /^Search$/i }).last();
    if (await searchBtn.isVisible({ timeout: 5000 })) {
      await searchBtn.focus();
      await page.keyboard.press("Tab");
      await page.waitForTimeout(200);
    }

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");
    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });
    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const fileName = `${exportDir}/${formattedDate}.xls`;

    // Primary save. If the target file is already open in Excel (or otherwise
    // locked by another process), Windows returns EBUSY and saveAs throws.
    // In that case, fall back to a timestamped filename so we can always
    // produce a fresh Excel for the voucher-list read step below.
    let savedPath = fileName;
    try {
      await download.saveAs(fileName);
    } catch (saveErr) {
      const locked = /EBUSY|EPERM|EACCES|locked/i.test(saveErr.message);
      if (!locked) throw saveErr;
      const stamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
      savedPath = `${exportDir}/${formattedDate} (${stamp}).xls`;
      console.warn(`  ⚠ Target file locked (likely open in Excel). Saving as: ${savedPath}`);
      await download.saveAs(savedPath);
    }

    console.log(`  ✅ Initial results saved to: ${savedPath}`);
    return savedPath;

  } catch (error) {
    console.error(`  ❌ Initial results download failed: ${error.message}`);
    return null;
  }
}

/**
 * Read every receipt number from the initial Select Receipt Excel export.
 *
 * Ramco's "Shift+E" export is a ZIP-packaged .xlsm (despite the .xls
 * extension); we unzip with Windows' bundled bsdtar and parse the worksheet
 * XML directly. This used to scan the entire workbook with a regex, but
 * Description cells like "BDO NAIA PESO-CA- AAI(012118002150)" contain
 * parenthesized account numbers that look exactly like Receipt Nos to a
 * naive scanner — the only safe approach is column-scoped extraction.
 *
 * Quirks of Ramco's specific xlsm format that this parser handles:
 *   - All elements are namespaced (`<x:row>`, `<x:c>`, `<x:v>`).
 *   - Data lives in the sheet named "mlreceiptinfo" — file index varies.
 *   - Row 1 is hidden template metadata; row 2 is the visible header.
 *   - Cells use inline strings (`t="str"`), not shared-string indices.
 *   - Some cells omit the `r="C3"` reference — those use sequential
 *     column position (track as we walk the row).
 *
 * Returns an array of unique receipt numbers in sheet order.
 */
function readVouchersFromExcel(xlsPath) {
  console.log(`\n📖 Reading receipt list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read receipt list.`);
    return [];
  }

  const tempDir = path.join(tmpdir(), `ramco-xlsx-${process.pid}-${Date.now()}`);
  mkdirSync(tempDir, { recursive: true });

  const decodeXmlEntities = (s) => s
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");

  // "C" → 2, "AA" → 26, etc. (0-based column index)
  const colLetterToIdx = (letter) => {
    let n = 0;
    for (const c of letter) n = n * 26 + (c.charCodeAt(0) - 64);
    return n - 1;
  };

  // Parse one sheet's XML into a list of rows. Each row is { hidden, cells[] }
  // where cells[i] is the resolved string value at column i (0-based).
  const parseSheet = (xml, sharedStrings) => {
    const rows = [];
    const rowPattern = /<(?:x:)?row\b([^>]*)>([\s\S]*?)<\/(?:x:)?row>/g;
    let rm;
    while ((rm = rowPattern.exec(xml)) !== null) {
      const rowAttrs = rm[1] || '';
      const rowInner = rm[2] || '';
      const hidden = /\bhidden="1"/.test(rowAttrs);

      // Match each <x:c> ... </x:c> (or self-closing). Capture attrs + inner.
      const cellPattern = /<(?:x:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:x:)?c>)/g;
      const cells = [];
      let curCol = 0;
      let cm;
      while ((cm = cellPattern.exec(rowInner)) !== null) {
        const attrs = cm[1] || '';
        const inner = cm[2] || '';
        const rMatch = attrs.match(/\br="([A-Z]+)\d+"/);
        if (rMatch) curCol = colLetterToIdx(rMatch[1]);
        while (cells.length < curCol) cells.push('');

        const tMatch = attrs.match(/\bt="([^"]+)"/);
        const cellType = tMatch ? tMatch[1] : '';

        let value = '';
        if (cellType === 'inlineStr') {
          const im = inner.match(/<(?:x:)?t\b[^>]*>([\s\S]*?)<\/(?:x:)?t>/);
          if (im) value = decodeXmlEntities(im[1]);
        } else {
          const vm = inner.match(/<(?:x:)?v>([\s\S]*?)<\/(?:x:)?v>/);
          if (vm) value = decodeXmlEntities(vm[1]);
          if (cellType === 's') {
            // Shared-string reference — value is an index into sharedStrings.
            const idx = parseInt(value, 10);
            value = Number.isFinite(idx) ? (sharedStrings[idx] || '') : '';
          }
          // For t="str" or numeric / no-type cells the <v> text IS the value.
        }

        cells[curCol] = value;
        curCol++;
      }
      rows.push({ hidden, cells });
    }
    return rows;
  };

  try {
    // Windows bsdtar understands ZIP containers; git-bash shadows it with
    // GNU tar (which does not), so call the Windows binary explicitly.
    const tarExe = process.env.SystemRoot
      ? path.join(process.env.SystemRoot, 'System32', 'tar.exe')
      : 'C:\\Windows\\System32\\tar.exe';

    execFileSync(tarExe, ['-xf', xlsPath, '-C', tempDir], { stdio: 'pipe' });

    // Build the shared-string table (only used by sheets that reference it
    // via `t="s"`; Ramco's mlreceiptinfo sheet uses inline strings instead).
    const sharedStrings = [];
    const sharedPath = path.join(tempDir, 'xl', 'sharedStrings.xml');
    if (existsSync(sharedPath)) {
      const xml = readFileSync(sharedPath, 'utf-8');
      const siPattern = /<(?:x:)?si\b[^>]*>([\s\S]*?)<\/(?:x:)?si>/g;
      let m;
      while ((m = siPattern.exec(xml)) !== null) {
        const tPattern = /<(?:x:)?t\b[^>]*>([\s\S]*?)<\/(?:x:)?t>/g;
        let text = '';
        let tm;
        while ((tm = tPattern.exec(m[1])) !== null) text += tm[1];
        sharedStrings.push(decodeXmlEntities(text));
      }
    }

    // Walk every sheet and stop at the first that has a "Receipt No." header
    // with data rows beneath it. (Ramco numbers sheets non-sequentially —
    // sheet5.xml is the data sheet in the sample; others are templates.)
    const sheetsDir = path.join(tempDir, 'xl', 'worksheets');
    if (!existsSync(sheetsDir)) {
      console.warn('  ⚠ No worksheet directory in Excel.');
      return [];
    }
    const files = readdirSync(sheetsDir)
      .filter(f => /^sheet\d+\.xml$/i.test(f))
      .sort((a, b) => parseInt(a.match(/\d+/)[0], 10) - parseInt(b.match(/\d+/)[0], 10));

    for (const file of files) {
      const xml = readFileSync(path.join(sheetsDir, file), 'utf-8');
      const rows = parseSheet(xml, sharedStrings);

      // Find header row — first non-hidden row containing "Receipt No.".
      let headerColIdx = -1;
      let headerRowIdx = -1;
      for (let i = 0; i < rows.length; i++) {
        if (rows[i].hidden) continue;
        const idx = rows[i].cells.findIndex(c => c.trim() === 'Receipt No.');
        if (idx >= 0) {
          headerColIdx = idx;
          headerRowIdx = i;
          break;
        }
      }
      if (headerColIdx < 0) continue;

      const seen = new Set();
      const out = [];
      for (let i = headerRowIdx + 1; i < rows.length; i++) {
        if (rows[i].hidden) continue;
        const v = (rows[i].cells[headerColIdx] || '').trim();
        if (!v) continue;
        if (/^0+$/.test(v)) continue; // skip all-zero placeholders
        if (!seen.has(v)) {
          seen.add(v);
          out.push(v);
        }
      }

      if (out.length > 0) {
        console.log(`  📍 "Receipt No." in ${file} at column index ${headerColIdx}`);
        console.log(`  ✔ Extracted ${out.length} receipt(s).`);
        return out;
      }
    }

    console.warn('  ⚠ Could not find "Receipt No." column in any sheet.');
    return [];
  } catch (e) {
    console.error(`  ❌ Failed to read Excel: ${e.message}`);
    return [];
  } finally {
    try { rmSync(tempDir, { recursive: true, force: true }); } catch (_) { }
  }
}

// ─── MAIN ──────────────────────────────────────────────────────────────────
async function main() {
  mkdirSync("screenshots", { recursive: true });
  mkdirSync("output", { recursive: true });

  const browser = await chromium.launch({ headless: CONFIG.headless, slowMo: CONFIG.slowMo });
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page = await context.newPage();

  try {
    const startTime = Date.now();
    await login(page);
    await changeContext(page, CONFIG.ouInstance);
    await goToLandingPage(page);

    const allResults = [];
    for (const menuPath of MENU_PATHS) {
      const result = await navigateMenu(page, menuPath);

      // Perform Search logic once we reach the Customer Receipt → View Receipt module
      if (menuPath.join(" ").includes("View Receipt")) {
        if (CONFIG.onlyVoucher) {
          const currentVoucherNo = CONFIG.onlyVoucher;
          console.log(`\n🎯 ONLY_CUSTOMER_RECEIPT="${currentVoucherNo}" — bypassing Excel list and processing this single receipt only.`);
          process.stdout.write(`\n📍 [Single]... ${currentVoucherNo}`);
          const result = await openVoucherViaF2(page, 0, currentVoucherNo, true);
          console.log(`\n📂 --- PROCESSING RECEIPT ${currentVoucherNo} (single) ---`);
          if (!result.ok) {
            console.error(`  ❌ Could not open receipt ${currentVoucherNo} — aborting.`);
          } else {
            await downloadVoucherDetail(page, currentVoucherNo);
            console.log("  🔙 Navigating back to search results...");
            await goBackFromDetail(page);
            console.log(`  ✅ Single receipt completed.`);
          }
        } else {
          const stats = await performVoucherSearch(page);

          // Download the initial Excel export — it contains *every* receipt
          // in the result set, independent of grid pagination / buffered view.
          const initialXlsPath = await downloadInitialSearchResults(page);

          // Read the full receipt list straight out of that Excel file. This
          // replaces the grid-scan + F2/ArrowDown walk: one authoritative list,
          // no keyboard navigation needed. If the Excel read fails we stop —
          // there is no safe fallback that handles paginated grids reliably.
          const voucherNumbers = readVouchersFromExcel(initialXlsPath);
          const totalRows = voucherNumbers.length;

          if (totalRows === 0) {
            console.error("  ❌ No receipts extracted from Excel — cannot process.");
          } else {
            if (stats.total && stats.total !== totalRows) {
              console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
            }

            // Honor RESUME_FROM_CUSTOMER_RECEIPT: locate the receipt in the Excel list
            // and start processing from that row. Preserves the original 1-based
            // "Row N/Total" numbering so logs still line up with the Excel index.
            let startIndex = 1;
            if (CONFIG.resumeFromReceipt) {
              const target = CONFIG.resumeFromReceipt.toUpperCase();
              const matchIdx = voucherNumbers.findIndex(v => String(v).trim().toUpperCase() === target);
              if (matchIdx === -1) {
                console.error(`  ❌ RESUME_FROM_CUSTOMER_RECEIPT="${CONFIG.resumeFromReceipt}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
                break;
              }
              startIndex = matchIdx + 1;
              console.log(`  ⏭ RESUME_FROM_CUSTOMER_RECEIPT="${CONFIG.resumeFromReceipt}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
            }

            console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} receipt(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

            for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
              const currentVoucherNo = voucherNumbers[processedCount - 1];
              process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${currentVoucherNo}`);

              // forceDirectEntry=true: paste the receipt into the Receipt No.
              // Direct Entry field and click "View Receipt". No grid click,
              // no F2, no ArrowDown — the receipt number came from the Excel
              // and is already authoritative.
              const result = await openVoucherViaF2(page, processedCount - 1, currentVoucherNo, true);
              console.log(`\n📂 --- PROCESSING RECEIPT ${currentVoucherNo} (${processedCount} OF ${totalRows}) ---`);

              if (!result.ok) {
                console.error(`  ❌ Could not open receipt ${currentVoucherNo} — skipping.`);
                continue;
              }

              // 1) View Receipt detail export (JSON header + Accounting Information Excel)
              await downloadVoucherDetail(page, currentVoucherNo);

              // 2) Back to Select Receipt for the next row.
              console.log("  🔙 Navigating back to search results...");
              await goBackFromDetail(page);

              console.log(`  ✅ Row ${processedCount} completed.`);
              await page.waitForTimeout(500);
            }
          }
        }

        const innerFrame = page.frame({ name: "ilboinnerframe" });
        if (innerFrame) {
          result.tables = await extractTableData(innerFrame);
          result.textSample = await innerFrame.evaluate(() => document.body.innerText.trim().slice(0, 500));
        }
      }

      allResults.push(result);
    }

    writeFileSync("output/extracted-data-customer-receipt.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-customer-receipt.json");

    const totalSeconds = Math.floor((Date.now() - startTime) / 1000);
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;

    console.log(`\n⏱  Total Process Runtime: ${minutes > 0 ? `${minutes}m ` : ""}${seconds}s`);
    console.log("✅ All done!");

    const shouldSignOut = await askWithCountdown("❓ Sign out? (Y/n)", 10);

    if (shouldSignOut.toLowerCase() === 'y' || shouldSignOut === '') {
      await signOut(page);
    } else {
      console.log("  ⏭ Skipping sign out.");
    }

    console.log("\n✅ Browser staying open.");
    await waitForEnter("👉 Press ENTER to close the browser...");

  } catch (err) {
    console.error("❌ Script error:", err.message);
    await waitForEnter("⚠ Error occurred. Press ENTER to close...");
  } finally {
    await browser.close();
    console.log("🔒 Browser closed.");
  }
}

main();
