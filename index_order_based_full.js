require("dotenv").config();
const { chromium } = require("playwright");
const { mkdirSync, writeFileSync, existsSync, readFileSync, readdirSync, rmSync } = require("fs");
const { execFileSync } = require("child_process");
const { tmpdir } = require("os");
const path = require("path");
const readline = require("readline");

// Default OU shown after login. If the OU env var is empty, null, or equals
// this string, the Change Context step is skipped.
const DEFAULT_OU_INSTANCE = "AAI WORLDWIDE LOGISTICS, INC. HO  OPERATING UNIT";

// Map full OU Instance label → short folder name used under
// output/Supplier Order Based Invoices/. Unknown OUs fall back to a sanitized
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
  // Dedicated OU Instance for the Supplier Order Based Invoice workflow, so it
  // can run against a different OU than the other scripts.
  ouInstance: (process.env.OU_INSTANCE_ORDER_BASED || "").trim(),
  headless: false,
  slowMo: 50,
  timeout: 30000,

  // Select Invoice search filters (dd/mm/yyyy — same format shown in the
  // Ramco UI). The Order Based Invoice screen filters by "Invoice Date From/To"
  // (not Pay Date / Request Date like the other Payable modules).

  invoiceDateFrom: process.env.INVOICE_DATE_FROM_ORDER_BASED || "01/04/2026",
  invoiceDateTo: process.env.INVOICE_DATE_TO_ORDER_BASED || "23/04/2026",

  // Invoice Category combo on the Select Invoice search screen (blank = leave as-is).
  invoiceCategory: (process.env.INVOICE_CATEGORY || "").trim(),

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // invoice number and begins processing from there. Blank = start at row 1.
  resumeFromInvoice: (process.env.RESUME_FROM_ORDER_BASED_INVOICE || "").trim(),

  // How long to wait for the Accounting Information grid to render after
  // clicking the footer link. Some invoices (many journal lines, server
  // load) take well over 20s. Configurable via ACCOUNTING_LOAD_TIMEOUT_MS in .env.
  accountingLoadTimeoutMs: Number(process.env.ACCOUNTING_LOAD_TIMEOUT_MS) || 360000,

  // Independent shorter timeout specifically for waiting for the grid structure (rows) 
  // to render once the page itself is loaded and the mask is gone.
  accountingGridTimeoutMs: Number(process.env.ACCOUNTING_GRID_TIMEOUT_MS) || 15000,

  // Single-invoice mode: if set, the Excel download/list is bypassed entirely
  // and only this exact Invoice No. is processed via Direct Entry. Takes
  // precedence over RESUME_FROM_ORDER_BASED_INVOICE.
  onlyInvoice: (process.env.ONLY_ORDER_BASED_INVOICE || process.env.ONLY_INVOICE || "").trim(),
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Supplier Order Based Invoices/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Supplier Order Based Invoices/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Order Based Invoice workflow:
// Payable Management → Supplier Order Based Invoice → View Invoice
const MENU_PATHS = [
  ["Payable Management", "Supplier Order Based Invoice", "View Invoice"],
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

  await page.waitForSelector('#userinfo', { timeout: CONFIG.timeout });
  await page.waitForTimeout(800);

  console.log('  🖱 Opening User Menu (#userinfo)...');
  await page.locator('#userinfo').first().click();
  await page.waitForTimeout(500);

  console.log('  🖱 Clicking "Change Context"...');
  const changeCtxItem = page.locator(
    '.x-menu-item-text:has-text("Change Context"), [qtip="Change Context"], a:has-text("Change Context")'
  ).first();
  await changeCtxItem.waitFor({ state: 'visible', timeout: CONFIG.timeout });
  await changeCtxItem.click();

  console.log('  ⏳ Waiting for SwitchContexts dialog...');
  await page.waitForTimeout(1500);

  const scopes = [page, ...page.frames()];
  let scope = null;
  let ouInputHandle = null;

  for (let attempt = 0; attempt < 10 && !ouInputHandle; attempt++) {
    for (const s of scopes) {
      try {
        const handle = await s.evaluateHandle(() => {
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

          let container = labelEl.parentElement;
          for (let i = 0; i < 6 && container; i++) {
            const inp = container.querySelector('input[type="text"], input:not([type])');
            if (inp && !labelEl.contains(inp)) return inp;
            container = container.parentElement;
          }
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

  console.log('  🖱 Clicking OK...');
  const okBtn = scope.locator(
    'button:has-text("OK"), .x-btn-text:has-text("OK"), input[type="button"][value="OK"]'
  ).first();
  if (await okBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
    await okBtn.click();
  } else {
    await page.keyboard.press('Enter');
  }

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
 * Orchestrates the Search workflow for the "Select Invoice" screen of the
 * Supplier Order Based Invoice module. Filters by Pay Date From/To and
 * Request Date From/To, then clicks Search.
 *
 * If either filter label isn't on the current screen (some Order Based UIs
 * use "Invoice Date" instead), fillDateField logs a warning and continues
 * — the Search button still runs, and the initial Excel download captures
 * whatever the default filter window resolved to.
 */
async function performInvoiceSearch(page) {
  console.log("\n🔍 Preparing to search Order Based invoices inside frame...");
  const frame = await getFrame(page);

  if (!frame) {
    throw new Error("Could not find ilboinnerframe");
  }

  try {
    // 1. Clear "User ID" field
    console.log('  ⌨ Attempting to clear "User ID" field...');

    const userIdField = frame.locator(
      'xpath=//label[contains(text(), "User ID")]/following::input[1] | ' +
      '//td[contains(text(), "User ID")]/following::input[1] | ' +
      '//span[contains(text(), "User ID")]/following::input[1]'
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
      console.warn('  ⚠️ Could not identify exact "User ID" input via label.');
    }

    // 2. Fill date filters. Same row layout as the other Payable screens:
    //    [<Label>From] [<input From>] [label "To"] [<input To>]
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
        await page.keyboard.press('Tab');
        return true;
      }
      console.warn(`  ⚠️ Date field not found for label "${labelText}" (position ${position}).`);
      return false;
    };

    // ── Helper: select a value in the "Invoice Category" combo ──────────────────
    const selectInvoiceCategory = async (category) => {
      const frameLocator = page.frameLocator("#ilboinnerframe");

      let field = frameLocator
        .locator('xpath=(//*[normalize-space(text())="Invoice Category"])[2]/following::input[1]')
        .first();

      if (!(await field.isVisible({ timeout: 3000 }).catch(() => false))) {
        console.warn('  ⚠️ Invoice Category [2] not visible — trying [1] as fallback...');
        field = frameLocator
          .locator('xpath=(//*[normalize-space(text())="Invoice Category"])[1]/following::input[1]')
          .first();
      }

      if (!(await field.isVisible({ timeout: 3000 }).catch(() => false))) {
        console.warn('  ⚠️ "Invoice Category" field not found — skipping.');
        return false;
      }

      await field.click();
      await page.waitForTimeout(500);

      const findAndClickExact = async () => {
        const items = frameLocator.locator(".x-combo-list-item");
        const count = await items.count().catch(() => 0);
        for (let i = 0; i < count; i++) {
          const text = (await items.nth(i).innerText()).trim();
          if (text === String(category).trim()) {
            await items.nth(i).click();
            return true;
          }
        }
        return false;
      };

      let clicked = await findAndClickExact();

      if (!clicked) {
        const comboList = frameLocator.locator(".x-combo-list-inner").first();
        if (await comboList.isVisible({ timeout: 2000 }).catch(() => false)) {
          await comboList.evaluate((el) => (el.scrollTop = 0));
          await page.waitForTimeout(200);
          for (let attempt = 0; attempt < 50; attempt++) {
            clicked = await findAndClickExact();
            if (clicked) break;
            await comboList.evaluate((el) => (el.scrollTop += 80));
            await page.waitForTimeout(200);
          }
        }
      }

      if (!clicked) {
        console.log(`    ℹ Combo item "${category}" not found in dropdown; typing directly...`);
        await field.focus();
        await page.keyboard.press("Control+A");
        await page.keyboard.press("Backspace");
        await page.keyboard.type(String(category), { delay: 50 });
        await page.keyboard.press("Enter");
      }

      await page.waitForTimeout(500);
      return clicked;
    };

    if (CONFIG.invoiceCategory) {
      console.log(`  📑 Selecting Search Criteria Invoice Category: ${CONFIG.invoiceCategory}`);
      await selectInvoiceCategory(CONFIG.invoiceCategory);
      console.log('  ✅ Invoice Category set.');
    }

    console.log(`  📅 Filling Invoice Date: ${CONFIG.invoiceDateFrom} → ${CONFIG.invoiceDateTo}`);
    await fillDateField("Invoice Date From", 1, CONFIG.invoiceDateFrom);
    await fillDateField("Invoice Date From", 2, CONFIG.invoiceDateTo);
    console.log('  ✅ Invoice Date range set.');

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
 * Invoice-number text pattern (Ramco-wide convention): e.g. "02APOI26-000137".
 * Matches 2 digits + uppercase letters + 2 digits + "-" + digits.
 */
const INVOICE_NO_PATTERN = /^\d{2}[A-Z]+\d{2}-\d+$/;

/**
 * Opens an invoice's detail screen by pasting its Invoice No. into Direct
 * Entry and clicking the "View Invoice" link.
 *
 * Returns { ok: boolean, invoiceNo: string|null }.
 */
async function openInvoiceViaDirectEntry(page, invoiceNo) {
  console.log(`\n🖱 Opening invoice via Direct Entry — ${invoiceNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  const viewLinkText = "View Invoice";
  console.log(`  🎯 Will click footer link: "${viewLinkText}"`);

  // Paste into "Invoice No." Direct Entry field.
  console.log(`  ⌨ Pasting into Direct Entry "Invoice No." field...`);
  const invInput = frame
    .locator('xpath=//*[normalize-space(text())="Invoice No."]/following::input[1]')
    .first();
  try {
    await invInput.waitFor({ state: "visible", timeout: 5000 });
    await invInput.click();
    await invInput.fill("");
    await invInput.fill(invoiceNo);
    console.log(`  ✔ Pasted ${invoiceNo}.`);
  } catch (e) {
    console.error(`  ❌ Failed to paste into Direct Entry: ${e.message}`);
    return { ok: false, invoiceNo };
  }

  // Click the "View Invoice" footer link. Primary: JS dispatch on the
  // innermost element whose *direct* text matches exactly. Fallback: locator.
  console.log(`  🖱 Clicking "${viewLinkText}"...`);
  let clicked = false;

  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  if (actualFrame) {
    try {
      clicked = await actualFrame.evaluate((targetText) => {
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== 'none' && s.visibility !== 'hidden';
        };
        for (const el of document.querySelectorAll('*')) {
          if (!isVisible(el)) continue;
          const directText = Array.from(el.childNodes)
            .filter(n => n.nodeType === Node.TEXT_NODE)
            .map(n => n.textContent.trim())
            .join('');
          if (directText !== targetText) continue;
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
      }, viewLinkText);
      if (clicked) console.log(`  ✔ Clicked "${viewLinkText}" via JS dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      const viewLink = frame.getByText(viewLinkText, { exact: true }).first();
      await viewLink.waitFor({ state: "visible", timeout: 5000 });
      await viewLink.click({ force: true });
      clicked = true;
      console.log(`  ✔ Clicked "${viewLinkText}" via locator.`);
    } catch (e) {
      console.error(`  ❌ Failed to click ${viewLinkText}: ${e.message}`);
      return { ok: false, invoiceNo };
    }
  }

  // Poll for one of two outcomes:
  //   (a) detail page attaches → continue
  //   (b) any Ramco/VirtualWorks error popup (e.g. "Invoice Number does not
  //       exist", "Error not defined") → close, retry once, else skip.
  const ERROR_TEXT_RE = /does not exist|not defined|virtualworks|error/i;

  const findErrorPopup = async () => {
    // (1) Inline #message-panel style errors in page + inner frame.
    for (const root of [page, frame]) {
      const panel = root.locator('#message-panel').first();
      const visible = await panel.isVisible({ timeout: 0 }).catch(() => false);
      if (!visible) continue;
      const text = (await panel.innerText().catch(() => "")).trim();
      if (ERROR_TEXT_RE.test(text)) {
        return { kind: 'panel', panel, text };
      }
    }
    // (2) Modal "VirtualWorks™ Message" dialogs — scan every frame for visible
    //     .x-window / .x-dlg / [role=dialog] containers with error-like text.
    for (const f of page.frames()) {
      let found = null;
      try {
        found = await f.evaluate(() => {
          const isVisible = (el) => {
            if (!el || el.offsetParent === null) return false;
            const s = window.getComputedStyle(el);
            return s.display !== 'none' && s.visibility !== 'hidden';
          };
          const popups = document.querySelectorAll('.x-window, .x-dlg, [role="dialog"]');
          for (const popup of popups) {
            if (!isVisible(popup)) continue;
            const text = (popup.innerText || '').trim();
            const lower = text.toLowerCase();
            if (
              lower.includes('virtualworks') ||
              lower.includes('not defined') ||
              lower.includes('does not exist') ||
              /\berror\b/.test(lower)
            ) {
              return { text };
            }
          }
          return null;
        });
      } catch (_) { }
      if (found) return { kind: 'dialog', panel: null, text: found.text };
    }
    return null;
  };

  const isDetailVisible = async () => {
    for (const t of ['Invoice Information', 'Supplier Information']) {
      const visible = await frame.locator(`text="${t}"`).first()
        .isVisible({ timeout: 0 })
        .catch(() => false);
      if (visible) return true;
    }
    return false;
  };

  const dismissPopup = async (found) => {
    if (found.kind === 'panel' && found.panel) {
      const panel = found.panel;
      const closeBtn = panel.locator('button').filter({ hasText: /^Close$/i }).first();
      await closeBtn.click({ force: true }).catch(() => { });
      await page.waitForTimeout(300);
      if (await panel.isVisible({ timeout: 0 }).catch(() => false)) {
        await panel.locator('.x-tool-close').first().click({ force: true }).catch(() => { });
        await page.waitForTimeout(300);
      }
      if (await panel.isVisible({ timeout: 0 }).catch(() => false)) {
        await page.keyboard.press("Escape");
        await page.waitForTimeout(300);
      }
      return;
    }
    // Modal dialog path — reuse the DOM-based dismisser.
    const closed = await dismissVirtualWorksErrorPopup(page);
    if (!closed) {
      await page.keyboard.press("Escape").catch(() => { });
      await page.waitForTimeout(300);
    }
  };

  // One-line summary of the popup text for logs (strip whitespace runs).
  const summarize = (t) => (t || '').replace(/\s+/g, ' ').trim().slice(0, 120);

  let detailLoaded = false;
  for (let attempt = 1; attempt <= 2; attempt++) {
    const deadline = Date.now() + 10000;
    let popup = null;
    while (Date.now() < deadline) {
      if (await isDetailVisible()) { detailLoaded = true; break; }
      popup = await findErrorPopup();
      if (popup) break;
      await page.waitForTimeout(300);
    }

    if (detailLoaded) break;

    if (popup) {
      console.warn(`  ⚠ Ramco error popup detected (attempt ${attempt}/2): "${summarize(popup.text)}". Closing...`);
      await dismissPopup(popup);

      if (attempt === 2) {
        console.error(`  ❌ Invoice ${invoiceNo} could not be opened after retry — skipping.`);
        return { ok: false, invoiceNo };
      }

      console.log(`  🔄 Retrying: re-paste + click "${viewLinkText}"...`);
      try {
        await invInput.click();
        await invInput.fill("");
        await invInput.fill(invoiceNo);
        await frame.getByText(viewLinkText, { exact: true }).first().click({ force: true });
      } catch (e) {
        console.error(`  ❌ Retry failed: ${e.message}`);
        return { ok: false, invoiceNo };
      }
      continue;
    }

    console.log(`  ⏳ No popup or detail within 10s on attempt ${attempt}, falling through...`);
    break;
  }

  if (!detailLoaded) {
    const attached = await frame.locator('text="Invoice Information"').first()
      .waitFor({ state: "attached", timeout: 15000 })
      .then(() => true)
      .catch(() => false);
    if (!attached) {
      // Sanity-dismiss any lingering error popup so it doesn't bleed into the
      // next iteration, then skip this invoice instead of pressing on.
      await dismissVirtualWorksErrorPopup(page).catch(() => { });
      console.error(`  ❌ Detail screen never attached for ${invoiceNo} — skipping.`);
      return { ok: false, invoiceNo };
    }
  }
  await page.waitForTimeout(1000);
  return { ok: true, invoiceNo };
}

/**
 * Scrapes the View Invoice header (Invoice / Supplier / Payment / Project
 * Information) and writes the result to `<exportDir>/Header Info - <invoiceNo>.txt`
 * as pretty-printed JSON.
 *
 * Must run BEFORE the Shift+E Excel export, since that keyboard sequence
 * shifts focus and can reflow the DOM.
 *
 * Mirrors saveVoucherInfoJson from the Sundry script, but adds support for
 * duplicated labels: on the Order Based Invoice screen "Account Description"
 * and "Supplier Name" each appear twice. The schema's third tuple element is
 * the occurrence index in DOM order (0-based; defaults to 0).
 */
async function saveInvoiceInfoJson(page, invoiceNo, exportDir) {
  console.log("  📋 Extracting Invoice / Supplier / Payment / Project Information → JSON...");
  try {
    const actualFrame = page.frames().find(
      (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
    );
    if (!actualFrame) {
      console.warn("    ⚠ Could not locate ilboinnerframe — skipping JSON dump.");
      return;
    }

    // Schema tuples: [visible label, json key, occurrenceIndex?]
    const schema = {
      invoice_information: [
        ["Invoice No.", "invoice_no"],
        ["Invoice Category", "invoice_category"],
        ["Status", "status"],
        ["Invoice Date", "invoice_date"],
        ["Finance Book", "finance_book"],
      ],
      supplier_information: [
        ["Supplier Registered At", "supplier_registered_at"],
        ["Supplier Code", "supplier_code"],
        ["Supplier Name", "supplier_name", 0],
        // Order Based invoices typically have a single "Supplier Account" block
        // rather than the Prepayment screen's separate Prepayment Account. The
        // label name is a best guess — if the screen shows something different,
        // change the first string to match (the second is our own JSON key).
        ["Supplier Account", "supplier_account"],
        ["Account Description", "supplier_account_description", 0],
        ["Supplier Control Account", "supplier_control_account"],
        ["Account Description", "supplier_control_account_description", 1],
        ["Supplier Invoice No.", "supplier_invoice_no"],
        ["Supplier Invoice Amount", "supplier_invoice_amount"],
        ["Supplier Invoice Date", "supplier_invoice_date"],
        ["Pay To Supplier Code", "pay_to_supplier_code"],
        ["Supplier Name", "pay_to_supplier_name", 1],
        ["Order %", "order_percent"],
      ],
      payment_information: [
        ["Currency", "currency"],
        ["Exchange Rate", "exchange_rate"],
        ["Electronic Payment", "electronic_payment"],
        ["Pay Term", "pay_term"],
        ["Anchor Date", "anchor_date"],
        ["Payment Priority", "payment_priority"],
        ["Payment Method", "payment_method"],
        ["Pay Mode", "pay_mode"],
        ["Payment Processing Point", "payment_processing_point"],
        ["Cost Center", "cost_center"],
        ["Total Invoice Amount", "total_invoice_amount"],
        ["Comments", "comments"],
      ],
      project_information: [
        ["Project OU", "project_ou"],
        ["Project /Sub-Project Code", "project_sub_project_code"],
        ["Project / Sub-Project Description", "project_sub_project_description"],
      ],
    };

    const numericKeys = new Set([
      "exchange_rate",
      "supplier_invoice_amount",
      "total_invoice_amount",
      "order_percent",
    ]);

    // Unique labels we need to locate in the DOM (duplicates deduped — we
    // fetch all occurrences per label once, then pick by index below).
    const allLabels = Array.from(new Set(
      Object.values(schema).flatMap((pairs) => pairs.map(([lbl]) => lbl))
    ));

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

      // Collect ALL visible label elements in DOM order (not just the smallest).
      // For each target label we keep every match so the caller can pick by
      // occurrence index when a label repeats (e.g. "Account Description").
      const findAllLabelEls = (lbl) => {
        const hits = [];
        for (const el of document.querySelectorAll("*")) {
          if (!isVisible(el)) continue;
          if (getDirectText(el) !== lbl) continue;
          hits.push(el);
        }
        // Deduplicate ancestor/descendant pairs: when outer and inner elements
        // both match (inner text bubbles up), keep only the innermost (leaf)
        // so the bounding rect hugs the label, not a panel.
        return hits.filter((el) =>
          !hits.some((other) => other !== el && el.contains(other))
        );
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

      const valueCandidates = Array.from(
        document.querySelectorAll(
          'input, textarea, select, .x-form-display-field, .x-form-field, div[class*="x-form"]'
        )
      ).filter((el) => {
        if (!isVisible(el) && el.tagName !== "INPUT") return false;
        const t = getDirectText(el);
        if (t && labels.includes(t)) return false;
        return true;
      });

      const valueForLabelEl = (labelEl) => {
        const labelRect = labelEl.getBoundingClientRect();
        let value = "";

        // Strategy 1: spatial proximity — closest candidate to the right on
        // the same horizontal band.
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
          if (!value) {
            const inner = best.querySelector("input, textarea, select, .x-form-display-field");
            if (inner) value = readValueFromEl(inner);
          }
        }

        // Strategy 2: adjacent <td>.
        if (!value) {
          const labelCell = labelEl.closest("td");
          if (labelCell && labelCell.nextElementSibling) {
            value = extractFromCell(labelCell.nextElementSibling);
          }
        }

        // Strategy 3: ExtJS form-item container.
        if (!value) {
          const item = labelEl.closest(".x-form-item, .x-form-element");
          if (item) value = extractFromCell(item);
        }

        // Strategy 4: walk a few following siblings.
        if (!value) {
          let sib = labelEl.nextElementSibling;
          for (let i = 0; i < 5 && sib && !value; i++, sib = sib.nextElementSibling) {
            value = extractFromCell(sib);
          }
        }

        if (value && labels.includes(value)) value = "";
        return value;
      };

      // Return map: label → array of values (one per occurrence, in DOM order).
      const map = {};
      for (const lbl of labels) {
        const els = findAllLabelEls(lbl);
        map[lbl] = els.map(valueForLabelEl);
      }
      return map;
    }, allLabels);

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
      for (const entry of pairs) {
        const [label, key, occurrence = 0] = entry;
        const list = valuesMap[label] || [];
        structured[section][key] = normalize(list[occurrence], key);
      }
    }
    // Never let the caller-supplied invoice number get lost to scraping misses.
    structured.invoice_information.invoice_no = invoiceNo;

    const outPath = `${exportDir}/Header Info - ${invoiceNo}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save invoice info JSON: ${error.message}`);
  }
}

/**
 * Downloads the Invoice detail Excel file from the View Invoice screen.
 * Mirrors downloadVoucherDetail from the Sundry script: Tab 9× to seat focus
 * on the grid area, then Shift+E to trigger Ramco's Excel export.
 */
async function downloadInvoiceDetail(page, invoiceNo = "Unknown") {
  console.log(`\n📥 Preparing to download Invoice Detail Excel (Invoice: ${invoiceNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });

    await page.waitForTimeout(1500);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Supplier Order Based Invoices/${OU_FOLDER}/${dateFolder}/${invoiceNo}`;
    mkdirSync(exportDir, { recursive: true });

    // Dump Invoice/Supplier/Payment/Project Information to JSON BEFORE the
    // Shift+E sequence, which shifts keyboard focus and can reflow the DOM.
    await saveInvoiceInfoJson(page, invoiceNo, exportDir);

    // Focus the Item/T/C/D Information grid directly. The Order Based Invoice
    // detail screen has only a few focusable header fields before the grid,
    // so the "Tab 9×" pattern used by the other scripts overshoots. Clicking
    // the first grid row seats focus reliably without counting Tabs.
    console.log("  🖱 Clicking first grid row to focus the Item/T/C/D grid...");
    const frame = page.frameLocator("#ilboinnerframe");
    try {
      await frame.locator('.x-grid3-row').first()
        .click({ force: true, timeout: 5000 });
    } catch (e) {
      console.warn(`  ⚠ Grid-row click failed (${e.message}); falling back to 1 Tab.`);
      await page.keyboard.press("Tab");
    }
    await page.waitForTimeout(300);

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");

    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });

    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const suggestedName = download.suggestedFilename();
    console.log(`  🔍 Server suggested name: "${suggestedName}"`);

    const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';

    const fileName = `${exportDir}/Invoice Info - (${invoiceNo}).${ext}`;

    await download.saveAs(fileName);
    console.log(`  ✅ Excel download successful!`);
    console.log(`  📁 Saved and renamed properly in your project → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Excel download failed: ${error.message}`);
  }
}

/**
 * Navigates to the dedicated Accounting Information page via the footer
 * hyperlink at the bottom of the View Invoice detail screen.
 *
 * The text "Accounting Information" can appear twice (section header +
 * footer link). We target the LAST match in DOM order, which is the footer.
 */
async function goToAccountingInformation(page) {
  console.log("\n📑 Navigating to 'Accounting Information' footer link...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  let clicked = false;

  if (actualFrame) {
    try {
      clicked = await actualFrame.evaluate(() => {
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== 'none' && s.visibility !== 'hidden';
        };
        const candidates = [];
        for (const el of document.querySelectorAll('*')) {
          if (!isVisible(el)) continue;
          const directText = Array.from(el.childNodes)
            .filter(n => n.nodeType === Node.TEXT_NODE)
            .map(n => n.textContent.trim())
            .join('');
          if (directText === 'Accounting Information') candidates.push(el);
        }
        if (candidates.length === 0) return false;
        const target = candidates[candidates.length - 1];
        target.scrollIntoView({ block: 'center' });
        const rect = target.getBoundingClientRect();
        const x = rect.left + rect.width / 2;
        const y = rect.top + rect.height / 2;
        ['mousedown', 'mouseup', 'click'].forEach(t => {
          target.dispatchEvent(new MouseEvent(t, {
            bubbles: true, cancelable: true, view: window,
            clientX: x, clientY: y, button: 0,
          }));
        });
        return true;
      });
      if (clicked) console.log("  ✔ Clicked 'Accounting Information' footer link via JS dispatch.");
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      const link = frame.getByText('Accounting Information', { exact: true }).last();
      await link.waitFor({ state: "visible", timeout: 5000 });
      await link.scrollIntoViewIfNeeded().catch(() => { });
      await link.click({ force: true });
      clicked = true;
      console.log("  ✔ Clicked 'Accounting Information' via locator (last).");
    } catch (e) {
      console.error(`  ❌ Failed to click Accounting Information: ${e.message}`);
      return false;
    }
  }

  const timeoutMs = CONFIG.accountingLoadTimeoutMs;
  const gridTimeoutMs = CONFIG.accountingGridTimeoutMs;

  console.log(`  ⏳ Waiting for Accounting Information page to load (up to ${Math.round(timeoutMs / 1000)}s)...`);
  
  const accDeadline = Date.now() + timeoutMs;
  let accLoaded = false;
  let gridTimeoutReached = false;
  let lastState = "";
  let gridWaitStartTime = null;
  
  while (Date.now() < accDeadline) {
    let state = null;
    for (const f of page.frames()) {
      try {
        state = await f.evaluate(() => {
          const isVisible = (el) => {
            if (!el || el.offsetParent === null) return false;
            const s = window.getComputedStyle(el);
            return s.display !== 'none' && s.visibility !== 'hidden';
          };

          // 1. Check for active loading masks (ExtJS / Ramco loading indicators)
          const masks = document.querySelectorAll('.ext-el-mask, .ext-el-mask-msg, .x-mask, .x-mask-msg');
          for (const mask of masks) {
            if (isVisible(mask)) return 'loading';
          }

          // 2. Check for Accounting Information title or breadcrumb on the frame
          let onAccountingPage = false;
          for (const el of document.querySelectorAll(
            '.x-panel-header-text, .breadcrumb, .x-toolbar, h1, h2, .page-title, td, span, div'
          )) {
            if (!isVisible(el)) continue;
            const txt = (el.innerText || '').trim();
            if (/\bAccounting Information\b/i.test(txt)) {
              onAccountingPage = true;
              break;
            }
          }
          if (!onAccountingPage) return 'navigating';

          // 3. Ensure there are visible grid rows or an empty data message
          const rows = document.querySelectorAll('.x-grid3-row');
          let hasVisibleRows = false;
          for (const r of rows) {
            if (isVisible(r)) {
              hasVisibleRows = true;
              break;
            }
          }
          
          const emptyMessages = document.querySelectorAll('.ext-empty-msg, .x-grid-empty, .x-grid3-empty');
          let hasEmptyMessage = false;
          for (const msg of emptyMessages) {
            if (isVisible(msg)) {
              hasEmptyMessage = true;
              break;
            }
          }

          if (!hasVisibleRows && !hasEmptyMessage) return 'waiting_for_rows';

          return 'ready';
        });
      } catch (_) { }
      if (state) break; // Use the state from the first frame that returns something
    }

    if (state === 'ready') {
      accLoaded = true;
      break;
    }
    
    if (state === 'waiting_for_rows') {
      if (!gridWaitStartTime) {
        gridWaitStartTime = Date.now();
      } else if (Date.now() - gridWaitStartTime > gridTimeoutMs) {
        gridTimeoutReached = true;
        break; // Secondary timeout reached, proceed anyway
      }
    } else {
      // Reset secondary timer if state changes back to navigating or loading
      gridWaitStartTime = null;
    }
    
    // Optional: Log state changes so we know what it's waiting for without spamming
    if (state && state !== lastState && state !== 'navigating') {
      if (state === 'loading') console.log("  … System is processing/loading (mask detected)...");
      if (state === 'waiting_for_rows') console.log(`  … On Accounting page; waiting for grid structure to render (up to ${Math.round(gridTimeoutMs / 1000)}s)...`);
      lastState = state;
    }

    await page.waitForTimeout(500);
  }

  if (accLoaded) {
    console.log("  ✔ Accounting Information grid detected and confirmed ready.");
    await page.waitForTimeout(1000); // Brief pause before interaction
  } else if (gridTimeoutReached) {
    console.warn(`  ⚠ Accounting grid structure not detected within ${Math.round(gridTimeoutMs / 1000)}s; proceeding anyway...`);
    await page.waitForTimeout(1000);
  } else {
    console.warn(`  ⚠ Accounting Information page not fully loaded within ${Math.round(timeoutMs / 1000)}s maximum timeout; proceeding anyway...`);
    await page.waitForTimeout(1000);
  }
  
  return true;
}

/**
 * Downloads the Excel file from the Accounting Information view for an invoice.
 */
async function downloadAccountingExcel(page, invoiceNo = "Unknown") {
  console.log(`\n📊 Preparing to download Accounting Excel (Invoice: ${invoiceNo})...`);

  try {
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Supplier Order Based Invoices/${OU_FOLDER}/${dateFolder}/${invoiceNo}`;
    mkdirSync(exportDir, { recursive: true });

    const frame = page.frameLocator("#ilboinnerframe");

    // Defensive re-check: even though goToAccountingInformation already
    // waits for the grid, confirm the grid is present here too —
    // if it isn't, the export will silently produce a stale/blank file.
    let rowCount = 0;
    try {
      await frame.locator('.x-grid3-row').first()
        .waitFor({ state: "visible", timeout: 8000 });
      rowCount = await frame.locator('.x-grid3-row').count().catch(() => 0);
      console.log(`  ✔ Accounting grid ready (${rowCount} row(s) visible).`);
    } catch (_) {
      console.warn("  ⚠ Accounting grid rows not visible — proceeding anyway after extra wait.");
      await page.waitForTimeout(1500);
    }

    // Click the first grid row to put real keyboard focus on the grid.
    try {
      const firstRow = frame.locator('.x-grid3-row').first();
      if (await firstRow.isVisible({ timeout: 2000 }).catch(() => false)) {
        await firstRow.click({ force: true, timeout: 3000 });
        console.log("  🖱 Clicked first grid row to anchor focus.");
        await page.waitForTimeout(400);
      }
    } catch (_) { /* fall through to Tab-based focus */ }

    console.log("  ⌨ Pressing Tab to focus the grid area...");
    await page.keyboard.press("Tab");
    await page.waitForTimeout(350);

    // Final stability beat.
    await page.waitForTimeout(600);

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");

    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });

    await page.keyboard.down('Shift');
    await page.waitForTimeout(80);
    await page.keyboard.press('E');
    await page.waitForTimeout(80);
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const suggestedName = download.suggestedFilename();
    console.log(`  🔍 Server suggested name: "${suggestedName}"`);

    const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
    const fileName = `${exportDir}/Accounting Info - (${invoiceNo}).${ext}`;

    await download.saveAs(fileName);
    console.log(`  ✅ Accounting Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);

    // Wait for Accounting Information download to settle completely before proceeding
    console.log("  ⏳ Waiting 3s for Accounting Information download process to settle...");
    await page.waitForTimeout(3000);

  } catch (error) {
    console.error(`  ❌ Accounting Excel download failed: ${error.message}`);
  }
}

// ─── TAX CALCULATION (Summary + Details) ───────────────────────────────────

/**
 * Navigates to the "View Tax Calculation - Summary" screen from View Invoice.
 */
async function goToTaxCalculationSummary(page) {
  console.log("\n📑 Navigating to 'View Tax Calculation - Summary'...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    await page.waitForTimeout(3000);

    console.log("  🔍 Searching for 'View Tax Calculation - Summary' link...");
    const link = frame.locator('text="View Tax Calculation - Summary"').first();

    if (await link.isVisible({ timeout: 10000 })) {
      console.log("  🖱 Clicking 'View Tax Calculation - Summary' link...");
      await link.click({ force: true, timeout: 5000 });
      console.log("  ✅ Clicked via standard click.");
    } else {
      console.warn("  ⚠ 'View Tax Calculation - Summary' link was not visible.");
      return false;
    }

    console.log("  ⏳ Waiting for View Tax Calculation - Summary info to load...");
    await page.waitForTimeout(5000);
    return true;
  } catch (error) {
    console.error(`  ❌ Failed to reach View Tax Calculation - Summary: ${error.message}`);
    return false;
  }
}

/**
 * Downloads the Excel file from the Tax Calculation - Summary view.
 */
async function downloadTaxSummaryExcel(page, invoiceNo = "Unknown") {
  console.log(`\n📊 Preparing to download Tax Summary Excel (Invoice: ${invoiceNo})...`);

  try {
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Supplier Order Based Invoices/${OU_FOLDER}/${dateFolder}/${invoiceNo}`;
    mkdirSync(exportDir, { recursive: true });

    await page.waitForTimeout(1000);

    console.log("  ⌨ Pressing Tab to focus the grid area...");
    await page.keyboard.press("Tab");
    await page.waitForTimeout(100);

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");
    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });

    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const suggestedName = download.suggestedFilename();
    const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
    const fileName = `${exportDir}/Tax Calculation Summary - (${invoiceNo}).${ext}`;

    await download.saveAs(fileName);
    console.log(`  ✅ Tax Summary Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);
  } catch (error) {
    console.error(`  ❌ Tax Summary Excel download failed: ${error.message}`);
  }
}

/**
 * Detects and closes the VirtualWorks error popup that appears when an
 * invoice has no Line level tax summary (message: "Summary Information for
 * the Line level taxes, not available. Save tax details in summary page").
 */
async function dismissVirtualWorksErrorPopup(page) {
  for (const f of page.frames()) {
    let closed = false;
    try {
      closed = await f.evaluate(() => {
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== 'none' && s.visibility !== 'hidden';
        };
        const popups = document.querySelectorAll('.x-window, .x-dlg, [role="dialog"]');
        for (const popup of popups) {
          if (!isVisible(popup)) continue;
          const text = (popup.innerText || '').toLowerCase();
          if (!text.includes('virtualworks') && !text.includes('not available')) continue;

          for (const btn of popup.querySelectorAll('button, .x-btn-text, a')) {
            const label = (btn.innerText || btn.textContent || '').trim().toLowerCase();
            if (label === 'close') {
              const rect = btn.getBoundingClientRect();
              const x = rect.left + rect.width / 2;
              const y = rect.top + rect.height / 2;
              ['mousedown', 'mouseup', 'click'].forEach(t => {
                btn.dispatchEvent(new MouseEvent(t, {
                  bubbles: true, cancelable: true, view: window,
                  clientX: x, clientY: y, button: 0,
                }));
              });
              return true;
            }
          }
          const x = popup.querySelector('.x-tool-close, .x-window-close');
          if (x) { x.click(); return true; }
        }
        return false;
      });
    } catch (_) { }
    if (closed) {
      console.log("  ✔ Closed VirtualWorks error popup.");
      await page.waitForTimeout(500);
      return true;
    }
  }
  return false;
}

/**
 * Navigates to 'View Tax Calculation - Details'. If Ramco responds with the
 * "Line level taxes not available" popup, we close it and return false.
 */
async function goToTaxCalculationDetails(page) {
  console.log("\n📑 Navigating to 'View Tax Calculation - Details'...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const link = frame.locator('text="View Tax Calculation - Details"').first();

    if (await link.isVisible({ timeout: 10000 })) {
      console.log("  🖱 Clicking 'View Tax Calculation - Details' link...");
      await link.click({ force: true, timeout: 5000 });
      console.log("  ✅ Clicked via standard click.");

      await page.waitForTimeout(2500);

      const popupClosed = await dismissVirtualWorksErrorPopup(page);
      if (popupClosed) {
        console.log("  ⚠ Tax summary not available for this invoice — skipping tax details.");
        return false;
      }

      console.log("  ⏳ Waiting for View Tax Calculation - Details info to load...");
      await page.waitForTimeout(2500);
      return true;
    } else {
      console.warn("  ⚠ 'View Tax Calculation - Details' link was not visible.");
      return false;
    }
  } catch (error) {
    console.error(`  ❌ Failed to reach View Tax Calculation - Details: ${error.message}`);
    return false;
  }
}

/**
 * Downloads the Tax Details Excel for the given (taxType, lineNo). If taxType
 * or lineNo are blank, the file is named without those qualifiers — useful
 * when we couldn't iterate lines.
 */
async function downloadTaxDetailsExcel(page, invoiceNo, taxType, lineNo) {
  const hasIter = taxType && lineNo !== undefined && lineNo !== null && lineNo !== '';
  const logLabel = hasIter ? `(${taxType} Line ${lineNo}) ` : '';
  console.log(`\n📊 Preparing to download Tax Details Excel ${logLabel}(Invoice: ${invoiceNo})...`);

  try {
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Supplier Order Based Invoices/${OU_FOLDER}/${dateFolder}/${invoiceNo}`;
    mkdirSync(exportDir, { recursive: true });

    await page.waitForTimeout(1000);

    console.log("  ⌨ Pressing Tab to focus the grid area...");
    await page.keyboard.press("Tab");
    await page.waitForTimeout(100);

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");
    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });

    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const fileSuffix = hasIter ? ` (${taxType} Line${lineNo})` : '';
    const fileName = `${exportDir}/Tax Calculation Details${fileSuffix} - (${invoiceNo}).xls`;

    await download.saveAs(fileName);
    console.log(`  ✅ Tax Details Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);
  } catch (error) {
    console.error(`  ❌ Tax Details Excel download failed: ${error.message}`);
  }
}

/**
 * Changes the "Tax Type" dropdown on the Tax Details screen (e.g. "VAT").
 */
async function changeTaxType(page, targetType) {
  console.log(`\n🔄 Changing Tax Type to '${targetType}'...`);
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const taxTypeField = frame.locator(
      'xpath=//label[contains(text(), "Tax Type")]/following::input[1] | ' +
      '//td[contains(text(), "Tax Type")]/following::input[1]'
    ).first();

    await taxTypeField.waitFor({ state: "visible", timeout: 10000 });
    await taxTypeField.click();
    await page.waitForTimeout(500);

    console.log(`  🖱 Selecting '${targetType}' from list...`);
    const option = page.locator(`.x-combo-list-item:has-text("${targetType}")`).first();

    if (await option.isVisible({ timeout: 5000 }).catch(() => false)) {
      await option.click();
    } else {
      console.warn(`    ⚠ Option '${targetType}' not found in list. Trying keyboard input...`);
      await taxTypeField.focus();
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type(targetType, { delay: 50 });
      await page.keyboard.press("Enter");
    }

    console.log(`  ⏳ Waiting for data to refresh...`);
    await page.waitForTimeout(2500);
    return true;
  } catch (error) {
    console.error(`  ❌ Failed to change Tax Type: ${error.message}`);
    return false;
  }
}

/**
 * Enumerates all values in the "Ref. Doc. Line No." dropdown. Tries the
 * ExtJS component's internal store first, falls back to scrolling through
 * the combo list and collecting visible items.
 */
async function getRefDocLineOptions(page) {
  console.log("  🔍 Detecting 'Ref. Doc. Line No.' options...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  try {
    const field = frame.locator(
      'xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | ' +
      '//td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]'
    ).first();

    await field.waitFor({ state: "visible", timeout: 5000 });

    let options = [];
    if (actualFrame) {
      try {
        options = await actualFrame.evaluate(() => {
          const labels = Array.from(document.querySelectorAll('label, td, span'));
          const labelEl = labels.find(el => el.innerText.trim().includes('Ref. Doc. Line No.'));
          if (labelEl) {
            const input = labelEl.closest('tr')?.querySelector('input') ||
              labelEl.parentElement?.querySelector('input') ||
              labelEl.closest('td')?.nextElementSibling?.querySelector('input');
            if (input && input.id) {
              const cmp = window.Ext?.getCmp(input.id);
              if (cmp && cmp.store) {
                const storeData = [];
                cmp.store.each(record => {
                  const val = record.get(cmp.valueField || cmp.displayField || 'text');
                  if (val !== undefined && val !== null && String(val).trim().length > 0) {
                    storeData.push(String(val).trim());
                  }
                });
                if (storeData.length > 0) return storeData;
              }
            }
          }
          return [];
        });
      } catch (e) {
        console.log(`    ℹ ExtJS store read failed: ${e.message}`);
      }
    }

    if (options.length > 0) {
      console.log(`  ✔ Found options: ${options.length}`);
      console.log(`    Options: ${options.join(", ")}`);
      return options;
    }

    console.log("    ℹ Falling back to dropdown scroll enumeration...");
    await field.click();
    await page.waitForTimeout(1000);

    const allOptions = new Set();

    let visibleItems = await page.locator('.x-combo-list-item').evaluateAll(els =>
      els.map(el => el.innerText.trim()).filter(txt => txt.length > 0)
    );
    visibleItems.forEach(item => allOptions.add(item));

    const comboList = page.locator('.x-combo-list-inner').first();
    const isComboVisible = await comboList.isVisible({ timeout: 2000 }).catch(() => false);

    if (isComboVisible) {
      let previousSize = 0;
      let scrollAttempts = 0;
      const maxScrollAttempts = 50;

      while (scrollAttempts < maxScrollAttempts) {
        await comboList.evaluate(el => el.scrollTop += 100);
        await page.waitForTimeout(300);

        visibleItems = await page.locator('.x-combo-list-item').evaluateAll(els =>
          els.map(el => el.innerText.trim()).filter(txt => txt.length > 0)
        );
        visibleItems.forEach(item => allOptions.add(item));

        if (allOptions.size === previousSize) {
          scrollAttempts++;
          if (scrollAttempts > 3) break;
        } else {
          scrollAttempts = 0;
        }
        previousSize = allOptions.size;
      }
    }

    await page.keyboard.press("Escape");
    await page.waitForTimeout(500);

    let uniqueOptions = [...allOptions];

    const allNumeric = uniqueOptions.every(opt => /^\d+$/.test(opt));
    if (allNumeric) {
      uniqueOptions.sort((a, b) => parseInt(a) - parseInt(b));
    }

    if (uniqueOptions.length === 0) uniqueOptions.push("1");

    console.log(`  ✔ Found options: ${uniqueOptions.length}`);
    console.log(`    Options: ${uniqueOptions.join(", ")}`);
    return uniqueOptions;
  } catch (error) {
    console.warn(`  ⚠ Could not detect line options: ${error.message}. Defaulting to ["1"].`);
    return ["1"];
  }
}

/**
 * Selects a specific line in the "Ref. Doc. Line No." dropdown.
 */
async function selectRefDocLine(page, lineNo) {
  console.log(`  🔢 Selecting Ref. Doc. Line No: ${lineNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const field = frame.locator(
      'xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | ' +
      '//td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]'
    ).first();

    await field.waitFor({ state: "visible", timeout: 5000 });
    await field.click();
    await page.waitForTimeout(500);

    let clicked = false;

    const findAndClickExact = async () => {
      const items = page.locator('.x-combo-list-item');
      const count = await items.count();
      for (let i = 0; i < count; i++) {
        const text = await items.nth(i).innerText();
        if (text.trim() === String(lineNo).trim()) {
          await items.nth(i).click();
          return true;
        }
      }
      return false;
    };

    clicked = await findAndClickExact();

    if (!clicked) {
      const comboList = page.locator('.x-combo-list-inner').first();
      const isComboVisible = await comboList.isVisible({ timeout: 2000 }).catch(() => false);

      if (isComboVisible) {
        await comboList.evaluate(el => el.scrollTop = 0);
        await page.waitForTimeout(200);

        for (let attempt = 0; attempt < 50; attempt++) {
          clicked = await findAndClickExact();
          if (clicked) break;

          await comboList.evaluate(el => el.scrollTop += 80);
          await page.waitForTimeout(200);
        }
      }
    }

    if (!clicked) {
      console.log(`    ℹ Scrolling didn't find item, using keyboard input...`);
      await field.focus();
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type(String(lineNo), { delay: 50 });
      await page.keyboard.press("Enter");
    }

    await page.waitForTimeout(2500);
    return true;
  } catch (error) {
    console.error(`    ❌ Failed to select Line ${lineNo}: ${error.message}`);
    return false;
  }
}

/**
 * Clicks the "Go back" button to return to the previous screen
 */
async function goBackFromDetail(page) {
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

    console.log("  ⏳ Waiting for previous screen to load...");
    const frame = page.frameLocator("#ilboinnerframe");
    await frame.locator('.x-grid3-row').first().waitFor({ state: "visible", timeout: 10000 }).catch(() => { });
    await page.waitForTimeout(1000);
  } catch (error) {
    console.error(`  ❌ Failed to navigate back: ${error.message}`);
  }
}

async function signOut(page) {
  console.log("\n🔒 Attempting to Sign Out...");

  try {
    console.log('  🖱 Clicking User Menu (#userinfo)...');
    const userMenu = page.locator('#userinfo').first();

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
 * Downloads the initial "Select Invoice" search-results grid. Uses
 * "Pay Date From" as the filename anchor when present; falls back to
 * today's date when that label isn't on screen.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const dateField = frame.locator(
      'xpath=//label[contains(text(), "Invoice Date From")]/following::input[1] | ' +
      '//td[contains(text(), "Invoice Date From")]/following::input[1]'
    ).first();
    const rawDate = await dateField.inputValue().catch(() => "");
    console.log(`  📅 Detected Invoice Date From: ${rawDate}`);
    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Supplier Order Based Invoices/${OU_FOLDER}/${dateFolder}`;
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
 * Read every invoice number from the initial Select Invoice Excel export.
 * Ramco's Shift+E export is a ZIP-packaged .xlsx — we extract with Windows'
 * bundled bsdtar and regex-match the invoice pattern across the sheet XMLs
 * (and sharedStrings.xml), giving us the full row count in one shot.
 */
function readInvoicesFromExcel(xlsPath) {
  console.log(`\n📖 Reading invoice list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read invoice list.`);
    return [];
  }

  const tempDir = path.join(tmpdir(), `ramco-xlsx-${process.pid}-${Date.now()}`);
  mkdirSync(tempDir, { recursive: true });

  try {
    const tarExe = process.env.SystemRoot
      ? path.join(process.env.SystemRoot, 'System32', 'tar.exe')
      : 'C:\\Windows\\System32\\tar.exe';

    execFileSync(tarExe, ['-xf', xlsPath, '-C', tempDir], { stdio: 'pipe' });

    const seen = new Set();
    const out = [];
    const addMatches = (text) => {
      // Order Based invoices only: prefix must be APOI (Accounts Payable Order Based
      // Invoice). Broader patterns leak PO numbers (e.g. "02PO26-000116") into
      // the list, causing Direct Entry lookups to fail with "Error not defined".
      const pattern = /\b(\d{2}APOI\d{2}-\d+)\b/gi;
      let m;
      while ((m = pattern.exec(text)) !== null) {
        const inv = m[1].toUpperCase();
        if (!seen.has(inv)) {
          seen.add(inv);
          out.push(inv);
        }
      }
    };

    const sheetsDir = path.join(tempDir, 'xl', 'worksheets');
    if (existsSync(sheetsDir)) {
      const files = readdirSync(sheetsDir)
        .filter(f => /^sheet\d+\.xml$/i.test(f))
        .sort((a, b) => {
          const na = parseInt(a.match(/\d+/)[0], 10);
          const nb = parseInt(b.match(/\d+/)[0], 10);
          return na - nb;
        });
      for (const f of files) {
        addMatches(readFileSync(path.join(sheetsDir, f), 'utf-8'));
      }
    }

    const sharedPath = path.join(tempDir, 'xl', 'sharedStrings.xml');
    if (existsSync(sharedPath)) {
      addMatches(readFileSync(sharedPath, 'utf-8'));
    }

    console.log(`  ✔ Extracted ${out.length} invoice(s) from Excel.`);
    return out;
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

      if (menuPath.join(" ").includes("View Invoice")) {
        const stats = await performInvoiceSearch(page);

        if (CONFIG.onlyInvoice) {
          // Single-invoice mode: skip Excel, process just this one via Direct Entry.
          const invoiceNo = CONFIG.onlyInvoice;
          console.log(`\n🎯 ONLY_INVOICE="${invoiceNo}" — bypassing Excel list and processing this single invoice only.`);
          process.stdout.write(`\n📍 [Single]... ${invoiceNo}`);

          const openResult = await openInvoiceViaDirectEntry(page, invoiceNo);
          console.log(`\n📂 --- PROCESSING INVOICE ${invoiceNo} (single) ---`);

          if (!openResult.ok) {
            console.error(`  ❌ Could not open invoice ${invoiceNo} — aborting.`);
          } else {
            // 1) View Invoice detail export
            await downloadInvoiceDetail(page, invoiceNo);

            // 2) Accounting Information drill-down
            const accountingOk = await goToAccountingInformation(page);
            if (accountingOk) {
              await downloadAccountingExcel(page, invoiceNo);
              console.log("  🔙 Returning from Accounting to View Invoice...");
              await goBackFromDetail(page);
            }

            // 3) Tax Calculation - Summary → download + drill into Details
            const summaryOpened = await goToTaxCalculationSummary(page);
            if (summaryOpened) {
              await downloadTaxSummaryExcel(page, invoiceNo);

              // --- NESTED: EWT Tax Type lines ---
              const ewtDetailsOpened = await goToTaxCalculationDetails(page);
              if (ewtDetailsOpened) {
                console.log("  📂 Processing EWT Tax Type lines...");
                const ewtLines = await getRefDocLineOptions(page);
                for (const line of ewtLines) {
                  const selected = await selectRefDocLine(page, line);
                  if (selected) {
                    await downloadTaxDetailsExcel(page, invoiceNo, "EWT", line);
                  }
                }

                // Back to Tax Summary, then re-enter Details fresh for VAT.
                console.log("  🔙 Returning from Tax Details (EWT done) to Tax Summary...");
                await goBackFromDetail(page);

                console.log("  📑 Re-entering Tax Details for VAT processing...");
                const vatDetailsOpened = await goToTaxCalculationDetails(page);
                if (vatDetailsOpened) {
                  const typeChanged = await changeTaxType(page, "VAT");
                  if (typeChanged) {
                    console.log("  📂 Processing VAT Tax Type lines...");
                    const vatLines = await getRefDocLineOptions(page);
                    for (const line of vatLines) {
                      const selected = await selectRefDocLine(page, line);
                      if (selected) {
                        await downloadTaxDetailsExcel(page, invoiceNo, "VAT", line);
                      }
                    }
                  }
                  console.log("  🔙 Returning from Tax Details (VAT done) to Tax Summary...");
                  await goBackFromDetail(page);
                }
              }

              // Back from Tax Summary to View Invoice.
              console.log("  🔙 Returning from Tax Summary to View Invoice...");
              await goBackFromDetail(page);
            } else {
              console.log("  ⏭ Tax Summary not available — skipping tax exports.");
            }

            // 4) Back from View Invoice to Select Invoice search grid.
            console.log("  🔙 Navigating back to search results...");
            await goBackFromDetail(page);

            console.log(`  ✅ Single invoice completed.`);
          }
        } else {
          // Download the initial Excel — contains every invoice in the result set.
          const initialXlsPath = await downloadInitialSearchResults(page);

          // Read the invoice list from that Excel.
          const invoiceNumbers = readInvoicesFromExcel(initialXlsPath);
          const totalRows = invoiceNumbers.length;

          if (totalRows === 0) {
            console.error("  ❌ No invoices extracted from Excel — cannot process.");
          } else {
            if (stats.total && stats.total !== totalRows) {
              console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
            }

            // Honor RESUME_FROM_ORDER_BASED_INVOICE
            let startIndex = 1;
            if (CONFIG.resumeFromInvoice) {
              const target = CONFIG.resumeFromInvoice.toUpperCase();
              const matchIdx = invoiceNumbers.findIndex(v => String(v).trim().toUpperCase() === target);
              if (matchIdx === -1) {
                console.error(`  ❌ RESUME_FROM_ORDER_BASED_INVOICE="${CONFIG.resumeFromInvoice}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
                break;
              }
              startIndex = matchIdx + 1;
              console.log(`  ⏭ RESUME_FROM_ORDER_BASED_INVOICE="${CONFIG.resumeFromInvoice}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
            }

            console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} invoice(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

            for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
              const invoiceNo = invoiceNumbers[processedCount - 1];
              process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${invoiceNo}`);

              const openResult = await openInvoiceViaDirectEntry(page, invoiceNo);
              console.log(`\n📂 --- PROCESSING INVOICE ${invoiceNo} (${processedCount} OF ${totalRows}) ---`);

              if (!openResult.ok) {
                console.error(`  ❌ Could not open invoice ${invoiceNo} — skipping.`);
                continue;
              }

              // 1) View Invoice detail export
              await downloadInvoiceDetail(page, invoiceNo);

              // 2) Accounting Information drill-down
              const accountingOk = await goToAccountingInformation(page);
              if (accountingOk) {
                await downloadAccountingExcel(page, invoiceNo);
                console.log("  🔙 Returning from Accounting to View Invoice...");
                await goBackFromDetail(page);
              }

              // 3) Tax Calculation - Summary → download + drill into Details
              const summaryOpened = await goToTaxCalculationSummary(page);
              if (summaryOpened) {
                await downloadTaxSummaryExcel(page, invoiceNo);

                // --- NESTED: EWT Tax Type lines ---
                const ewtDetailsOpened = await goToTaxCalculationDetails(page);
                if (ewtDetailsOpened) {
                  console.log("  📂 Processing EWT Tax Type lines...");
                  const ewtLines = await getRefDocLineOptions(page);
                  for (const line of ewtLines) {
                    const selected = await selectRefDocLine(page, line);
                    if (selected) {
                      await downloadTaxDetailsExcel(page, invoiceNo, "EWT", line);
                    }
                  }

                  // Back to Tax Summary, then re-enter Details fresh for VAT.
                  console.log("  🔙 Returning from Tax Details (EWT done) to Tax Summary...");
                  await goBackFromDetail(page);

                  console.log("  📑 Re-entering Tax Details for VAT processing...");
                  const vatDetailsOpened = await goToTaxCalculationDetails(page);
                  if (vatDetailsOpened) {
                    const typeChanged = await changeTaxType(page, "VAT");
                    if (typeChanged) {
                      console.log("  📂 Processing VAT Tax Type lines...");
                      const vatLines = await getRefDocLineOptions(page);
                      for (const line of vatLines) {
                        const selected = await selectRefDocLine(page, line);
                        if (selected) {
                          await downloadTaxDetailsExcel(page, invoiceNo, "VAT", line);
                        }
                      }
                    }
                    console.log("  🔙 Returning from Tax Details (VAT done) to Tax Summary...");
                    await goBackFromDetail(page);
                  }
                }

                // Back from Tax Summary to View Invoice.
                console.log("  🔙 Returning from Tax Summary to View Invoice...");
                await goBackFromDetail(page);
              } else {
                console.log("  ⏭ Tax Summary not available — skipping tax exports.");
              }

              // 4) Back from View Invoice to Select Invoice search grid.
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

    writeFileSync("output/extracted-data-order-based.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-order-based.json");

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
