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
// output/Adjustment Vouchers/. Unknown OUs fall back to a sanitized
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
  url:      process.env.RAMCO_URL      || "https://aai.ramcoes.com/RVW/extui/vwrt/LaunchPanel.htm",
  username: process.env.RAMCO_USERNAME || "",
  password: process.env.RAMCO_PASSWORD || "",
  // Dedicated OU Instance for the Adjustment Voucher workflow, so it
  // can run against a different OU than the other scripts.
  ouInstance: (process.env.OU_INSTANCE_ADJUSTMENT_VOUCHER || "").trim(),
  headless: false,
  slowMo:   50,
  timeout:  30000,

  // Select Adjustment Voucher search filters (dd/mm/yyyy — same format shown in
  // the Ramco UI). Mirrors the sibling Customer Balance modules, which filter by
  // "Document Date From/To". ⚠ VERIFY against the real Select screen (could be
  // "Voucher Date From/To"); the fill tries both labels.
  documentDateFrom: process.env.DOCUMENT_DATE_FROM_ADJUSTMENT_VOUCHER || "01/02/2026",
  documentDateTo:   process.env.DOCUMENT_DATE_TO_ADJUSTMENT_VOUCHER   || "28/02/2026",

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // voucher number and begins processing from there. Blank = start at row 1.
  resumeFromVoucher: (process.env.RESUME_FROM_ADJUSTMENT_VOUCHER || "").trim(),

  // Single-voucher mode: if set, the Excel download/list is bypassed entirely
  // and only this exact Voucher No. is processed via Direct Entry. Takes
  // precedence over RESUME_FROM_ADJUSTMENT_VOUCHER.
  onlyVoucher: (process.env.ONLY_ADJUSTMENT_VOUCHER || "").trim(),

  // Direct Entry "Voucher Type" dropdown value used when opening each voucher
  // (e.g. "Single Currency" or "Multi Currency"). Must match an option in that
  // dropdown — the selector matches leniently (case-insensitive contains).
  voucherType: (process.env.VOUCHER_TYPE_ADJUSTMENT_VOUCHER || "Single Currency").trim(),

  // How long to wait for a drill-down grid to render after clicking a link.
  // Configurable via env for one-off slow runs.
  accountingLoadTimeoutMs: Number(process.env.ACCOUNTING_LOAD_TIMEOUT_MS) || 60000,
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Adjustment Vouchers/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Adjustment Vouchers/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Adjustment Voucher workflow:
// Receivable Management → Customer Balance Adjustment → View Adjustment
// (the "View Adjustment" menu item opens the "View Adjustment Voucher" screen).
const MENU_PATHS = [
  ["Receivable Management", "Customer Balance Adjustment", "View Adjustment"],
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
  const pwSel   = "#ide_password";

  await loginFrame.waitForSelector(userSel, { state: "attached", timeout: CONFIG.timeout });
  await loginFrame.waitForSelector(pwSel,   { state: "attached", timeout: CONFIG.timeout });

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
      if (pw) ["keydown","keypress","keyup"].forEach(t =>
        pw.dispatchEvent(new KeyboardEvent(t, { key:"Enter", keyCode:13, bubbles:true }))
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
  await ouInputHandle.scrollIntoViewIfNeeded().catch(() => {});
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
  await page.waitForSelector('#userinfo', { timeout: CONFIG.timeout }).catch(() => {});

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
        await page.waitForSelector('#tb-bpc-menu', { timeout: 8000 }).catch(() => {});
        break;
      }
    } catch (_) {}
  }

  await page.waitForSelector('#tb-bpc-menu', { timeout: 10000 }).catch(() => {});
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
          await el.scrollIntoViewIfNeeded().catch(() => {});

          if (i < menuPath.length - 1) {
            console.log(`    🖱 Hovering "${label}"...`);
            await el.hover({ force: true }).catch(() => {});
            await page.waitForTimeout(400);
          }

          console.log(`    🖱 Clicking animated "${label}"...`);
          await el.click();
          await page.waitForTimeout(1000);
          console.log(`    ✔ Clicked "${label}" via: ${sel}`);
          clicked = true;
          break;
        }
      } catch (_) {}
    }

    if (!clicked) {
      console.warn(`    ⚠ Exact match for "${label}" failed. Trying generic visibility...`);
      try {
        const looseEl = page.getByText(label).filter({ visible: true }).first();
        if (await looseEl.isVisible({ timeout: 2000 }).catch(() => false)) {
          if (i < menuPath.length - 1) {
            await looseEl.hover().catch(() => {});
            await page.waitForTimeout(1000);
          }
          await looseEl.click();
          await page.waitForTimeout(2500);
          console.log(`    ✔ Clicked "${label}" via visibility filter`);
          clicked = true;
        }
      } catch (_) {}
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
  await page.waitForLoadState("networkidle").catch(() => {});
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
 * Orchestrates the Search workflow for the "Select Adjustment Voucher" screen of
 * the Customer Balance Adjustment module. Filters by Document Date From/To, then
 * clicks Search.
 */
async function performVoucherSearch(page) {
  console.log("\n🔍 Preparing to search adjustment vouchers inside frame...");
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

    // 2. Fill date filters. Same row layout as the other Receivable screens:
    //    [<Label>From] [<input From>] [label "To"] [<input To>]
    // The label is "Document Date From" on the sibling modules; some screens
    // use "Voucher Date From" — try both and use whichever is present.
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
      return false;
    };

    console.log('⌨ Attempting to select "All" in Currency dropdown...');

    const currencyField = frame.locator(
      'xpath=//label[contains(text(), "Currency")]/following::input[1] | ' +
      '//td[contains(text(), "Currency")]/following::input[1] | ' +
      '//span[contains(text(), "Currency")]/following::input[1]'
    ).first();

    if (await currencyField.isVisible({ timeout: 5000 }).catch(() => false)) {
        console.log('✅ Currency field found.');

        // Open dropdown
        await currencyField.click({ force: true });

        // Clear current value
        await currencyField.press('Control+A');
        await currencyField.press('Delete');

        // Type desired value
        await currencyField.fill('All');

        // Click actual dropdown option instead of ArrowDown
        const allOption = frame.locator(
          'xpath=//*[normalize-space()="All" and not(self::label)]'
        ).first();

        if (await allOption.isVisible({ timeout: 2000 }).catch(() => false)) {
            await allOption.click({ force: true });
            console.log('✅ Currency All selected.');
        } else {
            await page.keyboard.press('Enter');
            console.log('⚠️ Dropdown option not found, Enter pressed.');
        }

        // return true;
    } else {
        console.warn('⚠️ Could not identify exact Currency input.');
    }
    // return false;
 
    // Try each candidate label until the "From" field is found, then fill both.
    const DATE_LABELS = ["Document Date From", "Voucher Date From"];
    console.log(`  📅 Filling date range: ${CONFIG.documentDateFrom} → ${CONFIG.documentDateTo}`);
    let dateLabel = null;
    for (const lbl of DATE_LABELS) {
      if (await fillDateField(lbl, 1, CONFIG.documentDateFrom)) { dateLabel = lbl; break; }
    }
    if (dateLabel) {
      await fillDateField(dateLabel, 2, CONFIG.documentDateTo);
      console.log(`  ✅ Date range set via "${dateLabel}".`);
    } else {
      console.warn(`  ⚠️ Date filter not found (tried: ${DATE_LABELS.join(", ")}).`);
    }
 
    // 3. Click the Search button
    console.log('  🖱 Clicking "Search" button...');
    const searchBtn = frame.locator('button, .x-btn-text').filter({ hasText: /^Search$/i }).last();

    if (await searchBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await searchBtn.click({ force: true });
      console.log("  🚀 Search triggered.");

      let rowsVisible = await frame.locator('.x-grid3-row').first()
        .waitFor({ state: "visible", timeout: 15000 })
        .then(() => true)
        .catch(() => false);

      // Fallback: if the locator click didn't register (the .x-btn anchor can
      // swallow the synthetic click), re-fire via a real DOM MouseEvent on the
      // innermost element whose direct text is exactly "Search".
      if (!rowsVisible) {
        console.warn('  ⚠ No rows after click — retrying Search via JS dispatch...');
        const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
        if (actualFrame) {
          await actualFrame.evaluate(() => {
            const isVisible = (el) => {
              if (!el || el.offsetParent === null) return false;
              const s = window.getComputedStyle(el);
              return s.display !== 'none' && s.visibility !== 'hidden';
            };
            for (const el of document.querySelectorAll('button, .x-btn-text, a')) {
              if (!isVisible(el)) continue;
              if ((el.innerText || el.textContent || '').trim() !== 'Search') continue;
              const r = el.getBoundingClientRect();
              const x = r.left + r.width / 2, y = r.top + r.height / 2;
              ['mousedown', 'mouseup', 'click'].forEach(t =>
                el.dispatchEvent(new MouseEvent(t, { bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0 }))
              );
              return true;
            }
            return false;
          }).catch(() => {});
          rowsVisible = await frame.locator('.x-grid3-row').first()
            .waitFor({ state: "visible", timeout: 15000 })
            .then(() => true)
            .catch(() => false);
        }
      }

      if (!rowsVisible) console.log("  ⏳ Still no grid rows — the date range may simply have no records.");

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
 * Opens a adjustment voucher's detail screen by pasting its Adjustment Voucher No. into
 * Direct Entry and clicking the "View Adjustment Voucher" link.
 *
 * Returns { ok: boolean, voucherNo: string|null }.
 */
/**
 * Sets the Direct Entry "Voucher Type" dropdown (top of the Select Adjustment
 * Voucher screen, next to "View Adjustment Voucher") to `desiredType` — e.g.
 * "Single Currency" or "Multi Currency". This dropdown decides which kind of
 * voucher the "View Adjustment Voucher" link opens.
 *
 * It's the FIRST "Voucher Type" label on the page (the Search Criteria block has
 * another one lower down). We open the ExtJS combo via its trigger arrow (or
 * ArrowDown fallback) and click the matching list item. Returns true on success.
 */
async function selectDirectEntryVoucherType(page, desiredType) {
  console.log(`\n🔽 Setting Direct Entry "Voucher Type" → "${desiredType}"...`);
  const frame = page.frameLocator("#ilboinnerframe");

  // Combo input immediately after the FIRST "Voucher Type" label (Direct Entry).
  const comboInput = frame.locator(
    'xpath=(//*[normalize-space(text())="Voucher Type"])[1]/following::input[1]'
  ).first();

  if (!(await comboInput.isVisible({ timeout: 5000 }).catch(() => false))) {
    console.warn('  ⚠ Direct Entry "Voucher Type" dropdown not found — leaving as-is.');
    return false;
  }

  const current = (await comboInput.inputValue().catch(() => "")).trim();
  if (current && current.toLowerCase() === desiredType.toLowerCase()) {
    console.log(`  ✔ Already set to "${current}".`);
    return true;
  }
  console.log(`  ↻ Current value "${current || "(blank)"}" → changing to "${desiredType}".`);

  // Open the combo: click its trigger arrow; fall back to focus + ArrowDown.
  const trigger = frame.locator(
    'xpath=(//*[normalize-space(text())="Voucher Type"])[1]/following::input[1]' +
    '/following::img[contains(@class,"x-form-trigger")][1]'
  ).first();
  if (await trigger.isVisible({ timeout: 2000 }).catch(() => false)) {
    await trigger.click({ force: true });
  } else {
    await comboInput.click();
    await page.keyboard.press('ArrowDown');
  }
  await page.waitForTimeout(600);

  // Click the matching item in the dropdown list (case-insensitive contains).
  const item = frame.locator('.x-combo-list-item, li.x-combo-list-item, .x-combo-list .x-combo-list-item')
    .filter({ hasText: new RegExp(desiredType.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i') })
    .first();
  if (await item.isVisible({ timeout: 3000 }).catch(() => false)) {
    await item.click({ force: true });
    await page.waitForTimeout(400);
    const after = (await comboInput.inputValue().catch(() => "")).trim();
    console.log(`  ✔ Voucher Type now "${after}".`);
    return true;
  }

  // Fallback: type the value and commit with Enter.
  console.warn('  ⚠ Dropdown item not visible — typing the value instead...');
  await comboInput.click();
  await comboInput.fill("");
  await comboInput.type(desiredType, { delay: 25 });
  await page.waitForTimeout(400);
  await page.keyboard.press('Enter');
  await page.waitForTimeout(300);
  return true;
}

async function openVoucherViaDirectEntry(page, voucherNo) {
  console.log(`\n🖱 Opening adjustment voucher via Direct Entry — ${voucherNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  const viewLinkText = "View Adjustment Voucher";
  console.log(`  🎯 Will click footer link: "${viewLinkText}"`);

  // Set the Direct Entry "Voucher Type" dropdown first — it governs which kind
  // of voucher the "View Adjustment Voucher" link will open.
  await selectDirectEntryVoucherType(page, CONFIG.voucherType);

  // Paste into the Direct Entry "Voucher No." field (top-left of the screen).
  // Falls back to "Document No." in case this Select screen uses that label.
  // Exact text match avoids colliding with the "... From" search-criteria fields.
  console.log(`  ⌨ Pasting into Direct Entry "Voucher No." field...`);
  const invInput = frame
    .locator('xpath=//*[normalize-space(text())="Voucher No."]/following::input[1] | ' +
             '//*[normalize-space(text())="Document No."]/following::input[1]')
    .first();
  try {
    await invInput.waitFor({ state: "visible", timeout: 5000 });
    await invInput.click();
    await invInput.fill("");
    await invInput.fill(voucherNo);
    console.log(`  ✔ Pasted ${voucherNo}.`);
  } catch (e) {
    console.error(`  ❌ Failed to paste into Direct Entry: ${e.message}`);
    return { ok: false, voucherNo };
  }

  // Click the "View Adjustment Voucher" footer link. Primary: JS dispatch on the
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
      return { ok: false, voucherNo };
    }
  }

  // Poll for one of two outcomes:
  //   (a) detail page attaches → continue
  //   (b) any Ramco/VirtualWorks error popup (e.g. "Adjustment Voucher Number does
  //       not exist", "Error not defined") → close, retry once, else skip.
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
      } catch (_) {}
      if (found) return { kind: 'dialog', panel: null, text: found.text };
    }
    return null;
  };

  const isDetailVisible = async () => {
    // Section headers present on the View Adjustment Voucher detail screen.
    for (const t of ['Credit Documents', 'Debit Documents', 'Customer Information', 'voucher information']) {
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
      await closeBtn.click({ force: true }).catch(() => {});
      await page.waitForTimeout(300);
      if (await panel.isVisible({ timeout: 0 }).catch(() => false)) {
        await panel.locator('.x-tool-close').first().click({ force: true }).catch(() => {});
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
      await page.keyboard.press("Escape").catch(() => {});
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
        console.error(`  ❌ Adjustment voucher ${voucherNo} could not be opened after retry — skipping.`);
        return { ok: false, voucherNo };
      }

      console.log(`  🔄 Retrying: re-paste + click "${viewLinkText}"...`);
      try {
        await invInput.click();
        await invInput.fill("");
        await invInput.fill(voucherNo);
        await frame.getByText(viewLinkText, { exact: true }).first().click({ force: true });
      } catch (e) {
        console.error(`  ❌ Retry failed: ${e.message}`);
        return { ok: false, voucherNo };
      }
      continue;
    }

    console.log(`  ⏳ No popup or detail within 10s on attempt ${attempt}, falling through...`);
    break;
  }

  if (!detailLoaded) {
    const attached = await frame.locator('text="Credit Documents"').first()
      .waitFor({ state: "attached", timeout: 15000 })
      .then(() => true)
      .catch(() => false);
    if (!attached) {
      // Sanity-dismiss any lingering error popup so it doesn't bleed into the
      // next iteration, then skip this adjustment voucher instead of pressing on.
      await dismissVirtualWorksErrorPopup(page).catch(() => {});
      console.error(`  ❌ Detail screen never attached for ${voucherNo} — skipping.`);
      return { ok: false, voucherNo };
    }
  }
  await page.waitForTimeout(1000);
  return { ok: true, voucherNo };
}

/**
 * Scrapes the View Adjustment Voucher header (Voucher Information + Customer
 * Information) and writes the result to `<exportDir>/Header Info - <voucherNo>.txt`
 * as pretty-printed JSON.
 *
 * Must run BEFORE the Shift+E Excel export, since that keyboard sequence
 * shifts focus and can reflow the DOM. The extraction engine is label-agnostic,
 * so any label that isn't on the page simply yields null.
 */
async function saveVoucherInfoJson(page, voucherNo, exportDir) {
  console.log("  📋 Extracting Voucher / Customer Information → JSON...");
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
      voucher_information: [
        ["Voucher No.",   "voucher_no"],
        ["Voucher Date",  "voucher_date"],
        ["Voucher Type",  "voucher_type"],
        ["Status",        "status"],
        ["Finance Book",  "finance_book"],
      ],
      customer_information: [
        ["Customer Code",            "customer_code"],
        ["Customer Name",            "customer_name"],
        ["Customer Hierarchy",       "customer_hierarchy"],
        ["Credit Document Currency", "credit_document_currency"],
        ["Adjustment Sequence",      "adjustment_sequence"],
      ],
    };

    const numericKeys = new Set([]);

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
    // Never let the caller-supplied voucher number get lost to scraping misses.
    structured.voucher_information.voucher_no = voucherNo;

    const outPath = `${exportDir}/Header Info - ${voucherNo}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save adjustment voucher info JSON: ${error.message}`);
  }
}

/**
 * Returns true if the detail page's grid actually has data rows (and isn't
 * showing a "No records to display" placeholder). Mirrors the sibling scripts'
 * empty-grid guard so we don't sit waiting on a Shift+E export that will never
 * produce a download.
 */
async function pageHasGridData(page) {
  // Scan every frame and report "has data" if ANY frame shows at least one
  // VISIBLE grid row. We deliberately don't gate on "No records to display"
  // text: a drill-down page can carry several grids (e.g. one populated + one
  // empty placeholder), and that text would wrongly veto a grid that does have
  // rows. A truly empty ExtJS grid renders zero .x-grid3-row elements, so a
  // visible-row check is the reliable signal.
  for (const f of page.frames()) {
    try {
      const hasVisibleRow = await f.evaluate(() => {
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== 'none' && s.visibility !== 'hidden';
        };
        for (const r of document.querySelectorAll('.x-grid3-row')) {
          if (isVisible(r)) return true;
        }
        return false;
      });
      if (hasVisibleRow) return true;
    } catch (_) { /* detached/blank frame — try the next */ }
  }
  return false;
}

/**
 * Exports a single grid that sits under a given section header (e.g. "Credit
 * Documents" / "Debit Documents") via Shift+E, saving it as
 * `<fileLabel> - (<voucherNo>).<ext>`.
 *
 * Focus approach: a REAL Playwright .click() on the FIRST grid row that follows
 * the section header (not a synthetic MouseEvent — those don't move keyboard
 * focus, which leaves Shift+E with nothing focused and the download hangs).
 * Targeting "the row after this header" is what lets us export each grid
 * independently on a page that has more than one grid.
 *
 * Retry adds a Tab to lift focus from the row cell to the grid container.
 * Returns true on success, false if the grid is empty or the export fails.
 */
async function downloadGridAfterHeader(page, headerText, fileLabel, voucherNo, exportDir) {
  const frame = page.frameLocator("#ilboinnerframe");

  // First data row belonging to the grid that follows this section header.
  const rowAfterHeader = frame.locator(
    `xpath=(//*[normalize-space(text())="${headerText}"]/following::*[contains(@class,"x-grid3-row")])[1]`
  ).first();

  if (!(await rowAfterHeader.isVisible({ timeout: 4000 }).catch(() => false))) {
    console.log(`  ⏭ "${headerText}" grid has no rows — skipping ${fileLabel} export.`);
    return false;
  }

  const MAX_ATTEMPTS = 2;
  let lastErr = null;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      if (attempt > 1) {
        console.log(`  🔄 Retrying ${fileLabel} export (attempt ${attempt}/${MAX_ATTEMPTS})...`);
        await page.waitForTimeout(800);
      }

      console.log(`  🖱 Clicking first row of the "${headerText}" grid (real click → moves focus)...`);
      try {
        await rowAfterHeader.click({ force: true, timeout: 5000 });
      } catch (e) {
        console.warn(`  ⚠ Grid-row click failed (${e.message}).`);
      }
      await page.waitForTimeout(200);

      if (attempt > 1) {
        console.log("  ⌨ Pressing Tab once to park focus on the grid container...");
        await page.keyboard.press("Tab");
        await page.waitForTimeout(150);
      }

      console.log("  ⌨ Triggering Excel Export (Shift + E)...");
      const downloadPromise = page.waitForEvent('download', { timeout: 30000 });

      await page.keyboard.down('Shift');
      await page.keyboard.press('E');
      await page.keyboard.up('Shift');

      const download = await downloadPromise;
      const suggestedName = download.suggestedFilename();
      console.log(`  🔍 Server suggested name: "${suggestedName}"`);

      const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
      const fileName = `${exportDir}/${fileLabel} - (${voucherNo}).${ext}`;

      await download.saveAs(fileName);
      console.log(`  ✅ ${fileLabel} export successful!`);
      console.log(`  📁 Saved → ${fileName}`);
      return true;
    } catch (e) {
      lastErr = e;
      console.warn(`  ⚠ ${fileLabel} export attempt ${attempt}/${MAX_ATTEMPTS} failed: ${e.message}`);
    }
  }
  console.error(`  ❌ ${fileLabel} export failed after ${MAX_ATTEMPTS} attempts: ${lastErr ? lastErr.message : 'unknown'}`);
  return false;
}

/**
 * Dumps the header fields to JSON, then exports both line-item grids on the
 * View Adjustment Voucher detail screen — "Credit Documents" and "Debit
 * Documents" — each as its own Excel file.
 */
async function downloadVoucherDetail(page, voucherNo = "Unknown") {
  console.log(`\n📥 Preparing to download Adjustment Voucher detail (Voucher: ${voucherNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });

    await page.waitForTimeout(1500);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Adjustment Vouchers/${OU_FOLDER}/${dateFolder}/${voucherNo}`;
    mkdirSync(exportDir, { recursive: true });

    // Dump the header fields to JSON BEFORE the Shift+E sequence, which shifts
    // keyboard focus and can reflow the DOM.
    await saveVoucherInfoJson(page, voucherNo, exportDir);

    // Export each grid independently by targeting the row under its header.
    await downloadGridAfterHeader(page, "Credit Documents", "Credit Documents", voucherNo, exportDir);
    await downloadGridAfterHeader(page, "Debit Documents", "Debit Documents", voucherNo, exportDir);

  } catch (error) {
    console.error(`  ❌ Excel download failed: ${error.message}`);
  }
}

/**
 * Detects and closes any VirtualWorks/Ramco modal error popup. Used by
 * openVoucherViaDirectEntry to dismiss "Adjustment Voucher Number does not
 * exist" or similar errors before retrying.
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
    } catch (_) {}
    if (closed) {
      console.log("  ✔ Closed VirtualWorks error popup.");
      await page.waitForTimeout(500);
      return true;
    }
  }
  return false;
}

/**
 * Navigates to the "Accounting Information" sub-screen via the hyperlink at the
 * bottom-left of the View Adjustment Voucher detail screen (the action-links
 * block under "Reversal Details").
 *
 * Ramco renders these "hyperlinks" as styled <span>/<div> rather than real
 * <a> tags, so we use the same JS-dispatch trick (smallest visible leaf whose
 * direct text matches) with a Playwright locator fallback.
 *
 * This is a voucher-level drill-down (the GL postings for the whole voucher),
 * so — unlike the Transfer Note's row-scoped "View Account Postings" — no grid
 * row needs to be selected first.
 *
 * Returns false if the link can't be found/clicked, so the caller can skip the
 * accounting export and continue.
 */
async function goToAccountingInformation(page) {
  console.log("\n📑 Navigating to 'Accounting Information' link...");

  // Let the View Adjustment Voucher detail page fully settle before drilling in.
  // Clicking the link too soon triggers a Ramco-side "ValidateView" error popup
  // because the view hasn't finished initializing.
  console.log("  ⏳ Settling 2s before opening Accounting Information...");
  await page.waitForTimeout(2000);

  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  const LINK_TEXT = 'Accounting Information';
  let clicked = false;

  if (actualFrame) {
    try {
      clicked = await actualFrame.evaluate((linkText) => {
        const isVisible = (el) => {
          if (!el || el.offsetParent === null) return false;
          const s = window.getComputedStyle(el);
          return s.display !== 'none' && s.visibility !== 'hidden';
        };
        // Pick the SMALLEST visible element whose direct text exactly matches —
        // the leaf hyperlink, not a wrapping container.
        let best = null;
        let bestArea = Infinity;
        for (const el of document.querySelectorAll('*')) {
          if (!isVisible(el)) continue;
          const directText = Array.from(el.childNodes)
            .filter(n => n.nodeType === Node.TEXT_NODE)
            .map(n => n.textContent.trim())
            .join('');
          if (directText !== linkText) continue;
          const r = el.getBoundingClientRect();
          const area = r.width * r.height;
          if (area > 0 && area < bestArea) { best = el; bestArea = area; }
        }
        if (!best) return false;
        best.scrollIntoView({ block: 'center' });
        const r = best.getBoundingClientRect();
        const x = r.left + r.width / 2, y = r.top + r.height / 2;
        ['mousedown', 'mouseup', 'click'].forEach(t => {
          best.dispatchEvent(new MouseEvent(t, {
            bubbles: true, cancelable: true, view: window, clientX: x, clientY: y, button: 0,
          }));
        });
        return true;
      }, LINK_TEXT);
      if (clicked) console.log(`  ✔ Clicked '${LINK_TEXT}' via JS dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      const link = frame.getByText(LINK_TEXT, { exact: true }).last();
      await link.waitFor({ state: "visible", timeout: 5000 });
      await link.scrollIntoViewIfNeeded().catch(() => {});
      await link.click({ force: true });
      clicked = true;
      console.log(`  ✔ Clicked '${LINK_TEXT}' via locator (last).`);
    } catch (e) {
      console.warn(`  ⚠ No '${LINK_TEXT}' link found: ${e.message}`);
      return false;
    }
  }

  // The sub-screen renders effectively instantly, so poll fast on a short cap.
  // "Ready" = a grid row is visible AND we've left the voucher detail (its
  // "Credit Documents"/"Reversal Details" headers are gone) OR the breadcrumb
  // explicitly shows "Accounting Information". This breaks in well under a
  // second and never depends on a single fragile signal.
  const timeoutMs = Math.min(CONFIG.accountingLoadTimeoutMs, 5000);
  console.log(`  ⏳ Waiting for Accounting Information page to render (up to ${Math.round(timeoutMs / 1000)}s)...`);
  const accDeadline = Date.now() + timeoutMs;
  let accLoaded = false;
  while (Date.now() < accDeadline) {
    let ready = false;
    for (const f of page.frames()) {
      try {
        ready = await f.evaluate(() => {
          const isVisible = (el) => {
            if (!el || el.offsetParent === null) return false;
            const s = window.getComputedStyle(el);
            return s.display !== 'none' && s.visibility !== 'hidden';
          };
          // Any visible grid row yet?
          let hasRow = false;
          for (const r of document.querySelectorAll('.x-grid3-row')) {
            if (isVisible(r)) { hasRow = true; break; }
          }
          if (!hasRow) return false;
          // Direct-text scan once, reused for both checks below.
          const directTexts = [];
          for (const el of document.querySelectorAll('div, span, td, legend, .x-panel-header-text')) {
            if (!isVisible(el)) continue;
            const t = Array.from(el.childNodes)
              .filter(n => n.nodeType === Node.TEXT_NODE)
              .map(n => n.textContent.trim()).join('');
            if (t) directTexts.push(t);
          }
          // (a) Breadcrumb / header explicitly names the page.
          if (directTexts.some(t => /^Accounting Information$/i.test(t))) return true;
          // (b) We've navigated away from the voucher detail (its section
          //     headers are no longer present) and a grid is showing.
          const detailHeaders = ['Credit Documents', 'Debit Documents', 'Reversal Details'];
          const stillOnDetail = directTexts.some(t => detailHeaders.includes(t));
          return !stillOnDetail;
        });
      } catch (_) {}
      if (ready) break;
    }
    if (ready) { accLoaded = true; break; }
    await page.waitForTimeout(100);
  }
  if (accLoaded) {
    console.log("  ✔ Accounting Information grid rendered.");
    // Extra settle so the view is fully initialized before Tab/Shift+E fire.
    await page.waitForTimeout(1500);
  } else {
    console.warn(`  ⚠ Accounting Information grid not detected within ${Math.round(timeoutMs / 1000)}s; proceeding anyway.`);
    await page.waitForTimeout(1000);
  }
  return true;
}

/**
 * Downloads the Excel file from the Accounting Information sub-screen for a
 * voucher. Same focus strategy as the document grids: a real Playwright
 * .click() on the first grid row moves keyboard focus into the grid, then
 * Shift+E. Retry adds a Tab to lift focus to the grid container.
 */
async function downloadAccountingInformationExcel(page, voucherNo = "Unknown") {
  console.log(`\n📊 Preparing to download Accounting Information Excel (Voucher: ${voucherNo})...`);

  const dateFolder = new Date().toISOString().split('T')[0];
  const exportDir = `output/Adjustment Vouchers/${OU_FOLDER}/${dateFolder}/${voucherNo}`;
  mkdirSync(exportDir, { recursive: true });

  if (!(await pageHasGridData(page))) {
    console.log("  ⏭ Accounting Information has no rows — skipping export.");
    return;
  }

  // Focus approach mirrors index_sundry_receipt_full.js's downloadAccountingExcel:
  // on the Accounting Information drill-down a single Tab parks focus on the
  // grid container (a grid-row click instead lands focus inside a cell, which
  // swallows Shift+E). No row click — just Tab, then Shift+E. Each retry presses
  // Tab again, so a second attempt nudges focus one more step if needed.
  const MAX_ATTEMPTS = 2;
  let lastErr = null;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      if (attempt > 1) {
        console.log(`  🔄 Retrying Accounting Information export (attempt ${attempt}/${MAX_ATTEMPTS})...`);
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
      const fileName = `${exportDir}/Accounting Information - (${voucherNo}).${ext}`;

      await download.saveAs(fileName);
      console.log(`  ✅ Accounting Information Excel download successful!`);
      console.log(`  📁 Saved as → ${fileName}`);
      return;
    } catch (e) {
      lastErr = e;
      console.warn(`  ⚠ Accounting Information export attempt ${attempt}/${MAX_ATTEMPTS} failed: ${e.message}`);
    }
  }
  console.error(`  ❌ Accounting Information export failed after ${MAX_ATTEMPTS} attempts: ${lastErr ? lastErr.message : 'unknown'}`);
}

/**
 * Clicks the "Go back" button to return to the previous screen
 */
async function goBackFromDetail(page) {
  console.log("\n🔙 Navigating back to previous screen...");

  try {
    const backBtn = page.locator('#back button, button[qtip="Go back"], .x-tool-back').first();

    await backBtn.waitFor({ state: "attached", timeout: 10000 });
    await backBtn.scrollIntoViewIfNeeded().catch(() => {});

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
    await frame.locator('.x-grid3-row').first().waitFor({ state: "visible", timeout: 10000 }).catch(() => {});
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
      }).catch(() => {});
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

  return `${months[m-1]} ${d} ${y}`;
}

/**
 * Downloads the initial "Select Adjustment Voucher" search-results grid. Uses
 * the date filter value as the filename anchor when present; falls back to
 * today's date when that label isn't on screen.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const dateField = frame.locator(
      'xpath=//label[contains(text(), "Document Date From")]/following::input[1] | ' +
      '//td[contains(text(), "Document Date From")]/following::input[1] | ' +
      '//label[contains(text(), "Voucher Date From")]/following::input[1] | ' +
      '//td[contains(text(), "Voucher Date From")]/following::input[1]'
    ).first();
    const rawDate = await dateField.inputValue().catch(() => "");
    console.log(`  📅 Detected date filter value: ${rawDate}`);
    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Adjustment Vouchers/${OU_FOLDER}/${dateFolder}`;
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
 * Read every adjustment voucher number from the initial Select Adjustment Voucher Excel
 * export. Ramco's Shift+E export is a ZIP-packaged .xlsx — we extract with
 * Windows' bundled bsdtar, then locate the "Adjustment Voucher No." column in the
 * header row and read every value below it. Reading by column (rather than
 * regex) keeps working when an OU uses a non-standard note format.
 */
function readVouchersFromExcel(xlsPath) {
  console.log(`\n📖 Reading adjustment voucher list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read adjustment voucher list.`);
    return [];
  }

  const tempDir = path.join(tmpdir(), `ramco-xlsx-${process.pid}-${Date.now()}`);
  mkdirSync(tempDir, { recursive: true });

  try {
    const tarExe = process.env.SystemRoot
      ? path.join(process.env.SystemRoot, 'System32', 'tar.exe')
      : 'C:\\Windows\\System32\\tar.exe';

    execFileSync(tarExe, ['-xf', xlsPath, '-C', tempDir], { stdio: 'pipe' });

    const decodeXml = (s) => String(s)
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
      .replace(/&amp;/g, '&');

    // Cells with t="s" reference sharedStrings.xml by index.
    const sharedStrings = [];
    const sharedPath = path.join(tempDir, 'xl', 'sharedStrings.xml');
    if (existsSync(sharedPath)) {
      const sharedXml = readFileSync(sharedPath, 'utf-8');
      const siRegex = /<(?:\w+:)?si[^>]*>([\s\S]*?)<\/(?:\w+:)?si>/g;
      let s;
      while ((s = siRegex.exec(sharedXml)) !== null) {
        const tRegex = /<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
        let acc = '';
        let t;
        while ((t = tRegex.exec(s[1])) !== null) acc += t[1];
        sharedStrings.push(decodeXml(acc));
      }
    }

    // "AA" → 26, "C" → 2, etc. (0-based column index).
    const colLettersToIdx = (letters) => {
      let n = 0;
      for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
      return n - 1;
    };

    // Parse one <row>'s cells into a sparse array indexed by column position.
    // Cells without an `r=` attribute are positional — they take the slot
    // immediately after the previous cell.
    const parseRow = (rowXml) => {
      const cells = [];
      const cellRegex = /<(?:\w+:)?c\s*([^/>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
      let nextPos = 0;
      let m;
      while ((m = cellRegex.exec(rowXml)) !== null) {
        const attrs = m[1] || '';
        const inner = m[2] || '';
        const rMatch = attrs.match(/r="([A-Z]+)\d+"/);
        const pos = rMatch ? colLettersToIdx(rMatch[1]) : nextPos;
        nextPos = pos + 1;

        const tMatch = attrs.match(/t="([^"]+)"/);
        const type = tMatch ? tMatch[1] : '';
        let val = '';

        if (type === 'inlineStr') {
          const isMatch = inner.match(/<(?:\w+:)?is>([\s\S]*?)<\/(?:\w+:)?is>/);
          if (isMatch) {
            const tRegex = /<(?:\w+:)?t[^>]*>([\s\S]*?)<\/(?:\w+:)?t>/g;
            let tm;
            while ((tm = tRegex.exec(isMatch[1])) !== null) val += tm[1];
            val = decodeXml(val);
          }
        } else {
          const valMatch = inner.match(/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/);
          val = valMatch ? valMatch[1] : '';
          if (type === 's') {
            const idx = parseInt(val, 10);
            val = sharedStrings[idx] || '';
          } else {
            val = decodeXml(val);
          }
        }

        cells[pos] = val;
      }
      return cells;
    };

    const seen = new Set();
    const out = [];
    let foundHeader = false;

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
        const xml = readFileSync(path.join(sheetsDir, f), 'utf-8');
        const rowRegex = /<(?:\w+:)?row[^>]*>([\s\S]*?)<\/(?:\w+:)?row>/g;
        const rows = [];
        let r;
        while ((r = rowRegex.exec(xml)) !== null) rows.push(parseRow(r[1]));

        let headerRowIdx = -1;
        let noteColIdx = -1;
        for (let i = 0; i < rows.length && headerRowIdx < 0; i++) {
          for (let c = 0; c < rows[i].length; c++) {
            const v = (rows[i][c] || '').trim();
            // The results grid identifier column — "Voucher No." on this Select
            // screen; accept "Doc. No." / "Document No." as fallbacks.
            if (/^(Voucher|Document|Doc\.?)\s*No\.?$/i.test(v)) {
              headerRowIdx = i;
              noteColIdx = c;
              break;
            }
          }
        }

        if (headerRowIdx < 0) continue;
        foundHeader = true;

        for (let i = headerRowIdx + 1; i < rows.length; i++) {
          const val = (rows[i][noteColIdx] || '').trim();
          if (val && !seen.has(val)) {
            seen.add(val);
            out.push(val);
          }
        }
        if (out.length > 0) break;
      }
    }

    if (!foundHeader) {
      console.warn(`  ⚠ No "Voucher No." header found in any sheet — cannot extract list.`);
    }

    console.log(`  ✔ Extracted ${out.length} adjustment voucher(s) from Excel.`);
    return out;
  } catch (e) {
    console.error(`  ❌ Failed to read Excel: ${e.message}`);
    return [];
  } finally {
    try { rmSync(tempDir, { recursive: true, force: true }); } catch (_) {}
  }
}

// ─── MAIN ──────────────────────────────────────────────────────────────────
async function main() {
  mkdirSync("screenshots", { recursive: true });
  mkdirSync("output", { recursive: true });

  const browser = await chromium.launch({ headless: CONFIG.headless, slowMo: CONFIG.slowMo });
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page    = await context.newPage();

  try {
    const startTime = Date.now();
    await login(page);
    await changeContext(page, CONFIG.ouInstance);
    await goToLandingPage(page);

    const allResults = [];
    for (const menuPath of MENU_PATHS) {
      const result = await navigateMenu(page, menuPath);

      if (menuPath.join(" ").includes("View Adjustment")) {
        const stats = await performVoucherSearch(page);

        if (CONFIG.onlyVoucher) {
          // Single-note mode: skip Excel, process just this one via Direct Entry.
          const voucherNo = CONFIG.onlyVoucher;
          console.log(`\n🎯 ONLY_ADJUSTMENT_VOUCHER="${voucherNo}" — bypassing Excel list and processing this single voucher only.`);
          process.stdout.write(`\n📍 [Single]... ${voucherNo}`);

          const openResult = await openVoucherViaDirectEntry(page, voucherNo);
          console.log(`\n📂 --- PROCESSING ADJUSTMENT VOUCHER ${voucherNo} (single) ---`);

          if (!openResult.ok) {
            console.error(`  ❌ Could not open adjustment voucher ${voucherNo} — aborting.`);
          } else {
            // 1) Header JSON + Credit Documents & Debit Documents grids → Excel.
            await downloadVoucherDetail(page, voucherNo);

            // 2) Drill into "Accounting Information" and export that grid.
            const accountingOk = await goToAccountingInformation(page);
            if (accountingOk) {
              await downloadAccountingInformationExcel(page, voucherNo);
              console.log("  🔙 Returning from Accounting Information to View Adjustment Voucher...");
              await goBackFromDetail(page);
            }

            console.log("  🔙 Navigating back to search results...");
            await goBackFromDetail(page);

            console.log(`  ✅ Single adjustment voucher completed.`);
          }
        } else {
          // Download the initial Excel — contains every adjustment voucher in the result set.
          const initialXlsPath = await downloadInitialSearchResults(page);

          // Read the adjustment voucher list from that Excel.
          const voucherNumbers = readVouchersFromExcel(initialXlsPath);
          const totalRows = voucherNumbers.length;

          if (totalRows === 0) {
            console.error("  ❌ No adjustment vouchers extracted from Excel — cannot process.");
          } else {
            if (stats.total && stats.total !== totalRows) {
              console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
            }

            // Honor RESUME_FROM_ADJUSTMENT_VOUCHER
            let startIndex = 1;
            if (CONFIG.resumeFromVoucher) {
              const target = CONFIG.resumeFromVoucher.toUpperCase();
              const matchIdx = voucherNumbers.findIndex(v => String(v).trim().toUpperCase() === target);
              if (matchIdx === -1) {
                console.error(`  ❌ RESUME_FROM_ADJUSTMENT_VOUCHER="${CONFIG.resumeFromVoucher}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
                break;
              }
              startIndex = matchIdx + 1;
              console.log(`  ⏭ RESUME_FROM_ADJUSTMENT_VOUCHER="${CONFIG.resumeFromVoucher}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
            }

            console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} adjustment voucher(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

            for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
              const voucherNo = voucherNumbers[processedCount - 1];
              process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${voucherNo}`);

              const openResult = await openVoucherViaDirectEntry(page, voucherNo);
              console.log(`\n📂 --- PROCESSING ADJUSTMENT VOUCHER ${voucherNo} (${processedCount} OF ${totalRows}) ---`);

              if (!openResult.ok) {
                console.error(`  ❌ Could not open adjustment voucher ${voucherNo} — skipping.`);
                continue;
              }

              // 1) Header JSON + Credit Documents & Debit Documents grids → Excel.
              await downloadVoucherDetail(page, voucherNo);

              // 2) Drill into "Accounting Information" and export that grid.
              const accountingOk = await goToAccountingInformation(page);
              if (accountingOk) {
                await downloadAccountingInformationExcel(page, voucherNo);
                console.log("  🔙 Returning from Accounting Information to View Adjustment Voucher...");
                await goBackFromDetail(page);
              }

              // 3) Back from View Adjustment Voucher to Select Adjustment Voucher search grid.
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
          result.textSample = await innerFrame.evaluate(() => ((document.body && document.body.innerText) || "").trim().slice(0, 500));
        }
      }

      allResults.push(result);
    }

    writeFileSync("output/extracted-data-adjustment-voucher.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-adjustment-voucher.json");

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
