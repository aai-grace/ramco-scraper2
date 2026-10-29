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
// output/Sundry Receipts/. Unknown OUs fall back to a sanitized
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
  // Dedicated OU Instance for the Sundry Receipt workflow, so it
  // can run against a different OU than the other scripts.
  ouInstance: (process.env.OU_INSTANCE_SUNDRY_RECEIPT || "").trim(),
  headless: false,
  slowMo:   50,
  timeout:  30000,

  // Select Receipt search filters (dd/mm/yyyy — same format shown in the
  // Ramco UI). The Sundry Receipt screen filters by "Receipt Date From/To".
  receiptDateFrom: process.env.RECEIPT_DATE_FROM_SUNDRY_RECEIPT || "01/02/2026",
  receiptDateTo:   process.env.RECEIPT_DATE_TO_SUNDRY_RECEIPT   || "28/02/2026",

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // receipt number and begins processing from there. Blank = start at row 1.
  resumeFromReceipt: (process.env.RESUME_FROM_SUNDRY_RECEIPT || "").trim(),

  // Single-receipt mode: if set, the Excel download/list is bypassed entirely
  // and only this exact Receipt No. is processed via Direct Entry. Takes
  // precedence over RESUME_FROM_SUNDRY_RECEIPT.
  onlyReceipt: (process.env.ONLY_SUNDRY_RECEIPT || "").trim(),

  // How long to wait for the Accounting Information grid to render after
  // clicking the footer link. Configurable via env for one-off slow runs.
  accountingLoadTimeoutMs: Number(process.env.ACCOUNTING_LOAD_TIMEOUT_MS) || 60000,
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Sundry Receipts/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Sundry Receipts/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Sundry Receipt workflow:
// Receivable Management → Sundry Receipt → View Receipt
const MENU_PATHS = [
  ["Receivable Management", "Sundry Receipt", "View Receipt"],
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
 * Orchestrates the Search workflow for the "Select Receipt" screen of the
 * Sundry Receipt module. Filters by Receipt Date From/To, then clicks Search.
 */
async function performReceiptSearch(page) {
  console.log("\n🔍 Preparing to search sundry receipts inside frame...");
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

    console.log(`  📅 Filling Receipt Date: ${CONFIG.receiptDateFrom} → ${CONFIG.receiptDateTo}`);
    await fillDateField("Receipt Date From", 1, CONFIG.receiptDateFrom);
    await fillDateField("Receipt Date From", 2, CONFIG.receiptDateTo);
    console.log('  ✅ Receipt Date range set.');

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
 * Receipt-number text pattern: e.g. "08SREC26-000013".
 * Matches 2 digits + uppercase letters + 2 digits + "-" + digits.
 */
const RECEIPT_NO_PATTERN = /^\d{2}[A-Z]+\d{2}-\d+$/;

/**
 * Opens a receipt's detail screen by pasting its Receipt No. into Direct
 * Entry and clicking the "View Receipt" link.
 *
 * Returns { ok: boolean, receiptNo: string|null }.
 */
async function openReceiptViaDirectEntry(page, receiptNo) {
  console.log(`\n🖱 Opening receipt via Direct Entry — ${receiptNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  const viewLinkText = "View Receipt";
  console.log(`  🎯 Will click footer link: "${viewLinkText}"`);

  // Paste into "Receipt No." Direct Entry field.
  console.log(`  ⌨ Pasting into Direct Entry "Receipt No." field...`);
  const invInput = frame
    .locator('xpath=//*[normalize-space(text())="Receipt No."]/following::input[1]')
    .first();
  try {
    await invInput.waitFor({ state: "visible", timeout: 5000 });
    await invInput.click();
    await invInput.fill("");
    await invInput.fill(receiptNo);
    console.log(`  ✔ Pasted ${receiptNo}.`);
  } catch (e) {
    console.error(`  ❌ Failed to paste into Direct Entry: ${e.message}`);
    return { ok: false, receiptNo };
  }

  // Click the "View Receipt" footer link. Primary: JS dispatch on the
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
      return { ok: false, receiptNo };
    }
  }

  // Poll for one of two outcomes:
  //   (a) detail page attaches → continue
  //   (b) any Ramco/VirtualWorks error popup (e.g. "Receipt Number does not
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
      } catch (_) {}
      if (found) return { kind: 'dialog', panel: null, text: found.text };
    }
    return null;
  };

  const isDetailVisible = async () => {
    for (const t of ['Voucher Information', 'Instrument Information', 'Accounting Information']) {
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
        console.error(`  ❌ Receipt ${receiptNo} could not be opened after retry — skipping.`);
        return { ok: false, receiptNo };
      }

      console.log(`  🔄 Retrying: re-paste + click "${viewLinkText}"...`);
      try {
        await invInput.click();
        await invInput.fill("");
        await invInput.fill(receiptNo);
        await frame.getByText(viewLinkText, { exact: true }).first().click({ force: true });
      } catch (e) {
        console.error(`  ❌ Retry failed: ${e.message}`);
        return { ok: false, receiptNo };
      }
      continue;
    }

    console.log(`  ⏳ No popup or detail within 10s on attempt ${attempt}, falling through...`);
    break;
  }

  if (!detailLoaded) {
    const attached = await frame.locator('text="Voucher Information"').first()
      .waitFor({ state: "attached", timeout: 15000 })
      .then(() => true)
      .catch(() => false);
    if (!attached) {
      // Sanity-dismiss any lingering error popup so it doesn't bleed into the
      // next iteration, then skip this receipt instead of pressing on.
      await dismissVirtualWorksErrorPopup(page).catch(() => {});
      console.error(`  ❌ Detail screen never attached for ${receiptNo} — skipping.`);
      return { ok: false, receiptNo };
    }
  }
  await page.waitForTimeout(1000);
  return { ok: true, receiptNo };
}

/**
 * Scrapes the View Receipt header (Voucher / Project / Instrument
 * Information) and writes the result to `<exportDir>/Header Info - <receiptNo>.txt`
 * as pretty-printed JSON.
 *
 * Must run BEFORE the Shift+E Excel export, since that keyboard sequence
 * shifts focus and can reflow the DOM.
 */
async function saveReceiptInfoJson(page, receiptNo, exportDir) {
  console.log("  📋 Extracting Voucher / Project / Instrument Information → JSON...");
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
        ["Receipt No.",            "receipt_no"],
        ["Receipt Category",       "receipt_category"],
        ["Status",                 "status"],
        ["Receipt Date",           "receipt_date"],
        ["Finance Book",           "finance_book"],
        ["Receipt Method",         "receipt_method"],
        ["Receipt Route",          "receipt_route"],
        ["Receipt Mode",           "receipt_mode"],
        ["Currency",               "currency"],
        ["Exchange Rate",          "exchange_rate"],
        ["Remitter",               "remitter"],
        ["Bank / Cash Code",       "bank_cash_code"],
        ["Description",            "description"],
        ["Receipt Amount",         "receipt_amount"],
        ["Reference Document No.", "reference_document_no"],
        ["Remarks",                "remarks"],
        ["Sundry Payment No",      "sundry_payment_no"],
      ],
      project_information: [
        ["Project OU",                        "project_ou"],
        ["Project / Sub-Project Code",        "project_sub_project_code"],
        ["Project / Sub-Project Description", "project_sub_project_description"],
      ],
      instrument_information: [
        ["Instrument No.",     "instrument_no"],
        ["MICR No.",           "micr_no"],
        ["Instrument Amount",  "instrument_amount"],
        ["Instrument Date",    "instrument_date"],
        ["Bank Code",          "bank_code"],
        ["Instrument Status",  "instrument_status"],
        ["Instrument Type",    "instrument_type"],
      ],
    };

    const numericKeys = new Set([
      "exchange_rate",
      "receipt_amount",
      "instrument_amount",
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
    // Never let the caller-supplied receipt number get lost to scraping misses.
    structured.voucher_information.receipt_no = receiptNo;

    const outPath = `${exportDir}/Header Info - ${receiptNo}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save receipt info JSON: ${error.message}`);
  }
}

/**
 * Downloads the Receipt detail Excel file from the View Receipt screen.
 * Captures the on-page detail via Shift+E. The drill-down Accounting
 * Information page is exported separately (see goToAccountingInformation +
 * downloadAccountingExcel).
 */
async function downloadReceiptDetail(page, receiptNo = "Unknown") {
  console.log(`\n📥 Preparing to download Receipt Detail Excel (Receipt: ${receiptNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });

    await page.waitForTimeout(1500);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Sundry Receipts/${OU_FOLDER}/${dateFolder}/${receiptNo}`;
    mkdirSync(exportDir, { recursive: true });

    // Dump Voucher/Project/Instrument Information to JSON BEFORE the
    // Shift+E sequence, which shifts keyboard focus and can reflow the DOM.
    await saveReceiptInfoJson(page, receiptNo, exportDir);

    // Seat focus on the Accounting Information grid (which is embedded
    // directly on the View Receipt page) so Shift+E exports it correctly.
    console.log("  🖱 Clicking first grid row to focus the Accounting grid...");
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

    const fileName = `${exportDir}/Receipt Info - (${receiptNo}).${ext}`;

    await download.saveAs(fileName);
    console.log(`  ✅ Excel download successful!`);
    console.log(`  📁 Saved and renamed properly in your project → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Excel download failed: ${error.message}`);
  }
}

/**
 * Detects and closes any VirtualWorks/Ramco modal error popup. Used by
 * openReceiptViaDirectEntry to dismiss "Receipt Number does not exist" or
 * similar errors before retrying.
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
 * Navigates to the dedicated Accounting Information page via the footer
 * hyperlink at the bottom of the View Receipt detail screen.
 *
 * The text "Accounting Information" can appear twice on the page (the
 * inline section header and the footer link). We target the LAST match in
 * DOM order, which is the footer link.
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
      await link.scrollIntoViewIfNeeded().catch(() => {});
      await link.click({ force: true });
      clicked = true;
      console.log("  ✔ Clicked 'Accounting Information' via locator (last).");
    } catch (e) {
      console.error(`  ❌ Failed to click Accounting Information: ${e.message}`);
      return false;
    }
  }

  const timeoutMs = CONFIG.accountingLoadTimeoutMs;
  console.log(`  ⏳ Waiting for Accounting Information page to load (up to ${Math.round(timeoutMs / 1000)}s)...`);
  // Poll for signals that we're on the Accounting Information screen — accept
  // either breadcrumb confirmation (right page, grid still populating) OR
  // grid column headers (fully loaded).
  const accDeadline = Date.now() + timeoutMs;
  let accLoaded = false;
  let lastSignal = "none";
  while (Date.now() < accDeadline) {
    let signal = null;
    for (const f of page.frames()) {
      try {
        signal = await f.evaluate(() => {
          const isVisible = (el) => {
            if (!el || el.offsetParent === null) return false;
            const s = window.getComputedStyle(el);
            return s.display !== 'none' && s.visibility !== 'hidden';
          };
          // (a) Accounting-grid column headers — fully loaded.
          const markers = ['Account Code', 'Account Description', 'Debit Amount', 'Credit Amount', 'Base Amount', 'Dr. / Cr.'];
          const headers = document.querySelectorAll('.x-grid3-hd-inner, .x-grid3-hd');
          let hits = 0;
          for (const h of headers) {
            if (!isVisible(h)) continue;
            const txt = (h.innerText || '').trim();
            if (markers.includes(txt)) hits++;
          }
          if (hits >= 2) return 'grid';
          // (b) Breadcrumb / page title shows "Accounting Information".
          for (const el of document.querySelectorAll(
            '.x-panel-header-text, .breadcrumb, .x-toolbar, h1, h2, .page-title'
          )) {
            if (!isVisible(el)) continue;
            const txt = (el.innerText || '').trim();
            if (/\bAccounting Information\b/i.test(txt)) return 'breadcrumb';
          }
          return null;
        });
      } catch (_) {}
      if (signal) break;
    }
    if (signal === 'grid') { accLoaded = true; lastSignal = signal; break; }
    if (signal === 'breadcrumb' && lastSignal !== 'breadcrumb') {
      console.log("  … On Accounting Information page; waiting for grid rows...");
      lastSignal = signal;
    }
    await page.waitForTimeout(300);
  }
  if (accLoaded) {
    console.log("  ✔ Accounting Information grid detected.");
    await page.waitForTimeout(500);
  } else {
    console.warn(`  ⚠ Accounting Information grid not detected within ${Math.round(timeoutMs / 1000)}s; proceeding anyway.`);
  }
  return true;
}

/**
 * Returns true if the current page (View Receipt or the Accounting
 * Information sub-screen) has at least one data row in its main grid. False
 * when the grid shows an empty-state marker like "Found no rows to display!!!"
 * or "[No records to display]" in the paging toolbar.
 *
 * Some receipts have no Accounting lines for a given transaction. Without
 * this check, Tab+Shift+E would either export an empty file or hang waiting
 * for a download event that never fires.
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
 * Click the first row of the currently-visible grid to anchor keyboard focus
 * on it before triggering Shift+E. Required on the Accounting Information
 * screen, which mixes form fields (the Transaction No. dropdown) with the
 * grid — Tab-based focus would otherwise land on the dropdown and Shift+E
 * would type the literal "E" into it instead of exporting the grid.
 */
async function focusGridFirstRow(page) {
  const actualFrame = page.frames().find(
    (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
  );
  if (!actualFrame) return false;
  try {
    const firstRow = actualFrame.locator('.x-grid3-row').first();
    await firstRow.waitFor({ state: 'visible', timeout: 3000 });
    await firstRow.click();
    await page.waitForTimeout(300);
    return true;
  } catch (e) {
    console.warn(`  ⚠ Could not focus grid first row: ${e.message}`);
    return false;
  }
}

/**
 * Tab×N + Shift+E export of the currently-visible grid. Saves to
 * `<exportDir>/<label> - (<receiptNo>).<ext>`. Caller is responsible for
 * placing the right grid on screen before this is invoked.
 *
 * Skips silently (and returns false) if the grid has no rows. Retries the
 * Tab + Shift+E + download sequence up to MAX_ATTEMPTS — transient failures
 * (focus drift, slow server, network blip) usually recover on a second try.
 */
async function exportCurrentGrid(page, label, receiptNo, exportDir, tabCount = 1) {
  if (!(await pageHasGridData(page))) {
    console.log(`  ⏭ ${label} grid has no rows — skipping export.`);
    return false;
  }

  const MAX_ATTEMPTS = 2;
  let lastErr = null;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      if (attempt > 1) {
        console.log(`  🔄 Retrying ${label} export (attempt ${attempt}/${MAX_ATTEMPTS})...`);
        await page.waitForTimeout(800);
      }

      if (tabCount > 0) {
        console.log(`  ⌨ Pressing Tab ${tabCount} time${tabCount === 1 ? "" : "s"} (${label} grid)...`);
        for (let i = 0; i < tabCount; i++) {
          await page.keyboard.press("Tab");
          await page.waitForTimeout(100);
        }
      }

      console.log("  ⌨ Triggering Excel Export (Shift + E)...");
      const downloadPromise = page.waitForEvent("download", { timeout: 30000 });
      await page.keyboard.down("Shift");
      await page.keyboard.press("E");
      await page.keyboard.up("Shift");

      const download = await downloadPromise;
      const suggestedName = download.suggestedFilename();
      console.log(`  🔍 Server suggested name: "${suggestedName}"`);

      const ext = suggestedName.includes(".") ? suggestedName.split(".").pop() : "xls";
      const fileName = `${exportDir}/${label} - (${receiptNo}).${ext}`;

      await download.saveAs(fileName);
      console.log(`  ✅ ${label} download successful!`);
      console.log(`  📁 Saved → ${fileName}`);
      return true;
    } catch (e) {
      lastErr = e;
      console.warn(`  ⚠ ${label} attempt ${attempt}/${MAX_ATTEMPTS} failed: ${e.message}`);
    }
  }
  console.error(`  ❌ ${label} export failed after ${MAX_ATTEMPTS} attempts: ${lastErr ? lastErr.message : "unknown"}`);
  return false;
}

/**
 * Enumerates every value in the "Transaction No." dropdown on the
 * Accounting Information screen.
 *
 * A single receipt can split into more than one accounting transaction
 * (e.g. "08RARDN26-000001", "08RARDN26-000002", …); each transaction has
 * its own Accounting Information grid. A single Shift+E only captures the
 * transaction currently selected, so we enumerate every option, select each
 * in turn, and export.
 *
 * Strategy:
 *   1. Read the ExtJS combo's internal store directly via
 *      `Ext.getCmp(input.id).store` — returns the complete list regardless
 *      of which records are virtualized in the DOM.
 *   2. Fallback: open the dropdown and scroll `.x-combo-list-inner`,
 *      collecting `.x-combo-list-item` text on each tick.
 */
async function getTransactionNoOptions(page) {
  console.log("  🔍 Detecting 'Transaction No.' options...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(
    (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
  );

  try {
    const field = frame
      .locator(
        'xpath=//label[contains(text(), "Transaction No.")]/following::input[1] | ' +
          '//td[contains(text(), "Transaction No.")]/following::input[1] | ' +
          '//span[contains(text(), "Transaction No.")]/following::input[1]'
      )
      .first();

    await field.waitFor({ state: "visible", timeout: 5000 });

    // Strategy 1: read the ExtJS combo store directly.
    let options = [];
    if (actualFrame) {
      try {
        options = await actualFrame.evaluate(() => {
          const labels = Array.from(document.querySelectorAll("label, td, span"));
          const labelEl = labels.find((el) => el.innerText.trim().includes("Transaction No."));
          if (labelEl) {
            const input =
              labelEl.closest("tr")?.querySelector("input") ||
              labelEl.parentElement?.querySelector("input") ||
              labelEl.closest("td")?.nextElementSibling?.querySelector("input");
            if (input && input.id) {
              const cmp = window.Ext?.getCmp(input.id);
              if (cmp && cmp.store) {
                const storeData = [];
                cmp.store.each((record) => {
                  const val = record.get(cmp.valueField || cmp.displayField || "text");
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
      // De-dupe while preserving order.
      const unique = [...new Set(options)];
      console.log(`  ✔ Found options: ${unique.length}`);
      console.log(`    Options: ${unique.join(", ")}`);
      return unique;
    }

    // Strategy 2: open dropdown + scroll-collect (ExtJS virtualizes items
    // beyond the visible window, so a single read misses most rows).
    console.log("    ℹ Falling back to dropdown scroll enumeration...");
    await field.click();
    await page.waitForTimeout(1000);

    const allOptions = new Set();

    let visibleItems = await page
      .locator(".x-combo-list-item")
      .evaluateAll((els) => els.map((el) => el.innerText.trim()).filter((txt) => txt.length > 0));
    visibleItems.forEach((item) => allOptions.add(item));

    const comboList = page.locator(".x-combo-list-inner").first();
    const isComboVisible = await comboList.isVisible({ timeout: 2000 }).catch(() => false);

    if (isComboVisible) {
      let previousSize = 0;
      let scrollAttempts = 0;
      const maxScrollAttempts = 50;

      while (scrollAttempts < maxScrollAttempts) {
        await comboList.evaluate((el) => (el.scrollTop += 100));
        await page.waitForTimeout(300);

        visibleItems = await page
          .locator(".x-combo-list-item")
          .evaluateAll((els) =>
            els.map((el) => el.innerText.trim()).filter((txt) => txt.length > 0)
          );
        visibleItems.forEach((item) => allOptions.add(item));

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

    const uniqueOptions = [...allOptions];
    console.log(`  ✔ Found options: ${uniqueOptions.length}`);
    console.log(`    Options: ${uniqueOptions.join(", ")}`);
    return uniqueOptions;
  } catch (error) {
    console.warn(`  ⚠ Could not detect Transaction No. options: ${error.message}. Defaulting to current view.`);
    return [];
  }
}

/**
 * Selects a specific value in the "Transaction No." dropdown on the
 * Accounting Information screen. After the selection Ramco reloads the
 * Accounting Information grid for the chosen transaction.
 */
async function selectTransactionNo(page, txnNo) {
  console.log(`  🔢 Selecting Transaction No: ${txnNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const field = frame
      .locator(
        'xpath=//label[contains(text(), "Transaction No.")]/following::input[1] | ' +
          '//td[contains(text(), "Transaction No.")]/following::input[1] | ' +
          '//span[contains(text(), "Transaction No.")]/following::input[1]'
      )
      .first();

    await field.waitFor({ state: "visible", timeout: 5000 });
    await field.click();
    await page.waitForTimeout(500);

    let clicked = false;

    const findAndClickExact = async () => {
      const items = page.locator(".x-combo-list-item");
      const count = await items.count();
      for (let i = 0; i < count; i++) {
        const text = await items.nth(i).innerText();
        if (text.trim() === String(txnNo).trim()) {
          await items.nth(i).click();
          return true;
        }
      }
      return false;
    };

    clicked = await findAndClickExact();

    if (!clicked) {
      const comboList = page.locator(".x-combo-list-inner").first();
      const isComboVisible = await comboList.isVisible({ timeout: 2000 }).catch(() => false);

      if (isComboVisible) {
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
      console.log(`    ℹ Scrolling didn't find item, using keyboard input...`);
      await field.focus();
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type(String(txnNo), { delay: 50 });
      await page.keyboard.press("Enter");
    }

    // Give Ramco time to reload the Accounting Information grid for the
    // newly-selected transaction before the caller exports it.
    await page.waitForTimeout(2500);
    return true;
  } catch (error) {
    console.error(`    ❌ Failed to select Transaction No. ${txnNo}: ${error.message}`);
    return false;
  }
}

/**
 * Downloads the Excel file(s) from the Accounting Information view for a
 * receipt. The caller (goToAccountingInformation) has already navigated to
 * the Accounting Information screen.
 *
 * A receipt can span multiple accounting transactions; each renders its own
 * grid behind the "Transaction No." dropdown, and a single Shift+E only
 * captures the transaction currently selected. So we enumerate every
 * Transaction No., select it, and export one Excel per transaction.
 *
 * Filenames:
 *   - Single / unknown: `Accounting Info - (<receiptNo>).xls`  (original name)
 *   - Multiple:         `Accounting Info - <txnNo> - (<receiptNo>).xls`
 */
async function downloadAccountingExcel(page, receiptNo = "Unknown") {
  console.log(`\n📊 Preparing to download Accounting Excel (Receipt: ${receiptNo})...`);

  const dateFolder = new Date().toISOString().split('T')[0];
  const exportDir = `output/Sundry Receipts/${OU_FOLDER}/${dateFolder}/${receiptNo}`;
  mkdirSync(exportDir, { recursive: true });

  const options = await getTransactionNoOptions(page);

  if (options.length <= 1) {
    // Couldn't enumerate, or there's only one transaction — single export of
    // the current view. Anchor focus on the grid first (the Transaction No.
    // dropdown sits above the grid, so Tab-based focus would otherwise land
    // on the dropdown and Shift+E would type "E" into it).
    await focusGridFirstRow(page);
    await exportCurrentGrid(page, "Accounting Info", receiptNo, exportDir, 0);
  } else {
    for (let i = 0; i < options.length; i++) {
      const option = options[i];
      console.log(`\n  📍 Transaction No. "${option}" (${i + 1}/${options.length})...`);

      const selected = await selectTransactionNo(page, option);
      if (!selected) {
        console.warn(`  ⚠ Could not select Transaction No. "${option}" — skipping export.`);
        continue;
      }

      // selectTransactionNo leaves focus on the dropdown — re-anchor on the grid.
      await focusGridFirstRow(page);
      const safeOption = String(option).replace(/[<>:"/\\|?*]/g, "_");
      await exportCurrentGrid(page, `Accounting Info - ${safeOption}`, receiptNo, exportDir, 0);
    }
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
 * Downloads the initial "Select Receipt" search-results grid. Uses
 * "Receipt Date From" as the filename anchor when present; falls back to
 * today's date when that label isn't on screen.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const dateField = frame.locator(
      'xpath=//label[contains(text(), "Receipt Date From")]/following::input[1] | ' +
      '//td[contains(text(), "Receipt Date From")]/following::input[1]'
    ).first();
    const rawDate = await dateField.inputValue().catch(() => "");
    console.log(`  📅 Detected Receipt Date From: ${rawDate}`);
    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Sundry Receipts/${OU_FOLDER}/${dateFolder}`;
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
 * Read every receipt number from the initial Select Receipt Excel export.
 * Ramco's Shift+E export is a ZIP-packaged .xlsx — we extract with Windows'
 * bundled bsdtar, then locate the "Receipt No." column in the header row
 * and read every value below it. Reading by column (rather than regex) keeps
 * working when an OU uses a non-SREC receipt format (e.g. AAI Project
 * Logistics displays receipts as plain numerics like "0005788").
 */
function readReceiptsFromExcel(xlsPath) {
  console.log(`\n📖 Reading receipt list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read receipt list.`);
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
        let receiptColIdx = -1;
        for (let i = 0; i < rows.length && headerRowIdx < 0; i++) {
          for (let c = 0; c < rows[i].length; c++) {
            const v = (rows[i][c] || '').trim();
            if (/^Receipt\s*No\.?$/i.test(v)) {
              headerRowIdx = i;
              receiptColIdx = c;
              break;
            }
          }
        }

        if (headerRowIdx < 0) continue;
        foundHeader = true;

        for (let i = headerRowIdx + 1; i < rows.length; i++) {
          const val = (rows[i][receiptColIdx] || '').trim();
          if (val && !seen.has(val)) {
            seen.add(val);
            out.push(val);
          }
        }
        if (out.length > 0) break;
      }
    }

    if (!foundHeader) {
      console.warn(`  ⚠ No "Receipt No." header found in any sheet — cannot extract list.`);
    }

    console.log(`  ✔ Extracted ${out.length} receipt(s) from Excel.`);
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

      if (menuPath.join(" ").includes("View Receipt")) {
        const stats = await performReceiptSearch(page);

        if (CONFIG.onlyReceipt) {
          // Single-receipt mode: skip Excel, process just this one via Direct Entry.
          const receiptNo = CONFIG.onlyReceipt;
          console.log(`\n🎯 ONLY_SUNDRY_RECEIPT="${receiptNo}" — bypassing Excel list and processing this single receipt only.`);
          process.stdout.write(`\n📍 [Single]... ${receiptNo}`);

          const openResult = await openReceiptViaDirectEntry(page, receiptNo);
          console.log(`\n📂 --- PROCESSING RECEIPT ${receiptNo} (single) ---`);

          if (!openResult.ok) {
            console.error(`  ❌ Could not open receipt ${receiptNo} — aborting.`);
          } else {
            await downloadReceiptDetail(page, receiptNo);

            const accountingOk = await goToAccountingInformation(page);
            if (accountingOk) {
              await downloadAccountingExcel(page, receiptNo);
              console.log("  🔙 Returning from Accounting to View Receipt...");
              await goBackFromDetail(page);
            }

            console.log("  🔙 Navigating back to search results...");
            await goBackFromDetail(page);

            console.log(`  ✅ Single receipt completed.`);
          }
        } else {
          // Download the initial Excel — contains every receipt in the result set.
          const initialXlsPath = await downloadInitialSearchResults(page);

          // Read the receipt list from that Excel.
          const receiptNumbers = readReceiptsFromExcel(initialXlsPath);
          const totalRows = receiptNumbers.length;

          if (totalRows === 0) {
            console.error("  ❌ No receipts extracted from Excel — cannot process.");
          } else {
            if (stats.total && stats.total !== totalRows) {
              console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
            }

            // Honor RESUME_FROM_SUNDRY_RECEIPT
            let startIndex = 1;
            if (CONFIG.resumeFromReceipt) {
              const target = CONFIG.resumeFromReceipt.toUpperCase();
              const matchIdx = receiptNumbers.findIndex(v => String(v).trim().toUpperCase() === target);
              if (matchIdx === -1) {
                console.error(`  ❌ RESUME_FROM_SUNDRY_RECEIPT="${CONFIG.resumeFromReceipt}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
                break;
              }
              startIndex = matchIdx + 1;
              console.log(`  ⏭ RESUME_FROM_SUNDRY_RECEIPT="${CONFIG.resumeFromReceipt}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
            }

            console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} receipt(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

            for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
              const receiptNo = receiptNumbers[processedCount - 1];
              process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${receiptNo}`);

              const openResult = await openReceiptViaDirectEntry(page, receiptNo);
              console.log(`\n📂 --- PROCESSING RECEIPT ${receiptNo} (${processedCount} OF ${totalRows}) ---`);

              if (!openResult.ok) {
                console.error(`  ❌ Could not open receipt ${receiptNo} — skipping.`);
                continue;
              }

              // 1) View Receipt detail export (Voucher/Project/Instrument JSON +
              //    Shift+E Excel of the on-page receipt detail).
              await downloadReceiptDetail(page, receiptNo);

              // 2) Accounting Information drill-down via the footer hyperlink.
              const accountingOk = await goToAccountingInformation(page);
              if (accountingOk) {
                await downloadAccountingExcel(page, receiptNo);
                console.log("  🔙 Returning from Accounting to View Receipt...");
                await goBackFromDetail(page);
              }

              // 3) Back from View Receipt to Select Receipt search grid.
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

    writeFileSync("output/extracted-data-sundry-receipt.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-sundry-receipt.json");

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
