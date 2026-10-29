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
// output/Recurring Voucher Templates/. Unknown OUs fall back to a sanitized
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

// Sanitize a template number for use in Windows file/dir names. The raw value
// (e.g. "A2Z/CHD-32") can contain characters invalid in paths — replace each
// of /\:*?"<>| with "_" (so "A2Z/CHD-32" → "A2Z_CHD-32"). The ORIGINAL value is
// still used for Direct Entry and inside the JSON content.
function safeFileName(name) {
  return String(name).replace(/[\/\\:*?"<>|]/g, "_").trim();
}

// ─── CONFIG (loaded from .env) ─────────────────────────────────────────────
const CONFIG = {
  url:      process.env.RAMCO_URL      || "https://aai.ramcoes.com/RVW/extui/vwrt/LaunchPanel.htm",
  username: process.env.RAMCO_USERNAME || "",
  password: process.env.RAMCO_PASSWORD || "",
  ouInstance: (process.env.OU_INSTANCE_RECURRING_VOUCHER || "").trim(),
  headless: false,
  slowMo:   50,
  timeout:  30000,

  // View Recurring Voucher Template search filters. The search screen filters
  // by Posting Frequency / User ID / Status (all default to "All") plus an
  // optional Effective Date range. Leave the dropdown values blank to keep
  // them on "All"; leave the dates blank to skip them.
  postingFrequency:  (process.env.POSTING_FREQUENCY_RECURRING_VOUCHER || "").trim(), // e.g. "Monthly" — blank = All
  status:            (process.env.STATUS_RECURRING_VOUCHER            || "").trim(),  // e.g. "Authorized" — blank = All
  effectiveDateFrom: (process.env.EFFECTIVE_DATE_FROM_RECURRING_VOUCHER || "").trim(), // dd/mm/yyyy — blank = skip
  effectiveDateTo:   (process.env.EFFECTIVE_DATE_TO_RECURRING_VOUCHER   || "").trim(), // dd/mm/yyyy — blank = skip

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // template number and begins processing from there. Blank = start at row 1.
  resumeFromTemplate: (process.env.RESUME_FROM_RECURRING_VOUCHER || "").trim(),
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Recurring Voucher Templates/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Recurring Voucher Templates/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Recurring Voucher Template workflow:
// Book-keeping → Journal Voucher → View Recurring Voucher Template
const MENU_PATHS = [
  ["Book-keeping", "Journal Voucher", "View Recurring Voucher Template"],
];

// Label of the final menu item / search-result screen. Used in log lines and
// to detect when the search/iterate logic should run inside main().
const MODULE_LABEL = "View Recurring Voucher Template";

// The blue hyperlink that opens a template's detail screen from the search
// results (top-right of the results area in the screenshot).
const VIEW_LINK_TEXT = "View Recurring Voucher Template";

// Candidate labels for the Direct Entry "number" field used to open a template
// by pasting its number. The script tries each in order. (NOTE: confirm the
// real caption on the screen — it is usually "Template No." or "Voucher No.")
const DIRECT_ENTRY_LABELS = ["Template No.", "Voucher No.", "Recurring Voucher Template No."];

// Section headers that signal the detail screen has attached. The script
// considers the detail loaded if ANY of these is visible. These are the actual
// section headers on the View Recurring Voucher Template detail screen.
const DETAIL_READY_TEXTS = [
  "Accounting Information",
  "Project Information",
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
 * Best-effort setter for an ExtJS combo/dropdown identified by its label.
 * Used for the "Posting Frequency" and "Status" filters on the View Recurring
 * Voucher Template search screen. No-op (returns false) if the value is blank
 * or the field can't be located, so the dropdown keeps its default ("All").
 */
async function setComboByLabel(frame, page, labelText, value) {
  if (!value) return false; // blank → leave default "All"

  const field = frame.locator(
    `xpath=//label[contains(text(), "${labelText}")]/following::input[1] | ` +
    `//td[contains(text(), "${labelText}")]/following::input[1] | ` +
    `//span[contains(text(), "${labelText}")]/following::input[1]`
  ).first();

  if (!(await field.isVisible({ timeout: 3000 }).catch(() => false))) {
    console.warn(`  ⚠️ Combo "${labelText}" not found — leaving default.`);
    return false;
  }

  await field.click();
  await page.keyboard.down('Control');
  await page.keyboard.press('a');
  await page.keyboard.up('Control');
  await page.keyboard.press('Backspace');
  await field.type(value, { delay: 30 });
  await page.waitForTimeout(700);

  const item = frame.locator(
    `.x-combo-list-item:has-text("${value}"), li:has-text("${value}")`
  ).first();
  if (await item.isVisible({ timeout: 1500 }).catch(() => false)) {
    await item.click();
  } else {
    await page.keyboard.press('Enter');
  }
  await page.waitForTimeout(300);
  console.log(`  ✅ ${labelText} set to "${value}".`);
  return true;
}

/**
 * Orchestrates the Search workflow for the "View Recurring Voucher Template"
 * screen (see screenshot): clears User ID, optionally sets Posting Frequency /
 * Status / Effective Date filters, then clicks Search.
 */
async function performTemplateSearch(page) {
  console.log("\n🔍 Preparing to search recurring voucher templates inside frame...");
  const frame = await getFrame(page);

  if (!frame) {
    throw new Error("Could not find ilboinnerframe");
  }

  try {
    // 1. Clear "User ID" field so the search isn't scoped to the logged-in user.
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
      console.warn('  ⚠️ Could not identify exact "User ID" input via label — continuing.');
    }

    // 2. Optionally set the Posting Frequency / Status dropdowns. Blank config
    //    leaves each on its default "All" value.
    await setComboByLabel(frame, page, "Posting Frequency", CONFIG.postingFrequency);
    await setComboByLabel(frame, page, "Status", CONFIG.status);

    // 3. Optionally fill the Effective Date range. Each row in the UI is:
    //    [<Label>From] [<input From>] [label "To"] [<input To>] — so
    //    following::input[1] = From field, following::input[2] = To field.
    const fillDateField = async (labelText, position, value) => {
      if (!value) return false;
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

    if (CONFIG.effectiveDateFrom || CONFIG.effectiveDateTo) {
      console.log(`  📅 Filling Effective Date: ${CONFIG.effectiveDateFrom || "(skip)"} → ${CONFIG.effectiveDateTo || "(skip)"}`);
      await fillDateField("Effective Date From", 1, CONFIG.effectiveDateFrom);
      await fillDateField("Effective Date From", 2, CONFIG.effectiveDateTo);
      console.log('  ✅ Effective Date range set.');
    }

    // 4. Click the Search button
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
 * Loose pattern used only by the (now-unused) F2 popup reader. Recurring
 * Voucher Template numbers are free-form codes (e.g. "A2Z/CHD-32",
 * "AFME2026", "ACCRL_DNTN_NIBT'24"), so the main flow reads the Template No.
 * column directly rather than matching a pattern.
 */
const TEMPLATE_NO_PATTERN = /^[A-Z0-9][A-Z0-9._/'\-]+$/i;

/**
 * Enumerates every Template No. in the search-results grid by reading the
 * "Template No." column directly. Recurring Voucher Template numbers are
 * free-form alphanumeric codes, so we resolve the column from its ExtJS header
 * id (class `x-grid3-hd-<id>` ↔ cell class `x-grid3-td-<id>`) and read that
 * column's value per row, walking every pager page.
 *
 * Returns an array of unique template numbers in grid (display) order.
 */
async function getTemplateNumbers(page) {
  console.log("\n🔢 Collecting template numbers from the grid...");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  if (!actualFrame) {
    console.error("  ❌ Could not find ilboinnerframe.");
    return [];
  }

  // Total is just an early-break optimization; pager "Next disabled" is the
  // real termination signal. The toolbar here reads "1 - 5 / 23", so grab the
  // trailing number from the paging info / toolbar text.
  const total = await actualFrame.evaluate(() => {
    const candidates = [
      document.querySelector('.x-paging-info')?.innerText,
      document.querySelector('.x-paging-toolbar')?.innerText,
      ...Array.from(document.querySelectorAll('.x-toolbar')).map(t => t.innerText),
    ].filter(Boolean);
    for (const txt of candidates) {
      const m = txt.match(/\/\s*(\d+)/) || txt.match(/of\s+(\d+)/i) || txt.match(/(\d+)\s*$/);
      if (m) return parseInt(m[1], 10);
    }
    return 0;
  });
  console.log(`  📊 Paging toolbar reports total: ${total || "unknown"}`);

  const seen = new Set();
  const out = [];

  const scanCurrentPage = async () => {
    const batch = await actualFrame.evaluate(() => {
      // Resolve the Template No. column id from the header row.
      let colId = null;
      for (const h of document.querySelectorAll('.x-grid3-hd')) {
        const txt = (h.querySelector('.x-grid3-hd-inner')?.innerText || h.innerText || "").trim();
        if (txt === 'Template No.' || txt === 'Template No') {
          const m = h.className.match(/x-grid3-hd-([A-Za-z0-9_-]+)/);
          if (m && m[1] !== 'inner' && m[1] !== 'row') { colId = m[1]; break; }
        }
      }

      const found = [];
      document.querySelectorAll('.x-grid3-row').forEach(row => {
        let txt = "";
        if (colId) {
          // Read ONLY the Template No. column. Locked-column duplicate rows
          // lack this cell → empty → skipped (avoids picking row-number/checkbox).
          const cell = row.querySelector('.x-grid3-td-' + colId + ' .x-grid3-cell-inner')
                    || row.querySelector('.x-grid3-td-' + colId);
          txt = (cell?.innerText || cell?.textContent || "").trim();
        } else {
          // Fallback: first hyperlink-looking cell in the row.
          const link = row.querySelector('a');
          txt = (link?.innerText || link?.textContent || "").trim();
        }
        if (txt) found.push(txt);
      });
      return found;
    });

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

  // Reads the first grid row's full text — used as a cheap "did the page
  // change?" fingerprint after clicking Next (the .x-paging-info element this
  // screen lacks, so we can't diff the toolbar caption).
  const firstRowText = () => actualFrame.evaluate(() => {
    const row = document.querySelector('.x-grid3-row');
    return row ? (row.innerText || row.textContent || "").trim() : "";
  });

  let safety = 50;
  let pageNum = 1;
  while (safety-- > 0) {
    const added = await scanCurrentPage();
    console.log(`    📄 Page ${pageNum}: +${added} new template(s) (running total ${out.length}${total ? `/${total}` : ""})`);

    if (total > 0 && out.length >= total) break;

    const firstBefore = await firstRowText();

    const clickedNext = await actualFrame.evaluate(() => {
      const nextBtn = document.querySelector('.x-tbar-page-next');
      if (!nextBtn) return false;
      if (nextBtn.classList.contains('x-item-disabled')) return false;
      nextBtn.click();
      return true;
    });
    if (!clickedNext) break;

    try {
      await actualFrame.waitForFunction(
        (prev) => {
          const row = document.querySelector('.x-grid3-row');
          const cur = row ? (row.innerText || row.textContent || "").trim() : "";
          return cur && cur !== prev;
        },
        firstBefore,
        { timeout: 8000 }
      );
    } catch (_) {
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
    console.warn(`  ⚠ Collected ${out.length} of ${total} template(s).`);
  }
  console.log(`  ✔ Collected ${out.length} template(s)${total ? ` (of ${total} total)` : ""}: ${out.join(", ")}`);
  return out;
}

/**
 * Scans every frame for a template number belonging to the currently-open F2
 * Direct Entry popup. Priority (high → low):
 *   1) document.activeElement — the F2 popup input is focused immediately
 *      after F2 is pressed.
 *   2) Inputs / text inside visible ExtJS dialogs (.x-window, .x-dlg, etc.).
 *   3) Any visible <input>/<textarea> on the page (last resort).
 */
async function extractTemplateFromPopup(page) {
  const patternSrc = TEMPLATE_NO_PATTERN.source.replace(/^\^|\$$/g, ''); // match anywhere

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
    } catch (_) {}
  }
  return null;
}

/**
 * Locates the Direct Entry "number" input by trying each candidate label in
 * DIRECT_ENTRY_LABELS. Returns the first visible matching Playwright locator,
 * or null.
 */
function directEntryInput(frame) {
  // Build a single XPath that matches the input following ANY candidate label.
  const xpath = DIRECT_ENTRY_LABELS
    .map(lbl => `//*[normalize-space(text())="${lbl}"]/following::input[1]`)
    .join(" | ");
  return frame.locator(`xpath=${xpath}`).first();
}

/**
 * Opens a template's detail screen. The main loop runs Excel-driven with
 * forceDirectEntry=true: it pastes the template number into the Direct Entry
 * field and clicks the "View Recurring Voucher Template" link — no grid click,
 * no F2, no ArrowDown. The F2 path is retained for the non-forced case.
 *
 * Returns { ok: boolean, templateNo: string|null }.
 */
async function openTemplateViaF2(page, rowIndex, knownTemplateNo = null, forceDirectEntry = false) {
  console.log(`\n🖱 Opening template via Direct Entry — row ${rowIndex + 1}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  const rowsInDom = await frame.locator('.x-grid3-row').count().catch(() => 0);
  const rowOnCurrentPage = !forceDirectEntry && rowIndex < rowsInDom;

  let templateNo;

  if (!rowOnCurrentPage) {
    if (!knownTemplateNo) {
      console.error(`  ❌ Row ${rowIndex + 1} not on current grid page and no known template number to fall back on.`);
      return { ok: false, templateNo: null };
    }
    if (forceDirectEntry) {
      console.log(`  ⏭ Using Direct Entry (forced) with known template ${knownTemplateNo}.`);
    } else {
      console.log(`  ⏭ Row ${rowIndex + 1} not on current grid page — bypassing F2 and using known template ${knownTemplateNo}.`);
    }
    templateNo = knownTemplateNo;
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
        } catch (_) {}
      }
      if (!highlighted) {
        console.warn(`  ⚠ Could not highlight row ${rowIndex + 1}.`);
        return { ok: false, templateNo: null };
      }
      await page.waitForTimeout(500);

      console.log(`  ⌨ Pressing Tab to move keyboard focus into the grid...`);
      await page.keyboard.press('Tab');
      await page.waitForTimeout(300);
    }

    // 2. Press F2 → popup opens with template number pre-filled
    console.log(`  ⌨ Pressing F2...`);
    await page.keyboard.press('F2');
    await page.waitForTimeout(1200);

    // 3. Copy template number from popup, then dismiss the popup
    templateNo = await extractTemplateFromPopup(page);
    if (!templateNo) {
      if (knownTemplateNo) {
        console.warn(`  ⚠ Popup read failed; falling back to known template ${knownTemplateNo}.`);
        templateNo = knownTemplateNo;
      } else {
        console.error(`  ❌ Could not read template number from F2 popup.`);
        await page.keyboard.press('Escape').catch(() => {});
        return { ok: false, templateNo: null };
      }
    }
    console.log(`  📋 Copied template number: ${templateNo}`);

    await page.keyboard.press('Escape');
    await page.waitForTimeout(500);
  }

  // 4. Paste into the Direct Entry "number" field.
  console.log(`  ⌨ Pasting into Direct Entry field...`);
  const templateInput = directEntryInput(frame);
  try {
    await templateInput.waitFor({ state: "visible", timeout: 5000 });
    await templateInput.click();
    await templateInput.fill("");
    await templateInput.fill(templateNo);
    console.log(`  ✔ Pasted ${templateNo}.`);
  } catch (e) {
    console.error(`  ❌ Failed to paste into Direct Entry: ${e.message}`);
    return { ok: false, templateNo };
  }

  // 5. Trigger navigation — click the "View Recurring Voucher Template" link.
  console.log(`  🖱 Clicking "${VIEW_LINK_TEXT}"...`);
  let clicked = false;

  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  if (actualFrame) {
    try {
      clicked = await actualFrame.evaluate((linkText) => {
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
          if (directText !== linkText) continue;
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
      }, VIEW_LINK_TEXT);
      if (clicked) console.log(`  ✔ Clicked "${VIEW_LINK_TEXT}" via JS dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      const viewLink = frame.getByText(VIEW_LINK_TEXT, { exact: true }).first();
      await viewLink.waitFor({ state: "visible", timeout: 5000 });
      await viewLink.click({ force: true });
      clicked = true;
      console.log(`  ✔ Clicked "${VIEW_LINK_TEXT}" via locator.`);
    } catch (e) {
      console.error(`  ❌ Failed to click ${VIEW_LINK_TEXT}: ${e.message}`);
      return { ok: false, templateNo };
    }
  }

  // Poll for one of two outcomes after the View link click:
  //   (a) detail page attaches → break and continue normally
  //   (b) "... does not exist" popup (#message-panel ExtJS window) appears →
  //        close it, retry once, then skip if it still fails.
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
    for (const txt of DETAIL_READY_TEXTS) {
      const visible = await frame.locator(`text="${txt}"`).first()
        .isVisible({ timeout: 0 })
        .catch(() => false);
      if (visible) return true;
    }
    return false;
  };

  const dismissPopup = async (panel) => {
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
      console.warn(`  ⚠ "does not exist" popup detected (attempt ${attempt}/2). Closing...`);
      await dismissPopup(popup);

      if (attempt === 2) {
        console.error(`  ❌ Template ${templateNo} does not exist after retry — skipping.`);
        return { ok: false, templateNo };
      }

      console.log(`  🔄 Retrying: re-paste + click "${VIEW_LINK_TEXT}"...`);
      try {
        await templateInput.click();
        await templateInput.fill("");
        await templateInput.fill(templateNo);
        await frame.getByText(VIEW_LINK_TEXT, { exact: true }).first().click({ force: true });
      } catch (e) {
        console.error(`  ❌ Retry failed: ${e.message}`);
        return { ok: false, templateNo };
      }
      continue;
    }

    console.log(`  ⏳ No popup or detail within 10s on attempt ${attempt}, falling through...`);
    break;
  }

  // Final readiness wait (no-op if detail already attached above).
  if (!detailLoaded) {
    let waited = false;
    for (const txt of DETAIL_READY_TEXTS) {
      const ok = await frame.locator(`text="${txt}"`).first()
        .waitFor({ state: "attached", timeout: 3000 })
        .then(() => true)
        .catch(() => false);
      if (ok) { waited = true; break; }
    }
    if (!waited) console.log("  ⏳ Detail screen still loading (no known header matched)...");
  }
  await page.waitForTimeout(300);
  return { ok: true, templateNo };
}

/**
 * Counts data rows currently rendered in the grid (DOM-based).
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
 * Extracts the header fields of the View Recurring Voucher Template detail
 * screen via spatial-proximity scraping and writes the result to
 * `<exportDir>/Header Info - <templateNo>.txt` as pretty-printed JSON.
 *
 * NOTE: the label→key schema below is best-effort. Confirm the real field
 * captions on the detail screen and adjust the left-hand label strings — any
 * label that doesn't match simply yields null (harmless).
 *
 * Must run BEFORE downloadTemplateDetail's Tab + Shift+E sequence, since that
 * keyboard sequence moves focus and can shift the DOM layout.
 */
async function saveTemplateInfoJson(page, templateNo, exportDir) {
  console.log("  📋 Extracting Recurring Voucher Template header → JSON...");
  try {
    const actualFrame = page.frames().find(
      (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
    );
    if (!actualFrame) {
      console.warn("    ⚠ Could not locate ilboinnerframe — skipping JSON dump.");
      return;
    }

    // Labels match the View Recurring Voucher Template detail screen. All of
    // these sit ABOVE the "Accounting Information" grid, so the formMaxY ceiling
    // below resolves them correctly. (Data History fields sit below the grid and
    // are intentionally not scraped here.)
    const schema = {
      recurring_voucher_information: [
        ["Template No.",                          "template_no"],
        ["Status",                                "status"],
        ["Effective Date From",                   "effective_date_from"],
        ["Effective Date To",                     "effective_date_to"],
        ["Recurring Voucher Status",              "recurring_voucher_status"],
        ["Finance Book",                          "finance_book"],
        ["Posting Frequency",                     "posting_frequency"],
        ["Numbering Type For Recurring Vouchers", "numbering_type_for_recurring_vouchers"],
        ["Control Total",                         "control_total"],
        ["Remarks",                               "remarks"],
      ],
      project_information: [
        ["Project OU",                        "project_ou"],
        ["Project / Sub-Project Code",        "project_sub_project_code"],
        ["Project / Sub-Project Description", "project_sub_project_description"],
        ["Job Number",                        "job_number"],
        ["AFE Number",                        "afe_number"],
      ],
    };

    // Field keys whose value should be serialized as a Number, not a string.
    const numericKeys = new Set(["control_total"]);

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

      // All form labels we care about live ABOVE the "Accounting Information"
      // grid. Compute a Y ceiling from the grid section header and reject any
      // label candidate at or below it (so grid column headers don't shadow
      // the form labels).
      let formMaxY = Infinity;
      for (const el of document.querySelectorAll("*")) {
        if (!isVisible(el)) continue;
        if (getDirectText(el) !== "Accounting Information") continue;
        const r = el.getBoundingClientRect();
        if (r.top < formMaxY) formMaxY = r.top;
      }

      const findLabelEl = (lbl) => {
        let best = null;
        let bestArea = Infinity;
        for (const el of document.querySelectorAll("*")) {
          if (!isVisible(el)) continue;
          if (getDirectText(el) !== lbl) continue;
          const r = el.getBoundingClientRect();
          if (r.top >= formMaxY) continue;
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

      const valueCandidates = Array.from(
        document.querySelectorAll(
          'input, textarea, select, .x-form-display-field, .x-form-field, div[class*="x-form"]'
        )
      ).filter((el) => {
        if (!isVisible(el) && el.tagName !== "INPUT") return false;
        const t = getDirectText(el);
        if (t && labels.includes(t)) return false;
        const r = el.getBoundingClientRect();
        if (r.top >= formMaxY) return false;
        return true;
      });

      const map = {};
      for (const lbl of labels) {
        const labelEl = findLabelEl(lbl);
        if (!labelEl) { map[lbl] = ""; continue; }
        const labelRect = labelEl.getBoundingClientRect();

        let value = "";

        // Strategy 1 (primary for Ramco): spatial proximity. The value
        // element is visually to the RIGHT of the label on the same band.
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

        if (value && labels.includes(value)) value = "";

        map[lbl] = value;
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
      for (const [label, key] of pairs) {
        structured[section][key] = normalize(valuesMap[label], key);
      }
    }
    // Guarantee the template number is populated from the caller-provided value.
    structured.recurring_voucher_information.template_no = templateNo;

    const outPath = `${exportDir}/Header Info - ${safeFileName(templateNo)}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save template info JSON: ${error.message}`);
  }
}

/**
 * Dumps the header JSON, then downloads the Accounting Information grid that is
 * embedded on the View Recurring Voucher Template detail screen. Focus is
 * seated on the grid (click its first row) so Shift+E exports the accounting
 * rows. Saved as "Accounting Info - (templateNo).xls".
 */
async function downloadTemplateDetail(page, templateNo, exportDir) {
  console.log(`\n📥 Preparing to download Accounting Information (Template: ${templateNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });
    mkdirSync(exportDir, { recursive: true });

    await page.waitForTimeout(500);

    // Dump JSON BEFORE the focus/Shift+E sequence, which shifts keyboard focus
    // and can reflow the DOM.
    await saveTemplateInfoJson(page, templateNo, exportDir);

    const frame = page.frameLocator("#ilboinnerframe");

    const MAX_ATTEMPTS = 2;
    let lastErr = null;
    for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
      try {
        if (attempt > 1) {
          console.log(`  🔄 Retrying Accounting Excel export (attempt ${attempt}/${MAX_ATTEMPTS})...`);
          await page.waitForTimeout(800);
        }

        // Seat focus on the Accounting Information grid (embedded on the detail
        // page) so Shift+E exports it. Click the first grid row; fall back to Tab.
        console.log("  🖱 Clicking first Accounting grid row to focus it...");
        try {
          await frame.locator('.x-grid3-row').first().click({ force: true, timeout: 5000 });
        } catch (e) {
          console.warn(`  ⚠ Grid-row click failed (${e.message}); falling back to 1 Tab.`);
          await page.keyboard.press("Tab");
        }
        await page.waitForTimeout(300);

        console.log("  ⌨ Triggering Excel Export (Shift + E)...");
        const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
        await page.keyboard.down('Shift');
        await page.keyboard.press('E');
        await page.keyboard.up('Shift');

        const download = await downloadPromise;
        const suggestedName = download.suggestedFilename();
        console.log(`  🔍 Server suggested name: "${suggestedName}"`);

        const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
        const fileName = `${exportDir}/Accounting Info - (${safeFileName(templateNo)}).${ext}`;

        await download.saveAs(fileName);
        console.log(`  ✅ Accounting Excel download successful!`);
        console.log(`  📁 Saved → ${fileName}`);
        return;
      } catch (e) {
        lastErr = e;
        console.warn(`  ⚠ Accounting Excel attempt ${attempt}/${MAX_ATTEMPTS} failed: ${e.message}`);
      }
    }
    console.error(`  ❌ Accounting Excel download failed after ${MAX_ATTEMPTS} attempts: ${lastErr ? lastErr.message : 'unknown'}`);

  } catch (error) {
    console.error(`  ❌ Accounting Excel download failed: ${error.message}`);
  }
}

/**
 * Clicks the "View Recurring Voucher Status" footer link on the detail screen
 * and waits for the Recurring Voucher Status page (its own grid) to load.
 * Returns true on success.
 */
async function goToRecurringVoucherStatus(page) {
  console.log("\n📑 Navigating to 'View Recurring Voucher Status'...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  const linkText = "View Recurring Voucher Status";

  let clicked = false;

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
          el.scrollIntoView({ block: 'center' });
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
      }, linkText);
      if (clicked) console.log(`  ✔ Clicked "${linkText}" via JS dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      const link = frame.getByText(linkText, { exact: true }).first();
      await link.waitFor({ state: "visible", timeout: 5000 });
      await link.scrollIntoViewIfNeeded().catch(() => {});
      await link.click({ force: true });
      clicked = true;
      console.log(`  ✔ Clicked "${linkText}" via locator.`);
    } catch (e) {
      console.error(`  ❌ Failed to click ${linkText}: ${e.message}`);
      return false;
    }
  }

  // Wait for the Status page grid to render (poll up to ~20s). Accept either a
  // populated grid row or a page title/breadcrumb mentioning "Recurring Voucher
  // Status".
  console.log("  ⏳ Waiting for Recurring Voucher Status page to load...");
  const deadline = Date.now() + 20000;
  let loaded = false;
  while (Date.now() < deadline) {
    let signal = null;
    for (const f of page.frames()) {
      try {
        signal = await f.evaluate(() => {
          const isVisible = (el) => {
            if (!el || el.offsetParent === null) return false;
            const s = window.getComputedStyle(el);
            return s.display !== 'none' && s.visibility !== 'hidden';
          };
          const rows = document.querySelectorAll('.x-grid3-row');
          for (const r of rows) { if (isVisible(r)) return 'grid'; }
          for (const el of document.querySelectorAll(
            '.x-panel-header-text, .breadcrumb, .x-toolbar, h1, h2, .page-title'
          )) {
            if (!isVisible(el)) continue;
            if (/Recurring Voucher Status/i.test((el.innerText || '').trim())) return 'breadcrumb';
          }
          return null;
        });
      } catch (_) {}
      if (signal) break;
    }
    if (signal === 'grid') { loaded = true; break; }
    await page.waitForTimeout(300);
  }
  if (loaded) {
    console.log("  ✔ Recurring Voucher Status grid detected.");
    await page.waitForTimeout(500);
  } else {
    console.warn("  ⚠ Status grid not detected within 20s; proceeding anyway.");
  }
  return true;
}

/**
 * Downloads the Excel from the Recurring Voucher Status page (focus the grid,
 * then Shift+E). Saved as "Recurring Voucher Status - (templateNo).xls".
 */
async function downloadRecurringVoucherStatusExcel(page, templateNo, exportDir) {
  console.log(`\n📊 Preparing to download Recurring Voucher Status Excel (Template: ${templateNo})...`);
  mkdirSync(exportDir, { recursive: true });

  const frame = page.frameLocator("#ilboinnerframe");
  const MAX_ATTEMPTS = 2;
  let lastErr = null;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      if (attempt > 1) {
        console.log(`  🔄 Retrying Status Excel export (attempt ${attempt}/${MAX_ATTEMPTS})...`);
        await page.waitForTimeout(800);
      }

      console.log("  🖱 Focusing the Status grid...");
      try {
        await frame.locator('.x-grid3-row').first().click({ force: true, timeout: 5000 });
      } catch (e) {
        console.warn(`  ⚠ Grid-row click failed (${e.message}); falling back to 1 Tab.`);
        await page.keyboard.press("Tab");
      }
      await page.waitForTimeout(300);

      console.log("  ⌨ Triggering Excel Export (Shift + E)...");
      const downloadPromise = page.waitForEvent('download', { timeout: 30000 });
      await page.keyboard.down('Shift');
      await page.keyboard.press('E');
      await page.keyboard.up('Shift');

      const download = await downloadPromise;
      const suggestedName = download.suggestedFilename();
      console.log(`  🔍 Server suggested name: "${suggestedName}"`);

      const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
      const fileName = `${exportDir}/Recurring Voucher Status - (${safeFileName(templateNo)}).${ext}`;

      await download.saveAs(fileName);
      console.log(`  ✅ Status Excel download successful!`);
      console.log(`  📁 Saved → ${fileName}`);
      return;
    } catch (e) {
      lastErr = e;
      console.warn(`  ⚠ Status Excel attempt ${attempt}/${MAX_ATTEMPTS} failed: ${e.message}`);
    }
  }
  console.error(`  ❌ Status Excel download failed after ${MAX_ATTEMPTS} attempts: ${lastErr ? lastErr.message : 'unknown'}`);
}

/**
 * Clicks the "Go back" button to return to the search results grid
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
 * Downloads the initial search-results grid (Shift+E). Filename anchors on the
 * Effective Date From field when present, otherwise on today's date.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    let rawDate = "";
    const dateField = frame.locator('xpath=//label[contains(text(), "Effective Date From")]/following::input[1] | //td[contains(text(), "Effective Date From")]/following::input[1]').first();
    if (await dateField.isVisible({ timeout: 3000 }).catch(() => false)) {
      rawDate = await dateField.inputValue().catch(() => "");
    }
    console.log(`  📅 Detected Effective Date From: ${rawDate || "(none)"}`);

    const formattedDate = rawDate ? formatRamcoDate(rawDate) : new Date().toISOString().split('T')[0];
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Recurring Voucher Templates/${OU_FOLDER}/${dateFolder}`;
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
 * Read every Template No. from the initial search Excel export by locating the
 * "Template No." column in the header row and reading every value below it.
 * Ramco's "Shift+E" export is a ZIP-packaged .xlsx (despite the .xls
 * extension). Reading by COLUMN (rather than a regex) is what makes this work
 * for the free-form template codes (e.g. "A2Z/CHD-32", "ACCRL_DNTN_NIBT'24",
 * "AFME2026") — there is no fixed number pattern to match.
 *
 * Returns an array of unique template numbers in sheet (display) order.
 */
function readTemplatesFromExcel(xlsPath) {
  console.log(`\n📖 Reading template list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read template list.`);
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
        let templateColIdx = -1;
        for (let i = 0; i < rows.length && headerRowIdx < 0; i++) {
          for (let c = 0; c < rows[i].length; c++) {
            const v = (rows[i][c] || '').trim();
            if (/^Template\s*No\.?$/i.test(v)) {
              headerRowIdx = i;
              templateColIdx = c;
              break;
            }
          }
        }

        if (headerRowIdx < 0) continue;
        foundHeader = true;

        for (let i = headerRowIdx + 1; i < rows.length; i++) {
          const val = (rows[i][templateColIdx] || '').trim();
          if (val && !seen.has(val)) {
            seen.add(val);
            out.push(val);
          }
        }
        if (out.length > 0) break;
      }
    }

    if (!foundHeader) {
      console.warn(`  ⚠ No "Template No." header found in any sheet — cannot extract list.`);
    }

    console.log(`  ✔ Extracted ${out.length} template(s) from Excel.`);
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

      // Perform Search logic once we reach the View Recurring Voucher Template module
      if (menuPath.join(" ").includes(MODULE_LABEL)) {
        const stats = await performTemplateSearch(page);

        // Download the initial Excel export — it contains *every* template in
        // the result set, independent of grid pagination / buffered view.
        const initialXlsPath = await downloadInitialSearchResults(page);

        // Read the authoritative template list from that Excel by locating the
        // "Template No." column (template numbers are free-form codes, so we
        // read the column by header — same approach as index_sundry_receipt_full.js).
        const templateNumbers = readTemplatesFromExcel(initialXlsPath);
        const totalRows = templateNumbers.length;

        if (totalRows === 0) {
          console.error("  ❌ No templates extracted from Excel — cannot process.");
        } else {
          if (stats.total && stats.total !== totalRows) {
            console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
          }

          // Honor RESUME_FROM_RECURRING_VOUCHER: locate the template in the
          // Excel list and start processing from that row.
          let startIndex = 1;
          if (CONFIG.resumeFromTemplate) {
            const target = CONFIG.resumeFromTemplate.toUpperCase();
            const matchIdx = templateNumbers.findIndex(v => String(v).trim().toUpperCase() === target);
            if (matchIdx === -1) {
              console.error(`  ❌ RESUME_FROM_RECURRING_VOUCHER="${CONFIG.resumeFromTemplate}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
              break;
            }
            startIndex = matchIdx + 1;
            console.log(`  ⏭ RESUME_FROM_RECURRING_VOUCHER="${CONFIG.resumeFromTemplate}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
          }

          console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} template(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

          for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
            const currentTemplateNo = templateNumbers[processedCount - 1];
            process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${currentTemplateNo}`);

            // forceDirectEntry=true: paste the template into the Direct Entry
            // field and click the View link. No grid click, no F2, no ArrowDown.
            const result = await openTemplateViaF2(page, processedCount - 1, currentTemplateNo, true);
            console.log(`\n📂 --- PROCESSING TEMPLATE ${currentTemplateNo} (${processedCount} OF ${totalRows}) ---`);

            if (!result.ok) {
              console.error(`  ❌ Could not open template ${currentTemplateNo} — skipping.`);
              continue;
            }

            const dateFolder = new Date().toISOString().split('T')[0];
            const exportDir = `output/Recurring Voucher Templates/${OU_FOLDER}/${dateFolder}/${safeFileName(currentTemplateNo)}`;
            mkdirSync(exportDir, { recursive: true });

            // 1) Header JSON + Accounting Information grid Excel (on the detail page).
            await downloadTemplateDetail(page, currentTemplateNo, exportDir);

            // 2) Drill into "View Recurring Voucher Status" and download its grid,
            //    then return to the detail screen.
            const statusOk = await goToRecurringVoucherStatus(page);
            if (statusOk) {
              await downloadRecurringVoucherStatusExcel(page, currentTemplateNo, exportDir);
              console.log("  🔙 Returning from Recurring Voucher Status to detail...");
              await goBackFromDetail(page);
            }

            // 3) Back to the search results for the next row.
            console.log("  🔙 Navigating back to search results...");
            await goBackFromDetail(page);

            console.log(`  ✅ Row ${processedCount} completed.`);
            await page.waitForTimeout(500);
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

    writeFileSync("output/extracted-data-recurring-voucher.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-recurring-voucher.json");

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
