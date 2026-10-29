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
// output/Customer Debit Credit Notes/. Unknown OUs fall back to a sanitized
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
  ouInstance: (process.env.OU_INSTANCE_CUSTOMER_DEBIT_CREDIT || "").trim(),
  headless: false,
  slowMo:   50,
  timeout:  30000,

  // Select Debit / Credit Note search filters (dd/mm/yyyy — same format shown
  // in the Ramco UI).
  noteDateFrom: process.env.NOTE_DATE_FROM_CUSTOMER_DEBIT_CREDIT || "01/04/2026",
  noteDateTo:   process.env.NOTE_DATE_TO_CUSTOMER_DEBIT_CREDIT   || "23/04/2026",
  // Note Type & Note Category combos on the search screen (blank = leave as-is).
  noteType:     (process.env.NOTE_TYPE_CUSTOMER_DEBIT_CREDIT     || process.env.NOTE_TYPE     || "").trim(),
  noteCategory: (process.env.NOTE_CATEGORY_CUSTOMER_DEBIT_CREDIT || process.env.NOTE_CATEGORY || "").trim(),

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // note number and begins processing from there. Blank = start at row 1.
  resumeFromNote: (process.env.RESUME_FROM_CUSTOMER_DEBIT_CREDIT_NOTE || "").trim(),

  // How long to wait after clicking "Search" before deciding the grid is
  // genuinely empty. Ramco can show a "Please wait, Processing your
  // request..." overlay for several minutes on busy days / large result
  // sets — default 10 minutes to cover the worst case. Tunable via .env.
  searchWaitMs: parseInt(process.env.SEARCH_WAIT_MS_CUSTOMER_DEBIT_CREDIT || "600000", 10),
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Customer Debit Credit Notes/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Customer Debit Credit Notes/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Customer Debit/Credit Note workflow:
// Receivable Management → Customer Debit Credit Note → View Debit / Credit Note
const MENU_PATHS = [
  ["Receivable Management", "Customer Debit Credit Note", "View Debit / Credit Note"],
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
 * Orchestrates the Search workflow for the "Select Debit / Credit Note" screen.
 * Filters by Note Date From/To (Note Type / Note Category stay on "All" so
 * every Debit / Credit, Account Based / Item Based row comes back in a
 * single pass), then clicks Search.
 */
async function performNoteSearch(page) {
  console.log("\n🔍 Preparing to search debit/credit notes inside frame...");
  const frame = await getFrame(page);

  if (!frame) {
    throw new Error("Could not find ilboinnerframe");
  }

  try {
    // ── Helper: select a value in a combo box by label text ──────────────────────
    const selectComboByLabel = async (labelText, value) => {
      const frameLocator = page.frameLocator("#ilboinnerframe");

      const field = frameLocator
        .locator(
          `xpath=//label[contains(text(), "${labelText}")]/following::input[1] | ` +
          `//td[contains(text(), "${labelText}")]/following::input[1] | ` +
          `//span[contains(text(), "${labelText}")]/following::input[1] | ` +
          `//*[normalize-space(text())="${labelText}"]/following::input[1]`
        )
        .first();

      if (!(await field.isVisible({ timeout: 3000 }).catch(() => false))) {
        console.warn(`  ⚠️ "${labelText}" field not found in search criteria — skipping.`);
        return false;
      }

      await field.click();
      await page.waitForTimeout(500);

      const findAndClickExact = async () => {
        const items = frameLocator.locator(".x-combo-list-item");
        const count = await items.count().catch(() => 0);
        for (let i = 0; i < count; i++) {
          const text = (await items.nth(i).innerText()).trim();
          if (text.toLowerCase() === String(value).trim().toLowerCase()) {
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
        console.log(`    ℹ Combo item "${value}" not found in dropdown; typing directly...`);
        await field.focus();
        await page.keyboard.press("Control+A");
        await page.keyboard.press("Backspace");
        await page.keyboard.type(String(value), { delay: 50 });
        await page.keyboard.press("Enter");
      }

      await page.waitForTimeout(500);
      return true;
    };

    // Set Note Type search criteria dropdown if configured in .env
    if (CONFIG.noteType) {
      console.log(`  📑 Selecting Search Criteria Note Type: ${CONFIG.noteType}`);
      await selectComboByLabel("Note Type", CONFIG.noteType);
      console.log('  ✅ Note Type set.');
    }

    // Set Note Category search criteria dropdown if configured in .env
    if (CONFIG.noteCategory) {
      console.log(`  📑 Selecting Search Criteria Note Category: ${CONFIG.noteCategory}`);
      await selectComboByLabel("Note Category", CONFIG.noteCategory);
      console.log('  ✅ Note Category set.');
    }

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
        await frame.locator('input').first().focus().catch(() => {});
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

    console.log(`  📅 Filling Note Date: ${CONFIG.noteDateFrom} → ${CONFIG.noteDateTo}`);
    await fillDateField("Note Date From", 1, CONFIG.noteDateFrom);
    await fillDateField("Note Date From", 2, CONFIG.noteDateTo);
    console.log('  ✅ Note Date range set.');

    // 3. Click the Search button
    console.log('  🖱 Clicking "Search" button...');
    const searchBtn = frame.locator('button, .x-btn-text').filter({ hasText: /^Search$/i }).last();

    if (await searchBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await searchBtn.click();
      const searchWaitSecs = Math.round(CONFIG.searchWaitMs / 1000);
      console.log(`  🚀 Search triggered. Waiting up to ${searchWaitSecs}s for results...`);

      // The reliable "search done" signal is Ramco's "Please wait,
      // Processing your request..." overlay disappearing. We must poll
      // ALL frames, not just the ilboinnerframe — Ramco mounts modals at
      // the top level, so checking only the inner frame misses the overlay
      // entirely and the script proceeds with stale grid state.
      const isProcessingShown = async () => {
        for (const f of page.frames()) {
          try {
            const found = await f.evaluate(() =>
              /Please\s+wait.*Processing\s+your\s+request/i.test(document.body?.innerText || "")
            );
            if (found) return true;
          } catch (_) { /* frame detached, ignore */ }
        }
        return false;
      };

      // Wait up to 5 s for the overlay to APPEAR (confirms search started).
      const appearDeadline = Date.now() + 5000;
      let overlaySeen = false;
      while (Date.now() < appearDeadline) {
        if (await isProcessingShown()) {
          overlaySeen = true;
          console.log("  ⏳ 'Please wait, Processing your request...' overlay detected — search in progress.");
          break;
        }
        await page.waitForTimeout(200);
      }

      if (overlaySeen) {
        // Now wait for the overlay to DISAPPEAR — the real "search done" signal.
        const disappearDeadline = Date.now() + CONFIG.searchWaitMs;
        while (Date.now() < disappearDeadline) {
          if (!(await isProcessingShown())) {
            console.log("  ✅ Processing overlay cleared.");
            break;
          }
          await page.waitForTimeout(500);
        }
        if (await isProcessingShown()) {
          console.log("  ⏳ Processing overlay still showing after timeout — proceeding anyway.");
        }
      } else {
        // Overlay never appeared — likely a quick/cached query. Give the
        // grid a short buffer to paint before reading.
        console.log("  ℹ No processing overlay seen — short-circuiting with 2s buffer.");
        await page.waitForTimeout(2000);
      }

      // Settle delay so the grid finishes painting before we read it.
      await page.waitForTimeout(800);

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
      // The page has multiple .x-paging-info elements (Direct Entry section,
      // Search Results section, …). The first one often still shows
      // "[No records to display]" from a prior render — picking it gives
      // a false 0 total. Scan all of them and use the one with a real
      // count like "1 - 3 / 3" or "Displaying 1 - 3 of 3".
      const pagingElements = Array.from(document.querySelectorAll('.x-paging-info'));
      let total = 0;
      let pagingText = "";
      for (const el of pagingElements) {
        const text = (el.innerText || "").trim();
        if (!text) continue;
        const slashMatch = text.match(/(\d+)\s*\/\s*(\d+)/);     // "1 - 3 / 3"
        const ofMatch    = text.match(/of\s+(\d+)/i);            // "Displaying 1 - 3 of 3"
        if (slashMatch) {
          total = parseInt(slashMatch[2], 10);
          pagingText = text;
          break;
        } else if (ofMatch) {
          total = parseInt(ofMatch[1], 10);
          pagingText = text;
          break;
        }
      }

      // Count only data rows that are actually rendered with content —
      // ExtJS sometimes emits empty/template .x-grid3-row elements that
      // inflate the count.
      const visibleRows = Array.from(document.querySelectorAll('.x-grid3-row'))
        .filter(r => r.offsetParent !== null && (r.innerText || '').trim().length > 0)
        .length;

      return { total, visibleRows, pagingText };
    });

    if (stats.pagingText) {
      console.log(`  📋 Paging info: "${stats.pagingText}"`);
    }

    // Fallback: if paging info didn't yield a total but data rows are
    // visible, trust the row count — the toolbar text may be stale.
    if (stats.total === 0 && stats.visibleRows > 0) {
      console.log(`  ℹ Paging info reports 0 but ${stats.visibleRows} rows visible — using visible count as total.`);
      stats.total = stats.visibleRows;
    }

    console.log(`  ✔ Grid detected: Total ${stats.total} rows, ${stats.visibleRows} on current page.`);
    return stats;

  } catch (error) {
    console.error(`  ❌ Search automation failed: ${error.message}`);
    return { total: 0, visibleRows: 0 };
  }
}


/**
 * Note-number text pattern. Customer Debit / Credit Notes use purely-numeric
 * IDs (sometimes zero-padded, sometimes short like "968"), occasionally with
 * a single trailing letter that hints at the category (e.g. "0235794A" for
 * Account Based, "000005731" with no suffix). Different OUs use different
 * widths — A2Z Logistics emits 3-digit IDs, AAIHO emits 7-9 digits — so
 * we accept any number of digits with an optional trailing uppercase letter.
 */
const VOUCHER_NO_PATTERN = /^\d+[A-Z]?$/;

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
    } catch (_) {}
  }
  return null;
}

/**
 * Opens a note's detail screen via Direct Entry (mirroring the customer
 * receipt SOP, but adapted to the Select Debit / Credit Note screen):
 *   1. Paste the known Note No. into the "Note No." Direct Entry field.
 *   2. Click "View Account Based Note" or "View Item Based Note" based on
 *      the row's Note Category from the Excel export.
 *
 * The F2-popup branch is preserved for parity with the source script but
 * is unused on the main path — `forceDirectEntry=true` is passed for every
 * row because the Excel list already provides authoritative Note Nos.
 *
 * Returns { ok: boolean, voucherNo: string|null }.
 */
async function openVoucherViaF2(page, rowIndex, knownVoucherNo = null, forceDirectEntry = false, noteCategory = "", noteType = "") {
  // Customer DCN has two View links — pick the one matching the row category.
  const viewLinkText = /item/i.test(noteCategory || "")
    ? "View Item Based Note"
    : "View Account Based Note";
  console.log(`\n🖱 Opening note via Direct Entry — row ${rowIndex + 1} (${noteCategory || "Account Based"}, ${noteType || "default type"})...`);
  console.log(`  🎯 Will click footer link: "${viewLinkText}"`);
  const frame = page.frameLocator("#ilboinnerframe");

  // Align Direct Entry "Note Type" with the row's actual Note Type before
  // pasting. Without this step, a Debit row pasted while the combobox is
  // on the default "Credit" produces "Note No. and Note Type does not match".
  if (forceDirectEntry && noteType) {
    await setDirectEntryNoteType(page, noteType);
  }

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
        } catch (_) {}
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
        await page.keyboard.press('Escape').catch(() => {});
        return { ok: false, voucherNo: null };
      }
    }
    console.log(`  📋 Copied voucher number: ${voucherNo}`);

    await page.keyboard.press('Escape');
    await page.waitForTimeout(500);
  }

  // 4. Paste into "Note No." Direct Entry field (first match — excludes
  //    the "Note No. From" search-criteria field thanks to the exact-text anchor).
  console.log(`  ⌨ Pasting into Direct Entry "Note No." field...`);
  const voucherInput = frame
    .locator('xpath=//*[normalize-space(text())="Note No."]/following::input[1]')
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

  // 5. Trigger the search — click the matching "View … Note" footer link.
  //    Strategy A (primary): JS-dispatch a real mouse click on the innermost
  //    visible element whose *direct* text node exactly matches.
  //    Strategy B (fallback): Playwright getByText with exact match.
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

  // Poll for one of two outcomes after the click:
  //   (a) detail page attaches → break and continue normally
  //   (b) "Note Number does not exist" popup (#message-panel ExtJS window)
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
    // The View Account / Item Based Note detail screen carries a "Note
    // Information" section header (and a "Customer Information" header).
    // Either one signals the detail page has attached.
    for (const t of ['Note Information', 'Customer Information']) {
      const visible = await frame.locator(`text="${t}"`).first()
        .isVisible({ timeout: 0 })
        .catch(() => false);
      if (visible) return true;
    }
    return false;
  };

  const dismissPopup = async (panel) => {
    // Primary: Close button inside the panel (HTML shows <button class="x-btn-text">Close</button>).
    const closeBtn = panel.locator('button').filter({ hasText: /^Close$/i }).first();
    await closeBtn.click({ force: true }).catch(() => {});
    await page.waitForTimeout(300);
    // Backup: top-right X icon (.x-tool-close), in case Close didn't register.
    if (await panel.isVisible({ timeout: 0 }).catch(() => false)) {
      await panel.locator('.x-tool-close').first().click({ force: true }).catch(() => {});
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
      console.warn(`  ⚠ "Note Number does not exist" popup detected (attempt ${attempt}/2). Closing...`);
      await dismissPopup(popup);

      if (attempt === 2) {
        console.error(`  ❌ Note ${voucherNo} does not exist after retry — skipping.`);
        return { ok: false, voucherNo };
      }

      console.log(`  🔄 Retrying: re-paste + click "${viewLinkText}"...`);
      try {
        await voucherInput.click();
        await voucherInput.fill("");
        await voucherInput.fill(voucherNo);
        await frame.getByText(viewLinkText, { exact: true }).first().click({ force: true });
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
    await frame.locator('text="Note Information"').first()
      .waitFor({ state: "attached", timeout: 8000 })
      .catch(() => console.log("  ⏳ Detail screen still loading..."));
  }
  await page.waitForTimeout(300);
  // Don't claim success if the detail page never actually appeared —
  // otherwise downloadVoucherDetail would dump a Header Info JSON of the
  // still-visible Select Debit / Credit Note screen.
  if (!detailLoaded && !(await isDetailVisible())) {
    console.error(`  ❌ Direct Entry path: detail screen never attached for ${voucherNo}.`);
    return { ok: false, voucherNo };
  }
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
      await link.scrollIntoViewIfNeeded().catch(() => {});
      await link.click({ force: true, timeout: 5000 });
      clicked = true;
      console.log(`  ✔ Clicked voucher ${voucherNo} via locator.`);
    } catch (e) {
      console.error(`  ❌ Locator click failed: ${e.message}`);
    }
  }

  if (!clicked) return false;

  // Readiness signal for the View … Note screen
  await frame.locator('text="Note Information"').first()
    .waitFor({ state: "attached", timeout: 15000 })
    .catch(() => console.log("  ⏳ Detail screen still loading..."));
  await page.waitForTimeout(1000);
  return true;
}

/**
 * One-time setup: switch the Direct Entry "Note Category" combobox from its
 * default ("Account Based") to "All" so Direct Entry lookups can resolve
 * notes across every category — not just Account Based.
 *
 * The Select Debit / Credit Note screen has TWO "Note Category" labels (one
 * in Direct Entry, one in Search Criteria). The Direct Entry one appears
 * first in DOM order, so `(//…[text()="Note Category"])[1]/following::input[1]`
 * targets it.
 */
async function setDirectEntryNoteCategoryToAll(page) {
  console.log('\n⚙ Setting Direct Entry Note Category to "All" (one-time setup)...');
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const field = frame
      .locator('xpath=(//*[normalize-space(text())="Note Category"])[1]/following::input[1]')
      .first();

    await field.waitFor({ state: "visible", timeout: 5000 });
    await field.click();
    await page.waitForTimeout(500);

    let clicked = false;
    const items = page.locator('.x-combo-list-item');
    const count = await items.count().catch(() => 0);
    for (let i = 0; i < count; i++) {
      const text = (await items.nth(i).innerText()).trim();
      if (text === "All") {
        await items.nth(i).click();
        clicked = true;
        break;
      }
    }

    if (!clicked) {
      console.log('  ℹ Combo item "All" not found in dropdown; typing it directly...');
      await field.focus();
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type("All", { delay: 50 });
      await page.keyboard.press("Enter");
    }

    await page.waitForTimeout(500);
    console.log('  ✔ Note Category set to "All".');
    return true;
  } catch (error) {
    console.warn(`  ⚠ Could not set Note Category to All: ${error.message}`);
    return false;
  }
}

/**
 * Sets the Direct Entry "Note Type" combobox to match the row's actual
 * Note Type ("Debit" or "Credit"). Required before pasting a Note No. via
 * the Direct Entry fallback path because the combobox defaults to "Credit"
 * — and Ramco rejects mismatched type/number combinations with a
 * "Note No. and Note Type does not match" popup.
 *
 * The Direct Entry "Note Type" label is the FIRST "Note Type" label in
 * DOM order (Direct Entry section renders before Search Criteria), so
 * `(//*[text()="Note Type"])[1]/following::input[1]` targets it.
 */
async function setDirectEntryNoteType(page, noteType) {
  if (!noteType) return false;
  console.log(`  ⚙ Setting Direct Entry Note Type to "${noteType}"...`);
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const field = frame
      .locator('xpath=(//*[normalize-space(text())="Note Type"])[1]/following::input[1]')
      .first();

    await field.waitFor({ state: "visible", timeout: 5000 });
    await field.click();
    await page.waitForTimeout(400);

    let clicked = false;
    const items = page.locator(".x-combo-list-item");
    const count = await items.count().catch(() => 0);
    for (let i = 0; i < count; i++) {
      const text = (await items.nth(i).innerText()).trim();
      if (text.toLowerCase() === noteType.toLowerCase()) {
        await items.nth(i).click();
        clicked = true;
        break;
      }
    }

    if (!clicked) {
      console.log(`  ℹ Combo item "${noteType}" not visible in dropdown; typing it...`);
      await field.focus();
      await page.keyboard.press("Control+A");
      await page.keyboard.press("Backspace");
      await page.keyboard.type(noteType, { delay: 50 });
      await page.keyboard.press("Enter");
    }

    await page.waitForTimeout(400);
    console.log(`  ✔ Note Type set to "${noteType}".`);
    return true;
  } catch (error) {
    console.warn(`  ⚠ Could not set Note Type to "${noteType}": ${error.message}`);
    return false;
  }
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
 * Clicks the "View Account Based Note" link on the Select Debit / Credit Note
 * screen (legacy fallback path; the main Excel-driven loop uses Direct Entry
 * with the row's actual Note Category instead). Extracts the Note No. from
 * the loaded detail screen.
 */
async function triggerViewVoucher(page) {
  console.log('  🖱 Clicking "View Account Based Note" link...');
  const frameLocator = page.frameLocator("#ilboinnerframe");
  const viewBtn = frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /^View Account Based Note$/i }).first();

  const tryClick = async (locator, name) => {
    try {
      if (await locator.count() > 0) {
        await locator.scrollIntoViewIfNeeded().catch(() => {});
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
    const looseBtn = frameLocator.getByText('View Account Based Note', { exact: false }).first();
    clicked = await tryClick(looseBtn, "Flexible Link");
  }
  if (!clicked) {
    await frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /View Account Based Note/i }).first().evaluate(el => el.click()).catch(e => console.error("    ⚠ JS Click failed: " + e.message));
  }

  // Wait for the detail screen to appear ("Note Information" header).
  await frameLocator.locator('text="Note Information"').first().waitFor({ state: "attached", timeout: 15000 }).catch(() => {
    console.log("  ⏳ Detail screen still loading...");
  });
  await page.waitForTimeout(1000);

  // Extract Note No. from the detail screen
  let voucherNo = "Unknown";
  try {
    const frame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

    if (frame) {
      voucherNo = await frame.evaluate(() => {
        const elements = Array.from(document.querySelectorAll('label, td, span, div, b'));
        const labelEl = elements.find(el => {
            const text = el.innerText.trim();
            return text === 'Note No.' || text === 'Note No';
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
            const match = el.innerText.match(/Note No\.\s*[:\-]?\s*([A-Z0-9\-_]+)/i);
            if (match && match[1]) return match[1];
        }

        return "Unknown";
      });
    } else {
      console.warn("    ⚠ Could not find frame 'ilboinnerframe' for extraction.");
    }
  } catch (err) {
    console.warn("    ⚠ Failed to extract Note No: " + err.message);
  }

  return voucherNo;
}

/**
 * Scrapes the header sections of the View Account / Item Based Note screen
 * (Note Information, Customer Information, Payment Information, Project
 * Information) and writes the result to `<exportDir>/Header Info - <noteNo>.txt`
 * as pretty-printed JSON.
 *
 * Must run BEFORE downloadVoucherDetail's Tab + Shift+E sequence, since
 * that keyboard sequence moves focus and can shift the DOM layout.
 */
async function saveVoucherInfoJson(page, voucherNo, exportDir) {
  console.log("  📋 Extracting Note / Customer / Payment / Project Information → JSON...");
  try {
    const actualFrame = page.frames().find(
      (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
    );
    if (!actualFrame) {
      console.warn("    ⚠ Could not locate ilboinnerframe — skipping JSON dump.");
      return;
    }

    // Schema mirrors the supplier-side "Header Info" JSON, but the
    // Customer Debit / Credit Note screen replaces the Supplier sections
    // with Customer ones. Labels that do not exist on a given note type
    // simply scrape as null and are safe to leave in the schema.
    const schema = {
      note_information: [
        ["Note No.",       "note_no"],
        ["Note Type",      "note_type"],
        ["Note Category",  "note_category"],
        ["Status",         "status"],
        ["Note Date",      "note_date"],
        ["Finance Book",   "finance_book"],
        ["Note Amount",    "note_amount"],
        ["Currency",       "currency"],
      ],
      customer_information: [
        ["Customer Registered At", "customer_registered_at"],
        ["Customer Code",          "customer_code"],
        ["Customer Name",          "customer_name"],
        ["Customer Note No.",      "customer_note_no"],
        ["Customer Note Amount",   "customer_note_amount"],
        ["Customer Note Date",     "customer_note_date"],
        ["Customer Account Code",  "customer_account_code"],
        ["Account Description",    "account_description"],
      ],
      payment_information: [
        ["Currency",                 "currency"],
        ["Exchange Rate",            "exchange_rate"],
        ["Payment Processing Point", "payment_processing_point"],
        ["Electronic Payment",       "electronic_payment"],
        ["Payment Method",           "payment_method"],
        ["Pay Mode",                 "pay_mode"],
        ["Payment Priority",         "payment_priority"],
        ["Pay Term",                 "pay_term"],
        ["Anchor Date",              "anchor_date"],
        ["Total Amount",             "total_amount"],
        ["Comments",                 "comments"],
        ["Auto Adjust",              "auto_adjust"],
      ],
      project_information: [
        ["Project OU",                        "project_ou"],
        ["Project / Sub-Project Code",        "project_sub_project_code"],
        ["Project / Sub-Project Description", "project_sub_project_description"],
      ],
    };

    // Field keys whose value should be serialized as a Number, not a string.
    const numericKeys = new Set(["note_amount", "customer_note_amount", "exchange_rate", "total_amount"]);

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
      // visible pixels come from a wrapper <div>.
      const valueCandidates = Array.from(
        document.querySelectorAll(
          'input, textarea, select, .x-form-display-field, .x-form-field, div[class*="x-form"]'
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
    // Guarantee the note number is populated from the caller-provided
    // value, so this field is never null even if the scrape missed it.
    structured.note_information.note_no = voucherNo;

    const outPath = `${exportDir}/Header Info - ${voucherNo}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save note info JSON: ${error.message}`);
  }
}

/**
 * Returns true if the current page (View … Note or Accounting Information
 * sub-screen) has at least one data row in its main grid. False when the
 * grid shows an empty-state marker like "Found no rows to display!!!" or
 * "[No records to display]" in the paging toolbar.
 *
 * Some notes have no line items; some have no Accounting lines. Without
 * this check, Tab+Shift+E would either export an empty file or hang
 * waiting for a download event that never fires.
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
 * on it before triggering Shift+E. Required on screens that mix form fields
 * with a grid (e.g. View Tax Calculation - Details). On those screens,
 * Tab-based focus navigation lands on a form input (Ref. Doc. Line No.)
 * instead of the grid, so Shift+E types the literal "E" into the dropdown
 * rather than exporting the grid.
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
 * Clicks a footer hyperlink on the View … Note screen by exact text match
 * (e.g. "Accounting Information", "View Tax Details", "VAT", "Collaborate").
 * Opens whichever sub-screen the link points to.
 *
 * Strategy A (primary): JS-dispatch a real mouse click on the smallest
 * visible element whose direct text node exactly matches the label —
 * that's the leaf hyperlink, not a wrapping section/legend.
 * Strategy B (fallback): Playwright `.getByText(..., { exact: true }).last()`
 * — `.last()` because the footer hyperlink renders last in DOM order, after
 * any matching collapsed section header.
 *
 * Returns true on click success, false otherwise. Caller should `await
 * page.waitForTimeout(...)` for the sub-screen to attach.
 */
async function clickFooterLink(page, linkText) {
  console.log(`  🖱 Clicking "${linkText}" hyperlink...`);
  const frameLocator = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(
    (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
  );

  // VirtualWorks™ message popup ("TCAL Summary information not available",
  // "… not available", etc.) surfaces from clicks like
  // "View Tax Calculation - Summary" when Ramco has no data behind the link.
  // The sub-screen never attaches, so we need to dismiss the popup and either
  // retry once or skip — otherwise downstream Tab/Shift+E sequences hang.
  const findInfoPopup = async () => {
    for (const root of [page, frameLocator]) {
      const panel = root.locator('#message-panel').first();
      const visible = await panel.isVisible({ timeout: 0 }).catch(() => false);
      if (!visible) continue;
      const text = await panel.innerText().catch(() => "");
      if (/not available|does not exist|no data/i.test(text)) return { panel, text };
    }
    return null;
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

  const doClick = async () => {
    let clicked = false;
    if (actualFrame) {
      try {
        clicked = await actualFrame.evaluate((targetText) => {
          const isVisible = (el) => {
            if (!el || el.offsetParent === null) return false;
            const s = window.getComputedStyle(el);
            return s.display !== "none" && s.visibility !== "hidden";
          };
          let best = null;
          let bestArea = Infinity;
          for (const el of document.querySelectorAll("*")) {
            if (!isVisible(el)) continue;
            const directText = Array.from(el.childNodes)
              .filter((n) => n.nodeType === Node.TEXT_NODE)
              .map((n) => n.textContent.trim())
              .join("");
            if (directText !== targetText) continue;
            const r = el.getBoundingClientRect();
            const area = r.width * r.height;
            if (area > 0 && area < bestArea) {
              best = el;
              bestArea = area;
            }
          }
          if (!best) return false;
          best.scrollIntoView({ block: "center" });
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
        }, linkText);
        if (clicked) console.log(`  ✔ Clicked "${linkText}" via JS dispatch.`);
      } catch (e) {
        console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
      }
    }

    if (!clicked) {
      try {
        const link = frameLocator.getByText(linkText, { exact: true }).last();
        await link.waitFor({ state: "visible", timeout: 5000 });
        await link.scrollIntoViewIfNeeded().catch(() => {});
        await link.click({ force: true });
        clicked = true;
        console.log(`  ✔ Clicked "${linkText}" via locator fallback.`);
      } catch (e) {
        console.error(`  ❌ Failed to click "${linkText}": ${e.message}`);
      }
    }
    return clicked;
  };

  for (let attempt = 1; attempt <= 2; attempt++) {
    if (attempt > 1) {
      console.log(`  🔄 Retrying "${linkText}" click (attempt ${attempt}/2)...`);
    }
    const clicked = await doClick();
    if (!clicked) return false;

    // Poll up to 3 seconds for the "not available" popup — Ramco round-trips
    // the click to the server before deciding whether to render the popup,
    // so timing varies. If no popup appears, give the sub-screen another
    // ~1s to settle so keyboard focus lands before any later Tab/Shift+E.
    let popup = null;
    const deadline = Date.now() + 3000;
    while (Date.now() < deadline) {
      popup = await findInfoPopup();
      if (popup) break;
      await page.waitForTimeout(250);
    }

    if (!popup) {
      await page.waitForTimeout(1000);
      return true;
    }

    const popupText = popup.text.replace(/\s+/g, ' ').trim();
    console.warn(`  ⚠ VirtualWorks popup after "${linkText}" (attempt ${attempt}/2): ${popupText}`);
    await dismissPopup(popup.panel);

    if (attempt === 2) {
      console.error(`  ❌ "${linkText}" still shows popup after retry — skipping.`);
      return false;
    }
  }
  return false;
}

// Backwards-compatible alias — the existing call sites still work.
async function clickAccountingInformation(page) {
  return clickFooterLink(page, "Accounting Information");
}

/**
 * Tab×N + Shift+E export of the currently-visible grid. Saves to
 * `<exportDir>/<label> - (<voucherNo>).<ext>`. Caller is responsible for
 * placing the right grid on screen before this is invoked.
 *
 * Skips silently (and returns false) if the grid has no rows. Retries the
 * Tab + Shift+E + download sequence up to MAX_ATTEMPTS — transient
 * failures (focus drift, slow server, network blip) usually recover on a
 * second attempt.
 */
async function exportCurrentGrid(page, label, voucherNo, exportDir, tabCount = 1) {
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
      const fileName = `${exportDir}/${label} - (${voucherNo}).${ext}`;

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
 * Click a footer hyperlink on the current screen → wait for the
 * sub-screen to attach → export its grid via Tab×tabCount + Shift+E →
 * navigate back to the *previous* screen.
 *
 * `backWaitFor` selects the readiness target for the back navigation:
 *   - "note-info" (default): wait for the View … Note "Note Information"
 *     header — used when the parent screen IS View … Note.
 *   - "grid": wait for any `.x-grid3-row` — used when the parent screen
 *     is itself a sub-screen (e.g. "View Tax Calculation - Summary"),
 *     because the Summary renders a grid but no "Note Information" text.
 */
async function downloadFooterDrillDown(page, linkText, fileLabel, voucherNo, exportDir, tabCount = 1, backWaitFor = "note-info") {
  console.log(`\n📑 Drilling into "${linkText}" → exporting "${fileLabel}" Excel...`);
  const opened = await clickFooterLink(page, linkText);
  if (!opened) {
    console.error(`  ❌ Could not open "${linkText}" — skipping.`);
    return false;
  }

  const exported = await exportCurrentGrid(page, fileLabel, voucherNo, exportDir, tabCount);

  console.log(`  🔙 Returning from "${linkText}" → previous screen...`);
  await goBackFromDetail(page, { waitFor: backWaitFor });

  return exported;
}

/**
 * Enumerates all values in the "Ref. Doc. Line No." dropdown — pattern
 * ported from index_prepayment_full.js (`getRefDocLineOptions`).
 *
 * Strategy:
 *   1. Read the ExtJS combo's internal store directly via
 *      `Ext.getCmp(input.id).store` — this returns the *complete* list
 *      regardless of which records are virtualized in the DOM. Most
 *      reliable approach.
 *   2. Fallback: open the dropdown and scroll `.x-combo-list-inner`
 *      collecting items each tick — ExtJS combos virtualize the
 *      `.x-combo-list-item` children, so a one-shot read only sees the
 *      ~7 currently in view.
 */
async function getRefDocLineOptions(page) {
  console.log("  🔍 Detecting 'Ref. Doc. Line No.' options...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(
    (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
  );

  try {
    const field = frame
      .locator(
        'xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | ' +
          '//td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]'
      )
      .first();

    await field.waitFor({ state: "visible", timeout: 5000 });

    // Strategy 1: read the ExtJS combo store directly.
    let options = [];
    if (actualFrame) {
      try {
        options = await actualFrame.evaluate(() => {
          const labels = Array.from(document.querySelectorAll("label, td, span"));
          const labelEl = labels.find((el) => el.innerText.trim().includes("Ref. Doc. Line No."));
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
      console.log(`  ✔ Found options: ${options.length}`);
      console.log(`    Options: ${options.join(", ")}`);
      return options;
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

    let uniqueOptions = [...allOptions];

    // Sort numerically when every option is purely digits (1, 2, 3, ..., 10).
    const allNumeric = uniqueOptions.every((opt) => /^\d+$/.test(opt));
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
 * Selects a specific line in the "Ref. Doc. Line No." dropdown — pattern
 * ported from index_prepayment_full.js (`selectRefDocLine`).
 *
 * Strategy:
 *   1. Click the field to open the dropdown.
 *   2. Iterate `.x-combo-list-item` looking for an exact text match; click.
 *   3. Fallback: scroll the combo list 80px at a time, retrying the text
 *      match on each tick (handles virtualized lists).
 *   4. Fallback: clear the input and type the line number + Enter.
 */
async function selectRefDocLine(page, lineNo) {
  console.log(`  🔢 Selecting Ref. Doc. Line No: ${lineNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const field = frame
      .locator(
        'xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | ' +
          '//td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]'
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
        if (text.trim() === String(lineNo).trim()) {
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
 * Exports one Tax Details Excel per "Ref. Doc. Line No." option on the
 * View Tax Calculation - Details screen. Each line of the underlying
 * note has its own tax record — a single Shift+E only captures the line
 * currently selected, so we enumerate every option, select each in
 * turn, and export.
 *
 * Filenames:
 *   - Single line: `Tax Details - (<noteNo>).xls`
 *   - Multi-line:  `Tax Details Line <X> - (<noteNo>).xls` (one per line)
 */
async function downloadTaxDetailsPerLine(page, voucherNo, exportDir) {
  const options = await getRefDocLineOptions(page);

  if (options.length <= 1) {
    // Couldn't enumerate or only one line — single export of current view.
    // Anchor focus on the grid first: Tax Details has form fields above the
    // grid, so Tab-based focus lands on Ref. Doc. Line No. and Shift+E would
    // type "E" into the dropdown instead of exporting.
    await focusGridFirstRow(page);
    await exportCurrentGrid(page, "Tax Details", voucherNo, exportDir, 0);
    return;
  }

  for (let i = 0; i < options.length; i++) {
    const option = options[i];
    console.log(`\n  📍 Ref. Doc. Line No. "${option}" (${i + 1}/${options.length})...`);

    const selected = await selectRefDocLine(page, option);
    if (!selected) {
      console.warn(`  ⚠ Could not select line "${option}" — skipping export.`);
      continue;
    }

    // selectRefDocLine leaves focus on the dropdown — re-anchor on the grid.
    await focusGridFirstRow(page);
    const safeOption = String(option).replace(/[<>:"/\\|?*]/g, "_");
    await exportCurrentGrid(page, `Tax Details Line ${safeOption}`, voucherNo, exportDir, 0);
  }
}

/**
 * Enumerates every value in the "Transaction No." dropdown on the
 * Accounting Information screen — ported from `getRefDocLineOptions`.
 *
 * A single note can split into more than one accounting transaction
 * (e.g. "08RARDN26-000001", "08RARDN26-000002", …); each transaction has
 * its own Accounting Information grid. A single Shift+E only captures the
 * transaction currently selected, so we enumerate every option, select
 * each in turn, and export.
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
 * Accounting Information screen — ported from `selectRefDocLine`. After the
 * selection Ramco reloads the Accounting Information grid for the chosen
 * transaction.
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
 * Opens the "Accounting Information" sub-screen, selects the "Transaction No."
 * that matches the Note No., exports that one grid, then navigates back.
 *
 * The "Transaction No." dropdown can list more than one transaction for a
 * single note — most notably a *reversed* note also lists its reversal
 * document (e.g. Note No. "807" shows both "807" and "08RARDN26-000001").
 * Ramco defaults the dropdown to the reversal document, so without an explicit
 * selection we'd export the reversal's accounting entries instead of the
 * note's own. We therefore drive the selection from the Note No.: pick the
 * Transaction No. option equal to the Note No. and export only that one; the
 * reversal / other transactions are intentionally skipped.
 *
 * Filename: `Accounting Info - (<noteNo>).xls`
 *
 * Mirrors `downloadFooterDrillDown` (open → export → back).
 */
async function downloadAccountingInfoPerTransaction(page, voucherNo, exportDir) {
  console.log('\n📑 Drilling into "Accounting Information" → exporting Transaction No. matching Note No....');

  const opened = await clickFooterLink(page, "Accounting Information");
  if (!opened) {
    console.error('  ❌ Could not open "Accounting Information" — skipping.');
    return false;
  }

  const options = await getTransactionNoOptions(page);

  // Find the Transaction No. option corresponding to this note. Match exactly
  // first, then fall back to an alphanumeric-normalized compare so cosmetic
  // differences (padding, punctuation, casing) don't cause a miss.
  const target = String(voucherNo).trim();
  const norm = (s) => String(s).trim().replace(/[^A-Za-z0-9]/g, "").toUpperCase();
  const matched =
    options.find((o) => String(o).trim() === target) ||
    options.find((o) => norm(o) === norm(target));

  if (matched) {
    // Always pin the Transaction No. to the Note No. When a note has a
    // reversal, Ramco lists the reversal document alongside the note in this
    // dropdown AND defaults the selection to the reversal — so we must select
    // the Note-No.-matching option explicitly to export the note's own
    // accounting entries (not the reversal's). Selecting an already-selected
    // value just reloads the same grid, so this is safe even for a note with a
    // single transaction.
    console.log(`  🎯 Setting Transaction No. = Note No. "${target}" → "${matched}".`);
    const selected = await selectTransactionNo(page, matched);
    if (!selected) {
      console.warn(`  ⚠ Could not select Transaction No. "${matched}" — exporting current view.`);
    }
    // selectTransactionNo leaves focus on the dropdown — re-anchor on the grid
    // (the Transaction No. dropdown sits above the grid, so Tab-based focus
    // would otherwise land on the dropdown and Shift+E would type "E" into it).
    await focusGridFirstRow(page);
    await exportCurrentGrid(page, "Accounting Info", voucherNo, exportDir, 0);
  } else {
    // No Transaction No. option equals the Note No. (e.g. the dropdown
    // couldn't be enumerated, or the note's posting doc isn't keyed by Note
    // No.) — export whatever is currently shown so data isn't lost, and warn
    // when there were options to choose from.
    if (options.length > 0) {
      console.warn(
        `  ⚠ No Transaction No. option matched Note No. "${target}" ` +
        `(options: ${options.join(", ")}). Exporting current view.`
      );
    }
    await focusGridFirstRow(page);
    await exportCurrentGrid(page, "Accounting Info", voucherNo, exportDir, 0);
  }

  console.log('  🔙 Returning from "Accounting Information" → View … Note...');
  await goBackFromDetail(page, { waitFor: "note-info" });
  return true;
}

/**
 * Downloads the full per-note bundle from the View Account / Item Based
 * Note detail screen:
 *
 *   1. Header Info JSON (Note / Customer / Payment / Project Information)
 *   2. In-page "Note Info" line-items grid (Tab×9 + Shift+E)
 *   3. "Accounting Information" footer drill-down (Tab×1 + Shift+E)
 *   4. View Tax Calculation - Summary → View Tax Calculation - Details
 *      (one Tax Details Excel per Ref. Doc. Line No. option)
 *
 * Each footer drill-down clicks the link, exports the sub-screen's grid,
 * then navigates back so the next drill-down starts from the same
 * baseline. The main loop's `goBackFromDetail` after this function
 * returns to Select Debit / Credit Note as expected.
 */
async function downloadVoucherDetail(page, voucherNo = "Unknown") {
  console.log(`\n📥 Preparing to download per-note Excels (Note: ${voucherNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });
    await page.waitForTimeout(500);

    const dateFolder = new Date().toISOString().split("T")[0];
    const exportDir = `output/Customer Debit Credit Notes/${OU_FOLDER}/${dateFolder}/${voucherNo}`;
    mkdirSync(exportDir, { recursive: true });

    // 1) Header JSON — must run BEFORE any Tab/Shift+E sequence reflows
    //    the DOM and shifts keyboard focus.
    await saveVoucherInfoJson(page, voucherNo, exportDir);

    // 2) In-page Note Info line-items grid. Tab×9 parks focus on this
    //    grid (vs. Tab×1 for the footer drill-down sub-screens).
    await exportCurrentGrid(page, "Note Info", voucherNo, exportDir, 9);

    // 3) Accounting Information drill-down (View … Note → Accounting → back).
    //    A note can split into several accounting transactions, each with its
    //    own grid behind the "Transaction No." dropdown — export one Excel
    //    per Transaction No. value.
    await downloadAccountingInfoPerTransaction(page, voucherNo, exportDir);

    // 4) View Tax Calculation - Summary is the gateway to View Tax
    //    Details. The Summary screen itself isn't exported — it's used
    //    only as a navigation hub. Navigation is two-deep:
    //    View … Note → Summary → View Tax Details → back × 2.
    const summaryOpened = await clickFooterLink(page, "View Tax Calculation - Summary");
    if (summaryOpened) {
      // 4a) Drill from Summary into "View Tax Calculation - Details" and
      //     export one Tax Details Excel per Ref. Doc. Line No. option.
      const detailsOpened = await clickFooterLink(page, "View Tax Calculation - Details");
      if (detailsOpened) {
        await downloadTaxDetailsPerLine(page, voucherNo, exportDir);
        console.log('  🔙 Returning from "View Tax Calculation - Details" → Summary...');
        await goBackFromDetail(page, { waitFor: "grid" });
      } else {
        console.error('  ❌ Could not open "View Tax Calculation - Details" — skipping per-line tax export.');
      }

      // 4b) Pop back from Summary → View … Note so the main loop's
      //     goBackFromDetail correctly lands on Select Debit / Credit Note.
      console.log('  🔙 Returning from "View Tax Calculation - Summary" → View … Note...');
      await goBackFromDetail(page, { waitFor: "note-info" });
    } else {
      console.error('  ❌ Could not open "View Tax Calculation - Summary" — skipping tax exports.');
    }

  } catch (error) {
    console.error(`  ❌ downloadVoucherDetail failed: ${error.message}`);
  }
}

/**
 * Clicks the "Go back" button. Caller picks the readiness target so we
 * don't waste time waiting for an element the destination screen doesn't
 * have:
 *   - 'grid'      → Select Debit / Credit Note search results (.x-grid3-row)
 *   - 'note-info' → View … Note header ("Note Information")
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
    await backBtn.scrollIntoViewIfNeeded().catch(() => {});

    try {
      await backBtn.click({ force: true, timeout: 5000 });
      console.log("  ✔ Clicked 'Go back' (forced).");
    } catch (e) {
      console.warn("  ⚠ Direct 'Go back' click failed. Trying flexible search or JS dispatch...");
      const altBtn = page.locator('button, a, .x-btn-text').filter({ hasText: /Go back/i }).first();
      await altBtn.evaluate(el => el.click()).catch(err => console.error("    ❌ JS Click failed: " + err.message));
    }

    const frame = page.frameLocator("#ilboinnerframe");
    const target = waitFor === 'note-info'
      ? frame.locator('text="Note Information"').first()
      : frame.locator('.x-grid3-row').first();
    await target.waitFor({ state: "visible", timeout }).catch(() => {});
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
 * Downloads the initial Select Debit / Credit Note search results grid.
 * Uses "Note Date From" as the filename anchor.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const dateField = frame.locator('xpath=//label[contains(text(), "Note Date From")]/following::input[1] | //td[contains(text(), "Note Date From")]/following::input[1]').first();
    const rawDate = await dateField.inputValue();
    console.log(`  📅 Detected Note Date From: ${rawDate}`);

    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Customer Debit Credit Notes/${OU_FOLDER}/${dateFolder}`;
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
    // produce a fresh Excel for the note-list read step below.
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
 * Read every Note No. + Note Category from the initial Select Debit / Credit
 * Note Excel export. Note Category drives which footer link the loop clicks
 * ("View Account Based Note" vs "View Item Based Note"), so we need both.
 *
 * Ramco's "Shift+E" export is a ZIP-packaged .xlsm (despite the .xls
 * extension); we unzip with Windows' bundled bsdtar and parse the worksheet
 * XML directly using a column-scoped scan — naive regex over the workbook
 * picks up parenthesized account numbers in description cells that look
 * exactly like note numbers.
 *
 * Returns an array of `{ noteNo, noteCategory }` in sheet order.
 */
function readVouchersFromExcel(xlsPath) {
  console.log(`\n📖 Reading note list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read note list.`);
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

    // Walk every sheet and stop at the first that has a "Note No." header
    // with data rows beneath it. (Ramco numbers sheets non-sequentially —
    // template sheets sit alongside the data sheet.)
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

      // Find header row — first non-hidden row containing "Note No.".
      // While we're at it, capture the Note Category and Note Type column
      // indexes too. Note Type is needed because Direct Entry's Note Type
      // combobox defaults to "Credit"; pasting a Debit-row Note No. while
      // the combobox is on Credit produces "Note No. and Note Type does
      // not match", so the loop sets it to match the row before pasting.
      let noteNoCol = -1;
      let noteCatCol = -1;
      let noteTypeCol = -1;
      let headerRowIdx = -1;
      for (let i = 0; i < rows.length; i++) {
        if (rows[i].hidden) continue;
        for (let j = 0; j < rows[i].cells.length; j++) {
          const cell = rows[i].cells[j].trim().toLowerCase();
          if (noteNoCol === -1 && (cell === 'note no.' || cell === 'note no')) {
            noteNoCol = j;
            headerRowIdx = i;
          }
          if (noteCatCol === -1 && cell === 'note category') {
            noteCatCol = j;
          }
          if (noteTypeCol === -1 && cell === 'note type') {
            noteTypeCol = j;
          }
        }
        if (noteNoCol !== -1) break;
      }
      if (noteNoCol < 0) continue;

      const seen = new Set();
      const out = [];
      for (let i = headerRowIdx + 1; i < rows.length; i++) {
        if (rows[i].hidden) continue;
        const v = (rows[i].cells[noteNoCol] || '').trim();
        if (!v) continue;
        if (/^0+$/.test(v)) continue; // skip all-zero placeholders
        // Trust the column — Note No. values can be mixed alphanumeric
        // (e.g. "08ARDN2"), so don't filter by shape. The empty / all-zero
        // / dedupe guards above are enough to reject obvious junk rows.
        if (seen.has(v)) continue;
        seen.add(v);
        const noteCategory = noteCatCol >= 0 ? (rows[i].cells[noteCatCol] || '').trim() : '';
        const noteType = noteTypeCol >= 0 ? (rows[i].cells[noteTypeCol] || '').trim() : '';
        out.push({ noteNo: v, noteCategory, noteType });
      }

      if (out.length > 0) {
        console.log(
          `  📍 "Note No." in ${file} at column index ${noteNoCol}` +
            (noteCatCol >= 0 ? `, "Note Category" at ${noteCatCol}` : ' (no Note Category column)') +
            (noteTypeCol >= 0 ? `, "Note Type" at ${noteTypeCol}` : ' (no Note Type column)')
        );
        console.log(`  ✔ Extracted ${out.length} note(s).`);
        return out;
      }
    }

    console.warn('  ⚠ Could not find "Note No." column in any sheet.');
    return [];
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

      // Perform Search once we reach Customer Debit Credit Note → View Debit / Credit Note
      if (menuPath.join(" ").includes("View Debit / Credit Note")) {
        const stats = await performNoteSearch(page);

        // Empty result set — pressing Shift+E now would hang for 60 s
        // waiting on a download event Ramco never fires. Bail out early.
        if (!stats.total) {
          console.warn("  ⚠ Search returned 0 rows — skipping Excel export and note processing.");
          allResults.push(result);
          continue;
        }

        // Download the initial Excel export — it contains *every* note
        // in the result set, independent of grid pagination / buffered view.
        const initialXlsPath = await downloadInitialSearchResults(page);

        // Read the full note list (Note No. + Note Category) straight out of
        // that Excel file. This replaces any grid-scan walk: one authoritative
        // list, no keyboard navigation needed. If the Excel read fails we stop
        // — there is no safe fallback that handles paginated grids reliably.
        const notes = readVouchersFromExcel(initialXlsPath);
        const totalRows = notes.length;

        if (totalRows === 0) {
          console.error("  ❌ No notes extracted from Excel — cannot process.");
        } else {
          if (stats.total && stats.total !== totalRows) {
            console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
          }

          // Honor RESUME_FROM_CUSTOMER_DEBIT_CREDIT_NOTE: locate the note in
          // the Excel list and start processing from that row. Preserves the
          // original 1-based "Row N/Total" numbering so logs still line up
          // with the Excel index.
          let startIndex = 1;
          if (CONFIG.resumeFromNote) {
            const target = CONFIG.resumeFromNote.toUpperCase();
            const matchIdx = notes.findIndex(n => String(n.noteNo).trim().toUpperCase() === target);
            if (matchIdx === -1) {
              console.error(`  ❌ RESUME_FROM_CUSTOMER_DEBIT_CREDIT_NOTE="${CONFIG.resumeFromNote}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
              break;
            }
            startIndex = matchIdx + 1;
            console.log(`  ⏭ RESUME_FROM_CUSTOMER_DEBIT_CREDIT_NOTE="${CONFIG.resumeFromNote}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
          }

          console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} note(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

          // One-time Direct Entry setup: switch Note Category to "All" so
          // Direct Entry lookups resolve notes of either category. (Note
          // Type has no "All" option — only Debit / Credit — so we leave
          // whatever default Ramco shows; the View link choice is driven by
          // the row's category, not this dropdown.)
          await setDirectEntryNoteCategoryToAll(page);

          for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
            const { noteNo: currentNoteNo, noteCategory, noteType } = notes[processedCount - 1];
            process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${currentNoteNo} (${noteType || "?"} / ${noteCategory || "?"})`);

            // Direct Entry path: align the Note Type combobox with the
            // row's actual type, paste the Note No., then click the
            // matching "View … Note" footer link. The Excel-extracted
            // {noteNo, noteCategory, noteType} are authoritative.
            const result = await openVoucherViaF2(page, processedCount - 1, currentNoteNo, true, noteCategory, noteType);

            console.log(`\n📂 --- PROCESSING NOTE ${currentNoteNo} (${processedCount} OF ${totalRows}) ---`);

            if (!result.ok) {
              console.error(`  ❌ Could not open note ${currentNoteNo} — skipping.`);
              continue;
            }

            // 1) View … Note detail export (JSON header + Note Info Excel + Accounting Information Excel)
            await downloadVoucherDetail(page, currentNoteNo);

            // 2) Back to Select Debit / Credit Note for the next row.
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

    writeFileSync("output/extracted-data-customer-debit-credit.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-customer-debit-credit.json");

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
