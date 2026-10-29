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
// output/Purchase Requests/. Unknown OUs fall back to a sanitized
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
  ouInstance: (process.env.OU_INSTANCE_PURCHASE_REQUEST || "").trim(),
  headless: false,
  slowMo:   50,
  timeout:  30000,

  // View Purchase Request search filters (dd/mm/yyyy — same format shown in the Ramco UI)
  prDateFrom: process.env.PR_DATE_FROM_PURCHASE_REQUEST || "01/04/2026",
  prDateTo:   process.env.PR_DATE_TO_PURCHASE_REQUEST   || "03/04/2026",

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // PR number and begins processing from there. Blank = start at row 1.
  resumeFromPr: (process.env.RESUME_FROM_PURCHASE_REQUEST || "").trim(),
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Purchase Requests/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Purchase Requests/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Purchase Request workflow:
// Purchase and Subcontracting → Purchase Request → View Purchase Request
const MENU_PATHS = [
  ["Purchase and Subcontracting", "Purchase Request", "View Purchase Request"],
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
 * Orchestrates the Search workflow for the View Purchase Request screen.
 * (Filters by PR Date From/To, then clicks Search.)
 */
async function performPrSearch(page) {
  console.log("\n🔍 Preparing to search PRs inside frame...");
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
        await frame.locator('input').first().focus().catch(() => {});
        await page.keyboard.down('Control');
        await page.keyboard.press('a');
        await page.keyboard.up('Control');
        await page.keyboard.press('Backspace');
      }
    }

    // 2. Fill "PR Date From/To" filter
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

    console.log(`  📅 Filling PR Date: ${CONFIG.prDateFrom} → ${CONFIG.prDateTo}`);
    await fillDateField("PR Date From", 1, CONFIG.prDateFrom);
    await fillDateField("PR Date From", 2, CONFIG.prDateTo);
    console.log('  ✅ PR Date range set.');

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
 * PR-number text pattern, e.g. "02SUPAY26-000658".
 * Matches 2 digits + uppercase letters + 2 digits + "-" + digits.
 */
const PR_NO_PATTERN = /^\d{2}[A-Z]+\d{2}-\d+$/;

/**
 * Enumerates every PR No. currently rendered in the View Purchase Request
 * search-results grid. ExtJS grids render "hyperlinks" as styled <span>/<div>
 * elements, not <a> tags — so we scan grid-cell text and match the
 * PR-number pattern instead of relying on the anchor tag.
 *
 * Returns an array of unique PR numbers (e.g. ["02SUPAY26-000658", "02SUPAY26-000660"]).
 */
async function getPrNumbers(page) {
  console.log("\n🔢 Collecting PR numbers from the grid...");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  if (!actualFrame) {
    console.error("  ❌ Could not find ilboinnerframe.");
    return [];
  }

  // The View Purchase Request grid is a real paginated ExtJS grid
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
    }, PR_NO_PATTERN.source);

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
    console.log(`    📄 Page ${pageNum}: +${added} new PR(s) (running total ${out.length}${total ? `/${total}` : ""})`);

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
    console.warn(`  ⚠ Collected ${out.length} of ${total} PR(s).`);
  }
  console.log(`  ✔ Collected ${out.length} PR(s)${total ? ` (of ${total} total)` : ""}: ${out.join(", ")}`);
  return out;
}

/**
 * Scans every frame for a PR number belonging to the currently-open F2
 * Direct Entry popup. Priority (high → low):
 *   1) document.activeElement — the F2 popup input is focused immediately
 *      after F2 is pressed, so this is the most reliable source.
 *   2) Inputs / text inside visible ExtJS dialogs (.x-window, .x-dlg, etc.).
 *   3) Any visible <input>/<textarea> on the page (last resort — will match
 *      the Direct Entry "PR No." field even when stale).
 *
 * Without priorities 1 and 2, row 2+ would read the stale Direct Entry field
 * (still holding row 1's PR) and we'd open the same PR every time.
 */
async function extractPrFromPopup(page) {
  const patternSrc = PR_NO_PATTERN.source.replace(/^\^|\$$/g, ''); // match anywhere

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
 * Opens a PR's detail screen using the F2 Direct Entry shortcut, per the
 * user's annotated SOP:
 *   1. Press F2 (row already highlighted post-search) → popup shows PR number
 *   2. Copy PR number from popup → Escape the popup
 *   3. Paste into "PR No." Direct Entry field
 *   4. Click "View Purchase Request" to trigger navigation
 *
 * Row 1 is already highlighted after the search, so no row click is needed.
 * For rows 2+, we click the row at the given index to highlight it first.
 *
 * Returns { ok: boolean, prNo: string|null }.
 */
async function openPrViaF2(page, rowIndex, knownPrNo = null, forceDirectEntry = false) {
  console.log(`\n🖱 Opening PR via F2 — row ${rowIndex + 1}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  // If the target row isn't on the current grid page (e.g. row 6 when page 1
  // only holds rows 1-5) we can't click it or F2 it. Fall back to pasting
  // the already-collected PR number directly into Direct Entry.
  //
  // `forceDirectEntry=true` bypasses the grid-click + F2 path even when the
  // row index technically falls inside rowsInDom. Callers set this when the
  // PR number was already confirmed externally (e.g. the ArrowDown
  // continuation) and a second F2 would risk reading stale data — also
  // important because ExtJS grids with locked columns double the `.x-grid3-row`
  // count so the plain index comparison can't distinguish on-page vs off-page.
  const rowsInDom = await frame.locator('.x-grid3-row').count().catch(() => 0);
  const rowOnCurrentPage = !forceDirectEntry && rowIndex < rowsInDom;

  let prNo;

  if (!rowOnCurrentPage) {
    if (!knownPrNo) {
      console.error(`  ❌ Row ${rowIndex + 1} not on current grid page and no known PR number to fall back on.`);
      return { ok: false, prNo: null };
    }
    if (forceDirectEntry) {
      console.log(`  ⏭ Using Direct Entry (forced) with known PR ${knownPrNo}.`);
    } else {
      console.log(`  ⏭ Row ${rowIndex + 1} not on current grid page — bypassing F2 and using known PR ${knownPrNo}.`);
    }
    prNo = knownPrNo;
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
        return { ok: false, prNo: null };
      }
      await page.waitForTimeout(500);

      // Click alone doesn't move keyboard focus into the grid row on row 2+.
      // One Tab press shifts focus so F2 actually triggers the Direct Entry popup.
      console.log(`  ⌨ Pressing Tab to move keyboard focus into the grid...`);
      await page.keyboard.press('Tab');
      await page.waitForTimeout(300);
    }

    // 2. Press F2 → popup opens with PR number pre-filled
    console.log(`  ⌨ Pressing F2...`);
    await page.keyboard.press('F2');
    await page.waitForTimeout(1200);

    // 3. Copy PR number from popup, then dismiss the popup
    prNo = await extractPrFromPopup(page);
    if (!prNo) {
      // Popup extraction failed — but if we already know the PR from the
      // grid scan, we can still proceed by pasting it directly.
      if (knownPrNo) {
        console.warn(`  ⚠ Popup read failed; falling back to known PR ${knownPrNo}.`);
        prNo = knownPrNo;
      } else {
        console.error(`  ❌ Could not read PR number from F2 popup.`);
        await page.keyboard.press('Escape').catch(() => {});
        return { ok: false, prNo: null };
      }
    }
    console.log(`  📋 Copied PR number: ${prNo}`);

    await page.keyboard.press('Escape');
    await page.waitForTimeout(500);
  }

  // 4. Paste into "PR No." Direct Entry field (first match — excludes
  //    the "PR No. From" search-criteria field because of exact-text match).
  console.log(`  ⌨ Pasting into Direct Entry "PR No." field...`);
  const prInput = frame
    .locator('xpath=//*[normalize-space(text())="PR No."]/following::input[1]')
    .first();
  try {
    await prInput.waitFor({ state: "visible", timeout: 5000 });
    await prInput.click();
    await prInput.fill("");
    await prInput.fill(prNo);
    console.log(`  ✔ Pasted ${prNo}.`);
  } catch (e) {
    console.error(`  ❌ Failed to paste into Direct Entry: ${e.message}`);
    return { ok: false, prNo };
  }

  // 5. Trigger the search — click the "View Purchase Request" hyperlink.
  //    Strategy A (primary): JS-dispatch a real mouse click on the innermost
  //    visible element whose *direct* text node is exactly "View Purchase Request".
  //    Strategy B (fallback): Playwright getByText with exact match.
  console.log(`  🖱 Clicking "View Purchase Request"...`);
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
          // Only element whose *own* text nodes equal "View Purchase Request"
          const directText = Array.from(el.childNodes)
            .filter(n => n.nodeType === Node.TEXT_NODE)
            .map(n => n.textContent.trim())
            .join('');
          if (directText !== 'View Purchase Request') continue;
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
      if (clicked) console.log(`  ✔ Clicked "View Purchase Request" via JS dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
    }
  }

  if (!clicked) {
    try {
      const viewLink = frame.getByText('View Purchase Request', { exact: true }).first();
      await viewLink.waitFor({ state: "visible", timeout: 5000 });
      await viewLink.click({ force: true });
      clicked = true;
      console.log(`  ✔ Clicked "View Purchase Request" via locator.`);
    } catch (e) {
      console.error(`  ❌ Failed to click View Purchase Request: ${e.message}`);
      return { ok: false, prNo };
    }
  }

  // Poll for one of two outcomes after the View Purchase Request click:
  //   (a) detail page attaches → break and continue normally
  //   (b) "PR Number does not exist" popup (#message-panel ExtJS window)
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
      if (/PR Number does not exist/i.test(text)) return panel;
    }
    return null;
  };

  const isDetailVisible = async () => {
    return await frame.locator('text="PR Details"').first()
      .isVisible({ timeout: 0 })
      .catch(() => false);
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
      console.warn(`  ⚠ "PR Number does not exist" popup detected (attempt ${attempt}/2). Closing...`);
      await dismissPopup(popup);

      if (attempt === 2) {
        console.error(`  ❌ PR ${prNo} does not exist after retry — skipping.`);
        return { ok: false, prNo };
      }

      console.log(`  🔄 Retrying: re-paste + click "View Purchase Request"...`);
      try {
        await prInput.click();
        await prInput.fill("");
        await prInput.fill(prNo);
        await frame.getByText('View Purchase Request', { exact: true }).first().click({ force: true });
      } catch (e) {
        console.error(`  ❌ Retry failed: ${e.message}`);
        return { ok: false, prNo };
      }
      continue;
    }

    // Neither popup nor detail within 10s — fall through to the final wait.
    console.log(`  ⏳ No popup or detail within 10s on attempt ${attempt}, falling through...`);
    break;
  }

  // Final readiness signal (no-op if detail already attached above).
  if (!detailLoaded) {
    await frame.locator('text="PR Details"').first()
      .waitFor({ state: "attached", timeout: 15000 })
      .catch(() => console.log("  ⏳ Detail screen still loading..."));
  }
  await page.waitForTimeout(1000);
  return { ok: true, prNo };
}

/**
 * Counts data rows currently rendered in the View Purchase Request grid (DOM-based).
 * Fallback when getPrNumbers can't extract IDs — lets us still iterate
 * by row index and rely on F2 to surface the PR number.
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
 * now prefers openPrViaF2.
 */
async function clickPrByNumber(page, prNo) {
  console.log(`\n🖱 Clicking PR: ${prNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  let clicked = false;

  // Strategy 1 — JS click inside the grid-cell matching the PR text
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
      }, prNo);

      if (clicked) console.log(`  ✔ Clicked PR ${prNo} via grid-cell dispatch.`);
    } catch (e) {
      console.warn(`  ⚠ JS cell-click failed: ${e.message}`);
    }
  }

  // Strategy 2 — Playwright locator fallback (works if it's a real <a>)
  if (!clicked) {
    console.log("  ℹ Falling back to Playwright locator click...");
    const esc = prNo.replace(/[-/\\^$*+?.()|[\]{}]/g, '\\$&');
    const link = frame
      .locator('a, span, div')
      .filter({ hasText: new RegExp(`^\\s*${esc}\\s*$`) })
      .first();
    try {
      await link.waitFor({ state: "visible", timeout: 5000 });
      await link.scrollIntoViewIfNeeded().catch(() => {});
      await link.click({ force: true, timeout: 5000 });
      clicked = true;
      console.log(`  ✔ Clicked PR ${prNo} via locator.`);
    } catch (e) {
      console.error(`  ❌ Locator click failed: ${e.message}`);
    }
  }

  if (!clicked) return false;

  // Readiness signal for the View Purchase Request screen
  await frame.locator('text="PR Details"').first()
    .waitFor({ state: "attached", timeout: 15000 })
    .catch(() => console.log("  ⏳ Detail screen still loading..."));
  await page.waitForTimeout(1000);
  return true;
}

/**
 * Selects the first PR row (Stable Phase 1)
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
 * Clicks the "View Purchase Request" link/button on the View Purchase Request screen.
 * Extracts the PR No. from the loaded detail screen.
 */
async function triggerViewPr(page) {
  console.log('  🖱 Clicking "View Purchase Request" link...');
  const frameLocator = page.frameLocator("#ilboinnerframe");
  const viewBtn = frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /^View Purchase Request$/i }).first();

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
    const looseBtn = frameLocator.getByText('View Purchase Request', { exact: false }).first();
    clicked = await tryClick(looseBtn, "Flexible Link");
  }
  if (!clicked) {
    await frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /View Purchase Request/i }).first().evaluate(el => el.click()).catch(e => console.error("    ⚠ JS Click failed: " + e.message));
  }

  // Wait for the Purchase Request detail screen to appear.
  await frameLocator.locator('text="PR Details"').first().waitFor({ state: "attached", timeout: 15000 }).catch(() => {
    console.log("  ⏳ Detail screen still loading...");
  });
  await page.waitForTimeout(1000);

  // Extract PR No. from the detail screen
  let prNo = "Unknown";
  try {
    const frame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

    if (frame) {
      prNo = await frame.evaluate(() => {
        const elements = Array.from(document.querySelectorAll('label, td, span, div, b'));
        const labelEl = elements.find(el => {
            const text = el.innerText.trim();
            return text === 'PR No.' || text === 'PR No';
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
            const match = el.innerText.match(/PR No\.\s*[:\-]?\s*([A-Z0-9\-_]+)/i);
            if (match && match[1]) return match[1];
        }

        return "Unknown";
      });
    } else {
      console.warn("    ⚠ Could not find frame 'ilboinnerframe' for extraction.");
    }
  } catch (err) {
    console.warn("    ⚠ Failed to extract PR No: " + err.message);
  }

  return prNo;
}

/**
 * Scrapes the PR Details section of the View Purchase Request screen and writes
 * the result to `<exportDir>/Header Info - <prNo>.txt` as pretty-printed JSON.
 *
 * Must run BEFORE downloadPrDetail's Tab-9 + Shift+E sequence, since
 * that keyboard sequence moves focus and can shift the DOM layout.
 */
async function savePrDetailsJson(page, prNo, exportDir) {
  console.log("  📋 Extracting PR Details → JSON...");
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
    // Purchase Request screen with the JSON key we want in the output.
    const schema = {
      pr_details: [
        ["PR No.",                          "pr_no"],
        ["PR Date",                         "pr_date"],
        ["PR Type",                         "pr_type"],
        ["Mode",                            "mode"],
        ["Requester ID",                    "requester_id"],
        ["Folder",                          "folder"],
        ["PR Origin",                       "pr_origin"],
        ["Order At",                        "order_at"],
        ["Receipt At",                      "receipt_at"],
        ["PR Currency",                     "pr_currency"],
        ["Exchange Rate",                   "exchange_rate"],
        ["PR Value (Transaction Currency)", "pr_value_transaction_currency"],
        ["Approval Date",                   "approval_date"],
        ["Remarks",                         "remarks"],
        ["PR Value (Base Currency)",        "pr_value_base_currency"],
        ["Authorization Remarks",           "authorization_remarks"],
        ["Shortclose Reason Code",          "shortclose_reason_code"],
        ["Desc.",                           "desc"],
      ],
    };

    // Field keys whose value should be serialized as a Number, not a string.
    const numericKeys = new Set([
      "exchange_rate",
      "pr_value_transaction_currency",
      "pr_value_base_currency",
    ]);

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
    // Guarantee the PR number is populated from the caller-provided
    // value, so this field is never null even if the scrape missed it.
    structured.pr_details.pr_no = prNo;

    const outPath = `${exportDir}/Header Info - ${prNo}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save PR info JSON: ${error.message}`);
  }
}

/**
 * Downloads the Item Details Excel file from the View Purchase Request screen.
 */
async function downloadPrDetail(page, prNo = "Unknown") {
  console.log(`\n📥 Preparing to download Item Details Excel (PR: ${prNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });

    await page.waitForTimeout(1500);

    // Resolve the export folder up front — the JSON info dump and the Excel
    // download both land here, so they must agree on the date/PR path.
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Purchase Requests/${OU_FOLDER}/${dateFolder}/${prNo}`;
    mkdirSync(exportDir, { recursive: true });

    // Dump PR Details to JSON BEFORE the Tab-9 sequence, which shifts
    // keyboard focus and can reflow the DOM.
    await savePrDetailsJson(page, prNo, exportDir);

    console.log("  ⌨ Pressing Tab 9 times...");
    for (let i = 0; i < 9; i++) {
      await page.keyboard.press("Tab");
      await page.waitForTimeout(100);
    }

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");

    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });

    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const suggestedName = download.suggestedFilename();
    console.log(`  🔍 Server suggested name: "${suggestedName}"`);

    const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';

    const fileName = `${exportDir}/Item Details - (${prNo}).${ext}`;

    await download.saveAs(fileName);
    console.log(`  ✅ Excel download successful!`);
    console.log(`  📁 Saved and renamed properly in your project → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Excel download failed: ${error.message}`);
  }
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

/**
 * Navigates to the "Accounting Information" drill-down within the View Purchase Request screen.
 * (See Image 5 — Accounting Information footer link.  See Image 7 — resulting screen
 *  with credit/debit rows for accounts 1010102025 / 2010104002 etc.)
 */
/**
 * Navigates to the dedicated Accounting Information page via the footer
 * hyperlink at the bottom of View Purchase Request (Image 2). The text "Accounting
 * Information" also appears as a section header above the inline grid
 * (Image 1) — that one is NOT the target because clicking it just toggles
 * the section's collapse state. We target the LAST occurrence in DOM order
 * (the footer link comes after the section header).
 */
async function goToAccountingInformation(page) {
  console.log("\n📑 Navigating to 'Accounting Information' footer link...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  let clicked = false;

  // Strategy A — JS dispatch on the innermost visible element whose *direct*
  // text is exactly "Accounting Information", preferring the LAST match
  // (footer link comes after the section header in document order).
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

  // Strategy B (fallback) — Playwright locator, last exact-text match
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
      return;
    }
  }

  console.log("  ⏳ Waiting for accounting info page to load...");
  await page.waitForTimeout(2000);
}

/**
 * Downloads the Excel file from the Accounting Information view
 */
async function downloadAccountingExcel(page, prNo = "Unknown") {
  console.log(`\n📊 Preparing to download Accounting Excel (PR: ${prNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });

    console.log("  ⌨ Pressing Tab to focus the grid area...");
    for (let i = 0; i < 1; i++) {
        await page.keyboard.press("Tab");
        await page.waitForTimeout(100);
    }

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");

    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });

    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const suggestedName = download.suggestedFilename();
    console.log(`  🔍 Server suggested name: "${suggestedName}"`);

    const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Purchase Requests/${OU_FOLDER}/${dateFolder}/${prNo}`;
    mkdirSync(exportDir, { recursive: true });

    const fileName = `${exportDir}/Accounting Info - (${prNo}).xlsm`;

    await download.saveAs(fileName);
    console.log(`  ✅ Accounting Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Accounting Excel download failed: ${error.message}`);
  }
}

/**
 * Navigates to the "View Tax Calculation - Summary" screen from View Purchase Request.
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
 * Navigates to the "View Tax Calculation - Details" screen.
 * (See Image 6 — "View Tax Calculation - Details" footer link on View Purchase Request.)
 */
/**
 * Detects and closes the VirtualWorks error popup that appears when a PR
 * has no Line level tax summary (message: "Summary Information for the Line
 * level taxes, not available. Save tax details in summary page").
 *
 * Searches every frame for a visible ExtJS dialog/window whose text mentions
 * "VirtualWorks" or "not available", clicks its Close button (or X icon),
 * and returns true if one was dismissed.
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

          // Prefer the Close button
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
          // Fallback: X icon in top-right
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
 * Navigates to 'View Tax Calculation - Details'. If Ramco responds with the
 * "Line level taxes not available" popup instead of loading the screen, we
 * close it and return false so the caller skips the download for this PR.
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

      // Give Ramco a moment to either navigate OR pop up the error
      await page.waitForTimeout(2500);

      // If the "not available" popup came up, close it and bail — we stay on
      // View Purchase Request, so the main loop's final Go Back takes us to Select
      // PR, and the next row proceeds normally.
      const popupClosed = await dismissVirtualWorksErrorPopup(page);
      if (popupClosed) {
        console.log("  ⚠ Tax summary not available for this PR — skipping tax details.");
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

async function downloadTaxDetailsExcel(page, prNo, taxType, lineNo) {
  const hasIter = taxType && lineNo !== undefined && lineNo !== null && lineNo !== '';
  const logLabel = hasIter ? `(${taxType} Line ${lineNo}) ` : '';
  console.log(`\n📊 Preparing to download Tax Details Excel ${logLabel}(PR: ${prNo})...`);

  try {
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Purchase Requests/${OU_FOLDER}/${dateFolder}/${prNo}`;
    mkdirSync(exportDir, { recursive: true });

    await page.waitForTimeout(1000);

    console.log("  ⌨ Pressing Tab to focus the grid area...");
    for (let i = 0; i < 1; i++) {
        await page.keyboard.press("Tab");
        await page.waitForTimeout(100);
    }

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");
    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });

    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');

    const download = await downloadPromise;
    const fileSuffix = hasIter ? ` (${taxType} Line${lineNo})` : '';
    const fileName = `${exportDir}/Tax Calculation Details${fileSuffix} - (${prNo}).xls`;

    await download.saveAs(fileName);
    console.log(`  ✅ Tax Details Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Tax Details Excel download failed: ${error.message}`);
  }
}

async function changeTaxType(page, targetType) {
  console.log(`\n🔄 Changing Tax Type to '${targetType}'...`);
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const taxTypeField = frame.locator('xpath=//label[contains(text(), "Tax Type")]/following::input[1] | //td[contains(text(), "Tax Type")]/following::input[1]').first();

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

async function getRefDocLineOptions(page) {
  console.log("  🔍 Detecting 'Ref. Doc. Line No.' options...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  try {
    const field = frame.locator('xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | //td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]').first();

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
      console.log(`  ✔Found options: ${options.length}`);
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

    console.log(`  ✔Found options: ${uniqueOptions.length}`);
    console.log(`    Options: ${uniqueOptions.join(", ")}`);
    return uniqueOptions;
  } catch (error) {
    console.warn(`  ⚠ Could not detect line options: ${error.message}. Defaulting to ["1"].`);
    return ["1"];
  }
}

async function selectRefDocLine(page, lineNo) {
  console.log(`  🔢 Selecting Ref. Doc. Line No: ${lineNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const field = frame.locator('xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | //td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]').first();

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
 * Downloads the initial View Purchase Request search results grid.
 * Uses "PR Date From" as the filename anchor.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const dateField = frame.locator('xpath=//label[contains(text(), "PR Date From")]/following::input[1] | //td[contains(text(), "PR Date From")]/following::input[1]').first();
    const rawDate = await dateField.inputValue();
    console.log(`  📅 Detected PR Date From: ${rawDate}`);

    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Purchase Requests/${OU_FOLDER}/${dateFolder}`;
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
    // produce a fresh Excel for the PR-list read step below.
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
 * Read every PR number from the initial View Purchase Request Excel export.
 * Ramco's "Shift+E" export is actually a ZIP-packaged .xlsx (despite the
 * .xls extension), with the rows rendered inline in xl/worksheets/sheet*.xml.
 * We extract with Windows' bundled bsdtar (present on Windows 10 1803+) and
 * regex-match the PR pattern across the sheet XMLs — giving us the
 * full row count in one shot, no grid walking / F2 / ArrowDown needed.
 *
 * Returns an array of unique PR numbers in sheet order (which matches
 * the grid's display order).
 */
function readPrsFromExcel(xlsPath) {
  console.log(`\n📖 Reading PR list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read PR list.`);
    return [];
  }

  const tempDir = path.join(tmpdir(), `ramco-xlsx-${process.pid}-${Date.now()}`);
  mkdirSync(tempDir, { recursive: true });

  try {
    // Windows bsdtar understands ZIP containers; git-bash shadows it with
    // GNU tar (which does not), so call the Windows binary explicitly.
    const tarExe = process.env.SystemRoot
      ? path.join(process.env.SystemRoot, 'System32', 'tar.exe')
      : 'C:\\Windows\\System32\\tar.exe';

    execFileSync(tarExe, ['-xf', xlsPath, '-C', tempDir], { stdio: 'pipe' });

    const seen = new Set();
    const out = [];
    const addMatches = (text) => {
      const pattern = /\b(\d{2}[A-Z]+\d{2}-\d+)\b/g;
      let m;
      while ((m = pattern.exec(text)) !== null) {
        if (!seen.has(m[1])) {
          seen.add(m[1]);
          out.push(m[1]);
        }
      }
    };

    const sheetsDir = path.join(tempDir, 'xl', 'worksheets');
    if (existsSync(sheetsDir)) {
      // Sheets are numbered sheet1.xml, sheet2.xml, ... — iterate in order.
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

    // Also scan sharedStrings.xml — some Excel exports intern the PR
    // strings there instead of inlining them.
    const sharedPath = path.join(tempDir, 'xl', 'sharedStrings.xml');
    if (existsSync(sharedPath)) {
      addMatches(readFileSync(sharedPath, 'utf-8'));
    }

    console.log(`  ✔ Extracted ${out.length} PR(s) from Excel.`);
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

      // Perform Search logic once we reach the Purchase Request → View Purchase Request module
      if (menuPath.join(" ").includes("View Purchase Request")) {
        const stats = await performPrSearch(page);

        // Download the initial Excel export — it contains *every* PR
        // in the result set, independent of grid pagination / buffered view.
        const initialXlsPath = await downloadInitialSearchResults(page);

        // Read the full PR list straight out of that Excel file. This
        // replaces the grid-scan + F2/ArrowDown walk: one authoritative list,
        // no keyboard navigation needed. If the Excel read fails we stop —
        // there is no safe fallback that handles paginated grids reliably.
        const prNumbers = readPrsFromExcel(initialXlsPath);
        const totalRows = prNumbers.length;

        if (totalRows === 0) {
          console.error("  ❌ No PRs extracted from Excel — cannot process.");
        } else {
          if (stats.total && stats.total !== totalRows) {
            console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
          }

          // Honor RESUME_FROM_PURCHASE_REQUEST: locate the PR in the Excel list
          // and start processing from that row. Preserves the original 1-based
          // "Row N/Total" numbering so logs still line up with the Excel index.
          let startIndex = 1;
          if (CONFIG.resumeFromPr) {
            const target = CONFIG.resumeFromPr.toUpperCase();
            const matchIdx = prNumbers.findIndex(v => String(v).trim().toUpperCase() === target);
            if (matchIdx === -1) {
              console.error(`  ❌ RESUME_FROM_PURCHASE_REQUEST="${CONFIG.resumeFromPr}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
              break;
            }
            startIndex = matchIdx + 1;
            console.log(`  ⏭ RESUME_FROM_PURCHASE_REQUEST="${CONFIG.resumeFromPr}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
          }

          console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} PR(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

          for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
            const currentPrNo = prNumbers[processedCount - 1];
            process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${currentPrNo}`);

            // forceDirectEntry=true: paste the PR into the PR No.
            // Direct Entry field and click "View Purchase Request". No grid click,
            // no F2, no ArrowDown — the PR number came from the Excel
            // and is already authoritative.
            const result = await openPrViaF2(page, processedCount - 1, currentPrNo, true);
            console.log(`\n📂 --- PROCESSING PR ${currentPrNo} (${processedCount} OF ${totalRows}) ---`);

            if (!result.ok) {
              console.error(`  ❌ Could not open PR ${currentPrNo} — skipping.`);
              continue;
            }

            // 1) View Purchase Request detail export
            await downloadPrDetail(page, currentPrNo);

            // 2) Accounting Information drill-down
            // await goToAccountingInformation(page);
            // await downloadAccountingExcel(page, currentPrNo);

            // 3) Return from Accounting → View Purchase Request → View Purchase Request.
            console.log("  🔙 Returning from Accounting to View Purchase Request...");
            await goBackFromDetail(page);

            // console.log("  🔙 Navigating back to search results...");
            // await goBackFromDetail(page);

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

    writeFileSync("output/extracted-data-purchase-request.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-purchase-request.json");

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
