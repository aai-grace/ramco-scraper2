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
// output/Supplier Debit Credit Notes/. Unknown OUs fall back to a sanitized
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
  // Dedicated OU Instance for the Debit/Credit Note workflow so it can run
  // against a different OU than the Sundry script (which reads OU_INSTANCE).
  ouInstance: (process.env.OU_INSTANCE_DEBIT_CREDIT || "").trim(),
  headless: false,
  slowMo:   50,
  timeout:  30000,

  // Select Debit / Credit Note search filters (dd/mm/yyyy — same format shown
  // in the Ramco UI). Only Note Date From/To is filtered; Note Type stays on
  // "All" so both Debit and Credit notes are returned in one pass.
  noteDateFrom: process.env.NOTE_DATE_FROM || "01/04/2026",
  noteDateTo:   process.env.NOTE_DATE_TO   || "23/04/2026",

  // Resume support: if set, the Excel-driven loop fast-forwards to this
  // note number and begins processing from there. Blank = start at row 1.
  resumeFromNote: (process.env.RESUME_FROM_NOTE || "").trim(),
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// Short folder name for the active OU. Used to namespace exports under
// output/Supplier Debit Credit Notes/<OU_FOLDER>/<date>/...
const OU_FOLDER = deriveOuFolder(CONFIG.ouInstance || DEFAULT_OU_INSTANCE);
console.log(`📁 Output folder for this run: output/Supplier Debit Credit Notes/${OU_FOLDER}/`);

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Supplier Debit/Credit Note workflow:
// Payable Management → Supplier Debit credit note → View Debit / Credit Note
const MENU_PATHS = [
  ["Payable Management", "Supplier Debit credit note", "View Debit / Credit Note"],
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
 * Orchestrates the Search workflow for the "Select Debit / Credit Note" screen.
 * Filters by Note Date From/To (Note Type is left on "All" so both Debit and
 * Credit rows come back in a single pass), then clicks Search.
 */
async function performNoteSearch(page) {
  console.log("\n🔍 Preparing to search debit/credit notes inside frame...");
  const frame = await getFrame(page);

  if (!frame) {
    throw new Error("Could not find ilboinnerframe");
  }

  try {
    // 1. Clear "User ID" field (same pattern as Sundry — prefilled "CVL" etc.)
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

    // 2. Fill "Note Date From" / "Note Date To" filters.
    //    The label reads "Note Date From", and the row layout is
    //    [<Note Date From>] [input] [label "To"] [input] — so following::input[1]
    //    = From field, following::input[2] = To field (same trick as Sundry).
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
        await page.keyboard.press('Tab'); // commit the value
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
 * Note-number text pattern, e.g. "02APDN26-000098".
 * Matches 2 digits + uppercase letters + 2 digits + "-" + digits.
 */
const NOTE_NO_PATTERN = /^\d{2}[A-Z]+\d{2}-\d+$/;

/**
 * One-time setup: switch the Direct Entry "Note Category" combobox from its
 * default ("Account Based") to "All" so Direct Entry lookups can resolve
 * notes across every category — not just Account Based.
 *
 * The Select Debit / Credit Note screen has TWO "Note Category" labels (one in
 * Direct Entry, one in Search Criteria). The Direct Entry one appears first in
 * DOM order, so `(//…[text()="Note Category"])[1]/following::input[1]` targets it.
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
 * Opens a note's detail screen by pasting its Note No. into Direct Entry
 * and clicking the View link that matches the note's type + category.
 *
 * The Select Debit / Credit Note footer carries TWO sets of "View … Note"
 * links — one under Debit, one under Credit. Some Ramco builds disambiguate
 * the labels (e.g. "View Account Based Credit Note"), others reuse the same
 * generic text ("View Account Based Note") for both. We try the typed label
 * first, then fall back to selecting by DOM position (Debit = first
 * occurrence, Credit = second), since a Debit-labelled link clicked for a
 * Credit note returns: "Note No. is not Account Based Debit Note".
 *
 *   - noteType "Credit" + Account Based → "View Account Based Credit Note"
 *                                          | (or 2nd "View Account Based Note")
 *   - noteType "Debit"  + Account Based → "View Account Based Debit Note"
 *                                          | (or 1st "View Account Based Note")
 *   - "Item Based" variants follow the same pattern with "Item" instead.
 *
 * Returns { ok: boolean, noteNo: string|null }.
 */
async function openNoteViaDirectEntry(page, noteNo, noteCategory, noteType) {
  // If noteType wasn't supplied (Excel header missed), guess from the note no.
  // prefix — Ramco encodes type in it: "...DN..." = Debit, "...CN..." = Credit.
  let resolvedType = (noteType || "").trim();
  if (!resolvedType) {
    if (/[A-Z]*DN\d*-/i.test(noteNo)) resolvedType = "Debit";
    else if (/[A-Z]*CN\d*-/i.test(noteNo)) resolvedType = "Credit";
  }

  console.log(`\n🖱 Opening note via Direct Entry — ${noteNo} (${resolvedType || "unknown type"} / ${noteCategory || "unknown category"})...`);
  const frame = page.frameLocator("#ilboinnerframe");

  const isItem   = /item/i.test(noteCategory || "");
  const isCredit = /credit/i.test(resolvedType);
  const baseText  = isItem ? "View Item Based Note" : "View Account Based Note";
  const typedText = isItem
    ? (isCredit ? "View Item Based Credit Note"    : "View Item Based Debit Note")
    : (isCredit ? "View Account Based Credit Note" : "View Account Based Debit Note");
  console.log(`  🎯 Footer link target: "${typedText}" (fallback: "${baseText}" by position, ${isCredit ? "2nd" : "1st"})`);

  // Paste into "Note No." Direct Entry field (first match — excludes the
  // "Note No. From" search-criteria field thanks to the exact-text anchor).
  console.log(`  ⌨ Pasting into Direct Entry "Note No." field...`);
  const noteInput = frame
    .locator('xpath=//*[normalize-space(text())="Note No."]/following::input[1]')
    .first();
  try {
    await noteInput.waitFor({ state: "visible", timeout: 5000 });
    await noteInput.click();
    await noteInput.fill("");
    await noteInput.fill(noteNo);
    console.log(`  ✔ Pasted ${noteNo}.`);
  } catch (e) {
    console.error(`  ❌ Failed to paste into Direct Entry: ${e.message}`);
    return { ok: false, noteNo };
  }

  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));

  // Click the matching "View … Note" footer link. Strategy (in JS so we can
  // disambiguate two same-labelled links):
  //   1. Look for an element whose direct text matches typedText exactly.
  //   2. Otherwise, collect every visible element matching baseText, sort
  //      top-to-bottom / left-to-right, and pick by Credit/Debit index.
  const attemptClick = async () => {
    let info = null;
    if (actualFrame) {
      try {
        info = await actualFrame.evaluate(({ typedText, baseText, isCredit }) => {
          const isVisible = (el) => {
            if (!el || el.offsetParent === null) return false;
            const s = window.getComputedStyle(el);
            return s.display !== 'none' && s.visibility !== 'hidden';
          };
          const directTextOf = (el) => Array.from(el.childNodes)
            .filter(n => n.nodeType === Node.TEXT_NODE)
            .map(n => n.textContent.trim())
            .join('');
          const dispatch = (el) => {
            const rect = el.getBoundingClientRect();
            const x = rect.left + rect.width / 2;
            const y = rect.top  + rect.height / 2;
            ['mousedown', 'mouseup', 'click'].forEach(t => {
              el.dispatchEvent(new MouseEvent(t, {
                bubbles: true, cancelable: true, view: window,
                clientX: x, clientY: y, button: 0,
              }));
            });
          };

          // Pass 1: exact typed label.
          if (typedText && typedText !== baseText) {
            for (const el of document.querySelectorAll('*')) {
              if (!isVisible(el)) continue;
              if (directTextOf(el) === typedText) {
                dispatch(el);
                return { hit: 'typed', text: typedText, count: 1, index: 0 };
              }
            }
          }

          // Pass 2: every visible match of the base label, by reading order.
          const matches = [];
          for (const el of document.querySelectorAll('*')) {
            if (!isVisible(el)) continue;
            if (directTextOf(el) === baseText) {
              const rect = el.getBoundingClientRect();
              matches.push({ el, top: rect.top, left: rect.left });
            }
          }
          matches.sort((a, b) => a.top - b.top || a.left - b.left);
          if (matches.length === 0) return { hit: 'none', text: baseText, count: 0, index: -1 };
          const idx = matches.length > 1 ? (isCredit ? 1 : 0) : 0;
          dispatch(matches[idx].el);
          return { hit: 'base', text: baseText, count: matches.length, index: idx };
        }, { typedText, baseText, isCredit });
      } catch (e) {
        console.warn(`  ⚠ JS dispatch failed: ${e.message}`);
      }
    }

    if (info && info.hit !== 'none') {
      const tag = info.hit === 'base' ? `, ${info.index + 1}/${info.count}` : '';
      console.log(`  ✔ Clicked "${info.text}" via JS dispatch (${info.hit}${tag}).`);
      return true;
    }

    // Locator fallback: try the typed label, then the base label by index.
    try {
      const typedLink = frame.getByText(typedText, { exact: true }).first();
      if (await typedLink.isVisible({ timeout: 1500 }).catch(() => false)) {
        await typedLink.click({ force: true });
        console.log(`  ✔ Clicked "${typedText}" via locator.`);
        return true;
      }
      const baseAll = frame.getByText(baseText, { exact: true });
      const count = await baseAll.count().catch(() => 0);
      if (count === 0) {
        console.error(`  ❌ No footer link matching "${typedText}" or "${baseText}" found.`);
        return false;
      }
      const idx = count > 1 ? (isCredit ? 1 : 0) : 0;
      await baseAll.nth(idx).click({ force: true });
      console.log(`  ✔ Clicked "${baseText}" via locator (${idx + 1}/${count}).`);
      return true;
    } catch (e) {
      console.error(`  ❌ Failed to click footer link: ${e.message}`);
      return false;
    }
  };

  console.log(`  🖱 Clicking footer link...`);
  if (!(await attemptClick())) return { ok: false, noteNo };

  // Poll for one of three outcomes:
  //   (a) detail page attaches → continue
  //   (b) "Note Number does not exist" popup → close, retry once, else skip
  //   (c) "is not Account/Item Based Debit/Credit Note" popup → wrong link
  //       picked; close and skip (caller log will explain).
  const findErrorPopup = async () => {
    for (const root of [page, frame]) {
      const panel = root.locator('#message-panel').first();
      const visible = await panel.isVisible({ timeout: 0 }).catch(() => false);
      if (!visible) continue;
      const text = await panel.innerText().catch(() => "");
      if (/does not exist/i.test(text))                       return { panel, text, kind: 'missing' };
      if (/is not\s+(account|item)\s+based/i.test(text))      return { panel, text, kind: 'wrongType' };
    }
    return null;
  };

  const isDetailVisible = async () => {
    // The detail screen shows a "Note Information" header (mirroring
    // "Voucher Information" on the Sundry side). Fall back to generic
    // "Supplier Information" if that title isn't present.
    for (const t of ['Note Information', 'Supplier Information']) {
      const visible = await frame.locator(`text="${t}"`).first()
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
      const oneLine = (popup.text || "").replace(/\s+/g, " ").trim();
      if (popup.kind === 'wrongType') {
        // Wrong link picked: the screen rejected the note as not matching the
        // link's type/category. Retrying the same link won't help — skip and
        // surface the popup text so the user sees what Ramco said.
        console.error(`  ❌ Wrong link for ${noteNo}: "${oneLine}"`);
        await dismissPopup(popup.panel);
        return { ok: false, noteNo };
      }

      console.warn(`  ⚠ Popup on attempt ${attempt}/2: "${oneLine}". Closing...`);
      await dismissPopup(popup.panel);

      if (attempt === 2) {
        console.error(`  ❌ Note ${noteNo} still failing after retry — skipping.`);
        return { ok: false, noteNo };
      }

      console.log(`  🔄 Retrying: re-paste + click footer link...`);
      try {
        await noteInput.click();
        await noteInput.fill("");
        await noteInput.fill(noteNo);
      } catch (e) {
        console.error(`  ❌ Retry paste failed: ${e.message}`);
        return { ok: false, noteNo };
      }
      if (!(await attemptClick())) return { ok: false, noteNo };
      continue;
    }

    console.log(`  ⏳ No popup or detail within 10s on attempt ${attempt}, falling through...`);
    break;
  }

  if (!detailLoaded) {
    await frame.locator('text="Note Information"').first()
      .waitFor({ state: "attached", timeout: 15000 })
      .catch(() => console.log("  ⏳ Detail screen still loading..."));
  }
  await page.waitForTimeout(1000);
  return { ok: true, noteNo };
}

/**
 * Scrapes the four header sections of the View Account/Item Based Note screen
 * (Note Information, Supplier Information, Payment Information, Project
 * Information) and writes the result to `<exportDir>/Header Info - <noteNo>.txt`
 * as pretty-printed JSON.
 *
 * Must run BEFORE downloadNoteDetail's Tab-9 + Shift+E sequence — that
 * sequence moves keyboard focus and can reflow the DOM.
 */
async function saveNoteInfoJson(page, noteNo, exportDir) {
  console.log("  📋 Extracting Note / Supplier / Payment / Project Information → JSON...");
  try {
    const actualFrame = page.frames().find(
      (f) => f.name() === "ilboinnerframe" || f.url().includes("rvw_blank.aspx")
    );
    if (!actualFrame) {
      console.warn("    ⚠ Could not locate ilboinnerframe — skipping JSON dump.");
      return;
    }

    const schema = {
      note_information: [
        ["Note No.",      "note_no"],
        ["Note Type",     "note_type"],
        ["Status",        "status"],
        ["Note Date",     "note_date"],
        ["Finance Book",  "finance_book"],
      ],
      supplier_information: [
        ["Supplier Registered At", "supplier_registered_at"],
        ["Supplier Code",          "supplier_code"],
        ["Supplier Name",          "supplier_name"],
        ["Supplier Note No.",      "supplier_note_no"],
        ["Supplier Note Amount",   "supplier_note_amount"],
        ["Supplier Note Date",     "supplier_note_date"],
        ["Supplier Account Code",  "supplier_account_code"],
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
    const numericKeys = new Set(["exchange_rate", "supplier_note_amount", "total_amount"]);

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

      const map = {};
      for (const lbl of labels) {
        const labelEl = findLabelEl(lbl);
        if (!labelEl) { map[lbl] = ""; continue; }
        const labelRect = labelEl.getBoundingClientRect();

        let value = "";

        // Strategy 1: spatial proximity (the Ramco layout renders labels to
        // the LEFT of their value on the same horizontal band).
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
    // Guarantee the note number is populated from the caller-provided
    // value, so this field is never null even if the scrape missed it.
    structured.note_information.note_no = noteNo;

    const outPath = `${exportDir}/Header Info - ${noteNo}.txt`;
    writeFileSync(outPath, JSON.stringify(structured, null, 2), "utf8");
    console.log(`  ✅ Saved info JSON → ${outPath}`);
  } catch (error) {
    console.warn(`  ⚠ Failed to save note info JSON: ${error.message}`);
  }
}

/**
 * Downloads the Note detail Excel file from the View ... Note screen.
 * Mirrors downloadVoucherDetail from the Sundry script: Tab 9× to seat focus
 * on the grid area, then Shift+E to trigger Ramco's Excel export.
 */
async function downloadNoteDetail(page, noteNo = "Unknown") {
  console.log(`\n📥 Preparing to download Note Detail Excel (Note: ${noteNo})...`);

  try {
    mkdirSync("output/exports", { recursive: true });

    await page.waitForTimeout(1500);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Supplier Debit Credit Notes/${OU_FOLDER}/${dateFolder}/${noteNo}`;
    mkdirSync(exportDir, { recursive: true });

    // Dump the four header sections to JSON BEFORE the Tab-9 sequence,
    // which shifts keyboard focus and can reflow the DOM.
    await saveNoteInfoJson(page, noteNo, exportDir);

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

    const fileName = `${exportDir}/Note Info - (${noteNo}).${ext}`;

    await download.saveAs(fileName);
    console.log(`  ✅ Excel download successful!`);
    console.log(`  📁 Saved and renamed properly in your project → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Excel download failed: ${error.message}`);
  }
}

/**
 * Navigates to the dedicated Accounting Information page via the footer
 * hyperlink at the bottom of the View Account/Item Based Note detail screen.
 *
 * The text "Accounting Information" can appear twice: once as a section
 * header (collapsible — clicking it just toggles the section) and once as
 * the footer drill-down link (the one we want). We target the LAST match
 * in DOM order, which is the footer link.
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

  console.log("  ⏳ Waiting for accounting info page to load...");
  await page.waitForTimeout(2000);
  return true;
}

/**
 * Downloads the Excel file from the Accounting Information view for a note.
 */
async function downloadAccountingExcel(page, noteNo = "Unknown") {
  console.log(`\n📊 Preparing to download Accounting Excel (Note: ${noteNo})...`);

  try {
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
    const exportDir = `output/Supplier Debit Credit Notes/${OU_FOLDER}/${dateFolder}/${noteNo}`;
    mkdirSync(exportDir, { recursive: true });

    const fileName = `${exportDir}/Accounting Info - (${noteNo}).${ext}`;

    await download.saveAs(fileName);
    console.log(`  ✅ Accounting Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Accounting Excel download failed: ${error.message}`);
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
 * Downloads the initial "Select Debit / Credit Note" search-results grid.
 * Uses "Note Date From" as the filename anchor (per the Search Criteria panel).
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    const dateField = frame.locator(
      'xpath=//label[contains(text(), "Note Date From")]/following::input[1] | ' +
      '//td[contains(text(), "Note Date From")]/following::input[1]'
    ).first();
    const rawDate = await dateField.inputValue();
    console.log(`  📅 Detected Note Date From: ${rawDate}`);

    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/Supplier Debit Credit Notes/${OU_FOLDER}/${dateFolder}`;
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
 * Read every note number (and its Note Category) from the initial Select
 * Debit / Credit Note Excel export. Ramco's "Shift+E" export is a
 * ZIP-packaged .xlsx — we extract with Windows' bundled bsdtar and parse
 * the sheet XML / sharedStrings.xml to build ordered rows of:
 *    { noteNo, noteCategory }
 *
 * We need the category because it determines which footer link to click
 * ("View Account Based Note" vs "View Item Based Note").
 */
function readNotesFromExcel(xlsPath) {
  console.log(`\n📖 Reading note list from Excel: ${xlsPath}`);
  if (!xlsPath || !existsSync(xlsPath)) {
    console.warn(`  ⚠ Excel file not found — cannot read note list.`);
    return [];
  }

  const tempDir = path.join(tmpdir(), `ramco-xlsx-${process.pid}-${Date.now()}`);
  mkdirSync(tempDir, { recursive: true });

  try {
    const tarExe = process.env.SystemRoot
      ? path.join(process.env.SystemRoot, 'System32', 'tar.exe')
      : 'C:\\Windows\\System32\\tar.exe';

    execFileSync(tarExe, ['-xf', xlsPath, '-C', tempDir], { stdio: 'pipe' });

    // sharedStrings.xml holds interned text values referenced by <c t="s"><v>INDEX</v></c>
    // cells. Parse the <si> entries in order so we can resolve those indices.
    const sharedStrings = [];
    const sharedPath = path.join(tempDir, 'xl', 'sharedStrings.xml');
    if (existsSync(sharedPath)) {
      const xml = readFileSync(sharedPath, 'utf-8');
      const siRegex = /<si[^>]*>([\s\S]*?)<\/si>/g;
      const tRegex = /<t[^>]*>([\s\S]*?)<\/t>/g;
      let m;
      while ((m = siRegex.exec(xml)) !== null) {
        let combined = '';
        let tm;
        while ((tm = tRegex.exec(m[1])) !== null) {
          combined += tm[1];
        }
        sharedStrings.push(
          combined
            .replace(/&lt;/g, '<')
            .replace(/&gt;/g, '>')
            .replace(/&amp;/g, '&')
            .replace(/&quot;/g, '"')
            .replace(/&apos;/g, "'")
        );
      }
    }

    // Parse each sheet's rows into arrays of cell text.
    const noteNoPattern = /^\d{2}[A-Z]+\d{2}-\d+$/;
    const out = [];
    const seen = new Set();

    const parseRow = (rowXml) => {
      // <c r="A5" t="s" s="..."><v>42</v></c>   → sharedStrings[42]
      // <c r="B5"><v>123</v></c>                → inline number
      // <c r="C5" t="inlineStr"><is><t>...</t></is></c>
      const cells = [];
      const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
      let cm;
      while ((cm = cellRegex.exec(rowXml)) !== null) {
        const attrs = cm[1];
        const inner = cm[2];
        const refMatch = attrs.match(/r="([A-Z]+)\d+"/);
        const colLetters = refMatch ? refMatch[1] : "";
        const colIdx = colLetters
          ? colLetters.split('').reduce((n, c) => n * 26 + (c.charCodeAt(0) - 64), 0) - 1
          : cells.length;
        const typeMatch = attrs.match(/t="([^"]+)"/);
        const type = typeMatch ? typeMatch[1] : "";
        let text = "";
        if (type === "s") {
          const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);
          if (vMatch) {
            const idx = parseInt(vMatch[1], 10);
            text = sharedStrings[idx] || "";
          }
        } else if (type === "inlineStr") {
          const tMatch = inner.match(/<t[^>]*>([\s\S]*?)<\/t>/);
          if (tMatch) text = tMatch[1];
        } else {
          const vMatch = inner.match(/<v>([\s\S]*?)<\/v>/);
          if (vMatch) text = vMatch[1];
        }
        cells[colIdx] = (text || "").trim();
      }
      return cells;
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
        const xml = readFileSync(path.join(sheetsDir, f), 'utf-8');

        // Locate the header row first so we know which column holds
        // Note No. vs. Note Category. Ramco exports carry human-readable
        // headers in the first populated row.
        const rowRegex = /<row\b([^>]*)>([\s\S]*?)<\/row>/g;
        let rm;
        let noteNoCol = -1;
        let noteCatCol = -1;
        let noteTypeCol = -1;
        const rows = [];
        while ((rm = rowRegex.exec(xml)) !== null) {
          const cells = parseRow(rm[2]);
          rows.push(cells);
          if (noteNoCol === -1) {
            for (let i = 0; i < cells.length; i++) {
              // Normalize header text: lowercase, NBSP→space, collapse whitespace,
              // drop trailing punctuation. Ramco exports occasionally pad headers
              // with NBSP, double spaces, or trailing "." which broke exact match.
              const v = (cells[i] || "")
                .toLowerCase()
                .replace(/ /g, " ")
                .replace(/\s+/g, " ")
                .trim()
                .replace(/[.:]+$/, "");
              if (v === "note no") noteNoCol = i;
              else if (v === "note category") noteCatCol = i;
              else if (v === "note type") noteTypeCol = i;
            }
          }
        }

        // Pass 1: if we found the header columns, use them for ordered rows.
        if (noteNoCol >= 0) {
          for (const cells of rows) {
            const raw = cells[noteNoCol] || "";
            if (!noteNoPattern.test(raw)) continue;
            if (seen.has(raw)) continue;
            seen.add(raw);
            const category = noteCatCol  >= 0 ? (cells[noteCatCol]  || "") : "";
            const type     = noteTypeCol >= 0 ? (cells[noteTypeCol] || "") : "";
            out.push({ noteNo: raw, noteCategory: category, noteType: type });
          }
        } else {
          // Fallback: no header row detected — regex over the raw XML and
          // leave category/type blank (openNoteViaDirectEntry will guess from
          // the noteNo prefix: "...DN..." → Debit, "...CN..." → Credit).
          const pattern = /\b(\d{2}[A-Z]+\d{2}-\d+)\b/g;
          let pm;
          while ((pm = pattern.exec(xml)) !== null) {
            const v = pm[1];
            if (seen.has(v)) continue;
            seen.add(v);
            out.push({ noteNo: v, noteCategory: "", noteType: "" });
          }
        }
      }
    }

    // Safety net: scan sharedStrings too for any note numbers we might have missed.
    for (const s of sharedStrings) {
      const m2 = s.match(/\b(\d{2}[A-Z]+\d{2}-\d+)\b/);
      if (m2 && !seen.has(m2[1])) {
        seen.add(m2[1]);
        out.push({ noteNo: m2[1], noteCategory: "", noteType: "" });
      }
    }

    console.log(`  ✔ Extracted ${out.length} note(s) from Excel.`);
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

      if (menuPath.join(" ").includes("View Debit / Credit Note")) {
        const stats = await performNoteSearch(page);

        // Download the initial Excel export — contains every note in the
        // result set, independent of grid pagination / buffered view.
        const initialXlsPath = await downloadInitialSearchResults(page);

        // Read the full note list (Note No. + Note Category) from that Excel.
        const notes = readNotesFromExcel(initialXlsPath);
        const totalRows = notes.length;

        if (totalRows === 0) {
          console.error("  ❌ No notes extracted from Excel — cannot process.");
        } else {
          if (stats.total && stats.total !== totalRows) {
            console.warn(`  ⚠ Grid toolbar reports ${stats.total} total but Excel has ${totalRows}. Processing the Excel list.`);
          }

          // Honor RESUME_FROM_NOTE
          let startIndex = 1;
          if (CONFIG.resumeFromNote) {
            const target = CONFIG.resumeFromNote.toUpperCase();
            const matchIdx = notes.findIndex(n => String(n.noteNo).trim().toUpperCase() === target);
            if (matchIdx === -1) {
              console.error(`  ❌ RESUME_FROM_NOTE="${CONFIG.resumeFromNote}" not found in Excel list — aborting to avoid silently restarting from row 1.`);
              break;
            }
            startIndex = matchIdx + 1;
            console.log(`  ⏭ RESUME_FROM_NOTE="${CONFIG.resumeFromNote}" — skipping rows 1..${startIndex - 1}, resuming at row ${startIndex}.`);
          }

          console.log(`\n🚀 Starting continuous processing for ${totalRows - startIndex + 1} note(s) (Excel-driven, row ${startIndex}/${totalRows})...`);

          // One-time Direct Entry setup: switch Note Category to "All" so
          // Direct Entry lookups resolve notes of either category.
          // (Note Type has no "All" option — only Debit / Credit — so we
          // leave whatever default Ramco shows; the View link choice is
          // driven by the row's category, not this dropdown.)
          await setDirectEntryNoteCategoryToAll(page);

          for (let processedCount = startIndex; processedCount <= totalRows; processedCount++) {
            const { noteNo, noteCategory, noteType } = notes[processedCount - 1];
            process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}]... ${noteNo} (${noteType || "?"} / ${noteCategory || "?"})`);

            const openResult = await openNoteViaDirectEntry(page, noteNo, noteCategory, noteType);
            console.log(`\n📂 --- PROCESSING NOTE ${noteNo} (${processedCount} OF ${totalRows}) ---`);

            if (!openResult.ok) {
              console.error(`  ❌ Could not open note ${noteNo} — skipping.`);
              continue;
            }

            // 1) View ... Note detail export
            await downloadNoteDetail(page, noteNo);

            // 2) Accounting Information drill-down (footer link → dedicated
            //    accounting page → Shift+E export).
            const accountingOk = await goToAccountingInformation(page);
            if (accountingOk) {
              await downloadAccountingExcel(page, noteNo);
              console.log("  🔙 Returning from Accounting to View Note...");
              await goBackFromDetail(page);
            }

            // 3) Return from View ... Note → Select Debit / Credit Note.
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

    writeFileSync("output/extracted-data-debit-credit.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data-debit-credit.json");

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
