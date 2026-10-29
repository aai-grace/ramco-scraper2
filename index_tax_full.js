require("dotenv").config();
const { chromium } = require("playwright");
const { mkdirSync, writeFileSync, existsSync } = require("fs");
const readline = require("readline");

// ─── CONFIG (loaded from .env) ─────────────────────────────────────────────
const CONFIG = {
  url:      process.env.RAMCO_URL      || "https://aai.ramcoes.com/RVW/extui/vwrt/LaunchPanel.htm",
  username: process.env.RAMCO_USERNAME || "",
  password: process.env.RAMCO_PASSWORD || "",
  headless: false,
  slowMo:   50,   // Reduced from 300ms for faster processing
  timeout:  30000,

  // Select Voucher / search-screen date filters (dd/mm/yyyy — matches Ramco UI)
  payDateFrom:     process.env.PAY_DATE_FROM     || "01/04/2026",
  payDateTo:       process.env.PAY_DATE_TO       || "03/04/2026",
  requestDateFrom: process.env.REQUEST_DATE_FROM || "01/04/2026",
  requestDateTo:   process.env.REQUEST_DATE_TO   || "03/04/2026",
};

if (!CONFIG.username || !CONFIG.password) {
  console.error("❌ Missing credentials. Please set RAMCO_USERNAME and RAMCO_PASSWORD in your .env file.");
  process.exit(1);
}

// ─── MENU PATHS TO NAVIGATE ────────────────────────────────────────────────
// Add or remove paths as needed. Each array is a sequence of menu clicks.
const MENU_PATHS = [
  ["Payable Management", "Supplier Direct Invoice", "View Invoice"],
];

// ─── HELPERS ───────────────────────────────────────────────────────────────
/**
 * Consistent helper for terminal input using readline
 */
function waitForEnter(msg = "👉 Press ENTER to close the browser...") {
  const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
  return new Promise((resolve) => {
    rl.question("\n" + msg + "\n", () => {
      rl.close();
      resolve();
    });
  });
}

/**
 * Interactive question with a countdown timer
 */
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
  // await page.screenshot({ path: "screenshots/01-app-shell.png" });

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
    // await page.screenshot({ path: "screenshots/error-no-login-frame.png" });
    throw new Error("❌ Could not find login frame.");
  }

  // We know the IDs from previous runs
  const userSel = "#ide_username";
  const pwSel   = "#ide_password";

  // Wait for fields to be attached
  await loginFrame.waitForSelector(userSel, { state: "attached", timeout: CONFIG.timeout });
  await loginFrame.waitForSelector(pwSel,   { state: "attached", timeout: CONFIG.timeout });

  // Type credentials slowly so ExtJS registers each keystroke
  // Type credentials faster
  console.log("  ✏ Typing username...");
  await loginFrame.focus(userSel);
  await loginFrame.fill(userSel, "");
  await loginFrame.type(userSel, CONFIG.username, { delay: 30 }); // Faster typing
  await page.waitForTimeout(200);

  console.log("  ✏ Typing password...");
  await loginFrame.focus(pwSel);
  await loginFrame.fill(pwSel, "");
  await loginFrame.type(pwSel, CONFIG.password, { delay: 30 });
  await page.waitForTimeout(300);

  // await page.screenshot({ path: "screenshots/02-filled.png" });

  // Submit: press Enter on password field
  console.log("  🖱 Pressing Enter to submit...");
  await loginFrame.press(pwSel, "Enter");
  await page.waitForTimeout(1000);

  // Check if still on login page — if so, try JS click fallback
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

  // await page.screenshot({ path: "screenshots/03-after-login.png" });
  console.log("✅ Login complete.");
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
    '#home button', // From the user's HTML snippet
  ];

  for (const sel of homeSelectors) {
    try {
      const el = page.locator(sel).first();
      if (await el.isVisible({ timeout: 2000 }).catch(() => false)) {
        await el.click();
        console.log(`  ✔ Clicked home: ${sel}`);
        // Smart wait for dashboard menu instead of fixed 5s
        await page.waitForSelector('#tb-bpc-menu', { timeout: 8000 }).catch(() => {});
        break;
      }
    } catch (_) {}
  }

  // Ensure dashboard elements are visible
  await page.waitForSelector('#tb-bpc-menu', { timeout: 10000 }).catch(() => {});
  
  // await page.screenshot({ path: "screenshots/04-landing-page.png" });
  // console.log("📸 Landing page → screenshots/04-landing-page.png");
}

// ─── NAVIGATE MENU ─────────────────────────────────────────────────────────
async function navigateMenu(page, menuPath) {
  console.log(`\n📋 Navigating: ${menuPath.join(" → ")}`);

  // 1. Ensure BPC Menu is Expanded
  const isMenuCollapsed = await page.evaluate(() => {
    const tp = document.getElementById('ilbo-tp');
    const collapsedBar = document.getElementById('ilbo-tp-xcollapsed');
    // If Task Manager panel is missing or has collapsed class, or the collapsed bar is visible
    return !tp || tp.classList.contains('x-panel-collapsed') || (collapsedBar && window.getComputedStyle(collapsedBar).display !== 'none');
  });

  if (isMenuCollapsed) {
    console.log("  📂 Menu is collapsed. Expanding...");
    
    // Try clicking the BPC button or the expand tool
    const expandSelectors = [
      '#tb-bpc-menu button',
      '#ext-gen309', // The expand west tool from user HTML
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

  // 2. Iterate through menu path levels
  for (let i = 0; i < menuPath.length; i++) {
    const label = menuPath[i];
    console.log(`  [${i + 1}/${menuPath.length}] Looking for "${label}"...`);

    let clicked = false;

    // Enhanced selectors for ExtJS tree nodes and flyout menu items
    const selectors = [
      // ExtJS flyout menu items (standard)
      `.x-menu-item-text:has-text("${label}")`,
      `.x-menu-item:has-text("${label}")`,
      // ExtJS tree nodes
      `.x-tree-node-text:has-text("${label}")`,
      `.x-tree-node-anchor:has-text("${label}")`,
      // Generic selectors
      `span:has-text("${label}")`,
      `a:has-text("${label}")`,
      `div:has-text("${label}")`,
    ];

    for (const sel of selectors) {
      try {
        // Use a more flexible regex to handle potential leading/trailing whitespace or icons
        const el = page.locator(sel).filter({ hasText: new RegExp(`^\\s*${label}\\s*$`, 'i') }).first();
        
        // Wait a small moment for visibility check
        if (await el.isVisible({ timeout: 3000 }).catch(() => false)) {
          await el.scrollIntoViewIfNeeded().catch(() => {});
          
          if (i < menuPath.length - 1) {
            console.log(`    🖱 Hovering "${label}"...`);
            await el.hover({ force: true }).catch(() => {});
            await page.waitForTimeout(400); // Reduced delay
          }

          console.log(`    🖱 Clicking animated "${label}"...`);
          await el.click();
          await page.waitForTimeout(1000); // Reduced delay
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
      // await page.screenshot({ path: `screenshots/failed-${label.replace(/\s+/g, "_")}.png` });
      
      // If the last item failed, try search box as a final fallback
      if (i === menuPath.length - 1) {
        console.log("    🔍 Attempting search box fallback for final item...");
        await page.fill('#bpcsearch', label);
        await page.waitForTimeout(500);
        await page.press('#bpcsearch', 'Enter');
        
        // Verification: wait for the frame or breadcrumb to reflect the change
        await page.waitForTimeout(4000); 
        console.log("    ✔ Search submitted.");
        clicked = true; // Assume success for flow purposes, verify via screenshot
      }
    }
  }

  // Wait for the final page/panel to load
  console.log("  ⏳ Waiting for page to load...");
  await page.waitForLoadState("networkidle").catch(() => {});
  await page.waitForTimeout(3000);

  const screenshotPath = `screenshots/result-${menuPath.join("-").replace(/[\s/]+/g, "_")}.png`;
  // await page.screenshot({ path: screenshotPath, fullPage: true });
  // console.log(`  📸 Initial load screenshot → ${screenshotPath}`);

  return { menuPath, url: page.url(), screenshot: screenshotPath };
}

/**
 * Helper to get the actual Frame object (necessary for .evaluate and .waitForSelector)
 */
async function getFrame(page) {
  const frame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('ilrt/rvw_blank.aspx') || f.url().includes('ilbo'));
  if (!frame) {
    // Fallback search
    return page.mainFrame().childFrames()[0];
  }
  return frame;
}

/**
 * Orchestrates the Search workflow
 */
async function performInvoiceSearch(page) {
  console.log("\n🔍 Preparing to search invoices inside frame...");
  const frame = await getFrame(page);
  
  if (!frame) {
    throw new Error("Could not find ilboinnerframe");
  }

  try {
    // 1. Clear "User ID" field
    console.log('  ⌨ Attempting to clear "User ID" field...');
    
    // We look for an input field situated near a label or TD containing "User ID"
    const userIdField = frame.locator('input').filter({ hasText: "" }).and(
      frame.locator('xpath=//label[contains(text(), "User ID")]/following::input[1] | //td[contains(text(), "User ID")]/following::input[1] | //span[contains(text(), "User ID")]/following::input[1]')
    ).first();

    // Fallback strategy: find input by its currently visible value (e.g. "CVL" seen in user screenshot)
    const fallbackField = frame.locator('input[value*="CVL"]').first();

    if (await userIdField.isVisible({ timeout: 5000 }).catch(() => false)) {
      console.log('  ✅ User ID field found via label.');
      await userIdField.click();
      // Use Control+A and Backspace for aggressive clearing
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
        // Focus the first input and clear it
        await frame.locator('input').first().focus().catch(() => {});
        await page.keyboard.down('Control');
        await page.keyboard.press('a');
        await page.keyboard.up('Control');
        await page.keyboard.press('Backspace');
      }
    }

    // 2. Fill "Pay Date From/To" and "Request Date From/To" filters from .env.
    //    Each row in the UI is: [<Label>From] [<input From>] [label "To"] [<input To>]
    //    so following::input[1] = From field, following::input[2] = To field.
    //    fillDateField silently warns and continues if the label isn't on the
    //    current screen, so this is safe on pages that don't have these filters.
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

    console.log(`  📅 Filling Pay Date:     ${CONFIG.payDateFrom} → ${CONFIG.payDateTo}`);
    await fillDateField("Pay Date From", 1, CONFIG.payDateFrom);
    await fillDateField("Pay Date From", 2, CONFIG.payDateTo);
    console.log('  ✅ Pay Date range set.');

    console.log(`  📅 Filling Request Date: ${CONFIG.requestDateFrom} → ${CONFIG.requestDateTo}`);
    await fillDateField("Request Date From", 1, CONFIG.requestDateFrom);
    await fillDateField("Request Date From", 2, CONFIG.requestDateTo);
    console.log('  ✅ Request Date range set.');

    // 3. Click the Search button
    console.log('  🖱 Clicking "Search" button...');
    const searchBtn = frame.locator('button, .x-btn-text').filter({ hasText: /^Search$/i }).last();
    
    if (await searchBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await searchBtn.click();
      console.log("  🚀 Search triggered.");
      
      // Smart wait for grid rows instead of fixed 6s
      await frame.locator('.x-grid3-row').first().waitFor({ state: "visible", timeout: 15000 }).catch(() => {
        console.log("  ⏳ Still waiting for grid rows...");
      });
      
      // Handle potential "Invoice Number does not exist" popup
      const closeBtn = frame.locator('button').filter({ hasText: /^Close$/i }).first();
      if (await closeBtn.isVisible({ timeout: 1500 }).catch(() => false)) {
        console.warn("  ⚠ Search message popup detected. Closing it...");
        await closeBtn.click();
        await page.waitForTimeout(500);
      }
    } else {
      console.error("  ❌ Search button not found in frame.");
    }

    const screenshotPath = "screenshots/05-search-results.png";
    // await page.screenshot({ path: screenshotPath, fullPage: true });
    // console.log(`  📸 Search results captured → ${screenshotPath}`);

    // --- NEW: Extract Grid Stats (for Pagination) ---
    console.log("  📊 Extracting Grid metadata...");
    const stats = await frame.evaluate(() => {
      // Find the paging info text (e.g. "6 - 10 / 65")
      const pagingText = document.querySelector('.x-paging-info')?.innerText || "";
      const match = pagingText.match(/(\d+)\s*$/); // Get the last number
      const total = match ? parseInt(match[1]) : 0;
      
      // Also detect how many rows are visible on THIS page
      const visibleRows = document.querySelectorAll('.x-grid3-row').length;
      
      return { total, visibleRows };
    });
    
    console.log(`  ✔ Grid detected: Total ${stats.total} rows, ${stats.visibleRows} on current page.`);
    return stats;

  } catch (error) {
    console.error(`  ❌ Search automation failed: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-search.png" });
    return { total: 0, visibleRows: 0 };
  }
}


/**
 * Selects the first invoice row (Stable Phase 1)
 */
async function selectFirstRow(page) {
  console.log("\n✅ Selecting first row...");
  const frame = page.frameLocator("#ilboinnerframe");

  // 1. Refocus the "User ID" field to start the tabbing sequence
  console.log('  ⌨ Focusing "User ID" field to start keyboard sequence...');
  const userIdField = frame.locator('xpath=//label[contains(text(), "User ID")]/following::input[1] | //td[contains(text(), "User ID")]/following::input[1] | //span[contains(text(), "User ID")]/following::input[1]').first();
  
  await userIdField.waitFor({ state: "attached", timeout: 10000 });
  await userIdField.focus();
  await page.waitForTimeout(1000);

  // 2. Perform the Tab sequence (5 times)
  console.log("  ⌨ Pressing Tab 5 times...");
  for (let i = 0; i < 5; i++) {
    await page.keyboard.press("Tab");
    await page.waitForTimeout(150); // Faster tabs
  }

  // 3. Press Left Key and SpaceBar
  console.log("  ⌨ Pressing ArrowLeft and SpaceBar...");
  await page.keyboard.press("ArrowLeft");
  await page.waitForTimeout(200);
  await page.keyboard.press(" "); // Space to check
  console.log("  ✔ Row 1 selected.");
  await page.waitForTimeout(500);
}

/**
 * Selects the NEXT row using dynamic keyboard sequences
 */
async function selectNextRow(page, processedCount) {
  console.log(`\n✅ Selecting row ${processedCount}...`);
  
  if (processedCount >= 6) {
    // Row 6 and beyond: Use continuous ArrowDown + Space
    console.log(`  ⌨ [Row ${processedCount}] Pressing ArrowDown and SpaceBar...`);
    await page.keyboard.press("ArrowDown");
    await page.waitForTimeout(300);
    await page.keyboard.press(" "); // Space to check
  } else {
    // Row 2 to 5: Use Tab + Right + Space
    console.log(`  ⌨ [Row ${processedCount}] Pressing Tab, ArrowRight, and SpaceBar...`);
    await page.keyboard.press("Tab");
    await page.waitForTimeout(300);
    await page.keyboard.press("ArrowRight");
    await page.waitForTimeout(300);
    await page.keyboard.press(" "); // Space to check
  }
  
  console.log(`  ✔ Row ${processedCount} selected.`);
  await page.waitForTimeout(500);
}

/**
 * Common function to trigger the "View Expense Invoice" button
 */
async function triggerViewInvoice(page) {
  console.log('  🖱 Clicking "View Expense Invoice" button...');
  const frameLocator = page.frameLocator("#ilboinnerframe");
  const viewBtn = frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /^View Expense Invoice$/i }).first();
  
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

  let clicked = await tryClick(viewBtn, "Exact Button");
  if (!clicked) {
    const looseBtn = frameLocator.getByText('View Expense Invoice', { exact: false }).first();
    clicked = await tryClick(looseBtn, "Flexible Button");
  }
  if (!clicked) {
    await frameLocator.locator('a, button, .x-btn-text').filter({ hasText: /View Expense Invoice/i }).first().evaluate(el => el.click()).catch(e => console.error("    ⚠ JS Click failed: " + e.message));
  }

  // Smart wait for detail screen element instead of 10s
  await frameLocator.locator('text="Invoice Header"').first().waitFor({ state: "attached", timeout: 15000 }).catch(() => {
    console.log("  ⏳ Detail screen still loading...");
  });
  await page.waitForTimeout(1000); 

  // --- FIX: Access the actual Frame object for .evaluate() ---
  let invoiceNo = "Unknown";
  try {
    // Attempt to find the frame by name or ID
    const frame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
    
    if (frame) {
      invoiceNo = await frame.evaluate(() => {
        // Find the label "Invoice No."
        // We look for the exact label cell or label element
        const elements = Array.from(document.querySelectorAll('label, td, span, div, b'));
        const labelEl = elements.find(el => {
            const text = el.innerText.trim();
            return text === 'Invoice No.' || text === 'Invoice No';
        });
        
        if (labelEl) {
          // Option A: If it's a TD in a table (very common in Ramco/ExtJS), the value is usually in the next TD
          const td = labelEl.closest('td');
          if (td && td.nextElementSibling) {
              const val = td.nextElementSibling.innerText.trim();
              if (val) return val;
          }
          
          // Option B: Check if it's a label for an input
          if (labelEl.tagName === 'LABEL' && labelEl.getAttribute('for')) {
            const input = document.getElementById(labelEl.getAttribute('for'));
            if (input) return input.value || input.innerText;
          }

          // Option C: Look at sibling elements directly
          const next = labelEl.nextElementSibling;
          if (next && next.innerText.trim()) return next.innerText.trim();
        }

        // Option D: Final fallback - broad search for patterns like Invoice No.: 12345
        for (const el of elements) {
            const match = el.innerText.match(/Invoice No\.\s*[:\-]?\s*([A-Z0-9\-_]+)/i);
            if (match && match[1]) return match[1];
        }

        return "Unknown";
      });
    } else {
      console.warn("    ⚠ Could not find frame 'ilboinnerframe' for extraction.");
    }
  } catch (err) {
    console.warn("    ⚠ Failed to extract Invoice No: " + err.message);
  }

  return invoiceNo;
}

/**
 * Triggers the "View Expense Invoice" button click
 */
async function triggerViewExpenseInvoice(page) {
  console.log('  🖱 Clicking "View Expense Invoice" button...');
  const frame = page.frameLocator("#ilboinnerframe");
  
  const viewBtn = frame.locator('a, button, .x-btn-text').filter({ hasText: /^View Expense Invoice$/i }).first();
  
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

  let clicked = await tryClick(viewBtn, "Exact Button");
  
  if (!clicked) {
    const looseBtn = frame.getByText('View Expense Invoice', { exact: false }).first();
    clicked = await tryClick(looseBtn, "Flexible Button");
  }

  if (!clicked) {
    await frame.locator('a, button, .x-btn-text').filter({ hasText: /View Expense Invoice/i }).first().evaluate(el => el.click()).catch(e => console.error("    ⚠ JS Click failed: " + e.message));
  }

  // Smart wait for detail screen element instead of 10s
  await frame.locator('text="Invoice Header"').first().waitFor({ state: "attached", timeout: 15000 }).catch(() => {
    console.log("  ⏳ Detail screen still loading...");
  });
  await page.waitForTimeout(1000); 
}

/**
 * Navigates the detail screen and downloads the Excel file
 */
async function downloadExpenseDetail(page, invoiceNo = "Unknown") {
  console.log(`\n📥 Preparing to download Expense Detail Excel (Invoice: ${invoiceNo})...`);
  
  try {
    // Ensure exports directory exists
    mkdirSync("output/exports", { recursive: true });

    // 1. Give the detail screen a moment to fully load its data
    await page.waitForTimeout(1500); // Reduced from 5s

    // 2. Perform the 14-tab sequence as requested
    console.log("  ⌨ Pressing Tab 14 times...");
    for (let i = 0; i < 14; i++) {
      await page.keyboard.press("Tab");
      await page.waitForTimeout(100); // Faster tabs
    }

    // 3. Trigger the Export (Shift + E)
    console.log("  ⌨ Triggering Excel Export (Shift + E)...");
    
    // Set up the listener for the download event
    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });
    
    // Execute the keyboard shortcut
    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');
    
    const download = await downloadPromise;
    const suggestedName = download.suggestedFilename();
    console.log(`  🔍 Server suggested name: "${suggestedName}"`);
    
    // Determine the extension (fallback to .xls if none found)
    const ext = suggestedName.includes('.') ? suggestedName.split('.').pop() : 'xls';
    
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/List of Direct Expenses/${dateFolder}/${invoiceNo}`;
    mkdirSync(exportDir, { recursive: true });
    
    const fileName = `${exportDir}/Expense Info - (${invoiceNo}).${ext}`;
    
    await download.saveAs(fileName);
    console.log(`  ✅ Excel download successful!`);
    console.log(`  📁 Saved and renamed properly in your project → ${fileName}`);
    console.log(`  💡 (Note: The browser's internal download history may still show temporary UUID names).`);

  } catch (error) {
    console.error(`  ❌ Excel download failed: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-excel-download.png" });
  }
}

/**
 * Clicks the "Go back" button to return to the search results grid
 */
async function goBackFromDetail(page) {
  console.log("\n🔙 Navigating back to search results...");
  
  try {
    // Note: The "Go back" button is typically in the dashboard toolbar ABOVE the iframe
    const backBtn = page.locator('#back button, button[qtip="Go back"], .x-tool-back').first();
    
    await backBtn.waitFor({ state: "attached", timeout: 10000 });
    await backBtn.scrollIntoViewIfNeeded().catch(() => {});
    
    // Multi-strategy click
    try {
      await backBtn.click({ force: true, timeout: 5000 });
      console.log("  ✔ Clicked 'Go back' (forced).");
    } catch (e) {
      console.warn("  ⚠ Direct 'Go back' click failed. Trying flexible search or JS dispatch...");
      // Fallback to text filter if qtip/id failed
      const altBtn = page.locator('button, a, .x-btn-text').filter({ hasText: /Go back/i }).first();
      await altBtn.evaluate(el => el.click()).catch(err => console.error("    ❌ JS Click failed: " + err.message));
    }

    // Wait for the previous screen to reload
    console.log("  ⏳ Waiting for previous screen to load...");
    const frame = page.frameLocator("#ilboinnerframe");
    await frame.locator('.x-grid3-row').first().waitFor({ state: "visible", timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(1000);
    
    const screenshotPath = "screenshots/07-returned-to-grid.png";
    // await page.screenshot({ path: screenshotPath, fullPage: true });
    // console.log(`  📸 Landed back on grid → ${screenshotPath}`);

  } catch (error) {
    console.error(`  ❌ Failed to navigate back: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-go-back.png" });
  }
}

/**
 * Navigates to the "Accounting Information" view within the invoice detail screen
 */
async function goToAccountingInformation(page) {
  console.log("\n📑 Navigating to 'Accounting Information'...");
  const frame = page.frameLocator("#ilboinnerframe");

  try {
    // Try a more robust text-based selector first
    console.log("  🔍 Searching for 'Accounting Information' link...");
    const link = frame.locator('text="Accounting Information"').first();
    
    // Check if it's visible/attached with a longer wait
    await link.waitFor({ state: "attached", timeout: 15000 });
    await link.scrollIntoViewIfNeeded().catch(() => {});
    
    console.log("  🖱 Clicking 'Accounting Information' link...");
    try {
      // Try to click the specific element found
      await link.click({ force: true, timeout: 5000 });
      console.log("  ✔ Clicked via standard click.");
    } catch (e) {
      console.warn("  ⚠ Standard click failed, trying JS dispatch on all matching elements...");
      // Fallback: try JS click on any element containing the text
      const success = await frame.locator('*:has-text("Accounting Information")').evaluateAll(els => {
        const target = els.find(el => el.textContent.trim().toLowerCase() === 'accounting information');
        if (target) {
          target.click();
          return true;
        }
        return false;
      });
      
      if (success) {
        console.log("  ✔ Clicked via JS dispatch fallback.");
      } else {
        throw new Error("Could not find or click the 'Accounting Information' element.");
      }
    }

    // Wait for the information to render
    console.log("  ⏳ Waiting for accounting info to load...");
    // Smart wait for labels
    await frame.locator('text="Journal Details", text="Debit/Credit"').first().waitFor({ state: "attached", timeout: 10000 }).catch(() => {});
    await page.waitForTimeout(1000);
    
    const screenshotPath = "screenshots/06b-accounting-information.png";
    // await page.screenshot({ path: screenshotPath, fullPage: true });
    // console.log(`  📸 Accounting Information captured → ${screenshotPath}`);

  } catch (error) {
    console.error(`  ❌ Failed to reach Accounting Information: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-accounting-info.png" });
    
    // Last resort: debugging - log what we see in the frame
    // console.log("  📸 Debug: Taking screenshot of current state for analysis...");
    // await page.screenshot({ path: "screenshots/debug-failed-navigation.png" });
  }
}

/**
 * Downloads the Excel file from the Accounting Information view
 */
async function downloadAccountingExcel(page, invoiceNo = "Unknown") {
  console.log(`\n📊 Preparing to download Accounting Excel (Invoice: ${invoiceNo})...`);
  
  try {
    // Ensure exports directory exists
    mkdirSync("output/exports", { recursive: true });

    // 1. Focus the grid area. The user suggested "Tab and Shift + E".
    // Using 1 Tab press as requested.
    console.log("  ⌨ Pressing Tab to focus the grid area...");
    for (let i = 0; i < 1; i++) {
        await page.keyboard.press("Tab");
        await page.waitForTimeout(100);
    }

    // 2. Trigger the Export (Shift + E)
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
    const exportDir = `output/List of Direct Expenses/${dateFolder}/${invoiceNo}`;
    mkdirSync(exportDir, { recursive: true });

    const fileName = `${exportDir}/Accounting Info - (${invoiceNo}).xlsm`;
    
    await download.saveAs(fileName);
    console.log(`  ✅ Accounting Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Accounting Excel download failed: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-accounting-download.png" });
  }
}

/**
 * Navigates to the "View Tax Calculation - Summary" screen.
 * This follows the exact same pattern as 'Accounting Information'.
 * Returns true if the screen appears to have opened successfully.
 */
async function goToTaxCalculationSummary(page) {
  console.log("\n📑 Navigating to 'View Tax Calculation - Summary'...");
  const frame = page.frameLocator("#ilboinnerframe");
  
  try {
    // 1. Give the page a moment to stabilize after "Go Back"
    await page.waitForTimeout(3000);

    // 2. Locate and click the link (Same approach as Accounting Information)
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

    // 3. Verify screen loaded (Using same timeout approach as Accounting Information)
    console.log("  ⏳ Waiting for View Tax Calculation - Summary info to load...");
    await page.waitForTimeout(5000); 
    
    // Capture screenshot for confirmation
    // await page.screenshot({ path: "screenshots/06c-tax-calculation-summary.png", fullPage: true });
    // console.log("  ✅ View Tax Calculation - Summary captured.");
    return true;

  } catch (error) {
    console.error(`  ❌ Failed to reach View Tax Calculation - Summary: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-tax-summary-nav.png" });
    return false;
  }
}

/**
 * Handles the Shift+E export for Tax Summary Excel
 */
async function downloadTaxSummaryExcel(page, invoiceNo = "Unknown") {
  console.log(`\n📊 Preparing to download Tax Summary Excel (Invoice: ${invoiceNo})...`);
  const frame = page.frameLocator("#ilboinnerframe");
  
  try {
    // 1. Give the page a moment to focus
    await page.waitForTimeout(1000);

    // 2. Perform tabbing to focus the grid area
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
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
    const fileName = `output/exports/Tax Summary - (${invoiceNo}).xls`;
    
    await download.saveAs(fileName);
    console.log(`  ✅ Tax Summary Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Tax Summary Excel download failed: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-tax-summary-download.png" });
  }
}

/**
 * Navigates to the "View Tax Calculation - Details" screen from the Summary screen.
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
      
      console.log("  ⏳ Waiting for View Tax Calculation - Details info to load...");
      await page.waitForTimeout(5000); 
      
      // await page.screenshot({ path: "screenshots/06d-tax-calculation-details.png", fullPage: true });
      // console.log("  ✅ View Tax Calculation - Details captured.");
      return true;
    } else {
      console.warn("  ⚠ 'View Tax Calculation - Details' link was not visible.");
      return false;
    }

  } catch (error) {
    console.error(`  ❌ Failed to reach View Tax Calculation - Details: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-tax-details-nav.png" });
    return false;
  }
}

/**
 * Handles the Shift+E export for Tax Details Excel
 */
async function downloadTaxDetailsExcel(page, invoiceNo, taxType, lineNo) {
  console.log(`\n📊 Preparing to download Tax Details Excel (${taxType} Line ${lineNo}) (Invoice: ${invoiceNo})...`);
  
  try {
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/List of Direct Expenses/${dateFolder}/${invoiceNo}`;
    mkdirSync(exportDir, { recursive: true });

    // 1. Give the page a moment to focus
    await page.waitForTimeout(1000);

    // 2. Perform tabbing to focus the grid area
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
    const fileName = `${exportDir}/Tax Calculation Details (${taxType} Line${lineNo}) - (${invoiceNo}).xls`;
    
    await download.saveAs(fileName);
    console.log(`  ✅ Tax Details Excel download successful!`);
    console.log(`  📁 Saved as → ${fileName}`);

  } catch (error) {
    console.error(`  ❌ Tax Details Excel download failed: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-tax-details-download.png" });
  }
}

/**
 * Changes the "Tax Type" selection in the Tax Details screen.
 */
async function changeTaxType(page, targetType) {
  console.log(`\n🔄 Changing Tax Type to '${targetType}'...`);
  const frame = page.frameLocator("#ilboinnerframe");
  
  try {
    // 1. Locate the Tax Type dropdown field
    const taxTypeField = frame.locator('xpath=//label[contains(text(), "Tax Type")]/following::input[1] | //td[contains(text(), "Tax Type")]/following::input[1]').first();
    
    await taxTypeField.waitFor({ state: "visible", timeout: 10000 });
    await taxTypeField.click();
    await page.waitForTimeout(500);

    // 2. Select the target (e.g., VAT) from the expanded list
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

    // 3. Wait for the grid to refresh its data
    console.log(`  ⏳ Waiting for data to refresh...`);
    await page.waitForTimeout(2500); 
    
    // await page.screenshot({ path: `screenshots/06e-tax-details-${targetType.toLowerCase()}.png`, fullPage: true });
    // console.log(`  ✅ Tax Type change completed.`);
    return true;

  } catch (error) {
    console.error(`  ❌ Failed to change Tax Type: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-tax-type-change.png" });
    return false;
  }
}

/**
 * Fetches all available options from the "Ref. Doc. Line No." dropdown.
 * Uses multiple strategies to discover all options including scrolling
 * through the virtualized ExtJS combo list.
 */
async function getRefDocLineOptions(page) {
  console.log("  🔍 Detecting 'Ref. Doc. Line No.' options...");
  const frame = page.frameLocator("#ilboinnerframe");
  const actualFrame = page.frames().find(f => f.name() === 'ilboinnerframe' || f.url().includes('rvw_blank.aspx'));
  
  try {
    const field = frame.locator('xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | //td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]').first();
    
    await field.waitFor({ state: "visible", timeout: 5000 });
    
    // Strategy 1: Try to read the ExtJS store data directly (most reliable)
    let options = [];
    if (actualFrame) {
      try {
        options = await actualFrame.evaluate(() => {
          // Find the input element for Ref. Doc. Line No.
          const labels = Array.from(document.querySelectorAll('label, td, span'));
          const labelEl = labels.find(el => el.innerText.trim().includes('Ref. Doc. Line No.'));
          if (labelEl) {
            const input = labelEl.closest('tr')?.querySelector('input') || 
                          labelEl.parentElement?.querySelector('input') ||
                          labelEl.closest('td')?.nextElementSibling?.querySelector('input');
            if (input && input.id) {
              // Try to get the ExtJS component and its store
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

    // Strategy 2: Open dropdown and scroll through it to collect all items
    console.log("    ℹ Falling back to dropdown scroll enumeration...");
    await field.click();
    await page.waitForTimeout(1000);

    const allOptions = new Set();
    
    // Collect initially visible items
    let visibleItems = await page.locator('.x-combo-list-item').evaluateAll(els => 
      els.map(el => el.innerText.trim()).filter(txt => txt.length > 0)
    );
    visibleItems.forEach(item => allOptions.add(item));

    // Scroll through the combo list to find more items
    const comboList = page.locator('.x-combo-list-inner').first();
    const isComboVisible = await comboList.isVisible({ timeout: 2000 }).catch(() => false);
    
    if (isComboVisible) {
      let previousSize = 0;
      let scrollAttempts = 0;
      const maxScrollAttempts = 50; // Safety limit
      
      while (scrollAttempts < maxScrollAttempts) {
        // Scroll down in the combo list
        await comboList.evaluate(el => el.scrollTop += 100);
        await page.waitForTimeout(300);
        
        // Collect newly visible items
        visibleItems = await page.locator('.x-combo-list-item').evaluateAll(els => 
          els.map(el => el.innerText.trim()).filter(txt => txt.length > 0)
        );
        visibleItems.forEach(item => allOptions.add(item));
        
        // Check if we found new items
        if (allOptions.size === previousSize) {
          scrollAttempts++;
          if (scrollAttempts > 3) break; // No new items after 3 scroll attempts
        } else {
          scrollAttempts = 0; // Reset counter when we find new items
        }
        previousSize = allOptions.size;
      }
    }

    // Press Escape to close the dropdown
    await page.keyboard.press("Escape");
    await page.waitForTimeout(500);

    let uniqueOptions = [...allOptions];
    
    // Sort numerically if all options are numbers
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

/**
 * Selects a specific "Ref. Doc. Line No." from the dropdown.
 * Handles scrolling through virtualized ExtJS combo lists.
 */
async function selectRefDocLine(page, lineNo) {
  console.log(`  🔢 Selecting Ref. Doc. Line No: ${lineNo}...`);
  const frame = page.frameLocator("#ilboinnerframe");
  
  try {
    const field = frame.locator('xpath=//label[contains(text(), "Ref. Doc. Line No.")]/following::input[1] | //td[contains(text(), "Ref. Doc. Line No.")]/following::input[1]').first();
    
    await field.waitFor({ state: "visible", timeout: 5000 });
    await field.click();
    await page.waitForTimeout(500);

    // Try to find and click the exact option (with scrolling support)
    let clicked = false;
    
    // First, try to find the option in currently visible items
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

    // If not found, scroll through the combo list to find it
    if (!clicked) {
      const comboList = page.locator('.x-combo-list-inner').first();
      const isComboVisible = await comboList.isVisible({ timeout: 2000 }).catch(() => false);
      
      if (isComboVisible) {
        // Reset scroll to top first
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

    // Fallback: use keyboard input if scrolling didn't work
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
 * Automates the signout process
 */
async function signOut(page) {
  console.log("\n🔒 Attempting to Sign Out...");
  
  try {
    // 1. Trigger the User Menu (Cris V. Ladines)
    // ID from HTML: "userinfo"
    console.log('  🖱 Clicking User Menu (#userinfo)...');
    const userMenu = page.locator('#userinfo, button:has-text("Cris V. Ladines")').first();
    
    if (await userMenu.isVisible({ timeout: 5000 }).catch(() => false)) {
      await userMenu.click();
      await page.waitForTimeout(500); // Faster menu opening
    } else {
      console.warn("    ⚠ User menu not found, trying direct signout click...");
    }

    // 2. Click the Signout button
    // ID from HTML: "signout"
    console.log('  🖱 Clicking Signout (#signout)...');
    const signoutBtn = page.locator('#signout, [qtip="Signout"], .x-menu-item-text:text("Signout")').first();
    
    if (await signoutBtn.isVisible({ timeout: 5000 }).catch(() => false)) {
      await signoutBtn.click();
      console.log("  ✔ Signout button clicked.");
    } else {
      // Fallback to JS click if hidden by offsets
      await page.evaluate(() => {
        const btn = document.getElementById('signout') || 
                    Array.from(document.querySelectorAll('button')).find(b => b.innerText.includes('Signout'));
        if (btn) btn.click();
      }).catch(() => {});
    }

    await page.waitForTimeout(5000);
    // await page.screenshot({ path: "screenshots/08-signed-out.png" });
    // console.log("  📸 Sign out state captured → screenshots/08-signed-out.png");

  } catch (error) {
    console.error(`  ❌ Sign out failed: ${error.message}`);
    // await page.screenshot({ path: "screenshots/error-signout.png" });
  }
}

/**
 * Utility to convert DD/MM/YYYY to MMM D YYYY (e.g. 16/04/2026 -> Apr 16 2026)
 */
function formatRamcoDate(rawDate) {
  if (!rawDate) return "UnknownDate";
  // Remove any non-alphanumeric or slash characters
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
 * Downloads the initial search results grid before row-by-row processing starts.
 */
async function downloadInitialSearchResults(page) {
  console.log("\n📁 Processing Initial Search Results Download...");
  const frame = page.frameLocator("#ilboinnerframe");
  
  try {
    // 1. Get Invoice Date From to use as filename
    const dateField = frame.locator('xpath=//label[contains(text(), "Invoice Date From")]/following::input[1] | //td[contains(text(), "Invoice Date From")]/following::input[1]').first();
    const rawDate = await dateField.inputValue();
    console.log(`  📅 Detected Invoice Date: ${rawDate}`);
    
    const formattedDate = formatRamcoDate(rawDate);
    console.log(`  📂 Filename: ${formattedDate}.xls`);

    // 2. Ensure hierarchical directory exists
    const dateFolder = new Date().toISOString().split('T')[0];
    const exportDir = `output/List of Direct Expenses/${dateFolder}`;
    mkdirSync(exportDir, { recursive: true });

    // 3. Focus the grid area for download
    console.log("  ⌨ Focusing grid for initial download...");
    const searchBtn = frame.locator('button, .x-btn-text').filter({ hasText: /^Search$/i }).last();
    if (await searchBtn.isVisible({ timeout: 5000 })) {
      await searchBtn.focus();
      await page.keyboard.press("Tab"); // 1 tab as requested
      await page.waitForTimeout(200);
    }

    console.log("  ⌨ Triggering Excel Export (Shift + E)...");
    const downloadPromise = page.waitForEvent('download', { timeout: 60000 });
    await page.keyboard.down('Shift');
    await page.keyboard.press('E');
    await page.keyboard.up('Shift');
    
    const download = await downloadPromise;
    const fileName = `${exportDir}/${formattedDate}.xls`;
    await download.saveAs(fileName);
    
    console.log(`  ✅ Initial results saved to: ${fileName}`);
    return true;

  } catch (error) {
    console.error(`  ❌ Initial results download failed: ${error.message}`);
    return false;
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
    await goToLandingPage(page);

    const allResults = [];
    for (const menuPath of MENU_PATHS) {
      const result = await navigateMenu(page, menuPath);
      
      // Perform Search logic if we reached the Invoice module
      if (menuPath.join(" ").includes("View Invoice")) {
        const stats = await performInvoiceSearch(page);
        
        // --- NEW: Download Initial Results before starting row loop ---
        await downloadInitialSearchResults(page);

        const totalRows = stats.total || 65; 
        console.log(`\n🚀 Starting continuous processing for all ${totalRows} rows...`);

        const frame = page.frameLocator("#ilboinnerframe");

        for (let processedCount = 1; processedCount <= totalRows; processedCount++) {
          process.stdout.write(`\n📍 [Row ${processedCount}/${totalRows}] Navigation... `);
          
          // 1. Navigation to the row
          if (processedCount === 1) {
            await selectFirstRow(page);
          } else {
            await selectNextRow(page, processedCount);
          }

          // 2. Set Row-based Name (per user request)
          const invNo = `Row-${processedCount}`;
          console.log(`\n📂 --- PROCESSING ROW ${processedCount} OF ${totalRows} ---`);

          // 3. Trigger Detail View
          const actualInvNo = await triggerViewInvoice(page);
          console.log(`  🔍 Extracted Invoice No: ${actualInvNo}`);
          
          // Update the naming reference to use the real invoice number if found
          const currentInvNo = actualInvNo !== "Unknown" ? actualInvNo : invNo;

          // 4. Downloads
          await downloadExpenseDetail(page, currentInvNo);
          await goToAccountingInformation(page);
          await downloadAccountingExcel(page, currentInvNo);
          
          // 5. Navigate to 'View Tax Calculation - Summary' (Same approach as Accounting Information)
          console.log("  🔙 Returning from Accounting to Expense Invoice...");
          await goBackFromDetail(page); 
 
          const taxScreenOpened = await goToTaxCalculationSummary(page);
 
          if (taxScreenOpened) {
            // await downloadTaxSummaryExcel(page, currentInvNo);
            const detailsOpened = await goToTaxCalculationDetails(page);
            
            if (detailsOpened) {
              // --- NESTED ITERATION: Default Tax Type (EWT) ---
              console.log("  📂 Processing Default Tax Type (EWT) lines...");
              const defaultLines = await getRefDocLineOptions(page);
              for (const line of defaultLines) {
                const selected = await selectRefDocLine(page, line);
                if (selected) {
                  await downloadTaxDetailsExcel(page, currentInvNo, "EWT", line);
                }
              }
              
              // --- NESTED ITERATION: VAT Tax Type ---
              // Go back from Tax Details to Tax Summary first, then re-enter
              // to get a fresh dropdown state for VAT
              console.log("  🔙 Returning from Tax Details (EWT done) to Tax Summary...");
              await goBackFromDetail(page);
              
              // Re-enter Tax Details fresh for VAT processing
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
                      await downloadTaxDetailsExcel(page, currentInvNo, "VAT", line);
                    }
                  }
                }

                console.log("  🔙 Returning from Tax Details (VAT done) to Tax Summary...");
                await goBackFromDetail(page);
              }
            }

            // 6. Return from Tax Summary back to Expense Invoice
            console.log("  🔙 Returning from Tax Summary to Expense Invoice...");
            await goBackFromDetail(page); 
          } else {
            console.log("  ⏭ Skipping Tax Details View.");
          }
          
          // 7. Return to Grid (Final Back from Expense Invoice)
          console.log("  🔙 Navigating back to search results...");
          await goBackFromDetail(page); 
          
          console.log(`  ✅ Row ${processedCount} completed.`);
          await page.waitForTimeout(500); 
        }
        
        // Final extraction if any content is visible on result page
        const innerFrame = page.frame({ name: "ilboinnerframe" });
        if (innerFrame) {
          result.tables = await extractTableData(innerFrame);
          result.textSample = await innerFrame.evaluate(() => document.body.innerText.trim().slice(0, 500));
        }
      }
      
      allResults.push(result);
    }

    writeFileSync("output/extracted-data.json", JSON.stringify(allResults, null, 2));
    console.log("\n💾 Saved → output/extracted-data.json");

    const totalSeconds = Math.floor((Date.now() - startTime) / 1000);
    const minutes = Math.floor(totalSeconds / 60);
    const seconds = totalSeconds % 60;
    
    console.log(`\n⏱  Total Process Runtime: ${minutes > 0 ? `${minutes}m ` : ""}${seconds}s`);
    console.log("✅ All done!");
    
    // Use the 10s countdown instead of blocking prompt
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
    // await page.screenshot({ path: "screenshots/error.png" }).catch(() => {});
    await waitForEnter("⚠ Error occurred. Press ENTER to close...");
  } finally {
    await browser.close();
    console.log("🔒 Browser closed.");
  }
}

main();
