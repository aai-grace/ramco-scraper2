# Ramco VirtualWorks Scraper

A Node.js Playwright scraper for automating navigation and data extraction on Ramco VirtualWorks.

## Setup

```bash
# 1. Install dependencies
npm install

# 2. Install Playwright browser
npx playwright install chromium

# 3. Add your credentials in index.js
#    CONFIG.username = "your_username"
#    CONFIG.password = "your_password"

# 4. Run the scraper
node index.js
```

## How it works

### First Run (Discovery Mode)
On first run, leave `MENU_PATHS = []` as-is. The script will:
1. Log in to Ramco VirtualWorks
2. Dump all clickable elements to `output/discovered-elements.json`
3. Take screenshots to `screenshots/`

Open `output/discovered-elements.json` to see all menu item labels, then add them to `MENU_PATHS`.

### Adding Menu Paths
Edit `MENU_PATHS` in `index.js`:
```js
const MENU_PATHS = [
  ["Maintenance", "Work Orders"],
  ["Inventory", "Parts", "Search"],
];
```
Each array is a click sequence: top menu → submenu → page.

### Output
- `screenshots/` — PNG screenshots at each step
- `output/discovered-elements.json` — all clickable elements found after login
- `output/extracted-data.json` — tables and text extracted from navigated pages

## Tips
- Keep `headless: false` while building so you can watch the browser
- Increase `slowMo` (e.g. 500–1000) if the app is slow to respond
- Check `screenshots/error.png` if something goes wrong
