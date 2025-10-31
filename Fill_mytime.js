function mapProductName(productL1) {
  const normalized = (productL1 || '').trim().toLowerCase();

  const mappings = {
    'design compiler': 'Design Compiler / DC NXT',
    'design compiler nxt': 'Design Compiler / DC NXT',
    // Add more mappings as needed:
    // 'primetime': 'PrimeTime',
    // 'verdi': 'Verdi / Siloti',
  };

  return mappings[normalized] || productL1;
}

function mapCustomerName(customerName) {
  const normalized = (customerName || '').trim().toLowerCase();

  const mappings = {
    // === Examples; extend as needed ===
    'proxelera': 'Proxelera',
    // Add more mappings as needed:
    // 'abc corp': 'ABC',
    // 'abc corporation': 'ABC',
  };

  // Fallback to original value if no mapping found
  return mappings[normalized] || (customerName || '').trim();
}



function loadExcelData(callback) {
  const script = document.createElement('script');
  script.src = 'https://cdn.sheetjs.com/xlsx-0.20.0/package/dist/xlsx.full.min.js';
  script.onload = () => {
    console.log('SheetJS loaded');

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xls';
    input.style.display = 'none';

    input.addEventListener('change', (e) => {
      const file = e.target.files[0];
      const reader = new FileReader();

      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const excelData = XLSX.utils.sheet_to_json(worksheet);

        console.log('✅ Excel data loaded:', excelData);

        // ✅ Use the existing mapProductName function
        window.myExcelData = excelData.map(row => ({
          ...row,
          'Product L1': mapProductName(row['Product L1']),
          'Logo': mapCustomerName(row['Logo'])
        }));

        if (typeof callback === 'function') {
          callback(); // Start filling process
        }
      };

      reader.readAsArrayBuffer(file);
    });

    document.body.appendChild(input);
    input.click();
  };

  document.head.appendChild(script);
}

// Reusable helper
const delay = ms => new Promise(r => setTimeout(r, ms));

async function waitForElement(selector, timeout = 5000) {
  const start = performance.now();
  while (performance.now() - start < timeout) {
    const el = document.querySelector(selector);
    if (el && el.offsetParent !== null) return el; // visible and interactable
    await new Promise(r => setTimeout(r, 100));
  }
  return null;
}


async function pickDropdownOption(caretEl, optionText, timing = {}) {
  const { beforeOpen = 1000, afterOpen = 500, afterPick = 500 } = timing;
  if (!caretEl) return false;
  const normalize = s => (s || '').replace(/\u00a0/g, ' ').trim();

  await delay(beforeOpen);
  caretEl.click();
  await delay(afterOpen);

  const option = [...document.querySelectorAll('.ms-Dropdown-optionText')]
    .find(el => el.offsetParent !== null && normalize(el.textContent) === normalize(optionText));
  if (!option)  return false;

  option.click();
  await delay(afterPick);
  return true;
}

async function waitForText(text, timeout = 10000) {
  const start = performance.now();
  while (performance.now() - start < timeout) {
    if ([...document.querySelectorAll('*')]
      .some(e => e.textContent.trim() === text && e.offsetParent !== null)) {
      return true;
    }
    await new Promise(r => setTimeout(r, 150));
  }
  return false;
}

/**
 * Types text into an input, then loops until it can click the exact suggestion.
 * If "No results found" is visible, it backspaces 1 character, waits for options
 * to appear, then retypes that character before checking again.
 * This function NEVER times out; it keeps trying until success.
 *
 * @param {HTMLInputElement} inputEl - the input to type into
 * @param {string} targetText        - exact visible suggestion text to click
 * @param {object} [opts]
 * @param {number} [opts.perCharDelay=90]   - delay between typed chars
 * @param {number} [opts.checkInterval=140] - polling interval while waiting
 * @param {number} [opts.afterTypeWait=220] - pause after finishing typing
 * @param {boolean} [opts.debug=false]      - console.debug tracing
 * @returns {Promise<void>}
 */
async function selectFromSuggestionsForever(inputEl, targetText, opts = {}) {
  const {
    perCharDelay  = 90,
    checkInterval = 140,
    afterTypeWait = 220,
    debug         = false
  } = opts;

  if (typeof delay !== 'function') {
    throw new Error('selectFromSuggestionsForever: `delay(ms)` helper must be defined before this function.');
  }

  const log = (...args) => { if (debug) console.debug('[selectFromSuggestionsForever]', ...args); };
  const normalize = s => (s ?? '').replace(/\u00a0/g, ' ').trim();
  const eq = (a, b) => normalize(a) === normalize(b);
  const isVisible = el => !!(el && el.offsetParent !== null);

  const visiblePersonaItems = () =>
    Array.from(document.querySelectorAll('div.ms-Persona-primaryText')).filter(isVisible);

  const findSuggestionByText = (text) =>
    visiblePersonaItems().find(el => eq(el.textContent, text)) || null;

  const isNoResultsVisible = () => {
    const el = document.querySelector('.ms-Suggestions-none');
    return !!(isVisible(el) && /no results found/i.test(el.textContent || ''));
  };

  const clickSuggestion = (textEl) => {
    // Click the closest interactive container, not just the text node
    const clickable =
      textEl.closest('[role="option"], .ms-Suggestions-item, .ms-Persona, .ms-PeoplePicker-personaContent') || textEl;
    clickable.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    clickable.click();
  };

  const clearInput = () => {
    inputEl.focus();
    // Simulate selecting all and backspace for frameworks that track selection
    inputEl.setSelectionRange(0, inputEl.value.length);
    inputEl.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace' }));
    inputEl.value = '';
    inputEl.dispatchEvent(new Event('input', { bubbles: true }));
  };

  const typeText = async (text) => {
    inputEl.focus();
    clearInput();
    for (const ch of text) {
      inputEl.dispatchEvent(new KeyboardEvent('keydown', { key: ch }));
      inputEl.value += ch;
      inputEl.dispatchEvent(new Event('input', { bubbles: true }));
      await delay(perCharDelay);
    }
    await delay(afterTypeWait);
  };

  if (!normalize(targetText)) {
    throw new Error('selectFromSuggestionsForever: targetText is empty.');
  }

  // Initial type of the full target text
  await typeText(targetText);
  log('Typed initial text:', targetText);

  // Loop indefinitely until we can click the right suggestion
  while (true) {
    // 1) If exact suggestion is visible, click and exit
    const match = findSuggestionByText(targetText);
    if (match) {
      clickSuggestion(match);
      log('Clicked exact match.');
      await delay(120);
      return;
    }

    // 2) If "No results found" is visible: backspace 1 char, then WAIT for options to appear, then retype that char
    if (isNoResultsVisible()) {
      const current = inputEl.value || '';
      log('No results visible. Current value:', current);

      // If the input drifted from target, retype the full target
      if (!eq(current, targetText)) {
        log('Input drift detected. Retyping targetText.');
        await typeText(targetText);
        continue;
      }

      let lastChar = '';
      if (current.length > 0) {
        lastChar = current.slice(-1);

        // Backspace one char
        inputEl.focus();
        inputEl.dispatchEvent(new KeyboardEvent('keydown', { key: 'Backspace' }));
        inputEl.value = current.slice(0, -1);
        inputEl.dispatchEvent(new Event('input', { bubbles: true }));
        log('Backspaced one char:', lastChar);
      } else {
        // Nothing to backspace → retype full text
        log('Nothing to backspace. Retyping full targetText.');
        await typeText(targetText);
        continue;
      }

      // 🔻 WAIT (no timeout) for options to appear or the exact match to show up
      while (true) {
        const direct = findSuggestionByText(targetText);
        if (direct) {
          clickSuggestion(direct);
          log('Clicked match during wait.');
          await delay(120);
          return;
        }

        const anyOptions = visiblePersonaItems().length > 0;
        const stillNoRes = isNoResultsVisible();
        if (anyOptions && !stillNoRes) {
          log('Options appeared after backspace.');
          break; // proceed to retype last char
        }

        await delay(checkInterval);
      }

      // Retype that one char now that options are present
      inputEl.dispatchEvent(new KeyboardEvent('keydown', { key: lastChar }));
      inputEl.value += lastChar;
      inputEl.dispatchEvent(new Event('input', { bubbles: true }));
      log('Retyped last char:', lastChar);
      await delay(afterTypeWait);

      // Continue → outer loop will re-check for the target suggestion
      continue;
    }

    // 3) Neither exact match nor "No results" → keep the input consistent and poll
    if (!eq(inputEl.value || '', targetText)) {
      log('Input changed unexpectedly. Restoring targetText.');
      await typeText(targetText);
    } else {
      await delay(checkInterval);
    }
  }
}

function fillCommentBox(text) {
  if (!text || text.trim() === '') return; // Skip if text is undefined or blank

  const el = document.querySelector('textarea.ms-TextField-field');
  if (el) {
    el.focus();
    el.value = `Support Case: ${text}`;
    el.dispatchEvent(new Event('input', { bubbles: true }));
  }
}

async function waitForSaveAndCloseToDisappear() {
  while (true) {
    const button = Array.from(document.querySelectorAll('span.ms-Button-label'))
      .find(el => el.textContent.trim() === "Save and Close");

    if (!button || button.offsetParent === null) {
      return true; // Button is gone or hidden
    }

    await new Promise(r => setTimeout(r, 200));
  }
}

// Fill week values in two passes to avoid row disappearing:
window.fillWeekFromExcel = async function (dataRows) {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm = s => (s || '').replace(/\u00a0/g, ' ').trim().toLowerCase();
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];

  const isZeroish = (v) => {
    if (v == null) return true;
    const s = String(v).replace(/%/g, '').replace(/[^\d.,-]/g, '').replace(',', '.').trim();
    if (s === '') return true;
    const n = Number(s);
    return !isFinite(n) ? false : n === 0;
  };

  const targetText = (v) => (v == null ? '' : String(v));

  for (const r of (dataRows || [])) {
    const logo     = (r['Logo'] || '').trim();
    const prod     = (r['Product L1'] || '').trim();
    const category = (r['Category'] || '').trim();
    const activity = (r['Activity'] || '').trim();

    let rowEl;

    if (norm(category) === 'administration') {
      // Match only Category and Activity
      rowEl = [...document.querySelectorAll('.ms-DetailsRow-fields[data-automationid="DetailsRowFields"]')]
        .find(el =>
          norm(el.querySelector('[data-automation-key="headerColumn0"]')?.textContent) === norm(category) &&
          norm(el.querySelector('[data-automation-key="headerColumn1"]')?.textContent) === norm(activity)
        );
    } else {
      // Match all four: Logo, Product L1, Category, Activity
      rowEl = [...document.querySelectorAll('.ms-DetailsRow-fields[data-automationid="DetailsRowFields"]')]
        .find(el =>
          norm(el.querySelector('[data-automation-key="headerColumn2"]')?.textContent) === norm(logo) &&
          norm(el.querySelector('[data-automation-key="headerColumn3"]')?.textContent) === norm(prod) &&
          norm(el.querySelector('[data-automation-key="headerColumn0"]')?.textContent) === norm(category) &&
          norm(el.querySelector('[data-automation-key="headerColumn1"]')?.textContent) === norm(activity)
        );
    }

    if (!rowEl) continue;

    const vals = days.map(d => r[d] ?? r['%work'] ?? r['% work'] ?? '');

    // --- PASS 1: Fill non-zero values ---
    for (let i = 0; i < 5; i++) {
      if (isZeroish(vals[i])) continue;
      const input = rowEl.querySelector(`[data-automation-key="dataColumn${i}"] input`);
      if (!input) continue;

      input.focus();
      input.value = targetText(vals[i]);
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await sleep(40);
    }

    const hasAnyNonZero = vals.some(v => !isZeroish(v));
    if (!hasAnyNonZero) continue;

    // --- PASS 2: Clear zero values ---
    for (let i = 0; i < 5; i++) {
      if (!isZeroish(vals[i])) continue;
      const input = rowEl.querySelector(`[data-automation-key="dataColumn${i}"] input`);
      if (!input) continue;

      input.focus();
      input.value = '';
      input.dispatchEvent(new Event('input', { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await sleep(30);
    }

    await sleep(80);
  }
};

// Start everything (process sequentially)
loadExcelData(async () => {
  for (const row of window.myExcelData) {
    const logo = mapCustomerName(row["Logo"]);
    const tool = mapProductName(row["Product L1"]);
    const comment = row["Case Number"];

    const clickNew = () => document.querySelector('.ms-Button-label')?.closest('button')?.click();

    // Keep trying until "New Time Entry" is visible
    let opened = false;
    while (!opened) {
      clickNew();
      // small delay before checking
      await new Promise(r => setTimeout(r, 600));
      opened = await waitForText('New Time Entry', 1500);
    }
    // Proceed with your existing flow once it's open
    const category = [...document.querySelectorAll('.ms-Dropdown-caretDownWrapper')].filter(e => e.offsetParent)[0];
    const activity = [...document.querySelectorAll('.ms-Dropdown-caretDownWrapper')].filter(e => e.offsetParent)[1];

// === NEW: use Category and Activity directly from Excel, with defaults ===
const categoryText = (row["Category"] || '').trim() || 'Post-Sales';
const activityText = (row["Activity"] || '').trim() || 'Reactive/Tape-out support';



// Pick Category only if it's not "Pre-Sales" (default)
if (categoryText !== 'Pre-Sales') {
  await pickDropdownOption(category, categoryText);
}

// Pick Activity only if it's not one of the default values
const defaultActivities = [
  'Marketing events',
  'Account management',
  'Paid Consulting',
  'Administrative',
  'Recurring issues and crashes',
  'Architecture, Spec, Use cases'
];

if (!defaultActivities.includes(activityText)) {
  await pickDropdownOption(activity, activityText);
}

await new Promise(async resolve => {
  const doSaveAndClose = async () => {
    Array.from(document.querySelectorAll('span.ms-Button-label'))
      .find(el => el.textContent.trim() === "Save and Close")
      ?.closest('button')?.click();
    await waitForSaveAndCloseToDisappear();
    resolve();
  };

  // Case 1: Administration → skip filling
  if (categoryText === 'Administration') {
    await doSaveAndClose();
    return;
  }

  // Case 2: Normal flow → wait for inputs and fill
  const customerLogo = await waitForElement('input[placeholder="Search for Customer"]');
  const productName  = await waitForElement('input[placeholder="Search for Product"]');

  // Select Logo (will keep trying until it succeeds)
  await selectFromSuggestionsForever(customerLogo, logo, {
    perCharDelay: 100,
    checkInterval: 2000,
    afterTypeWait: 220
  });

  // Select Product (will keep trying until it succeeds)
  await selectFromSuggestionsForever(productName, tool, {
    perCharDelay: 100,
    checkInterval: 2000,
    afterTypeWait: 220
  });

  // fillCommentBox(comment);
  await doSaveAndClose();
  await new Promise(r => setTimeout(r, 1000));
});

    // tiny buffer before next row (optional)
    await new Promise(r => setTimeout(r, 1000));
  }

  // === After the loop finishes, fill Mon–Fri for all rows ===
  await new Promise(r => setTimeout(r, 500));        // let the grid render all items
  await window.fillWeekFromExcel(window.myExcelData); // fill week % for every Excel row
  console.log('✅ Fill success');
});


