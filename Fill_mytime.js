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
          'Product L1': mapProductName(row['Product L1'])
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


async function pickDropdownOption(caretEl, optionText, timing = {}) {
  const { beforeOpen = 1000, afterOpen = 500, afterPick = 500 } = timing;
  if (!caretEl) return false;
  const normalize = s => (s || '').replace(/\u00a0/g, ' ').trim();

  await delay(beforeOpen);
  caretEl.click();
  await delay(afterOpen);

  const option = [...document.querySelectorAll('.ms-Dropdown-optionText')]
    .find(el => el.offsetParent !== null && normalize(el.textContent) === normalize(optionText));

  if (!option) return false;

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

function populateTextbox(element, text, callback) {
  element.focus();
  text.split('').forEach((char, index) => {
    setTimeout(() => {
      const keyEvent = new KeyboardEvent('keydown', { key: char });
      element.dispatchEvent(keyEvent);
      element.value += char;
      const inputEvent = new Event('input', { bubbles: true });
      element.dispatchEvent(inputEvent);

      if (index === text.length - 1 && typeof callback === 'function') {
        setTimeout(callback, 300);
      }
    }, index * 100);
  });
}

function waitForLogoToAppearAndClick(text, callback) {
  const interval = setInterval(() => {
    const match = Array.from(document.querySelectorAll('div.ms-Persona-primaryText'))
      .find(el => el.textContent.trim() === text);
    if (match) {
      clearInterval(interval);
      match.click();
      if (typeof callback === 'function') {
        setTimeout(callback, 300);
      }
    }
  }, 200);
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

async function waitForSaveAndCloseToDisappear(timeout = 10000) {
  const start = performance.now();
  while (performance.now() - start < timeout) {
    const button = Array.from(document.querySelectorAll('span.ms-Button-label'))
      .find(el => el.textContent.trim() === "Save and Close");
    if (!button || button.offsetParent === null) {
      return true;
    }
    await new Promise(r => setTimeout(r, 200));
  }
  return false;
}
// Fill week values in two passes to avoid row disappearing:
// Pass 1: fill non-zero days first; Pass 2: clear zero days.
window.fillWeekFromExcel = async function (dataRows) {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const norm  = s => (s || '').replace(/\u00a0/g, ' ').trim().toLowerCase();
  const days  = ['Monday','Tuesday','Wednesday','Thursday','Friday'];

  // Decide if a target is "zero/empty"
  const isZeroish = (v) => {
    if (v == null) return true;
    const s = String(v).replace(/%/g, '').replace(/[^\d.,-]/g, '').replace(',', '.').trim();
    if (s === '') return true;
    const n = Number(s);
    return !isFinite(n) ? false : n === 0;
  };

  // Get the target text we want to set into the cell (leave as-is from Excel)
  const targetText = (v) => (v == null ? '' : String(v));

  for (const r of (dataRows || [])) {
    const logo = (r['Logo'] || '').trim();
    const prod = (r['Product L1'] || '').trim();
    const isFTO = norm(logo) === 'fto' || norm(prod) === 'fto';

    // Find the row: FTO -> headerColumn1 === "Vacation, LOA"; else normal Logo+Product
    let rowEl;
    if (isFTO) {
      rowEl = [...document.querySelectorAll('.ms-DetailsRow-fields[data-automationid="DetailsRowFields"]')]
        .find(el => norm(el.querySelector('[data-automation-key="headerColumn1"]')?.textContent) === norm('Vacation, LOA'));
      if (!rowEl) continue;

      // Click the "Vacation, LOA" cell to ensure the row is active/visible
      const headerCell = rowEl.querySelector('[data-automation-key="headerColumn1"]');
      headerCell?.scrollIntoView({ block: 'center' });
      headerCell?.click();
      await sleep(60);
    } else {
      rowEl = [...document.querySelectorAll('.ms-DetailsRow-fields[data-automationid="DetailsRowFields"]')]
        .find(el =>
          norm(el.querySelector('[data-automation-key="headerColumn2"]')?.textContent) === norm(logo) &&
          norm(el.querySelector('[data-automation-key="headerColumn3"]')?.textContent) === norm(prod)
        );
      if (!rowEl) continue;
    }

    // Build target values array for Mon..Fri (fallback to %work if day missing)
    const vals = days.map(d => r[d] ?? r['%work'] ?? r['% work'] ?? '');

    // --- PASS 1: set NON-ZERO targets first ---
    for (let i = 0; i < 5; i++) {
      if (isZeroish(vals[i])) continue; // skip zeros for now
      const input = rowEl.querySelector(`[data-automation-key="dataColumn${i}"] input`);
      if (!input) continue;

      input.focus();
      // Replace content directly (no prior delete) to avoid empty row state
      input.value = targetText(vals[i]);
      input.dispatchEvent(new Event('input',  { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await sleep(40);
    }

    // If all are zeroish, skip clearing to avoid deleting the row entirely
    const hasAnyNonZero = vals.some(v => !isZeroish(v));
    if (!hasAnyNonZero) {
      continue; // nothing to fill without risking disappearance
    }

    // --- PASS 2: now clear ZERO targets safely (row already has a non-zero) ---
    for (let i = 0; i < 5; i++) {
      if (!isZeroish(vals[i])) continue; // only zero targets
      const input = rowEl.querySelector(`[data-automation-key="dataColumn${i}"] input`);
      if (!input) continue;

      input.focus();
      // Clear the cell
      input.value = '';
      input.dispatchEvent(new Event('input',  { bubbles: true }));
      input.dispatchEvent(new Event('change', { bubbles: true }));
      await sleep(30);
    }

    await sleep(80); // small gap before next row
  }
};


// Start everything (process sequentially)
loadExcelData(async () => {
  for (const row of window.myExcelData) {
    const logo = row["Logo"];
    const tool = mapProductName(row["Product L1"]);
    const comment = row["Case Number"];

    const customerLogo = document.querySelector('input[placeholder="Search for Customer"]');
    const productName = document.querySelector('input[placeholder="Search for Product"]');

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

    // === NEW: choose texts based on FTO ===
    // Treat these as "FTO" types (adjust list as needed)
    const isFTO = v => /^\s*(FTO|PTO|Flex Time Off)\s*$/i.test((v || '').trim());
    const categoryText = (isFTO(logo) || isFTO(tool)) ? 'Administration' : 'Post-Sales';
    const activityText = (isFTO(logo) || isFTO(tool)) ? 'Vacation, LOA' : 'Reactive/Tape-out support';

    // Pick Category and Activity
    await pickDropdownOption(category, categoryText);
    await pickDropdownOption(activity, activityText);



// Await the entire fill + save + dialog close flow (FTO vs non-FTO)
await new Promise(resolve => {
  const doSaveAndClose = async () => {
    Array.from(document.querySelectorAll('span.ms-Button-label'))
      .find(el => el.textContent.trim() === "Save and Close")
      ?.closest('button')?.click();
    await waitForSaveAndCloseToDisappear();
    resolve();
  };

  // --- Case 1: FTO -> only Save & Close, skip populate flow ---
  if (isFTO(tool)) {
    (async () => { await doSaveAndClose(); })();
    return;
  }

  // --- Case 2: Normal flow -> populate then Save & Close ---
  populateTextbox(customerLogo, logo, () => {
    waitForLogoToAppearAndClick(logo, () => {
      populateTextbox(productName, tool, () => {
        waitForLogoToAppearAndClick(tool, async () => {
          fillCommentBox(comment);
          await doSaveAndClose();
        });
      });
    });
  });
});


    // tiny buffer before next row (optional)
    await new Promise(r => setTimeout(r, 1000));
  }

  // === After the loop finishes, fill Mon–Fri for all rows ===
  await new Promise(r => setTimeout(r, 500));        // let the grid render all items
  await window.fillWeekFromExcel(window.myExcelData); // fill week % for every Excel row
  console.log('✅ Fill success');
});
