/** SPCA Pet Pantry — Order Form Auto-Populate (Slides → PDF)
 * Production version — Trigger-Safe
 * Generates new labeled PDF when form response is received
 */

const CONFIG = {
  SOURCE_SHEET_ID: '1JrfUHDAPMCIvOSknKoN3vVR6KQZTKUaNLpsRru7cekU',
  RESPONSE_SHEET_NAME: 'Form Responses 1',
  GUIDELINES_SHEET_ID: '1ujUGOzUlw7G1m8rPJgqV2qKOr5PIkEL9QVjEiMNswNU',
  SLIDES_TEMPLATE_ID: '1SxWI9modxXjQNwxpeFNKFcbnKvxt55J0MwjnVBqevbM',
  OUTPUT_FOLDER_ID: '1ccWalGXHxLJVN-G92GTzv8OIF_rHq7QO',
  PET_SLOTS: 6,
  FORM_ID_COLUMN: 'FormID',
  TZ: Session.getScriptTimeZone(),
  BARCODE: {
    PLACEHOLDER: '{{barcode}}',
    TYPE: 'code128',
    PX: { width: 1100, height: 500 },
    TARGET_HEIGHT_IN: 2.5,
  },
};

/**
 * === MAIN ENTRY (trigger) ===
 * Runs automatically when a new form response is received.
 * Populates FormID first, then generates PDF and URL.
 */
function onFormSubmit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.RESPONSE_SHEET_NAME) return;

    const rowIndex = e.range.getRow();
    if (rowIndex <= 1) return; // skip header row

    generateOrderForm_(rowIndex);
  } catch (err) {
    Logger.log(`❌ onFormSubmit failed: ${err}`);
  }
}


/**
 * === INTERNAL CORE ===
 * Handles FormID assignment, merge, barcode, and PDF generation.
 * Skips regeneration unless explicitly forced.
 *
 * @param {number} rowIndex - The target row number (1-based)
 * @param {boolean} [force=false] - Force regeneration even if a PDF already exists
 */
function generateOrderForm_(rowIndex, force = false) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SOURCE_SHEET_ID);
    const sh = ss.getSheetByName(CONFIG.RESPONSE_SHEET_NAME);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const values = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const nv = Object.fromEntries(headers.map((h, i) => [h, values[i]]));

    // --- Column indexes ---
    const idCol = headers.indexOf('Generated PDF ID') + 1;
    const urlCol = headers.indexOf('Generated PDF URL') + 1;
    const tsCol = headers.indexOf('Generated At') + 1;

    // --- Check for existing generation ---
    const alreadyHasPdf =
      idCol > 0 && urlCol > 0 &&
      sh.getRange(rowIndex, idCol).getValue() &&
      sh.getRange(rowIndex, urlCol).getValue();

    // --- Skip duplicate generation unless forced ---
    if (alreadyHasPdf && !force) {
      Logger.log(`⏩ Skipped row ${rowIndex}: already has a generated PDF`);
      return { ok: true, skipped: true };
    }

    // --- Assign FormID if missing ---
    if (!nv[CONFIG.FORM_ID_COLUMN]) {
      const formIdCol = headers.indexOf(CONFIG.FORM_ID_COLUMN) + 1;
      if (!formIdCol) throw new Error('FormID column missing');
      const newId = generateNextFormId_();
      sh.getRange(rowIndex, formIdCol).setValue(newId);
      nv[CONFIG.FORM_ID_COLUMN] = newId;
      Logger.log(`🆕 Assigned FormID ${newId} to row ${rowIndex}`);
    }

    // --- Compute and generate new PDF ---
    const summary = computeAnimalSummary_(nv, rowIndex);
    const merge = buildMergeMap_(nv, summary, rowIndex);
    const pdfFile = mergeSlidesToPdf_(merge, makeOutputName_(nv));

    // --- Write back PDF info ---
    if (!alreadyHasPdf) {
      // First-time generation → write to main columns
      if (idCol > 0) sh.getRange(rowIndex, idCol).setValue(pdfFile.getId());
      if (urlCol > 0) sh.getRange(rowIndex, urlCol).setValue(pdfFile.getUrl());
      if (tsCol > 0) sh.getRange(rowIndex, tsCol).setValue(new Date());
    } else {
      // Forced regeneration (preserve originals)
      const regenIdCol = headers.indexOf('Regenerated PDF ID') + 1;
      const regenUrlCol = headers.indexOf('Regenerated PDF URL') + 1;
      const regenTsCol = headers.indexOf('Last Regenerated At') + 1;

      if (regenIdCol > 0) sh.getRange(rowIndex, regenIdCol).setValue(pdfFile.getId());
      if (regenUrlCol > 0) sh.getRange(rowIndex, regenUrlCol).setValue(pdfFile.getUrl());
      if (regenTsCol > 0) sh.getRange(rowIndex, regenTsCol).setValue(new Date());
    }

    Logger.log(
      `✅ ${force ? 'Regenerated' : 'Generated'} form for row ${rowIndex} → ${pdfFile.getUrl()}`
    );
    return { ok: true, url: pdfFile.getUrl(), regenerated: force };

  } catch (err) {
    Logger.log(`❌ generateOrderForm_ failed: ${err}`);
    return { ok: false, message: err.message };
  }
}

/**
 * Generates the next sequential 12-digit Form ID and saves it in Script Properties.
 */
function generateNextFormId_() {
  const props = PropertiesService.getScriptProperties();
  const last = Number(props.getProperty('LAST_FORM_ID') || 100000000542);
  const next = last + 1;
  props.setProperty('LAST_FORM_ID', next);
  return String(next).padStart(12, '0');
}

/**
 * Recursively clears any unreplaced {{placeholders}} inside shapes, tables, or groups.
 * Replaces with equal-length spaces to preserve layout.
 */
function clearPlaceholdersInElement_(el) {
  const type = el.getPageElementType();
  try {
    if (type === SlidesApp.PageElementType.SHAPE) {
      const text = el.asShape().getText();
      const raw = text.asString();
      if (raw.includes('{{')) {
        const cleaned = raw.replace(/\{\{[^}]+\}\}/g, m => ' '.repeat(m.length));
        text.setText(cleaned);
      }
    }
    if (type === SlidesApp.PageElementType.TABLE) {
      const table = el.asTable();
      for (let r = 0; r < table.getNumRows(); r++) {
        for (let c = 0; c < table.getRow(r).getNumCells(); c++) {
          const text = table.getCell(r, c).getText();
          const raw = text.asString();
          if (raw.includes('{{')) {
            const cleaned = raw.replace(/\{\{[^}]+\}\}/g, m => ' '.repeat(m.length));
            text.setText(cleaned);
          }
        }
      }
    }
    if (type === SlidesApp.PageElementType.GROUP) {
      el.asGroup().getChildren().forEach(child => clearPlaceholdersInElement_(child));
    }
  } catch (err) {
    Logger.log(`⚠️ clearPlaceholdersInElement_ skipped ${type}: ${err}`);
  }
}

/** === MERGE + PDF EXPORT === **/
function mergeSlidesToPdf_(mergeMap, outName) {
  const template = DriveApp.getFileById(CONFIG.SLIDES_TEMPLATE_ID);
  const tmp = template.makeCopy(`_tmp_${outName}`, DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID));
  const pres = SlidesApp.openById(tmp.getId());

  pres.getSlides().forEach(slide => {
    Object.entries(mergeMap).forEach(([ph, val]) => {
      try {
        slide.replaceAllText(ph, String(val ?? ''));
      } catch (e) {
        Logger.log('⚠️ replaceAllText failed for %s: %s', ph, e);
      }
    });
  });

  const formId = mergeMap['{{FormID}}'] || mergeMap['{{formId}}'];
  if (formId) {
    try { insertBarcodeIntoSlides_(pres, formId); } catch (e) { Logger.log('⚠️ barcode skip: %s', e); }
  }

  // Clean up unreplaced placeholders safely (preserve layout)
  pres.getSlides().forEach(slide => {
    slide.getPageElements().forEach(el => clearPlaceholdersInElement_(el));
  });

  pres.saveAndClose();

  const blob = tmp.getAs(MimeType.PDF).setName(`${outName}.pdf`);
  const pdf = DriveApp.getFolderById(CONFIG.OUTPUT_FOLDER_ID).createFile(blob);
  tmp.setTrashed(true);
  return pdf;
}

function makeOutputName_(nv) {
  const ts = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyyMMdd_HHmm');
  const first = nv['First Name'] || 'Unknown';
  const last = nv['Last Name'] || '';
  return `PetPantryForm_${first}_${last}_${ts}`;
}

/** === BARCODE === **/
function insertBarcodeIntoSlides_(pres, formId) {
  const slides = pres.getSlides();
  const targetHeightPt = CONFIG.BARCODE.TARGET_HEIGHT_IN * 72;
  const ratio = CONFIG.BARCODE.PX.width / CONFIG.BARCODE.PX.height;
  const blob = UrlFetchApp.fetch(buildBarcodeUrl_(String(formId))).getBlob();

  slides.forEach(slide => {
    slide.getPageElements().forEach(el => {
      if (el.getPageElementType() !== SlidesApp.PageElementType.SHAPE) return;
      const shape = el.asShape();
      const txt = shape.getText().asString();
      if (!txt.includes(CONFIG.BARCODE.PLACEHOLDER)) return;
      const box = { l: el.getLeft(), t: el.getTop(), w: el.getWidth(), h: el.getHeight() };
      let h = Math.min(targetHeightPt, box.h), w = h * ratio;
      if (w > box.w) { w = box.w; h = w / ratio; }
      const left = box.l + (box.w - w) / 2;
      const top = box.t + (box.h - h) / 2;
      shape.getText().replaceAllText(CONFIG.BARCODE.PLACEHOLDER, '');
      slide.insertImage(blob, left, top, w, h);
    });
  });
}

function buildBarcodeUrl_(text) {
  const base = 'https://quickchart.io/barcode';
  const p = CONFIG.BARCODE;
  const qs = `?text=${encodeURIComponent(text)}&type=${p.TYPE}&format=png&width=${p.PX.width}&height=${p.PX.height}&margin=0`;
  return base + qs;
}

/** === MERGE MAP === **/
function buildMergeMap_(nv, summary, rowIndex) {
  const merge = {};

  // Basic info
  merge['{{formDate}}'] = Utilities.formatDate(new Date(), CONFIG.TZ, 'M/d/yyyy');
  merge['{{FormID}}'] = nv['FormID'] || '';
  merge['{{fullName}}'] = `${nv['First Name'] || ''} ${nv['Last Name'] || ''}`.trim();
  merge['{{phone}}'] = nv['Phone Number'] || '';
  merge['{{email}}'] = nv['Email Address'] || '';

  const contactRaw = nv['Preferred Contact Method'] || nv['Contact Method'] || '';
  let contactClean = '';
  if (contactRaw) {
    const matches = contactRaw.match(/\b(Text|Email|Phone)\b(?=\s*[–—-])/gi);
    if (matches && matches.length) {
      const unique = [...new Set(matches.map(m =>
        m.charAt(0).toUpperCase() + m.slice(1).toLowerCase()
      ))];
      contactClean = unique.join(', ');
    }
  }
  merge['{{Contact}}'] = contactClean;

  merge['{{addressLine1}}'] = nv['Address Line 1'] || '';
  merge['{{addressLine2}}'] = nv['Address Line 2'] || '';
  merge['{{city}}'] = nv['Town/City'] || '';
  merge['{{state}}'] = nv['State'] || '';
  merge['{{zip}}'] = nv['Zip Code'] || '';
  merge['{{newClient}}'] = nv['Returning Client'] || '';
  merge['{{services}}'] = nv['Additional Services'] || '';
  merge['{{pickupWindow}}'] =
    /(sat|sun|weekend)/i.test(String(nv['Pick-up Window'] || '').toLowerCase()) ? 'Saturday' : 'Weekday';
  merge['{{todaysDate}}'] = Utilities.formatDate(new Date(), CONFIG.TZ, 'M/d/yyyy');
  merge['{{lastName}}'] = nv['Last Name'] || '';

  // Pet slots
  for (let i = 1; i <= CONFIG.PET_SLOTS; i++) {
    const prefix = `Pet ${i}`;
    merge[`{{${i}Name}}`] = nv[`${prefix} Name`] || '';
    merge[`{{${i}Species}}`] = nv[`${prefix} Species`] || '';
    merge[`{{${i}Breed}}`] = nv[`${prefix} Breed`] || '';
    merge[`{{${i}Color}}`] = nv[`${prefix} Color`] || '';
    merge[`{{${i}Age}}`] = nv[`${prefix} Age`] || '';
    merge[`{{${i}Units}}`] = nv[`${prefix} Units`] || '';
    merge[`{{${i}Weight}}`] = nv[`${prefix} Weight`] || '';
    merge[`{{${i}Sex}}`] = nv[`${prefix} Sex`] || '';
    merge[`{{${i}SPN}}`] = nv[`${prefix} Spay/Neuter`] || '';
  }

  // Summary
  merge['{{adultDogCount}}'] = summary.adultDogCount || 0;
  merge['{{puppyCount}}'] = summary.puppyCount || 0;
  merge['{{adultCatCount}}'] = summary.adultCatCount || 0;
  merge['{{kittenCount}}'] = summary.kittenCount || 0;
  merge['{{dogSizes}}'] = summary.dogSizes || '';
  merge['{{otherSpecies}}'] = summary.otherSpecies || '';

  // Recommended items
  try {
    const itemsMap = buildItemPlaceholderMap_(nv, rowIndex);
    Object.assign(merge, itemsMap);
  } catch (err) {
    Logger.log('⚠️ buildItemPlaceholderMap_ failed: %s', err);
  }
  return merge;
}

/** === ANIMAL SUMMARY === **/
function computeAnimalSummary_(nv, rowIndex) {
  let dogCount = 0, catCount = 0, puppyCount = 0, kittenCount = 0;
  const otherSpecies = [];
  const sizeCounts = { Miniature: 0, Small: 0, Medium: 0, Large: 0, Giant: 0 };

  for (let i = 1; i <= CONFIG.PET_SLOTS; i++) {
    const species = String(nv[`Pet ${i} Species`] || '').toLowerCase().trim();
    if (!species) continue;
    const units = String(nv[`Pet ${i} Units`] || '').toLowerCase();
    const weight = parseWeightLbs_(nv[`Pet ${i} Weight`]);

    if (species === 'dog') {
      dogCount++;
      if (units.startsWith('month')) puppyCount++;
      else if (weight) sizeCounts[dogSizeFromWeight_(weight)]++;
    } else if (species === 'cat') {
      catCount++;
      if (units.startsWith('month')) kittenCount++;
    } else {
      const pretty = capitalizeFirst_(species);
      if (pretty && !otherSpecies.includes(pretty)) otherSpecies.push(pretty);
    }
  }

  const adultDogCount = dogCount - puppyCount;
  const adultCatCount = catCount - kittenCount;
  const dogSizes = Object.entries(sizeCounts)
    .filter(([_, c]) => c > 0)
    .map(([k, c]) => `${c} ${k}`)
    .join(', ');

  // Write counts back to sheet
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SOURCE_SHEET_ID);
    const sh = ss.getSheetByName(CONFIG.RESPONSE_SHEET_NAME);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const map = Object.fromEntries(headers.map((h, i) => [h.trim(), i + 1]));
    const cols = [
      map['CountAdultDogs'], map['CountPuppies'],
      map['CountAdultCats'], map['CountKittens']
    ];
    if (cols.every(c => c)) {
      sh.getRange(rowIndex, cols[0], 1, 4)
        .setValues([[adultDogCount, puppyCount, adultCatCount, kittenCount]]);
    }
  } catch (err) {
    Logger.log('⚠️ summary write failed: %s', err);
  }

  return {
    adultDogCount, puppyCount,
    adultCatCount, kittenCount,
    dogSizes, otherSpecies: otherSpecies.join(', ')
  };
}

/** === HELPERS === **/
function getFirst(v) { return Array.isArray(v) ? v[0] : v; }
function parseWeightLbs_(raw) {
  const m = String(raw || '').match(/(\d+(\.\d+)?)/);
  return m ? parseFloat(m[1]) : null;
}
function dogSizeFromWeight_(lbs) {
  if (lbs < 12) return 'Miniature';
  if (lbs < 25) return 'Small';
  if (lbs < 50) return 'Medium';
  if (lbs < 100) return 'Large';
  return 'Giant';
}
function capitalizeFirst_(s) {
  const str = String(s || '').trim();
  return str ? str.charAt(0).toUpperCase() + str.slice(1) : '';
}
function parseRequested_(csv) {
  if (!csv) return [];
  return String(csv).split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
}
function toNum_(val) {
  if (val === null || val === '' || val === undefined) return null;
  const str = String(val).trim();
  const n = parseFloat(str);
  return isNaN(n) ? null : n;
}
function inferSpecies_(name) {
  const n = String(name || '').toLowerCase();
  if (n.includes('dog')) return 'dog';
  if (n.includes('cat')) return 'cat';
  return null;
}

/** === LOAD GUIDELINES === **/
function loadAvailability_() {
  const ss = SpreadsheetApp.openById(CONFIG.GUIDELINES_SHEET_ID);
  const sh = ss.getSheets()[0];
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(h => String(h || '').trim().toLowerCase());
  const idx = (name) => headers.findIndex(h => h === name.toLowerCase());

  const iItem = idx('item');
  const iPlaceholder = idx('placeholders');
  const iPerPet = idx('per pet');
  const iHHMax = idx('household max');
  const iNotes = idx('notes');
  const iAmountGiven = idx('amountgiven');
  const iAmtPH = idx('amount placeholder');

  const map = {};
  values.forEach(row => {
    const item = String(row[iItem] || '').trim();
    if (!item) return;
    const key = item.toLowerCase();
    const perPet = toNum_(row[iPerPet]);
    const hhMax = toNum_(row[iHHMax]);

    map[key] = {
      item,
      placeholder: String(row[iPlaceholder] || '').trim(),
      perPet,
      householdMax: hhMax,
      notes: String(row[iNotes] || '').trim(),
      amountGiven: String(row[iAmountGiven] || '').trim(),
      amountPlaceholder: String(row[iAmtPH] || '').trim(),
    };
  });

  Logger.log('✅ Loaded %s guideline items from "%s"', Object.keys(map).length, sh.getName());
  return map;
}

/**
 * Build placeholder → replacement map for Recommended Item List
 * using Distribution Guidelines logic and Puppy/Kitten auto-handling.
 */
function buildItemPlaceholderMap_(nv, rowIndex) {
  const ss = SpreadsheetApp.openById(CONFIG.SOURCE_SHEET_ID);
  const sh = ss.getSheetByName(CONFIG.RESPONSE_SHEET_NAME);
  const guidelines = loadAvailability_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);

  // Extract Resources Requested list
  const requestedCsv = getFirst(nv['Resources Requested']) || getFirst(nv['Requested Resources']) || '';
  const requestedItems = requestedCsv
    .split(',')
    .map(s => s.trim().toLowerCase())
    .filter(Boolean);

  // Get counts from sheet (or summary columns)
  const getCol = (name) => headers.indexOf(name) + 1;
  const getVal = (name) => {
    const c = getCol(name);
    return c > 0 ? Number(sh.getRange(rowIndex, c).getValue() || 0) : 0;
  };
  const countAdultDogs = getVal('CountAdultDogs');
  const countPuppies   = getVal('CountPuppies');
  const countAdultCats = getVal('CountAdultCats');
  const countKittens   = getVal('CountKittens');

  const placeholderMap = {};
  const amountMap = {};
  const log = (msg, ...args) => Logger.log(`[buildItemPlaceholderMap_] ${msg}`, ...args);

  // Helper to compute the display text for a given guideline row
  function computeLine(item, count) {
    const rule = guidelines[item.toLowerCase()];
    if (!rule) return null;

    const perPet = rule.perPet;
    const hhMax = rule.householdMax;
    let qty = 0;

    if (perPet == null) {
      qty = hhMax && !isNaN(hhMax) ? Number(hhMax) : 0;
    } else {
      const rawQty = count * perPet;
      qty = hhMax && !isNaN(hhMax) ? Math.min(rawQty, hhMax) : rawQty;
    }

    if (!qty || qty <= 0) return null;

    const finalText = `${qty} ${rule.notes || ''}`.trim();
    return { qty, text: finalText, rule };
  }

  // --- STANDARD REQUESTED ITEMS ---
  requestedItems.forEach(req => {
    const rule = guidelines[req];
    if (!rule) return; // skip unmatched
    const species = inferSpecies_(req);
    let count = 0;
    if (species === 'dog') count = countAdultDogs;
    else if (species === 'cat') count = countAdultCats;

    const result = computeLine(req, count);
    if (!result) return;

    placeholderMap[rule.placeholder] = result.text;
    if (rule.amountPlaceholder) amountMap[rule.amountPlaceholder] = rule.amountGiven || '';
  });

  // --- PUPPY LOGIC ---
  if (
    countPuppies > 0 &&
    requestedItems.some(r =>
      r.includes('dog food') || r.includes('wet dog food') || r.includes('canned dog food')
    )
  ) {
    ['Dry Puppy Food', 'Wet Puppy Food'].forEach(item => {
      const result = computeLine(item, countPuppies);
      if (!result) return;
      const rule = guidelines[item.toLowerCase()];
      if (rule) {
        placeholderMap[rule.placeholder] = result.text;
        if (rule.amountPlaceholder) amountMap[rule.amountPlaceholder] = rule.amountGiven || '';
      }
    });
  }

  // --- KITTEN LOGIC ---
  if (
    countKittens > 0 &&
    requestedItems.some(r =>
      r.includes('cat food') || r.includes('wet cat food') || r.includes('canned cat food')
    )
  ) {
    ['Dry Kitten Food', 'Wet Kitten Food'].forEach(item => {
      const result = computeLine(item, countKittens);
      if (!result) return;
      const rule = guidelines[item.toLowerCase()];
      if (rule) {
        placeholderMap[rule.placeholder] = result.text;
        if (rule.amountPlaceholder) amountMap[rule.amountPlaceholder] = rule.amountGiven || '';
      }
    });
  }

  // --- COMBINE ---
  const combined = Object.assign({}, placeholderMap, amountMap);
  log('✅ Built item placeholder map with %s entries', Object.keys(combined).length);
  return combined;
}