/** CONFIG **/
const CFG = {
  SHEET_CF1: 'EN + Local',
  SHEET_CF2: 'All languages',
  LIMIT_CF1: 18005,
  LIMIT_CF1_NM: 18010,
  LIMIT_CF1_VM: 18015,
  LIMIT_CF2: 18020,
  LIMIT_CF2_NM: 18025,
  LIMIT_CF2_VM: 18030,
  MA: 'Music & Audio'   // category name to control
};

/**
 * Opens the Channel Picker sidebar
 */
function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Channel Picker')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Sidebar bootstrap */
function getMeta(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  if (!sheet) throw new Error('Sheet not found: ' + sheetName);

  const sheetNames = ss.getSheets().map(s => s.getName());
  const lastRow = sheet.getLastRow();
  const langs = new Set(), cats = new Set();

  if (lastRow >= 2) {
    // Read E..F (language, category)
    sheet.getRange(2, 5, lastRow - 1, 2).getValues().forEach(([lg, ct]) => {
      if (lg) langs.add(String(lg));
      if (ct) cats.add(String(ct));
    });
  }
  return {
    active: sheet.getName(),
    sheetNames,
    languages: Array.from(langs).sort(),
    categories: Array.from(cats).sort(),
    rules: {
      CF1: { sheet: CFG.SHEET_CF1, limit: CFG.LIMIT_CF1, target: 'Filtered_CF1' },
      CF2: { sheet: CFG.SHEET_CF2, limit: CFG.LIMIT_CF2, target: 'Filtered_CF2' }
    }
  };
}

function buildFiltered(opts) {
  const ss = SpreadsheetApp.getActive();
  const tag = String(opts.tag || '').toUpperCase();
  if (!tag) throw new Error('Pick a campaign tag.');

  const isCF1 = tag.startsWith('CF1');
  const isCF2 = tag.startsWith('CF2');
  const isNM = tag.includes('NM');
  const isVM = tag.includes('VM');

  const requiredSheet = isCF1 ? CFG.SHEET_CF1 : CFG.SHEET_CF2;
  if (opts.sourceSheet !== requiredSheet) {
    throw new Error(`"${tag}" must use sheet "${requiredSheet}". Current: "${opts.sourceSheet}".`);
  }

  const sheet = ss.getSheetByName(opts.sourceSheet);
  if (!sheet) throw new Error('Sheet not found: ' + opts.sourceSheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) throw new Error('No data to process.');

  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // Source columns for language/category (E/F)
  const COL_LANG = 5;
  const COL_CAT  = 6;

  // Allow-lists from UI
  let allowLangs = Array.isArray(opts.includeLanguages) && opts.includeLanguages.length ? new Set(opts.includeLanguages) : null;
  let allowCats  = Array.isArray(opts.includeCategories) && opts.includeCategories.length ? new Set(opts.includeCategories) : null;

  // Enforce NM/VM rules
  if (isNM) { if (allowCats) allowCats.delete(CFG.MA); }
  if (isVM) { allowCats = new Set([CFG.MA]); opts.includeBlankCategory = false; }

  const includeBlankLang = !!opts.includeBlankLanguage;
  const includeBlankCat  = !!opts.includeBlankCategory;

  // Dynamic limit values based on campaign type
  let defaultLimit;
  if (isCF1) {
    if (isNM) defaultLimit = CFG.LIMIT_CF1_NM;
    else if (isVM) defaultLimit = CFG.LIMIT_CF1_VM;
    else defaultLimit = CFG.LIMIT_CF1;
  } else if (isCF2) {
    if (isNM) defaultLimit = CFG.LIMIT_CF2_NM;
    else if (isVM) defaultLimit = CFG.LIMIT_CF2_VM;
    else defaultLimit = CFG.LIMIT_CF2;
  }

  const limit = Math.max(0, Number(opts.limit || defaultLimit));

  // Filter
  const filtered = [];
  for (const row of rows) {
    const lang = String(row[COL_LANG - 1] || '').trim();
    const cat  = String(row[COL_CAT  - 1] || '').trim();

    // Language filter
    if (allowLangs) {
      if (!lang && !includeBlankLang) continue;
      if (lang && !allowLangs.has(lang)) continue;
    } else if (!includeBlankLang && !lang) {
      continue;
    }

    // Category filter  
    if (allowCats) {
      if (!cat && !includeBlankCat) continue;
      if (cat && !allowCats.has(cat)) continue;
    } else if (!includeBlankCat && !cat) {
      continue;
    }

    // NM rule: exclude Music & Audio
    if (isNM && cat === CFG.MA) continue;

    filtered.push(row);
  }

  // Sort if desired
  const sortKey = String(opts.sortKey || 'none').toUpperCase();
  const sortDir = String(opts.sortDir || 'DESC').toUpperCase();
  let sorted = filtered;
  if (sortKey !== 'NONE') {
    const colIndex = ({ A: 1, E: 5, F: 6 })[sortKey] || 1;
    sorted = filtered.slice().sort((a, b) => {
      const va = a[colIndex - 1], vb = b[colIndex - 1];
      const na = typeof va === 'number' ? va : Number(String(va).replace(/[^0-9.-]/g, ''));
      const nb = typeof vb === 'number' ? vb : Number(String(vb).replace(/[^0-9.-]/g, ''));
      const bothNum = !isNaN(na) && !isNaN(nb);
      let cmp = bothNum ? (na - nb) : String(va).localeCompare(String(vb));
      return sortDir === 'ASC' ? cmp : -cmp;
    });
  }

  const finalRows = limit ? sorted.slice(0, limit) : sorted;

  // Create/write target sheet
  const targetName = `Filtered_${tag.replace(/\s+/g, '_')}`;
  const out = ss.getSheetByName(targetName) || ss.insertSheet(targetName);
  out.clear();

  // Keep original header and data as-is
  const outHeader = [...header];
  out.getRange(1, 1, 1, outHeader.length).setValues([outHeader]);

  // Copy rows without modification - Channel IDs in column B are already prepared
  const enhancedRows = finalRows.map(row => {
    const r = [...row];
    // Fill up any shorter rows to match header length
    while (r.length < outHeader.length) r.push('');
    return r;
  });

  if (enhancedRows.length) {
    out.getRange(2, 1, enhancedRows.length, outHeader.length).setValues(enhancedRows);
  }

  out.setFrozenRows(1);
  try {
    out.getRange(1, 2, out.getLastRow(), 1).setHorizontalAlignment('center').setFontWeight('bold'); // B
    out.getRange(1, 3, out.getLastRow(), 1).setHorizontalAlignment('left');                         // C
  } catch(_) {}

  return {
    targetName,
    kept: finalRows.length,
    totalAfterFilter: sorted.length,
    limit,
    note: isNM ? 'NM rule: "Music & Audio" excluded.' : (isVM ? 'VM rule: only "Music & Audio".' : ''),
    channelIdsGenerated: enhancedRows.filter(r => r[1]).length
  };
}

/**
 * NEW FUNCTION: Extracts Channel ID from YouTube URL
 * @param {string} url - YouTube URL
 * @returns {string} Channel ID or empty string if none found
 */
function extractChannelIdFromUrl(url) {
  if (!url || typeof url !== 'string') return '';
  
  const urlStr = String(url).trim();
  
  // Different YouTube URL formats to handle:
  // https://www.youtube.com/channel/UC_x5XG1OV2P6uZZ5FSM9Ttw
  // https://www.youtube.com/c/channelname
  // https://www.youtube.com/user/username
  // https://youtube.com/watch?v=videoid
  
  try {
    // Channel ID pattern (UC followed by 22 characters)
    const channelIdMatch = urlStr.match(/\/channel\/([a-zA-Z0-9_-]{24})/);
    if (channelIdMatch) return channelIdMatch[1];
    
    // Custom URL pattern - return as is (can be used for manual mapping later)
    const customMatch = urlStr.match(/\/c\/([^\/\?&#]+)/);
    if (customMatch) return `@${customMatch[1]}`; // Prefix with @ for custom names
    
    // User pattern
    const userMatch = urlStr.match(/\/user\/([^\/\?&#]+)/);
    if (userMatch) return `user:${userMatch[1]}`; // Prefix with user:
    
    // If nothing else, try to extract domain-specific part
    if (urlStr.includes('youtube.com/') || urlStr.includes('youtu.be/')) {
      const urlParts = urlStr.split('/');
      const lastPart = urlParts[urlParts.length - 1];
      if (lastPart && lastPart.length > 3) {
        return lastPart.split(/[\?&#]/)[0]; // Remove query parameters
      }
    }
    
  } catch (e) {
    console.log(`Error extracting channel ID from URL: ${urlStr}`, e);
  }
  
  return ''; // No Channel ID found
}

/**
 * NEW FUNCTION: Validates and repairs Channel ID data in a sheet
 * @param {string} sheetName - Name of the sheet to validate
 * @returns {Object} Validation result
 */
function validateAndRepairChannelIds(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    return { error: `Sheet "${sheetName}" not found` };
  }
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { message: 'No data rows to process' };
  }
  
  const dataRange = sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), 3));
  const values = dataRange.getValues();
  
  let repaired = 0;
  let validated = 0;
  
  const repairedValues = values.map(row => {
    const url = String(row[1] || '').trim(); // Column B
    let channelId = String(row[2] || '').trim(); // Column C
    
    if (url && !channelId) {
      // Try to extract Channel ID from URL
      const extractedId = extractChannelIdFromUrl(url);
      if (extractedId) {
        row[2] = extractedId;
        repaired++;
      }
    } else if (channelId) {
      validated++;
    }
    
    return row;
  });
  
  // Update the sheet if we repaired anything
  if (repaired > 0) {
    dataRange.setValues(repairedValues);
  }
  
  return {
    processed: values.length,
    repaired: repaired,
    validated: validated,
    message: `Processed ${values.length} rows. Repaired ${repaired} Channel IDs, validated ${validated} existing IDs.`
  };
}

/**
 * NEW FUNCTION: Repairs Channel IDs for all Filtered_* sheets
 */
function repairAllChannelIds() {
  const ss = SpreadsheetApp.getActive();
  const allSheets = ss.getSheets();
  const filteredSheets = allSheets.filter(sheet => sheet.getName().startsWith('Filtered_'));
  
  const results = [];
  
  filteredSheets.forEach(sheet => {
    const result = validateAndRepairChannelIds(sheet.getName());
    results.push({
      sheetName: sheet.getName(),
      ...result
    });
  });
  
  // Log the results
  console.log('=== CHANNEL ID REPAIR RESULTS ===');
  results.forEach(result => {
    console.log(`${result.sheetName}: ${result.message || result.error}`);
  });
  
  const totalRepaired = results.reduce((sum, r) => sum + (r.repaired || 0), 0);
  SpreadsheetApp.getUi().alert(`Channel ID repair completed. ${totalRepaired} IDs repaired across ${filteredSheets.length} sheets.`);
  
  return results;
}
