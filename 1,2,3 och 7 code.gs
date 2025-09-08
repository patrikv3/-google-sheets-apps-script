/**
 * @OnlyCurrentDoc
 *
 * Detta skript kombinerar funktioner för arkhantering, TTT-konvertering,
 * CSV-generering (backend) och UI-hantering för Google Ads-verktyget.
 * Baserat på ursprungliga Skript 1 och Skript 7.
 * Inkluderar även targeting-funktionalitet från v1.6
 */

// ===============================================
// KONFIGURATION (från Skript 7)
// ===============================================
const CSV_TEMPLATE_SHEET_NAME = "CSV Template"; // <--- *** UPPDATERA VID BEHOV ***
const ALL_AGE_OPTIONS = ["18-24", "25-34", "35-44", "45-54", "55-64", "65-up", "Unknown"];
const ALL_GENDER_OPTIONS = ["Male", "Female", "Unknown"]; 
const NEGATIVE_TOPICS_SHEET_NAME = "Negative Topics List";
const NEGATIVE_LISTS_SHEET_NAME = "Excl lists NL";

// --- Konfiguration för Targeting ---
const TARGETING_SHEET = "Targeting";

// Definiera master-arken och kolumnerna för targeting
const MASTER_SHEETS_CONFIG = {
  Topics: {
    name: "Topics", 
    nameCol: "Topics",
    idCol: "Topics ID",
    masterKeyCol: "Category path",
    masterValueCol: "ID"
  },
  Affinities: {
    name: "Affinities",
    nameCol: "Affinities",
    idCol: "Affinities ID",
    masterKeyCol: "Category path",
    masterValueCol: "ID"
  },
  InMarket: {
    name: "In-market",
    nameCol: "In-market",
    idCol: "In-market ID",
    masterKeyCol: "Category path",
    masterValueCol: "ID"
  },
  LifeEvents: {
    name: "Life-events",
    nameCol: "Life events",
    idCol: "Life events ID",
    masterKeyCol: "Category path",
    masterValueCol: "ID"
  }, 
  DetailedDemographics: {
    name: "Detailed demographics",
    nameCol: "Detailed demographics",
    idCol: "Detailed demographics ID",
    masterKeyCol: "Category path",     // <<--- ÄNDRA TILL DETTA!
    masterValueCol: "ID"
  } 
};

// ===============================================
// VID ÖPPNING & MENY (Uppdaterad version - Ny ordning och namn)
// ===============================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // --- Kör funktioner som ska köras vid öppning (Setup) ---
  try {
    removeDeletedSheets_();
    Logger.log("onOpen: removeDeletedSheets_ kördes.");
  } catch (e) {
    Logger.log(`Fel vid körning av removeDeletedSheets_ i onOpen: ${e}`);
  }
  try {
    manageSheets_();
    Logger.log("onOpen: manageSheets_ kördes.");
  } catch (e) {
    Logger.log(`Fel vid körning av manageSheets_ i onOpen: ${e}`);
     ui.alert(`Ett fel uppstod vid hantering av TTT-ark vid öppning: ${e.message}`);
  }
  try {
    resetFilterInSFData_();
    Logger.log("onOpen: resetFilterInSFData_ kördes.");
  } catch (e) {
    Logger.log(`Fel vid körning av resetFilterInSFData_ i onOpen: ${e}`);
  }

  // --- Skapa Meny (Ny Ordning) ---
  ui.createMenu('Ad Tools')
    .addItem('1. Clear Data validation and rows in TTT', 'clearValidationInTTT') // Flyttad hit och omdöpt
    .addItem('2. Convert TTT', 'convertTTT') // Ny numrering
    .addItem('3. Update Video/Ad/Tracking', 'updateVideoAdTrackingValues') // Ny numrering
    .addItem('4. Update TTT from SF Data', 'updateTTTFromSFData') // Ny numrering
    .addItem('5. Update Geo/Language', 'updateTTTGeoLanguage') // Ny numrering
    .addItem('6. Update Campaign Mapping', 'updateCampaignMapping') // Ny numrering
    .addItem('7. Build Compilation Sheet', 'buildCampaignSummary') // Ny numrering
    .addSeparator()
    .addItem('Generate Ads CSV', 'showPopup_') // Flyttad hit
    .addItem('Populera Targeting IDs', 'runPopulateTargetingIDs')
    .addItem('Lägg till targeting på befintliga kampanjer', 'showTargetingPopup')
    .addSeparator()
    .addItem('Run Full Update (Steps 2-7), Not working yet', 'runFullUpdateAndCompile') // Flyttad hit och etikett justerad
    .addToUi();

  Logger.log("onOpen: Meny 'Ad Tools' skapades med ny ordning."); // Justerad loggtext
}

/**
 * ===============================================
 * MASTER FUNCTION - KÖR FLERA STEG
 * ===============================================
 */
function runFullUpdateAndCompile() {
  const ui = SpreadsheetApp.getUi();
  const startTime = new Date();
  Logger.log("runFullUpdateAndCompile: Startar fullständig uppdatering och kompilering.");
  SpreadsheetApp.getActiveSpreadsheet().toast("Startar fullständig uppdatering (Steg 1-6)...", "Status", 10);

  try {
    Logger.log("runFullUpdateAndCompile: Kör Steg 1: convertTTT");
    SpreadsheetApp.getActiveSpreadsheet().toast("Kör Steg 1/6: Convert TTT...", "Pågår...", 5);
    convertTTT(); 

    Logger.log("runFullUpdateAndCompile: Kör Steg 2: updateVideoAdTrackingValues");
     SpreadsheetApp.getActiveSpreadsheet().toast("Kör Steg 2/6: Update Video/Ad/Tracking...", "Pågår...", 5);
    updateVideoAdTrackingValues();

    Logger.log("runFullUpdateAndCompile: Kör Steg 3: updateTTTFromSFData");
     SpreadsheetApp.getActiveSpreadsheet().toast("Kör Steg 3/6: Update TTT from SF Data...", "Pågår...", 5);
    updateTTTFromSFData();

    Logger.log("runFullUpdateAndCompile: Kör Steg 4: updateTTTGeoLanguage");
     SpreadsheetApp.getActiveSpreadsheet().toast("Kör Steg 4/6: Update Geo/Language...", "Pågår...", 5);
    updateTTTGeoLanguage();

    Logger.log("runFullUpdateAndCompile: Kör Steg 5: updateCampaignMapping");
     SpreadsheetApp.getActiveSpreadsheet().toast("Kör Steg 5/6: Update Campaign Mapping...", "Pågår...", 5);
    updateCampaignMapping();

    Logger.log("runFullUpdateAndCompile: Kör Steg 6: buildCampaignSummary");
     SpreadsheetApp.getActiveSpreadsheet().toast("Kör Steg 6/6: Build Compilation Sheet...", "Pågår...", 5);

    const endTime = new Date();
    const duration = (endTime.getTime() - startTime.getTime()) / 1000; 
    Logger.log(`runFullUpdateAndCompile: Hela processen klar på ${duration.toFixed(1)} sekunder.`);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Fullständig uppdatering (1-6) klar! (${duration.toFixed(1)} s)`, "Slutförd", 10);
    ui.alert(`Fullständig uppdatering och kompilering (steg 1-6) är klar!\nTotal tid: ${duration.toFixed(1)} sekunder.`);

  } catch (e) {
    Logger.log(`Ett allvarligt fel inträffade under runFullUpdateAndCompile: ${e}\nStack: ${e.stack}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Fel under uppdateringsprocessen!", "FEL", 10);
    ui.alert(`Ett fel inträffade under den fullständiga uppdateringen:\n\n${e.message}\n\nProcessen avbröts. Kontrollera vilket steg som kördes senast och se loggen för detaljer.`);
  }
}

// ===============================================
// ARK-HANTERING 
// ===============================================
function resetFilterInSFData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sfDataSheet = ss.getSheetByName("SF Data");
  if (sfDataSheet) {
    const currentFilter = sfDataSheet.getFilter();
    if (currentFilter) {
      currentFilter.remove();
      Logger.log('Filter borttaget från "SF Data".');
    }
  } else {
    Logger.log('Arket "SF Data" hittades inte, kan inte återställa filter.');
  }
}

function clearValidationInTTT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "TTT";
  const sheet = ss.getSheetByName(sheetName);
  const ui = SpreadsheetApp.getUi();

  if (sheet) {
    try {
      // 1. Kontrollera om ett filter är aktivt
      const currentFilter = sheet.getFilter();
      if (currentFilter) {
        Logger.log(`Filter är aktivt på arket "${sheetName}". Avbryter clearValidationInTTT.`);
        ui.alert("Filter Active on TTT Sheet", "Please remove the filter from the 'TTT' sheet and try again.", ui.ButtonSet.OK);
        return; // Avbryt funktionen
      }
      Logger.log(`Inget filter aktivt på arket "${sheetName}". Fortsätter.`);

      // 2. Rensa datavalideringar
      sheet.getDataRange().clearDataValidations(); 
      Logger.log("Datavalidering borttagen från TTT-arket.");
      // ui.alert("All data validation has been removed from the TTT sheet."); // Kanske inte behövs om nästa alert kommer

      // 3. Ta bort överflödiga tomma rader
      // Använder din existerande funktion removeTrailingEmptyRows.
      // Om threshold är 4, krävs det minst 5 (4+1) konsekutiva tomma rader i kolumn C efter sista datan för att rader ska tas bort.
      const referenceColumnForEmptyRows = "C";
      const emptyRowThreshold = 4; // Detta betyder att 5 eller fler tomma rader triggar borttagning

      Logger.log(`Anropar removeTrailingEmptyRows från clearValidationInTTT för ark: "${sheetName}", refKolumn: "${referenceColumnForEmptyRows}", tröskel: ${emptyRowThreshold}.`);
      removeTrailingEmptyRows(sheetName, referenceColumnForEmptyRows, emptyRowThreshold); 
      // removeTrailingEmptyRows kommer att ge sitt eget toast-meddelande om rader tas bort.
      
      ui.alert("TTT sheet processing complete", "Data validations have been cleared. Trailing empty rows (if any, based on criteria) have been removed.", ui.ButtonSet.OK);

    } catch (e) {
      Logger.log(`Kunde inte slutföra bearbetning av TTT-arket: ${e.message}\nStack: ${e.stack}`);
      ui.alert("Error Processing TTT Sheet", `An error occurred: ${e.message}`, ui.ButtonSet.OK);
    }
  } else {
    Logger.log(`Kunde inte hitta arket "${sheetName}".`);
    ui.alert("Sheet Not Found", `Sheet "${sheetName}" could not be found.`, ui.ButtonSet.OK);
  }
}

function removeTrailingEmptyRows(sheetName = "TTT", referenceColumn = "C", threshold = 4) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  const ui = SpreadsheetApp.getUi();
  if (!sheet) {
    Logger.log(`removeTrailingEmptyRows: Arket "${sheetName}" hittades inte.`);
    return;
  }
  const maxRows = sheet.getMaxRows();
  let lastRowWithDataInRefCol = 0;
  const referenceColumnNumber = sheet.getRange(referenceColumn + "1").getColumn();
  const columnValues = sheet.getRange(1, referenceColumnNumber, maxRows, 1).getValues();
  for (let i = columnValues.length - 1; i >= 0; i--) {
    if (String(columnValues[i][0]).trim() !== "") {
      lastRowWithDataInRefCol = i + 1; 
      break;
    }
  }
  Logger.log(`removeTrailingEmptyRows: Ark "${sheetName}". Max rader: ${maxRows}. Sista rad med data i kolumn ${referenceColumn}: ${lastRowWithDataInRefCol}.`);
  if (lastRowWithDataInRefCol === 0) {
    Logger.log(`removeTrailingEmptyRows: Ingen data hittades i kolumn ${referenceColumn} i arket "${sheetName}". Inga rader tas bort.`);
    return;
  }
  if (lastRowWithDataInRefCol >= maxRows) {
    Logger.log(`removeTrailingEmptyRows: Sista raden med data är den sista raden i arket. Inga överflödiga tomma rader att ta bort från "${sheetName}".`);
    return;
  }
  let consecutiveEmptyRows = 0;
  const firstRowToScanForEmptiness = lastRowWithDataInRefCol + 1;
  const rowsToConsiderForThreshold = threshold + 1; 
  if (firstRowToScanForEmptiness > maxRows) {
     Logger.log(`removeTrailingEmptyRows: Inga rader att skanna efter sista datapunkten i kolumn ${referenceColumn}.`);
     return;
  }
  for (let i = 0; i < rowsToConsiderForThreshold; i++) {
    const currentRowBeingScanned = firstRowToScanForEmptiness + i;
    if (currentRowBeingScanned > maxRows) { 
      break;
    }
    if (String(columnValues[currentRowBeingScanned - 1][0]).trim() === "") { 
      consecutiveEmptyRows++;
    } else {
      Logger.log(`removeTrailingEmptyRows: Hittade data på rad ${currentRowBeingScanned} i kolumn ${referenceColumn} vid tröskelkontroll. Tar inte bort några rader.`);
      return; 
    }
  }
  Logger.log(`removeTrailingEmptyRows: Skannade upp till ${rowsToConsiderForThreshold} rader efter sista data. Hittade ${consecutiveEmptyRows} konsekutiva tomma celler i kolumn ${referenceColumn}. Tröskelvärde: > ${threshold}.`);
  if (consecutiveEmptyRows > threshold) { 
    const startDeleteRow = lastRowWithDataInRefCol + 1;
    const numRowsToDelete = maxRows - lastRowWithDataInRefCol;
    if (numRowsToDelete > 0) {
      try {
        Logger.log(`removeTrailingEmptyRows: Tar bort ${numRowsToDelete} rader från arket "${sheetName}", från och med rad ${startDeleteRow}.`);
        sheet.deleteRows(startDeleteRow, numRowsToDelete);
        SpreadsheetApp.getActiveSpreadsheet().toast(`${numRowsToDelete} tomma rader har tagits bort från slutet av arket "${sheetName}".`);
      } catch (e) {
        Logger.log(`removeTrailingEmptyRows: Fel vid borttagning av rader från arket "${sheetName}": ${e}`);
        ui.alert(`Ett fel uppstod när tomma rader skulle tas bort från "${sheetName}": ${e.message}`);
      }
    } else {
      Logger.log(`removeTrailingEmptyRows: Antal rader att ta bort är ${numRowsToDelete}. Inga rader tas bort.`);
    }
  } else {
    Logger.log(`removeTrailingEmptyRows: Villkoret ej uppfyllt (hittade ${consecutiveEmptyRows} tomma, tröskel är > ${threshold}). Inga rader tas bort från arket "${sheetName}".`);
  }
}

function removeDeletedSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetsToDelete = ss.getSheets().filter(sh => sh.getName() === "Delete");
  if (sheetsToDelete.length > 0) {
    sheetsToDelete.forEach(sh => {
      try {
        ss.deleteSheet(sh);
        Logger.log(`Tog bort arket "Delete" (ID: ${sh.getSheetId()}).`);
      } catch (e) {
        Logger.log(`Kunde inte ta bort ark "Delete" (ID: ${sh.getSheetId()}): ${e}`);
        SpreadsheetApp.getUi().alert(`Kunde inte ta bort ett ark med namnet "Delete". Det kan behöva tas bort manuellt.`);
      }
    });
  } else {
    Logger.log('Inga ark med namnet "Delete" hittades.');
  }
}

function manageSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let tttSheet = null;
  let copySheet = null;
  sheets.forEach(sh => {
    const name = sh.getName();
    if (name === "TTT") { tttSheet = sh; }
    else if (name.includes("TTT") && name.includes("Copy")) {
      if (name.toLowerCase().startsWith("copy of ttt")) { copySheet = sh; }
    }
  });
  Logger.log(`manageSheets_: Hittade TTT: ${tttSheet ? tttSheet.getName() : 'Nej'}, Copy of TTT: ${copySheet ? copySheet.getName() : 'Nej'}`);
  try {
    if (copySheet && tttSheet) {
      Logger.log(`manageSheets_: Både TTT och Copy of TTT finns. Byter namn på TTT till Delete och Copy of TTT till TTT.`);
      tttSheet.setName("Delete"); tttSheet.hideSheet(); 
      copySheet.setName("TTT"); copySheet.showSheet(); 
    } else if (copySheet && !tttSheet) {
      Logger.log(`manageSheets_: Endast Copy of TTT finns. Byter namn till TTT.`);
      copySheet.setName("TTT"); copySheet.showSheet();
    } else if (tttSheet && !copySheet) {
       Logger.log(`manageSheets_: Endast TTT finns. Gör inget namnbyte.`);
       tttSheet.showSheet(); 
    } else {
       Logger.log(`manageSheets_: Varken TTT eller Copy of TTT hittades. Ingen åtgärd för namnbyte.`);
    }
    const finalSheets = ss.getSheets();
    let finalTttSheetExists = false;
    finalSheets.forEach(sh => {
      if (sh.getName() === "TTT") { sh.showSheet(); finalTttSheetExists = true; }
      else if (sh.getName() !== "Delete") { sh.hideSheet(); }
    });
     Logger.log(`manageSheets_: TTT-arket är ${finalTttSheetExists ? '' : 'inte '}synligt. Övriga ark (utom Delete) dolda.`);
  } catch (e) {
    Logger.log(`Ett fel inträffade i manageSheets_: ${e}`);
    SpreadsheetApp.getUi().alert(`Ett fel inträffade vid hantering av TTT-ark: ${e.message}. Kontrollera arkens namn manuellt.`);
  }
}

// ===============================================
// HJÄLPFUNKTION: Hitta Kolumnindex
// ===============================================
function findHeaderIndex_(headerArray, headerName) {
  if (!headerName || !Array.isArray(headerArray)) {
     Logger.log(`findHeaderIndex_ anropades med ogiltiga argument. headerName: ${headerName}, headerArray: ${Array.isArray(headerArray)}`);
     return -1;
  }
  const target = headerName.toLowerCase().trim();
  for (let i = 0; i < headerArray.length; i++) {
    const currentHeader = headerArray[i];
    if (typeof currentHeader === 'string' && currentHeader.trim().toLowerCase() === target) {
      return i + 1; 
    }
  }
  return -1; 
}

/**
 * HJÄLPFUNKTION: Skapar och lägger till CSV-rader för specifik inriktning.
 * Använder nu 'name' och 'id' från targetingItem. Prioriterar ID för relevanta kolumner.
 */
function addSpecificTargetingRows_(
    csvData, 
    campaignNameForCsv, 
    adGroupNameForOutput, 
    targetingItemsArray, 
    col, 
    campaignTemplateRowDataLength,
    defaultCriterionStatus,
    logPrefix = ""
) {
    if (!targetingItemsArray || targetingItemsArray.length === 0) {
        return;
    }

    Logger.log(`${logPrefix} Adding ${targetingItemsArray.length} targeting item(s) to AG: "${adGroupNameForOutput}" in C: "${campaignNameForCsv}"...`);
    targetingItemsArray.forEach(targetingItem => {
        const targetRow = createEmptyRow_(campaignTemplateRowDataLength); 
        setValue_(targetRow, col.campaign, campaignNameForCsv);
        setValue_(targetRow, col.adGroup, adGroupNameForOutput);
        setValue_(targetRow, col.status, targetingItem.status || defaultCriterionStatus); 

        const itemTypeLower = String(targetingItem.itemType || "").toLowerCase();
        const itemName = targetingItem.name || ""; 
        const itemId = targetingItem.id || "";   

        if (itemTypeLower === "audience") {
            const criterionTypeValue = targetingItem.audienceCategory || "Audience"; 
            setValue_(targetRow, col.criterionType, criterionTypeValue);
            
            if (itemId && col.id !== -1) { 
                setValue_(targetRow, col.id, itemId);
            }
            if (col.audienceName !== -1) { 
                setValue_(targetRow, col.audienceName, itemName); 
            } else if (!itemId || col.id === -1) { 
                 Logger.log(`${logPrefix}  WARNING: Neither col.id (for criterion ID) nor col.audienceName is mapped for Audience '${itemName}'.`);
            }
        } else if (itemTypeLower === "topic") {
            setValue_(targetRow, col.criterionType, "Topic");
            if (itemId && col.topicId !== -1) { 
                 setValue_(targetRow, col.topicId, itemId);
            }
            if (col.topic !== -1) { 
                setValue_(targetRow, col.topic, itemName); 
            } else if (!itemId || col.topicId === -1) {
                 Logger.log(`${logPrefix}  WARNING: Neither col.topicId nor col.topic (for name) is mapped for Topic '${itemName}'.`);
            }
        } else {
            Logger.log(`${logPrefix}  WARNING: Unknown Targeting Item Type "${targetingItem.itemType || 'undefined'}". Skipping for Name: ${itemName}.`);
            return; 
        }
        csvData.push(targetRow);
    });
}

// ===============================================
// TTT KONVERTERING (från Skript 1)
// ===============================================

/**
 * Huvudfunktionen "Convert TTT":
 * Körs från menyn 'Ad Tools'.
 * 1) Byter rubriker i rad 2 (Placement name -> Campaign, etc.)
 * 2) Skapar nya kolumner (alla grupper + extra)
 * 3) Tar bort eventuella dubblettkolumner
 * 4) Kontrollerar/fixar datum i "Start Date TTT" och "End Date TTT"
 * 5) Fyller standardvärden i "Networks", "Budget" och "Budget type"
 * endast på rader där "Campaign" inte är tom.
 * 6) Lägger till kolumnen "Opportunity Name". **SÄKERSTÄLLD FUNKTIONALITET 2025-05-07**
 */
function convertTTT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheet = ss.getSheetByName("TTT");

  if (!sheet) {
    ui.alert('Arket "TTT" hittades inte. Kontrollera att det finns ett ark med exakt det namnet och att det är synligt.\n\n(Kördes `manageSheets` korrekt vid öppning?)');
    Logger.log('convertTTT: Arket "TTT" hittades inte.');
    return;
  }
  if (sheet.isSheetHidden()) {
    ui.alert('Arket "TTT" är dolt. Gör det synligt och försök igen.\n\n(Kördes `manageSheets` korrekt vid öppning?)');
    Logger.log('convertTTT: Arket "TTT" är dolt.');
    return;
  }

  try {
    Logger.log("convertTTT: Startar konvertering.");
    SpreadsheetApp.getActiveSpreadsheet().toast("Startar konvertering av TTT...", "Status", 5);

    Logger.log("convertTTT: Kör Steg 1: renameHeaders_");
    renameHeaders_(sheet);
    Logger.log("convertTTT: Steg 1/6 (renameHeaders_) klart.");

    Logger.log("convertTTT: Kör Steg 2: insertAllColumns_");
    insertAllColumns_(sheet);
    Logger.log("convertTTT: Steg 2/6 (insertAllColumns_) klart.");

    Logger.log("convertTTT: Kör Steg 3: removeDuplicateColumns_");
    removeDuplicateColumns_(sheet);
    Logger.log("convertTTT: Steg 3/6 (removeDuplicateColumns_) klart.");

    Logger.log("convertTTT: Kör Steg 4: checkTTTDatesAndFix_");
    checkTTTDatesAndFix_(sheet);
    Logger.log("convertTTT: Steg 4/6 (checkTTTDatesAndFix_) klart.");

    Logger.log("convertTTT: Kör Steg 5: fillDefaultsWhereCampaignNotEmpty_");
    fillDefaultsWhereCampaignNotEmpty_(sheet);
    Logger.log("convertTTT: Steg 5/6 (fillDefaultsWhereCampaignNotEmpty_) klart.");

    Logger.log("convertTTT: Kör Steg 6: insertOpportunityNameColumn_");
    insertOpportunityNameColumn_(sheet); // Denna funktion skapar "Opportunity Name"
    Logger.log("convertTTT: Steg 6/6 (insertOpportunityNameColumn_) klart.");

    SpreadsheetApp.getActiveSpreadsheet().toast("Konvertering av TTT klar!", "Status", 5);
    ui.alert("Konvertering av TTT klar!");
    Logger.log("convertTTT: Konvertering slutförd framgångsrikt.");

  } catch (e) {
    Logger.log(`Ett fel inträffade under convertTTT: ${e.message}\nStack: ${e.stack || 'Ingen stack tillgänglig'}`);
    ui.alert(`Ett fel inträffade under konverteringen av TTT:\n\n${e.message}\n\nSe loggen (Utföranden) för mer information.`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Fel vid konvertering!", "Fel", 10);
  }
}

// --- Hjälpfunktioner för convertTTT ---

function renameHeaders_(sheet) {
  const headerRow = 2;
  if (sheet.getMaxRows() < headerRow || sheet.getMaxColumns() < 1) {
     Logger.log(`renameHeaders_: Arket har inte tillräckligt med rader/kolumner för att hitta rubriker på rad ${headerRow}.`);
     throw new Error(`Arket "${sheet.getName()}" verkar tomt eller saknar rubrikrad ${headerRow}.`);
  }
  const lastCol = sheet.getLastColumn();
   if (lastCol === 0) {
       Logger.log(`renameHeaders_: Arket "${sheet.getName()}" verkar helt tomt (getLastColumn() == 0).`);
       throw new Error(`Arket "${sheet.getName()}" verkar helt tomt.`);
   }
  const rng = sheet.getRange(headerRow, 1, 1, lastCol);
  const headers = rng.getValues()[0];
  let changed = false;
  const newHeaders = headers.map(header => {
    const txt = String(header || "").trim();
    const low = txt.toLowerCase();
    let newHeader = txt; 
    if (low === "placement name") newHeader = "Campaign";
    else if (low === "click tag") newHeader = "Tracking template";
    else if (low.includes("landing page url") && low !== "landing page url") newHeader = "Landing Page";
    else if (low.includes("language") && low !== "language") newHeader = "Language";
    else if (low.includes("rotation") && low !== "rotation") newHeader = "Rotation";
    else if (low === "start date") newHeader = "Start Date TTT";
    else if (low === "end date") newHeader = "End Date TTT";
    else if (low === "device") newHeader = "Device TTT";
    if (newHeader !== txt) { changed = true; }
    return newHeader;
  });
  if (changed) { 
    rng.setValues([newHeaders]);
    Logger.log("renameHeaders_: Rubriker uppdaterade.");
  } else { 
    Logger.log("renameHeaders_: Inga rubriker behövde ändras."); 
  }
}

function insertAllColumns_(sheet) {
  const headerRow = 2;
   if (sheet.getMaxRows() < headerRow || sheet.getLastColumn() < 1) {
       Logger.log("insertAllColumns_: Kan inte läsa rubriker, arket för litet.");
       throw new Error(`Arket "${sheet.getName()}" saknar rubrikrad ${headerRow} eller kolumner.`);
   }
  let headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newColGroups = [ 
    { ref: "Campaign", columns: [ { name: "Location", hidden: false }, { name: "ID", hidden: false }, { name: "Campaign Type", hidden: false }, { name: "Bid Strategy Type", hidden: false }, { name: "Networks", hidden: true }, { name: "Budget", hidden: true }, { name: "Budget type", hidden: true }, { name: "Client Rate (converted)", hidden: true }, { name: "Media Products", hidden: false } ] },
    { ref: "Format", columns: [ { name: "Ad Group Type", hidden: false }, { name: "Ad type", hidden: false } ] },
    // Viktigt: Se till att "Companion Videos" INTE skapas här om den ska vara referens för Opportunity Name
    // Om "Companion Videos" är en av de ursprungliga kolumnerna i TTT är det OK.
    // Om den är ny, och "Opportunity Name" ska komma EFTER den, så kan den inte döljas förrän "Opportunity Name" är på plats,
    // eller så måste "Opportunity Name" ha en annan referenspunkt.
    // För nu, antar vi att "Companion Videos" finns eller skapas och är synlig.
    { ref: "Device TTT", columns: [ { name: "Desktop Bid Modifier", hidden: true }, { name: "Mobile Bid Modifier", hidden: true }, { name: "Tablet Bid Modifier", hidden: true }, { name: "TV Screen Bid Modifier", hidden: true }, { name: "Age", hidden: false }, { name: "Gender", hidden: true }, { name: "Parental status", hidden: true } ] },
    { ref: "Youtube Link", columns: [ { name: "Video ID", hidden: false }, { name: "Ad Name", hidden: false }, { name: "Companion Videos", hidden: false} ] }, // Exempel: Companion Videos skapas här
    { ref: "Tracking template", columns: [ { name: "Final URL", hidden: false }, { name: "Display URL", hidden: false } ] },
    { ref: "Start Date TTT", columns: [ { name: "Start Date SF", hidden: false } ] },
    { ref: "End Date TTT", columns: [ { name: "End Date SF", hidden: false } ] },
    { ref: "Device TTT", columns: [ { name: "Device Targeting", hidden: false } ] },
    { ref: "Client Rate (converted)", columns: [ { name: "Total Client Cost (converted)", hidden: true }, { name: "Total Campaign Placement Days", hidden: true } ] },
    { ref: "Media Products", columns: [ { name: "Multiformat ads", hidden: false } ] }
  ];

  newColGroups.sort((a, b) => { 
    const refA = a && a.ref; const refB = b && b.ref; if (!refA) return 1; if (!refB) return -1;
    const indexA = findHeaderIndex_(headers, refA); const indexB = findHeaderIndex_(headers, refB);
    if (indexA === -1) return 1; if (indexB === -1) return -1; return indexB - indexA; 
  });

  let columnsInsertedTotal = 0;
  newColGroups.forEach(group => {
    if (!group || !group.ref || !Array.isArray(group.columns) || group.columns.length === 0) { Logger.log(`insertAllColumns_: Hoppar över ogiltigt eller tomt gruppobjekt: ${JSON.stringify(group)}`); return; }
    const refIdx = findHeaderIndex_(headers, group.ref);
    if (refIdx === -1) { Logger.log(`insertAllColumns_: Referenskolumn "${group.ref}" hittades inte för gruppen [${group.columns.map(c=>c.name).join(', ')}]. Hoppar över.`); return; }
    if (findHeaderIndex_(headers, group.columns[0].name) !== -1) { Logger.log(`insertAllColumns_: Första kolumnen "${group.columns[0].name}" (för ref "${group.ref}") finns redan. Antar att gruppen redan är infogad.`); return; }
    
    const numColsToInsert = group.columns.length;
    sheet.insertColumnsAfter(refIdx, numColsToInsert);
    columnsInsertedTotal += numColsToInsert;
    Logger.log(`insertAllColumns_: Infogade ${numColsToInsert} kolumn(er) [${group.columns.map(c=>c.name).join(', ')}] efter "${group.ref}" (index ${refIdx}).`);
    const targetHeaderRange = sheet.getRange(headerRow, refIdx + 1, 1, numColsToInsert);
    const newHeaderNames = group.columns.map(c => c.name);
    targetHeaderRange.setValues([newHeaderNames]).setBackground("black").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    group.columns.forEach((col, i) => { 
      if (col.hidden) { 
        const colToHideIndex = refIdx + 1 + i;
        try { sheet.hideColumn(sheet.getRange(1, colToHideIndex)); } catch (e) { Logger.log(`insertAllColumns_: Kunde inte dölja kolumn ${colToHideIndex} ("${col.name}"): ${e}`);}
      } 
    });
    headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0]; // Uppdatera headers array efter varje infogning
  });
  Logger.log(`insertAllColumns_: Totalt ${columnsInsertedTotal} kolumner infogade.`);
}

/**
 * Infogar kolumnen "Opportunity Name". Försöker efter "Companion Videos", annars sist.
 * **UPPDATERAD 2025-05-07 för robusthet.**
 */
function insertOpportunityNameColumn_(sheet) {
  const headerRow = 2;
  if (sheet.getMaxRows() < headerRow || sheet.getLastColumn() === 0) { // Kontroll om arket har några kolumner alls
    Logger.log("insertOpportunityNameColumn_: Arket är tomt eller saknar rubrikrad. Kan inte infoga Opportunity Name.");
    throw new Error("TTT-arket verkar tomt, kan inte lägga till Opportunity Name."); // Kasta fel för att stoppa convertTTT
  }
  // Hämta headers på nytt för att säkerställa att vi har den senaste versionen
  let currentHeaders = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const newColName = "Opportunity Name";

  if (findHeaderIndex_(currentHeaders, newColName) !== -1) {
    Logger.log(`insertOpportunityNameColumn_: Kolumnen "${newColName}" finns redan.`);
    return;
  }

  const companionVideosCol = "Companion Videos";
  let insertAfterIndex = findHeaderIndex_(currentHeaders, companionVideosCol); // 1-baserat

  if (insertAfterIndex === -1) {
    Logger.log(`insertOpportunityNameColumn_: Kolumnen "${companionVideosCol}" hittades INTE. Försöker lägga till "${newColName}" som sista kolumn.`);
    insertAfterIndex = sheet.getLastColumn(); 
    // Om getLastColumn() fortfarande är 0 här (vilket det inte borde vara efter insertAllColumns_),
    // behöver vi ett grundfall, men det är osannolikt.
    if (insertAfterIndex === 0) { // Fallback om arket på något sätt blev helt tomt
        Logger.log(`insertOpportunityNameColumn_: Arket har inga kolumner, infogar ${newColName} som första.`);
        sheet.insertColumnBefore(1); // Infoga som allra första kolumn
        insertAfterIndex = 0; // Så att newColActualIndex blir 1
    }
  } else {
    Logger.log(`insertOpportunityNameColumn_: Hittade "${companionVideosCol}" på kolumnindex ${insertAfterIndex}. Infogar "${newColName}" efter denna.`);
  }
  
  try {
    if (insertAfterIndex > 0) {
        sheet.insertColumnsAfter(insertAfterIndex, 1); 
    } else { // Om insertAfterIndex blev 0 (ovan fallback), har vi redan infogat med insertColumnBefore
        // inget mer att göra här för att skapa kolumnen.
    }
    const newColActualIndex = insertAfterIndex + 1; 
    const rng = sheet.getRange(headerRow, newColActualIndex);
    rng.setValue(newColName); 
    rng.setBackground("black").setFontColor("white").setFontWeight("bold").setHorizontalAlignment("center");
    Logger.log(`insertOpportunityNameColumn_: Kolumn "${newColName}" infogad/satt på faktiskt kolumnindex ${newColActualIndex}.`);
  } catch (e) {
     Logger.log(`insertOpportunityNameColumn_: FEL vid försök att infoga kolumn efter index ${insertAfterIndex}. Fel: ${e.message} \nStack: ${e.stack}`);
     SpreadsheetApp.getUi().alert(`Kunde inte infoga kolumnen "Opportunity Name": ${e.message}`);
  }
}

function removeDuplicateColumns_(sheet) {
  const headerRow = 2;
  if (sheet.getMaxRows() < headerRow || sheet.getLastColumn() < 1) { Logger.log("removeDuplicateColumns_: Kan inte läsa rubriker..."); return; }
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  const seen = {};
  let deletedCount = 0;
  for (let c = lastCol - 1; c >= 0; c--) {
    const headerText = String(headers[c] || "").trim().toLowerCase();
    if (!headerText) { 
      Logger.log(`removeDuplicateColumns_: Hittade tom rubrik i kolumn ${c + 1}. Tar bort.`);
      sheet.deleteColumn(c + 1); deletedCount++; continue;
    }; 
    if (!seen[headerText]) {
      seen[headerText] = true; 
    } else {
      sheet.deleteColumn(c + 1); 
      deletedCount++;
      Logger.log(`removeDuplicateColumns_: Tog bort dubblettkolumn "${headers[c]}" (index ${c + 1}).`);
    }
  }
  Logger.log(`removeDuplicateColumns_: Tog bort totalt ${deletedCount} dubblett-/tomma kolumner.`);
}

function checkTTTDatesAndFix_(sheet) {
  const headerRow = 2; const firstDataRow = 3;
  if (sheet.getMaxRows() < firstDataRow || sheet.getLastColumn() < 1) { Logger.log("checkTTTDatesAndFix_: Inga datarader..."); return; }
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const startColName = "Start Date TTT"; const endColName = "End Date TTT";
  const startIdx = findHeaderIndex_(headers, startColName); 
  const endIdx = findHeaderIndex_(headers, endColName);   
  if (startIdx === -1 && endIdx === -1) { Logger.log(`checkTTTDatesAndFix_: Varken "${startColName}"/"${endColName}" hittades.`); return; }
  const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow) { Logger.log("checkTTTDatesAndFix_: Inga datarader under rubriken."); return; }
  const numRows = lastRow - firstDataRow + 1;
  const dataRange = sheet.getRange(firstDataRow, 1, numRows, sheet.getLastColumn());
  const data = dataRange.getValues();
  let invalidFound = false; let datesChanged = false;
  for (let r = 0; r < data.length; r++) {
    if (startIdx !== -1) {
      const cellValue = data[r][startIdx - 1]; 
      const originalValueString = String(cellValue); 
      if (cellValue !== null && cellValue !== "") { 
        const newValue = parseToMMDDYYYY_(cellValue);
        if (newValue === null) {
          invalidFound = true; Logger.log(`checkTTTDatesAndFix_: Ogiltigt startdatum rad ${firstDataRow + r}: "${cellValue}"`);
        } else {
          if (String(newValue) !== originalValueString) { data[r][startIdx - 1] = newValue; datesChanged = true; }
        }
      }
    }
    if (endIdx !== -1) {
      const cellValue = data[r][endIdx - 1];
      const originalValueString = String(cellValue);
      if (cellValue !== null && cellValue !== "") {
        const newValue = parseToMMDDYYYY_(cellValue);
        if (newValue === null) {
          invalidFound = true; Logger.log(`checkTTTDatesAndFix_: Ogiltigt slutdatum rad ${firstDataRow + r}: "${cellValue}"`);
        } else {
          if (String(newValue) !== originalValueString) { data[r][endIdx - 1] = newValue; datesChanged = true; }
        }
      }
    }
  }
  if (datesChanged) { dataRange.setValues(data); Logger.log("checkTTTDatesAndFix_: Datumkolumner uppdaterade."); }
  else { Logger.log("checkTTTDatesAndFix_: Inga datum behövde uppdateras."); }
  if (invalidFound) { SpreadsheetApp.getUi().alert(`Varning: Ogiltiga datum i "${startColName}"/"${endColName}"...`); }
}

function parseToMMDDYYYY_(val) {
  if (!val) return null; 
  let dateObj = null;
  if (val instanceof Date) { dateObj = val; }
  else if (typeof val === 'string') {
    const trimmedVal = val.trim();
    const replaced = trimmedVal.replace(/[.\-]/g, '/');
    dateObj = new Date(replaced);
    if (isNaN(dateObj.getTime())) {
      const parts = replaced.split('/');
      if (parts.length === 3) {
        const p1 = parseInt(parts[0], 10); const p2 = parseInt(parts[1], 10); const p3 = parseInt(parts[2], 10);
        if (p1 > 12 && p2 <= 12 && p3 > 1900) { dateObj = new Date(p3, p2 - 1, p1); }
        else if (p1 <= 12 && p3 > 1900) { dateObj = new Date(p3, p1 - 1, p2); }
        else { dateObj = new Date(replaced); }
      }
    }
  } else if (typeof val === 'number') {
    if (val > 0 && val < 60000) { 
      const excelEpoch = new Date(1899, 11, 30); 
      const jsDate = new Date(excelEpoch.getTime() + val * 24 * 60 * 60 * 1000);
      const utcDate = new Date(Date.UTC(jsDate.getFullYear(), jsDate.getMonth(), jsDate.getDate()));
      if (!isNaN(utcDate.getTime())) { dateObj = utcDate; }
    }
  }
  if (dateObj && !isNaN(dateObj.getTime())) {
    try { return Utilities.formatDate(dateObj, "UTC", "MM/dd/yyyy"); }
    catch (e) { Logger.log(`Fel vid formatering av datumobjekt ${dateObj}: ${e}`); return null; }
  }
  return null;
}

function fillDefaultsWhereCampaignNotEmpty_(sheet) {
  const headerRow = 2; const firstDataRow = 3; const lastRow = sheet.getLastRow();
  if (lastRow < firstDataRow || sheet.getLastColumn() < 1) { Logger.log("fillDefaults...: Inga datarader..."); return; }
  const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const campaignIdx = findHeaderIndex_(headers, "Campaign");
  const networksIdx = findHeaderIndex_(headers, "Networks");
  const budgetIdx = findHeaderIndex_(headers, "Budget");
  const budgetTypeIdx = findHeaderIndex_(headers, "Budget type");
  if (campaignIdx === -1) { Logger.log('fillDefaults...: "Campaign" hittades inte.'); SpreadsheetApp.getUi().alert('Varning: "Campaign" saknas i TTT...'); return; }
  if (networksIdx === -1 && budgetIdx === -1 && budgetTypeIdx === -1) { Logger.log("fillDefaults...: Ingen av Networks, Budget, Budget type hittades."); return; }
  const numRows = lastRow - firstDataRow + 1;
  const dataRange = sheet.getRange(firstDataRow, 1, numRows, sheet.getLastColumn());
  const data = dataRange.getValues();
  let valuesChanged = false;
  for (let r = 0; r < data.length; r++) {
    const campaignValue = data[r][campaignIdx - 1]; 
    if (campaignValue === null || String(campaignValue).trim() === "") { continue; }
    if (networksIdx !== -1) { const networkValue = data[r][networksIdx - 1]; if (networkValue === null || String(networkValue).trim() === "") { data[r][networksIdx - 1] = "Youtube;YouTube Videos"; valuesChanged = true; } }
    if (budgetIdx !== -1) { const budgetValue = data[r][budgetIdx - 1]; if (budgetValue === null || String(budgetValue).trim() === "") { data[r][budgetIdx - 1] = "1"; valuesChanged = true; } }
    if (budgetTypeIdx !== -1) { const budgetTypeValue = data[r][budgetTypeIdx - 1]; if (budgetTypeValue === null || String(budgetTypeValue).trim() === "") { data[r][budgetTypeIdx - 1] = "Daily"; valuesChanged = true; } }
  }
  if (valuesChanged) { dataRange.setValues(data); Logger.log("fillDefaults...: Standardvärden ifyllda."); }
  else { Logger.log("fillDefaults...: Inga standardvärden behövde fyllas i."); }
}

function testConvertTTT() {
  Logger.log("Kör testConvertTTT...");
  convertTTT();
  Logger.log("testConvertTTT klar.");
}

// ===============================================
// POPUP & CSV-GENERERING BACKEND
// ===============================================

/**
 * Visar popup-fönstret för CSV-generering.
 * Kallas från menyn 'Ad Tools'.
 */
function showPopup_() {
  try {
    // Se till att HTML-filen heter 'popup.html' i ditt Apps Script-projekt
    const html = HtmlService.createHtmlOutputFromFile('popup')
        .setWidth(1050) 
        .setHeight(750);
    SpreadsheetApp.getUi().showModalDialog(html, "Select Campaign Options & Generate CSV");
    Logger.log("showPopup_: Popup-fönster visat.");
  } catch (e) {
    Logger.log(`Popup Error: ${e}\n${e.stack}`);
    SpreadsheetApp.getUi().alert(`Kunde inte öppna popup-fönstret: ${e.message}\n\nKontrollera att filen 'popup.html' finns i projektet.`);
  }
}

/**
 * Genererar CSV-strängen baserat på användarens val från popupen.
 * **KORRIGERAD: Återställd targetingOnlyMode-logik från original, integrerad med Excl lists NL.**
 */
function generateCsv(selections) {
  Logger.log("Initiating CSV generation..."); // Från ditt original
  try {
    Logger.log(`Received selections (structure may vary): ${JSON.stringify(selections)}`); // Från ditt original

    const targetingOnlyMode = selections.targetingOnlyMode || false;
    const importAdsOnlyMode = selections.importMode === 'adsOnly';
    const channelsOnlyMode = selections.channelsOnlyMode || false;
    // Nya flaggor från popupen
    const addNegativeKeywords = selections.addNegativeKeywords || false;
    const addNegativePlacements = selections.addNegativePlacements || false;
    const addChannels = selections.addChannels || false;

    // Validering från ditt original
    if (!selections ||
      (!targetingOnlyMode && !channelsOnlyMode && (!selections.selectedCampaigns || !selections.selectedCFs || !selections.selectedAdGroupTypes)) ||
      (targetingOnlyMode && (!selections.selectedCFs || !selections.selectedAdGroupTypes)) ||
      (channelsOnlyMode && (!selections.selectedCampaigns || !selections.selectedCFs || !selections.selectedAdGroupTypes))) {
      Logger.log("generateCsv: Error - Ogiltiga eller ofullständiga val mottagna för aktuellt läge.");
      return "Error: Invalid selections data received from popup for the current mode.";
    }

    Logger.log(`generateCsv mode: TargetingOnly=${targetingOnlyMode}, AdsOnly=${importAdsOnlyMode}, ChannelsOnly=${channelsOnlyMode}, AddNegKeywords=${addNegativeKeywords}, AddNegPlacements=${addNegativePlacements}`);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName(CSV_TEMPLATE_SHEET_NAME);
    if (!templateSheet) { return `Error: CSV Template sheet "${CSV_TEMPLATE_SHEET_NAME}" not found.`; }

    // templateHeaders och campaignTemplateRowData från ditt original
    const templateLastCol = templateSheet.getLastColumn() > 0 ? templateSheet.getLastColumn() : 1;
    const templateHeaders = templateSheet.getRange(1, 1, 1, templateLastCol).getValues()[0].map(h => String(h || "").trim());
    let campaignTemplateRowData = Array(templateHeaders.length).fill("");
    if (templateSheet.getLastRow() > 1) {
        campaignTemplateRowData = templateSheet.getRange(2, 1, 1, templateLastCol).getValues()[0].map(val => String(val || "").trim());
    }
    while (campaignTemplateRowData.length < templateHeaders.length) { campaignTemplateRowData.push(""); }

    const headerMap = {};
    templateHeaders.forEach((header, index) => { if (header) { headerMap[header.toLowerCase().trim()] = index; } });

    // getIndex funktion från ditt original
    const getIndex = (primaryName, fallbackName = null) => {
      if (!primaryName) return -1;
      const primaryLower = primaryName.toLowerCase().trim();
      if (headerMap.hasOwnProperty(primaryLower)) { return headerMap[primaryLower] + 1; }
      if (fallbackName) { const fallbackLower = fallbackName.toLowerCase().trim(); if (headerMap.hasOwnProperty(fallbackLower)) { return headerMap[fallbackLower] + 1; } }
      Logger.log(`VARNING (getIndex): Varken "${primaryName}"/"${fallbackName || ''}" hittades i CSV Template.`);
      return -1;
    };

    // Kolumnmappning från ditt original, uppdaterad för delade listor
    const col = {
      campaign: getIndex("Campaign", "Kampanj"), adGroup: getIndex("Ad group", "Annonsgrupp"), status: getIndex("Status"),
      criterionType: getIndex("Criterion type", "Kriterietyp"), campaignType: getIndex("Campaign type", "Kampanjtyp"),
      bidStrategyType: getIndex("Bid Strategy Type"), startDate: getIndex("Campaign start date", "Start Date"),
      endDate: getIndex("Campaign end date", "End Date"), budget: getIndex("Campaign daily budget", "Budget"),
      networks: getIndex("Networks", "Nätverk"), languages: getIndex("Languages", "Språk"),
      id: getIndex("ID"), location: getIndex("Location"), maxCpv: getIndex("Max CPV", "Max. CPV"),
      targetCpv: getIndex("Target CPV"), targetCpm: getIndex("Target CPM"),
      adGroupType: getIndex("Ad group type", "Annonsgruppstyp"), gender: getIndex("Gender", "Kön"),
      age: getIndex("Age", "Ålder"), videoId: getIndex("Video ID", "Video-id"), adName: getIndex("Ad name", "Annonsnamn"),
      displayUrl: getIndex("Display URL", "Visningsadress"), finalUrl: getIndex("Final URL", "Slutlig webbadress"),
      trackingTemplate: getIndex("Tracking template", "Mall för spårning"), videoAdStatus: getIndex("Video ad status"),
      adType: getIndex("Ad type"), cta: getIndex("Call to action", "CTA"), headline: getIndex("Headline", "Headline 1"),
      longHeadline: getIndex("Long headline 1"), description1: getIndex("Description", "Description 1"),
      description2: getIndex("Description 2"), inventoryType: getIndex("Inventory type"),
      multiformat: getIndex("Multiformat ads"), optimizedTargeting: getIndex("Optimized targeting"),
      audienceTargeting: getIndex("Audience targeting"), flexibleReach: getIndex("Flexible reach"),
      desktopBidModifier: getIndex("Desktop Bid Modifier"), mobileBidModifier: getIndex("Mobile Bid Modifier"),
      tabletBidModifier: getIndex("Tablet Bid Modifier"), tvBidModifier: getIndex("TV Screen Bid Modifier"),
      negativeTopicId: getIndex("Negative Topic ID", "ID"),
      topic: getIndex("Topic"), topicId: getIndex("Topic ID", "ID"),
      audienceName: getIndex("Audience segment"), budgetPeriod: getIndex("Budget period", "Budget type"),
      // Nya/justerade för delade listor
      sharedSetName: getIndex("Shared set name"),
      sharedSetType: getIndex("Shared set type"),
      // För placement/channels
      placement: getIndex("Placement"),
      channelId: getIndex("Channel ID")
    };

    // DEBUG: Logga alla kolumnmappningar
    Logger.log("=== KOLUMNMAPPNINGAR ===");
    Logger.log(`CSV Template headers: ${templateHeaders.join(', ')}`);
    Logger.log(`Placement kolumn index: ${col.placement}`);
    Logger.log(`Channel ID kolumn index: ${col.channelId}`);
    if (col.channelId === -1) {
      Logger.log("VARNING: Channel ID kolumn hittades inte i CSV Template!");
    }
    if (col.placement === -1) {
      Logger.log("VARNING: Placement kolumn hittades inte i CSV Template!");
    }

    // Validering för delade list-kolumner
    if ((addNegativeKeywords || addNegativePlacements) && (col.sharedSetName === -1 || col.sharedSetType === -1)) {
      return "Error: Kolumnerna 'Shared set name' och/eller 'Shared set type' saknas i din CSV Template, men behövs för att hantera negativa listor.";
    }

    // Ladda negativa topics (din befintliga kod)
    let negativeTopicData = [];
    if (typeof NEGATIVE_TOPICS_SHEET_NAME !== 'undefined' && NEGATIVE_TOPICS_SHEET_NAME) {
      const negTopicSheet = ss.getSheetByName(NEGATIVE_TOPICS_SHEET_NAME);
      if (negTopicSheet) {
        try {
          const negTopicsRange = negTopicSheet.getDataRange();
          const negTopicsValues = negTopicsRange.getValues();
          if (negTopicsValues.length > 1) {
            const negTopicHeaders = negTopicsValues[0].map(h => String(h || "").trim().toLowerCase());
            const idColNegIdx = negTopicHeaders.indexOf("id");
            const topicColNegIdx = negTopicHeaders.indexOf("topic");
            if (idColNegIdx !== -1 && topicColNegIdx !== -1) {
              for (let k = 1; k < negTopicsValues.length; k++) {
                if (negTopicsValues[k][idColNegIdx] && negTopicsValues[k][topicColNegIdx]) {
                  negativeTopicData.push({ id: String(negTopicsValues[k][idColNegIdx]).trim(), topicName: String(negTopicsValues[k][topicColNegIdx]).trim() });
                }
              }
              Logger.log(`generateCsv: Laddade ${negativeTopicData.length} negativa ämnen.`);
            } else { Logger.log(`generateCsv: VARNING - Kolumner ('id', 'topic') för negativa ämnen saknas i '${NEGATIVE_TOPICS_SHEET_NAME}'.`); }
          } else { Logger.log(`generateCsv: INFO - Inga datarader (förutom rubrik) i fliken för negativa ämnen '${NEGATIVE_TOPICS_SHEET_NAME}'.`); }
        } catch (e) { Logger.log(`generateCsv: FEL vid läsning av negativa topics från '${NEGATIVE_TOPICS_SHEET_NAME}': ${e.message}`); }
      } else { Logger.log(`generateCsv: VARNING - Fliken för negativa ämnen ('${NEGATIVE_TOPICS_SHEET_NAME}') hittades inte.`); }
    }

    // --- Läs mappningar från Excl lists NL ---
    const cfSharedListMappings = {};
    if ((addNegativeKeywords || addNegativePlacements) && typeof NEGATIVE_LISTS_SHEET_NAME !== 'undefined' && NEGATIVE_LISTS_SHEET_NAME) {
      const exclSheet = ss.getSheetByName(NEGATIVE_LISTS_SHEET_NAME);
      if (exclSheet) {
        const data = exclSheet.getDataRange().getValues();
        const HEADER_ROW_INDEX_EXCL = 0; 
        const DATA_START_ROW_INDEX_EXCL = HEADER_ROW_INDEX_EXCL + 1;

        if (data.length >= DATA_START_ROW_INDEX_EXCL) {
          const exclHeaders = data[HEADER_ROW_INDEX_EXCL].map(h => String(h || "").trim().toLowerCase());
          
          const cfApplicabilityColHeader = "cf1, cf2, cf3, cf4, cf5, cf6, cf7"; 
          const sharedSetTypeColHeader = "shared set type";                     
          const sharedSetNameColHeader = "shared set name";                     

          const cfColIdx = exclHeaders.indexOf(cfApplicabilityColHeader.toLowerCase());
          const typeColIdx = exclHeaders.indexOf(sharedSetTypeColHeader.toLowerCase());
          const nameColIdx = exclHeaders.indexOf(sharedSetNameColHeader.toLowerCase());

          if (cfColIdx !== -1 && typeColIdx !== -1 && nameColIdx !== -1) {
            for (let i = DATA_START_ROW_INDEX_EXCL; i < data.length; i++) {
              const row = data[i];
              const cfApplicabilityStr = String(row[cfColIdx] || "").trim().toUpperCase();
              const sharedSetTypeStrInput = String(row[typeColIdx] || "").trim().toLowerCase();
              const sharedSetNameStr = String(row[nameColIdx] || "").trim();

              if (cfApplicabilityStr && sharedSetTypeStrInput && sharedSetNameStr) {
                let correctSharedSetType = ""; 
                if (sharedSetTypeStrInput.includes("keyword")) {
                    correctSharedSetType = "Shared negative keyword list"; 
                } else if (sharedSetTypeStrInput.includes("placement")) {
                    correctSharedSetType = "Shared negative placement list"; 
                }

                if (correctSharedSetType) {
                    const cfCodesForRow = cfApplicabilityStr.split(/[;, ]+/).map(cfToken => cfToken.trim()).filter(cfToken => cfToken);
                    cfCodesForRow.forEach(cfCode => {
                        // Lista av CF-koder att applicera negativa listor på (inklusive varianter)
                        let cfCodesToApply = [cfCode];
                        
                        // Om bas-CF1 eller CF2 hittas, lägg även till deras varianter
                        if (cfCode === 'CF1') {
                            cfCodesToApply.push('CF1 NM', 'CF1 VM');
                        } else if (cfCode === 'CF2') {
                            cfCodesToApply.push('CF2 NM', 'CF2 VM');
                        }
                        
                        Logger.log(`NEGATIVE LISTS: CF '${cfCode}' expanderas till: ${cfCodesToApply.join(', ')}`);
                        
                        // Applicera negativa listor på alla CF-koder (bas + varianter)
                        cfCodesToApply.forEach(targetCf => {
                            if (!cfSharedListMappings[targetCf]) {
                                cfSharedListMappings[targetCf] = { keywordLists: [], placementLists: [] };
                            }
                            if (correctSharedSetType === "Shared negative keyword list") { 
                                if (cfSharedListMappings[targetCf].keywordLists.indexOf(sharedSetNameStr) === -1) {
                                    cfSharedListMappings[targetCf].keywordLists.push(sharedSetNameStr);
                                }
                            } else if (correctSharedSetType === "Shared negative placement list") { 
                                if (cfSharedListMappings[targetCf].placementLists.indexOf(sharedSetNameStr) === -1) {
                                    cfSharedListMappings[targetCf].placementLists.push(sharedSetNameStr);
                                }
                            }
                        });
                    });
                }
              }
            }
            Logger.log(`Hämtade mappningar för delade listor: ${JSON.stringify(cfSharedListMappings)}`);
          } else {
            let missingHeadersLog = [];
            if(cfColIdx === -1) missingHeadersLog.push(`'${cfApplicabilityColHeader}'`);
            if(typeColIdx === -1) missingHeadersLog.push(`'${sharedSetTypeColHeader}'`);
            if(nameColIdx === -1) missingHeadersLog.push(`'${sharedSetNameColHeader}'`);
            Logger.log(`VARNING: Kunde inte hitta alla nödvändiga rubriker (${missingHeadersLog.join(", ")}) på rad ${HEADER_ROW_INDEX_EXCL + 1} i '${NEGATIVE_LISTS_SHEET_NAME}'.`);
          }
        }
      } else {
        Logger.log(`VARNING: Arket '${NEGATIVE_LISTS_SHEET_NAME}' hittades inte.`);
      }
    }
    
    // Ladda positiv targeting data (din befintliga kod)
    let columnarTargetingData = {};
    if (!importAdsOnlyMode) {
      const targetingSheet = ss.getSheetByName(TARGETING_SHEET); // Använder din globala konstant
      if (targetingSheet) {
        const targetingSheetValues = targetingSheet.getDataRange().getValues();
        if (targetingSheetValues.length > 1) {
          const sheetHeaders = targetingSheetValues[0].map(h => String(h || "").trim());
          for (const typeKeyInMasterConfig in MASTER_SHEETS_CONFIG) {
            const config = MASTER_SHEETS_CONFIG[typeKeyInMasterConfig];
            const nameColHeader = config.nameCol;
            const idColHeader = config.idCol;
            const nameColIdx = sheetHeaders.indexOf(nameColHeader);
            const idColIdx = sheetHeaders.indexOf(idColHeader);
            const dataStorageKey = config.name;
            if (nameColIdx !== -1) {
              columnarTargetingData[dataStorageKey] = [];
              for (let rowIndex = 1; rowIndex < targetingSheetValues.length; rowIndex++) {
                const nameValue = String(targetingSheetValues[rowIndex][nameColIdx] || "").trim();
                let idValue = "";
                if (idColIdx !== -1) {
                  idValue = String(targetingSheetValues[rowIndex][idColIdx] || "").trim();
                }
                if (nameValue) {
                  columnarTargetingData[dataStorageKey].push({ name: nameValue, id: idValue });
                }
              }
              if (idColIdx === -1 && columnarTargetingData[dataStorageKey] && columnarTargetingData[dataStorageKey].length > 0) {
                Logger.log(`INFO: ID column "${idColHeader}" not found for type "${dataStorageKey}" in '${TARGETING_SHEET}'. Only names were loaded.`);
              }
              Logger.log(`Loaded ${columnarTargetingData[dataStorageKey] ? columnarTargetingData[dataStorageKey].length : 0} items for type "${dataStorageKey}" from '${TARGETING_SHEET}'.`);
            } else {
              Logger.log(`WARNING: Name column "${nameColHeader}" for type "${dataStorageKey}" not found in '${TARGETING_SHEET}'.`);
            }
          }
        }
        Logger.log(`Targeting data loaded for keys: ${Object.keys(columnarTargetingData).join(", ")}`);
      } else { Logger.log(`WARNING: Sheet '${TARGETING_SHEET}' not found.`); }
    }

    const csvData = [];
    csvData.push(templateHeaders);

    const defaultMaxCpv_placeholder = "0.50";
    const defaultAdGroupStatus = "Paused";
    const defaultAdStatus = "Enabled";
    const fallbackAdGroupType = "Video";
    const defaultCriterionStatus = "Enabled";
    const DEFAULT_TARGETING_SETTING = "Targeting";

    // Targeting Only Mode (din befintliga kod, från ditt original)
    if (targetingOnlyMode) {
      Logger.log("=== TARGETING ONLY MODE AKTIVERAT ===");
      Logger.log(`Valda kampanjer: ${JSON.stringify(selections.selectedCampaigns)}`);
      Logger.log(`Valda CF:er: ${JSON.stringify(selections.selectedCFs)}`);
      Logger.log(`Valda AdGroup typer: ${JSON.stringify(selections.selectedAdGroupTypes)}`);
      Logger.log(`Add Channels: ${selections.addChannels}`);
      Logger.log(`CF Variant: ${JSON.stringify(selections.cfVariant)}`);
      
      const targetingInputSheet = ss.getSheetByName(TARGETING_SHEET); // Använder din globala konstant
      if (!targetingInputSheet) {
        Logger.log("Error: Targeting Input sheet 'Targeting' not found.");
        return "Error: Targeting Input sheet 'Targeting' not found.";
      }
      const inputData = targetingInputSheet.getDataRange().getValues();
      if (inputData.length <= 1) {
        Logger.log("Error: Targeting Input sheet has no data rows.");
        return "Error: No data rows found in the Targeting Input sheet.";
      }
      const inputHeaders = inputData[0].map(h => String(h || "").trim().toLowerCase());
      const topicsIdColIdx = inputHeaders.indexOf("topics id");
      const affinitiesIdColIdx = inputHeaders.indexOf("affinities id");
      const inMarketIdColIdx = inputHeaders.indexOf("in-market id");
      const detailedDemographicsIdColIdx = inputHeaders.indexOf("detailed demographics id");
      const lifeEventsIdColIdx = inputHeaders.indexOf("life events id");
      const idColumns = [
          { idx: topicsIdColIdx, type: "Topics", criterionType: "Topics", targetingType: "topic", label: "Topics ID" },
          { idx: affinitiesIdColIdx, type: "Affinity", criterionType: "Affinity", targetingType: "audience", label: "Affinities ID" },
          { idx: inMarketIdColIdx, type: "In-market", criterionType: "In-market", targetingType: "audience", label: "In-market ID" },
          { idx: detailedDemographicsIdColIdx, type: "Detailed demographics", criterionType: "Detailed demographics", targetingType: "audience", label: "Detailed demographics ID" },
          { idx: lifeEventsIdColIdx, type: "Life Event", criterionType: "Life Event", targetingType: "audience", label: "Life events ID" }
      ].filter(colDef => colDef.idx !== -1);

      if (idColumns.length === 0) {
        Logger.log("Warning: No targeting ID columns found in Targeting Input sheet.");
        return "Error: No targeting ID columns found in the Targeting Input sheet.";
      }

      let addedRows = 0;
      const selectedCampaigns = selections.selectedCampaigns || [];
      if (selectedCampaigns.length === 0) {
        Logger.log("No campaigns selected in Targeting Only mode.");
        return "Error: No campaigns selected for Targeting Only mode.";
      }
      
      selectedCampaigns.forEach(campaignData => {
          const baseCampaignName = campaignData.name;
          if (!baseCampaignName) return;
          const allAds = campaignData.ads || [];
          const evenAds = allAds.filter(ad => String(ad.rotation || "").trim().toLowerCase() === 'even');
          const normalAds = allAds.filter(ad => String(ad.rotation || "").trim().toLowerCase() !== 'even');

          Logger.log(`=== PROCESSING CF: ${cf} ===`);
           (selections.selectedCFs || []).forEach(cf => {
               // CF innehåller nu fullständiga namn som "CF1 NM", "CF1 VM" etc
               const targetCampaignNames = [];
               if (normalAds.length > 0 || allAds.length === 0) {
                   targetCampaignNames.push(`${baseCampaignName} ${cf}`);
               }
               evenAds.forEach((adData, evenIndex) => {
                   targetCampaignNames.push(`${baseCampaignName} ${cf} - ${adData.videoId || `EVEN_AD_NO_ID_${evenIndex + 1}`}`);
               });
              
              targetCampaignNames.forEach(fullCampaignName => {
                      Logger.log(`=== PROCESSING AD GROUP TYPE: ${adGroupType} för CF: ${cf} ===`);
                      (selections.selectedAdGroupTypes || []).forEach(adGroupType => {
                      const lowerCaseAdGroupType = adGroupType.toLowerCase();
                      
                      // Extrahera bas-CF för regelkontroll (t.ex. "CF1" från "CF1 NM")
                      const baseCf = cf.split(' ')[0];
                      if (lowerCaseAdGroupType === 'topics' && baseCf !== 'CF3' && baseCf !== 'CF5') { Logger.log(`TARGETING: Skipping 'Topics' for ${fullCampaignName} (Not CF3 or CF5)`); return; }
                      if (lowerCaseAdGroupType === 'comscore' && baseCf !== 'CF1') { Logger.log(`TARGETING: Skipping 'Comscore' for ${fullCampaignName} (Not CF1)`); return; }
                      if (lowerCaseAdGroupType === 'keywords' && baseCf !== 'CF3' && baseCf !== 'CF5') { Logger.log(`TARGETING: Skipping 'Keywords' for ${fullCampaignName} (Not CF3 or CF5)`); return; }
                       if (lowerCaseAdGroupType === 'channels' && (baseCf === 'CF3' || baseCf === 'CF5')) { Logger.log(`TARGETING: Skipping 'Channels' for ${fullCampaignName} (Is CF3 or CF5)`); return; }
                       if (baseCf === 'CF7' && lowerCaseAdGroupType !== 'channels' && lowerCaseAdGroupType !== 'demo') { Logger.log(`TARGETING: Skipping '${adGroupType}' for ${fullCampaignName} (Only Channels/Demo allowed)`); return; }
                       
                         // Hantera channels från Filtered_* sheets om "channels" är valt och addChannels är aktiverat
                         if (lowerCaseAdGroupType === 'channels' && selections.addChannels) {
                              Logger.log(`TARGETING: Hanterar channels för CF: "${cf}", AG: "${adGroupType}", addChannels: ${selections.addChannels}`);
                              Logger.log(`TARGETING: Exakt CF-värde skickas till getChannelsForCF_: "${cf}"`);
                              const channelsData = getChannelsForCF_(cf, selections.cfVariant);
                              Logger.log(`TARGETING: Fick ${channelsData ? channelsData.length : 0} channels från getChannelsForCF_ för CF: "${cf}"`);
                              Logger.log(`TARGETING: Förväntat sheet namn: Filtered_${cf.replace(/ /g, '_')}`);
                             
                              if (channelsData && channelsData.length > 0) {
                                channelsData.forEach((channelData, index) => {
                                    const channelRow = createEmptyRow_(templateHeaders.length);
                                    setValue_(channelRow, col.campaign, fullCampaignName);
                                    setValue_(channelRow, col.adGroup, adGroupType);
                                    setValue_(channelRow, col.status, defaultCriterionStatus);
                                    setValue_(channelRow, col.criterionType, "Placement");
                                    // Sätt Placement med URL (kolumn A) och Channel ID med Channel ID (kolumn B)
                                    if (col.placement !== -1) {
                                        setValue_(channelRow, col.placement, channelData.url);
                                        Logger.log(`TARGETING: Satte Placement (col ${col.placement}) till: ${channelData.url}`);
                                    }
                                    if (col.channelId !== -1) {
                                        setValue_(channelRow, col.channelId, channelData.channelId);
                                        Logger.log(`TARGETING: Satte Channel ID (col ${col.channelId}) till: ${channelData.channelId}`);
                                    }
                                    csvData.push(channelRow);
                                    addedRows++;
                                    Logger.log(`TARGETING: Lade till channel rad ${index + 1}/${channelsData.length} för ${fullCampaignName} - ${adGroupType}`);
                                });
                                Logger.log(`TARGETING: SLUTFÖRT - Added ${channelsData.length} channels for CF: ${cf}, AG: ${adGroupType}, Campaign: ${fullCampaignName}`);
                            } else {
                                Logger.log(`TARGETING: VARNING - Inga channels hittades för CF: ${cf}, Campaign: ${fullCampaignName}, AdGroup: ${adGroupType}. Kontrollera att sheet 'Filtered_${cf.replace(' ', '_')}' finns och innehåller data.`);
                            }
                            return; // Fortsätt inte med vanlig targeting för channels
                        }
                      
                      for (let i = 1; i < inputData.length; i++) {
                          for (const idCol of idColumns) {
                              const idValue = String(inputData[i][idCol.idx] || "").trim();
                              if (idValue) {
                                  let matchFound = false;
                                  if (idCol.label.toLowerCase().includes(lowerCaseAdGroupType)) { matchFound = true; }
                                  else if (lowerCaseAdGroupType === "topics" && idCol.label.toLowerCase().includes("topics")) { matchFound = true; }
                                  else if (lowerCaseAdGroupType === "affinities" && idCol.label.toLowerCase().includes("affinities")) { matchFound = true; }
                                  else if (lowerCaseAdGroupType === "in-market" && idCol.label.toLowerCase().includes("in-market")) { matchFound = true; }
                                  else if (lowerCaseAdGroupType === "detailed demographics" && idCol.label.toLowerCase().includes("detailed demographics")) { matchFound = true; }
                                  else if (lowerCaseAdGroupType === "life events" && idCol.label.toLowerCase().includes("life events")) { matchFound = true; }
                                  
                                  if (matchFound) {
                                      const targetRow = createEmptyRow_(templateHeaders.length);
                                      setValue_(targetRow, col.campaign, fullCampaignName);
                                      setValue_(targetRow, col.adGroup, adGroupType);
                                      setValue_(targetRow, col.id, idValue);
                                      setValue_(targetRow, col.status, defaultCriterionStatus);
                                      if (idCol.targetingType === "topic") {
                                          setValue_(targetRow, col.criterionType, "Topic");
                                          setValue_(targetRow, col.topic, "Topics"); 
                                      } else if (idCol.targetingType === "audience") {
                                          setValue_(targetRow, col.criterionType, "Audience");
                                          setValue_(targetRow, col.audienceName, idCol.criterionType);
                                      }
                                      csvData.push(targetRow);
                                      addedRows++;
                                  }
                              }
                          }
                      }
                  });
              });
          });
      });
      if (addedRows === 0) {
        return "Error: No targeting data rows were generated. Check your Targeting Input sheet data and selected campaign/ad group types.";
      }
      Logger.log(`Targeting Only Mode: Added ${addedRows} targeting rows to the CSV.`);

    } else if (channelsOnlyMode) {
      Logger.log("=== CHANNELS ONLY MODE SUPER-OPTIMERAD ===");
      Logger.log(`Valda kampanjer: ${selections.selectedCampaigns?.length || 0}`);
      Logger.log(`Valda CF:er: ${selections.selectedCFs?.length || 0}`);
      Logger.log(`Valda AdGroup typer: ${selections.selectedAdGroupTypes?.length || 0}`);
      
      let addedChannelRows = 0;
      const selectedCampaigns = selections.selectedCampaigns || [];
      
      if (selectedCampaigns.length === 0) {
        return "Error: No campaigns selected for Channels Only mode.";
      }
      
      // ULTRA-OPTIMERING: Förhämta alla channels för valda CF:er
      const cfChannelsCache = {};
      console.log("CHANNELS ONLY: Snabbladdning av channels...");
      
      (selections.selectedCFs || []).forEach(cf => {
        if (!cfChannelsCache[cf]) {
          cfChannelsCache[cf] = getChannelsForCF_(cf, selections.cfVariant);
          // Endast console.log för stora dataset (snabbare än Logger.log)
          if (cfChannelsCache[cf]?.length > 5000) {
            console.log(`Loaded ${cfChannelsCache[cf].length} channels for ${cf}`);
          }
        }
      });
      
      // MEGA-BATCH bearbetning: Processera allt i minsta möjliga loopar
      selectedCampaigns.forEach(campaignData => {
        const baseCampaignName = campaignData.name;
        if (!baseCampaignName) return;
        
        const allAds = campaignData.ads || [];
        const evenAds = allAds.filter(ad => String(ad.rotation || "").trim().toLowerCase() === 'even');
        const normalAds = allAds.filter(ad => String(ad.rotation || "").trim().toLowerCase() !== 'even');
        
        (selections.selectedCFs || []).forEach(cf => {
          const channelsData = cfChannelsCache[cf];
          if (!channelsData?.length) return; // Snabb exit utan logging
          
          const targetCampaignNames = [];
          
          // Generera kampanjnamn
          if (normalAds.length > 0 || allAds.length === 0) {
            targetCampaignNames.push(`${baseCampaignName} ${cf}`);
          }
          evenAds.forEach((adData, evenIndex) => {
            targetCampaignNames.push(`${baseCampaignName} ${cf} - ${adData.videoId || `EVEN_AD_NO_ID_${evenIndex + 1}`}`);
          });
          
          // ULTRA-SNABB: Pre-beräkna storlek och batch-skapa allt
          const totalRowsForThisCF = targetCampaignNames.length * (selections.selectedAdGroupTypes?.length || 0) * channelsData.length;
          const batchRows = new Array(totalRowsForThisCF);
          let batchIndex = 0;
          
          targetCampaignNames.forEach(fullCampaignName => {
            (selections.selectedAdGroupTypes || []).forEach(adGroupType => {
              // Skapa alla channel rows för denna kombination utan extra logging
              channelsData.forEach(channelData => {
                const channelRow = createEmptyRow_(templateHeaders.length);
                setValue_(channelRow, col.campaign, fullCampaignName);
                setValue_(channelRow, col.adGroup, adGroupType);
                setValue_(channelRow, col.status, defaultCriterionStatus);
                setValue_(channelRow, col.criterionType, "Placement");
                
                if (col.placement !== -1) setValue_(channelRow, col.placement, channelData.url);
                if (col.channelId !== -1) setValue_(channelRow, col.channelId, channelData.channelId);
                
                batchRows[batchIndex++] = channelRow;
              });
            });
          });
          
          // Bulk-lägg till alla rows (mycket snabbare än push i loop)
          csvData.push(...batchRows);
          addedChannelRows += batchRows.length;
          
          // Minimal logging för stora CF:er
          if (batchRows.length > 10000) {
            console.log(`Added ${batchRows.length} rows for CF: ${cf}`);
          }
        });
      });
      
      if (addedChannelRows === 0) {
        return "Error: No channel data rows were generated. Check your Filtered_* sheets and selected options.";
      }
      Logger.log(`CHANNELS ONLY SUPER-OPTIMERAD: ${addedChannelRows} channel rows generated`);

    } else { // Full/Ads Only Mode
      (selections.selectedCampaigns || []).forEach(campaignDataFromPopup => {
        const baseCampaignName = campaignDataFromPopup.name;
        if (!baseCampaignName) { return; }

        // Fullständig baseMetadata definition från ditt original
        const baseMetadata = {
            campaign: baseCampaignName, 
            geoId: campaignDataFromPopup.geoId, 
            locationName: campaignDataFromPopup.locationName || "",
            language: campaignDataFromPopup.language, 
            campaignType: campaignDataFromPopup.campaignType, 
            startDate: campaignDataFromPopup.startDate,
            endDate: campaignDataFromPopup.endDate, 
            bidStrategyType: campaignDataFromPopup.bidStrategyType,
            clientRatePercent: campaignDataFromPopup.clientRatePercent || "80",
            deviceStates: campaignDataFromPopup.deviceStates || { desktop: true, mobile: true, tablet: true, tv: true }
        };
        const allAds = campaignDataFromPopup.ads || [];
        const evenAds = allAds.filter(ad => String(ad.rotation || "").trim().toLowerCase() === 'even');
        const normalAds = allAds.filter(ad => String(ad.rotation || "").trim().toLowerCase() !== 'even');
        Logger.log(`Split Results for "${baseCampaignName}": Even Ads (${evenAds.length}), Normal Ads (${normalAds.length})`);
          
        if (allAds.length === 0 && selections.importMode !== 'adsOnly') { 
           Logger.log(`(Full Build) INFO: No ads for "${baseCampaignName}", but structure can be created if Ad Group Types are selected.`);
        } else if (allAds.length === 0 && selections.importMode === 'adsOnly'){
           Logger.log(`(Ads Only) INFO: No ads to add for "${baseCampaignName}". Skipping.`);
           return;
        }

        selections.selectedCFs.forEach(cf => {
          const currentCfMapping = cfSharedListMappings[cf] || { keywordLists: [], placementLists: [] };

          if (normalAds.length > 0 || (evenAds.length === 0 && normalAds.length === 0 && !importAdsOnlyMode)) {
            const campaignNameForCsv_Normal = `${baseCampaignName} ${cf}`;
            generateCampaignAndAdGroupRows_(csvData, campaignNameForCsv_Normal, baseMetadata, cf,
              normalAds.length > 0 ? normalAds : [], selections, col, campaignTemplateRowData,
              fallbackAdGroupType, defaultMaxCpv_placeholder, defaultAdGroupStatus, defaultCriterionStatus,
              defaultAdStatus, negativeTopicData, importAdsOnlyMode ? {} : columnarTargetingData, DEFAULT_TARGETING_SETTING
            );

            if (addNegativeKeywords && currentCfMapping.keywordLists.length > 0) {
              currentCfMapping.keywordLists.forEach(listName => {
                const assocRow = createEmptyRow_(templateHeaders.length);
                setValue_(assocRow, col.campaign, campaignNameForCsv_Normal);
                setValue_(assocRow, col.sharedSetName, listName);
                setValue_(assocRow, col.sharedSetType, "Shared negative keyword list");
                csvData.push(assocRow);
              });
            }
            if (addNegativePlacements && currentCfMapping.placementLists.length > 0 && cf !== "CF1" && cf !== "CF2" && cf !== "CF7") {
              currentCfMapping.placementLists.forEach(listName => {
                const assocRow = createEmptyRow_(templateHeaders.length);
                setValue_(assocRow, col.campaign, campaignNameForCsv_Normal);
                setValue_(assocRow, col.sharedSetName, listName);
                setValue_(assocRow, col.sharedSetType, "Shared negative placement list");
                csvData.push(assocRow);
              });
            }
          }

          if (evenAds.length > 0) {
            evenAds.forEach((adData, evenIndex) => {
              let campaignNameForCsv_Even = `${baseCampaignName} ${cf} - ${adData.videoId || `EVEN_AD_NO_ID_${evenIndex + 1}`}`;
              generateCampaignAndAdGroupRows_(csvData, campaignNameForCsv_Even, baseMetadata, cf,
                [adData], selections, col, campaignTemplateRowData, fallbackAdGroupType,
                defaultMaxCpv_placeholder, defaultAdGroupStatus, defaultCriterionStatus,
                defaultAdStatus, negativeTopicData, importAdsOnlyMode ? {} : columnarTargetingData, DEFAULT_TARGETING_SETTING
              );
              
              if (addNegativeKeywords && currentCfMapping.keywordLists.length > 0) {
                currentCfMapping.keywordLists.forEach(listName => {
                  const assocRow = createEmptyRow_(templateHeaders.length);
                  setValue_(assocRow, col.campaign, campaignNameForCsv_Even);
                  setValue_(assocRow, col.sharedSetName, listName);
                  setValue_(assocRow, col.sharedSetType, "Shared negative keyword list");
                  csvData.push(assocRow);
                });
              }
              if (addNegativePlacements && currentCfMapping.placementLists.length > 0 && cf !== "CF1" && cf !== "CF2" && cf !== "CF7") {
                 currentCfMapping.placementLists.forEach(listName => {
                  const assocRow = createEmptyRow_(templateHeaders.length);
                  setValue_(assocRow, col.campaign, campaignNameForCsv_Even);
                  setValue_(assocRow, col.sharedSetName, listName);
                  setValue_(assocRow, col.sharedSetType, "Shared negative placement list");
                  csvData.push(assocRow);
                });
              }
            });
          }
        });
      });
    }

    if (csvData.length <= 1 ) { 
        return "Error: No data rows were generated (excluding headers). Check logs and Excl lists NL structure.";
    }
    const csvContent = csvData.map(rowArray => rowArray.map(cellValue => escapeCsvCell_(cellValue)).join(',')).join('\n');
    Logger.log(`generateCsv: CSV genererad, längd: ${csvContent.length}.`);
    return csvContent;

  } catch (e) {
    Logger.log(`FATAL Error in generateCsv: ${e.message}\nStack: ${e.stack || 'Ingen stack tillgänglig'}`);
    return `Error generating CSV: ${e.message}.`;
  }
}

/**
 * ===============================================
 * HJÄLPFUNKTION för generateCsv - Skapar rader för en kampanj/CF-kombination.
 * **UPPDATERAD: Hanterar positiva och negativa (topics, keywords, placements) listor.**
 * @param {Array<Array<string>>} csvData
 * @param {string} campaignNameForCsv
 * @param {object} baseMetadata
 * @param {string} cf
 * @param {Array<object>} adsForThisCampaign
 * @param {object} selections
 * @param {object} col
 * @param {Array<string>} campaignTemplateRowData
 * @param {string} fallbackAdGroupType
 * @param {string} defaultMaxCpv_unused
 * @param {string} defaultAdGroupStatus
 * @param {string} defaultCriterionStatus
 * @param {string} defaultAdStatus
 * @param {Array<object>} negativeTopicData
 * @param {object} columnarTargetingData
 * @param {string} defaultTargetingSetting
 * @param {Array<string>} negativeKeywordData - NY
 * @param {Array<string>} negativePlacementData - NY
 * ===============================================
 */

/**
 * ===============================================
 * HJÄLPFUNKTION för generateCsv - Skapar rader för en kampanj/CF-kombination.
 * (Detta är den återställda originalversionen som förutsätts)
 * ===============================================
 */
function generateCampaignAndAdGroupRows_(csvData, campaignNameForCsv, baseMetadata, cf, adsForThisCampaign, selections, col, campaignTemplateRowData, fallbackAdGroupType, defaultMaxCpv_unused, defaultAdGroupStatus, defaultCriterionStatus, defaultAdStatus, negativeTopicData, columnarTargetingData, defaultTargetingSetting) {
  Logger.log(`generateCampaignAndAdGroupRows_: C: "${campaignNameForCsv}" (CF: ${cf}), ads: ${adsForThisCampaign ? adsForThisCampaign.length : 0}`);
  const importMode = selections.importMode || 'full';
  const targetingOnlyMode = selections.targetingOnlyMode || false;

  let campaignRow = null;
  let createdLocationRows = [];

  if (importMode !== 'adsOnly' && !targetingOnlyMode) {
    campaignRow = Array.isArray(campaignTemplateRowData) ? [...campaignTemplateRowData] : createEmptyRow_(campaignTemplateRowData.length); // Använder campaignTemplateRowData.length

    setValue_(campaignRow, col.campaign, campaignNameForCsv);
    setValue_(campaignRow, col.startDate, formatDateForAds_(baseMetadata.startDate));
    setValue_(campaignRow, col.endDate, formatDateForAds_(baseMetadata.endDate));
    setValue_(campaignRow, col.campaignType, baseMetadata.campaignType);
    setValue_(campaignRow, col.bidStrategyType, baseMetadata.bidStrategyType);
    let languageValue = baseMetadata.language || "en";
    if (cf === 'CF1' || cf === 'CF7') languageValue = 'all';
    setValue_(campaignRow, col.languages, languageValue);
    let inventoryValue = 'Standard inventory';
    if (cf === 'CF3') inventoryValue = 'Limited inventory';
    setValue_(campaignRow, col.inventoryType, inventoryValue);
    if (col.multiformat !== -1 && adsForThisCampaign.length > 0 && adsForThisCampaign[0] && adsForThisCampaign[0].multiformat) {
      setValue_(campaignRow, col.multiformat, adsForThisCampaign[0].multiformat);
    }
    if (col.flexibleReach !== -1) {
      setValue_(campaignRow, col.flexibleReach, "[]");
    }
    if (col.networks !== -1) {
      setValue_(campaignRow, col.networks, "Youtube;YouTube Videos");
    }
    if (col.budget !== -1) {
      setValue_(campaignRow, col.budget, "1");
    }
    if (col.budgetPeriod && col.budgetPeriod !== -1) {
      setValue_(campaignRow, col.budgetPeriod, "Daily");
    }

    const geoIdString = baseMetadata.geoId || "";
    const locationNameString = baseMetadata.locationName || "";
    const containsMultipleLocations = geoIdString.includes(";");
    if (containsMultipleLocations) {
      if (col.id !== -1) setValue_(campaignRow, col.id, "");
      if (col.location !== -1) setValue_(campaignRow, col.location, "");
      const locationIds = geoIdString.split(';').map(id => id.trim()).filter(id => id);
      if (locationIds.length > 0 && col.campaign !== -1 && col.criterionType !== -1 && col.id !== -1 && col.status !== -1) {
        locationIds.forEach(locId => {
          const locationRow = createEmptyRow_(campaignTemplateRowData.length);
          setValue_(locationRow, col.campaign, campaignNameForCsv);
          setValue_(locationRow, col.criterionType, "Location");
          setValue_(locationRow, col.id, locId);
          setValue_(locationRow, col.status, defaultCriterionStatus);
          createdLocationRows.push(locationRow);
        });
      }
    } else if (geoIdString) {
      if (col.id !== -1) setValue_(campaignRow, col.id, geoIdString);
      if (col.location !== -1) setValue_(campaignRow, col.location, locationNameString);
    } else {
      if (col.id !== -1) setValue_(campaignRow, col.id, "");
      if (col.location !== -1) setValue_(campaignRow, col.location, "");
    }
    const deviceStates = baseMetadata.deviceStates;
    const bidModifierValue = -100;
    if (deviceStates) {
      if (!deviceStates.desktop && col.desktopBidModifier !== -1 && typeof col.desktopBidModifier === 'number') { setValue_(campaignRow, col.desktopBidModifier, bidModifierValue); }
      if (!deviceStates.mobile && col.mobileBidModifier !== -1 && typeof col.mobileBidModifier === 'number') { setValue_(campaignRow, col.mobileBidModifier, bidModifierValue); }
      if (!deviceStates.tablet && col.tabletBidModifier !== -1 && typeof col.tabletBidModifier === 'number') { setValue_(campaignRow, col.tabletBidModifier, bidModifierValue); }
      if (!deviceStates.tv && col.tvBidModifier !== -1 && typeof col.tvBidModifier === 'number') { setValue_(campaignRow, col.tvBidModifier, bidModifierValue); }
    }
    csvData.push(campaignRow);
    if (createdLocationRows.length > 0) {
      createdLocationRows.forEach(row => csvData.push(row));
    }

    const topicIdCsvColForNegativeCampaign = (col.negativeTopicId && col.negativeTopicId !== -1) ? col.negativeTopicId : (col.topicId && col.topicId !== -1 ? col.topicId : col.id);
    if (cf === 'CF3' && negativeTopicData && negativeTopicData.length > 0) {
      if (col.campaign !== -1 && col.criterionType !== -1 && topicIdCsvColForNegativeCampaign !== -1 && col.topic !== -1 && col.status !== -1) {
        Logger.log(`  (Full Build C: "${campaignNameForCsv}") Adding ${negativeTopicData.length} negative topics for CF3.`);
        negativeTopicData.forEach(negTopic => {
          const negTopicRow = createEmptyRow_(campaignTemplateRowData.length);
          setValue_(negTopicRow, col.campaign, campaignNameForCsv);
          setValue_(negTopicRow, col.adGroup, "");
          setValue_(negTopicRow, col.criterionType, 'Campaign Negative Topic');
          setValue_(negTopicRow, topicIdCsvColForNegativeCampaign, negTopic.id);
          setValue_(negTopicRow, col.topic, negTopic.topicName);
          setValue_(negTopicRow, col.status, defaultCriterionStatus);
          csvData.push(negTopicRow);
        });
      } else {
        Logger.log(`  WARNING: Cannot add negative topics for CF3 for C: "${campaignNameForCsv}". Required columns might be missing.`);
      }
    }
  }

  (selections.selectedAdGroupTypes || []).forEach(adGroupTypeNameSelected => {
    const lowerCaseAdGroupType = adGroupTypeNameSelected.toLowerCase();
    if (lowerCaseAdGroupType === 'topics' && cf !== 'CF3' && cf !== 'CF5') { return; }
    if (lowerCaseAdGroupType === 'comscore' && cf !== 'CF1') { return; }
    if (lowerCaseAdGroupType === 'keywords' && cf !== 'CF3' && cf !== 'CF5') { return; }
    if (lowerCaseAdGroupType === 'channels' && (cf === 'CF3' || cf === 'CF5')) { return; }
    if (cf === 'CF7' && lowerCaseAdGroupType !== 'channels' && lowerCaseAdGroupType !== 'demo') { return; }

    // Lägg till channels från Filtered_* sheets om "channels" är valt och addChannels är aktiverat
    if (lowerCaseAdGroupType === 'channels' && selections.addChannels) {
      Logger.log(`generateCampaignAndAdGroupRows_: Hanterar channels för CF: "${cf}", AG: "${adGroupTypeNameSelected}", addChannels: ${selections.addChannels}`);
      Logger.log(`generateCampaignAndAdGroupRows_: Rensat CF-namn innan getChannelsForCF_: "${cf.trim()}"`);
      const channelsData = getChannelsForCF_(cf, selections.cfVariant);
      Logger.log(`generateCampaignAndAdGroupRows_: Fick ${channelsData ? channelsData.length : 0} channels från getChannelsForCF_`);
      Logger.log(`generateCampaignAndAdGroupRows_: Sheet name som används: Filtered_${cf.replace(/ /g, '_')} (från CF: "${cf}")`);
      
      if (channelsData && channelsData.length > 0) {
        channelsData.forEach((channelData, index) => {
          const channelRow = createEmptyRow_(campaignTemplateRowData.length);
          setValue_(channelRow, col.campaign, campaignNameForCsv);
          setValue_(channelRow, col.adGroup, adGroupTypeNameSelected);
          setValue_(channelRow, col.adGroupType, determinedAdGroupTypeForCSV);
          setValue_(channelRow, col.status, defaultCriterionStatus);
          setValue_(channelRow, col.criterionType, "Placement");
          // Sätt Placement med URL (kolumn A) och Channel ID med Channel ID (kolumn B)
          if (col.placement !== -1) {
            setValue_(channelRow, col.placement, channelData.url);
            Logger.log(`generateCampaignAndAdGroupRows_: Satte Placement (col ${col.placement}) till: ${channelData.url}`);
          }
          if (col.channelId !== -1) {
            setValue_(channelRow, col.channelId, channelData.channelId);
            Logger.log(`generateCampaignAndAdGroupRows_: Satte Channel ID (col ${col.channelId}) till: ${channelData.channelId}`);
          }
          setValue_(channelRow, col.targetingSetting, defaultTargetingSetting);
          csvData.push(channelRow);
          Logger.log(`generateCampaignAndAdGroupRows_: Lade till channel rad ${index + 1}/${channelsData.length} för ${campaignNameForCsv} - ${adGroupTypeNameSelected}`);
        });
        Logger.log(`generateCampaignAndAdGroupRows_: SLUTFÖRT - Added ${channelsData.length} channels for CF: ${cf}, AG: ${adGroupTypeNameSelected}, Campaign: ${campaignNameForCsv}`);
        return; // Fortsätt inte med vanlig ad group skapning för channels
      } else {
        Logger.log(`generateCampaignAndAdGroupRows_: VARNING - Inga channels hittades för CF: ${cf}, Campaign: ${campaignNameForCsv}, AdGroup: ${adGroupTypeNameSelected}. Kontrollera att sheet 'Filtered_${cf.replace(' ', '_')}' finns och innehåller data.`);
      }
    }

    const adGroupName = adGroupTypeNameSelected;
    let campaignSpecificAdGroupType = fallbackAdGroupType;
    if (adsForThisCampaign && adsForThisCampaign.length > 0 && adsForThisCampaign[0] && adsForThisCampaign[0].adGroupType) {
      campaignSpecificAdGroupType = adsForThisCampaign[0].adGroupType;
    }
    const isVrcEfficientReach = (adsForThisCampaign && adsForThisCampaign.length > 0 && adsForThisCampaign[0] && (adsForThisCampaign[0].mediaProduct || "").trim().toLowerCase() === "video reach campaign efficient reach");
    let determinedAdGroupTypeForCSV = isVrcEfficientReach ? "Responsive video" : campaignSpecificAdGroupType;
    if (determinedAdGroupTypeForCSV.toLowerCase() === "non-skippable in-stream") { determinedAdGroupTypeForCSV = "Nonskippable instream"; }
    else if (determinedAdGroupTypeForCSV.toLowerCase() === "skippable in-stream") { determinedAdGroupTypeForCSV = "Skippable instream"; }

    if ((importMode !== 'adsOnly' && !targetingOnlyMode)) {
      const adGroupRow = createEmptyRow_(campaignTemplateRowData.length);
      setValue_(adGroupRow, col.campaign, campaignNameForCsv);
      setValue_(adGroupRow, col.adGroup, adGroupName);
      setValue_(adGroupRow, col.adGroupType, determinedAdGroupTypeForCSV);
      setValue_(adGroupRow, col.status, defaultAdGroupStatus);
      if (adsForThisCampaign && adsForThisCampaign.length > 0 && adsForThisCampaign[0] && adsForThisCampaign[0].clientRate !== undefined) {
        const originalRateStr = String(adsForThisCampaign[0].clientRate).replace(',', '.'); const originalRate = parseFloat(originalRateStr);
        const percentValStr = String(baseMetadata.clientRatePercent).replace(',', '.'); const percentVal = parseFloat(percentValStr);
        if (!isNaN(originalRate) && !isNaN(percentVal) && percentVal >= 0) {
          const calculatedRate = (originalRate * (percentVal / 100));
          let finalRateToSet = (calculatedRate % 1 !== 0) ? calculatedRate.toFixed(2) : calculatedRate.toString();
          const bidStrategyUpper = (baseMetadata.bidStrategyType || "").toUpperCase().trim();
          let targetColForRate = -1;
          if (bidStrategyUpper === "TARGET CPV") targetColForRate = col.targetCpv;
          else if (bidStrategyUpper === "TARGET CPM") targetColForRate = col.targetCpm;
          else if (bidStrategyUpper === "MANUAL CPV") targetColForRate = col.maxCpv;
          if (targetColForRate !== -1 && typeof targetColForRate === 'number') { setValue_(adGroupRow, targetColForRate, finalRateToSet); }
        }
      }
      if (col.optimizedTargeting !== -1) setValue_(adGroupRow, col.optimizedTargeting, "Disabled");
      if (col.audienceTargeting !== -1) setValue_(adGroupRow, col.audienceTargeting, "Audience segments");
      if (col.flexibleReach !== -1) setValue_(adGroupRow, col.flexibleReach, "Genders;Ages;Parental status;Household incomes");
      csvData.push(adGroupRow);

      if (col.age !== -1 && col.criterionType !== -1 && defaultCriterionStatus) {
        const selectedAges = selections.selectedAges || [];
        const agesToExclude = ALL_AGE_OPTIONS.filter(ageOption => !selectedAges.includes(ageOption));
        if (agesToExclude.length > 0 && agesToExclude.length < ALL_AGE_OPTIONS.length) {
          agesToExclude.forEach(ageToExclude => {
            const r = createEmptyRow_(campaignTemplateRowData.length);
            setValue_(r, col.campaign, campaignNameForCsv); setValue_(r, col.adGroup, adGroupName);
            setValue_(r, col.age, ageToExclude); setValue_(r, col.criterionType, "Negative Age");
            setValue_(r, col.status, defaultCriterionStatus); csvData.push(r);
          });
        }
      }
      if (col.gender !== -1 && col.criterionType !== -1 && defaultCriterionStatus) {
        const selectedGenders = selections.selectedGenders || [];
        const gendersToExclude = ALL_GENDER_OPTIONS.filter(genderOption => !selectedGenders.includes(genderOption));
        if (gendersToExclude.length > 0 && gendersToExclude.length < ALL_GENDER_OPTIONS.length) {
          gendersToExclude.forEach(genderToExclude => {
            const r = createEmptyRow_(campaignTemplateRowData.length);
            setValue_(r, col.campaign, campaignNameForCsv); setValue_(r, col.adGroup, adGroupName);
            setValue_(r, col.gender, genderToExclude); setValue_(r, col.criterionType, "Negative Gender");
            setValue_(r, col.status, defaultCriterionStatus); csvData.push(r);
          });
        }
      }
    }
    if (columnarTargetingData && Object.keys(columnarTargetingData).length > 0) {
      if (columnarTargetingData.audiences && columnarTargetingData.audiences.length > 0 && col.campaign !== -1 && col.adGroup !== -1 && col.audienceName !== -1 && col.criterionType !== -1) {
        columnarTargetingData.audiences.forEach(audience => {
          if (audience.name) {
            const audienceRow = createEmptyRow_(campaignTemplateRowData.length);
            setValue_(audienceRow, col.campaign, campaignNameForCsv);
            setValue_(audienceRow, col.adGroup, adGroupName);
            setValue_(audienceRow, col.audienceName, audience.name);
            setValue_(audienceRow, col.criterionType, defaultTargetingSetting || "Targeting");
            setValue_(audienceRow, col.status, defaultCriterionStatus);
            if (audience.id && col.id !== -1) {
              setValue_(audienceRow, col.id, audience.id);
            }
            csvData.push(audienceRow);
          }
        });
      }
      if (columnarTargetingData.topics && columnarTargetingData.topics.length > 0 && col.campaign !== -1 && col.adGroup !== -1 && col.topic !== -1 && col.criterionType !== -1) {
        columnarTargetingData.topics.forEach(topicData => {
          if (topicData.name) {
            const topicRow = createEmptyRow_(campaignTemplateRowData.length);
            setValue_(topicRow, col.campaign, campaignNameForCsv);
            setValue_(topicRow, col.adGroup, adGroupName);
            setValue_(topicRow, col.topic, topicData.name);
            setValue_(topicRow, col.criterionType, defaultTargetingSetting || "Targeting");
            setValue_(topicRow, col.status, defaultCriterionStatus);
            const topicIdCol = (col.topicId !== -1) ? col.topicId : col.id;
            if (topicData.id && topicIdCol !== -1) {
              setValue_(topicRow, topicIdCol, topicData.id);
            }
            csvData.push(topicRow);
          }
        });
      }
    }
    if (!targetingOnlyMode && adsForThisCampaign && adsForThisCampaign.length > 0) {
      adsForThisCampaign.forEach((adData, adIndex) => {
        if (!adData.videoId && !adData.finalUrl) { return; }
        const adRow = createEmptyRow_(campaignTemplateRowData.length);
        setValue_(adRow, col.campaign, campaignNameForCsv); setValue_(adRow, col.adGroup, adGroupName); setValue_(adRow, col.videoId, adData.videoId);
        const adMediaProductLower = (adData.mediaProduct || "").trim().toLowerCase();
        const adIsVrcEfficientReach = (adMediaProductLower === "video reach campaign efficient reach");
        let finalAdName;
        if (adIsVrcEfficientReach) { finalAdName = `Xs_[${adData.videoId || 'NoVideoID'}]_VRC`; }
        else { finalAdName = adData.adName || `${adGroupName} Ad ${String(adData.videoId || adData.finalUrl || adIndex + 1).slice(-5)}`; }
        setValue_(adRow, col.adName, finalAdName);
        let finalAdTypeForCSV;
        if (adIsVrcEfficientReach) { finalAdTypeForCSV = "Responsive video ad"; }
        else {
          if (adData.adType) { finalAdTypeForCSV = adData.adType; }
          else {
            const agTypeLower = determinedAdGroupTypeForCSV.toLowerCase();
            if (agTypeLower.includes("skippable")) finalAdTypeForCSV = "Skippable in-stream ad";
            else if (agTypeLower.includes("nonskippable")) finalAdTypeForCSV = "Non-skippable in-stream ad";
            else if (agTypeLower.includes("responsive video")) finalAdTypeForCSV = "Responsive video ad";
            else if (agTypeLower.includes("audio")) finalAdTypeForCSV = "Audio ad";
            else if (agTypeLower.includes("bumper")) finalAdTypeForCSV = "Bumper ad";
            else finalAdTypeForCSV = "";
          }
        }
        setValue_(adRow, col.adType, finalAdTypeForCSV);
        const adGroupTypeForDisplayUrlCheck = (determinedAdGroupTypeForCSV || "").toLowerCase();
        if (col.displayUrl !== -1) {
          if (adGroupTypeForDisplayUrlCheck.includes("responsive video") || adGroupTypeForDisplayUrlCheck.includes("efficient reach")) {
            setValue_(adRow, col.displayUrl, "");
          } else {
            setValue_(adRow, col.displayUrl, adData.displayUrl);
          }
        }
        setValue_(adRow, col.finalUrl, adData.finalUrl);
        setValue_(adRow, col.trackingTemplate, adData.trackingTemplate);
        setValue_(adRow, col.cta, adData.cta);
        setValue_(adRow, col.headline, adData.headline);
        if (col.longHeadline !== -1) { setValue_(adRow, col.longHeadline, adData.longHeadline); }
        setValue_(adRow, col.description1, adData.description1);
        setValue_(adRow, col.description2, adData.description2);
        const adStatusColumnIndex = (col.videoAdStatus !== -1) ? col.videoAdStatus : col.status;
        if (adStatusColumnIndex !== -1) { setValue_(adRow, adStatusColumnIndex, defaultAdStatus); }
        csvData.push(adRow);
      });
    }
  });
}

/**
 * ===============================================
 * HÄMTAR CHANNELS FRÅN FILTERED_* SHEETS
 * ===============================================
 */
function getChannelsForCF_(cf, cfVariant = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const channels = [];
  
  try {
    Logger.log(`=== getChannelsForCF_ DEBUG ===`);
    Logger.log(`Input CF: "${cf}"`);
    
    // Lista alla tillgängliga sheets för debugging
    const allSheets = ss.getSheets().map(s => s.getName()).filter(name => name.startsWith('Filtered_'));
    Logger.log(`Tillgängliga Filtered_ sheets: ${allSheets.join(', ')}`);
    
    // Bestäm vilket sheet som ska användas baserat på CF
    const cleanCF = cf.trim();
    let sheetName = null;
    
    // Förbättrad mapping för CF-varianter
    if (cleanCF.includes(' NM') || cleanCF.includes(' VM')) {
      // Direkt mapping för CF med varianter (CF1 NM -> Filtered_CF1_NM)
      sheetName = `Filtered_${cleanCF.replace(/ /g, '_')}`;
    } else {
      // Grundläggande CF utan variant (CF1 -> Filtered_CF1)
      sheetName = `Filtered_${cleanCF}`;
    }
    
    Logger.log(`Söker efter sheet: "${sheetName}" för CF: "${cf}"`);
    
    // Hämta sheet och kontrollera om det finns
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`FEL: Sheet '${sheetName}' finns inte! Tillgängliga: ${allSheets.join(', ')}`);
      
      // Försök alternativ mappning om första försöket misslyckas
      const alternativeNames = [
        `Filtered_${cleanCF.replace(' ', '_')}`,
        `Filtered_${cleanCF.replace(/\s+/g, '_')}`,
        `Filtered_${cleanCF}`
      ];
      
      for (const altName of alternativeNames) {
        const altSheet = ss.getSheetByName(altName);
        if (altSheet) {
          Logger.log(`Hittade alternativ sheet: "${altName}"`);
          return getChannelsFromSheet_(altSheet, cf);
        }
      }
      
      return [];
    }
    
    return getChannelsFromSheet_(sheet, cf);
    
  } catch (e) {
    Logger.log(`FEL i getChannelsForCF_ för CF ${cf}: ${e.message}`);
    return [];
  }
}

/**
 * Hjälpfunktion för att extrahera channels från ett sheet
 */
function getChannelsFromSheet_(sheet, cf) {
  const channels = [];
  
  try {
    // ULTRA-OPTIMERING: Hämta endast data vi behöver
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return []; // Snabb exit
    
    // Minimal logging för stora sheets
    if (lastRow > 20000) {
      console.log(`Processing ${sheet.getName()} with ${lastRow} rows...`);
    }
    
    // Hämta all data på en gång (snabbaste metoden)
    const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    
    // MEGA-SNABB processing: Pre-allokera array och använd for-loop istället för forEach
    channels.length = 0; // Säkerställ tom array
    
    for (let i = 0; i < data.length; i++) {
      const channelId = String(data[i][1] || "").trim();
      
      if (channelId) {
        channels.push({
          url: String(data[i][0] || "").trim(), // URL kan vara tom
          channelId: channelId
        });
      }
    }
    
    // Endast logging för stora datasets
    if (channels.length > 10000) {
      console.log(`Processed ${channels.length} channels from ${sheet.getName()}`);
    }
    
    return channels;
    
  } catch (e) {
    console.error(`Error in getChannelsFromSheet_ for ${sheet.getName()}: ${e.message}`);
    return [];
  }
}

// ===============================================
// HJÄLPFUNKTIONER FÖR CSV-GENERERING
// ===============================================

function setValue_(rowArray, colIndex, value) {
  // colIndex är 1-baserat från getIndex, så vi konverterar till 0-baserat för arrayen
  const zeroBasedIndex = colIndex - 1; 
  if (typeof colIndex === 'number' && colIndex >= 1 && zeroBasedIndex >= 0 && zeroBasedIndex < rowArray.length) {
    rowArray[zeroBasedIndex] = value; 
  } else {
    // Denna loggning är bra för felsökning om något går fel med kolumnindex
    // Logger.log(`setValue_ anropades med ogiltigt colIndex: ${colIndex} (0-baserat blir ${zeroBasedIndex}). Radlängd: ${rowArray.length}. Försökte sätta värde: "${value}" för kolumn som kanske inte finns i CSV-mallen eller är felmappad.`);
  }
}

function createEmptyRow_(length) {
  const validLength = (typeof length === 'number' && length > 0) ? Math.floor(length) : 0;
  return Array(validLength).fill("");
}

function formatDateForAds_(dateValue) {
  if (!dateValue) return ""; // Returnera tom sträng om inget datumvärde finns
  try {
    let dateObj = null;
    // Kontrollera om det redan är ett JavaScript Date-objekt
    if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
      dateObj = dateValue;
    } 
    // Hantera strängar (försök parsa olika format)
    else if (typeof dateValue === 'string' && dateValue.trim() !== '') {
      let cleanedDateStr = dateValue.trim().replace(/-/g, '/'); // Ersätt bindestreck med snedstreck för bättre kompatibilitet
      // Försök med standard new Date()
      let potentialDate = new Date(cleanedDateStr);
      // Om det misslyckas, och det ser ut som DD/MM/YYYY eller MM/DD/YYYY
      if (isNaN(potentialDate.getTime())) {
          const parts = cleanedDateStr.split('/');
          if (parts.length === 3) {
              const p1 = parseInt(parts[0],10);
              const p2 = parseInt(parts[1],10);
              const p3 = parseInt(parts[2],10);
              // Enkel heuristik: om p3 är ett år, försök MM/DD/YYYY och DD/MM/YYYY
              if (p3 > 1900 && p3 < 2100) {
                  if (p1 > 0 && p1 <=12 && p2 > 0 && p2 <= 31) { // Möjlig MM/DD/YYYY
                      potentialDate = new Date(p3, p1 - 1, p2);
                  }
                  if (isNaN(potentialDate.getTime()) && p2 > 0 && p2 <=12 && p1 > 0 && p1 <= 31) { // Möjlig DD/MM/YYYY
                       potentialDate = new Date(p3, p2 - 1, p1);
                  }
              }
          }
      }
      if (!isNaN(potentialDate.getTime())) {
        dateObj = potentialDate;
      }
    } 
    // Hantera Excel datum nummer (serienummer)
    else if (typeof dateValue === 'number' && dateValue > 0 && dateValue < 2958466) { // Omfång för Excel-datum
      const excelEpochDiff = 25569; // Dagar mellan 1970-01-01 och 1900-01-01 (Excel-epok)
      const dateInMilliseconds = (dateValue - excelEpochDiff) * 24 * 60 * 60 * 1000;
      let potentialDate = new Date(dateInMilliseconds);
       // Justera för UTC för att undvika tidszonsfel vid konvertering från Excel-tal
      if (!isNaN(potentialDate.getTime())) {
         dateObj = new Date(Date.UTC(potentialDate.getUTCFullYear(), potentialDate.getUTCMonth(), potentialDate.getUTCDate()));
      }
    }

    if (dateObj && !isNaN(dateObj.getTime())) {
      const year = dateObj.getUTCFullYear();
      const month = String(dateObj.getUTCMonth() + 1).padStart(2, '0'); // Månader är 0-indexerade
      const day = String(dateObj.getUTCDate()).padStart(2, '0');
      // Rimlighetskontroll för årtal
      if (year < 1970 || year > 2070) { // Justerat övre gräns för rimlighet
          Logger.log(`formatDateForAds_: Ogiltigt årtal (${year}) från värde: ${dateValue}, konverterat objekt: ${dateObj.toUTCString()}`);
          return ""; 
      }
      return `${year}-${month}-${day}`; // Format YYYY-MM-DD som Google Ads Editor föredrar
    } else {
      Logger.log(`formatDateForAds_: Kunde inte konvertera datumvärde: "${dateValue}" till ett giltigt datumobjekt.`);
      return ""; // Returnera tom sträng om datumet är ogiltigt
    }
  } catch (e) { 
    Logger.log(`formatDateForAds_ Allvarligt fel vid konvertering av datum "${dateValue}": ${e.message} \nStack: ${e.stack}`); 
    return ""; // Returnera tom sträng vid fel
  }
}

function escapeCsvCell_(cellValue) {
    if (cellValue === null || cellValue === undefined) { return ''; }
    const stringValue = String(cellValue);
    // Om cellen innehåller kommatecken, citationstecken eller radbrytningar, omslut med citationstecken
    // och dubbla eventuella befintliga citationstecken.
    if (/[",\n\r]/.test(stringValue)) {
        const escapedString = stringValue.replace(/"/g, '""'); // Dubbla alla citationstecken
        return `"${escapedString}"`; // Omslut hela strängen med citationstecken
    }
    return stringValue; // Returnera originalsträngen om ingen specialhantering behövs
}

/**
 * Hämtar kampanjstrukturen från "Compilation"-arket för att användas i popupen.
 * Inkluderar Location Name, Rotation, Multiformat ads, Client Rate etc.
 * **UPPDATERAD 2025-05-07: Säkerställer alla fält och robust loggning.**
 */
function getCampaignStructure() {
  const COMPILATION_SHEET_NAME = "Compilation";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(COMPILATION_SHEET_NAME);
  Logger.log(`getCampaignStructure: Starting processing for sheet: "${COMPILATION_SHEET_NAME}"`);

  if (!sheet) {
    Logger.log(`getCampaignStructure: Error - Sheet '${COMPILATION_SHEET_NAME}' not found.`);
    return { error: `Sheet '${COMPILATION_SHEET_NAME}' not found.` };
  }
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  if (data.length < 2) { 
    Logger.log(`getCampaignStructure: Error - Sheet '${COMPILATION_SHEET_NAME}' has less than 2 rows.`);
    return { error: `Sheet '${COMPILATION_SHEET_NAME}' needs headers and at least one data row.` };
  }

  try {
    const compilationHeaders = data[0].map(h => String(h || "").trim());
    const headersLower = compilationHeaders.map(h => h.toLowerCase());

    const findIndex = (name) => {
        const lowerName = name.toLowerCase().trim();
        const index = headersLower.indexOf(lowerName);
        if (index === -1) {
            Logger.log(`>>> VARNING getCampaignStructure: Kolumnrubriken '${name}' hittades INTE i arket '${COMPILATION_SHEET_NAME}'. Headers found: [${compilationHeaders.join(", ")}]`);
        }
        return index; 
    };

    Logger.log(`getCampaignStructure: Found Headers in "${COMPILATION_SHEET_NAME}": [${compilationHeaders.join(", ")}]`);

    const indices = {
        campaign       : findIndex("Campaign"),
        geoId          : findIndex("ID"),
        locationName   : findIndex("Location"),
        campaignType   : findIndex("Campaign Type"),
        bidStrategy    : findIndex("Bid Strategy Type"),
        startDate      : findIndex("Start Date SF"),
        endDate        : findIndex("End Date SF"),
        language       : findIndex("Language"),
        rotation       : findIndex("Rotation"),
        deviceTargeting: findIndex("Device Targeting"),
        multiformatAds : findIndex("Multiformat ads"),
        adType         : findIndex("Ad type"),
        adGroupType    : findIndex("Ad Group Type"),
        clientRate     : findIndex("Client Rate"),     
        adName         : findIndex("Ad Name"),
        videoId        : findIndex("Video ID"),
        tracking       : findIndex("Tracking template"),
        finalUrl       : findIndex("Final URL"),
        displayUrl     : findIndex("Display URL"),
        cta            : findIndex("CTA"),
        headline       : findIndex("Headline"),
        longHeadline   : findIndex("Long Headline"),
        desc1          : findIndex("Description 1"),
        desc2          : findIndex("Description 2"),
        mediaProduct   : findIndex("Media Products")
    };

    const criticalKeys = ["campaign", "rotation", "videoId", "finalUrl"]; 
    if(indices.clientRate === -1) Logger.log(">>> VARNING getCampaignStructure: Kolumnen 'Client Rate' hittades inte i Compilation, procentberäkning kommer att misslyckas eller använda fallback.");
    
    const missingCriticalKeys = criticalKeys.filter(key => indices[key] === -1);
    if (missingCriticalKeys.length > 0) {
      const errorMsg = `Kritiska kolumner saknas eller kunde inte hittas i '${COMPILATION_SHEET_NAME}': ${missingCriticalKeys.join(", ")}. Kan inte fortsätta.`;
      Logger.log(`getCampaignStructure: Error - ${errorMsg}`);
      return { error: errorMsg };
    }
    Logger.log(`getCampaignStructure: Kritiska kolumnindex hittade. Rotation index: ${indices.rotation}, Client Rate index: ${indices.clientRate}`);

    let campaigns = {};
    let currentCampaignName = null;
    let currentCampaignMetadata = null;

    const getDataSafely = (rowIndex, colIndex) => {
        const row = data[rowIndex]; 
        if (!row) return "";
        if (colIndex !== -1 && colIndex < row.length) { 
            return (row[colIndex] ?? "").toString().trim();
        }
        return "";
    };

    for (let i = 1; i < data.length; i++) {
      if (data[i].every(cell => String(cell || "").trim() === "")) continue; 
      const campaignNameInData = getDataSafely(i, indices.campaign);

      if (campaignNameInData) { 
        currentCampaignName = campaignNameInData;
        currentCampaignMetadata = {
            campaign       : currentCampaignName, 
            geoId          : getDataSafely(i, indices.geoId), 
            locationName   : getDataSafely(i, indices.locationName),
            startDate      : getDataSafely(i, indices.startDate), 
            endDate        : getDataSafely(i, indices.endDate),
            campaignType   : getDataSafely(i, indices.campaignType), 
            language       : getDataSafely(i, indices.language),
            bidStrategyType: getDataSafely(i, indices.bidStrategy), 
            rotation       : getDataSafely(i, indices.rotation),
            deviceTargeting: getDataSafely(i, indices.deviceTargeting)
        };
        if (!campaigns[currentCampaignName]) {
            campaigns[currentCampaignName] = { metadata: currentCampaignMetadata, rows: [] };
        } else {
             campaigns[currentCampaignName].metadata = currentCampaignMetadata; 
        }
      } else if (currentCampaignName && campaigns[currentCampaignName]) { 
        const videoIdValue = getDataSafely(i, indices.videoId);
        const finalUrlValue = getDataSafely(i, indices.finalUrl);
        const isLikelyAdRow = !campaignNameInData && (videoIdValue || finalUrlValue);

        if (isLikelyAdRow) {
          const adDataObj = {
              videoId         : videoIdValue,
              trackingTemplate: getDataSafely(i, indices.tracking),
              finalUrl        : finalUrlValue,
              displayUrl      : getDataSafely(i, indices.displayUrl),
              adName          : getDataSafely(i, indices.adName),
              adType          : getDataSafely(i, indices.adType),
              adGroupType     : getDataSafely(i, indices.adGroupType),
              clientRate      : getDataSafely(i, indices.clientRate), 
              cta             : getDataSafely(i, indices.cta),
              headline        : getDataSafely(i, indices.headline),
              longHeadline    : getDataSafely(i, indices.longHeadline),
              description1    : getDataSafely(i, indices.desc1),
              description2    : getDataSafely(i, indices.desc2),
              mediaProduct    : getDataSafely(i, indices.mediaProduct),
              multiformat     : getDataSafely(i, indices.multiformatAds),
              rotation        : getDataSafely(i, indices.rotation) 
          };
          Logger.log(`getCampaignStructure - Ad Data for ${currentCampaignName}, row ${i+1}: ${JSON.stringify(adDataObj)}`);
          campaigns[currentCampaignName].rows.push(adDataObj);
        }
      }
    } 

    const detectedCampaigns = Object.keys(campaigns);
    if (detectedCampaigns.length === 0) {
        Logger.log(`getCampaignStructure: Warning - No campaigns with ads were successfully processed from '${COMPILATION_SHEET_NAME}'.`);
        return { error: `No campaigns found or processed in '${COMPILATION_SHEET_NAME}'. Check sheet structure and data.` };
    }
     Logger.log(`getCampaignStructure: Finished. Detected ${detectedCampaigns.length} campaigns.`);
    return { campaigns: campaigns };

  } catch (e) {
    Logger.log(`FATAL Error in getCampaignStructure: ${e}\n${e.stack}`);
    return { error: `Error reading or processing '${COMPILATION_SHEET_NAME}': ${e.message}.` };
  }
}

// ===============================================
// TTT DATA UPDATERS (Video/Ad/Tracking & SF Data) - Uppdaterad 2025-05-01
// ===============================================
// Innehåller funktioner för att uppdatera TTT-arket med:
// - Video ID, Ad Name, Final URL, Display URL (baserat på Youtube Link, Tracking template ELLER Landing Page (If NO Click Tag))
// - Data från SF Data-arket (baserat på matchande Placement ID och datum)
//
// Baserat på ursprungliga Skript 2 och Skript 3.
// Kräver att findHeaderIndex_ finns definierad (från Core-skriptet).

// ===============================================
// DEL 1: VIDEO/AD/TRACKING UPDATER (från Skript 2 - Uppdaterad för Landing Page fallback)
// ===============================================

/**
 * Uppdaterar värdena i kolumnerna Video ID, Ad Name, Final URL och Display URL i arket "TTT".
 * Använder "Tracking template" i första hand för Final URL.
 * Om "Tracking template" är tom, används "Landing Page (If NO Click Tag)" som Final URL.
 * Modifierar även Tracking template (lägger till suffix) och rensar Final URL (tar bort dclid).
 * Bör köras EFTER convertTTT.
 */
function updateVideoAdTrackingValues() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetName = "TTT";
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    ui.alert(`Arket '${sheetName}' hittades inte. Kör först 'Convert TTT' och säkerställ att arket finns och är synligt.`);
    Logger.log(`updateVideoAdTrackingValues: Arket '${sheetName}' hittades inte.`);
    return;
  }
  if (sheet.isSheetHidden()) {
    ui.alert(`Arket '${sheetName}' är dolt. Gör det synligt och försök igen.`);
    Logger.log(`updateVideoAdTrackingValues: Arket '${sheetName}' är dolt.`);
    return;
  }

  const headerRow = 2;
  const firstDataRow = 3;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < firstDataRow || lastCol < 1) {
    ui.alert(`Inga data rader (från rad ${firstDataRow}) finns i '${sheetName}' att uppdatera.`);
    Logger.log(`updateVideoAdTrackingValues: Inga datarader att bearbeta.`);
    return;
  }

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Startar uppdatering av Video/Ad/Tracking...", "Status", 5);
    Logger.log("updateVideoAdTrackingValues: Startar uppdatering.");

    // --- Hämta rubriker och kolumnindex ---
    const headerRange = sheet.getRange(headerRow, 1, 1, lastCol);
    const headers = headerRange.getValues()[0];

    const youtubeLinkColIdx = findHeaderIndex_(headers, "Youtube Link");
    const videoIDColIdx = findHeaderIndex_(headers, "Video ID");
    const adNameColIdx = findHeaderIndex_(headers, "Ad Name");
    const trackingTempColIdx = findHeaderIndex_(headers, "Tracking template");
    const finalURLColIdx = findHeaderIndex_(headers, "Final URL");
    const displayURLColIdx = findHeaderIndex_(headers, "Display URL");
    const formatColIdx = findHeaderIndex_(headers, "Format");
    const landingPageNoClickTagColIdx = findHeaderIndex_(headers, "Landing Page (If NO Click Tag)");

    // Validera att nödvändiga kolumner finns
    const requiredCoreCols = ["Youtube Link", "Video ID", "Ad Name", "Tracking template", "Final URL", "Display URL"];
    const missingCoreCols = requiredCoreCols.filter(name => findHeaderIndex_(headers, name) === -1);

    let errorMsg = "";
    if (missingCoreCols.length > 0) {
      errorMsg += `Saknar nödvändiga kolumner i '${sheetName}': ${missingCoreCols.join(", ")}. `;
      Logger.log(`updateVideoAdTrackingValues: Saknar kolumner: ${missingCoreCols.join(", ")}`);
    }
    
    if (landingPageNoClickTagColIdx === -1) {
      Logger.log(`updateVideoAdTrackingValues: Varning - Kolumnen "Landing Page (If NO Click Tag)" saknas. Kan inte sätta Final URL när Tracking template är tom.`);
      ui.alert(`Varning: Kolumnen "Landing Page (If NO Click Tag)" saknas i '${sheetName}'.\n\nFinal URL kommer inte att sättas för rader där "Tracking template" är tom.`);
    }

    if (missingCoreCols.length > 0) { 
      ui.alert(errorMsg.trim() + "\n\nAvbryter uppdatering.");
      return;
    }

    if (formatColIdx === -1) {
      Logger.log(`updateVideoAdTrackingValues: Varning - Kolumnen "Format" saknas. Ad Name kommer inte innehålla format-suffix.`);
    }

    // --- Hämta data (från rad 3 till sista) ---
    const numRows = lastRow - firstDataRow + 1;
    const dataRange = sheet.getRange(firstDataRow, 1, numRows, lastCol);
    const data = dataRange.getValues();
    let primaryUpdatesMade = false;

    // --- Bearbeta data (Primära uppdateringar) ---
    for (let r = 0; r < data.length; r++) {
      let rowChanged = false;
      const currentRow = data[r];

      // Extrahera Video ID & skapa Ad Name
      if (youtubeLinkColIdx !== -1 && videoIDColIdx !== -1 && adNameColIdx !== -1) {
        const ytLink = String(currentRow[youtubeLinkColIdx - 1] || "").trim();
        if (ytLink) { // Kör bara om det finns en Youtube-länk
            const extractedVid = extractYoutubeID_(ytLink);
            if (extractedVid !== "" && extractedVid !== String(currentRow[videoIDColIdx - 1])) {
                currentRow[videoIDColIdx - 1] = extractedVid; 
                rowChanged = true;
            }
            // Skapa Ad Name även om Video ID inte ändrades, men baserat på det (potentiellt) nya/befintliga Video ID
            if (extractedVid) { // Använd extraherat Video ID om det finns
                let formatVal = (formatColIdx !== -1 && currentRow[formatColIdx - 1]) ? String(currentRow[formatColIdx - 1]).trim() : "";
                const newAdName = `Xs_[${extractedVid}]_${formatVal}`;
                if (newAdName !== String(currentRow[adNameColIdx - 1])) { 
                    currentRow[adNameColIdx - 1] = newAdName; 
                    rowChanged = true; // Markera ändring även om bara Ad Name justeras
                }
            }
        }
      }


      // Hämta Final URL och Display URL
      let newFinalUrl = null; 
      if (trackingTempColIdx !== -1 && finalURLColIdx !== -1) {
        const trackVal = String(currentRow[trackingTempColIdx - 1] || "").trim();
        if (trackVal) { 
          const resolvedUrl = getFinalRedirect_(trackVal, r + firstDataRow);
          if (resolvedUrl && resolvedUrl !== String(currentRow[finalURLColIdx - 1])) {
            newFinalUrl = resolvedUrl; 
            currentRow[finalURLColIdx - 1] = newFinalUrl;
            rowChanged = true;
          } else if (resolvedUrl) { // Om resolvedUrl är samma som befintlig, använd den ändå för Display URL-logik
            newFinalUrl = resolvedUrl;
          } else { // Om resolvedUrl är null/tom (t.ex. getFinalRedirect misslyckades grovt)
            newFinalUrl = String(currentRow[finalURLColIdx - 1] || ""); // Behåll befintlig
          }
        } else if (landingPageNoClickTagColIdx !== -1) { 
          const landingPageUrl = String(currentRow[landingPageNoClickTagColIdx - 1] || "").trim();
          if (landingPageUrl && landingPageUrl !== String(currentRow[finalURLColIdx - 1])) {
            newFinalUrl = landingPageUrl; 
            currentRow[finalURLColIdx - 1] = newFinalUrl;
            rowChanged = true;
            Logger.log(`Rad ${r + firstDataRow}: Använder Landing Page (If NO Click Tag) som Final URL: "${newFinalUrl}"`);
          } else if (landingPageUrl) { // Även om samma, använd den för Display URL-logik
            newFinalUrl = landingPageUrl;
          } else {
             newFinalUrl = String(currentRow[finalURLColIdx - 1] || "");
          }
        } else {
          newFinalUrl = String(currentRow[finalURLColIdx - 1] || "");
        }
      }

      // Uppdatera Display URL
      if (displayURLColIdx !== -1 && newFinalUrl !== null) { 
        const dispUrl = extractDomain_(newFinalUrl);
        if (dispUrl && dispUrl !== String(currentRow[displayURLColIdx - 1])) {
          currentRow[displayURLColIdx - 1] = dispUrl;
          rowChanged = true; 
        } else if (!dispUrl && String(currentRow[displayURLColIdx - 1] || "") !== "") {
          // Valfritt: Rensa Display URL om Final URL är ogiltig/tom och Display URL hade ett värde
          // currentRow[displayURLColIdx - 1] = "";
          // rowChanged = true; 
        }
      }
      
      if (rowChanged) {
        primaryUpdatesMade = true;
      }
    } 

    if (primaryUpdatesMade) {
      dataRange.setValues(data);
      Logger.log("updateVideoAdTrackingValues: Primära uppdateringar sparade.");
    } else {
      Logger.log("updateVideoAdTrackingValues: Inga primära uppdateringar behövde göras.");
    }

    // --- Sekundära justeringar (Tracking Template Suffix & dclid Cleanup) ---
    if (trackingTempColIdx !== -1) {
        const ttRange = sheet.getRange(firstDataRow, trackingTempColIdx, numRows, 1);
        const ttData = ttRange.getValues();
        const suffixToAdd = ";dc_transparent=1;?{lpurl}";
        let ttChanged = false;
        for (let r = 0; r < ttData.length; r++) {
            const current = String(ttData[r][0] || "").trim();
            if (current !== "" && !current.includes(suffixToAdd)) { 
                ttData[r][0] = current + suffixToAdd; 
                ttChanged = true; 
            }
        }
        if (ttChanged) { ttRange.setValues(ttData); Logger.log(`updateVideoAdTrackingValues: Suffix tillagt i Tracking template.`); }
    }

    if (finalURLColIdx !== -1) {
        const finalRange = sheet.getRange(firstDataRow, finalURLColIdx, numRows, 1);
        const finalData = finalRange.getValues();
        const dclidString = "&dclid=";
        let finalCleaned = false;
        for (let r = 0; r < finalData.length; r++) {
            const urlVal = String(finalData[r][0] || ""); 
            const idx = urlVal.indexOf(dclidString);
            if (idx !== -1) { 
                finalData[r][0] = urlVal.substring(0, idx); 
                finalCleaned = true; 
            }
        }
        if (finalCleaned) { finalRange.setValues(finalData); Logger.log(`updateVideoAdTrackingValues: Rensat dclid från Final URL.`); }
    }

    SpreadsheetApp.getActiveSpreadsheet().toast("Uppdatering av Video/Ad/Tracking klar!", "Status", 5);
    ui.alert(`Uppdatering av Video ID, Ad Name, Final URL och Display URL är klar i '${sheetName}'.`);
     Logger.log("updateVideoAdTrackingValues: Uppdatering slutförd.");

  } catch (e) {
    Logger.log(`Ett fel inträffade i updateVideoAdTrackingValues: ${e}\nStack: ${e.stack}`);
    ui.alert(`Ett fel inträffade vid uppdatering av Video/Ad/Tracking:\n\n${e.message}\n\nSe loggen (Visa > Loggar) för mer information.`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Fel vid uppdatering!", "Fel", 10);
  }
}

// --- Hjälpfunktioner för updateVideoAdTrackingValues ---
function extractYoutubeID_(url) {
  if (!url || typeof url !== 'string') return "";
  const patterns = [ /youtu\.be\/([a-zA-Z0-9_-]{11})/, /\?v=([a-zA-Z0-9_-]{11})/, /&v=([a-zA-Z0-9_-]{11})/, /embed\/([a-zA-Z0-9_-]{11})/, /v\/([a-zA-Z0-9_-]{11})/, /shorts\/([a-zA-Z0-9_-]{11})/ ];
  for (let i = 0; i < patterns.length; i++) { const match = url.match(patterns[i]); if (match && match[1]) { return match[1]; } } return ""; 
}

function extractDomain_(url) {
  if (!url || typeof url !== 'string') return ""; 
  try { 
    const parsedUrl = new URL(url); 
    return parsedUrl.hostname; 
  } catch (e) { 
    const match = url.match(/^(?:https?:\/\/)?(?:www\.)?([^\/]+)/i); 
    if (match && match[1]) { return match[1].split(':')[0]; } 
    Logger.log(`extractDomain_: Kunde inte extrahera domän från "${url}". Fel: ${e.message}`); return ""; 
  } 
}

function getFinalRedirect_(url, rowNumForLog = 0) {
  if (!url || typeof url !== 'string' || !url.toLowerCase().startsWith('http')) { 
    if (url) Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Ogiltig start-URL "${url}". Måste börja med http/https.`);
    return url; 
  }
  let currentUrl = url;
  if (url.includes("{") || url.includes("}")) { 
    currentUrl = url.replace(/{/g, "%7B").replace(/}/g, "%7D");
    Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Ersatte {{}} i URL: ${currentUrl}`);
  }
  const options = { followRedirects: false, muteHttpExceptions: true }; 
  const maxRedirects = 10; 
  let redirectsFollowed = 0; 
  const visitedUrls = new Set(); 
  visitedUrls.add(currentUrl); 
  while (redirectsFollowed < maxRedirects) { 
    let response; 
    try { 
      response = UrlFetchApp.fetch(currentUrl, options); 
    } catch (e) { 
      Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Nätverksfel vid hämtning av "${currentUrl}". Fel: ${e.message}. Returnerar senaste giltiga URL.`); 
      return currentUrl; 
    }
    const responseCode = response.getResponseCode(); 
    const headers = response.getHeaders(); 
    const locationHeader = headers['Location'] || headers['location']; 
    if (responseCode >= 300 && responseCode < 400 && locationHeader) { 
      let nextUrl = locationHeader.trim(); 
      if (!nextUrl.toLowerCase().startsWith('http')) { 
        try { 
          const baseUrl = new URL(currentUrl); 
          nextUrl = new URL(nextUrl, baseUrl).toString(); 
        } catch(e) { 
          Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Kunde inte konstruera absolut URL från relativ "${locationHeader}" och bas "${currentUrl}". Fel: ${e.message}. Returnerar nuvarande URL.`); 
          return currentUrl; 
        }
      }
      if (visitedUrls.has(nextUrl)) { 
        Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Omdirigeringsloop upptäckt vid "${nextUrl}". Returnerar denna URL.`); 
        return nextUrl; 
      }
      currentUrl = nextUrl; 
      visitedUrls.add(currentUrl); 
      redirectsFollowed++; 
      Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Följde omdirigering #${redirectsFollowed} till: ${currentUrl}`);
    } else { 
      if (responseCode >= 400) {
         Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Fick status ${responseCode} för "${currentUrl}". Returnerar denna URL som slutgiltig (trots fel).`);
      } else if (responseCode >= 300 && !locationHeader) {
         Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Fick status ${responseCode} men ingen Location header för "${currentUrl}". Returnerar denna URL.`);
      }
      return currentUrl; 
    }
  } 
  Logger.log(`getFinalRedirect_ (Rad ${rowNumForLog}): Max antal omdirigeringar (${maxRedirects}) nåddes. Returnerar den sista URL:en: ${currentUrl}`); 
  return currentUrl;
}

function testUpdateVideoAdTrackingValues() { 
  Logger.log("Kör testUpdateVideoAdTrackingValues..."); 
  updateVideoAdTrackingValues(); 
  Logger.log("testUpdateVideoAdTrackingValues klar."); 
}

// ===============================================
// TARGETING FUNKTIONER (från TargetingHelper v1.6)
// ===============================================

/**
 * Visar popup-fönstret för targeting.
 * Kallas från menyn 'Ad Tools'.
 */
function showTargetingPopup() {
  const html = HtmlService.createHtmlOutputFromFile('targetingPopup')
    .setWidth(850)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Lägg till targeting på befintliga kampanjer/adgroups');
}

/**
 * Wrapperfunktion för targeting menyvalet
 */
function runPopulateTargetingIDs() {
    Logger.log("===> runPopulateTargetingIDs: Startar...");
    try {
        SpreadsheetApp.getActiveSpreadsheet().toast("Startar ID-lookup...", "Pågår", 5);
        populateTargetingIDs();
        Logger.log("===> runPopulateTargetingIDs: populateTargetingIDs() klar.");
        const ui = SpreadsheetApp.getUi();
        ui.alert('ID-lookup klar! Kolumnerna för "ID" i arket "' + TARGETING_SHEET + '" har uppdaterats.\n\nKontrollera loggen (Visa > Körningar) för detaljer och eventuella varningar.');
        SpreadsheetApp.getActiveSpreadsheet().toast("ID-lookup klar!", "Status", 5);
    } catch (e) {
        Logger.log("===> runPopulateTargetingIDs: FEL FÅNGAT!");
        Logger.log("Fel vid populateTargetingIDs: " + e + "\nStack: " + e.stack);
        SpreadsheetApp.getUi().alert("Ett allvarligt fel uppstod: " + e.message + "\nSe loggen för detaljer.");
        SpreadsheetApp.getActiveSpreadsheet().toast("Fel vid ID-lookup!", "Fel", 10);
    }
    Logger.log("===> runPopulateTargetingIDs: Avslutar.");
}

/**
 * Huvudfunktion: Läser namn från TARGETING_SHEET, slår upp ID:n i master-arken
 * och skriver tillbaka ID:na till TARGETING_SHEET.
 */
function populateTargetingIDs() {
    Logger.log("===> populateTargetingIDs: Funktion startad.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(TARGETING_SHEET);

    if (!targetSheet) {
      Logger.log(`===> populateTargetingIDs: FEL - Arket "${TARGETING_SHEET}" hittades inte.`);
      throw new Error(`Hittade inte arket som heter "${TARGETING_SHEET}".`);
    }
    Logger.log(`===> populateTargetingIDs: Hittade arket "${TARGETING_SHEET}".`);

    const lookupMaps = {};
    Logger.log("===> populateTargetingIDs: Försöker bygga lookup maps...");
    for (const key in MASTER_SHEETS_CONFIG) {
        const config = MASTER_SHEETS_CONFIG[key];
        Logger.log(` ---> Försöker bygga map för: ${config.name}`);
        try {
             lookupMaps[key] = buildLookupMap_(config.name, config.masterKeyCol, config.masterValueCol);
             if (lookupMaps[key] === null) {
                 Logger.log(`    - ERROR/SKIP: Lookup map build misslyckades för ${config.name}. Kontrollera arknamn/kolumner. Denna typ kommer att hoppas över.`);
             } else if (lookupMaps[key].size === 0) {
                 Logger.log(`    - WARNING: Kartan för ${config.name} är tom (inga giltiga nyckel/värde-par hittades i master-arket).`);
             } else {
                 Logger.log(`    - SUCCESS: Karta byggd för ${config.name} (${lookupMaps[key].size} entries).`);
             }
        } catch (mapError) {
             Logger.log(`    - CRITICAL ERROR building map for ${config.name}: ${mapError}\n${mapError.stack}`);
             lookupMaps[key] = null; 
        }
    }
    Logger.log("===> populateTargetingIDs: Klar med försök att bygga lookup maps.");

    const targetRange = targetSheet.getDataRange();
    const targetData = targetRange.getValues();
    Logger.log(`===> populateTargetingIDs: Läste ${targetData.length} rader från ${TARGETING_SHEET}.`);
    if (targetData.length < 2) {
        Logger.log(`Arket "${TARGETING_SHEET}" innehåller ingen data under rubrikraden. Avslutar.`);
        SpreadsheetApp.getActiveSpreadsheet().toast("Ingen data att bearbeta i Targeting-arket.", "Info", 5);
        return;
    }
    const targetHeaders = targetData[0].map(h => String(h || "").trim());
    Logger.log(`===> populateTargetingIDs: Targeting Headers: ${targetHeaders.join(" | ")}`);

    const targetIndices = {};
    for (const key in MASTER_SHEETS_CONFIG) {
        const config = MASTER_SHEETS_CONFIG[key];
        const nameColIdx = targetHeaders.indexOf(config.nameCol);
        const idColIdx = targetHeaders.indexOf(config.idCol);
        if (nameColIdx === -1 || idColIdx === -1) {
            Logger.log(` ---> VARNING: Kolumn "${config.nameCol}" eller "${config.idCol}" hittades inte i "${TARGETING_SHEET}". Kan inte processa typen: ${key}.`);
            targetIndices[key] = null;
        } else {
            targetIndices[key] = { name: nameColIdx, id: idColIdx };
        }
    }

    let idsWrittenTotal = 0;
    let rowsProcessed = targetData.length - 1;

    for (const key in MASTER_SHEETS_CONFIG) {
        const config = MASTER_SHEETS_CONFIG[key];
        const indices = targetIndices[key];
        const lookupMap = lookupMaps[key]; 

        if (!indices || !lookupMap) {
             Logger.log(` ---> Hoppar över ${key} pga saknade kolumner i ${TARGETING_SHEET} eller misslyckad lookup map för ${config.name}.`);
             continue; 
        }

        const nameColIdx_0 = indices.name;
        const idColIdx_0 = indices.id;
        Logger.log(` ---> Bearbetar typ: ${key}. Namnkolumn: "${config.nameCol}" [${nameColIdx_0}], ID-kolumn: "${config.idCol}" [${idColIdx_0}]`);

        const idsToWrite = [];
        let idsFoundCount = 0;
        let namesNotFound = []; 

        for (let i = 1; i < targetData.length; i++) { 
            const row = targetData[i];
            let foundId = ""; 

            if (nameColIdx_0 < row.length) {
                const nameToLookupRaw = row[nameColIdx_0] ? String(row[nameColIdx_0]).trim() : null;

                if (nameToLookupRaw && nameToLookupRaw !== "") {
                    let nameForLookup = nameToLookupRaw;

                    // Steg 1: Ersätt alla ">" (med eventuella mellanslag runt) med "/"
                    nameForLookup = nameForLookup.replace(/\s*>\s*/g, '/');
                    
                    // Steg 2: Standardisera alla "/" till ett enda tecken utan mellanslag runt
                    nameForLookup = nameForLookup.replace(/\s*\/\s*/g, '/');
                    
                    foundId = lookupMap.get(nameForLookup) || "";

                    if (foundId) {
                       idsFoundCount++;
                    } else {
                       if (namesNotFound.indexOf(nameToLookupRaw) === -1) { 
                           namesNotFound.push(nameToLookupRaw);
                       }
                    }
                }
            }
            idsToWrite.push([foundId]);
        } 

        if (namesNotFound.length > 0) {
            Logger.log(`      - VARNING (${key}): Följande namn från kolumn "${config.nameCol}" hittades inte i master-arket "${config.name}" eller matchade inte formatet: ${namesNotFound.slice(0, 20).join(", ")}` + (namesNotFound.length > 20 ? ` ... och ${namesNotFound.length - 20} till.` : ''));
        }

        if (idsToWrite.length > 0) {
             try {
                 targetSheet.getRange(2, idColIdx_0 + 1, idsToWrite.length, 1).setValues(idsToWrite);
                 Logger.log(` ---> Skrev ${idsFoundCount} av ${idsToWrite.length} möjliga ID:n till kolumn "${config.idCol}" för typ ${key}.`);
                 idsWrittenTotal += idsFoundCount;
             } catch (writeError) {
                  Logger.log(` ---> FEL vid skrivning till kolumn "${config.idCol}" för typ ${key}: ${writeError}`);
                  SpreadsheetApp.getUi().alert(`Ett fel uppstod vid skrivning till kolumn "${config.idCol}". Skriptet fortsätter. Se logg.`);
             }
        }
    } 

    Logger.log(`===> populateTargetingIDs: Klar med kolumn-processning. Totalt ${idsWrittenTotal} ID:n skrevs över ${rowsProcessed} rader.`);
} 

/**
 * Hjälpfunktion: Bygger en lookup map (Key -> Value) från ett master-ark.
 */
function buildLookupMap_(sheetName, keyColumnName, valueColumnName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        Logger.log(`buildLookupMap_: Master-arket "${sheetName}" hittades inte.`);
        return null; 
    }
    const dataRange = sheet.getDataRange();
    const data = dataRange.getValues();
    if (data.length < 2) {
        Logger.log(`buildLookupMap_: Master-arket "${sheetName}" har ingen data under rubriken.`);
        return new Map(); 
    }
    const headers = data[0].map(h => String(h || "").trim());
    const keyColIdx = headers.indexOf(keyColumnName); 
    const valColIdx = headers.indexOf(valueColumnName); 

    if (keyColIdx === -1 || valColIdx === -1) {
        Logger.log(`buildLookupMap_: Nödvändig kolumn "${keyColumnName}" (index ${keyColIdx}) eller "${valueColumnName}" (index ${valColIdx}) hittades inte i master-arket "${sheetName}".`);
        return null; 
    }

    const lookupMap = new Map();
    let duplicates = 0;
    let validEntries = 0;
    for (let i = 1; i < data.length; i++) {
        if (data[i].length <= Math.max(keyColIdx, valColIdx)) { continue; }
        const keyRaw = data[i][keyColIdx] ? String(data[i][keyColIdx]).trim() : null; 
        const value = data[i][valColIdx] ? String(data[i][valColIdx]).trim() : null; 

        if (keyRaw && keyRaw !== "" && value && value !== "") {
            let keyNormalized = keyRaw;

            // Steg 1: Ta bort eventuellt inledande slash (/)
            if (keyNormalized.startsWith('/')) {
                keyNormalized = keyNormalized.substring(1);
            }

            // Steg 2: Ersätt alla ">" (med eventuella mellanslag runt) med "/"
            keyNormalized = keyNormalized.replace(/\s*>\s*/g, '/');
            
            // Steg 3: Standardisera alla "/" till ett enda tecken utan mellanslag runt
            keyNormalized = keyNormalized.replace(/\s*\/\s*/g, '/');

            if (keyNormalized) { 
                if (lookupMap.has(keyNormalized)) {
                    duplicates++;
                }
                lookupMap.set(keyNormalized, value); 
                validEntries++;
            }
        }
    }
    if (duplicates > 0) {
        Logger.log(`buildLookupMap_: Byggde map för "${sheetName}" med ${lookupMap.size} unika nycklar (${validEntries} totalt lästa). ${duplicates} dubblettnyckelvärden upptäcktes (senaste värdet för varje nyckel behölls).`);
    } else {
        Logger.log(`buildLookupMap_: Byggde map för "${sheetName}" med ${lookupMap.size} unika nycklar (${validEntries} totalt lästa). Inga dubbletter.`);
    }
    return lookupMap;
}

/**
 * Funktion för att generera endast targeting CSV för befintliga kampanjer/adgroups
 */
function generateTargetingOnlyCSV(selections) {
  Logger.log("Genererar targeting-only CSV med inställningar:", JSON.stringify(selections));
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const compilationSheet = ss.getSheetByName('Compilation');
    const targetingSheet = ss.getSheetByName(TARGETING_SHEET);
    
    if (!compilationSheet || !targetingSheet) {
      Logger.log("Saknar nödvändiga ark: Compilation eller " + TARGETING_SHEET);
      return "Error: Saknar nödvändiga ark: Compilation eller " + TARGETING_SHEET;
    }
    
    // Hämtar befintliga targeting-data
    const targetingData = getTargetingData(selections);
    
    if (!targetingData || targetingData.length === 0) {
      return "Error: Ingen targeting-data hittades för de valda kriterierna.";
    }
    
    // Genererar CSV för targeting
    const csvRows = [];
    // Lägger till CSV-header
    csvRows.push('"Action Type","Customer ID","Campaign","Ad group","Criterion Type","Audience","Keyword","Placement","Category","Topic","Exclude"');
    
    // För varje targeting-rad, generera motsvarande CSV-rad
    targetingData.forEach(targeting => {
      const campaignName = selections.campaignName;
      const adGroup = targeting.adGroup;
      const criterionType = targeting.criterionType;
      const targetingValue = targeting.targetingValue;
      const targetingId = targeting.targetingId || '';
      
      // Bestäm vilket fält som ska användas baserat på criterion type
      let audienceField = '';
      let keywordField = '';
      let placementField = '';
      let categoryField = '';
      let topicField = '';
      
      switch(criterionType.toLowerCase()) {
        case 'audience':
          audienceField = targetingId || targetingValue;
          break;
        case 'keyword':
          keywordField = targetingValue;
          break;
        case 'placement':
          placementField = targetingValue;
          break;
        case 'category':
          categoryField = targetingId || targetingValue;
          break;
        case 'topic':
          topicField = targetingId || targetingValue;
          break;
      }
      
      // Skapa CSV-rad
      const csvRow = [
        'Add', // Action Type
        '', // Customer ID (tom)
        campaignName,
        adGroup,
        criterionType,
        audienceField,
        keywordField,
        placementField,
        categoryField,
        topicField,
        'false' // Exclude (alltid false för targeting)
      ];
      
      csvRows.push('"' + csvRow.join('","') + '"');
    });
    
    return csvRows.join('\n');
  } catch (e) {
    Logger.log("Fel vid generering av targeting-CSV: " + e.message);
    return "Error: " + e.message;
  }
}

/**
 * Hämtar targeting-data baserat på valda kriterier
 */
function getTargetingData(selections) {
  const targetingType = selections.targetingType;
  const campaignName = selections.campaignName;
  const adGroups = selections.adGroups || [];
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetingSheet = ss.getSheetByName(TARGETING_SHEET);
  
  if (!targetingSheet) {
    Logger.log("Targeting ark saknas");
    return [];
  }
  
  const data = targetingSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log("Ingen data i targeting ark");
    return [];
  }
  
  const headers = data[0].map(h => String(h).trim());
  
  // Hitta kolumner för targeting-typer baserat på MASTER_SHEETS_CONFIG
  const targetingColumns = {};
  const targetingIdColumns = {};
  
  for (const key in MASTER_SHEETS_CONFIG) {
    const config = MASTER_SHEETS_CONFIG[key];
    const nameColIdx = headers.indexOf(config.nameCol);
    const idColIdx = headers.indexOf(config.idCol);
    
    if (nameColIdx !== -1) {
      targetingColumns[key.toLowerCase()] = nameColIdx;
    }
    
    if (idColIdx !== -1) {
      targetingIdColumns[key.toLowerCase()] = idColIdx;
    }
  }
  
  // Hitta ad group-kolumnen
  const adGroupColIdx = headers.indexOf('Ad Group');
  if (adGroupColIdx === -1) {
    Logger.log("Kolumnen 'Ad Group' saknas i targeting ark");
    return [];
  }
  
  // Filtrera targeting-data
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const adGroup = row[adGroupColIdx] ? String(row[adGroupColIdx]).trim() : '';
    
    // Kontrollera om denna ad group är vald
    if (!adGroup || (adGroups.length > 0 && !adGroups.includes(adGroup))) {
      continue;
    }
    
    // Leta efter targeting-data i denna rad
    for (const type in targetingColumns) {
      const colIdx = targetingColumns[type];
      const idColIdx = targetingIdColumns[type] || -1;
      
      if (colIdx !== -1 && row[colIdx] && String(row[colIdx]).trim() !== '') {
        const targetingValue = String(row[colIdx]).trim();
        let targetingId = '';
        
        if (idColIdx !== -1 && row[idColIdx]) {
          targetingId = String(row[idColIdx]).trim();
        }
        
        // Om specifik targeting-typ valdes, filtrera på den
        if (targetingType && targetingType.toLowerCase() !== 'alla' && 
            targetingType.toLowerCase() !== type.toLowerCase()) {
          continue;
        }
        
        // Bestäm criterionType
        let criterionType;
        switch (type.toLowerCase()) {
          case 'topics':
            criterionType = 'Topic';
            break;
          case 'affinities':
          case 'inmarket':
          case 'lifeevents':
            criterionType = 'Audience';
            break;
          default:
            criterionType = 'Audience'; // Standard
        }
        
        result.push({
          adGroup: adGroup,
          criterionType: criterionType,
          targetingValue: targetingValue,
          targetingId: targetingId,
          targetingType: type
        });
      }
    }
  }
  
  return result;
}

/**
 * Funktion för att hämta kampanjnamn från Compilation-arket
 */
function getCampaignNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const compilationSheet = ss.getSheetByName('Compilation');
    
    if (!compilationSheet) {
      return { error: "Compilation-arket hittades inte" };
    }
    
    const data = compilationSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: "Ingen data i Compilation-arket" };
    }
    
    const headers = data[0].map(h => String(h).trim());
    const campaignColIdx = headers.indexOf('Campaign');
    
    if (campaignColIdx === -1) {
      return { error: "Kolumnen 'Campaign' hittades inte i Compilation-arket" };
    }
    
    const campaigns = {};
    for (let i = 1; i < data.length; i++) {
      const campaignName = data[i][campaignColIdx] ? String(data[i][campaignColIdx]).trim() : '';
      if (campaignName) {
        campaigns[campaignName] = true;
      }
    }
    
    return { campaigns: Object.keys(campaigns) };
  } catch (e) {
    return { error: "Fel vid hämtning av kampanjer: " + e.message };
  }
}

/**
 * Funktion för att hämta ad groups för en specifik kampanj
 */
function getAdGroupsForCampaign(campaignName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const compilationSheet = ss.getSheetByName('Compilation');
    
    if (!compilationSheet) {
      return { error: "Compilation-arket hittades inte" };
    }
    
    const data = compilationSheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { error: "Ingen data i Compilation-arket" };
    }
    
    const headers = data[0].map(h => String(h).trim());
    const campaignColIdx = headers.indexOf('Campaign');
    const adGroupColIdx = headers.indexOf('Ad Group');
    
    if (campaignColIdx === -1 || adGroupColIdx === -1) {
      return { error: "Kolumnen 'Campaign' eller 'Ad Group' hittades inte" };
    }
    
    const adGroups = {};
    for (let i = 1; i < data.length; i++) {
      const currentCampaign = data[i][campaignColIdx] ? String(data[i][campaignColIdx]).trim() : '';
      if (currentCampaign === campaignName) {
        const adGroup = data[i][adGroupColIdx] ? String(data[i][adGroupColIdx]).trim() : '';
        if (adGroup) {
          adGroups[adGroup] = true;
        }
      }
    }
    
    return { adGroups: Object.keys(adGroups) };
  } catch (e) {
    return { error: "Fel vid hämtning av ad groups: " + e.message };
  }
}

/**
 * Uppdaterar TTT-arket med data från SF Data-arket baserat på matchning av
 * Placement ID (rensat), Start Date TTT och End Date TTT.
 * Förväntar sig specifika kolumnnamn i båda arken.
 */
function updateTTTFromSFData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tttSheetName = "TTT";
  const sfSheetName = "SF Data";

  const sheetTTT = ss.getSheetByName(tttSheetName);
  const sheetSF = ss.getSheetByName(sfSheetName);

  // --- Validera att arken finns ---
  if (!sheetTTT) {
    ui.alert(`Arket "${tttSheetName}" saknas. Kör först "Convert TTT".`);
    Logger.log(`updateTTTFromSFData: Arket "${tttSheetName}" saknas.`);
    return;
  }
  if (sheetTTT.isSheetHidden()) {
    ui.alert(`Arket '${tttSheetName}' är dolt. Gör det synligt och försök igen.`);
    Logger.log(`updateTTTFromSFData: Arket '${tttSheetName}' är dolt.`);
    return;
  }
  if (!sheetSF) {
    ui.alert(`Arket "${sfSheetName}" saknas eller är felstavat.`);
    Logger.log(`updateTTTFromSFData: Arket "${sfSheetName}" saknas.`);
    return;
  }

  const tttHeaderRow = 2; // Rubriker i TTT på rad 2
  const sfHeaderRow = 1;  // Rubriker i SF Data på rad 1
  const tttFirstDataRow = tttHeaderRow + 1;
  const sfFirstDataRow = sfHeaderRow + 1;

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Startar uppdatering från SF Data...", "Status", 5);
    Logger.log("updateTTTFromSFData: Startar uppdatering.");

    // --- Hämta och validera rubriker och index i TTT ---
    const tttLastCol = sheetTTT.getLastColumn();
    if (sheetTTT.getMaxRows() < tttHeaderRow || tttLastCol < 1) throw new Error(`Arket "${tttSheetName}" saknar rubriker på rad ${tttHeaderRow}.`);
    const tttHeaders = sheetTTT.getRange(tttHeaderRow, 1, 1, tttLastCol).getValues()[0];
    const tttIndex = {}; // Objekt för att mappa TTT-rubrik => 0-baserat index
    tttHeaders.forEach((h, i) => { tttIndex[String(h || "").trim()] = i; });

    const requiredTTTCols = [
      "Placement ID", "Start Date TTT", "End Date TTT", "Opportunity Name",
      "Start Date SF", "End Date SF", "Bid Strategy Type",
      "Location", "Language", "Age", "Gender",
      "Device Targeting", "Media Products", "Client Rate (converted)",
      "Total Client Cost (converted)", "Total Campaign Placement Days"
    ];
    const missingTTT = requiredTTTCols.filter(col => tttIndex[col] === undefined);
    if (missingTTT.length > 0) {
      throw new Error(`Saknar följande nödvändiga kolumner i "${tttSheetName}" (rad ${tttHeaderRow}):\n${missingTTT.join(", ")}`);
    }

    // --- Hämta och validera rubriker och index i SF Data ---
    const sfLastCol = sheetSF.getLastColumn();
    if (sheetSF.getMaxRows() < sfHeaderRow || sfLastCol < 1) throw new Error(`Arket "${sfSheetName}" saknar rubriker på rad ${sfHeaderRow}.`);
    const sfHeaders = sheetSF.getRange(sfHeaderRow, 1, 1, sfLastCol).getValues()[0];
    const sfIndex = {}; // Objekt för att mappa SF-rubrik => 0-baserat index
    sfHeaders.forEach((h, i) => { sfIndex[String(h || "").trim()] = i; });

    const requiredSFCols = [
      "PL Number", "Placement Start Date", "Placement End Date", "Opportunity Name",
      "Cost Method", "Geo Delivered", "Language(s)", "Demographic",
      "Device Targeting", "Media Products", "Client Rate (converted)",
      "Total Client Cost (converted)", "Total Campaign Placement Days"
    ];
    const missingSF = requiredSFCols.filter(col => sfIndex[col] === undefined);
    if (missingSF.length > 0) {
      throw new Error(`Saknar följande nödvändiga kolumner i "${sfSheetName}" (rad ${sfHeaderRow}):\n${missingSF.join(", ")}`);
    }

    // --- Hämta data från TTT ---
    const lastRowTTT = sheetTTT.getLastRow();
    if (lastRowTTT < tttFirstDataRow) {
      Logger.log(`updateTTTFromSFData: Inga datarader funna i "${tttSheetName}".`);
      ui.alert(`Inga datarader att uppdatera hittades i "${tttSheetName}" (från rad ${tttFirstDataRow}).`);
      return;
    }
    const numRowsTTT = lastRowTTT - tttFirstDataRow + 1; // Korrekt antal rader
    const tttDataRange = sheetTTT.getRange(tttFirstDataRow, 1, numRowsTTT, tttLastCol);
    const tttData = tttDataRange.getValues();

    // --- Hämta data från SF Data ---
    const lastRowSF = sheetSF.getLastRow();
    if (lastRowSF < sfFirstDataRow) {
      Logger.log(`updateTTTFromSFData: Inga datarader funna i "${sfSheetName}".`);
      ui.alert(`Inga datarader att hämta data från hittades i "${sfSheetName}" (från rad ${sfFirstDataRow}).`);
      return;
    }
    const numRowsSF = lastRowSF - sfFirstDataRow + 1; // Korrekt antal rader
    const sfDataRange = sheetSF.getRange(sfFirstDataRow, 1, numRowsSF, sfLastCol);
    const sfData = sfDataRange.getValues();

    // --- Bygg lookup-dictionary från SF Data ---
    const sfDict = {};
    sfData.forEach((rowSF, index) => {
      const plNumClean = cleanPlacementId_(rowSF[sfIndex["PL Number"]]);
      const startDateFormatted = formatDateForDictKey_(rowSF[sfIndex["Placement Start Date"]]);
      const endDateFormatted = formatDateForDictKey_(rowSF[sfIndex["Placement End Date"]]);

      if (!plNumClean || !startDateFormatted || !endDateFormatted) {
        Logger.log(`updateTTTFromSFData: Hoppar över rad ${sfFirstDataRow + index} i SF Data p.g.a. ogiltig nyckeldata (PL: ${plNumClean}, Start: ${startDateFormatted}, End: ${endDateFormatted}).`);
        return; 
      }
      const key = `${plNumClean}|${startDateFormatted}|${endDateFormatted}`;
      if (sfDict.hasOwnProperty(key)) {
        Logger.log(`updateTTTFromSFData: Varning - Dubblettnyckel "${key}" hittad i SF Data (rad ${sfFirstDataRow + index}). Tidigare värde skrivs över.`);
      }
      sfDict[key] = rowSF;
    });
    Logger.log(`updateTTTFromSFData: Skapade SF Data dictionary med ${Object.keys(sfDict).length} unika nycklar.`);

    // --- Loopa igenom TTT-data och uppdatera baserat på matchning ---
    let rowsUpdated = 0;
    let rowsMatched = 0;
    tttData.forEach((rowTTT, index) => {
      const placementIdClean = cleanPlacementId_(rowTTT[tttIndex["Placement ID"]]);
      const startDateTTTFormatted = formatDateForDictKey_(rowTTT[tttIndex["Start Date TTT"]]);
      const endDateTTTFormatted = formatDateForDictKey_(rowTTT[tttIndex["End Date TTT"]]);

      if (!placementIdClean || !startDateTTTFormatted || !endDateTTTFormatted) {
        Logger.log(`updateTTTFromSFData: Hoppar över rad ${tttFirstDataRow + index} i TTT p.g.a. ogiltig nyckeldata för matchning (PL: ${placementIdClean}, Start: ${startDateTTTFormatted}, End: ${endDateTTTFormatted}).`);
        return; 
      }
      const tttKey = `${placementIdClean}|${startDateTTTFormatted}|${endDateTTTFormatted}`;

      if (sfDict.hasOwnProperty(tttKey)) {
        rowsMatched++;
        const rowSF = sfDict[tttKey]; 
        let tttRowChanged = false;
        const updateIfNeeded = (tttColName, sfColName, transformFn = null) => {
          const tttColIdx = tttIndex[tttColName];
          const sfColIdx = sfIndex[sfColName];
          let sfValue = rowSF[sfColIdx];
          if (transformFn) {
            sfValue = transformFn(sfValue, rowSF); 
          }
          const currentValueStr = String(rowTTT[tttColIdx] ?? "");
          const newValueStr = String(sfValue ?? "");
          if (currentValueStr !== newValueStr) {
            rowTTT[tttColIdx] = sfValue; 
            tttRowChanged = true;
          }
        };

        updateIfNeeded("Opportunity Name", "Opportunity Name");
        updateIfNeeded("Start Date SF", "Placement Start Date"); 
        updateIfNeeded("End Date SF", "Placement End Date"); 
        updateIfNeeded("Bid Strategy Type", "Cost Method", (costMethod) => {
          const cm = String(costMethod || "").trim().toUpperCase();
          if (cm === "CPV") return "Manual CPV";
          if (cm === "CPM") return "Target CPM";
          return costMethod; 
        });
        updateIfNeeded("Location", "Geo Delivered");
        updateIfNeeded("Language", "Language(s)");
        const demoValue = rowSF[sfIndex["Demographic"]];
        const parsedDemo = parseDemographicField_(demoValue);
        if (String(rowTTT[tttIndex["Age"]] ?? "") !== String(parsedDemo.age ?? "")) { rowTTT[tttIndex["Age"]] = parsedDemo.age; tttRowChanged = true; }
        if (String(rowTTT[tttIndex["Gender"]] ?? "") !== String(parsedDemo.gender ?? "")) { rowTTT[tttIndex["Gender"]] = parsedDemo.gender; tttRowChanged = true; }
        updateIfNeeded("Device Targeting", "Device Targeting");
        updateIfNeeded("Media Products", "Media Products");
        updateIfNeeded("Client Rate (converted)", "Client Rate (converted)");
        updateIfNeeded("Total Client Cost (converted)", "Total Client Cost (converted)");
        updateIfNeeded("Total Campaign Placement Days", "Total Campaign Placement Days");
        if (tttRowChanged) {
          rowsUpdated++;
        }
      } 
    }); 

    if (rowsUpdated > 0) {
      tttDataRange.setValues(tttData);
      Logger.log(`updateTTTFromSFData: Uppdatering från SF Data klar. ${rowsMatched} rader matchade, ${rowsUpdated} rader hade ändringar.`);
      ui.alert(`Uppdatering från SF Data slutförd.\n${rowsMatched} rader i TTT matchade en rad i SF Data.\n${rowsUpdated} av dessa rader hade värden som uppdaterades.`);
    } else {
      Logger.log(`updateTTTFromSFData: Uppdatering från SF Data klar. ${rowsMatched} rader matchade, men inga värden behövde ändras.`);
      ui.alert(`Uppdatering från SF Data slutförd.\n${rowsMatched} rader i TTT matchade en rad i SF Data.\nInga värden behövde dock uppdateras.`);
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("Uppdatering från SF Data klar!", "Status", 5);

  } catch (e) {
    Logger.log(`Ett fel inträffade i updateTTTFromSFData: ${e}\nStack: ${e.stack}`);
    ui.alert(`Ett fel inträffade vid uppdatering från SF Data:\n\n${e.message}\n\nSe loggen (Visa > Loggar) för mer information.`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Fel vid uppdatering!", "Fel", 10);
  }
}

// --- Hjälpfunktioner för updateTTTFromSFData ---
function cleanPlacementId_(val) {
  return String(val || "").trim().replace(/^PL\s*/i, "");
}
function formatDateForDictKey_(dateVal) {
  return parseToMMDDYYYY_(dateVal);
}
function parseDemographicField_(demoVal) {
  const result = { age: "", gender: "" };
  if (!demoVal || typeof demoVal !== 'string') return result;
  let trimmed = demoVal.trim();
  if (trimmed.startsWith("[") && trimmed.endsWith("]")) {
    trimmed = trimmed.substring(1, trimmed.length - 1);
  }
  if (trimmed === "") {
    result.age = ""; 
    result.gender = "";
    return result;
  }
  const tokens = trimmed.split(/[,;]+/);
  const agesFound = [];
  let genderStr = "";
  const allowedAgeGroups = ["18-24", "25-34", "35-44", "45-54", "55-64", "65-up", "65+", "unknown"];
  tokens.forEach(token => {
    const t = token.trim();
    const tLower = t.toLowerCase();
    if (allowedAgeGroups.includes(tLower)) {
      if (tLower === "65+") agesFound.push("65-up");
      else if (tLower === "unknown") agesFound.push("Unknown"); 
      else agesFound.push(t); 
    } else {
      const upperToken = t.toUpperCase();
      switch (upperToken) {
        case "A": genderStr = "Male, Female, Unknown"; break;
        case "M": if (!genderStr.includes("Unknown")) { genderStr = genderStr ? `${genderStr}, Male` : "Male"; } break;
        case "F": if (!genderStr.includes("Unknown")) { genderStr = genderStr ? `${genderStr}, Female` : "Female"; } break;
        case "MF": case "M-F": case "M/F": if (!genderStr.includes("Unknown")) { genderStr = "Male, Female"; } break;
        case "U": case "UNKNOWN": if (!genderStr.includes("Unknown") && !genderStr.includes("Male, Female, Unknown")) { genderStr = genderStr ? `${genderStr}, Unknown` : "Unknown"; } break;
      }
    }
  });
  if (genderStr.includes("Male") && genderStr.includes("Female") && genderStr.includes("Unknown")) {
    genderStr = "Male, Female, Unknown";
  } else if (genderStr.includes("Male") && genderStr.includes("Female")) {
    if (!genderStr.includes("Unknown")) genderStr = "Male, Female";
  }
  result.gender = genderStr.split(',').map(g => g.trim()).filter(g => g).join(', ');
  result.age = agesFound.join(", ");
  if (!result.age && !result.gender && demoVal.trim() && demoVal.trim() !== "[]") {
    result.age = demoVal; 
    Logger.log(`parseDemographicField_: Kunde inte tolka Demographic "${demoVal}". Behåller originalvärdet i Age-fältet.`);
  }
  return result;
}

function testUpdateTTTFromSFData() {
  Logger.log("Kör testUpdateTTTFromSFData...");
  updateTTTFromSFData();
  Logger.log("testUpdateTTTFromSFData klar.");
}

/**
 * Uppdaterar Geo ID, Språk och Platsnamn i TTT-arket.
 * Specifik logik för Belgien baserat på "Location" och kampanjnamn.
 * **UPPDATERAD: Förfinad Belgien-logik enligt nya specifikationer.**
 */
function updateTTTGeoLanguage() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const tttSheetName = "TTT";
  const geoSheetName = "Geo Codes";
  const tttSheet = ss.getSheetByName(tttSheetName);
  const geoSheet = ss.getSheetByName(geoSheetName);

  if (!tttSheet || !geoSheet) { 
      ui.alert(`Kunde inte hitta ett eller båda arken: "${tttSheetName}", "${geoSheetName}"`); 
      return; 
  }
  if (tttSheet.isSheetHidden()) { ui.alert(`Arket '${tttSheetName}' är dolt.`); return; }

  const tttHeaderRow = 2; 
  const geoHeaderRow = 1;
  const tttFirstDataRow = tttHeaderRow + 1; 
  const geoFirstDataRow = geoHeaderRow + 1;

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Startar uppdatering av Geo/Language...", "Status", 5);
    Logger.log("updateTTTGeoLanguage: Startar uppdatering med ny Belgien-logik.");

    const tttLastCol = tttSheet.getLastColumn();
    if (tttSheet.getMaxRows() < tttHeaderRow || tttLastCol < 1) throw new Error(`Arket "${tttSheetName}" saknar rubriker.`);
    const tttHeaders = tttSheet.getRange(tttHeaderRow, 1, 1, tttLastCol).getValues()[0];
    const locationColIdxTTT = findHeaderIndex_(tttHeaders, "Location");
    const idColIdxTTT = findHeaderIndex_(tttHeaders, "ID");
    const languageColIdxTTT = findHeaderIndex_(tttHeaders, "Language");
    const campaignColIdxTTT = findHeaderIndex_(tttHeaders, "Campaign"); 

    if (locationColIdxTTT === -1 || idColIdxTTT === -1 || languageColIdxTTT === -1 || campaignColIdxTTT === -1) {
      throw new Error(`Saknar nödvändiga kolumner i "${tttSheetName}": Location, ID, Language, Campaign.`);
    }

    const geoLastCol = geoSheet.getLastColumn();
    if (geoSheet.getMaxRows() < geoHeaderRow || geoLastCol < 1) throw new Error(`Arket "${geoSheetName}" saknar rubriker.`);
    const geoHeaders = geoSheet.getRange(geoHeaderRow, 1, 1, geoLastCol).getValues()[0];
    const canonicalNameColIdxGeo = findHeaderIndex_(geoHeaders, "Canonical Name");
    const criteriaIDColIdxGeo = findHeaderIndex_(geoHeaders, "Criteria ID");
    const countryCodeColIdxGeo = findHeaderIndex_(geoHeaders, "Country Code");
    if (canonicalNameColIdxGeo === -1 || criteriaIDColIdxGeo === -1 || countryCodeColIdxGeo === -1) {
      throw new Error(`Saknar kolumner i "${geoSheetName}": Canonical Name, Criteria ID, Country Code.`);
    }
    const geoLastRow = geoSheet.getLastRow();
    if (geoLastRow < geoFirstDataRow) { ui.alert(`Inga data i "${geoSheetName}"...`); return; }
    const geoData = geoSheet.getRange(geoFirstDataRow, 1, (geoLastRow - geoFirstDataRow + 1), geoLastCol).getValues();
    
    const geoMapping = {}; 
    let belgiumCountryId = null; // Specifikt för lands-ID för Belgien

    geoData.forEach((row) => {
      const canonicalName = String(row[canonicalNameColIdxGeo - 1] || "").trim();
      const criteriaID = String(row[criteriaIDColIdxGeo - 1] || "").trim(); 
      const countryCode = String(row[countryCodeColIdxGeo - 1] || "").trim();
      const nameLower = canonicalName.toLowerCase();

      if (canonicalName && criteriaID && countryCode) { 
        geoMapping[nameLower] = { criteriaID: criteriaID, countryCode: countryCode, name: canonicalName };
        if (nameLower === "belgium") { // Spara lands-ID för Belgien separat
            belgiumCountryId = criteriaID;
        }
      }
    });
    
    if (!belgiumCountryId) {
         Logger.log(`VARNING: Kunde inte hitta lands-ID för "Belgium" i 'Geo Codes'. Fallback för Belgien utan FR/NL i kampanjnamn kanske inte fungerar som väntat.`);
         // ui.alert kan läggas till här om det är kritiskt
    }
    Logger.log(`geoMapping skapad. Lands-ID för Belgien: ${belgiumCountryId || 'Inte hittat'}`);
    
    const tttLastRow = tttSheet.getLastRow();
    if (tttLastRow < tttFirstDataRow) { ui.alert(`Inga datarader i "${tttSheetName}"...`); return; }
    const tttDataRange = tttSheet.getRange(tttFirstDataRow, 1, (tttLastRow - tttFirstDataRow + 1), tttLastCol);
    const tttData = tttDataRange.getValues();
    
    let rowsUpdatedCount = 0; 
    const notFoundSet = new Set();

    for (let r = 0; r < tttData.length; r++) {
      const tttRow = tttData[r];
      let originalLocation = String(tttRow[locationColIdxTTT - 1] || "").trim();
      const campaignNameValue = String(tttRow[campaignColIdxTTT - 1] || "").trim();
      let originalId = String(tttRow[idColIdxTTT - 1] || "").trim();
      let originalLang = String(tttRow[languageColIdxTTT - 1] || "").trim();

      let newLocation = originalLocation;
      let newId = originalId;
      let newLang = originalLang;
      
      let rowChangedThisIteration = false;

      const locationLower = originalLocation.toLowerCase();
      const campaignLower = campaignNameValue.toLowerCase();

      if (locationLower === "belgium") {
        Logger.log(`Rad ${r + tttFirstDataRow}: Hittade "Belgium" i Location. Kampanj: "${campaignNameValue}"`);
        let specificRegionApplied = false;

        if (campaignLower.includes("nl")) {
          newLocation = "Flanders";
          newId = "9069523";
          newLang = "nl;en";
          Logger.log(` -> NL i kampanjnamn. Satt: Location="${newLocation}", ID="${newId}", Lang="${newLang}"`);
          specificRegionApplied = true;
        } else if (campaignLower.includes("fr")) {
          newLocation = "Wallonia;Brussels";
          newId = "9069524;20052";
          newLang = "fr;en";
          Logger.log(` -> FR i kampanjnamn. Satt: Location="${newLocation}", ID="${newId}", Lang="${newLang}"`);
          specificRegionApplied = true;
        }

        if (!specificRegionApplied) {
          // Om "Location" var "Belgium" men ingen "nl" eller "fr" i kampanjnamnet.
          // Använd lands-ID för Belgien och standardiserade språk. Behåll "Location" som "Belgium" (eller dess kanoniska form).
          if (belgiumCountryId) {
            const belgiumCountryCanonicalName = geoMapping["belgium"] ? geoMapping["belgium"].name : originalLocation; // Hämta kanoniskt namn om möjligt
            
            newLocation = belgiumCountryCanonicalName; // Kan vara "Belgium" eller t.ex. "Belgien" om det är så i Geo Codes
            newId = belgiumCountryId;
            newLang = "fr;nl;en";
            Logger.log(` -> Ingen NL/FR i kampanj. Fallback till land Belgien. Satt: Location="${newLocation}", ID="${newId}", Lang="${newLang}"`);
          } else {
            Logger.log(` -> Ingen NL/FR i kampanj, och lands-ID för Belgien kunde inte hittas. Behåller ursprungliga TTT-värden för "${originalLocation}".`);
          }
        }
      } else if (originalLocation) { // Standardhantering för andra icke-tomma platser
        const match = geoMapping[locationLower];
        if (match) {
          newLocation = match.name; // Uppdatera till kanoniskt namn
          newId = match.criteriaID;
          newLang = match.countryCode ? match.countryCode.toLowerCase() + ";en" : originalLang; // Behåll original om ingen landskod finns
          Logger.log(`Rad ${r + tttFirstDataRow}: Standardlookup för "${originalLocation}". Satt: Location="${newLocation}", ID="${newId}", Lang="${newLang}"`);
        } else {
          if (!notFoundSet.has(originalLocation)) { 
            Logger.log(`Location "${originalLocation}" på rad ${r + tttFirstDataRow} hittades ej i Geo Codes.`); 
            notFoundSet.add(originalLocation); 
          }
        }
      }
      // Om originalLocation är tomt, och det inte var ett "Belgium"-fall, görs inga ändringar.

      // Jämför och uppdatera endast om något faktiskt har ändrats
      if (originalLocation !== newLocation) {
        tttRow[locationColIdxTTT - 1] = newLocation;
        rowChangedThisIteration = true;
      }
      if (originalId !== newId) {
        tttRow[idColIdxTTT - 1] = newId;
        rowChangedThisIteration = true;
      }
      if (originalLang !== newLang) {
        tttRow[languageColIdxTTT - 1] = newLang;
        rowChangedThisIteration = true;
      }

      if (rowChangedThisIteration) {
        rowsUpdatedCount++;
      }
    } // Slut på TTT-loop

    if (rowsUpdatedCount > 0) { 
      tttDataRange.setValues(tttData); 
      Logger.log(`updateTTTGeoLanguage: ${rowsUpdatedCount} rader uppdaterades.`);
    } else {
       Logger.log(`updateTTTGeoLanguage: Inga rader behövde uppdateras.`);
    }

    let message = `Geo/Language uppdatering klar. ${rowsUpdatedCount} rader hade ändringar.`;
    if (notFoundSet.size > 0) { message += `\n${notFoundSet.size} unika platser (ej Belgien) hittades ej i Geo Codes.`; }
    ui.alert(message);
    SpreadsheetApp.getActiveSpreadsheet().toast("Uppdatering av Geo/Language klar!", "Status", 5);
    
  } catch (e) {
    Logger.log(`Fel i updateTTTGeoLanguage: ${e}\nStack: ${e.stack}`);
    ui.alert(`Fel vid uppdatering Geo/Language:\n\n${e.message}`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Fel vid Geo/Lang uppdatering!", "Fel", 10);
  }
}

// ===============================================
// KAMPANJMAPPNINGS-UPPDATERAREN (från Skript 5)
// ===============================================

/**
 * Uppdaterar Campaign Mapping från Campaign Type.
 * Skapar och uppdaterar kolumnen "Campaign Mapping".
 */
function updateCampaignMapping() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const sheetName = "TTT";
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) { ui.alert(`Arket "${sheetName}" hittades inte...`); return; }
  if (sheet.isSheetHidden()) { ui.alert(`Arket '${sheetName}' är dolt...`); return; }

  const headerRow = 2; const firstDataRow = 3;
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Startar uppdatering av Campaign Mapping...", "Status", 5);
    Logger.log("updateCampaignMapping: Startar.");
    const lastCol = sheet.getLastColumn();
    if (sheet.getMaxRows() < headerRow || lastCol < 1) throw new Error(`Arket "${sheetName}" saknar rubriker...`);
    const headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    const colIndices = {
        mediaProducts: findHeaderIndex_(headers, "Media Products"),
        campaignType: findHeaderIndex_(headers, "Campaign Type"),
        adGroupType: findHeaderIndex_(headers, "Ad Group Type"),
        adType: findHeaderIndex_(headers, "Ad type"),
        campaign: findHeaderIndex_(headers, "Campaign"),
        oppName: findHeaderIndex_(headers, "Opportunity Name"),
        multiformat: findHeaderIndex_(headers, "Multiformat ads"),
        adName: findHeaderIndex_(headers, "Ad Name"),
        videoId: findHeaderIndex_(headers, "Video ID")
    };
    const requiredMappings = { 
        "Media Products": colIndices.mediaProducts, "Campaign Type": colIndices.campaignType,
        "Ad Group Type": colIndices.adGroupType, "Ad type": colIndices.adType, "Campaign": colIndices.campaign,
        "Opportunity Name": colIndices.oppName, "Multiformat ads": colIndices.multiformat,
        "Ad Name": colIndices.adName, "Video ID": colIndices.videoId
    };
    const missingCols = Object.entries(requiredMappings).filter(([_, index]) => index === -1).map(([key]) => key);
    if (missingCols.length > 0) { throw new Error(`Saknar kolumner i "${sheetName}": ${missingCols.join(", ")}`); }

    const lastRow = sheet.getLastRow();
    if (lastRow < firstDataRow) { /* ... hantera inga datarader ... */ return; }
    const numRows = lastRow - firstDataRow + 1;
    const range = sheet.getRange(firstDataRow, 1, numRows, lastCol);
    const values = range.getValues();

    let opCodeValue = ""; const opRegex = /OP\d{6}/i;
    for (let i = 0; i < values.length; i++) { /* ... hitta opCodeValue ... */ if (values[i][colIndices.oppName - 1]) { const match = String(values[i][colIndices.oppName - 1]).trim().match(opRegex); if (match) { opCodeValue = match[0].toUpperCase(); break; } } }
    // if (!opCodeValue) { Logger.log(`updateCampaignMapping: Ingen OP-kod hittades.`); } // Kan tas bort

    let rowsUpdatedCount = 0;
    for (let i = 0; i < values.length; i++) {
        const currentRow = values[i];
        const mediaProd = String(currentRow[colIndices.mediaProducts - 1] || "").trim();
        let rowChanged = false;
        let calculatedCampaignType = "", calculatedAdGroupType = "", calculatedAdType = "", calculatedMultiformat = "";
        let calculatedAdName = String(currentRow[colIndices.adName - 1] || ""); // Behåll befintligt som default
        const currentVideoId = String(currentRow[colIndices.videoId - 1] || "").trim();

        if (mediaProd) {
            const mediaProdLower = mediaProd.toLowerCase();
            // Logger.log(`Row ${i + firstDataRow} (updateCampaignMapping): Media Product: "${mediaProd}"`); // Kan tas bort

            // Sätt Campaign Type, Ad Group Type, Ad Type, Multiformat baserat på mediaProdLower
            // ... (samma logik som tidigare för dessa) ...
             if (mediaProdLower.includes("non-skip")) { calculatedCampaignType = "Video - Non-skippable"; } else if (mediaProdLower.startsWith("audio")) { calculatedCampaignType = "Video - Audio"; } else { calculatedCampaignType = "Video"; }
             switch (mediaProdLower) { case "video reach campaign efficient reach": calculatedAdGroupType = "Responsive video"; break; case "trueview in-stream": calculatedAdGroupType = "Skippable in-stream"; break; case "video reach campaign efficient reach 2.0": calculatedAdGroupType = "Responsive video"; break; case "bumpers": case "bumsers": calculatedAdGroupType = "Responsive video"; break; case "trueview for reach": calculatedAdGroupType = "Responsive video"; break; case "shorts": calculatedAdGroupType = "Responsive video"; break; case "audio ads": calculatedAdGroupType = "Audio"; break; default: if (mediaProdLower.startsWith("non-skip")) { calculatedAdGroupType = "Nonskippable instream"; } else { calculatedAdGroupType = "Standard"; } break; }
             switch (calculatedAdGroupType) { case "Skippable in-stream": calculatedAdType = "Skippable in-stream ad"; break; case "Nonskippable instream": calculatedAdType = "Non-skippable in-stream ad"; break; case "Audio": calculatedAdType = "Audio ad"; break; case "Responsive video": case "efficient reach": calculatedAdType = "Responsive video ad"; break; default: calculatedAdType = ""; break; }
             switch (mediaProdLower) { case "video views campaign": case "video reach campaign efficient reach 2.0": case "video reach campaign efficient reach": calculatedMultiformat = "In-stream ads;In-feed ads;Shorts ads"; break; case "shorts": calculatedMultiformat = "Shorts ads"; break; case "in-feed video ad": calculatedMultiformat = "In-feed ads"; break; case "trueview for reach": case "bumpers": case "bumsers": calculatedMultiformat = "In-stream ads"; break; default: calculatedMultiformat = ""; break; }
       

            // +++ START PÅ NY OCH UTÖKAD LOGIK FÖR Ad Name +++
            if (currentVideoId) { 
                let expectedAdName = null; 

                if (mediaProdLower.includes("bumper")) {
                    if (calculatedAdName.startsWith("X")) {
                        calculatedAdName = calculatedAdName.replace(/^X/, '6');
                        Logger.log(`Row ${i + firstDataRow}: Uppdaterar Ad Name till "${calculatedAdName}" (Bumper: X → 6)`);
                        rowChanged = true; 
                    }
                } else if (mediaProdLower.includes("non-skip") && mediaProdLower.includes("15")) { 
                    if (calculatedAdName.startsWith("X")) {
                        calculatedAdName = calculatedAdName.replace(/^X/, '15');
                        Logger.log(`Row ${i + firstDataRow}: Uppdaterar Ad Name till "${calculatedAdName}" (15s Non-skip: X → 15)`);
                        rowChanged = true;
                    }
                } else if (mediaProdLower.includes("non-skip") && mediaProdLower.includes("20")) { 
                    if (calculatedAdName.startsWith("X")) {
                        calculatedAdName = calculatedAdName.replace(/^X/, '20');
                        Logger.log(`Row ${i + firstDataRow}: Uppdaterar Ad Name till "${calculatedAdName}" (20s Non-skip: X → 20)`);
                        rowChanged = true;
                    }
                } else if (mediaProd.toLowerCase() === "video reach campaign efficient reach") { 
                    expectedAdName = `Xs_[${currentVideoId}]_VRC`;
                    if (calculatedAdName !== expectedAdName) {
                        calculatedAdName = expectedAdName;
                        Logger.log(`Row ${i + firstDataRow}: Uppdaterar Ad Name till "${calculatedAdName}" baserat på Media Product "${mediaProd}"`);
                        rowChanged = true;
                    }
                } else if (mediaProd.toLowerCase() === "video reach campaign efficient reach 2.0") {
                    expectedAdName = `Xs_[${currentVideoId}]_VRC_2.0`;
                    if (calculatedAdName !== expectedAdName) {
                        calculatedAdName = expectedAdName;
                        Logger.log(`Row ${i + firstDataRow}: Uppdaterar Ad Name till "${calculatedAdName}" baserat på Media Product "${mediaProd}"`);
                        rowChanged = true;
                    }
                } else if (mediaProd.toLowerCase() === "trueview for reach") {
                    expectedAdName = `Xs_[${currentVideoId}]_TVFR`;
                    if (calculatedAdName !== expectedAdName) {
                        calculatedAdName = expectedAdName;
                        Logger.log(`Row ${i + firstDataRow}: Uppdaterar Ad Name till "${calculatedAdName}" baserat på Media Product "${mediaProd}"`);
                        rowChanged = true;
                    }
                }
            } else { 
                 // Logger.log(`Row ${i + firstDataRow}: Kan inte sätta specifikt Ad Name format eftersom Video ID saknas.`);
            }
            // +++ SLUT PÅ NY OCH UTÖKAD LOGIK FÖR Ad Name +++

        } // Avslutande måsvinge för "if (mediaProd)"

        // Uppdatera värden i arrayen om de ändrats
        if (String(currentRow[colIndices.campaignType - 1] ?? "") !== calculatedCampaignType) { currentRow[colIndices.campaignType - 1] = calculatedCampaignType; rowChanged = true; }
        if (String(currentRow[colIndices.adGroupType - 1] ?? "") !== calculatedAdGroupType) { currentRow[colIndices.adGroupType - 1] = calculatedAdGroupType; rowChanged = true; }
        if (String(currentRow[colIndices.adType - 1] ?? "") !== calculatedAdType) { currentRow[colIndices.adType - 1] = calculatedAdType; rowChanged = true; }
        if (String(currentRow[colIndices.multiformat - 1] ?? "") !== calculatedMultiformat) { currentRow[colIndices.multiformat - 1] = calculatedMultiformat; rowChanged = true; }
        // Uppdatera Ad Name OM det ändrades av logiken ovan
        if (String(currentRow[colIndices.adName - 1] ?? "") !== calculatedAdName) { 
            currentRow[colIndices.adName - 1] = calculatedAdName; 
            rowChanged = true; 
        }

        // Lägg till OP-kod på kampanjnamn (som tidigare)
        if (opCodeValue) {
            const currentCampaign = String(currentRow[colIndices.campaign - 1] || "").trim();
            if (currentCampaign && !currentCampaign.endsWith(` ${opCodeValue}`)) { 
                currentRow[colIndices.campaign - 1] = `${currentCampaign} ${opCodeValue}`; 
                rowChanged = true; 
            }
        }
        if (rowChanged) { rowsUpdatedCount++; }
    } // End for loop

    if (rowsUpdatedCount > 0) { 
        range.setValues(values); 
        ui.alert(`Campaign Mapping klar. ${rowsUpdatedCount} rader påverkades.`);
    } else { 
        ui.alert("Campaign Mapping klar. Inga rader behövde uppdateras.");
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("Uppdatering av Mapping klar!", "Status", 5);
    Logger.log(`updateCampaignMapping: Klar. ${rowsUpdatedCount} rader uppdaterades.`);
  } catch (e) { 
      Logger.log(`Fel i updateCampaignMapping: ${e}\nStack: ${e.stack || 'Ingen stack tillgänglig'}`);
      ui.alert(`Fel vid Campaign Mapping:\n\n${e.message}`);
      SpreadsheetApp.getActiveSpreadsheet().toast("Fel vid Mapping!", "Fel", 10);
  }
}

function findHeaderIndex_(headers, headerName) {
    const lowerCaseHeaderName = headerName.toLowerCase();
    for (let i = 0; i < headers.length; i++) {
        if (headers[i].toLowerCase() === lowerCaseHeaderName) {
            return i + 1; // Returnera 1-baserat index
        }
    }
    return -1; // Returnera -1 om huvudet inte hittades
}

/**
 * ===============================================
 * COMPILATION SHEET BUILDER (Uppdaterad för Rotation & Client Rate)
 * ===============================================
 * Läser data från TTT-arket, grupperar per kampanj och bygger
 * "Compilation"-arket. Inkluderar nu "Rotation", "Multiformat ads" och "Client Rate".
 */
function buildCampaignSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  const tttSheetName = "TTT"; const compilationSheetName = "Compilation";
  const tttSheet = ss.getSheetByName(tttSheetName);

  if (!tttSheet) { ui.alert(`Arket "${tttSheetName}" saknas.`); Logger.log(`buildCampaignSummary: Arket "${tttSheetName}" saknas.`); return; }
  if (tttSheet.isSheetHidden()) { ui.alert(`Arket '${tttSheetName}' är dolt.`); Logger.log(`buildCampaignSummary: Arket '${tttSheetName}' är dolt.`); return; }

  const headerRowIndex = 1; // Antag att rubriker i TTT är på rad 2 (0-indexerat blir 1)
  const dataStartIndex = headerRowIndex + 1; // Data börjar på rad 3

  try {
    SpreadsheetApp.getActiveSpreadsheet().toast("Startar bygge av Compilation-ark...", "Status", 5);
    Logger.log("buildCampaignSummary: Startar.");

    const allData = tttSheet.getDataRange().getValues();
    if (allData.length < dataStartIndex) { 
        ui.alert(`Inte tillräckligt med data i "${tttSheetName}" (behöver minst ${dataStartIndex} rader).`);
        Logger.log(`buildCampaignSummary: Inte tillräckligt med data. Rader: ${allData.length}, Krävs: ${dataStartIndex}`);
        return;
    }

    const headersTTT = allData[headerRowIndex].map(h => String(h || "").trim()); 
    Logger.log(`buildCampaignSummary: Hittade rubriker i TTT: ${headersTTT.join(" | ")}`);

    const requiredColsSetup = [
      { compilationHeader: "Campaign", tttHeader: "Campaign" },
      { compilationHeader: "ID", tttHeader: "ID" },
      { compilationHeader: "Campaign Type", tttHeader: "Campaign Type" },
      { compilationHeader: "Bid Strategy Type", tttHeader: "Bid Strategy Type" },
      { compilationHeader: "Start Date SF", tttHeader: "Start Date SF" },
      { compilationHeader: "End Date SF", tttHeader: "End Date SF" },
      { compilationHeader: "Language", tttHeader: "Language" },
      { compilationHeader: "Rotation", tttHeader: "Rotation" },
      { compilationHeader: "Device Targeting", tttHeader: "Device Targeting" }, 
      { compilationHeader: "Multiformat ads", tttHeader: "Multiformat ads" }, 
      { compilationHeader: "Ad type", tttHeader: "Ad type" },
      { compilationHeader: "Ad Group Type", tttHeader: "Ad Group Type" },
      { compilationHeader: "Client Rate", tttHeader: "Client Rate (converted)" }, // <<< NY KOLUMN HÄR
      { compilationHeader: "Ad Name", tttHeader: "Ad Name" },
      { compilationHeader: "Video ID", tttHeader: "Video ID" },
      { compilationHeader: "Tracking template", tttHeader: "Tracking template" },
      { compilationHeader: "Final URL", tttHeader: "Final URL" },
      { compilationHeader: "Display URL", tttHeader: "Display URL" },
      { compilationHeader: "CTA", tttHeader: "CTA" },
      { compilationHeader: "Headline", tttHeader: "Headline" },
      { compilationHeader: "Long Headline", tttHeader: "Long Headline" },
      { compilationHeader: "Description 1", tttHeader: "Description 1" },
      { compilationHeader: "Description 2", tttHeader: "Description 2" }
    ];

    const compilationHeadersOutput = requiredColsSetup.map(col => col.compilationHeader);
    const totalColumnsOutput = compilationHeadersOutput.length;

    const colIndexesTTT = [];
    const missingColsTTT = [];
    requiredColsSetup.forEach(colSetup => {
      const index = findHeaderIndex_(headersTTT, colSetup.tttHeader); 
      if (index === -1) {
        missingColsTTT.push(colSetup.tttHeader);
      }
      colIndexesTTT.push(index - 1); 
    });

    if (missingColsTTT.length > 0) {
      throw new Error(`Saknar följande kolumner i TTT-arkets rubrikrad (rad 2): ${missingColsTTT.join(", ")}`);
    }

    const campaignNameColIdxTTT = colIndexesTTT[0]; 
    const videoIdColIdxTTT = colIndexesTTT[requiredColsSetup.findIndex(c => c.compilationHeader === "Video ID")]; 

    const campaignMap = {};
    const dataRows = allData.slice(dataStartIndex); 
    Logger.log(`buildCampaignSummary: Bearbetar ${dataRows.length} datarader från TTT.`);

    dataRows.forEach((row, rowIndex) => {
      if (row.every(cell => String(cell || "").trim() === "")) { 
          Logger.log(`buildCampaignSummary: Hoppar över helt tom rad ${dataStartIndex + rowIndex + 1} i TTT.`);
          return;
      }
      const maxNeededIndex = Math.max(...colIndexesTTT.filter(idx => idx !== undefined && idx > -2)); // idx > -2 eftersom -1 är ogiltigt
      if (row.length <= maxNeededIndex) {
          Logger.log(`buildCampaignSummary: Hoppar över rad ${dataStartIndex + rowIndex + 1} i TTT, för få kolumner (${row.length} vs ${maxNeededIndex + 1} behövs).`);
          return;
      }
      const campaignName = String(row[campaignNameColIdxTTT] || "").trim();
      if (!campaignName) { 
          Logger.log(`buildCampaignSummary: Hoppar över rad ${dataStartIndex + rowIndex + 1} i TTT, saknar kampanjnamn i kolumn ${campaignNameColIdxTTT + 1}.`);
          return;
      }

      const numCampaignInfoCols = 9; 

      if (!campaignMap[campaignName]) {
        const campaignInfo = [];
        for (let k = 0; k < numCampaignInfoCols; k++) {
            campaignInfo.push(row[colIndexesTTT[k]] ?? "");
        }
        campaignMap[campaignName] = { info: campaignInfo, ads: [] };
      }

      const adData = [];
      // totalColumnsOutput är antalet kolumner vi vill ha i Compilation.
      // colIndexesTTT har samma längd och mappar till TTT.
      for (let k = numCampaignInfoCols; k < totalColumnsOutput; k++) { 
          adData.push(row[colIndexesTTT[k]] ?? "");
      }
      const videoIdValue = row[videoIdColIdxTTT] ?? "";

      if (String(videoIdValue).trim() !== "") {
        const rotationIndexInTTTArray = requiredColsSetup.findIndex(c => c.compilationHeader === "Rotation");
        const rotationValue = row[colIndexesTTT[rotationIndexInTTTArray]] ?? "";
        campaignMap[campaignName].ads.push({ data: adData, rotation: rotationValue });
      }
    });

    const uniqueCampaignCount = Object.keys(campaignMap).length;
    Logger.log(`buildCampaignSummary: Hittade ${uniqueCampaignCount} unika kampanjer att bearbeta.`);
    if (uniqueCampaignCount === 0) { ui.alert(`Kunde inte identifiera några kampanjer med annonser i TTT-arket.`); return; }

    const output = [];
    output.push(compilationHeadersOutput); 

    const numCampaignInfoColsOutput = 9; // Samma som numCampaignInfoCols

    for (const campName in campaignMap) {
      const campaignData = campaignMap[campName];
      const campaignInfo = campaignData.info; 
      const ads = campaignData.ads;          

      const campaignRowOutput = [...campaignInfo, ...Array(totalColumnsOutput - numCampaignInfoColsOutput).fill("")];
      const adTypeIndexInOutput = compilationHeadersOutput.indexOf("Ad type");
      if (adTypeIndexInOutput >= numCampaignInfoColsOutput) campaignRowOutput[adTypeIndexInOutput] = ""; 
      output.push(campaignRowOutput);

      ads.forEach(adObject => {
        const adDataArray = adObject.data; 
        const rotationValue = adObject.rotation;
        const adRowCampaignPlaceholders = Array(numCampaignInfoColsOutput).fill("");
        const rotationIndexInOutput = compilationHeadersOutput.indexOf("Rotation");
        if (rotationIndexInOutput < numCampaignInfoColsOutput && rotationIndexInOutput !== -1) { 
            adRowCampaignPlaceholders[rotationIndexInOutput] = rotationValue;
        }
        const adRowOutput = [...adRowCampaignPlaceholders, ...adDataArray];
        output.push(adRowOutput);
      });
      output.push(Array(totalColumnsOutput).fill("")); 
    }

    if (output.length > 1 && output[output.length - 1].every(cell => cell === "")) { output.pop(); } 

    let summarySheet = ss.getSheetByName(compilationSheetName);
    if (!summarySheet) { summarySheet = ss.insertSheet(compilationSheetName); }
    else { summarySheet.clearContents(); summarySheet.clearFormats(); } 
    summarySheet.showSheet(); summarySheet.activate();

    if (output.length > 1) { 
      summarySheet.getRange(1, 1, output.length, totalColumnsOutput).setValues(output);
      summarySheet.getRange(1, 1, 1, totalColumnsOutput).setFontWeight("bold").setBackground("#f0f0f0");
      summarySheet.autoResizeColumns(1, totalColumnsOutput);
      ui.alert(`Sammanställning klar i arket "${compilationSheetName}"!`);
    } else {
      ui.alert(`Kunde inte bygga sammanställningen. Ingen data att skriva.`);
      Logger.log("buildCampaignSummary: Ingen outputdata genererades.");
    }
     Logger.log("buildCampaignSummary: Klar.");
  } catch (e) {
    Logger.log(`Fel i buildCampaignSummary: ${e}\nStack: ${e.stack || 'Ingen stack tillgänglig'}`);
    ui.alert(`Fel vid bygge av Compilation-arket:\n\n${e.message}\n\nKontrollera loggen (Utföranden).`);
    SpreadsheetApp.getActiveSpreadsheet().toast("Fel vid bygge av Compilation!", "Fel", 10);
  }
}

// ===============================================
// Lägg till onOpen_TargetingHelper i huvudmenyn
// ===============================================

/**
 * Utökar onOpen för att inkludera Targeting Helper-menyn
 */
function onOpen_TargetingHelper() {
  SpreadsheetApp.getUi()
      .createMenu('Targeting Helper')
      .addItem('Populate Targeting IDs', 'runPopulateTargetingIDs')
      .addItem('Add Targeting to Existing Campaigns', 'showTargetingPopup_')
      .addToUi();
}
