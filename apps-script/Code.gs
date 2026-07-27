const SPREADSHEET_ID = '10dvG_p7S-SuOTslaI048d3Hd-wobV34F_ZJf1SfwG5g';
const INVENTORY_SHEET = 'Inventory';
const LOG_SHEET = 'Scan Log';

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = safeCallback_(p.callback);
  let result;

  try {
    if (p.action === 'ping') {
      result = { ok: true, message: 'Backend connected' };
    } else if (p.action === 'add') {
      result = addCardFromRequest_(p);
    } else if (p.action === 'undo') {
      result = undoLastScan_();
    } else {
      result = { ok: true, message: 'Pokémon Rapid Scanner backend' };
    }
  } catch (error) {
    result = { ok: false, error: error.message || String(error) };
  }

  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let inventory = ss.getSheetByName(INVENTORY_SHEET);
  let log = ss.getSheetByName(LOG_SHEET);

  if (!inventory) inventory = ss.insertSheet(INVENTORY_SHEET);
  if (!log) log = ss.insertSheet(LOG_SHEET);

  const inventoryHeaders = [
    'Card Key', 'Card ID', 'Name', 'Set', 'Set Code', 'Number',
    'Rarity', 'Supertype', 'Subtypes', 'HP', 'Abilities', 'Attacks',
    'Finish', 'Quantity', 'Image URL', 'Last Scanned'
  ];

  const logHeaders = [
    'Timestamp', 'Card ID', 'Name', 'Set', 'Number', 'Finish',
    'Previous Qty', 'New Qty', 'Confidence', 'OCR Name',
    'OCR Number', 'OCR Text', 'Status', 'Card Key'
  ];

  if (inventory.getLastRow() === 0 || inventory.getRange(1, 1).getValue() !== 'Card Key') {
    inventory.clear();
    inventory.getRange(1, 1, 1, inventoryHeaders.length).setValues([inventoryHeaders]);
    inventory.setFrozenRows(1);
    inventory.getRange(1, 1, 1, inventoryHeaders.length)
      .setFontWeight('bold').setBackground('#fbbc04');
  }

  if (log.getLastRow() === 0 || log.getRange(1, 1).getValue() !== 'Timestamp') {
    log.clear();
    log.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
    log.setFrozenRows(1);
    log.getRange(1, 1, 1, logHeaders.length)
      .setFontWeight('bold').setBackground('#34a853').setFontColor('#ffffff');
  }

  inventory.autoResizeColumns(1, inventoryHeaders.length);
  log.autoResizeColumns(1, logHeaders.length);
  return { ok: true };
}

function addCardFromRequest_(p) {
  setupSheets();

  const card = JSON.parse(p.card || '{}');
  if (!card.id) throw new Error('Missing card data.');

  const finish = p.finish || 'Unspecified';
  const key = card.id + '|' + finish.toLowerCase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const inventory = ss.getSheetByName(INVENTORY_SHEET);
  const log = ss.getSheetByName(LOG_SHEET);
  const now = new Date();

  const found = inventory.getRange('A:A').createTextFinder(key)
    .matchEntireCell(true).findNext();

  let previousQty = 0;
  let newQty = 1;

  if (found) {
    const row = found.getRow();
    previousQty = Number(inventory.getRange(row, 14).getValue()) || 0;
    newQty = previousQty + 1;
    inventory.getRange(row, 14).setValue(newQty);
    inventory.getRange(row, 16).setValue(now);
  } else {
    inventory.appendRow([
      key,
      card.id || '',
      card.name || '',
      card.set || '',
      card.setCode || '',
      card.number || '',
      card.rarity || '',
      card.supertype || '',
      Array.isArray(card.subtypes) ? card.subtypes.join(', ') : '',
      card.hp || '',
      card.abilities || '',
      card.attacks || '',
      finish,
      1,
      card.image || '',
      now
    ]);
  }

  log.appendRow([
    now,
    card.id || '',
    card.name || '',
    card.set || '',
    card.number || '',
    finish,
    previousQty,
    newQty,
    p.confidence || '',
    p.ocrName || '',
    p.ocrNumber || '',
    p.ocrText || '',
    'Added',
    key
  ]);

  return { ok: true, quantity: newQty, cardKey: key, name: card.name };
}

function undoLastScan_() {
  setupSheets();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const inventory = ss.getSheetByName(INVENTORY_SHEET);
  const log = ss.getSheetByName(LOG_SHEET);

  if (log.getLastRow() < 2) return { ok: false, message: 'Nothing to undo.' };

  const row = log.getLastRow();
  const values = log.getRange(row, 1, 1, 14).getValues()[0];
  const key = values[13];
  const found = inventory.getRange('A:A').createTextFinder(key)
    .matchEntireCell(true).findNext();

  if (found) {
    const invRow = found.getRow();
    const qty = Number(inventory.getRange(invRow, 14).getValue()) || 0;
    if (qty <= 1) inventory.deleteRow(invRow);
    else inventory.getRange(invRow, 14).setValue(qty - 1);
  }

  log.deleteRow(row);
  return { ok: true, name: values[2] };
}

function safeCallback_(value) {
  const callback = String(value || '');
  return /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback) ? callback : '';
}
