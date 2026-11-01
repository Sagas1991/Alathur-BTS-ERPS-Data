/**
 * Alathur NW ERPS Migration — Field App
 * Backend for the "Work Data" sheet.
 *
 * SETUP:
 * 1. Open the spreadsheet → Extensions → Apps Script.
 * 2. Replace Code.gs content with this file.
 * 3. Create a new HTML file named exactly "Index" and paste Index.html into it.
 * 4. Deploy → New deployment → Web app.
 *      Execute as: Me
 *      Who has access: Anyone with the link (or "Anyone within [org]" if you have Workspace)
 * 5. Open the resulting URL on your phone and add it to your home screen.
 *
 * If you rename the "Work Data" tab, update SHEET_NAME below to match.
 */

const SHEET_NAME = 'Work Data';
const FIRST_DATA_ROW = 3; // row 1 = group headers, row 2 = column headers, data starts row 3

// 1-indexed column map, matching the sheet exactly as of 14-Jul-2026
// NOTE: column E (A_END) is now "Site is in Ring/Linear" (dropdown: Linear / Ring /
// Linear but co-located) and column G (B_END) is no longer used for any logic.
const COL = {
  SL_NO: 1, CLUSTER: 2, SOURCE_NODE: 3, SOURCE_IP: 4, A_END: 5, ENODE_B: 6, B_END: 7,
  // H = Cluster Diagram (image) — intentionally skipped
  VLAN4G_1: 9, VLAN4G_2: 10, VLAN4G_3: 11, VLAN4G_4: 12, IWF_MAC: 13,
  CPU_PORT_MGMT_VLAN: 14, CPU_PORT_IP: 15, NEXT_HOP_IP_NETMASK: 16,
  VLAN_3G: 17, RNC: 18, RNC_PORT: 19,
  BSC: 20, UPLINK_BSC_PORT: 21, BSC_EXIST_PORT: 22, BSC_NEW_PORT: 23,
  VLAN_2G: 24, E1: 25, CEF_MAC: 26, IWF_MAC_C1: 27,
  SINK_NODE: 28, SINK_PORT: 29, SINK_IP: 30, SINK_NE: 31,
  REQUEST_DATE: 32, DONE_DATE: 33,
  P1: 34, P2: 35, P3: 36, P4: 37, P5: 38, E1_12: 39, E2_34: 40,
  CEF_CONFIG_STATUS: 41, MIGRATION_2G: 42, MIGRATION_3G: 43,
  // 44 (AR) = Print hyperlink column — not read by the app
  ERPS_STATUS: 45, VRRP_STATUS: 46
};

const LAST_COL = COL.VRRP_STATUS;

// Fields the field-app is allowed to write back. Everything else is read-only from the phone.
const EDITABLE_FIELDS = {
  sinkPort: COL.SINK_PORT,
  doneDate: COL.DONE_DATE,
  p1: COL.P1, p2: COL.P2, p3: COL.P3, p4: COL.P4, p5: COL.P5,
  e1_12: COL.E1_12, e2_34: COL.E2_34,
  cefConfigStatus: COL.CEF_CONFIG_STATUS, migration2g: COL.MIGRATION_2G, migration3g: COL.MIGRATION_3G
};

// ---------------------------------------------------------------------------
// ADMIN ACCESS (everyone else is view-only)
// ---------------------------------------------------------------------------
// Anyone with the app link can VIEW. Only someone who logs in with the admin
// PIN gets a token that the server accepts for saving. The check is done on
// the server for every save, so hiding buttons in the app is not the only
// protection. Set or change the PIN from the sheet menu:
//   📋 ERPS Reports > 🔐 Set Admin PIN

const ADMIN_TOKEN_TTL_SECONDS = 21600; // 6 hours, the longest Apps Script cache allows

function hashPin_(pin, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

/** Menu item: the sheet owner sets or changes the admin PIN. Changing it logs out every device. */
function setAdminPin() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('Set admin PIN',
    'Enter a new admin PIN (at least 4 characters). Only people who know it can edit data in the mobile app.',
    ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const pin = res.getResponseText().trim();
  if (pin.length < 4) { ui.alert('The PIN must be at least 4 characters. Nothing was changed.'); return; }
  const salt = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({
    ADMIN_PIN_SALT: salt,
    ADMIN_PIN_HASH: hashPin_(pin, salt),
    ADMIN_EPOCH: Utilities.getUuid()   // a new epoch makes every older login invalid
  });
  ui.alert('Admin PIN saved. Anyone who was logged in as admin has been logged out.');
}

/** Called by the app's login box. Returns {ok:true, token} or {ok:false, error}. */
function adminLogin(pin) {
  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty('ADMIN_PIN_SALT');
  const hash = props.getProperty('ADMIN_PIN_HASH');
  if (!salt || !hash) {
    return { ok: false, error: 'No admin PIN has been set yet. Set it from the sheet menu: ERPS Reports > Set Admin PIN.' };
  }
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('admin_fails') || 0);
  if (fails >= 5) return { ok: false, error: 'Too many wrong attempts. Try again in 10 minutes.' };

  if (hashPin_(String(pin || '').trim(), salt) !== hash) {
    cache.put('admin_fails', String(fails + 1), 600);
    return { ok: false, error: 'Wrong PIN.' };
  }
  cache.remove('admin_fails');
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('admin_' + token, props.getProperty('ADMIN_EPOCH') || 'none', ADMIN_TOKEN_TTL_SECONDS);
  return { ok: true, token: token };
}

function isAdminToken_(token) {
  if (!token) return false;
  const stored = CacheService.getScriptCache().get('admin_' + token);
  if (stored === null) return false;
  return stored === (PropertiesService.getScriptProperties().getProperty('ADMIN_EPOCH') || 'none');
}

/** Used by the app to confirm a saved login is still valid. */
function checkAdmin(token) { return isAdminToken_(token); }

function adminLogout(token) {
  if (token) CacheService.getScriptCache().remove('admin_' + token);
  return true;
}

/** Every function that changes data calls this first. */
function requireAdmin_(token) {
  if (!isAdminToken_(token)) throw new Error('NOT_ADMIN');
}

/** For maintenance helpers: only the script owner, running from the editor, may use them. */
function isOwnerRequest_() {
  const active = Session.getActiveUser().getEmail();
  const owner = Session.getEffectiveUser().getEmail();
  return !!active && !!owner && active === owner;
}

function requireOwner_() {
  if (!isOwnerRequest_()) throw new Error('Only the sheet owner can run this.');
}

// ---------------------------------------------------------------------------
// SINK NODE PORTS
// ---------------------------------------------------------------------------

const SINK_PORTS_SHEET_NAME = 'Sink Node Ports';
// Layout: A=Sink Node, B=Port ID, C=Connected To

// Google Sheets auto-detects things like "1/1/1" as a date the instant
// they're written, even with the cell pre-formatted as plain text. To make
// this bug structurally impossible, Port IDs are stored using a lookalike
// "fraction slash" (U+2044) instead of a real '/' — visually identical,
// but invisible to Sheets' date parser. Converted back to a normal '/' on
// every read/display so the person never sees the difference.
const SAFE_SLASH_ = '\u2044';
function toSafeSlash_(s) { return (s === undefined || s === null) ? '' : s.toString().split('/').join(SAFE_SLASH_); }
function fromSafeSlash_(s) { return (s === undefined || s === null) ? '' : s.toString().split(SAFE_SLASH_).join('/'); }

function getSinkPortsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SINK_PORTS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SINK_PORTS_SHEET_NAME);
    sheet.getRange(1, 1, 1, 3).setValues([['Sink Node', 'Port ID', 'Connected To']]);
  }
  // Lock column B (Port ID) to plain text so Sheets never silently
  // auto-converts values like "1/1" into an actual date again.
  sheet.getRange(2, 2, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  return sheet;
}

/** Converts a cell value back to safe display text: recovers a normal '/'
 * from the safe-slash marker, and also recovers from any older row that
 * Sheets already silently auto-parsed as a Date (e.g. "1/1" -> 1/1/2001). */
function portIdToText_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Kolkata', 'M/d');
  }
  return fromSafeSlash_(v);
}

function comparePortIds_(a, b) {
  const pa = portIdToText_(a).toString().split('/').map(Number);
  const pb = portIdToText_(b).toString().split('/').map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const x = pa[i] || 0, y = pb[i] || 0;
    if (x !== y) return x - y;
  }
  return 0;
}

function nextPortId_(existingPorts) {
  if (existingPorts.length === 0) return '1/1/1';
  let maxPort = existingPorts[0].portId;
  existingPorts.forEach(function (p) { if (comparePortIds_(p.portId, maxPort) > 0) maxPort = p.portId; });
  const parts = maxPort.toString().split('/').map(Number);
  parts[parts.length - 1] = (parts[parts.length - 1] || 0) + 1;
  return parts.join('/');
}

/** Returns every distinct Sink Node name, from Work Data plus any already tracked in Sink Node Ports. */
function getSinkNodes() {
  const workSheet = getSheet_();
  const lastRow = workSheet.getLastRow();
  const set = {};
  if (lastRow >= FIRST_DATA_ROW) {
    const values = workSheet.getRange(FIRST_DATA_ROW, COL.SINK_NODE, lastRow - FIRST_DATA_ROW + 1, 1).getValues();
    values.forEach(function (v) { const s = clean_(v[0]); if (s) set[s] = true; });
  }
  const portSheet = getSinkPortsSheet_();
  const lastPortRow = portSheet.getLastRow();
  if (lastPortRow >= 2) {
    const nodeVals = portSheet.getRange(2, 1, lastPortRow - 1, 1).getValues();
    nodeVals.forEach(function (v) { const s = clean_(portIdToText_(v[0])); if (s) set[s] = true; });
  }
  return Object.keys(set).sort();
}

/** Returns all ports currently tracked for one Sink Node, sorted by port ID. */
function getSinkNodePorts(nodeName) {
  const sheet = getSinkPortsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const target = clean_(nodeName).toLowerCase();
  const ports = [];
  values.forEach(function (r, i) {
    const rowSite = clean_(portIdToText_(r[0])).toLowerCase();
    if (rowSite === target) {
      ports.push({ rowIndex: i + 2, portId: portIdToText_(r[1]), connectedTo: portIdToText_(r[2]) });
    }
  });
  ports.sort(function (a, b) { return comparePortIds_(a.portId, b.portId); });
  return ports;
}

/** Appends a new port row for a node with an auto-suggested next port ID. */
function addSinkPort(nodeName, token) {
  requireAdmin_(token);
  return addSinkPort_(nodeName);
}

function addSinkPort_(nodeName) {
  const existing = getSinkNodePorts(nodeName);
  const newId = nextPortId_(existing);
  const sheet = getSinkPortsSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow([nodeName, toSafeSlash_(newId), '']);
  } finally {
    lock.releaseLock();
  }
  return { portId: newId };
}

/** Deletes one port row by its sheet row index. */
function deleteSinkPort(rowIndex, token) {
  requireAdmin_(token);
  const sheet = getSinkPortsSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.deleteRow(Number(rowIndex));
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

/** Batch-saves Port ID + Connected To for every currently-displayed port row of a node. */
function saveSinkNodePorts(nodeName, ports, token) {
  requireAdmin_(token);
  const sheet = getSinkPortsSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    ports.forEach(function (p) {
      if (p.rowIndex) {
        sheet.getRange(p.rowIndex, 2, 1, 2).setValues([[toSafeSlash_(p.portId), p.connectedTo]]);
      }
    });
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// DEBUG / MAINTENANCE HELPERS — not used by the web app.
// Run manually from the Apps Script editor (pick the function in the
// dropdown at the top, then click Run) and read the Execution log.
// ---------------------------------------------------------------------------

function debugKoyalmannam() {
  const workSheet = getSheet_();
  const lastRow = workSheet.getLastRow();
  const values = workSheet.getRange(FIRST_DATA_ROW, 1, lastRow - FIRST_DATA_ROW + 1, LAST_COL).getValues();

  Logger.log('--- Searching Work Data for any ENode B containing "koyalmannam" ---');
  values.forEach(function (row, i) {
    const enodeB = (row[COL.ENODE_B - 1] || '').toString();
    if (enodeB.toLowerCase().indexOf('koyalmannam') !== -1) {
      Logger.log('Row ' + (FIRST_DATA_ROW + i) + ' ENode B: "' + enodeB + '" (length ' + enodeB.length + ')');
      Logger.log('Char codes: ' + enodeB.split('').map(function (c) { return c.charCodeAt(0); }).join(','));
    }
  });

  Logger.log('--- Searching Power_Levels for any site containing "koyalmannam" ---');
  const powerSheet = getPowerSheet_();
  if (!powerSheet) { Logger.log('Power_Levels sheet not found!'); return; }
  const powerLastRow = powerSheet.getLastRow();
  const powerValues = powerSheet.getRange(2, 1, Math.max(powerLastRow - 1, 0), 1).getValues();
  powerValues.forEach(function (row, i) {
    const site = (row[0] || '').toString();
    if (site.toLowerCase().indexOf('koyalmannam') !== -1) {
      Logger.log('Power_Levels row ' + (i + 2) + ' Site Name: "' + site + '" (length ' + site.length + ')');
      Logger.log('Char codes: ' + site.split('').map(function (c) { return c.charCodeAt(0); }).join(','));
    }
  });
}

function debugSinkPortFix() {
  requireOwner_();
  const sheet = getSinkPortsSheet_();
  Logger.log('--- Simulating Add Port for a brand-new test node ---');
  const testNode = 'DEBUG_TEST_NODE';
  const result = addSinkPort_(testNode);
  Logger.log('addSinkPort returned: ' + JSON.stringify(result));

  const raw = sheet.getRange(sheet.getLastRow(), 2).getValue();
  Logger.log('Raw cell value just written, typeof: ' + typeof raw + ', toString: ' + raw);
  Logger.log('Is it a Date object: ' + (Object.prototype.toString.call(raw) === '[object Date]'));

  const readBack = getSinkNodePorts(testNode);
  Logger.log('Read back via getSinkNodePorts: ' + JSON.stringify(readBack));

  // Clean up the test row
  sheet.deleteRow(sheet.getLastRow());
  Logger.log('Test row removed.');
}

/**
 * ONE-TIME CLEANUP — removes corrupted/test rows (date-mangled Port IDs,
 * bare numbers, blank nodes) from the Sink Node Ports sheet.
 */
function cleanupBadSinkPortRows() {
  requireOwner_();
  const sheet = getSinkPortsSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('Nothing to clean.'); return; }
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const rowsToDelete = [];
  values.forEach(function (r, i) {
    const siteName = clean_(r[0]);
    const rawPortId = r[1];
    const isDateMangled = Object.prototype.toString.call(rawPortId) === '[object Date]';
    const isBareNumber = typeof rawPortId === 'number';
    const isBlankNode = !siteName;
    if (isDateMangled || isBareNumber || isBlankNode) rowsToDelete.push(i + 2);
  });
  Logger.log('Deleting rows: ' + JSON.stringify(rowsToDelete));
  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sheet.deleteRow(r); });
  Logger.log('Done. Deleted ' + rowsToDelete.length + ' row(s).');
}

function debugSinkPorts() {
  Logger.log('--- getSinkNodes() ---');
  Logger.log(JSON.stringify(getSinkNodes()));

  Logger.log('--- getSinkNodePorts("Thenkurissi") ---');
  Logger.log(JSON.stringify(getSinkNodePorts('Thenkurissi')));

  const sheet = getSinkPortsSheet_();
  Logger.log('--- Sink Node Ports sheet raw rows ---');
  const lastRow = sheet.getLastRow();
  Logger.log('Last row: ' + lastRow);
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    values.forEach(function (r, i) { Logger.log((i + 2) + ': ' + JSON.stringify(r)); });
  }
}

function debugPower() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getPowerSheet_();
  Logger.log('Power Levels sheet found: ' + (sheet !== null));
  if (!sheet) {
    Logger.log('Available sheet names: ' + ss.getSheets().map(function (s) { return s.getName(); }).join(', '));
    return;
  }
  Logger.log('Last row: ' + sheet.getLastRow());
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 6).getValues();
  Logger.log('Row data:');
  values.forEach(function (r, i) { Logger.log((i + 2) + ': ' + JSON.stringify(r)); });

  Logger.log('--- getPowerData("RP-04703-Vemballur") ---');
  Logger.log(JSON.stringify(getPowerData('RP-04703-Vemballur')));

  Logger.log('--- getPowerData(sheet ENode B value from row 4) ---');
  const workSheet = getSheet_();
  const actualEnodeB = workSheet.getRange(4, COL.ENODE_B).getValue();
  Logger.log('Actual ENode B string: "' + actualEnodeB + '"');
  Logger.log(JSON.stringify(getPowerData(actualEnodeB)));
}

/** Diagnoses the Cluster Diagram cell for row 3 (change testRow below to test another cluster's first row). */
function debugDiagram() {
  const testRow = 3;
  const sheet = getSheet_();
  const value = sheet.getRange(testRow, DIAGRAM_COL).getValue();

  Logger.log('--- Cell H' + testRow + ' value ---');
  Logger.log('typeof: ' + typeof value);
  Logger.log('String form: ' + value);
  Logger.log('Is CellImage (has getContentUrl): ' + (value && typeof value.getContentUrl === 'function'));

  Logger.log('--- Merged ranges touching H' + testRow + ' ---');
  const mergedRanges = sheet.getRange(testRow, DIAGRAM_COL).getMergedRanges();
  mergedRanges.forEach(function (r) {
    Logger.log('Merged range: ' + r.getA1Notation());
  });

  Logger.log('--- Floating images on sheet ---');
  const images = sheet.getImages();
  Logger.log('Count: ' + images.length);
  images.forEach(function (img, i) {
    const a = img.getAnchorCell();
    Logger.log('Image ' + i + ' anchored at row=' + a.getRow() + ' col=' + a.getColumn());
  });

  Logger.log('--- getClusterDiagram(' + testRow + ') result ---');
  const contentUrl = (value && typeof value.getContentUrl === 'function') ? value.getContentUrl() : null;
  Logger.log('contentUrl: ' + contentUrl);
  if (contentUrl) {
    try {
      const resp = UrlFetchApp.fetch(contentUrl, { muteHttpExceptions: true });
      Logger.log('Fetch response code: ' + resp.getResponseCode());
      Logger.log('Fetch content-type: ' + resp.getBlob().getContentType());
      Logger.log('Fetch bytes length: ' + resp.getBlob().getBytes().length);
    } catch (e) {
      Logger.log('Fetch threw: ' + e.message);
    }
  }
  const result = getClusterDiagram(testRow);
  Logger.log(result ? ('Got data URI, length=' + result.length) : 'Returned null');
}

function debugGetSiteData() {
  Logger.log('--- Testing getClusters() ---');
  Logger.log(JSON.stringify(getClusters()));

  Logger.log('--- Testing getSites("PKD_Cluster-14") ---');
  Logger.log(JSON.stringify(getSites('PKD_Cluster-14')));

  Logger.log('--- Testing getSiteData(8) ---');
  Logger.log(JSON.stringify(getSiteData(8)));
}

// ---------------------------------------------------------------------------
// SHEET MENU + STATUS REPORT (now includes Cluster Diagrams)
// ---------------------------------------------------------------------------

/**
 * Runs automatically whenever the spreadsheet is opened. Adds a custom
 * menu so the status report can be viewed right inside Google Sheets,
 * no need to open the mobile web app.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 ERPS Reports')
    .addItem('📊 View Dashboard', 'showDashboardDialog')
    .addItem('View Status Report', 'showStatusReportDialog')
    .addItem('📗 Export ERPS Report (Excel)', 'exportErpsReportExcel')
    .addItem('🔐 Set Admin PIN', 'setAdminPin')
    .addToUi();
}

/**
 * Opens the Migration Status Report (grouped by cluster, with ERPS/VRRP
 * status, the cluster diagram, and CEF/2G/3G migration pills) as a modal
 * dialog inside Google Sheets.
 *
 * Diagrams are loaded AFTER the dialog opens (a few at a time) rather than
 * being embedded up-front, so the report appears instantly even with many
 * large JPEGs. The Print button stays disabled until every diagram has
 * finished loading, so a printout never has gaps.
 */
function showStatusReportDialog() {
  const sites = sortByClusterNo_(getStatusReportData());

  function esc(s) {
    return (s === undefined || s === null) ? '' : s.toString()
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function pill(value) {
    const v = (value || '').toString().trim();
    const done = v === 'Completed';
    const cls = done ? 'done' : 'pending';
    return '<span class="pill ' + cls + '">' + esc(v || 'Not started') + '</span>';
  }

  const now = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Kolkata', 'dd MMM yyyy');

  let rowsHtml = '';
  let lastCluster = null;
  const diagramJobs = [];

  sites.forEach(function (site) {
    if (site.cluster !== lastCluster) {
      if (lastCluster !== null) rowsHtml += '</tbody></table></div>';
      rowsHtml += '<div class="cluster-block"><table><tbody>';
      rowsHtml += '<tr class="cluster-row"><td colspan="4">' + esc(site.cluster) +
        '<span class="extra">ERPS: ' + esc(site.erpsStatus || 'Not started') +
        ' &nbsp;·&nbsp; VRRP: ' + esc(site.vrrpStatus || 'Not started') + '</span></td></tr>';

      if (site.diagramRow) {
        const id = 'diag-' + diagramJobs.length;
        diagramJobs.push({ id: id, row: site.diagramRow });
        rowsHtml += '<tr class="diagram-row"><td colspan="4"><div id="' + id +
          '" class="diagram-box">Loading diagram…</div></td></tr>';
      }

      // Column titles repeated for every cluster, directly above its sites
      rowsHtml += '<tr class="col-head"><td>ENode B</td><td>CEF Config</td>' +
        '<td>2G Migration</td><td>3G Migration</td></tr>';

      lastCluster = site.cluster;
    }
    rowsHtml += '<tr><td>' + esc(site.enodeB) + '</td><td>' + pill(site.cefConfigStatus) +
      '</td><td>' + pill(site.migration2g) + '</td><td>' + pill(site.migration3g) + '</td></tr>';
  });

  if (lastCluster !== null) rowsHtml += '</tbody></table></div>';

  const hasJobs = diagramJobs.length > 0;
  const jobsJson = JSON.stringify(diagramJobs).replace(/</g, '\\u003c');

  const html = '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:20px;background:#fff;color:#111;}' +
    '.creator{text-align:right;font-size:11px;font-weight:700;color:#1F3B57;margin-bottom:8px;}' +
    '.title{text-align:center;margin-bottom:6px;}' +
    '.title h1{font-size:18px;color:#1F3B57;margin:0 0 2px;}' +
    '.title .sub{font-size:12px;color:#555;font-weight:600;}' +
    '.meta{text-align:center;font-size:11px;color:#666;margin-bottom:16px;}' +
    'table{width:100%;border-collapse:collapse;font-size:12.5px;}' +
    'th{background:#1F3B57;color:#fff;text-align:left;padding:8px 10px;font-size:11px;text-transform:uppercase;}' +
    '.cluster-row td{background:#dce6f0;font-weight:700;padding:7px 10px;color:#1F3B57;border-top:2px solid #1F3B57;}' +
    '.extra{float:right;font-weight:600;font-size:11px;color:#3a5f83;}' +
    '.diagram-row td{padding:8px 10px;text-align:center;border-bottom:1px solid #ddd;background:#fafbfc;}' +
    '.col-head td{background:#eef2f7;color:#111;font-weight:800;font-size:11px;text-transform:uppercase;padding:7px 10px;border-top:1px solid #9aa9b8;border-bottom:2px solid #1F3B57;}' +
    '.cluster-block{page-break-after:always;break-after:page;margin-bottom:18px;}' +
    '.cluster-block:last-child{page-break-after:auto;break-after:auto;margin-bottom:0;}' +
    '@page{margin:12mm;}' +
    '.diagram-box{font-size:11px;color:#888;}' +
    '.diagram-box img{max-width:100%;max-height:400px;object-fit:contain;}' +
    'td{padding:6px 10px;border-bottom:1px solid #ddd;}' +
    '.pill{display:inline-block;padding:2px 8px;border-radius:999px;font-weight:700;font-size:11px;}' +
    '.pill.done{background:#d7f2e0;color:#1a7a44;}' +
    '.pill.pending{background:#fdecc8;color:#8a5a00;}' +
    '.printBtn{display:block;width:max-content;margin:16px auto 0;padding:10px 20px;background:#43D17A;border:none;border-radius:8px;font-weight:700;font-size:14px;cursor:pointer;}' +
    '.printBtn:disabled{background:#b8c4bd;cursor:default;}' +
    'tr{page-break-inside:avoid;}' +
    '@media print{ .no-print{display:none !important;} }' +
    '</style></head><body>' +
    '<div class="creator">Created by Sarath Ganga S, JE</div>' +
    '<div class="title"><h1>ALATHUR NW ERPS MIGRATION</h1><div class="sub">Migration Status Report</div></div>' +
    '<div class="meta">Generated on ' + esc(dateStr) + ' &nbsp;·&nbsp; ' + sites.length + ' sites</div>' +
    '<div id="blocks">' + rowsHtml + '</div>' +
    '<button id="printBtn" class="printBtn no-print" onclick="window.print()"' + (hasJobs ? ' disabled' : '') + '>' +
    (hasJobs ? 'Loading diagrams…' : '🖨 Print / Save PDF') + '</button>' +
    '<script>' +
    'var jobs = ' + jobsJson + ';' +
    'var pending = jobs.length, next = 0;' +
    'var btn = document.getElementById("printBtn");' +
    'function finishOne(){ pending--; if (pending <= 0) { btn.disabled = false; btn.textContent = "🖨 Print / Save PDF"; } }' +
    'function runNext(){' +
    '  if (next >= jobs.length) return;' +
    '  var j = jobs[next++];' +
    '  google.script.run' +
    '    .withSuccessHandler(function(uri){' +
    '      var el = document.getElementById(j.id);' +
    '      if (uri) { el.innerHTML = ""; var img = document.createElement("img"); img.src = uri; el.appendChild(img); }' +
    '      else { el.textContent = "No diagram available"; }' +
    '      finishOne(); runNext();' +
    '    })' +
    '    .withFailureHandler(function(){' +
    '      document.getElementById(j.id).textContent = "Could not load diagram";' +
    '      finishOne(); runNext();' +
    '    })' +
    '    .getClusterDiagram(j.row);' +
    '}' +
    'for (var k = 0; k < 3; k++) runNext();' +
    '</script>' +
    '</body></html>';

  const output = HtmlService.createHtmlOutput(html).setWidth(860).setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(output, 'Migration Status Report');
}

// ---------------------------------------------------------------------------
// DASHBOARD (inside Google Sheets)
// ---------------------------------------------------------------------------

/**
 * Opens a dashboard inside Google Sheets: overall counts (total sites,
 * CEF / 2G / 3G done, pending sites), a per-cluster breakdown, and every
 * pending site with its reason(s). The reason is simply the status text
 * currently entered in the CEF / 2G / 3G status cell ("Not started" if the
 * cell is blank). Sites where a technology does not exist (e.g. "No 2G",
 * "4G Only site") are not counted as pending.
 */
function showDashboardDialog() {
  const d = getDashboardData();
  const t = d.totals;
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Kolkata', 'dd MMM yyyy, hh:mm a');

  function esc(s) {
    return (s === undefined || s === null) ? '' : s.toString()
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function pct(a, b) { return b > 0 ? Math.round((a / b) * 100) : 0; }

  function card(label, value, sub, percent, cls) {
    return '<div class="card ' + cls + '"><div class="lbl">' + esc(label) + '</div>' +
      '<div class="val">' + esc(value) + '</div>' +
      '<div class="csub">' + esc(sub) + '</div>' +
      (percent === null ? '' : '<div class="bar"><div class="fill" style="width:' + percent + '%"></div></div>') +
      '</div>';
  }

  const pending2g = t.mig2gApplicable - t.mig2g;
  const pending3g = t.mig3gApplicable - t.mig3g;
  const pendingCef = t.cefApplicable - t.cef;

  const cards =
    card('Total Sites', t.sites, d.clusters.length + ' clusters', null, 'c-blue') +
    card('2G Migrated', t.mig2g, 'of ' + t.mig2gApplicable + ' · ' + pending2g + ' pending', pct(t.mig2g, t.mig2gApplicable), 'c-pink') +
    card('3G Migrated', t.mig3g, 'of ' + t.mig3gApplicable + ' · ' + pending3g + ' pending', pct(t.mig3g, t.mig3gApplicable), 'c-yellow') +
    card('CEF Configured', t.cef, 'of ' + t.cefApplicable + ' · ' + pendingCef + ' pending', pct(t.cef, t.cefApplicable), 'c-green') +
    card('Pending Sites', d.pending.length, 'sites with work left', null, 'c-red');

  // Per-cluster table
  const pendingByCluster = {};
  d.pending.forEach(function (p) {
    pendingByCluster[p.cluster] = (pendingByCluster[p.cluster] || 0) + 1;
  });

  let clusterRows = '';
  d.clusters.forEach(function (c) {
    const pc = pendingByCluster[c.cluster] || 0;
    clusterRows += '<tr><td class="cn">' + esc(c.cluster) + '</td><td>' + c.sites + '</td>' +
      '<td>' + c.cef + '/' + c.cefApplicable + '</td>' +
      '<td>' + c.mig2g + '/' + c.mig2gApplicable + '</td>' +
      '<td>' + c.mig3g + '/' + c.mig3gApplicable + '</td>' +
      '<td>' + (pc ? '<span class="badge bad">' + pc + ' pending</span>' : '<span class="badge ok">Complete</span>') + '</td></tr>';
  });

  // Pending sites with reasons, grouped by cluster
  let pendingHtml = '';
  let lastCluster = null;
  d.pending.forEach(function (p) {
    if (p.cluster !== lastCluster) {
      pendingHtml += '<div class="pc-head">' + esc(p.cluster) + '</div>';
      lastCluster = p.cluster;
    }
    function chip(label, status) {
      const cls = status === 'Not started' ? 'chip ns' : 'chip';
      return '<span class="' + cls + '"><b>' + label + '</b> ' + esc(status) + '</span>';
    }
    let reasons = '';
    if (!p.cefDone) reasons += chip('CEF', p.cefStatus);
    if (!p.mig2gDone) reasons += chip('2G', p.mig2gStatus);
    if (!p.mig3gDone) reasons += chip('3G', p.mig3gStatus);
    pendingHtml += '<div class="prow"><div class="pname">' + esc(p.site) + '</div><div class="reasons">' + reasons + '</div></div>';
  });
  if (!d.pending.length) {
    pendingHtml = '<div class="allgood">🎉 No pending sites. Everything is complete.</div>';
  }

  const html = '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:20px;background:#f4f6f9;color:#111;}' +
    '.creator{text-align:right;font-size:11px;font-weight:700;color:#1F3B57;margin-bottom:6px;}' +
    '.title{text-align:center;margin-bottom:4px;}' +
    '.title h1{font-size:18px;color:#1F3B57;margin:0 0 2px;}' +
    '.title .sub{font-size:12px;color:#555;font-weight:600;}' +
    '.meta{text-align:center;font-size:11px;color:#666;margin-bottom:16px;}' +
    '.cards{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:18px;}' +
    '.card{flex:1 1 150px;background:#fff;border-radius:10px;padding:12px 14px;border-top:4px solid #999;box-shadow:0 1px 3px rgba(0,0,0,0.12);}' +
    '.card .lbl{font-size:11px;font-weight:700;text-transform:uppercase;color:#666;}' +
    '.card .val{font-size:30px;font-weight:800;color:#111;margin:2px 0;}' +
    '.card .csub{font-size:11px;color:#777;}' +
    '.bar{height:6px;background:#e5e9ee;border-radius:3px;margin-top:8px;overflow:hidden;}' +
    '.fill{height:100%;background:#2f855a;}' +
    '.c-blue{border-top-color:#2b6cb0;} .c-pink{border-top-color:#d9777d;} .c-yellow{border-top-color:#d6a400;}' +
    '.c-green{border-top-color:#2f855a;} .c-red{border-top-color:#c53030;} .c-red .val{color:#c53030;}' +
    'h2{font-size:13px;color:#1F3B57;text-transform:uppercase;margin:18px 0 8px;}' +
    'table{width:100%;border-collapse:collapse;font-size:12.5px;background:#fff;border-radius:8px;overflow:hidden;}' +
    'th{background:#1F3B57;color:#fff;text-align:left;padding:8px 10px;font-size:11px;text-transform:uppercase;}' +
    'td{padding:7px 10px;border-bottom:1px solid #e3e7ec;}' +
    'td.cn{font-weight:700;color:#1F3B57;}' +
    '.badge{display:inline-block;padding:2px 9px;border-radius:999px;font-weight:700;font-size:11px;}' +
    '.badge.ok{background:#d7f2e0;color:#1a7a44;} .badge.bad{background:#fdecc8;color:#8a5a00;}' +
    '.plist{background:#fff;border-radius:8px;overflow:hidden;}' +
    '.pc-head{background:#dce6f0;color:#1F3B57;font-weight:700;font-size:12px;padding:7px 12px;border-top:2px solid #1F3B57;}' +
    '.prow{display:flex;gap:12px;padding:8px 12px;border-bottom:1px solid #e3e7ec;align-items:flex-start;}' +
    '.pname{flex:0 0 38%;font-weight:700;font-size:12.5px;}' +
    '.reasons{flex:1;display:flex;flex-wrap:wrap;gap:6px;}' +
    '.chip{background:#fdecc8;color:#6b4500;border-radius:6px;padding:3px 8px;font-size:11.5px;}' +
    '.chip.ns{background:#e8ebef;color:#555;}' +
    '.allgood{padding:24px;text-align:center;color:#1a7a44;font-weight:700;}' +
    '.printBtn{display:block;width:max-content;margin:18px auto 0;padding:10px 20px;background:#43D17A;border:none;border-radius:8px;font-weight:700;font-size:14px;cursor:pointer;}' +
    '.prow{page-break-inside:avoid;}' +
    '@media print{ .no-print{display:none !important;} body{background:#fff;} }' +
    '</style></head><body>' +
    '<div class="creator">Created by Sarath Ganga S, JE</div>' +
    '<div class="title"><h1>ALATHUR NW ERPS MIGRATION</h1><div class="sub">Migration Dashboard</div></div>' +
    '<div class="meta">As of ' + esc(stamp) + '</div>' +
    '<div class="cards">' + cards + '</div>' +
    '<h2>Cluster-wise progress</h2>' +
    '<table><thead><tr><th>Cluster</th><th>Sites</th><th>CEF</th><th>2G</th><th>3G</th><th>Status</th></tr></thead>' +
    '<tbody>' + clusterRows + '</tbody></table>' +
    '<h2>Pending sites and reasons (' + d.pending.length + ')</h2>' +
    '<div class="plist">' + pendingHtml + '</div>' +
    '<button class="printBtn no-print" onclick="window.print()">🖨 Print / Save PDF</button>' +
    '</body></html>';

  const output = HtmlService.createHtmlOutput(html).setWidth(920).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(output, 'Migration Dashboard');
}

// ---------------------------------------------------------------------------
// ERPS EXCEL REPORT (Ring / Linear per cluster, ERPS Completed / Not Feasible)
// ---------------------------------------------------------------------------

/**
 * Classifies a cluster's ERPS Status text (column "ERPS Status").
 * Matching is case-insensitive and ignores spaces, dashes and underscores:
 *   contains "not feasible" / "infeasible"  -> 'notfeasible'
 *   contains "complete" (but not "not complete"/"incomplete") -> 'completed'
 *   anything else, including blank          -> 'other'
 */
function erpsCategory_(status) {
  const compact = clean_(status).toLowerCase().replace(/[\s_\-]+/g, '');
  if (!compact) return 'other';
  if (compact.indexOf('notfeasible') !== -1 || compact.indexOf('infeasible') !== -1) return 'notfeasible';
  if (compact.indexOf('notcomplet') !== -1 || compact.indexOf('incomplet') !== -1) return 'other';
  if (compact.indexOf('complet') !== -1) return 'completed';
  return 'other';
}

/**
 * Turns the "Site is in Ring/Linear" dropdown value (column E) into a category:
 *   "Ring"                              -> 'ring'
 *   "Linear but co-located"             -> 'colocated'  (counted inside Linear)
 *   "Linear"                            -> 'linear'
 *   blank or anything else              -> 'unmarked'
 * Matching ignores capital letters, spaces, dashes and underscores.
 */
function topologyCategory_(value) {
  const compact = clean_(value).toLowerCase().replace(/[\s_\-]+/g, '');
  if (!compact) return 'unmarked';
  if (compact.indexOf('colocat') !== -1 || compact.indexOf('collocat') !== -1) return 'colocated';
  if (compact.indexOf('ring') !== -1) return 'ring';
  if (compact.indexOf('linear') !== -1) return 'linear';
  return 'unmarked';
}

/**
 * Counts, for every cluster, how many sites (ENode B) are marked Ring and
 * how many are Linear, using the "Site is in Ring/Linear" dropdown in
 * column E. "Linear but co-located" sites are included in the Linear count
 * and also counted separately. Column G (B End) is not used.
 */
function analyzeRingLinear_() {
  const rows = getDataRows_().filter(function (r) { return r.cluster && r.eNodeB; });
  const byCluster = {}, order = [], sites = [];

  rows.forEach(function (r) {
    if (!byCluster[r.cluster]) {
      byCluster[r.cluster] = {
        cluster: r.cluster, status: r.erpsStatus,
        total: 0, ring: 0, linear: 0, colocated: 0, unmarked: 0
      };
      order.push(r.cluster);
    }
    const c = byCluster[r.cluster];
    const cat = topologyCategory_(r.aEnd); // r.aEnd = column E
    c.total++;
    if (cat === 'ring') {
      c.ring++;
    } else if (cat === 'colocated') {
      c.linear++;
      c.colocated++;
    } else if (cat === 'linear') {
      c.linear++;
    } else {
      c.unmarked++;
    }
    sites.push({ cluster: r.cluster, eNodeB: r.eNodeB, value: r.aEnd, category: cat });
  });

  const clusters = order.map(function (name) { return byCluster[name]; });
  return { clusters: sortByClusterNo_(clusters), sites: sortByClusterNo_(sites) };
}

/** Writes one formatted report sheet: title, credit line, header row, data rows. */
function writeReportSheet_(sheet, title, subtitle, headers, rows, widths, centerCols) {
  sheet.clear();
  sheet.getRange(1, 1).setValue(title).setFontWeight('bold').setFontSize(14).setFontColor('#1F3B57');
  sheet.getRange(2, 1).setValue(subtitle).setFontStyle('italic').setFontColor('#333333');

  sheet.getRange(4, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold').setFontColor('#111111').setBackground('#dce6f0')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true)
    .setBorder(true, true, true, true, true, true, '#1F3B57', SpreadsheetApp.BorderStyle.SOLID);

  if (rows.length) {
    sheet.getRange(5, 1, rows.length, headers.length)
      .setValues(rows)
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#9aa9b8', SpreadsheetApp.BorderStyle.SOLID);
    centerCols.forEach(function (c) {
      sheet.getRange(5, c, rows.length, 1).setHorizontalAlignment('center');
    });
  } else {
    sheet.getRange(5, 1).setValue('None').setFontStyle('italic').setFontColor('#777777');
  }

  widths.forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
  sheet.setRowHeight(4, 52);
  sheet.setFrozenRows(4);
}

/**
 * Builds the ERPS Excel report and saves it to the root of My Drive:
 *   Ring vs Linear     per cluster: total sites, sites in ring, sites linear
 *                      (including "Linear but co-located"), and separately how
 *                      many of those linear sites are co-located
 *   ERPS Completed     cluster names with ERPS status Completed
 *   ERPS Not Feasible  cluster names with ERPS status Not feasible
 *   Site Detail        the Ring/Linear value of every ENode B
 * Ring/Linear comes from the "Site is in Ring/Linear" dropdown (column E).
 * A temporary Google Sheet is created, exported as .xlsx, then trashed.
 */
function exportErpsReportExcel() {
  const ui = SpreadsheetApp.getUi();
  let tmpId = null;
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast('Building the report. This takes about 10-30 seconds…', 'ERPS Report', 30);
    Logger.log('1/5 Reading Work Data and counting Ring / Linear sites…');
    const data = analyzeRingLinear_();
    const clusters = data.clusters;
    clusters.forEach(function (c) { c.category = erpsCategory_(c.status); });
    Logger.log('2/5 Analysed ' + clusters.length + ' clusters and ' + data.sites.length + ' sites.');

    const completed = clusters.filter(function (c) { return c.category === 'completed'; });
    const notFeasible = clusters.filter(function (c) { return c.category === 'notfeasible'; });
    const sum = function (key) { return clusters.reduce(function (a, c) { return a + c[key]; }, 0); };

    const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
    const now = new Date();
    const stamp = Utilities.formatDate(now, tz, 'dd MMM yyyy, hh:mm a');
    const fileStamp = Utilities.formatDate(now, tz, 'yyyyMMdd_HHmm');
    const subtitle = 'Created by Sarath Ganga S, JE   |   Generated on ' + stamp;
    const heading = 'ALATHUR NW ERPS MIGRATION - ';

    const tmp = SpreadsheetApp.create('tmp_ERPS_Report_' + fileStamp);
    tmpId = tmp.getId();

    // 1) Ring vs Linear, per cluster (with a TOTAL row)
    const ringRows = clusters.map(function (c, i) {
      return [i + 1, c.cluster, c.total, c.ring, c.linear, c.colocated, c.unmarked];
    });
    ringRows.push(['', 'TOTAL', sum('total'), sum('ring'), sum('linear'), sum('colocated'), sum('unmarked')]);

    const ringSheet = tmp.getSheets()[0];
    ringSheet.setName('Ring vs Linear');
    writeReportSheet_(ringSheet, heading + 'Sites in Ring / Linear per Cluster', subtitle,
      ['Sl. No.', 'Cluster No.', 'Total No. of Sites', 'No. of Sites in Ring',
       'No. of Sites Linear (incl. Linear but Co-located)', 'Of which Linear but Co-located', 'Not Yet Marked'],
      ringRows, [60, 170, 110, 110, 200, 150, 100], [1, 3, 4, 5, 6, 7]);
    ringSheet.getRange(4 + ringRows.length, 1, 1, 7).setFontWeight('bold').setBackground('#eef2f7');
    ringSheet.getRange(6 + ringRows.length, 1)
      .setValue('Note: Linear includes "Linear but co-located". Total = Ring + Linear + Not Yet Marked.')
      .setFontStyle('italic').setFontColor('#555555');

    // 2) ERPS Completed
    writeReportSheet_(tmp.insertSheet('ERPS Completed'), heading + 'ERPS Completed Clusters', subtitle,
      ['Sl. No.', 'Cluster No.', 'ERPS Status', 'Total Sites'],
      completed.map(function (c, i) { return [i + 1, c.cluster, c.status, c.total]; }),
      [60, 170, 160, 100], [1, 3, 4]);

    // 3) ERPS Not Feasible
    writeReportSheet_(tmp.insertSheet('ERPS Not Feasible'), heading + 'ERPS Not Feasible Clusters', subtitle,
      ['Sl. No.', 'Cluster No.', 'ERPS Status', 'Total Sites'],
      notFeasible.map(function (c, i) { return [i + 1, c.cluster, c.status, c.total]; }),
      [60, 170, 160, 100], [1, 3, 4]);

    // 4) Site-level detail, so every count can be checked
    const catLabel = { ring: 'Ring', linear: 'Linear', colocated: 'Linear but co-located', unmarked: 'Not marked' };
    writeReportSheet_(tmp.insertSheet('Site Detail'), heading + 'Ring / Linear Status of Each Site', subtitle,
      ['Cluster No.', 'ENode B', 'Site is in Ring/Linear (as entered)', 'Counted As'],
      data.sites.map(function (s) { return [s.cluster, s.eNodeB, s.value || '', catLabel[s.category]]; }),
      [160, 230, 230, 180], [3, 4]);

    SpreadsheetApp.flush();
    Logger.log('3/5 Report sheets built. Exporting to Excel…');

    // Export the temporary sheet as .xlsx and save it to Drive
    const exportUrl = 'https://docs.google.com/spreadsheets/d/' + tmpId + '/export?format=xlsx';
    const resp = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('Excel export failed (HTTP ' + resp.getResponseCode() + ').');
    }
    Logger.log('4/5 Excel exported (' + resp.getBlob().getBytes().length + ' bytes). Saving to Drive…');
    const fileName = 'ERPS_Report_' + fileStamp + '.xlsx';
    const file = DriveApp.createFile(resp.getBlob().setName(fileName));
    Logger.log('5/5 Saved: ' + file.getUrl());

    const dl = 'https://drive.google.com/uc?export=download&amp;id=' + file.getId();
    const btn = 'display:inline-block;margin:4px;padding:10px 18px;border-radius:8px;font-weight:700;text-decoration:none;color:#111;';
    const html = '<div style="font-family:Arial,sans-serif;padding:16px;text-align:center;">' +
      '<h3 style="color:#1F3B57;margin:0 0 6px;">Excel report ready</h3>' +
      '<div style="font-size:12px;color:#555;margin-bottom:16px;">' + fileName +
      '<br>Saved in My Drive</div>' +
      '<a href="' + dl + '" target="_blank" style="' + btn + 'background:#43D17A;">⬇ Download Excel</a>' +
      '<a href="' + file.getUrl() + '" target="_blank" style="' + btn + 'background:#dce6f0;">Open in Drive</a>' +
      '</div>';
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(430).setHeight(190), 'ERPS Excel Report');
  } catch (e) {
    Logger.log('ERROR: ' + e.message + ' | ' + e.stack);
    ui.alert('Could not create the Excel report', e.message, ui.ButtonSet.OK);
  } finally {
    if (tmpId) {
      try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (x) { /* ignore */ }
    }
  }
}

// ---------------------------------------------------------------------------
// WEB APP ENTRY POINTS
// ---------------------------------------------------------------------------

function doGet(e) {
  // This in-Google page (and the print links) are for the sheet owner only.
  // Everyone else uses the installable field app, which talks to doPost below.
  if (!isOwnerRequest_()) {
    return ContentService.createTextOutput('ERPS Field App data service. Please open the field app from the link you were given.');
  }
  if (e && e.parameter && e.parameter.printRow) {
    return renderSinglePrintPage_(e.parameter.printRow);
  }
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Alathur NW ERPS — Field App')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=5, user-scalable=yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------------------------------------------------------------------------
// FIELD APP API (used by the installable app on GitHub Pages)
// ---------------------------------------------------------------------------
// The installable app sends one POST per request: {fn, args}. Anyone with the
// link can READ. Saving needs the ADMIN login token (the PIN), so only the
// owner can change data. Only the functions listed in apiDispatch_ can be
// reached this way.

/** A cheap "has anything changed?" stamp: when the spreadsheet file was last modified. */
function getDataVersion() {
  try {
    return String(DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getLastUpdated().getTime());
  } catch (e) {
    return String(Date.now());   // if Drive cannot be read, always look changed
  }
}

/** Power Levels sheet as a lookup by lower-case site name (first match wins, like getPowerData). */
function buildPowerMap_() {
  const map = {};
  const sheet = getPowerSheet_();
  if (!sheet) return map;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  values.forEach(function (row, i) {
    const key = clean_(row[0]).toLowerCase();
    if (!key || map[key]) return;
    map[key] = {
      rowIndex: i + 2, siteName: row[0], objectName: row[1], laserType: row[2],
      txPower: row[3], rxPower: row[4], fixStatus: row[5] || ''
    };
  });
  return map;
}

/** Sink Node Ports sheet as {lower-case node name: [ports sorted by port ID]}, read once. */
function getAllSinkPortsMap_() {
  const sheet = getSinkPortsSheet_();
  const lastRow = sheet.getLastRow();
  const map = {};
  if (lastRow < 2) return map;
  const values = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  values.forEach(function (r, i) {
    const key = clean_(portIdToText_(r[0])).toLowerCase();
    if (!key) return;
    if (!map[key]) map[key] = [];
    map[key].push({ rowIndex: i + 2, portId: portIdToText_(r[1]), connectedTo: portIdToText_(r[2]) });
  });
  Object.keys(map).forEach(function (k) {
    map[k].sort(function (a, b) { return comparePortIds_(a.portId, b.portId); });
  });
  return map;
}

/**
 * Everything the field app needs to work offline, in one reply: the cluster
 * and site lists, the full detail of every site, search index, dashboard,
 * power cases, status report, checklist and sink ports. Diagrams are fetched
 * separately (one at a time) because they are large.
 */
function getOfflineBundle() {
  const sheet = getSheet_();
  const rows = getDataRows_();

  const bundle = {
    generatedAt: new Date().toISOString(),
    version: getDataVersion(),
    clusters: [], sites: {}, siteData: {}, searchIndex: [],
    dashboard: null, powerCases: [], statusReport: [], checklist: [],
    sinkNodes: [], sinkPorts: {}, diagramRows: []
  };

  rows.forEach(function (r) {
    if (r.cluster && bundle.clusters.indexOf(r.cluster) === -1) bundle.clusters.push(r.cluster);
  });
  bundle.clusters.forEach(function (c) { bundle.sites[c] = []; });
  rows.forEach(function (r) {
    if (bundle.sites[r.cluster]) {
      bundle.sites[r.cluster].push({
        rowIndex: r.rowIndex,
        label: r.eNodeB || r.aEnd || ('Row ' + r.rowIndex),
        aEnd: r.aEnd, eNodeB: r.eNodeB, bEnd: r.bEnd
      });
    }
    bundle.searchIndex.push({
      rowIndex: r.rowIndex, cluster: r.cluster,
      label: r.eNodeB || r.aEnd || ('Row ' + r.rowIndex)
    });
  });

  if (rows.length) {
    const first = rows[0].rowIndex;
    const numRows = rows[rows.length - 1].rowIndex - first + 1;
    const values = sheet.getRange(first, 1, numRows, LAST_COL).getValues();
    const powerMap = buildPowerMap_();
    rows.forEach(function (r) {
      bundle.siteData[r.rowIndex] = buildSiteData_(r.rowIndex, values[r.rowIndex - first], r, powerMap);
    });
  }

  const seenDiagram = {};
  Object.keys(bundle.siteData).forEach(function (k) {
    const d = bundle.siteData[k].header.diagramRow;
    if (d && !seenDiagram[d]) { seenDiagram[d] = true; bundle.diagramRows.push(d); }
  });

  bundle.dashboard = getDashboardData();
  bundle.powerCases = getAllPowerCases();
  bundle.statusReport = getStatusReportData();
  bundle.checklist = getChecklistData();

  bundle.sinkNodes = getSinkNodes();
  const portsMap = getAllSinkPortsMap_();
  bundle.sinkNodes.forEach(function (n) {
    bundle.sinkPorts[n] = portsMap[clean_(n).toLowerCase()] || [];
  });
  return bundle;
}

/** The only door into the data for the installable app. */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = { ok: true, result: apiDispatch_(req) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function apiDispatch_(req) {
  const fn = String(req.fn || '');
  const args = Array.isArray(req.args) ? req.args : [];

  switch (fn) {
    // admin login (protected by the PIN and its lock-out)
    case 'adminLogin':  return adminLogin.apply(null, args);
    case 'checkAdmin':  return checkAdmin.apply(null, args);
    case 'adminLogout': return adminLogout.apply(null, args);

    // reading: open to anyone with the link
    case 'getDataVersion':    return getDataVersion();
    case 'getOfflineBundle':  return getOfflineBundle();
    case 'getClusterDiagram': return getClusterDiagram.apply(null, args);
    case 'getSinkNodePorts':  return getSinkNodePorts.apply(null, args);

    // saving: each of these checks the admin token itself
    case 'saveSiteData':      return saveSiteData.apply(null, args);
    case 'savePowerStatus':   return savePowerStatus.apply(null, args);
    case 'saveSinkNodePorts': return saveSinkNodePorts.apply(null, args);
    case 'addSinkPort':       return addSinkPort.apply(null, args);
    case 'deleteSinkPort':    return deleteSinkPort.apply(null, args);

    default: throw new Error('Unknown request');
  }
}

// ---------------------------------------------------------------------------
// SHARED HELPERS
// ---------------------------------------------------------------------------

function getSheet_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found. Check the tab name.');
  return sheet;
}

function formatDate_(v) {
  if (!v && v !== 0) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || 'Asia/Kolkata', 'dd/MM/yyyy');
  }
  return v.toString();
}

function clean_(v) {
  return (v || '').toString().replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '').trim();
}

/**
 * Sorts report rows in ascending order of cluster number (the number at the
 * end of the name, e.g. "PKD_Cluster-26" -> 26). Sites inside a cluster keep
 * their original sheet order. Clusters without a number go to the end.
 */
function sortByClusterNo_(sites) {
  function num(name) {
    const m = (name || '').toString().match(/(\d+)\s*$/);
    return m ? Number(m[1]) : Number.MAX_SAFE_INTEGER;
  }
  return sites
    .map(function (s, i) { return { s: s, i: i }; })
    .sort(function (a, b) {
      const na = num(a.s.cluster), nb = num(b.s.cluster);
      if (na !== nb) return na - nb;
      if (a.s.cluster !== b.s.cluster) return a.s.cluster < b.s.cluster ? -1 : 1;
      return a.i - b.i;
    })
    .map(function (x) { return x.s; });
}

/**
 * Reads all data rows once, forward-filling Cluster No./Source Node/Source IP
 * across the blank rows caused by visually-merged cells, and stops at the
 * first fully-blank row (end of real data).
 */
function getDataRows_() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - FIRST_DATA_ROW + 1;
  if (numRows <= 0) return [];

  const values = sheet.getRange(FIRST_DATA_ROW, 1, numRows, LAST_COL).getValues();
  const rows = [];
  let currentCluster = '', currentSourceNode = '', currentSourceIP = '', currentDiagramRow = null;
  let currentErpsStatus = '', currentVrrpStatus = '';

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowIndex = FIRST_DATA_ROW + i;

    const aEnd = clean_(row[COL.A_END - 1]);
    const eNodeB = clean_(row[COL.ENODE_B - 1]);
    const bEnd = clean_(row[COL.B_END - 1]);
    const sourceIP = clean_(row[COL.SOURCE_IP - 1]);

    if (!aEnd && !eNodeB && !bEnd && !sourceIP) break; // end of real data

    const clusterVal = clean_(row[COL.CLUSTER - 1]);
    const sourceNodeVal = clean_(row[COL.SOURCE_NODE - 1]);

    if (clusterVal) {
      currentCluster = clusterVal;
      currentDiagramRow = rowIndex; // the diagram image lives on this same first row of the cluster
      // ERPS/VRRP Status are also merged per-cluster, entered only on this same first row
      currentErpsStatus = clean_(row[COL.ERPS_STATUS - 1]);
      currentVrrpStatus = clean_(row[COL.VRRP_STATUS - 1]);
    }
    if (sourceNodeVal) currentSourceNode = sourceNodeVal;
    if (sourceIP) currentSourceIP = sourceIP;

    rows.push({
      rowIndex: rowIndex,
      cluster: currentCluster,
      sourceNode: currentSourceNode,
      sourceIP: sourceIP || currentSourceIP,
      aEnd: aEnd,
      eNodeB: eNodeB,
      bEnd: bEnd,
      diagramRow: currentDiagramRow,
      erpsStatus: currentErpsStatus,
      vrrpStatus: currentVrrpStatus
    });
  }
  return rows;
}

/** Returns unique cluster numbers in sheet order, e.g. ["PKD_Cluster-26", "PKD_Cluster-27"] */
function getClusters() {
  const rows = getDataRows_();
  const seen = [];
  rows.forEach(function (r) {
    if (r.cluster && seen.indexOf(r.cluster) === -1) seen.push(r.cluster);
  });
  return seen;
}

/**
 * Returns a flat list of every site (cluster + name + row) for the
 * search box, so a site can be found by name without first picking
 * its cluster.
 */
function getSiteSearchIndex() {
  return getDataRows_().map(function (r) {
    return {
      rowIndex: r.rowIndex,
      cluster: r.cluster,
      label: r.eNodeB || r.aEnd || ('Row ' + r.rowIndex)
    };
  });
}

/** Returns the site rows mapped under one cluster, for the second dropdown */
function getSites(clusterNo) {
  return getDataRows_()
    .filter(function (r) { return r.cluster === clusterNo; })
    .map(function (r) {
      return {
        rowIndex: r.rowIndex,
        label: r.eNodeB || r.aEnd || ('Row ' + r.rowIndex),
        aEnd: r.aEnd,
        eNodeB: r.eNodeB,
        bEnd: r.bEnd
      };
    });
}

// Columns typed as bare "n/n/n" port labels (e.g. "2/24/26") get silently
// auto-parsed by Sheets as an actual date the moment they're entered — the
// same trap we hit with Sink Node Port IDs. Genuine dates (Request/Done
// Date) should still display as dd/MM/yyyy, but these need reconstructing
// in the M/d/yy order they were actually typed in, or the recovered text
// comes out completely different from what the field tech entered.
const PORT_LABEL_COLUMNS_ = [];
function initPortLabelColumns_() {
  if (PORT_LABEL_COLUMNS_.length) return;
  PORT_LABEL_COLUMNS_.push(COL.BSC_EXIST_PORT, COL.BSC_NEW_PORT, COL.UPLINK_BSC_PORT, COL.RNC_PORT);
}

/** Returns every tab's data for one site (one sheet row) */
function getSiteData(rowIndex) {
  rowIndex = Number(rowIndex);
  const sheet = getSheet_();
  const values = sheet.getRange(rowIndex, 1, 1, LAST_COL).getValues()[0];
  const ctx = getDataRows_().filter(function (r) { return r.rowIndex === rowIndex; })[0] || {};
  return buildSiteData_(rowIndex, values, ctx, null);
}

/**
 * Builds the site object from one row of values. Shared by getSiteData (one
 * site) and getOfflineBundle (all sites at once, using a prebuilt powerMap).
 */
function buildSiteData_(rowIndex, values, ctx, powerMap) {
  initPortLabelColumns_();
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const g = function (c) {
    const v = values[c - 1];
    // Sheets silently auto-parses things like "2/24/16" as a real Date.
    // Convert it back to plain text so it never leaks a raw Date object
    // through google.script.run (which can hang on nested Date values).
    if (Object.prototype.toString.call(v) === '[object Date]') {
      if (PORT_LABEL_COLUMNS_.indexOf(c) !== -1) {
        return Utilities.formatDate(v, tz, 'M/d/yy'); // matches how it was actually typed
      }
      return formatDate_(v); // genuine dates: dd/MM/yyyy
    }
    return v;
  };

  return {
    rowIndex: rowIndex,
    header: {
      cluster: ctx.cluster || '',
      sourceNode: ctx.sourceNode || '',
      sourceIP: ctx.sourceIP || '',
      aEnd: g(COL.A_END),
      eNodeB: g(COL.ENODE_B),
      bEnd: g(COL.B_END),
      diagramRow: ctx.diagramRow || rowIndex,
      ringLinear: g(COL.A_END),           // column E: Ring / Linear / Linear but co-located
      erpsStatus: ctx.erpsStatus || '',   // column AS (cluster level)
      vrrpStatus: ctx.vrrpStatus || ''    // column AT (cluster level)
    },
    fourG: {
      vlan1: g(COL.VLAN4G_1), vlan2: g(COL.VLAN4G_2), vlan3: g(COL.VLAN4G_3), vlan4: g(COL.VLAN4G_4),
      iwfMac: g(COL.IWF_MAC),
      cpuPortMgmtVlan: g(COL.CPU_PORT_MGMT_VLAN), cpuPortIp: g(COL.CPU_PORT_IP), nextHopIpNetmask: g(COL.NEXT_HOP_IP_NETMASK)
    },
    threeG: {
      vlan3g: g(COL.VLAN_3G), rnc: g(COL.RNC), rncPort: g(COL.RNC_PORT)
    },
    twoG: {
      bsc: g(COL.BSC), uplinkBscPort: g(COL.UPLINK_BSC_PORT),
      bscExistingPort: g(COL.BSC_EXIST_PORT), bscNewPort: g(COL.BSC_NEW_PORT),
      vlan2g: g(COL.VLAN_2G), e1: g(COL.E1), cefMac: g(COL.CEF_MAC), iwfMacC1: g(COL.IWF_MAC_C1)
    },
    sink: {
      sinkNode: g(COL.SINK_NODE), sinkPort: g(COL.SINK_PORT), sinkIp: g(COL.SINK_IP), sinkNe: g(COL.SINK_NE)
    },
    service: {
      requestDate: formatDate_(g(COL.REQUEST_DATE)), doneDate: formatDate_(g(COL.DONE_DATE))
    },
    cef: {
      p1: g(COL.P1), p2: g(COL.P2), p3: g(COL.P3), p4: g(COL.P4), p5: g(COL.P5),
      e1_12: g(COL.E1_12), e2_34: g(COL.E2_34)
    },
    status: {
      cefConfigStatus: g(COL.CEF_CONFIG_STATUS),
      migration2g: g(COL.MIGRATION_2G),
      migration3g: g(COL.MIGRATION_3G)
    },
    power: powerMap ? (powerMap[clean_(g(COL.ENODE_B)).toLowerCase()] || null) : getPowerData(g(COL.ENODE_B))
  };
}

// ---------------------------------------------------------------------------
// CLUSTER DIAGRAM (column H, JPEG images)
// ---------------------------------------------------------------------------

const DIAGRAM_COL = 8; // column H

/**
 * Returns the Cluster Diagram image from column H as a data URI, so the
 * web app / report can display it without needing separate auth/hosting.
 * Supports both "Insert image > In cell" (CellImage) and legacy floating
 * images anchored over the cell.
 */
function getClusterDiagram(rowIndex) {
  rowIndex = Number(rowIndex);
  const sheet = getSheet_();

  // 1) Image inserted "in cell"
  const value = sheet.getRange(rowIndex, DIAGRAM_COL).getValue();
  if (value && typeof value.getContentUrl === 'function') {
    const dataUri = fetchImageAsDataUri_(value.getContentUrl());
    if (dataUri) return dataUri;
  }

  // 2) Fallback: floating image anchored over this cell
  const images = sheet.getImages();
  for (let i = 0; i < images.length; i++) {
    const img = images[i];
    const anchor = img.getAnchorCell();
    if (anchor.getRow() === rowIndex && anchor.getColumn() === DIAGRAM_COL) {
      try {
        return blobToDataUri_(img.getBlob());
      } catch (e) { /* try next image */ }
    }
  }

  return null;
}

function blobToDataUri_(blob) {
  const base64 = Utilities.base64Encode(blob.getBytes());
  const contentType = blob.getContentType() || 'image/jpeg';
  return 'data:' + contentType + ';base64,' + base64;
}

function fetchImageAsDataUri_(url) {
  if (!url) return null;
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return null;
    return blobToDataUri_(resp.getBlob());
  } catch (e) {
    return null;
  }
}

// ---------------------------------------------------------------------------
// POWER LEVELS
// ---------------------------------------------------------------------------

const POWER_SHEET_NAME = 'Power Levels'; // display label only — lookup below is resilient to renaming

/** Finds the power-levels sheet regardless of whether it's named
 * "Power Levels", "Power_Levels", different case, etc. — so a future
 * rename can't silently break this feature again. */
function getPowerSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exact = ss.getSheetByName(POWER_SHEET_NAME) || ss.getSheetByName('Power_Levels');
  if (exact) return exact;
  const target = 'powerlevels';
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const norm = sheets[i].getName().toLowerCase().replace(/[\s_]/g, '');
    if (norm === target) return sheets[i];
  }
  return null;
}
// Power Levels tab layout: A=Site Name, B=Object Name, C=Laser Type, D=Tx Power(dBm), E=Rx Power(dBm), F=Power Fix Status

/**
 * Returns every case currently listed in the Power Levels sheet, each
 * cross-referenced against Work Data so the app can jump straight to
 * that site (cluster + row) when tapped.
 */
function getAllPowerCases() {
  const sheet = getPowerSheet_();
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  const workRows = getDataRows_();
  const lookup = {};
  workRows.forEach(function (r) {
    const key = clean_(r.eNodeB).toLowerCase();
    if (key) lookup[key] = { cluster: r.cluster, rowIndex: r.rowIndex };
  });

  return values
    .map(function (row, i) {
      const siteName = clean_(portIdToText_(row[0]));
      if (!siteName) return null;
      const match = lookup[siteName.toLowerCase()];
      return {
        rowIndex: i + 2,
        siteName: siteName,
        objectName: row[1],
        laserType: row[2],
        txPower: row[3],
        rxPower: row[4],
        fixStatus: (row[5] || '').toString().trim(),
        cluster: match ? match.cluster : '',
        workRowIndex: match ? match.rowIndex : null
      };
    })
    .filter(function (p) { return p !== null; });
}

/**
 * Returns the power-issue row for a given site name (case-insensitive match
 * against ENode B), or null if that site has no flagged power issue.
 */
function getPowerData(siteName) {
  const sheet = getPowerSheet_();
  if (!sheet) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  const target = clean_(siteName).toLowerCase();

  for (let i = 0; i < values.length; i++) {
    const rowSite = clean_(values[i][0]).toLowerCase();
    if (rowSite && rowSite === target) {
      return {
        rowIndex: i + 2,
        siteName: values[i][0],
        objectName: values[i][1],
        laserType: values[i][2],
        txPower: values[i][3],
        rxPower: values[i][4],
        fixStatus: values[i][5] || ''
      };
    }
  }
  return null;
}

/** Writes back just the Power Fix Status for one Power Levels row. */
function savePowerStatus(rowIndex, status, token) {
  requireAdmin_(token);
  rowIndex = Number(rowIndex);
  const sheet = getPowerSheet_();
  if (!sheet) return { success: false, error: 'Power Levels sheet not found' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.getRange(rowIndex, 6).setValue(status);
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// STATUS REPORT DATA, CHECKLIST, DASHBOARD
// ---------------------------------------------------------------------------

/**
 * Returns Cluster No. / ENode B / CEF Config / 2G Migration / 3G Migration
 * (plus ERPS/VRRP status and the cluster's diagram row) for every site —
 * the compact status listing used for the report submitted to the manager.
 * Grouped in natural sheet order (cluster grouping already resolved via
 * forward-fill).
 */
function getStatusReportData() {
  const sheet = getSheet_();
  const rows = getDataRows_();
  if (rows.length === 0) return [];

  const firstRow = rows[0].rowIndex;
  const lastRow = rows[rows.length - 1].rowIndex;
  const numRows = lastRow - firstRow + 1;
  const statusValues = sheet.getRange(firstRow, COL.CEF_CONFIG_STATUS, numRows, 3).getValues();

  return rows.map(function (r) {
    const idx = r.rowIndex - firstRow;
    return {
      cluster: r.cluster,
      enodeB: r.eNodeB || r.aEnd,
      cefConfigStatus: (statusValues[idx][0] || '').toString().trim(),
      migration2g: (statusValues[idx][1] || '').toString().trim(),
      migration3g: (statusValues[idx][2] || '').toString().trim(),
      erpsStatus: r.erpsStatus,
      vrrpStatus: r.vrrpStatus,
      diagramRow: r.diagramRow
    };
  });
}

/**
 * Returns the exact field set used on the printed field checklist, for
 * every site currently in Work Data, in sheet order with cluster grouping
 * already resolved (forward-filled past merged cells).
 */
function getChecklistData() {
  initPortLabelColumns_();
  const sheet = getSheet_();
  const rows = getDataRows_();
  if (rows.length === 0) return [];

  const firstRow = rows[0].rowIndex;
  const lastRow = rows[rows.length - 1].rowIndex;
  const numRows = lastRow - firstRow + 1;
  const values = sheet.getRange(firstRow, 1, numRows, LAST_COL).getValues();
  const tz = Session.getScriptTimeZone() || 'Asia/Kolkata';

  function cellText(row, col) {
    const v = row[col - 1];
    if (Object.prototype.toString.call(v) === '[object Date]') {
      if (PORT_LABEL_COLUMNS_.indexOf(col) !== -1) return Utilities.formatDate(v, tz, 'M/d/yy');
      return formatDate_(v);
    }
    return v;
  }

  return rows.map(function (r) {
    const row = values[r.rowIndex - firstRow];
    const vlan4gParts = [COL.VLAN4G_1, COL.VLAN4G_2, COL.VLAN4G_3, COL.VLAN4G_4]
      .map(function (c) { return (cellText(row, c) || '').toString().trim(); })
      .filter(Boolean);
    return {
      rowIndex: r.rowIndex,
      cluster: r.cluster,
      enodeB: r.eNodeB || r.aEnd,
      vlan4g: vlan4gParts.join(' / '),
      iwfMac: cellText(row, COL.IWF_MAC),
      vlan3g: cellText(row, COL.VLAN_3G),
      rnc: cellText(row, COL.RNC),
      bsc: cellText(row, COL.BSC),
      bscExistingPort: cellText(row, COL.BSC_EXIST_PORT),
      bscNewPort: cellText(row, COL.BSC_NEW_PORT),
      vlan2g: cellText(row, COL.VLAN_2G),
      sinkPort: cellText(row, COL.SINK_PORT)
    };
  });
}

/** Returns the checklist fields for just one site row (used by the
 * per-row print link in the sheet). */
function getChecklistDataForRow(rowIndex) {
  rowIndex = Number(rowIndex);
  const all = getChecklistData();
  for (let i = 0; i < all.length; i++) {
    if (all[i].rowIndex === rowIndex) return all[i];
  }
  return null;
}

/** Builds a standalone, print-ready HTML page for a single site — this is
 * what opens when someone clicks the 🖨 Print link in the sheet itself. */
function renderSinglePrintPage_(rowIndex) {
  const site = getChecklistDataForRow(rowIndex);
  if (!site) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;">No site found for this row. It may be a header row or a row without an ENode B.</p>');
  }

  function esc(s) {
    return (s === undefined || s === null) ? '' : s.toString()
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  function row(label, value) {
    const v = (value || '').toString().trim();
    return '<div class="r"><div class="l">' + esc(label) + '</div><div class="v">' + (v ? esc(v) : '—') + '</div></div>';
  }

  const html = '<!DOCTYPE html><html><head><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>' +
    'body{font-family:Arial,sans-serif;margin:0;padding:24px;background:#fff;}' +
    '.title{text-align:center;margin-bottom:20px;}' +
    '.title h1{font-size:18px;color:#1F3B57;margin:0 0 2px;}' +
    '.title .sub{font-size:12px;color:#555;font-weight:600;}' +
    '.card{border:1px solid #999;border-radius:6px;overflow:hidden;max-width:500px;margin:0 auto;}' +
    '.hdr{background:#1F3B57;color:#fff;padding:10px 14px;font-weight:700;font-size:15px;}' +
    '.hdr .c{font-weight:400;font-size:11px;color:#cbd8e6;margin-left:6px;}' +
    '.r{display:flex;border-top:1px solid #ccc;font-size:13px;}' +
    '.l{flex:0 0 42%;background:#f5f7fa;color:#333;font-weight:600;padding:8px 12px;}' +
    '.v{flex:1;color:#111;font-weight:600;padding:8px 12px;background:#fff;}' +
    '.checks{display:flex;gap:16px;background:#f2f2f2;padding:12px 14px;border-top:1px solid #ccc;font-size:13px;font-weight:600;flex-wrap:wrap;}' +
    '.checks span{display:inline-flex;align-items:center;gap:6px;}' +
    '.box{width:13px;height:13px;border:1.5px solid #333;display:inline-block;}' +
    '.printBtn{display:block;width:max-content;margin:18px auto 0;padding:10px 20px;background:#43D17A;border:none;border-radius:8px;font-weight:700;font-size:14px;cursor:pointer;}' +
    '@media print{ .no-print{display:none !important;} }' +
    '</style></head><body>' +
    '<div class="title"><h1>ALATHUR NW ERPS MIGRATION</h1><div class="sub">Field Verification Checklist</div></div>' +
    '<div class="card">' +
    '<div class="hdr">' + esc(site.enodeB) + '<span class="c">(' + esc(site.cluster) + ')</span></div>' +
    row('4G VLAN', site.vlan4g) + row('IWF MAC', site.iwfMac) + row('3G VLAN', site.vlan3g) +
    row('RNC', site.rnc) + row('BSC', site.bsc) + row('BSC Existing Port', site.bscExistingPort) +
    row('BSC New Port', site.bscNewPort) + row('2G VLAN', site.vlan2g) + row('Sink Port', site.sinkPort) +
    '<div class="checks"><span><span class="box"></span>CEF Migration</span>' +
    '<span><span class="box"></span>2G Migration</span><span><span class="box"></span>3G Migration</span></div>' +
    '</div>' +
    '<button class="printBtn no-print" onclick="window.print()">🖨 Print / Save PDF</button>' +
    '<script>window.onload = function(){ setTimeout(function(){ window.print(); }, 300); };</script>' +
    '</body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('Print — ' + site.enodeB);
}

/**
 * Aggregates every site's status into overall totals, a per-cluster
 * breakdown, and a flat list of anything not yet fully complete —
 * for the dashboard view.
 */
function getDashboardData() {
  const sheet = getSheet_();
  const rows = getDataRows_();
  if (rows.length === 0) {
    return {
      totals: { sites: 0, cef: 0, cefApplicable: 0, mig2g: 0, mig2gApplicable: 0, mig3g: 0, mig3gApplicable: 0 },
      clusters: [], pending: []
    };
  }

  const firstRow = rows[0].rowIndex;
  const lastRow = rows[rows.length - 1].rowIndex;
  const numRows = lastRow - firstRow + 1;

  const statusValues = sheet.getRange(firstRow, COL.CEF_CONFIG_STATUS, numRows, 3).getValues();
  // statusValues[i] = [cefConfig, migration2g, migration3g] for row (firstRow + i)

  // Classifies a status as 'done', 'na' (facility doesn't exist there —
  // e.g. "No 2G", "4G Only site" — never counted as pending), or
  // 'pending' (blank, or any other in-progress/error text).
  function classify(text, techKeyword) {
    const v = (text || '').toString().trim();
    const lower = v.toLowerCase();
    if (lower === 'completed') return 'done';
    if (lower.indexOf('4g only') !== -1) return 'na';
    if (techKeyword && lower === 'no ' + techKeyword) return 'na';
    return 'pending';
  }

  const clusterMap = {}; // clusterName -> { sites, cef, cefApplicable, mig2g, mig2gApplicable, mig3g, mig3gApplicable }
  const clusterOrder = [];
  const pending = [];

  rows.forEach(function (r) {
    const idx = r.rowIndex - firstRow;
    const cefVal = statusValues[idx][0];
    const mig2gVal = statusValues[idx][1];
    const mig3gVal = statusValues[idx][2];

    const cefClass = classify(cefVal, null);
    const mig2gClass = classify(mig2gVal, '2g');
    const mig3gClass = classify(mig3gVal, '3g');

    if (!clusterMap[r.cluster]) {
      clusterMap[r.cluster] = {
        cluster: r.cluster, sites: 0,
        cef: 0, cefApplicable: 0, mig2g: 0, mig2gApplicable: 0, mig3g: 0, mig3gApplicable: 0
      };
      clusterOrder.push(r.cluster);
    }
    const c = clusterMap[r.cluster];
    c.sites += 1;
    if (cefClass !== 'na') { c.cefApplicable += 1; if (cefClass === 'done') c.cef += 1; }
    if (mig2gClass !== 'na') { c.mig2gApplicable += 1; if (mig2gClass === 'done') c.mig2g += 1; }
    if (mig3gClass !== 'na') { c.mig3gApplicable += 1; if (mig3gClass === 'done') c.mig3g += 1; }

    if (cefClass === 'pending' || mig2gClass === 'pending' || mig3gClass === 'pending') {
      pending.push({
        rowIndex: r.rowIndex,
        cluster: r.cluster,
        site: r.eNodeB || r.aEnd,
        cefStatus: (cefVal || '').toString().trim() || 'Not started',
        mig2gStatus: (mig2gVal || '').toString().trim() || 'Not started',
        mig3gStatus: (mig3gVal || '').toString().trim() || 'Not started',
        cefDone: cefClass !== 'pending', mig2gDone: mig2gClass !== 'pending', mig3gDone: mig3gClass !== 'pending'
      });
    }
  });

  const clusters = clusterOrder.map(function (name) { return clusterMap[name]; });

  const totals = clusters.reduce(function (acc, c) {
    acc.sites += c.sites;
    acc.cef += c.cef; acc.cefApplicable += c.cefApplicable;
    acc.mig2g += c.mig2g; acc.mig2gApplicable += c.mig2gApplicable;
    acc.mig3g += c.mig3g; acc.mig3gApplicable += c.mig3gApplicable;
    return acc;
  }, { sites: 0, cef: 0, cefApplicable: 0, mig2g: 0, mig2gApplicable: 0, mig3g: 0, mig3gApplicable: 0 });

  return { totals: totals, clusters: clusters, pending: pending };
}

// ---------------------------------------------------------------------------
// SAVE
// ---------------------------------------------------------------------------

/**
 * Writes back only the editable fields: Sink Port, Done Date, and the
 * CEF Card Port Details block (P1-P5, E1(1-2), E2(3-4)), plus the three
 * status fields.
 * payload keys: any of EDITABLE_FIELDS' keys.
 */
function saveSiteData(rowIndex, payload, token) {
  requireAdmin_(token);
  rowIndex = Number(rowIndex);
  const sheet = getSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    Object.keys(EDITABLE_FIELDS).forEach(function (key) {
      if (payload[key] !== undefined) {
        const col = EDITABLE_FIELDS[key];
        let value = payload[key];
        if (key === 'doneDate' && value) {
          // <input type="date"> sends yyyy-mm-dd
          const parts = value.split('-');
          if (parts.length === 3) {
            value = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
          }
        }
        sheet.getRange(rowIndex, col).setValue(value);
      }
    });
    SpreadsheetApp.flush();
    return { success: true, savedAt: new Date().toISOString() };
  } catch (e) {
    return { success: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}
