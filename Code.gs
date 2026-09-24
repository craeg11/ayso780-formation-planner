const SHEETS = {
  teams: 'Teams',
  roster: 'Team Rosters',
  plans: 'Lineup Plans',
  sync: 'Sync Receipts'
};

const TEAM_HEADERS = ['Team ID', 'Team name', 'Coach', 'PIN hash', 'Age group', 'Gender', 'Season', 'Year'];
const ROSTER_HEADERS = ['Team ID', 'Display name', 'Jersey #'];
const PLAN_HEADERS = ['Plan ID', 'Saved at', 'Team ID', 'Plan name', 'Date', 'Formation', 'Plan data'];
const SYNC_HEADERS = ['Operation ID', 'Saved at', 'Status', 'Result'];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Lineup Planner')
    .addItem('Set up planner sheets', 'setupPlanner')
    .addItem('Open planner', 'showPlannerSidebar')
    .addToUi();
}

function showPlannerSidebar() {
  setupPlanner();
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('index').setTitle('AYSO Lineup Planner')
  );
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  try {
    let result;
    if (params.action === 'ping') result = { ok: true, version: 'pwa-sync-v2' };
    else if (params.action === 'getAppData') result = getAppData();
    else if (params.action === 'getTeamData') result = getTeamData(params.teamId, params.pin);
    else if (params.action === 'getSyncStatus') result = getSyncStatus_(params.operationId);
    else throw new Error('Unknown request.');
    return jsonOutput_(result, params.prefix || params.callback);
  } catch (err) {
    return jsonOutput_({ error: err.message || String(err) }, params.prefix || params.callback);
  }
}

function doPost(e) {
  let params = {};
  let operationId = '';
  try {
    params = JSON.parse((e.postData && e.postData.contents) || '{}');
    operationId = String(params.operationId || '').trim();
    if (operationId) {
      const prior = getSyncStatus_(operationId);
      if (prior.status === 'success') {
        return jsonOutput_(prior.result);
      }
    }
    const result = dispatchAction_(params);
    if (operationId) recordSyncStatus_(operationId, 'success', result);
    return jsonOutput_(result);
  } catch (err) {
    const result = { error: err.message || String(err) };
    if (operationId) recordSyncStatus_(operationId, 'error', result);
    return jsonOutput_(result);
  }
}

function dispatchAction_(params) {
  switch (params.action) {
    case 'getAppData': return getAppData();
    case 'getTeamData': return getTeamData(params.teamId, params.pin);
    case 'saveTeam': return saveTeam(params.team);
    case 'deleteTeam': return deleteTeam(params.teamId, params.pin);
    case 'saveRoster': return saveRoster(params.teamId, params.pin, params.roster);
    case 'savePlan': return savePlan(params.pin, params.plan);
    case 'deletePlan': return deletePlan(params.teamId, params.planId, params.pin);
    default: throw new Error('Unknown action.');
  }
}

function jsonOutput_(data, callback) {
  const json = JSON.stringify(data);
  const safeCallback = String(callback || '');
  if (safeCallback && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(safeCallback)) {
    return ContentService.createTextOutput(`${safeCallback}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function setupPlanner() {
  const ss = spreadsheet_();
  ensureSheet_(ss, SHEETS.teams, TEAM_HEADERS);
  ensureSheet_(ss, SHEETS.roster, ROSTER_HEADERS);
  ensureSheet_(ss, SHEETS.plans, PLAN_HEADERS);
  ensureSheet_(ss, SHEETS.sync, SYNC_HEADERS);
}

function getAppData() {
  setupPlanner();
  return { teams: readTeams_(getSheet_(SHEETS.teams)) };
}

function getTeamData(teamId, pin) {
  setupPlanner();
  verifyTeamPin_(teamId, pin);
  return {
    roster: readRoster_(getSheet_(SHEETS.roster), teamId),
    plans: readPlans_(getSheet_(SHEETS.plans), teamId)
  };
}

function saveTeam(team) {
  setupPlanner();
  const clean = cleanTeam_(team || {});
  const sheet = getSheet_(SHEETS.teams);
  const id = String((team || {}).id || '').trim() || Utilities.getUuid();
  const data = sheet.getDataRange().getValues();
  const existing = data.findIndex((row, index) => index > 0 && row[0] === id);
  const duplicate = data.find((row, index) => index > 0 && row[1] === clean.name && row[0] !== id);
  if (duplicate) {
    // Two devices may independently create the same generated team while offline.
    // A matching PIN proves the coach intended the existing team, so reconcile IDs
    // instead of creating a duplicate or permanently blocking the local sync queue.
    try {
      verifyTeamPin_(duplicate[0], clean.pin);
    } catch (error) {
      throw new Error('A team with that generated name already exists. Use that team\'s PIN to open it.');
    }
    return {
      id: duplicate[0],
      name: duplicate[1],
      coach: duplicate[2],
      ageGroup: duplicate[4],
      gender: duplicate[5],
      season: duplicate[6],
      year: duplicate[7],
      reconciledFrom: id
    };
  }
  if (existing > 0) verifyTeamPin_(id, team.currentPin || clean.pin);

  const row = [id, clean.name, clean.coach, hashPin_(clean.pin), clean.ageGroup, clean.gender, clean.season, clean.year];
  if (existing > 0) sheet.getRange(existing + 1, 1, 1, TEAM_HEADERS.length).setValues([row]);
  else sheet.appendRow(row);
  return { id, name: clean.name, coach: clean.coach, ageGroup: clean.ageGroup, gender: clean.gender, season: clean.season, year: clean.year };
}

function deleteTeam(teamId, pin) {
  setupPlanner();
  const id = String(teamId || '').trim();
  const team = verifyTeamPin_(id, pin);
  rewriteWithoutTeam_(getSheet_(SHEETS.teams), TEAM_HEADERS.length, 0, id);
  rewriteWithoutTeam_(getSheet_(SHEETS.roster), ROSTER_HEADERS.length, 0, id);
  rewriteWithoutTeam_(getSheet_(SHEETS.plans), PLAN_HEADERS.length, 2, id);
  return { deleted: team.name };
}

function saveRoster(teamId, pin, roster) {
  setupPlanner();
  verifyTeamPin_(teamId, pin);
  const players = (roster || []).map(cleanPlayer_).filter(player => player.name);
  const labels = new Set();
  players.forEach(player => {
    if (labels.has(player.label)) throw new Error(`Duplicate player: ${player.label}`);
    labels.add(player.label);
  });
  const sheet = getSheet_(SHEETS.roster);
  const retained = sheet.getDataRange().getValues().slice(1).filter(row => row[0] !== teamId && row[1]);
  const rows = [...retained, ...players.map(player => [teamId, player.name, player.jersey])];
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), ROSTER_HEADERS.length).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, ROSTER_HEADERS.length).setValues(rows);
  return readRoster_(sheet, teamId);
}

function savePlan(pin, plan) {
  setupPlanner();
  const safe = validatePlan_(plan || {});
  verifyTeamPin_(safe.teamId, pin);
  const sheet = getSheet_(SHEETS.plans);
  const id = safe.id || Utilities.getUuid();
  const data = sheet.getDataRange().getValues();
  const existing = data.findIndex((row, index) => index > 0 && row[0] === id);
  if (existing > 0 && data[existing][2] !== safe.teamId) {
    throw new Error('This plan belongs to a different team.');
  }
  const row = [id, new Date(), safe.teamId, safe.name, safe.date, safe.formation, JSON.stringify({ ...safe, id })];
  if (existing > 0) sheet.getRange(existing + 1, 1, 1, PLAN_HEADERS.length).setValues([row]);
  else sheet.appendRow(row);
  return { ...safe, id };
}

function deletePlan(teamId, planId, pin) {
  setupPlanner();
  verifyTeamPin_(teamId, pin);
  const id = String(planId || '').trim();
  if (!id) throw new Error('Choose a saved plan to delete.');
  const sheet = getSheet_(SHEETS.plans);
  const row = sheet.getDataRange().getValues().slice(1).find(record => record[0] === id && record[2] === teamId);
  if (!row) throw new Error('That saved plan no longer exists.');
  const plan = JSON.parse(row[6]);
  rewriteWithoutPlan_(sheet, id);
  return { deleted: plan.name || 'Saved plan' };
}

function getSyncStatus_(operationId) {
  setupPlanner();
  const id = String(operationId || '').trim();
  if (!id) return { status: 'missing' };
  const row = getSheet_(SHEETS.sync).getDataRange().getValues().slice(1).reverse()
    .find(record => record[0] === id);
  if (!row) return { status: 'pending' };
  let result = {};
  try { result = JSON.parse(row[3] || '{}'); } catch (error) { result = { error: 'Unreadable sync receipt.' }; }
  return { status: row[2], result };
}

function recordSyncStatus_(operationId, status, result) {
  if (!operationId) return;
  setupPlanner();
  const sheet = getSheet_(SHEETS.sync);
  const data = sheet.getDataRange().getValues();
  const existing = data.findIndex((row, index) => index > 0 && row[0] === operationId);
  const row = [operationId, new Date(), status, JSON.stringify(result || {})];
  if (existing > 0) sheet.getRange(existing + 1, 1, 1, SYNC_HEADERS.length).setValues([row]);
  else sheet.appendRow(row);
}

function cleanTeam_(team) {
  const ageGroup = String(team.ageGroup || '').trim();
  const gender = String(team.gender || '').trim();
  const coach = String(team.coach || '').trim().slice(0, 80);
  const season = String(team.season || '').trim();
  const year = String(team.year || '').trim();
  const pin = String(team.pin || '').trim();
  if (!['U19', 'U16', 'U14'].includes(ageGroup)) throw new Error('Choose an age group.');
  if (!['Boys', 'Girls'].includes(gender)) throw new Error('Choose Boys or Girls.');
  if (!coach) throw new Error('Enter the coach name.');
  if (!['Fall', 'All Stars', 'Spring'].includes(season)) throw new Error('Choose a season.');
  if (!/^(20)\d{2}$/.test(year)) throw new Error('Enter a four-digit year.');
  if (!/^\d{4}$/.test(pin)) throw new Error('PIN must be a 4-digit number.');
  const coachCode = coach.replace(/[^A-Za-z0-9]+/g, '');
  if (!coachCode) throw new Error('Coach name needs at least one letter or number.');
  return { ageGroup, gender, coach, season, year, pin, name: `${ageGroup}${gender === 'Girls' ? 'G' : 'B'}-${coachCode}_${season}_${year}` };
}

function verifyTeamPin_(teamId, pin) {
  const id = String(teamId || '').trim();
  const providedPin = String(pin || '').trim();
  const row = getSheet_(SHEETS.teams).getDataRange().getValues().slice(1).find(record => record[0] === id);
  if (!row) throw new Error('Team not found.');
  const storedPin = String(row[3] || '').trim();
  const valid = /^\d{4}$/.test(storedPin) ? storedPin === providedPin : storedPin === hashPin_(providedPin);
  if (!valid) throw new Error('Incorrect team PIN.');
  return { id: row[0], name: row[1], coach: row[2] };
}

function hashPin_(pin) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pin));
  return bytes.map(value => (value + 256) % 256).map(value => value.toString(16).padStart(2, '0')).join('');
}

function spreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('PLANNER_SPREADSHEET_ID');
  const ss = storedId ? SpreadsheetApp.openById(storedId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open this project from its Google Sheet and run setupPlanner once before deploying as a web app.');
  if (!storedId) props.setProperty('PLANNER_SPREADSHEET_ID', ss.getId());
  return ss;
}

function getSheet_(name) {
  return spreadsheet_().getSheetByName(name);
}

function ensureSheet_(ss, name, headers) {
  const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (existing.join('|') !== headers.join('|')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rewriteWithoutTeam_(sheet, columnCount, teamIdColumn, teamId) {
  const retained = sheet.getDataRange().getValues().slice(1).filter(row => row[teamIdColumn] !== teamId);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), columnCount).clearContent();
  if (retained.length) sheet.getRange(2, 1, retained.length, columnCount).setValues(retained);
}

function rewriteWithoutPlan_(sheet, planId) {
  const retained = sheet.getDataRange().getValues().slice(1).filter(row => row[0] !== planId);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), PLAN_HEADERS.length).clearContent();
  if (retained.length) sheet.getRange(2, 1, retained.length, PLAN_HEADERS.length).setValues(retained);
}

function readTeams_(sheet) {
  return sheet.getDataRange().getValues().slice(1)
    .map(([id, name, coach, , ageGroup, gender, season, year]) => ({ id: String(id || ''), name: String(name || ''), coach: String(coach || ''), ageGroup: String(ageGroup || ''), gender: String(gender || ''), season: String(season || ''), year: String(year || '') }))
    .filter(team => team.id && team.name)
    .sort((a, b) => a.name.localeCompare(b.name));
}

function readRoster_(sheet, teamId) {
  return sheet.getDataRange().getValues().slice(1)
    .filter(row => row[0] === teamId)
    .map(([, name, jersey]) => cleanPlayer_({ name, jersey }))
    .filter(player => player.name);
}

function readPlans_(sheet, teamId) {
  return sheet.getDataRange().getValues().slice(1).reverse()
    .filter(row => row[2] === teamId)
    .map(row => { try { return JSON.parse(row[6]); } catch (error) { return null; } })
    .filter(Boolean).slice(0, 30);
}

function cleanPlayer_(player) {
  const name = String(player.name || '').trim().slice(0, 50);
  const jersey = String(player.jersey || '').trim().slice(0, 10);
  return { name, jersey, label: jersey ? `${name} ${jersey}` : name };
}

function validatePlan_(plan) {
  const teamId = String(plan.teamId || '').trim();
  const name = String(plan.name || '').trim().slice(0, 80) || 'Lineup plan';
  const formation = String(plan.formation || '').trim().slice(0, 30);
  const date = String(plan.date || '').trim().slice(0, 20);
  if (!teamId) throw new Error('Choose a team before saving.');
  if (!formation) throw new Error('Choose a formation before saving.');
  const quarters = plan.quarters || {};
  const result = { id: plan.id || '', teamId, name, date, formation, absent: uniqueStrings_(plan.absent), quarters: {} };
  ['Q1', 'Q2', 'Q3', 'Q4'].forEach(quarter => {
    const current = quarters[quarter] || {};
    const lineup = uniqueStrings_(current.lineup);
    const substitutes = uniqueStrings_(current.substitutes).slice(0, 6);
    if (lineup.some(player => substitutes.includes(player))) throw new Error(`${quarter}: a player cannot be both a starter and a substitute.`);
    if (result.absent.some(player => lineup.includes(player) || substitutes.includes(player))) throw new Error(`${quarter}: an unavailable player is selected.`);
    result.quarters[quarter] = { lineup, substitutes };
  });
  return result;
}

function uniqueStrings_(values) {
  return [...new Set((values || []).map(value => String(value || '').trim()).filter(Boolean))];
}
