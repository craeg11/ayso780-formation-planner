const SHEETS = {
  teams: 'Teams',
  roster: 'Team Rosters',
  plans: 'Lineup Plans'
};

const TEAM_HEADERS = ['Team ID', 'Team name', 'Coach', 'PIN'];
const ROSTER_HEADERS = ['Team ID', 'Display name', 'Jersey #'];
const PLAN_HEADERS = ['Plan ID', 'Saved at', 'Team ID', 'Plan name', 'Date', 'Formation', 'Plan data'];

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
  const action = e ? e.parameter.action : '';
  let result = {};
  
  if (action === 'getAppData') result = getAppData();
  if (action === 'getTeamData') result = getTeamData(e.parameter.teamId, e.parameter.pin);
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents || '{}');
    const action = params.action;
    let result = {};

    if (action === 'getAppData') result = getAppData();
    else if (action === 'getTeamData') result = getTeamData(params.teamId, params.pin);
    else if (action === 'saveTeam') result = saveTeam(params.team);
    else if (action === 'deleteTeam') result = deleteTeam(params.teamId, params.pin);
    else if (action === 'saveRoster') result = saveRoster(params.teamId, params.pin, params.roster);
    else if (action === 'savePlan') result = savePlan(params.pin, params.plan);
    else if (action === 'deletePlan') result = deletePlan(params.teamId, params.planId, params.pin);

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupPlanner() {
  const ss = spreadsheet_();
  ensureSheet_(ss, SHEETS.teams, TEAM_HEADERS);
  ensureSheet_(ss, SHEETS.roster, ROSTER_HEADERS);
  ensureSheet_(ss, SHEETS.plans, PLAN_HEADERS);
}

function getAppData() {
  setupPlanner();
  return {
    teams: readTeams_(getSheet_(SHEETS.teams))
  };
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
  const name = String((team || {}).name || '').trim().slice(0, 80);
  const coach = String((team || {}).coach || '').trim().slice(0, 80);
  const pin = String((team || {}).pin || '').trim();
  
  if (!name) throw new Error('Enter a team name.');
  if (!/^\d{4}$/.test(pin)) throw new Error('PIN must be a 4-digit number.');

  const sheet = getSheet_(SHEETS.teams);
  const id = String((team || {}).id || '').trim() || Utilities.getUuid();
  const data = sheet.getDataRange().getValues();
  const i = data.findIndex((r, n) => n > 0 && r[0] === id);
  
  if (i > 0) {
    verifyTeamPin_(id, team.currentPin || pin);
  }

  const row = [id, name, coach, pin];
  
  if (i > 0) {
    sheet.getRange(i + 1, 1, 1, TEAM_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  
  return { id, name, coach };
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
  
  const players = (roster || []).map(cleanPlayer_).filter(p => p.name);
  const labels = new Set();
  
  players.forEach(p => {
    if (labels.has(p.label)) throw new Error(`Duplicate player: ${p.label}`);
    labels.add(p.label);
  });
  
  const sheet = getSheet_(SHEETS.roster);
  const retained = sheet.getDataRange().getValues().slice(1).filter(r => r[0] !== teamId && r[1]);
  const rows = [...retained, ...players.map(p => [teamId, p.name, p.jersey])];
  
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), ROSTER_HEADERS.length).clearContent();
  
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, ROSTER_HEADERS.length).setValues(rows);
  }
  
  return readRoster_(sheet, teamId);
}

function savePlan(pin, plan) {
  setupPlanner();
  const safe = validatePlan_(plan || {});
  verifyTeamPin_(safe.teamId, pin);
  
  const sheet = getSheet_(SHEETS.plans);
  const id = safe.id || Utilities.getUuid();
  const row = [
    id,
    new Date(),
    safe.teamId,
    safe.name,
    safe.date,
    safe.formation,
    JSON.stringify({ ...safe, id })
  ];
  
  const data = sheet.getDataRange().getValues();
  const i = data.findIndex((r, n) => n > 0 && r[0] === id);
  
  if (i > 0) {
    sheet.getRange(i + 1, 1, 1, PLAN_HEADERS.length).setValues([row]);
  } else {
    sheet.appendRow(row);
  }
  
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

function verifyTeamPin_(teamId, pin) {
  const id = String(teamId || '').trim();
  const providedPin = String(pin || '').trim();
  
  const sheet = getSheet_(SHEETS.teams);
  const row = sheet.getDataRange().getValues().slice(1).find(r => r[0] === id);
  
  if (!row) throw new Error('Team not found.');
  
  const storedPin = String(row[3] || '').trim();
  if (storedPin && storedPin !== providedPin) {
    throw new Error('Incorrect team PIN.');
  }
  
  return { id: row[0], name: row[1], coach: row[2] };
}

function spreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('PLANNER_SPREADSHEET_ID');
  const ss = storedId ? SpreadsheetApp.openById(storedId) : SpreadsheetApp.getActiveSpreadsheet();
  
  if (!ss) {
    throw new Error('Open this project from its Google Sheet and run setupPlanner once before deploying as a web app.');
  }
  if (!storedId) {
    props.setProperty('PLANNER_SPREADSHEET_ID', ss.getId());
  }
  return ss;
}

function getSheet_(name) {
  return spreadsheet_().getSheetByName(name);
}

function rewriteWithoutTeam_(sheet, columnCount, teamIdColumn, teamId) {
  const retained = sheet.getDataRange().getValues().slice(1).filter(row => row[teamIdColumn] !== teamId);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), columnCount).clearContent();
  if (retained.length) {
    sheet.getRange(2, 1, retained.length, columnCount).setValues(retained);
  }
}

function rewriteWithoutPlan_(sheet, planId) {
  const retained = sheet.getDataRange().getValues().slice(1).filter(row => row[0] !== planId);
  sheet.getRange(2, 1, Math.max(sheet.getMaxRows() - 1, 1), PLAN_HEADERS.length).clearContent();
  if (retained.length) {
    sheet.getRange(2, 1, retained.length, PLAN_HEADERS.length).setValues(retained);
  }
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

function readTeams_(sheet) {
  return sheet.getDataRange().getValues().slice(1)
    .map(([id, name, coach]) => ({
      id: String(id || ''),
      name: String(name || ''),
      coach: String(coach || '')
    }))
    .filter(t => t.id && t.name)
    .sort((a, b) => a.name.localeCompare(b.name));
}

function readRoster_(sheet, teamId) {
  return sheet.getDataRange().getValues().slice(1)
    .filter(r => r[0] === teamId)
    .map(([, name, jersey]) => cleanPlayer_({ name, jersey }))
    .filter(p => p.name);
}

function readPlans_(sheet, teamId) {
  return sheet.getDataRange().getValues().slice(1)
    .reverse()
    .filter(r => r[2] === teamId)
    .map(r => {
      try { return JSON.parse(r[6]); } catch (e) { return null; }
    })
    .filter(Boolean)
    .slice(0, 30);
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
  const result = {
    id: plan.id || '',
    teamId,
    name,
    date,
    formation,
    absent: uniqueStrings_(plan.absent),
    quarters: {}
  };
  
  ['Q1', 'Q2', 'Q3', 'Q4'].forEach(q => {
    const current = quarters[q] || {};
    const lineup = uniqueStrings_(current.lineup);
    const substitutes = uniqueStrings_(current.substitutes).slice(0, 6);
    
    if (lineup.some(p => substitutes.includes(p))) {
      throw new Error(`${q}: a player cannot be both a starter and a substitute.`);
    }
    if (result.absent.some(p => lineup.includes(p) || substitutes.includes(p))) {
      throw new Error(`${q}: an unavailable player is selected.`);
    }
    
    result.quarters[q] = { lineup, substitutes };
  });
  
  return result;
}

function uniqueStrings_(values) {
  return [...new Set((values || []).map(v => String(v || '').trim()).filter(Boolean))];
}
