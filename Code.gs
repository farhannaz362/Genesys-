/**
 * GENESYS + — Unified WFM Backend v2.1
 * ─────────────────────────────────────
 * Paste as Code.gs in Google Apps Script.
 * Paste updated Index.html as the HTML file.
 *
 * ROUTER MODES:
 *   (default)          → Agent Portal
 *   ?mode=coordinator  → Team Coordinator Portal
 *   ?mode=manager      → Manager Portal
 *   ?mode=farhan       → Admin Portal
 */

/* ═══════════════════════════════════════
   CONFIG
═══════════════════════════════════════ */
const GX = {
  appName: 'Genesys +',
  tz: 'Asia/Karachi',
  staffingPrefix:    'Staffing Matrix',
  breakControlPrefix:'Break Control',
  breakSlotsPrefix:  'Break Slots',
  BREAK_SUBMIT_WINDOW_MINS: 90,
  PULSE_GRACE_MINS:          30,
  AUTO_BREAK_THRESHOLD_MINS: 90,
  BREAK_SLOT_THRESHOLD:      60,
  CALENDAR_ID: '',
  sheets: {
    exceptions:  'Exception Logs',
    liveMonitor: 'Live Monitor',
    access:      'Access_Control',
    performance: 'Agent Performance',
    activity:    'GenesysX_Activity',
    sysActivity: 'Sys_Activity',
    audit:       'Headcount_Audit'
  },
  cols: {
    shiftStart:    0,
    shiftEnd:      1,
    id:            2,
    name:          3,
    dayStart:      4,
    role:          11,
    teamLead:      12,
    utcStart:      14,
    utcEnd:        15,
    email:         16,
    linkedSheetId: 17
  },
  routes: {
    AGENT:                    ['agent','performance'],
    TEAM_COORDINATOR_LOCKED:  ['agent','performance'],
    TEAM_COORDINATOR:         ['agent','performance','monitor','exceptions','headcount','breakCalendar','manager'],
    MANAGER:                  ['monitor','exceptions','headcount','balancing','breakCalendar','manager','logs'],
    ADMIN:                    ['agent','performance','monitor','exceptions','headcount','balancing','breakCalendar','manager','admin','logs']
  }
};

/* ═══════════════════════════════════════
   ENTRY POINTS
═══════════════════════════════════════ */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.mode  = (e && e.parameter && e.parameter.mode) ? String(e.parameter.mode) : '';
  return template.evaluate()
    .setTitle('Genesys +')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Genesys +')
    .addItem('Open Portal',            'gxOpenPortal')
    .addSeparator()
    .addItem('Refresh Live Monitor',   'gxRefreshLiveMonitorMenu')
    .addItem('Recalculate Headcount',  'gxRecalculateHeadcountMenu')
    .addItem('Run Balancing Engine',   'gxRunBalancingMenu')
    .addItem('Setup Mock Test Data',   'gxSetupMockData')
    .addSeparator()
    .addItem('Setup 5-Min Trigger',    'gxSetupTriggers_')
    .addToUi();
}

function gxOpenPortal() {
  const url  = ScriptApp.getService().getUrl();
  const html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '","_blank");google.script.host.close();<\/script>'
  ).setWidth(10).setHeight(10);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Genesys +…');
}

function gxRefreshLiveMonitorMenu()  { gxGetLiveMonitorData_({ autoSync:true, writeSnapshot:true }); SpreadsheetApp.getActiveSpreadsheet().toast('Live Monitor refreshed','Genesys +',4); }
function gxRecalculateHeadcountMenu(){ const r = gxCalculateHeadcount_({ writeSheet:true }); SpreadsheetApp.getActiveSpreadsheet().toast(r.message,'Genesys +',4); }
function gxRunBalancingMenu()        { const r = gxGetBalancingData_(); SpreadsheetApp.getActiveSpreadsheet().toast('Balancing: ' + r.suggestions.length + ' suggestions','Genesys +',4); }
function gxScheduledRefresh()        { gxGetLiveMonitorData_({ autoSync:true, writeSnapshot:true }); gxAutoSubmitBreaks_(); }

/* ═══════════════════════════════════════
   PUBLIC API DISPATCHER
═══════════════════════════════════════ */
function gxApi(action, payload) {
  payload = payload || {};
  try {
    switch (action) {
      case 'bootstrap':        return gxBootstrap(payload.mode || '');
      case 'liveMonitor':      { const ctx = gxGetAccessContext_(SpreadsheetApp.getActiveSpreadsheet(), payload.mode||''); const d = gxGetLiveMonitorData_({ autoSync:true, writeSnapshot:true, context:ctx }); return gxOk_(d); }
      case 'exceptions':       return gxOk_({ rows: gxGetExceptionLogs_(payload.limit || 200) });
      case 'submitException':  return gxOk_(gxSubmitException_(payload));
      case 'markExceptionDone':return gxOk_(gxMarkExceptionDone_(payload));
      case 'headcount':        return gxOk_(gxCalculateHeadcount_({ writeSheet:!!payload.writeSheet }));
      case 'headcountGrids':   return gxOk_(gxGetHeadcountGrids_(payload));
      case 'balancing':        return gxOk_(gxGetBalancingData_(payload));
      case 'applySuggestion':  return gxOk_(gxApplyBalancingSuggestion_(payload));
      case 'breakCalendar':    return gxOk_(gxGetBreakCalendar_(payload));
      case 'moveBreak':        return gxOk_(gxMoveBreak_(payload));
      case 'submitBreak':      return gxOk_(gxSubmitBreak_(payload));
      case 'performance':      return gxOk_(gxGetPerformance_(payload));
      case 'logs':             return gxOk_({ rows: gxGetActivityLogs_(payload.limit || 200) });
      case 'accessControl':    return gxOk_(gxGetAccessRows_());
      case 'saveAccess':       return gxOk_(gxSaveAccessRow_(payload));
      case 'revokeAccess':     return gxOk_(gxRevokeAccessRow_(payload));
      case 'adminInitialize':  return gxOk_(gxAdminInitialize_());
      case 'setupMockData':    return gxOk_(gxSetupMockData());
      case 'setupTriggers':    return gxOk_(gxSetupTriggers_());
      case 'overrideState':    return gxOk_(gxOverrideAgentState_(payload));
      case 'runReorganiser':   return gxOk_(gxRunReorganiser_());
      default: throw new Error('Unknown action: ' + action);
    }
  } catch(err) {
    gxLog_('API_ERROR','',action + ': ' + (err.message||err),'SYSTEM');
    return { ok:false, error: err.message || String(err) };
  }
}

function gxBootstrap(mode) {
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  gxEnsureCoreSheets_(ss);
  const context = gxGetAccessContext_(ss, mode);
  // Flag when the script runs as its deployer (common when "Execute as: Me" is set).
  // The frontend uses this to skip the strict email-match check on sign-in.
  context.isDeployer = (gxGetUserEmail_() === gxGetDeployerEmail_());
  const roster  = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const agents  = gxGetAgents_(ss, context);
  const live    = gxGetLiveMonitorData_({ autoSync:true, writeSnapshot:true, context:context });
  const stats   = gxBuildStats_(live.rows);
  return {
    ok:true, appName:GX.appName, now:gxNowStamp_(),
    context, routes:context.routes, defaultView:context.defaultView,
    roster:{ name:roster.name, found:!!roster.sheet },
    agents, live, stats,
    exceptionTypes: ['Late Arrival','Early Departure','Absent','Overtime']
  };
}

function gxOk_(data) { data = data || {}; data.ok = true; return data; }

/* ═══════════════════════════════════════
   ACCESS ROUTER
═══════════════════════════════════════ */
function gxGetAccessContext_(ss, requestedMode) {
  const email = gxGetUserEmail_();
  const mode  = gxNormalizeMode_(requestedMode);
  const fallback = {
    email, mode,
    name:       email ? gxTitleCase_(email.split('@')[0].replace(/[._]/g,' ')) : 'Genesys User',
    role:       'AGENT', roleLabel:'Agent',
    team:'', agentId:'',
    routes:     GX.routes.AGENT,
    defaultView:'agent',
    accessLocked:false, accessMessage:''
  };

  const rosterIdentity = gxFindRosterIdentityByEmail_(ss, email);
  const accessIdentity = gxFindAccessIdentityByEmail_(ss, email);
  const op = gxOperationalDate_(new Date());

  /* ── AGENT (default) ── */
  if (mode === 'agent') {
    if (!rosterIdentity) return Object.assign({},fallback,{ accessLocked:true, accessMessage:'Email not found in current roster. Please contact your manager.' });
    const shift    = gxBuildShiftWindow_(op.ymd, rosterIdentity.shiftStart, rosterIdentity.shiftEnd);
    const nowMs    = Date.now();
    const minsFromStart = (nowMs - shift.start.getTime()) / 60000;
    const minsToStart   = (shift.start.getTime() - nowMs) / 60000;
    const breakAlreadyDone = gxHasSubmittedBreak_(ss, email, op.ymd);
    const windowOpen = breakAlreadyDone || minsToStart <= 15 || (minsFromStart >= 0 && minsFromStart <= GX.BREAK_SUBMIT_WINDOW_MINS);
    if (!windowOpen && !breakAlreadyDone) {
      return Object.assign({},fallback,rosterIdentity,{
        role:'AGENT', roleLabel:'Agent', routes:GX.routes.AGENT, defaultView:'agent',
        accessLocked:true, accessMessage:'Submission window has closed (90 minutes from shift start). Please contact your coordinator.'
      });
    }
    return Object.assign({},fallback,rosterIdentity,{ role:'AGENT', roleLabel:'Agent', routes:GX.routes.AGENT, defaultView:'agent' });
  }

  /* ── COORDINATOR ── */
  if (mode === 'coordinator') {
    if (!accessIdentity || !['TEAM_COORDINATOR','ADMIN'].includes(accessIdentity.role)) {
      return Object.assign({},fallback,rosterIdentity||{},{ mode, accessLocked:true, accessMessage:'Coordinator access not granted in Access_Control.' });
    }
    const submitted = gxHasSubmittedBreak_(ss, email, op.ymd);
    const locked    = !submitted && accessIdentity.role !== 'ADMIN';
    return Object.assign({},fallback,rosterIdentity||{},accessIdentity,{
      role:'TEAM_COORDINATOR', roleLabel:'Team Coordinator', mode,
      routes:    locked ? GX.routes.TEAM_COORDINATOR_LOCKED : GX.routes.TEAM_COORDINATOR,
      defaultView:'agent', accessLocked:locked,
      accessMessage: locked ? 'Submit your own breaks to unlock coordinator tools.' : ''
    });
  }

  /* ── MANAGER ── */
  if (mode === 'manager') {
    if (!accessIdentity || !['MANAGER','ADMIN'].includes(accessIdentity.role)) {
      return Object.assign({},fallback,rosterIdentity||{},{ mode, accessLocked:true, accessMessage:'Manager access not granted in Access_Control.' });
    }
    return Object.assign({},fallback,rosterIdentity||{},accessIdentity,{
      role:'MANAGER', roleLabel:'Manager', mode,
      routes:GX.routes.MANAGER, defaultView:'monitor'
    });
  }

  /* ── ADMIN ── */
  if (mode === 'farhan') {
    if (!accessIdentity || accessIdentity.role !== 'ADMIN') {
      return Object.assign({},fallback,rosterIdentity||{},{ mode, accessLocked:true, accessMessage:'Admin access not granted in Access_Control.' });
    }
    return Object.assign({},fallback,rosterIdentity||{},accessIdentity,{
      role:'ADMIN', roleLabel:'Admin', mode,
      routes:GX.routes.ADMIN, defaultView:'monitor'
    });
  }

  return fallback;
}

function gxFindAccessIdentityByEmail_(ss, email) {
  if (!email) return null;
  const sheet = ss.getSheetByName(GX.sheets.access);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const vals    = sheet.getDataRange().getDisplayValues();
  const headers = vals[0].map(h => h.trim().toLowerCase());
  const ec = headers.indexOf('email'), rc = headers.indexOf('role'), nc = headers.indexOf('name'), tc = headers.indexOf('team');
  for (let i = 1; i < vals.length; i++) {
    const rowEmail = String(vals[i][ec]||'').trim().toLowerCase();
    if (rowEmail && rowEmail === email) {
      const role = gxNormalizeRole_(vals[i][rc]);
      return { email, name:vals[i][nc]||'', role, roleLabel:gxRoleLabel_(role), team:vals[i][tc]||'', agentId:'' };
    }
  }
  return null;
}

function gxFindRosterIdentityByEmail_(ss, email) {
  if (!email) return null;
  const rObj  = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const sheet = rObj.sheet;
  if (!sheet || sheet.getLastRow() < 3) return null;
  const rows = sheet.getRange(3, 1, sheet.getLastRow()-2, Math.min(20,sheet.getLastColumn())).getDisplayValues();
  for (let i = 0; i < rows.length; i++) {
    const rowEmail = String(rows[i][GX.cols.email]||'').trim().toLowerCase();
    // CRITICAL: skip blank emails, require exact non-empty match
    if (!rowEmail) continue;
    if (rowEmail !== email) continue;
    const rawName = String(rows[i][GX.cols.name]||'');
    const role    = rawName.includes('- TC') ? 'TEAM_COORDINATOR' : 'AGENT';
    return {
      email, role, roleLabel:gxRoleLabel_(role),
      name:       gxCleanName_(rawName) || email,
      team:       gxCleanName_(rows[i][GX.cols.teamLead]),
      agentId:    String(rows[i][GX.cols.id]||'').trim(),
      shiftStart: gxTimeString_(rows[i][GX.cols.shiftStart]),
      shiftEnd:   gxTimeString_(rows[i][GX.cols.shiftEnd]),
      rosterRow:  3 + i
    };
  }
  return null;
}

function gxNormalizeMode_(m) {
  m = String(m||'').trim().toLowerCase();
  if (m === 'coordinator' || m === 'tc')   return 'coordinator';
  if (m === 'manager'     || m === 'mgr')  return 'manager';
  if (m === 'farhan'      || m === 'admin')return 'farhan';
  return 'agent';
}

function gxNormalizeRole_(r) {
  r = String(r||'').trim().toUpperCase().replace(/[\s\-]+/g,'_');
  if (['ADMIN','SUPER_ADMIN'].includes(r))                                  return 'ADMIN';
  if (['MANAGER','MGR','SUPERVISOR'].includes(r))                           return 'MANAGER';
  if (['TEAM_COORDINATOR','TC','TEAM_LEAD','COORDINATOR'].includes(r))      return 'TEAM_COORDINATOR';
  return 'AGENT';
}

function gxRoleLabel_(role) {
  return { AGENT:'Agent', TEAM_COORDINATOR:'Team Coordinator', MANAGER:'Manager', ADMIN:'Admin' }[role] || 'Agent';
}

function gxHasSubmittedBreak_(ss, email, opYmd) {
  return !!gxBuildBreakMap_(ss, opYmd)[String(email||'').trim().toLowerCase()];
}

/* ═══════════════════════════════════════
   CORE DATA — AGENTS
═══════════════════════════════════════ */
function gxGetAgents_(ss, context) {
  const rObj  = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const sheet = rObj.sheet;
  if (!sheet || sheet.getLastRow() < 3) return [];
  const rows   = sheet.getRange(3, 1, sheet.getLastRow()-2, Math.min(20,sheet.getLastColumn())).getDisplayValues();
  const agents = [];
  rows.forEach((r,i) => {
    const id    = String(r[GX.cols.id]  ||'').trim();
    const name  = gxCleanName_(r[GX.cols.name]);
    const email = String(r[GX.cols.email]||'').trim().toLowerCase();
    if (!id || !name || id === 'ID') return;
    const team = gxCleanName_(r[GX.cols.teamLead]);
    if (context && context.role === 'TEAM_COORDINATOR' && context.team && team !== context.team) return;
    agents.push({
      row:          3+i, id, name, team,
      role:         String(r[GX.cols.role]||'').trim(),
      email,
      linkedSheetId:String(r[GX.cols.linkedSheetId]||'').trim(),
      shiftStart:   gxTimeString_(r[GX.cols.shiftStart]),
      shiftEnd:     gxTimeString_(r[GX.cols.shiftEnd]),
      utcStart:     gxTimeString_(r[GX.cols.utcStart]),
      utcEnd:       gxTimeString_(r[GX.cols.utcEnd])
    });
  });
  return agents.sort((a,b) => a.name.localeCompare(b.name));
}

/* ═══════════════════════════════════════
   LIVE MONITOR
═══════════════════════════════════════ */
function gxGetLiveMonitorData_(options) {
  options = options || {};
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const context = options.context || gxGetAccessContext_(ss,'');
  const rObj    = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const roster  = rObj.sheet;
  if (!roster) return { rows:[], stats:{}, error:'Missing roster: '+rObj.name };

  const op           = gxOperationalDate_(new Date());
  const dayIndex     = gxDayIndexForRoster_(roster, op.ymd);
  const agents       = gxGetAgents_(ss, context);
  const pulseMap     = gxBuildPulseMap_(agents);
  const breakMap     = gxBuildBreakMap_(ss, op.ymd);
  const exceptionMap = gxBuildExceptionMap_(ss, op.ymd);
  const now          = new Date();
  const rows         = [];

  agents.forEach(agent => {
    const rosterCell  = String(roster.getRange(agent.row, 5+dayIndex).getDisplayValue()||'').trim().toLowerCase();
    const scheduled   = rosterCell === '' || rosterCell === 'ed';
    if (!scheduled && rosterCell !== 'absent') return;

    const shift         = gxBuildShiftWindow_(op.ymd, agent.shiftStart, agent.shiftEnd);
    const pulse         = pulseMap[agent.id] || null;

    // Pulse only valid if: within shift hours AND less than 9 hours ago
    const nineHoursAgo = now.getTime() - 9*60*60*1000;
    const pulseActive   = !!(pulse && pulse.date
      && pulse.date.getTime() >= shift.start.getTime() - 5*60000
      && pulse.date.getTime() <= shift.end.getTime() + 15*60000
      && pulse.date.getTime() >= nineHoursAgo);

    const breakInfo     = breakMap[agent.email] || null;
    const breakSubmitted = !!breakInfo;

    if (options.autoSync) gxAutoSyncAttendance_(ss, roster, agent, dayIndex, op.ymd, shift, pulse, exceptionMap);

    const exception     = gxLatestExceptionFor_(exceptionMap, agent.id);
    const isAbsent      = !!(exception && exception.type === 'Absent' && exception.status !== 'Recovered');

    // Status: absent agents show as OFFLINE with exception=Absent
    let status = isAbsent ? 'OFFLINE' : gxCombinedStatus_(pulseActive, breakSubmitted);

    if (!isAbsent) {
      const breakPhase = breakInfo ? gxBreakPhase_(breakInfo.breaks, now) : { phase:'NONE' };
      if (breakPhase.phase === 'ON_BREAK') status = 'BREAK';
      else if (breakPhase.phase === 'WRAP_UP') status = 'BREAK WRAP UP';

      const minsToEnd = (shift.end.getTime() - now.getTime()) / 60000;
      if (minsToEnd <= 20 && minsToEnd > 0 && status === 'ONLINE') status = 'SHIFT WRAP UP';
    }

    const nextBreak = (!isAbsent && breakInfo) ? gxNextBreakInfo_(breakInfo, now) : null;

    rows.push({
      id:agent.id, name:agent.name, team:agent.team, role:agent.role, email:agent.email,
      shift:          agent.shiftStart + '–' + agent.shiftEnd,
      shiftStartMs:   shift.start.getTime(),
      shiftEndMs:     shift.end.getTime(),
      pulseActive,
      pulseLast:      pulse && pulse.date ? Utilities.formatDate(pulse.date, GX.tz,'HH:mm') : '',
      breakSubmitted: isAbsent ? false : breakSubmitted,
      breakPlan:      (!isAbsent && breakInfo) ? breakInfo.plan : '',
      breaks:         (!isAbsent && breakInfo) ? breakInfo.breaks : [],
      nextBreak,
      status,
      exception:      exception ? exception.type  : '',
      exceptionStatus:exception ? exception.status : '',
      lastActivity:   pulse && pulse.date ? gxHumanAge_(pulse.date, now) : 'No signal',
      risk:           gxRiskForRow_(status, pulseActive, isAbsent ? false : breakSubmitted, exception, shift, now)
    });
  });

  if (options.writeSnapshot) gxWriteLiveSnapshot_(ss, rows);
  return { rows, stats:gxBuildStats_(rows), generatedAt:gxNowStamp_(), opDate:op.ymd, rosterName:rObj.name };
}

function gxCombinedStatus_(pulse, breakDone) {
  if (pulse  && breakDone)  return 'ONLINE';
  if (pulse  && !breakDone) return 'ACTIVE';
  if (!pulse && breakDone)  return 'IDLE';
  return 'OFFLINE';
}

function gxNextBreakInfo_(breakInfo, now) {
  if (!breakInfo || !breakInfo.breaks || !breakInfo.breaks.length) return null;
  const nowMs = now.getTime();
  for (const b of breakInfo.breaks) {
    if (nowMs < b.endMs) {
      return { label:b.label, start:b.start, startMs:b.startMs, endMs:b.endMs, duration:b.duration, ongoing: nowMs >= b.startMs };
    }
  }
  return null;
}

function gxAutoSyncAttendance_(ss, roster, agent, dayIndex, opYmd, shift, pulse, exMap) {
  const now      = new Date();
  const records  = (exMap[agent.id]||[]);
  const openAbs  = records.find(r => r.type==='Absent'       && r.status!=='Recovered');
  const existLate= records.find(r => r.type==='Late Arrival' && r.status!=='Cancelled');
  const pulseTime = pulse && pulse.date ? pulse.date : null;

  if (!pulseTime && now.getTime() > shift.start.getTime() + GX.PULSE_GRACE_MINS*60000 && !openAbs && !existLate) {
    gxAppendException_(ss, {
      agentId:agent.id, agentName:agent.name, teamLead:agent.team,
      date:opYmd, shift:agent.shiftStart+' - '+agent.shiftEnd,
      type:'Absent', timeValue:'', endTime:'', duration:'',
      source:'AUTO_NO_PULSE', status:'Open', notes:'Auto: no pulse within '+GX.PULSE_GRACE_MINS+' min of shift start'
    });
    roster.getRange(agent.row, 5+dayIndex).setValue('Absent');
    gxLog_('AGENT_AUTO_ABSENT', agent.id,'No pulse grace expired','SYSTEM');
    return;
  }

  if (pulseTime && pulseTime.getTime() > shift.start.getTime() + GX.PULSE_GRACE_MINS*60000) {
    const lateTime = Utilities.formatDate(pulseTime, GX.tz, 'HH:mm');
    if (openAbs) {
      gxRecoverAbsentAsLate_(ss, openAbs.rowNumber, lateTime, pulseTime, agent);
      roster.getRange(agent.row, 5+dayIndex).clearContent();
    } else if (!existLate) {
      gxAppendException_(ss, {
        agentId:agent.id, agentName:agent.name, teamLead:agent.team,
        date:opYmd, shift:agent.shiftStart+' - '+agent.shiftEnd,
        type:'Late Arrival', timeValue:lateTime, endTime:'', duration:gxDuration_(agent.shiftStart, lateTime),
        source:'AUTO_PULSE', status:'Open', notes:'Auto: first pulse detected after grace'
      });
      gxLog_('AGENT_LATE_AUTO', agent.id,'Late arrival at '+lateTime,'SYSTEM');
    }
  }
}

function gxAutoSubmitBreaks_() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const op  = gxOperationalDate_(new Date());
  const all = gxGetAgents_(ss, {});
  const breakMap = gxBuildBreakMap_(ss, op.ymd);
  const now = new Date();
  all.forEach(agent => {
    if (!agent.email) return;
    const shift = gxBuildShiftWindow_(op.ymd, agent.shiftStart, agent.shiftEnd);
    const minsFromStart = (now.getTime() - shift.start.getTime()) / 60000;
    if (minsFromStart < GX.AUTO_BREAK_THRESHOLD_MINS) return;
    if (breakMap[agent.email]) return;
    const mid   = shift.start.getTime() + (shift.end.getTime()-shift.start.getTime())/2;
    const b1Ms  = mid - 30*60000;
    const b2Ms  = mid + 60*60000;
    const fmt   = ms => { const d=new Date(ms); return pad_(d.getHours())+':'+pad_(d.getMinutes()); };
    gxSubmitBreak_({
      agentId:agent.id, name:agent.name, email:agent.email,
      shift:agent.shiftStart+' - '+agent.shiftEnd,
      shiftDate:op.ymd, plan:'T1', b1:fmt(b1Ms), b2:fmt(b2Ms), b3:'',
      submittedBy:'AUTO_SYSTEM', allowDuplicate:false
    });
    gxLog_('AUTO_BREAK_SUBMITTED', agent.id,'Auto-assigned breaks (90min window expired)','SYSTEM');
  });
}

/* ── PULSE MAP ── */
function gxBuildPulseMap_(agents) {
  const cache = {}, result = {};
  agents.forEach(agent => {
    if (!agent.linkedSheetId || !agent.email) return;
    const pack = gxGetPulseCache_(agent.linkedSheetId, cache);
    if (!pack || pack.error) return;
    let hit = pack.map[agent.email];
    if (!hit) Object.keys(pack.map).some(k => { if (k.includes(agent.email)||agent.email.includes(k)){ hit=pack.map[k]; return true; } return false; });
    if (hit) result[agent.id] = { date:hit };
  });
  return result;
}

function gxGetPulseCache_(sheetId, cache) {
  if (cache[sheetId] !== undefined) return cache[sheetId];
  try {
    const ss  = SpreadsheetApp.openById(sheetId);
    const tab = ss.getSheetByName('Sys_Activity');
    if (!tab) return cache[sheetId] = { error:'NO_SYS_ACTIVITY', map:{} };
    const last = tab.getLastRow();
    if (last < 1) return cache[sheetId] = { map:{} };
    const start = Math.max(1, last-1000);
    const vals  = tab.getRange(start, 1, last-start+1, 2).getValues();
    const map   = {};
    vals.forEach(r => {
      const em = String(r[0]||'').trim().toLowerCase();
      if (!em) return;
      const dt = gxParseDate_(r[1]);
      if (dt && (!map[em] || dt > map[em])) map[em] = dt;
    });
    return cache[sheetId] = { map };
  } catch(e) { return cache[sheetId] = { error:'ACCESS_DENIED', map:{} }; }
}

/* ── BREAK MAP ── */
function gxBuildBreakMap_(ss, opYmd) {
  const name  = gxWeeklySheetName_(GX.breakControlPrefix, opYmd);
  const sheet = ss.getSheetByName(name);
  const map   = {};
  if (!sheet || sheet.getLastRow() < 2) return map;
  const vals = sheet.getRange(2,1,sheet.getLastRow()-1, Math.max(12,sheet.getLastColumn())).getValues();
  vals.forEach((r,i) => {
    const email   = String(r[7]||'').trim().toLowerCase();
    const rowDate = r[8] instanceof Date ? Utilities.formatDate(r[8],GX.tz,'yyyy-MM-dd') : String(r[8]||'').trim();
    if (!email || rowDate !== opYmd || !r[4]) return;
    const plan   = String(r[3]||'');
    const breaks = [];
    gxPushBreak_(breaks, r[4], plan.includes('T2') ? 'T2-30B1' : 'T1-30B1', 30);
    gxPushBreak_(breaks, r[5], plan.includes('T2') ? 'T2-15B1' : 'T1-30B2', plan.includes('T2') ? 15 : 30);
    if (plan.includes('T2')) gxPushBreak_(breaks, r[6], 'T2-15B2', 15);
    breaks.sort((a,b) => a.startMs-b.startMs);
    map[email] = { rowNumber:i+2, plan, breaks };
  });
  return map;
}

function gxPushBreak_(arr, value, label, mins) {
  if (!value || String(value).includes('N/A')) return;
  const dt = gxParseDate_(value);
  if (!dt) return;
  arr.push({
    label, duration:mins,
    start:  Utilities.formatDate(dt,GX.tz,'HH:mm'),
    startMs:dt.getTime(),
    endMs:  dt.getTime() + mins*60000
  });
}

function gxBreakPhase_(breaks, now) {
  if (!breaks || !breaks.length) return { phase:'NONE', label:'' };
  const nowMs = now.getTime();
  for (const b of breaks) {
    const wrapStart = b.startMs - 10*60000;
    if (nowMs >= b.startMs && nowMs < b.endMs)   return { phase:'ON_BREAK',  label:b.start };
    if (nowMs >= wrapStart  && nowMs < b.startMs) return { phase:'WRAP_UP',   label:b.start };
    if (b.startMs > nowMs)                        return { phase:'UPCOMING',  label:b.start };
  }
  return { phase:'COMPLETED', label:'Done' };
}

/* ── EXCEPTION MAP ── */
function gxBuildExceptionMap_(ss, opYmd) {
  const map  = {};
  const rows = gxGetExceptionLogs_(1000).filter(r => r.date === opYmd);
  rows.forEach(r => { if (!map[r.agentId]) map[r.agentId]=[]; map[r.agentId].push(r); });
  return map;
}

function gxLatestExceptionFor_(map, agentId) {
  return (map[agentId]||[]).find(r => r.status !== 'Cancelled') || null;
}

/* ── SNAPSHOT ── */
function gxWriteLiveSnapshot_(ss, rows) {
  const sheet = gxEnsureSheet_(ss, GX.sheets.liveMonitor,
    ['Last Updated','Agent','ID','Team','Shift','Pulse','Break','Status','Exception','Last Activity','Risk']);
  if (sheet.getLastRow() > 1) sheet.getRange(2,1,sheet.getLastRow()-1,11).clearContent();
  if (!rows.length) return;
  const stamp = gxNowStamp_();
  const vals  = rows.map(r => [stamp,r.name,r.id,r.team,r.shift,
    r.pulseActive?'ACTIVE':'OFFLINE', r.breakSubmitted?'SUBMITTED':'PENDING',
    r.status, r.exception, r.lastActivity, r.risk]);
  sheet.getRange(2,1,vals.length,11).setValues(vals);
}

/* ── STATS ── */
function gxBuildStats_(rows) {
  rows = rows || [];
  const cnt = s => rows.filter(r=>r.status===s).length;
  return {
    scheduled: rows.length, online:cnt('ONLINE'), active:cnt('ACTIVE'),
    idle:cnt('IDLE'), offline:cnt('OFFLINE'), absent:0,
    break:cnt('BREAK'),
    // Absent counted by exception, not status (status shows OFFLINE for absent)
    absent: rows.filter(r => r.exception === 'Absent').length,
    alerts: rows.filter(r=>['HIGH','CRITICAL'].includes(r.risk)).length,
    adherence: rows.length ? Math.round((cnt('ONLINE')/rows.length)*100) : 0
  };
}

/* ── RISK ── */
function gxRiskForRow_(status, pulse, breakDone, exception, shift, now) {
  if (exception && exception.type === 'Absent') return 'CRITICAL';
  if (exception && ['Late Arrival','Early Departure'].includes(exception.type)) return 'HIGH';
  if (!pulse && now.getTime() > shift.start.getTime() + 20*60000) return 'HIGH';
  if (!breakDone && now.getTime() > shift.start.getTime() + 60*60000) return 'MEDIUM';
  if (status === 'OFFLINE') return 'MEDIUM';
  return 'LOW';
}

/* ═══════════════════════════════════════
   EXCEPTION CENTER
═══════════════════════════════════════ */
function gxSubmitException_(payload) {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const rObj   = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const roster = rObj.sheet;
  if (!roster) throw new Error('Missing roster: '+rObj.name);

  const agent  = gxFindAgentById_(ss, payload.agentId);
  if (!agent) throw new Error('Agent not found: '+payload.agentId);
  const type   = gxNormalizeExType_(payload.type);
  const opYmd  = payload.date || gxOperationalDate_(new Date()).ymd;
  const dayIdx = gxDayIndexForRoster_(roster, opYmd);

  if (type === 'Overtime') {
    const shiftStart = gxParseTime_(agent.shiftStart);
    const shiftEnd   = gxParseTime_(agent.shiftEnd);
    const otStart    = gxParseTime_(payload.timeValue);
    const otEnd      = gxParseTime_(payload.endTime);
    const inside = (t) => {
      if (shiftEnd.decimal > shiftStart.decimal)
        return t.decimal >= shiftStart.decimal && t.decimal <= shiftEnd.decimal;
      return t.decimal >= shiftStart.decimal || t.decimal <= shiftEnd.decimal;
    };
    if (inside(otStart) || inside(otEnd)) throw new Error('Overtime must be outside scheduled shift hours.');
  }

  if (['Late Arrival','Early Departure'].includes(type) && payload.timeValue) {
    const shift = gxBuildShiftWindow_(opYmd, agent.shiftStart, agent.shiftEnd);
    const tvMs  = gxBuildDateAt_(opYmd, payload.timeValue).getTime();
    if (tvMs < shift.start.getTime() || tvMs > shift.end.getTime())
      throw new Error(type + ' time must be within shift hours (' + agent.shiftStart + '–' + agent.shiftEnd + ').');
  }

  let timeValue = payload.timeValue || '';
  let duration  = '';
  if (type === 'Late Arrival' && payload.timeValue) {
    duration = gxDuration_(agent.shiftStart, payload.timeValue);
  } else if (type === 'Early Departure' && payload.timeValue) {
    duration = gxDuration_(payload.timeValue, agent.shiftEnd);
  } else if (type === 'Overtime') {
    timeValue = (payload.timeValue||'') + ' - ' + (payload.endTime||'');
    duration  = gxDuration_(payload.timeValue, payload.endTime);
  }

  const saved = gxAppendException_(ss, {
    agentId:agent.id, agentName:agent.name, teamLead:agent.team,
    date: type === 'Overtime' ? (payload.otDate||opYmd) : opYmd,
    shift: type === 'Overtime' ? opYmd : agent.shiftStart+' - '+agent.shiftEnd,
    type, timeValue, endTime:payload.endTime||'', duration,
    source:payload.source||'MANUAL', status:'Open', notes:payload.notes||''
  });

  if (type === 'Absent')          roster.getRange(agent.row, 5+dayIdx).setValue('Absent');
  if (type === 'Early Departure') roster.getRange(agent.row, 5+dayIdx).setValue('ED');
  if (type === 'Late Arrival') {
    const cell = roster.getRange(agent.row, 5+dayIdx);
    if (cell.getDisplayValue().toLowerCase() === 'absent') cell.clearContent();
  }

  gxCalculateHeadcount_({ writeSheet:true });
  gxLog_('EXCEPTION_LOGGED', agent.id, type+' '+timeValue, payload.source||gxGetUserEmail_());
  return { message:'Exception logged', rowNumber:saved.rowNumber };
}

function gxAppendException_(ss, entry) {
  const sheet    = gxEnsureExceptionSheet_(ss);
  const existing = gxReadExceptionObjects_(sheet, 800);
  const dupe = existing.some(r =>
    r.agentId === entry.agentId && r.date === entry.date && r.type === entry.type &&
    String(r.timeValue||'') === String(entry.timeValue||'') && r.status !== 'Cancelled');
  if (dupe) return { message:'Duplicate ignored', rowNumber:-1 };
  const dayName = gxDayName_(entry.date);
  sheet.appendRow([new Date(), entry.agentId, entry.agentName, entry.teamLead, dayName, entry.date, entry.shift,
    entry.type, entry.timeValue||'', entry.duration||'', entry.source||'SYSTEM', entry.status||'Open', entry.notes||'']);
  return { rowNumber:sheet.getLastRow() };
}

function gxRecoverAbsentAsLate_(ss, rowNumber, lateTime, pulseDate, agent) {
  const sheet = gxEnsureExceptionSheet_(ss);
  const dur   = gxDuration_(agent.shiftStart, lateTime);
  sheet.getRange(rowNumber,1).setValue(new Date());
  sheet.getRange(rowNumber,8).setValue('Late Arrival');
  sheet.getRange(rowNumber,9).setValue(lateTime);
  sheet.getRange(rowNumber,10).setValue(dur);
  sheet.getRange(rowNumber,11).setValue('AUTO_PULSE_RECOVERY');
  sheet.getRange(rowNumber,12).setValue('Recovered');
  sheet.getRange(rowNumber,13).setValue('Recovered from auto-absent; pulse at '+Utilities.formatDate(pulseDate,GX.tz,'HH:mm:ss'));
}

function gxMarkExceptionDone_(payload) {
  if (!payload.rowNumber || payload.rowNumber < 2) throw new Error('Invalid row number.');
  const sheet = gxEnsureExceptionSheet_(SpreadsheetApp.getActiveSpreadsheet());
  sheet.getRange(payload.rowNumber, 12).setValue('Completed');
  sheet.getRange(payload.rowNumber, 13).setValue('Completed by '+gxGetUserEmail_()+' at '+gxNowStamp_());
  gxLog_('EXCEPTION_COMPLETED','','Row '+payload.rowNumber, gxGetUserEmail_());
  return { message:'Exception marked completed.' };
}

function gxGetExceptionLogs_(limit) {
  const sheet = gxEnsureExceptionSheet_(SpreadsheetApp.getActiveSpreadsheet());
  return gxReadExceptionObjects_(sheet, limit||200);
}

function gxReadExceptionObjects_(sheet, limit) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const last  = sheet.getLastRow();
  const count = Math.min(limit||1000, last-1);
  const start = Math.max(2, last-count+1);
  return sheet.getRange(start,1,count,13).getDisplayValues().map((r,i) => ({
    rowNumber:start+i, timestamp:r[0], agentId:String(r[1]||'').trim(),
    agentName:r[2], teamLead:r[3], day:r[4], date:r[5], shift:r[6],
    type:r[7], timeValue:r[8], duration:r[9], source:r[10], status:r[11], notes:r[12]
  })).reverse();
}

function gxNormalizeExType_(type) {
  const t = String(type||'').trim().toLowerCase();
  if (t.includes('late'))  return 'Late Arrival';
  if (t.includes('early')) return 'Early Departure';
  if (t.includes('abs'))   return 'Absent';
  if (t.includes('over') || t === 'ot') return 'Overtime';
  throw new Error('Unsupported exception type: '+type);
}

/* ═══════════════════════════════════════
   HEADCOUNT GRIDS
═══════════════════════════════════════ */
function gxGetHeadcountGrids_(payload) {
  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const op       = gxOperationalDate_(new Date());
  const name     = gxWeeklySheetName_(GX.breakSlotsPrefix, op.ymd);
  const sheet    = ss.getSheetByName(name);
  const weekIdx  = Number(payload && payload.weekIndex) || 0;
  const showAll  = !!(payload && payload.showAll);

  const weeks = [];
  if (sheet && sheet.getLastRow() >= 1) {
    const headerVals = sheet.getRange(1, 2, 1, 28).getDisplayValues()[0];
    for (let w = 0; w < 4; w++) {
      const startDate = headerVals[w*7] || ('Week '+(w+1));
      weeks.push(startDate);
    }
  } else {
    for (let w=0;w<4;w++) weeks.push('Week '+(w+1));
  }

  const colOffset = showAll ? 0 : weekIdx * 7;
  const numCols   = showAll ? 28 : 7;

  function readGrid(startRow, numRows) {
    if (!sheet || sheet.getLastRow() < startRow + numRows - 1) return { headers:[], rows:[] };
    const headerRow = sheet.getRange(1, 2+colOffset, 1, numCols).getDisplayValues()[0];
    const dataRange = sheet.getRange(startRow, 2+colOffset, numRows, numCols).getDisplayValues();
    const noteRange = sheet.getRange(startRow, 2+colOffset, numRows, numCols).getNotes();
    const rows = [];
    for (let r = 0; r < numRows; r++) {
      const cells = dataRange[r].map((v,c) => ({ value:v, display:v, note:noteRange[r][c] }));
      rows.push({ label:pad_(r)+':00', cells });
    }
    return { headers: headerRow, rows };
  }

  // Break Audits: read from Break Control and build grid
  const breakAuditGrid = gxBuildBreakAuditGrid_(ss, op.ymd, colOffset, numCols, weeks, showAll ? null : weekIdx);

  return {
    weeks,
    currentWeek: weekIdx,
    showAll,
    slots:      readGrid(4,   24),
    breaks:     breakAuditGrid,
    adherence:  readGrid(108, 24),
    live:       readGrid(134, 24)
  };
}

function gxBuildBreakAuditGrid_(ss, opYmd, colOffset, numCols, weeks, weekIdx) {
  // Build break audit data from Break Control sheets for display
  // Returns grid with headers (dates) and rows (hours 0-23), cells = minutes occupied + notes
  const headers = [];
  const rowData = Array.from({length:24}, () => Array(numCols).fill(0));
  const rowNotes= Array.from({length:24}, () => Array(numCols).fill(''));

  // Build date list (28 days from Monday of current week)
  const op     = gxOperationalDate_(new Date());
  const monday = new Date(op.date);
  const diff   = monday.getDay()===0?-6:1-monday.getDay();
  monday.setDate(monday.getDate()+diff);

  for (let c=0;c<numCols;c++) {
    const d = new Date(monday);
    d.setDate(monday.getDate() + colOffset + c);
    headers.push(Utilities.formatDate(d, GX.tz, 'dd/MM'));
    const ymd = Utilities.formatDate(d, GX.tz, 'yyyy-MM-dd');
    const bcName = gxWeeklySheetName_(GX.breakControlPrefix, ymd);
    const bcSheet = ss.getSheetByName(bcName);
    if (!bcSheet || bcSheet.getLastRow() < 2) continue;
    const rows = bcSheet.getRange(2,1,bcSheet.getLastRow()-1,Math.max(12,bcSheet.getLastColumn())).getValues();
    rows.forEach(r => {
      const email   = String(r[7]||'').trim().toLowerCase();
      const rowDate = r[8] instanceof Date ? Utilities.formatDate(r[8],GX.tz,'yyyy-MM-dd') : String(r[8]||'').trim();
      if (!email || rowDate !== ymd || !r[4]) return;
      const agentName = String(r[1]||'').trim();
      const plan      = String(r[3]||'');
      const addBreak  = (val, label, dur) => {
        if (!val || String(val).includes('N/A')) return;
        const dt = gxParseDate_(val);
        if (!dt) return;
        const h = Number(Utilities.formatDate(dt, GX.tz, 'H'));
        if (h < 0 || h > 23) return;
        rowData[h][c]  += dur;
        rowNotes[h][c] += (rowNotes[h][c]?'\n':'') + agentName + ': ' + label;
      };
      if (plan.includes('T2')) {
        addBreak(r[4],'T2-30B1',30); addBreak(r[5],'T2-15B1',15); addBreak(r[6],'T2-15B2',15);
      } else {
        addBreak(r[4],'T1-30B1',30); addBreak(r[5],'T1-30B2',30);
      }
    });
  }

  const gridRows = rowData.map((cells,h) => ({
    label: pad_(h)+':00',
    cells: cells.map((v,c) => ({ value:v, display:v>0?String(v):'', note:rowNotes[h][c] }))
  }));
  return { headers, rows:gridRows };
}

/* ── HEADCOUNT CALCULATION ── */
function gxCalculateHeadcount_(options) {
  options = options || {};
  const ss      = SpreadsheetApp.getActiveSpreadsheet();
  const rObj    = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const roster  = rObj.sheet;
  if (!roster || roster.getLastRow() < 3) return { message:'No roster data', grid:[] };

  const dates      = roster.getRange(1,5,1,7).getDisplayValues()[0];
  const dateValues = roster.getRange(1,5,1,7).getValues()[0];
  const rows       = roster.getRange(3,1,roster.getLastRow()-2, Math.min(30,roster.getLastColumn())).getDisplayValues();
  const exceptions = gxGetExceptionLogs_(2000);
  const grid  = Array.from({length:7}, () => Array(24).fill(0));
  const notes = Array.from({length:7}, () => Array.from({length:24}, () => []));

  rows.forEach(agent => {
    const id   = String(agent[GX.cols.id]  ||'').trim();
    const name = gxCleanName_(agent[GX.cols.name]);
    if (!id || !name || id === 'ID') return;

    const start   = gxParseTime_(agent[GX.cols.utcStart]||agent[GX.cols.shiftStart]);
    const end     = gxParseTime_(agent[GX.cols.utcEnd]  ||agent[GX.cols.shiftEnd]);
    if (start.h < 0 || end.h < 0) return;
    const startDec = start.decimal;
    const duration = end.decimal > startDec ? end.decimal - startDec : 24 - startDec + end.decimal;

    for (let d = 0; d < 7; d++) {
      const status = String(agent[GX.cols.dayStart+d]||'').trim().toLowerCase();
      if (!(status==='' || status==='ed')) continue;
      const targetDate = dateValues[d];
      const targetText = dates[d];
      const dayExc = exceptions.filter(e => e.agentId===id && (gxSameYmd_(targetDate,e.date)||e.date===targetText));
      if (dayExc.some(e => e.type==='Absent' && e.status!=='Recovered')) continue;

      const late  = dayExc.find(e => e.type==='Late Arrival');
      const early = dayExc.find(e => e.type==='Early Departure');
      const ot    = dayExc.find(e => e.type==='Overtime' && e.status==='Approved');
      let actualStartOff = 0, actualEndOff = duration;

      if (late && late.timeValue) {
        const t = gxParseTime_(late.timeValue);
        let logUtc = ((t.h-5+24)%24) + t.m/60;
        let off = logUtc - startDec;
        if (off < -12) off += 24; if (off > 12) off -= 24;
        if (off > 0) actualStartOff = Math.min(off, duration);
      }
      if (early && early.timeValue) {
        const t = gxParseTime_(early.timeValue);
        let logUtc = ((t.h-5+24)%24) + t.m/60;
        let off = logUtc - startDec;
        if (off < -12) off += 24; if (off > 12) off -= 24;
        if (off > 0) actualEndOff = Math.min(off, duration);
      }

      for (let i = 0; i < duration; i++) {
        const worked = Math.max(0, Math.min(actualEndOff,i+1) - Math.max(actualStartOff,i));
        if (worked <= 0) continue;
        const hour      = (Math.floor(startDec)+i) % 24;
        const dayOffset = Math.floor((Math.floor(startDec)+i)/24);
        const rawDay    = d + dayOffset;
        if (rawDay >= 0 && rawDay <= 6) { grid[rawDay][hour] += worked; notes[rawDay][hour].push(name); }
      }

      if (ot && ot.timeValue) {
        const parts = ot.timeValue.split('-').map(s=>s.trim());
        if (parts.length === 2) {
          const otS = gxParseTime_(parts[0]), otE = gxParseTime_(parts[1]);
          if (otS.h >= 0 && otE.h >= 0) {
            const otDur = otE.decimal > otS.decimal ? otE.decimal - otS.decimal : 24 - otS.decimal + otE.decimal;
            for (let i = 0; i < Math.ceil(otDur); i++) {
              const worked = Math.min(1, otDur - i);
              const hour   = (Math.floor(otS.decimal)+i)%24;
              if (hour >= 0 && hour < 24) { grid[d][hour] += worked; notes[d][hour].push(name+' (OT)'); }
            }
          }
        }
      }
    }
  });

  if (options.writeSheet) {
    const op    = gxOperationalDate_(new Date());
    const sName = gxWeeklySheetName_(GX.breakSlotsPrefix, op.ymd);
    let sSheet  = ss.getSheetByName(sName);
    if (!sSheet) { sSheet = ss.insertSheet(sName); gxInitBreakSlotsSheet_(sSheet); }
    const weekOffset = gxCurrentWeekOffset_(sSheet, op.ymd);
    const startCol   = 2 + weekOffset;
    for (let h = 0; h < 24; h++) {
      for (let d = 0; d < 7; d++) {
        const row = 134 + h;
        const col = startCol + d;
        if (col <= sSheet.getLastColumn()) {
          sSheet.getRange(row, col).setValue(Number(grid[d][h].toFixed(2)));
          const noteText = notes[d][h].length ? notes[d][h].join('\n') : '';
          if (noteText) sSheet.getRange(row,col).setNote(noteText);
        }
      }
    }
    gxLog_('HEADCOUNT_CALCULATED','','Written to '+sName,'SYSTEM');
  }

  const packed = [];
  for (let d=0;d<7;d++) for (let h=0;h<24;h++)
    packed.push({ dayIndex:d, day:dates[d], hour:h, value:Number(grid[d][h].toFixed(2)), agents:notes[d][h] });
  return { message:'Headcount calculated', grid:packed, rosterName:rObj.name, generatedAt:gxNowStamp_() };
}

function gxCurrentWeekOffset_(sheet, opYmd) {
  try {
    const headerRow = sheet.getRange(1,2,1,28).getDisplayValues()[0];
    for (let w=0;w<4;w++) {
      const dateStr = headerRow[w*7];
      if (dateStr && gxSameYmd_(gxMondayOfWeek_(opYmd), dateStr)) return w*7;
    }
  } catch(e) {}
  return 0;
}

function gxMondayOfWeek_(ymd) {
  const d = new Date(ymd+'T12:00:00');
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1-day;
  d.setDate(d.getDate()+diff);
  return Utilities.formatDate(d, GX.tz, 'yyyy-MM-dd');
}

/* ═══════════════════════════════════════
   BALANCING ENGINE
═══════════════════════════════════════ */
function gxGetBalancingData_(payload) {
  payload = payload || {};
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const op     = gxOperationalDate_(new Date());
  const opYmd  = payload.date || op.ymd;
  const live   = gxGetLiveMonitorData_({ autoSync:false, writeSnapshot:false });
  const breakMap = gxBuildBreakMap_(ss, opYmd);
  const hours  = Array.from({length:24}, (_,h) => ({ hour:h, planned:0, actual:0, breakLoad:0, available:0, deficit:0 }));

  live.rows.forEach(r => {
    const start = new Date(r.shiftStartMs), end = new Date(r.shiftEndMs);
    for (let h=0;h<24;h++) {
      const bkt = gxBuildDateAt_(opYmd, pad_(h)+':00');
      const bktEnd = new Date(bkt.getTime()+60*60000);
      if (bkt < end && bktEnd > start) {
        hours[h].planned++;
        if (['ONLINE','ACTIVE','IDLE','BREAK','OVERTIME'].includes(r.status)) hours[h].actual++;
      }
    }
  });

  Object.values(breakMap).forEach(pack => {
    pack.breaks.forEach(b => {
      const h = Number(Utilities.formatDate(new Date(b.startMs),GX.tz,'H'));
      if (hours[h]) hours[h].breakLoad++;
    });
  });

  hours.forEach(h => { h.available = Math.max(0,h.actual-h.breakLoad); h.deficit = Math.max(0,h.planned-h.available); });

  const deficits = hours.filter(h=>h.deficit>0).map(h => ({
    hour:h.hour, from:pad_(h.hour)+':00', to:pad_((h.hour+1)%24)+':00',
    deficit:h.deficit, severity: h.deficit>=3?'CRITICAL':h.deficit===2?'HIGH':'MEDIUM',
    planned:h.planned, available:h.available, breakLoad:h.breakLoad
  }));

  const suggestions = gxBuildSuggestions_(deficits, hours, breakMap);
  return { hours, deficits, suggestions, generatedAt:gxNowStamp_() };
}

function gxBuildSuggestions_(deficits, hours, breakMap) {
  const sugs = [];
  deficits.forEach(def => {
    const targets = hours
      .filter(h => h.hour !== def.hour && h.deficit === 0 && h.breakLoad <= 1 && h.planned > 0)
      .sort((a,b) => Math.abs(a.hour-def.hour)-Math.abs(b.hour-def.hour));
    if (!targets.length) return;
    const targetHour = targets[0];
    Object.keys(breakMap).some(email => {
      const pack = breakMap[email];
      const br   = pack.breaks.find(b => Number(Utilities.formatDate(new Date(b.startMs),GX.tz,'H'))===def.hour);
      if (!br) return false;
      sugs.push({
        email, rowNumber:pack.rowNumber, breakLabel:br.label,
        breakKey: br.label.includes('B2')?'b2':br.label.includes('B3')?'b3':'b1',
        fromTime:br.start, toTime:pad_(targetHour.hour)+':00',
        impact:'+'+Math.min(def.deficit,1)+' capacity restored in '+def.from,
        severity:def.severity, reason:'Move break from deficit hour '+def.from+' to '+pad_(targetHour.hour)+':00'
      });
      return true;
    });
  });
  return sugs.slice(0,12);
}

function gxApplyBalancingSuggestion_(payload) {
  if (!payload || !payload.rowNumber || !payload.breakKey || !payload.toTime) throw new Error('Missing suggestion data.');
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const op     = gxOperationalDate_(new Date());
  const opYmd  = payload.date || op.ymd;
  const name   = gxWeeklySheetName_(GX.breakControlPrefix, opYmd);
  const sheet  = ss.getSheetByName(name);
  if (!sheet) throw new Error('Break Control sheet missing: '+name);
  const col   = payload.breakKey==='b2'?6:payload.breakKey==='b3'?7:5;
  sheet.getRange(payload.rowNumber,col).setValue(gxBreakStamp_(opYmd,payload.toTime));
  sheet.getRange(payload.rowNumber,col).setNote('Moved by Balancing Engine '+gxNowStamp_()+'\n'+payload.reason);
  gxLog_('BREAK_MOVED_BALANCING','',payload.fromTime+'→'+payload.toTime,gxGetUserEmail_());
  return { message:'Break moved to '+payload.toTime };
}

/* ═══════════════════════════════════════
   BREAK SUBMISSION
═══════════════════════════════════════ */
function gxSubmitBreak_(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const opYmd = payload.shiftDate || gxOperationalDate_(new Date()).ymd;
  const name  = gxWeeklySheetName_(GX.breakControlPrefix, opYmd);
  const sheet = gxEnsureSheet_(ss, name, ['Submitted At','Agent','Shift','Plan','Break 1','Break 2','Break 3','Email','Shift Date','Approval','Sync','Submitted By']);
  const email = String(payload.email||'').trim().toLowerCase();
  if (!email) throw new Error('Missing agent email.');
  if (!payload.b1||!payload.b2) throw new Error('Break 1 and 2 are required.');

  const vals = sheet.getLastRow()>1 ? sheet.getRange(2,1,sheet.getLastRow()-1,12).getDisplayValues() : [];
  const dupe = vals.some((r,i)=>i>0 && String(r[7]).trim().toLowerCase()===email && String(r[8]).trim()===opYmd && r[4]);
  if (dupe && !payload.allowDuplicate) throw new Error('Breaks already submitted for this shift.');

  const plan = String(payload.plan||'').includes('T2') ? 'T2 Breaks' : 'T1 Breaks';
  const b1   = gxBreakStamp_(opYmd, payload.b1);
  const b2   = gxBreakStamp_(opYmd, payload.b2);
  const b3   = plan==='T2 Breaks' ? gxBreakStamp_(opYmd, payload.b3) : 'N/A - T1 Break';
  const by   = payload.submittedBy || gxGetUserEmail_() || 'Genesys +';
  sheet.appendRow([gxNowStamp_(), payload.name||email, "'"+String(payload.shift||''), plan, b1, b2, b3, email, opYmd, 'Approved', '', by]);

  if (GX.CALENDAR_ID) {
    try {
      const cal = CalendarApp.getCalendarById(GX.CALENDAR_ID);
      if (cal) {
        [[payload.b1,30],[payload.b2,plan==='T2 Breaks'?15:30],[payload.b3,15]].forEach(([t,d],i) => {
          if (!t||t==='N/A') return;
          const start = gxBuildDateAt_(opYmd,t);
          const end   = new Date(start.getTime()+d*60000);
          cal.createEvent((payload.name||email)+' — '+(plan==='T2 Breaks'?['T2-30B1','T2-15B1','T2-15B2'][i]:'T1-30B'+(i+1)), start, end);
        });
      }
    } catch(e) {}
  }

  gxUpdateBreakAudit_(ss, opYmd, email, payload.name, payload.b1, payload.b2, payload.b3, plan);
  gxLog_('BREAK_SUBMITTED', payload.agentId||email, plan+' '+[payload.b1,payload.b2,payload.b3||''].filter(Boolean).join(','), by);
  return { message:'Break plan submitted', sheetName:name };
}

function gxUpdateBreakAudit_(ss, opYmd, email, name, b1, b2, b3, plan) {
  const sName  = gxWeeklySheetName_(GX.breakSlotsPrefix, opYmd);
  const sSheet = ss.getSheetByName(sName);
  if (!sSheet) return;
  const weekOffset = gxCurrentWeekOffset_(sSheet, opYmd);
  const d = new Date(opYmd+'T12:00:00').getDay();
  const dayOff = d === 0 ? 6 : d-1;
  const col = 2 + weekOffset + dayOff;
  [[b1,'T2-30B1'],[b2,'T2-15B1'],[b3,'T2-15B2']].forEach(([t,lbl]) => {
    if (!t||t==='N/A') return;
    const parsed = gxParseDate_(t);
    if (!parsed) return;
    const h = Number(Utilities.formatDate(parsed,GX.tz,'H'));
    const row = 30 + h;
    if (row > 52 || row < 30) return;
    const cell    = sSheet.getRange(row,col);
    const curVal  = Number(cell.getValue())||0;
    cell.setValue(curVal+1);
    const curNote = cell.getNote()||'';
    cell.setNote(curNote + (curNote?'\n':'') + name+': '+lbl);
  });
}

/* ── BREAK CALENDAR ── */
function gxGetBreakCalendar_(payload) {
  payload = payload || {};
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const op     = gxOperationalDate_(new Date());
  const opYmd  = payload.date || op.ymd;
  const name   = gxWeeklySheetName_(GX.breakControlPrefix, opYmd);
  const sheet  = ss.getSheetByName(name);
  const agents = gxGetAgents_(ss, gxGetAccessContext_(ss, payload.mode||''));
  const byEmail = {};
  agents.forEach(a => byEmail[a.email] = a);
  const events = [], submitted = {};

  if (sheet && sheet.getLastRow() > 1) {
    const values = sheet.getRange(2,1,sheet.getLastRow()-1,Math.max(12,sheet.getLastColumn())).getValues();
    values.forEach((row,idx) => {
      const email   = String(row[7]||'').trim().toLowerCase();
      const rowDate = row[8] instanceof Date ? Utilities.formatDate(row[8],GX.tz,'yyyy-MM-dd') : String(row[8]||'').trim();
      if (!email||rowDate!==opYmd||!row[4]) return;
      submitted[email] = true;
      const agent = byEmail[email] || { name:row[1], id:'', email };
      const plan  = String(row[3]||'');
      gxPushCalEvent_(events, row[4], agent, idx+2, name, 'b1', plan.includes('T2')?'T2-30B1':'T1-30B1', 30, opYmd);
      gxPushCalEvent_(events, row[5], agent, idx+2, name, 'b2', plan.includes('T2')?'T2-15B1':'T1-30B2', plan.includes('T2')?15:30, opYmd);
      if (plan.includes('T2')) gxPushCalEvent_(events, row[6], agent, idx+2, name, 'b3', 'T2-15B2', 15, opYmd);
    });
  }

  const unscheduled = agents.filter(a => a.email && !submitted[a.email]).map(a => ({
    id:a.id, name:a.name, email:a.email, shift:a.shiftStart+' - '+a.shiftEnd
  }));

  events.sort((a,b) => a.startMs-b.startMs);
  return { events, unscheduled, sheetName:name, opDate:opYmd };
}

function gxPushCalEvent_(events, value, agent, rowNumber, sheetName, breakKey, label, duration, opYmd) {
  if (!value||String(value).includes('N/A')) return;
  const dt = gxParseDate_(value);
  if (!dt) return;
  const h    = Number(Utilities.formatDate(dt,GX.tz,'H'));
  const m    = Number(Utilities.formatDate(dt,GX.tz,'m'));
  const CAL_START_H = 5;
  let offset = (h - CAL_START_H)*60 + m;
  const evYmd = Utilities.formatDate(dt,GX.tz,'yyyy-MM-dd');
  if (evYmd > opYmd) offset += 24*60;
  if (offset < 0)    offset += 24*60;
  events.push({
    id:agent.email+'-'+breakKey, agentId:agent.id, name:agent.name, email:agent.email,
    rowNumber, sheetName, breakKey, label, time:pad_(h)+':'+pad_(m),
    startMs:dt.getTime(), duration, offsetMins:offset
  });
}

function gxMoveBreak_(payload) {
  if (!payload||!payload.sheetName||!payload.rowNumber||!payload.breakKey||!payload.toTime)
    throw new Error('Missing move payload.');
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(payload.sheetName);
  if (!sheet) throw new Error('Break sheet missing: '+payload.sheetName);
  const col   = payload.breakKey==='b2'?6:payload.breakKey==='b3'?7:5;
  const opYmd = payload.date || gxOperationalDate_(new Date()).ymd;
  const oldVal = sheet.getRange(payload.rowNumber,col).getDisplayValue();
  sheet.getRange(payload.rowNumber,col).setValue(gxBreakStamp_(opYmd,payload.toTime));
  sheet.getRange(payload.rowNumber,col).setNote('Moved in calendar by '+gxGetUserEmail_()+'\nFrom: '+oldVal+'\nAt: '+gxNowStamp_());
  gxLog_('BREAK_MOVED_CALENDAR',payload.agentId||'',oldVal+' → '+payload.toTime,gxGetUserEmail_());
  return { message:'Break moved to '+payload.toTime };
}

/* ═══════════════════════════════════════
   PERFORMANCE
═══════════════════════════════════════ */
function gxGetPerformance_(payload) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(GX.sheets.performance) || ss.getSheetByName('Agent Perfomance');
  if (!sheet||sheet.getLastRow()<2) return { rows:[] };
  const vals = sheet.getDataRange().getDisplayValues();
  const rows = [];
  for (let i=1;i<vals.length;i++) {
    rows.push({
      week:vals[i][0], name:vals[i][1]||vals[i][11], shift:vals[i][2],
      art:vals[i][3], efficiency:vals[i][4], escalations:vals[i][5],
      aht:vals[i][6], nps:vals[i][7], quality:vals[i][8],
      audits:vals[i][9], activeTime:vals[i][10], email:vals[i][11]
    });
  }
  return { rows:rows.slice(-200).reverse() };
}

/* ═══════════════════════════════════════
   LOGS
═══════════════════════════════════════ */
function gxGetActivityLogs_(limit) {
  const sheet = gxEnsureSheet_(SpreadsheetApp.getActiveSpreadsheet(), GX.sheets.activity,
    ['Timestamp','Event','Agent ID','Details','By']);
  if (sheet.getLastRow()<2) return [];
  const last  = sheet.getLastRow();
  const count = Math.min(limit||200,last-1);
  const start = Math.max(2,last-count+1);
  return sheet.getRange(start,1,count,5).getDisplayValues().reverse().map(r => ({
    timestamp:r[0],event:r[1],agentId:r[2],details:r[3],by:r[4]
  }));
}

/* ═══════════════════════════════════════
   ACCESS CONTROL
═══════════════════════════════════════ */
function gxGetAccessRows_() {
  const sheet = gxEnsureSheet_(SpreadsheetApp.getActiveSpreadsheet(), GX.sheets.access, ['Email','Role','Name','Team']);
  if (sheet.getLastRow()<2) return { rows:[] };
  const vals = sheet.getRange(2,1,sheet.getLastRow()-1,4).getDisplayValues();
  return { rows:vals.map((r,i)=>({ rowNumber:i+2,email:r[0],role:r[1],name:r[2],team:r[3] })) };
}

function gxSaveAccessRow_(payload) {
  if (!payload.email||!payload.role) throw new Error('Email and role required.');
  const sheet = gxEnsureSheet_(SpreadsheetApp.getActiveSpreadsheet(), GX.sheets.access, ['Email','Role','Name','Team']);
  const email = String(payload.email).trim().toLowerCase();
  const vals  = sheet.getLastRow()>1 ? sheet.getRange(2,1,sheet.getLastRow()-1,4).getDisplayValues() : [];
  let rowNum  = 0;
  vals.some((r,i)=>{ if (String(r[0]).trim().toLowerCase()===email){ rowNum=i+2; return true; } return false; });
  const roleLabel = gxRoleLabel_(gxNormalizeRole_(payload.role));
  const row = [[email, roleLabel, payload.name||'', payload.team||'']];
  if (rowNum) sheet.getRange(rowNum,1,1,4).setValues(row);
  else        sheet.getRange(sheet.getLastRow()+1,1,1,4).setValues(row);
  gxLog_('ACCESS_SAVED','',email+' → '+roleLabel,gxGetUserEmail_());
  return { message:'Access saved for '+email };
}

function gxRevokeAccessRow_(payload) {
  if (!payload.email) throw new Error('Email required.');
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const sht = ss.getSheetByName(GX.sheets.access);
  if (!sht) return { message:'Sheet missing.' };
  const vals  = sht.getRange(2,1,sht.getLastRow()-1,1).getDisplayValues();
  const email = String(payload.email).trim().toLowerCase();
  for (let i=0;i<vals.length;i++) {
    if (String(vals[i][0]).trim().toLowerCase()===email) {
      sht.deleteRow(i+2);
      gxLog_('ACCESS_REVOKED','',email,gxGetUserEmail_());
      return { message:'Access revoked for '+email };
    }
  }
  return { message:'Email not found in Access_Control.' };
}

/* ═══════════════════════════════════════
   REORGANISER
═══════════════════════════════════════ */
function gxRunReorganiser_() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const rObj   = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const roster = rObj.sheet;
  if (!roster) throw new Error('Missing roster: '+rObj.name);
  const agents   = gxGetAgents_(ss, {});
  const template = ss.getSheetByName('Agent Template') || ss.getSheets()[0];
  let created=0, updated=0, skipped=0;
  agents.forEach(agent => {
    if (!agent.linkedSheetId || !agent.email) { skipped++; return; }
    try {
      SpreadsheetApp.openById(agent.linkedSheetId);
      updated++;
    } catch(e) {
      try {
        const newSheet = ss.insertSheet(agent.name, { template });
        newSheet.setName(agent.name + ' — ' + agent.id);
        roster.getRange(agent.row, 18).setValue(ss.getId());
        created++;
        gxLog_('REORGANISER_CREATED_SHEET',agent.id,agent.name,gxGetUserEmail_());
      } catch(e2) { skipped++; }
    }
  });
  gxLog_('REORGANISER_RUN','','Created:'+created+' Updated:'+updated+' Skipped:'+skipped,gxGetUserEmail_());
  return { message:`Reorganiser complete. Created: ${created}, Updated: ${updated}, Skipped: ${skipped}.` };
}

/* ═══════════════════════════════════════
   ADMIN
═══════════════════════════════════════ */
function gxAdminInitialize_() {
  gxEnsureCoreSheets_(SpreadsheetApp.getActiveSpreadsheet());
  return { message:'Genesys + core sheets initialized.' };
}

function gxSetupTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => { if (t.getHandlerFunction()==='gxScheduledRefresh') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('gxScheduledRefresh').timeBased().everyMinutes(5).create();
  return { message:'5-minute refresh trigger created.' };
}

function gxOverrideAgentState_(payload) {
  gxLog_('STATE_OVERRIDE',payload.agentId||'','Forced: '+payload.state,gxGetUserEmail_());
  return { message:'State override logged for '+(payload.agentId||'unknown') };
}

/* ═══════════════════════════════════════
   MOCK DATA
═══════════════════════════════════════ */
function gxSetupMockData() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  gxEnsureCoreSheets_(ss);
  const userEmail = gxGetUserEmail_() || 'farhan@example.com';
  const op        = gxOperationalDate_(new Date());
  const ssId      = ss.getId();

  const rosterName = gxWeeklySheetName_(GX.staffingPrefix, op.ymd);
  let roster = ss.getSheetByName(rosterName) || ss.insertSheet(rosterName);
  roster.clear();

  const monday = new Date(op.date);
  const diff   = monday.getDay()===0 ? -6 : 1-monday.getDay();
  monday.setDate(monday.getDate()+diff);
  const header = Array(30).fill('');
  header[0]='Shift Start';header[1]='Shift End';header[2]='ID';header[3]='Agent Name';
  for (let i=0;i<7;i++){ const d=new Date(monday);d.setDate(monday.getDate()+i);header[4+i]=d; }
  header[11]='Role';header[12]='Team Lead';header[14]='UTC Start';header[15]='UTC End';
  header[16]='Email';header[17]='Linked Sheet ID';
  roster.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold');
  roster.getRange(2,1,1,header.length).setValues([Array(header.length).fill('')]);

  const mock = [
    ['09:00','18:00','GX001','Farhan Naz - TC','','','','','','','','TC',   'Team Alpha','','04:00','13:00',userEmail,ssId],
    ['09:00','18:00','GX002','Alexandra Chen', '','','','','','','','Chat', 'Team Alpha','','04:00','13:00','alexandra.chen@example.com',ssId],
    ['08:00','17:00','GX003','Marcus Johnson', '','','','','','','','Chat', 'Team Alpha','','03:00','12:00','marcus.johnson@example.com',ssId],
    ['10:00','19:00','GX004','Priya Sharma',   '','','','','','','','Chat', 'Team Beta', '','05:00','14:00','priya.sharma@example.com',ssId],
    ['09:00','18:00','GX005','David Okafor',   '','','','','','','','Chat', 'Team Beta', '','04:00','13:00','david.okafor@example.com',ssId],
    ['12:00','21:00','GX006','James Wilson',   '','','','','','','','Chat', 'Team Gamma','','07:00','16:00','james.wilson@example.com',ssId],
    ['11:00','20:00','GX007','Aisha Nakamura', '','','','','','','','Chat', 'Team Gamma','','06:00','15:00','aisha.nakamura@example.com',ssId],
    ['08:00','17:00','GX008','Carlos Mendes',  '','','','','','','','Chat', 'Team Beta', '','03:00','12:00','carlos.mendes@example.com',ssId],
    ['09:00','18:00','GX009','Sophie Laurent', '','','','','','','','Chat', 'Team Alpha','','04:00','13:00','sophie.laurent@example.com',ssId],
    ['10:00','19:00','GX010','Kwame Asante',   '','','','','','','','Chat', 'Team Gamma','','05:00','14:00','kwame.asante@example.com',ssId]
  ];
  roster.getRange(3,1,mock.length,mock[0].length).setValues(mock);
  roster.setFrozenRows(2);

  const access = gxEnsureSheet_(ss, GX.sheets.access, ['Email','Role','Name','Team']);
  access.clear();
  access.getRange(1,1,1,4).setValues([['Email','Role','Name','Team']]).setFontWeight('bold');
  access.getRange(2,1,4,4).setValues([
    [userEmail,             'Admin',            'Farhan Naz',     ''],
    ['tc.alpha@example.com','Team Coordinator', 'Alpha TC',       'Team Alpha'],
    ['manager@example.com', 'Manager',          'Ops Manager',    ''],
    ['admin@example.com',   'Admin',            'Admin User',     '']
  ]);

  const sys = gxEnsureSheet_(ss, GX.sheets.sysActivity, ['Agent Ext','ActivityStamp']);
  sys.clear();
  sys.getRange(1,1,1,2).setValues([['Agent Ext','ActivityStamp']]).setFontWeight('bold');
  const now = new Date();
  sys.getRange(2,1,8,2).setValues([
    [userEmail,                      new Date(now.getTime()-5*60000)],
    ['alexandra.chen@example.com',   new Date(now.getTime()-2*60000)],
    ['marcus.johnson@example.com',   new Date(now.getTime()-8*60000)],
    ['david.okafor@example.com',     new Date(now.getTime()-12*60000)],
    ['james.wilson@example.com',     new Date(now.getTime()-18*60000)],
    ['aisha.nakamura@example.com',   new Date(now.getTime()-1*60000)],
    ['sophie.laurent@example.com',   new Date(now.getTime()-3*60000)],
    ['kwame.asante@example.com',     new Date(now.getTime()-4*60000)]
  ]);

  const bcName = gxWeeklySheetName_(GX.breakControlPrefix, op.ymd);
  const bc = gxEnsureSheet_(ss,bcName,['Submitted At','Agent','Shift','Plan','Break 1','Break 2','Break 3','Email','Shift Date','Approval','Sync','Submitted By']);
  bc.clear();
  bc.getRange(1,1,1,12).setValues([['Submitted At','Agent','Shift','Plan','Break 1','Break 2','Break 3','Email','Shift Date','Approval','Sync','Submitted By']]).setFontWeight('bold');
  bc.getRange(2,1,6,12).setValues([
    [gxNowStamp_(),'Farhan Naz',    "'09:00 - 18:00",'T1 Breaks',gxBreakStamp_(op.ymd,'11:00'),gxBreakStamp_(op.ymd,'14:30'),'N/A - T1 Break',userEmail,op.ymd,'Approved','','Mock'],
    [gxNowStamp_(),'Alexandra Chen',"'09:00 - 18:00",'T1 Breaks',gxBreakStamp_(op.ymd,'11:00'),gxBreakStamp_(op.ymd,'14:30'),'N/A - T1 Break','alexandra.chen@example.com',op.ymd,'Approved','','Mock'],
    [gxNowStamp_(),'Marcus Johnson',"'08:00 - 17:00",'T2 Breaks',gxBreakStamp_(op.ymd,'10:30'),gxBreakStamp_(op.ymd,'13:00'),gxBreakStamp_(op.ymd,'15:00'),'marcus.johnson@example.com',op.ymd,'Approved','','Mock'],
    [gxNowStamp_(),'David Okafor',  "'09:00 - 18:00",'T1 Breaks',gxBreakStamp_(op.ymd,'10:00'),gxBreakStamp_(op.ymd,'15:30'),'N/A - T1 Break','david.okafor@example.com',op.ymd,'Approved','','Mock'],
    [gxNowStamp_(),'Aisha Nakamura',"'11:00 - 20:00",'T2 Breaks',gxBreakStamp_(op.ymd,'12:30'),gxBreakStamp_(op.ymd,'15:00'),gxBreakStamp_(op.ymd,'17:00'),'aisha.nakamura@example.com',op.ymd,'Approved','','Mock'],
    [gxNowStamp_(),'Kwame Asante',  "'10:00 - 19:00",'T1 Breaks',gxBreakStamp_(op.ymd,'11:00'),gxBreakStamp_(op.ymd,'16:00'),'N/A - T1 Break','kwame.asante@example.com',op.ymd,'Approved','','Mock']
  ]);

  const perf = gxEnsureSheet_(ss,GX.sheets.performance,['Week','Name','Shift','ART','Efficiency','Escalations','AHT','NPS','Quality','Audits','Active Time','Email']);
  perf.clear();
  perf.getRange(1,1,1,12).setValues([['Week','Name','Shift','ART','Efficiency','Escalations','AHT','NPS','Quality','Audits','Active Time','Email']]).setFontWeight('bold');
  const perfRows = [];
  mock.forEach((r,idx) => {
    for (let w=0;w<4;w++) {
      perfRows.push(['2026-W'+(20-w),gxCleanName_(r[3]),r[0]+' - '+r[1],
        (2.3+idx*.04).toFixed(2),(88+(idx%7)).toFixed(2)+'%',idx%4,
        (12+(idx%5)).toFixed(2),-20-idx,(90+(idx%9)).toFixed(2)+'%',
        5+idx,(92+(idx%6)).toFixed(2)+'%',r[16]]);
    }
  });
  perf.getRange(2,1,perfRows.length,12).setValues(perfRows);

  const ex = gxEnsureExceptionSheet_(ss);
  ex.clear();
  ex.getRange(1,1,1,13).setValues([['Timestamp','AgentID','AgentName','Team Lead','Day','Date','Shift','Type','Time Value','Duration','Source','Status','Notes']]).setFontWeight('bold');
  gxAppendException_(ss,{agentId:'GX008',agentName:'Carlos Mendes',teamLead:'Team Beta',date:op.ymd,shift:'08:00 - 17:00',type:'Late Arrival',timeValue:'08:44',endTime:'',duration:'00:44',source:'Mock',status:'Open',notes:'Mock late arrival'});
  gxAppendException_(ss,{agentId:'GX006',agentName:'James Wilson',teamLead:'Team Gamma',date:op.ymd,shift:'12:00 - 21:00',type:'Overtime',timeValue:'21:00 - 22:00',endTime:'',duration:'01:00',source:'Mock',status:'Approved',notes:'Mock OT'});

  const slotsName = gxWeeklySheetName_(GX.breakSlotsPrefix, op.ymd);
  let slotsSheet  = ss.getSheetByName(slotsName);
  if (!slotsSheet) { slotsSheet = ss.insertSheet(slotsName); gxInitBreakSlotsSheet_(slotsSheet); }

  gxGetLiveMonitorData_({ autoSync:false, writeSnapshot:true });
  gxLog_('MOCK_DATA_CREATED','SYSTEM','Full mock suite created',userEmail);
  return { message:'Mock test data created. Roster: '+rosterName };
}

function gxInitBreakSlotsSheet_(sheet) {
  const op     = gxOperationalDate_(new Date());
  const monday = new Date(op.date);
  const diff   = monday.getDay()===0?-6:1-monday.getDay();
  monday.setDate(monday.getDate()+diff);
  const headers = ['Hour'];
  for (let i=0;i<28;i++) { const d=new Date(monday); d.setDate(monday.getDate()+i); headers.push(Utilities.formatDate(d,GX.tz,'dd/MM/yyyy')); }
  sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  const grids = [[4,27],[30,52],[108,131],[134,157]];
  grids.forEach(([start,end]) => {
    for (let r=start;r<=end;r++) sheet.getRange(r,1).setValue(pad_(r-start)+':00');
  });
  for (let r=4;r<=27;r++) {
    for (let c=2;c<=29;c++) sheet.getRange(r,c).setValue(180);
  }
}

/* ═══════════════════════════════════════
   SHEET HELPERS
═══════════════════════════════════════ */
function gxEnsureCoreSheets_(ss) {
  gxEnsureExceptionSheet_(ss);
  gxEnsureSheet_(ss, GX.sheets.activity,    ['Timestamp','Event','Agent ID','Details','By']);
  gxEnsureSheet_(ss, GX.sheets.access,      ['Email','Role','Name','Team']);
  gxEnsureSheet_(ss, GX.sheets.liveMonitor, ['Last Updated','Agent','ID','Team','Shift','Pulse','Break','Status','Exception','Last Activity','Risk']);
  gxEnsureSheet_(ss, GX.sheets.performance, ['Week','Name','Shift','ART','Efficiency','Escalations','AHT','NPS','Quality','Audits','Active Time','Email']);
  gxEnsureSheet_(ss, GX.sheets.sysActivity, ['Agent Ext','ActivityStamp']);
  const op = gxOperationalDate_(new Date());
  gxEnsureSheet_(ss, gxWeeklySheetName_(GX.breakControlPrefix,op.ymd),
    ['Submitted At','Agent','Shift','Plan','Break 1','Break 2','Break 3','Email','Shift Date','Approval','Sync','Submitted By']);
}

function gxEnsureExceptionSheet_(ss) {
  return gxEnsureSheet_(ss, GX.sheets.exceptions,
    ['Timestamp','AgentID','AgentName','Team Lead','Day','Date','Shift','Type','Time Value','Duration','Source','Status','Notes']);
}

function gxEnsureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (headers && headers.length) sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  } else if (headers && headers.length && sheet.getLastRow()===0) {
    sheet.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

function gxGetDynamicSheet_(ss, prefix) {
  const op   = gxOperationalDate_(new Date());
  const name = gxWeeklySheetName_(prefix, op.ymd);
  let sheet  = ss.getSheetByName(name);
  if (!sheet) {
    // Fallback: find the most recent sheet whose name starts with prefix
    const all = ss.getSheets();
    let best = null, bestDate = null;
    for (const s of all) {
      const n = s.getName();
      if (!n.startsWith(prefix)) continue;
      // Extract date portion dd.MM.yyyy from end of sheet name
      const m = n.match(/(\d{2})\.(\d{2})\.(\d{4})$/);
      if (!m) continue;
      const d = new Date(Number(m[3]), Number(m[2])-1, Number(m[1]));
      if (!bestDate || d > bestDate) { bestDate = d; best = s; }
    }
    sheet = best;
  }
  return { sheet, name };
}

/** Exposed to frontend so setup wizard can show the expected sheet name */
function gxGetExpectedSheetName() {
  const op = gxOperationalDate_(new Date());
  return gxWeeklySheetName_(GX.staffingPrefix, op.ymd);
}

function gxFindAgentById_(ss, id) {
  const rObj  = gxGetDynamicSheet_(ss, GX.staffingPrefix);
  const sheet = rObj.sheet;
  if (!sheet) return null;
  const rows = sheet.getRange(3,1,sheet.getLastRow()-2,Math.min(20,sheet.getLastColumn())).getDisplayValues();
  for (let i=0;i<rows.length;i++) {
    if (String(rows[i][GX.cols.id]).trim()===String(id).trim()) {
      return {
        row:3+i, id:String(rows[i][GX.cols.id]).trim(),
        name:gxCleanName_(rows[i][GX.cols.name]), team:gxCleanName_(rows[i][GX.cols.teamLead]),
        email:String(rows[i][GX.cols.email]||'').trim().toLowerCase(),
        shiftStart:gxTimeString_(rows[i][GX.cols.shiftStart]),
        shiftEnd:gxTimeString_(rows[i][GX.cols.shiftEnd])
      };
    }
  }
  return null;
}

/* ═══════════════════════════════════════
   DATE / TIME HELPERS
═══════════════════════════════════════ */
function gxOperationalDate_(now) {
  const pkt = new Date(now.getTime()+5*60*60*1000);
  if (pkt.getUTCHours()<5) pkt.setUTCDate(pkt.getUTCDate()-1);
  const ymd   = Utilities.formatDate(pkt,'GMT','yyyy-MM-dd');
  const parts = ymd.split('-').map(Number);
  return { ymd, date:new Date(parts[0],parts[1]-1,parts[2],12,0,0) };
}

function gxDayIndexForRoster_(roster, opYmd) {
  const vals = roster.getRange(1,5,1,7).getValues()[0];
  for (let i=0;i<vals.length;i++) {
    const ymd = vals[i] instanceof Date ? Utilities.formatDate(vals[i],GX.tz,'yyyy-MM-dd') : String(vals[i]);
    if (ymd===opYmd || gxSameYmd_(vals[i],opYmd)) return i;
  }
  const d = new Date(opYmd+'T12:00:00');
  return (d.getDay()+6)%7;
}

function gxBuildShiftWindow_(ymd, startTime, endTime) {
  const start = gxBuildDateAt_(ymd, startTime);
  let   end   = gxBuildDateAt_(ymd, endTime);
  if (end <= start) end = new Date(end.getTime()+24*60*60*1000);
  return { start, end };
}

function gxBuildDateAt_(ymd, hhmm) {
  const p = String(ymd).split('-').map(Number);
  const t = gxParseTime_(hhmm);
  return new Date(p[0],p[1]-1,p[2],Math.max(0,t.h),t.m||0,0);
}

function gxParseTime_(value) {
  if (value===null||value===undefined||value==='') return { h:-1,m:0,decimal:-1 };
  if (value instanceof Date) {
    const h=Number(Utilities.formatDate(value,GX.tz,'H')),m=Number(Utilities.formatDate(value,GX.tz,'m'));
    return { h,m,decimal:h+m/60 };
  }
  let s=String(value).trim().toLowerCase();
  const isPM=s.includes('pm'),isAM=s.includes('am');
  s=s.replace('pm','').replace('am','').trim();
  let h=0,m=0;
  if (s.includes(':')) { const [hs,ms]=s.split(':'); h=parseInt(hs)||0; m=parseInt(ms)||0; }
  else { h=parseInt(s)||0; }
  if (isPM&&h<12) h+=12; if (isAM&&h===12) h=0;
  return { h,m,decimal:h+m/60 };
}

function gxTimeString_(value) {
  const t=gxParseTime_(value); if (t.h<0) return ''; return pad_(t.h)+':'+pad_(t.m);
}

function gxParseDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const s=String(value).trim();
  let m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+(\d{1,2}):(\d{2})/);
  if (m) return new Date(+m[3],+m[2]-1,+m[1],+m[4],+m[5]);
  m=s.match(/^(\d{2})\s+(\w{3})\s+(\d{4}),\s+(\d{2}):(\d{2})/);
  if (m) { const months={'Jan':0,'Feb':1,'Mar':2,'Apr':3,'May':4,'Jun':5,'Jul':6,'Aug':7,'Sep':8,'Oct':9,'Nov':10,'Dec':11}; return new Date(+m[3],months[m[2]],+m[1],+m[4],+m[5]); }
  const d=new Date(s); return isNaN(d.getTime()) ? null : d;
}

function gxBreakStamp_(opYmd, hhmm) {
  if (!hhmm) return '';
  const dt = gxBuildDateAt_(opYmd, hhmm);
  if (gxParseTime_(hhmm).h < 5) dt.setDate(dt.getDate()+1);
  return Utilities.formatDate(dt, GX.tz, 'dd MMM yyyy, HH:mm:ss');
}

function gxDuration_(start, end) {
  if (!start||!end) return '';
  const s=gxParseTime_(start), e=gxParseTime_(end);
  let mins=(e.h*60+e.m)-(s.h*60+s.m);
  if (mins<0) mins+=24*60;
  return pad_(Math.floor(mins/60))+':'+pad_(mins%60);
}

function gxSameYmd_(value, ymd) {
  if (!value||!ymd) return false;
  const d=value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return String(value).includes(String(ymd));
  return Utilities.formatDate(d,GX.tz,'yyyy-MM-dd')===ymd;
}

function gxWeeklySheetName_(prefix, ymd) {
  const date=new Date(ymd+'T12:00:00');
  const day=date.getDay();
  const diff=day===0?-6:1-day;
  date.setDate(date.getDate()+diff);
  return prefix+' '+Utilities.formatDate(date,GX.tz,'dd.MM.yyyy');
}

function gxDayName_(ymd) { return Utilities.formatDate(new Date(ymd+'T12:00:00'),GX.tz,'EEEE'); }
function gxCleanName_(text) { if (!text) return ''; return String(text).replace(/[\s\u00A0]*[-–—].*$/i,'').trim(); }
function gxTitleCase_(text) { return String(text||'').replace(/\b\w/g,c=>c.toUpperCase()); }
function gxNowStamp_() { return Utilities.formatDate(new Date(),GX.tz,'dd MMM yyyy HH:mm:ss'); }
function gxHumanAge_(date, now) {
  const diff=Math.max(0,Math.floor((now.getTime()-date.getTime())/60000));
  if (diff<1)  return 'Just now';
  if (diff<60) return diff+' min ago';
  return Math.floor(diff/60)+' hr ago';
}
function gxGetUserEmail_() {
  try { return String(Session.getActiveUser().getEmail()||'').trim().toLowerCase(); } catch(e) { return ''; }
}

function gxGetDeployerEmail_() {
  try { return String(Session.getEffectiveUser().getEmail()||'').trim().toLowerCase(); } catch(e) { return ''; }
}
function gxLog_(event, agentId, details, by) {
  try {
    const sheet=gxEnsureSheet_(SpreadsheetApp.getActiveSpreadsheet(),GX.sheets.activity,['Timestamp','Event','Agent ID','Details','By']);
    sheet.appendRow([new Date(),event,agentId||'',details||'',by||gxGetUserEmail_()||'SYSTEM']);
  } catch(e) {}
}
function pad_(n) { return String(n).padStart(2,'0'); }
