// Code.gs — OBSI Support Console (STRICT; SSR; Admin+SRP)
// BUILD: ASP-CANON-2025-10-23-EMAILCC-MERGED-v1 (+ EmailToggle, Gmail API send)
//
// Canon rules:
// - All displayed dates use MM/dd/yyyy
// - All displayed times use HH:mm (24-hour)
// - Email bundles: HTML + text; no links; de-duplicated recipients
// - Hard delete for Tickets/ Work Orders on Delete actions
// - Added getActiveEngineers_() helper
// - FIX: requesterEmailFromTicket_ returns RequestedByEmail || ContactEmail

function include(file){ return HtmlService.createHtmlOutputFromFile(file).getContent(); }

// ---------- Canon constants ----------
const FILE_ID='1hBLUJSVeNMjYo5ieYZRlhPUAFdqB7eEPkP5v0HzRlKM';
const BUILD  ='ASP-CANON-2025-10-23-EMAILCC-MERGED-v1';
const ASP_URL='https://script.google.com/macros/s/AKfycbzQPan1Ww69SeyOPWH6CqswfAMLMjEXCIWFHRVIkbwvHpdt7m-D7AHmkPTu2IOLtyS7BQ/exec';

// Email identity (universal)
const SUPPORT_EMAIL = 'support@obsisupport.com';
const SUPPORT_NAME  = 'OBSI Support';

// Sheets
const SH_CLIENTS='Clients',SH_MARKETS='Markets',SH_CALLSIGNS='CallSigns',SH_ENG='Engineers',
      SH_TKT='SupportTickets',SH_TREC='SupportRecords',SH_WO='WorkOrders',SH_WREC='WorkOrderRecords',
      SH_CONTACTS='Contacts',SH_CLIENTCONTACTS='ClientContacts',SH_EMAILCC='EmailCC';

// ===== Email Controls (toggle only; no quota display) =====
const EMAIL_TOGGLE_PROP = 'ASP_EmailEnabled';
const EMAIL_DEFAULT_ENABLED = true;

// ---------- Utils ----------
const ss_=()=>SpreadsheetApp.openById(FILE_ID);
const sh_ = n=>{ const s=ss_().getSheetByName(n); if(!s) throw new Error('Missing sheet: '+n); return s; };
const hdr_= s=>s.getRange(1,1,1,Math.max(1,s.getLastColumn())).getValues()[0].map(h=>String(h||'').trim());

function tryRead_(name){ try{ return read_(name); }catch(_){ return []; } }
function tryGetSheet_(name){ try{ return sh_(name); }catch(_){ return null; } }

function read_(name){
  const s=sh_(name),v=s.getDataRange().getValues(); if(v.length<2) return [];
  const h=v[0], out=[];
  for(let r=1;r<v.length;r++){
    const row=v[r]; if(row.every(c=>c===''||c===null)) continue;
    const x={}; for(let c=0;c<h.length;c++){ if(h[c]) x[h[c]]=row[c]; }
    out.push(x);
  }
  return out;
}
function writeByKey_(sheet,key,val,updates){
  const s=sh_(sheet),vals=s.getDataRange().getValues(); if(!vals.length) throw new Error('No data in '+sheet);
  const h=vals[0].map(x=>String(x||'').trim()), k=h.indexOf(key); if(k<0) throw new Error('Missing key '+key+' in '+sheet);
  for(let r=1;r<vals.length;r++){
    if(String(vals[r][k])===String(val)){
      Object.entries(updates).forEach(([K,V])=>{
        const c=h.indexOf(K);
        if(c>=0) s.getRange(r+1,c+1).setValue(V);
      });
      return;
    }
  }
  throw new Error(sheet+': no row where '+key+'=='+val);
}
function appendByHdr_(sheet,obj){
  const s=sh_(sheet),h=hdr_(s);
  s.appendRow(h.map(H=>Object.prototype.hasOwnProperty.call(obj,H)?obj[H]:'')); 
}

const now_ = ()=>new Date();
const tz_  = ()=>Session.getScriptTimeZone()||'America/Chicago';
const fmt_ = (d,fmt)=>Utilities.formatDate(d, tz_(), fmt);

const fmtDateMDY_ = d => (d instanceof Date ? fmt_(d,'MM/dd/yyyy') : '');
const fmtIsoMDY_  = iso => { if(!iso) return ''; const d=new Date(iso); return isNaN(d)?'':fmtDateMDY_(d); };
const fmtTimeHM_  = tstr => {
  const t=String(tstr||'').trim();
  if(!t) return '';
  const m=t.match(/^(\d{1,2})\/?(\d{2})$/) || t.match(/^(\d{1,2}):(\d{2})$/);
  if(!m) return '';
  const H=+m[1],M=+m[2];
  if(H>23||M>59) return '';
  return `${String(H).padStart(2,'0')}:${String(M).padStart(2,'0')}`;
};

function hhmm_(s){
  return fmtTimeHM_(s);
}
function totalHHMM_(a,b){
  const s=hhmm_(a),e=hhmm_(b);
  if(!s||!e) return '';
  const [sh,sm]=s.split(':').map(Number),[eh,em]=e.split(':').map(Number);
  let m=(eh*60+em)-(sh*60+sm);
  if(m<0) m+=1440;
  return `${String(Math.floor(m/60)).padStart(2,'0')}:${String(m%60).padStart(2,'0')}`;
}
function dateOnlyFromISO_(iso){
  if(!iso) return '';
  const y=+iso.slice(0,4),m=+iso.slice(5,7)-1,d=+iso.slice(8,10);
  return new Date(y,m,d,0,0,0,0);
}

function perfFromIso_(iso){
  const Z = tz_();
  if (!iso) {
    const now = new Date();
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 12, 0, 0, 0);
    return Utilities.formatDate(d, Z, 'MM/dd/yyyy');
  }
  const y = +iso.slice(0,4), m = +iso.slice(5,7)-1, d = +iso.slice(8,10);
  const atNoon = new Date(y, m, d, 12, 0, 0, 0);
  return Utilities.formatDate(atNoon, Z, 'MM/dd/yyyy');
}

function parsePerfToMDY_(perf){
  if(!perf) return '';
  if(perf instanceof Date) return fmtDateMDY_(perf);
  const s=String(perf).trim();
  let m=s.match(/^(\d{2})\/(\d{2})\/(\d{2})$/);
  if(m){
    const dd=+m[1], mm=+m[2], yy=2000+(+m[3]);
    return fmtDateMDY_(new Date(yy,mm-1,dd,12,0,0,0));
  }
  m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if(m){
    const mm=+m[1], dd=+m[2], yy=+m[3];
    return fmtDateMDY_(new Date(yy,mm-1,dd,12,0,0,0));
  }
  const dt=new Date(s);
  return isNaN(dt)?'':fmtDateMDY_(dt);
}

const hasCol_ = (sheet,col)=>{
  const s=tryGetSheet_(sheet);
  if(!s) return false;
  return hdr_(s).includes(col);
};

function norm_(s){
  const raw = s==null ? '' : String(s);
  return raw.replace(/[\u2018-\u201B]/g,"'")
            .replace(/[\u201C-\u201E]/g,'"')
            .replace(/[\u2013\u2014]/g,'-')
            .replace(/[^\x09\x0A\x0D\x20-\x7E]/g,'')
            .replace(/\s+/g,' ')
            .trim();
}
function esc_(s){
  return String(s||'').replace(/[&<>"]/g, m=>({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'
  }[m]));
}

// ----- Normalized ticket / work order readers (force IDs to strings) -----
function tickets_(){
  return read_(SH_TKT).map(function(t){
    const out = Object.assign({}, t);
    out.TicketID = String(t.TicketID || '').trim();
    return out;
  });
}
function workOrders_(){
  return read_(SH_WO).map(function(w){
    const out = Object.assign({}, w);
    out.WorkOrderID = String(w.WorkOrderID || '').trim();
    return out;
  });
}

// ----- Robust TicketID matcher -----
function eqTicketId_(a,b){
  const A = String(a || '').trim();
  const B = String(b || '').trim();
  if (!A || !B) return false;
  if (A === B) return true;

  const Aa = A.toUpperCase();
  const Bb = B.toUpperCase();

  if (Aa[0] === 'T' && Aa.slice(1) === Bb) return true;
  if (Bb[0] === 'T' && Bb.slice(1) === Aa) return true;

  const nA = Number(Aa.replace(/^T/i,''));
  const nB = Number(Bb.replace(/^T/i,''));
  if (!isNaN(nA) && !isNaN(nB) && nA === nB) return true;

  return false;
}

// ----- Default sorting helpers: newest → oldest -----
function _stampTicket_(t){
  const cand = t.LastUpdate || t.Timestamp || t.Time;
  const d = cand ? new Date(cand) : null;
  return d && !isNaN(d.getTime()) ? d.getTime() : 0;
}
function _sortTicketsNewest_(arr){
  return (arr || []).slice().sort((a,b)=>_stampTicket_(b) - _stampTicket_(a));
}
function _stampWorkOrder_(w){
  const cand = w.LastUpdate || w.Timestamp || w.Time || w.ScheduledDate;
  const d = cand ? new Date(cand) : null;
  return d && !isNaN(d.getTime()) ? d.getTime() : 0;
}
function _sortWorkOrdersNewest_(arr){
  return (arr || []).slice().sort((a,b)=>_stampWorkOrder_(b) - _stampWorkOrder_(a));
}

// ===== Email Controls (toggle only) =====
function isEmailEnabled_() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(EMAIL_TOGGLE_PROP);
  return raw == null ? EMAIL_DEFAULT_ENABLED : raw === 'true';
}
function assertEmailEnabledOrThrow_() {
  if (!isEmailEnabled_()) throw new Error('Email sending is currently DISABLED by admin toggle.');
}
function getEmailControls_() {
  return { enabled: isEmailEnabled_() };
}
function getEmailControls(){ return getEmailControls_(); }
function setEmailEnabled(enabled){
  PropertiesService.getScriptProperties().setProperty(EMAIL_TOGGLE_PROP, String(!!enabled));
  return getEmailControls_();
}

// ---------- Engineers/Contacts ----------
function engineers_(){ return read_(SH_ENG); }
function getEngineerById_(id){ return engineers_().find(e=>String(e.EngineerID)===String(id))||null; }
function engineerEmailById_(id){
  const e=getEngineerById_(id);
  return e ? String(e.EngineerEmail||e.Email||'').trim() : '';
}
function notifyEngineersEmails_(){
  const list = engineers_().filter(e=>{
    const n = String(e.Notifications||e.Notify||'').trim().toUpperCase();
    const active = 'Active' in e ? String(e.Active||'').trim().toUpperCase() : 'Y';
    const yn = (n==='Y'||n==='YES'||n==='TRUE'||n==='1');
    const act = (active==='Y'||active==='YES'||active==='TRUE'||active==='1');
    return yn && act;
  });
  return list.map(e=>String(e.EngineerEmail||e.Email||'').trim()).filter(Boolean);
}
function getEngineersPickerCore_(excludeId){
  const ex=String(excludeId||'');
  return engineers_().filter(e=>String(e.EngineerID||'')!==ex)
    .map(e=>({
      EngineerID  : String(e.EngineerID||''),
      EngineerName: String(e.EngineerName||''),
      Email       : String(e.EngineerEmail||e.Email||'')||''
    }))
    .sort((a,b)=>a.EngineerName.localeCompare(b.EngineerName));
}

// Corporate emails: prefer legacy Contacts if present; else derive from ClientContacts (look for CORP rows)
function corpEmails_(clientID){
  if(!clientID) return [];
  const legacy = tryRead_(SH_CONTACTS);
  if (legacy.length){
    const corp = String(clientID)+'COR';
    return legacy.filter(c=>String(c.ClientID)===String(clientID)
                          && String(c.MarketID||'')===corp
                          && (c.Email||c.ContactEmail))
                 .map(c=>String(c.Email||c.ContactEmail).trim())
                 .filter(Boolean);
  }
  const cc = tryRead_(SH_CLIENTCONTACTS);
  if (!cc.length) return [];
  const target = String(clientID)+'COR';
  const uniq = new Set();
  cc.forEach(r=>{
    if(String(r.ClientID)!==String(clientID)) return;
    Object.keys(r).forEach(k=>{
      if(/^MarketID\s*\d+$/i.test(k) && String(r[k]||'')===target){
        const em = String(r.ContactEmail||'').trim();
        if(em) uniq.add(em);
      }
    });
  });
  return Array.from(uniq);
}

// Preferred email from ticket
function requesterEmailFromTicket_(t){
  return String(t.RequestedByEmail || t.ContactEmail || '').trim();
}

// Active engineers
function getActiveEngineers_(){
  return engineers_()
    .filter(e=>{
      const active = String(e.Active||'').trim().toUpperCase();
      return active==='Y'||active==='YES'||active==='TRUE'||active==='1';
    })
    .map(e=>({
      EngineerID  : String(e.EngineerID||''),
      EngineerName: String(e.EngineerName||''),
      Email       : String(e.EngineerEmail||e.Email||'')
    }))
    .filter(x=>x.EngineerID||x.EngineerName||x.Email)
    .sort((a,b)=>a.EngineerName.localeCompare(b.EngineerName));
}

/**
 * buildTicketSubject_ — new subject:
 *   ClientName / MarketName / CALLSIGN / Priority / Status
 *   (no TicketID)
 */
function buildTicketSubject_(t, statusOpt, priorityOpt) {
  function safe(v){ return (v == null ? '' : String(v)).trim(); }

  const client   = safe(t.ClientName);
  const market   = safe(t.MarketName);
  const call     = safe(t.CallSign);
  const status   = safe(statusOpt || t.Status);
  const priority = safe(priorityOpt || t.Priority);

  const parts = [];
  if (client)   parts.push(client);
  if (market)   parts.push(market);
  if (call)     parts.push(call.toUpperCase());
  if (priority) parts.push(priority);
  if (status)   parts.push(status);

  return parts.join(' / ');
}

/**
 * sendEmail_ — Gmail API via Advanced Gmail service
 */
function sendEmail_(toList, ccList, subject, textBody, htmlBody, bccList){
  assertEmailEnabledOrThrow_();

  const uniq = a => Array.from(new Set((a||[]).map(x=>String(x||'').trim()).filter(Boolean)));
  const toU = uniq(toList);
  let ccU = uniq(ccList).filter(e=>!toU.includes(e));
  let bccU = uniq(bccList).filter(e=>!toU.includes(e) && !ccU.includes(e));

  if (!toU.length) return { ok:false, reason:'no-to' };

  var nl = '\r\n';
  var boundary = 'OBSI_BOUNDARY_' + new Date().getTime();

  var headers = [];

  if (SUPPORT_NAME || SUPPORT_EMAIL) {
    var fromLine = SUPPORT_NAME && SUPPORT_EMAIL
      ? SUPPORT_NAME + ' <' + SUPPORT_EMAIL + '>'
      : (SUPPORT_NAME || SUPPORT_EMAIL);
    headers.push('From: ' + fromLine);
  }

  headers.push('To: ' + toU.join(', '));
  if (ccU.length)  headers.push('Cc: ' + ccU.join(', '));
  if (bccU.length) headers.push('Bcc: ' + bccU.join(', '));
  headers.push('Subject: ' + subject);
  headers.push('MIME-Version: 1.0');
  if (SUPPORT_EMAIL) headers.push('Reply-To: ' + SUPPORT_EMAIL);

  var body;
  if (htmlBody) {
    headers.push('Content-Type: multipart/alternative; boundary="' + boundary + '"');
    body =
      headers.join(nl) + nl + nl +
      '--' + boundary + nl +
      'Content-Type: text/plain; charset="UTF-8"' + nl + nl +
      (textBody || '') + nl + nl +
      '--' + boundary + nl +
      'Content-Type: text/html; charset="UTF-8"' + nl + nl +
      htmlBody + nl + nl +
      '--' + boundary + '--';
  } else {
    headers.push('Content-Type: text/plain; charset="UTF-8"');
    body = headers.join(nl) + nl + nl + (textBody || '');
  }

  var raw = Utilities.base64EncodeWebSafe(body);
  // Standard message object for Gmail API
  var message = { raw: raw };

  try {
    var sent = Gmail.Users.Messages.send(message, 'me');
    return { ok:true, to:toU, cc:ccU, bcc:bccU, id: sent && sent.id };
  } catch (e) {
    throw new Error('Gmail API send failed: ' + (e && e.message ? e.message : e));
  }
}

/* =========================================================
   EMAILCC ROUTING
   ========================================================= */

function emailCCListByMarket_(marketID){
  try{
    const s = sh_(SH_EMAILCC);
    const v = s.getDataRange().getValues();
    if(!v.length) return [];
    const H = v[0];
    const iMID = H.indexOf('MarketID');
    if(iMID<0) return [];
    const out=[];
    for(let r=1;r<v.length;r++){
      const row=v[r];
      if(String(row[iMID])===String(marketID)){
        for(let c=0;c<row.length;c++){
          if(c!==iMID && row[c] && String(row[c]).indexOf('@')>=0) out.push(String(row[c]).trim());
        }
      }
    }
    return Array.from(new Set(out));
  }catch(_){ return []; }
}

function routeEmailForTicket_(t){
  const to  = String(t.ContactEmail||'').trim();
  const ccs = emailCCListByMarket_(t.MarketID||'');
  if(to) return {to:[to], cc:ccs};
  if(ccs.length) return {to:[ccs[0]], cc:ccs.slice(1)};
  return {to:[], cc:[]};
}

// ---------- Email bundle builders ----------
function getTicketBundle_(ticketId) {
  const tidNorm = String(ticketId||'').trim();
  const tickets = tickets_();
  const recsAll = read_(SH_TREC);

  const t = tickets.find(x => eqTicketId_(x.TicketID, tidNorm));
  if (!t) return null;

  const recs = recsAll
    .filter(r => eqTicketId_(r.TicketID, tidNorm) && String(r.Status||'').toUpperCase()!=='DELETED')
    .sort((a,b)=>new Date(a.Timestamp)-new Date(b.Timestamp));

  return { ticket: t, records: recs };
}

function composeTicketEmailBundle_(bundle, heading, outro) {
  const t = bundle.ticket, recs = bundle.records || [];

  const toMDY = s => {
    const raw = String(s || '').trim();
    if (!raw) return '';
    const m = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (m) {
      let a = +m[1], b = +m[2], y = m[3];
      const yyyy = y.length === 2 ? 2000 + (+y) : +y;
      let mm = a, dd = b;
      if (a > 12 && b <= 12) { mm = b; dd = a; }
      return `${String(mm).padStart(2,'0')}/${String(dd).padStart(2,'0')}/${String(yyyy).padStart(4,'0')}`;
    }
    const d = new Date(raw);
    return isNaN(d) ? '' : fmtDateMDY_(d);
  };

  const pairs = [
    ['Client', t.ClientName],
    ['Market', t.MarketName],
    ['Station', String(t.CallSign || '').toUpperCase()],
    t.Priority ? ['Priority', t.Priority] : ['Priority', 'Not Assigned'],
    t.Status ? ['Status', t.Status] : ['Status', 'Not Set'],
    (t.RequestedBy || t.RequestedByName) ? ['Requested by', t.RequestedBy || t.RequestedByName] : ['Requested by', 'Not Provided'],
    t.Time ? ['Date of Request', fmtDateMDY_(new Date(t.Time))] : ['Date of Request', 'Not Provided']
  ].filter(Boolean);

  const lines = [];
  lines.push(norm_(heading), '');
  lines.push('Ticket ID: ' + t.TicketID, '');
  pairs.forEach(([k, v]) => lines.push(k + ': ' + (v!=null?v:'')));
  if (t.SupportRequest) {
    lines.push('', 'Support Request:', norm_(t.SupportRequest));
  }

  lines.push('');
  if (recs.length) {
    lines.push('Support Records:');
    recs.slice().sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp))
      .forEach(r => {
        const dateLine = toMDY(r.PerfDate || r.PerFdate || '');
        const who = r.EngineerName || '';
        const detail = norm_(r.SupportProvided || '');
        if (dateLine || who) lines.push('  ' + dateLine + (dateLine && who ? '\t' : '') + who);
        if (detail) lines.push('  ' + detail);
        lines.push('');
      });
  } else {
    lines.push('No support records yet.');
  }
  lines.push('', norm_(outro || '-- Offerdahl Broadcast Service, Inc.'));
  const text = lines.join('\n');

  const rr = recs.length
    ? recs.slice().sort((a,b) => new Date(b.Timestamp) - new Date(a.Timestamp))
        .map(r => {
          const d   = esc_(toMDY(r.PerfDate || r.PerFdate || ''));
          const who = esc_(r.EngineerName || '');
          const n   = esc_(norm_(r.SupportProvided || ''));
          const row1 = (d || who) ? '<div><span style="white-space:pre">'+d+(d && who ? '    ' : '')+who+'</span></div>' : '';
          const row2 = n ? '<div style="margin-top:2px">'+n+'</div>' : '';
          return '<tr><td style="padding:6px 0;border-bottom:1px solid #eee">'+row1+row2+'</td></tr>';
        }).join('')
    : '<tr><td style="padding:4px 0;color:#999">No support records yet.</td></tr>';

  const html = '<div style="font:14px/1.35 -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif">'
    + '<h3 style="margin:0 0 8px 0">'+esc_(heading)+'</h3>'
    + '<p><strong>Ticket ID:</strong> '+esc_(t.TicketID)+'</p>'
    + '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin:8px 0">'
    + pairs.map(([k,v]) => '<tr><td style="padding:2px 8px 2px 0;color:#444"><strong>'+esc_(k)+'</strong></td><td>'+esc_(v!=null?v:'')+'</td></tr>').join('')
    + '</table>'
    + (t.SupportRequest ? '<p style="margin:10px 0 4px"><strong>Support Request:</strong></p><div>'+esc_(norm_(t.SupportRequest))+'</div>' : '')
    + '<p style="margin:12px 0 4px"><strong>Support Records:</strong></p>'
    + '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;margin:0 0 8px 0">'+rr+'</table>'
    + '<p style="color:#555">Offerdahl Broadcast Service, Inc.</p>'
    + '</div>';

  const subject = buildTicketSubject_(t);

  return { subject, text, html };
}

function getWorkOrderBundle_(woId) {
  const widNorm = String(woId||'').trim();
  const w = workOrders_().find(x => String(x.WorkOrderID||'').trim() === widNorm);
  if (!w) return null;
  const t = w.TicketID
    ? tickets_().find(x => eqTicketId_(x.TicketID, w.TicketID))
    : null;
  const recs = read_(SH_WREC)
    .filter(r => String(r.WorkOrderID) === String(w.WorkOrderID) && String(r.Status||'').toUpperCase()!=='DELETED')
    .sort((a,b)=>new Date(a.Timestamp)-new Date(b.Timestamp));
  return { workOrder: w, parentTicket: t, records: recs };
}

// ---------- ST Detail RPC (for panel refresh) ----------
function getTicketDetail(ticketId){
  const tidNorm = String(ticketId||'').trim();
  if(!tidNorm){
    return { ok:false, error:'TicketID required.' };
  }

  const tickets = tickets_();
  const recs    = read_(SH_TREC);

  let t = tickets.find(x => eqTicketId_(x.TicketID, tidNorm));

  // FILTER OUT DELETED CHILD RECORDS
  const records = recs
    .filter(r => eqTicketId_(r.TicketID, tidNorm))
    .filter(r => {
      const st  = String(r.Status || '').toUpperCase();
      const txt = String(r.SupportProvided || '');
      if (st === 'DELETED') return false;
      if (txt.indexOf('[DELETED]') === 0) return false;
      return true;
    })
    .sort((a,b)=>new Date(b.Timestamp) - new Date(a.Timestamp));

  if (!t){
    if (!records.length){
      return { ok:false, error:'Ticket not found in Tickets or SupportRecords.' };
    }
    t = {
      TicketID       : tidNorm,
      ClientID       : '',
      ClientName     : '',
      MarketID       : '',
      MarketName     : '',
      CallSign       : '',
      Status         : '',
      Priority       : '',
      RequestedBy    : '',
      RequestedByName: '',
      RequestedByEmail: '',
      ContactEmail   : '',
      SupportRequest : '',
      Time           : '',
      LastUpdate     : ''
    };
  }

  return { ok:true, ticket:t, records:records };
}

// === Work Order email bundle ===
function composeWorkOrderEmailBundle_(bundle, heading, outro) {
  const w = bundle.workOrder;
  const t = bundle.parentTicket;
  const recs = bundle.records || [];

  const subject = 'Work Order ' + w.WorkOrderID;

  const details = [
    ['Client', w.ClientName],
    ['Market', w.MarketName],
    ['Station', String(w.CallSign || '').toUpperCase()],
    ['Priority', w.Priority || 'Not Set'],
    ['Status', w.Status || 'Not Set'],
    ['Engineer', w.EngineerName || '(unassigned)'],
    ['Scheduled Date', w.ScheduledDate ? fmtDateMDY_(new Date(w.ScheduledDate)) : '(not scheduled)']
  ];

  const lines = [];
  lines.push(norm_(heading), '');
  details.forEach(function(pair){
    var label = pair[0], value = pair[1];
    lines.push(label + ': ' + (value != null ? value : ''));
  });

  if (w.WorkOrderText) {
    lines.push('');
    lines.push('Work Requested:');
    lines.push('  ' + norm_(w.WorkOrderText));
  }

    lines.push('');
  // Label child rows as Tasks in plain-text emails
  lines.push('Tasks:');
  recs.forEach(r => {
    const perfDate = fmtDateMDY_(new Date(r.PerfDate || r.PerFdate || ''));
    const engineer = r.EngineerName || '(No engineer)';
    const workNote = norm_(r.WorkPerformed || 'No details available');
    lines.push(perfDate + '    ' + engineer);
    lines.push('  ' + workNote);
    lines.push('---------------------------------------');
  });


  lines.push('', norm_(outro || '-- Offerdahl Broadcast Service, Inc.'));
  const text = lines.join('\n');

  const recordRows = recs.length
    ? recs.map(r => {
        const perfDate = esc_(fmtDateMDY_(new Date(r.PerfDate || r.PerFdate || '')));
        const engineer = esc_(r.EngineerName || '(No engineer)');
        const note = esc_(norm_(r.WorkPerformed || 'No details available'));
        return (
          '<tr><td style="padding:2px 6px 2px 0;vertical-align:top;color:#555">' +
          perfDate +
          '    ' +
          engineer +
          '</td><td style="padding:2px 0">' +
          note +
          '</td></tr>'
        );
      }).join('')
    : '<tr><td colspan="2" style="padding:4px 0;color:#999">No work records yet.</td></tr>';

  const html =
    '<div style="font:14px/1.35 -apple-system,BlinkMacSystemFont,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif">' +
    '<h3 style="margin:0 0 8px 0">' + esc_(heading) + '</h3>' +
    '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin:8px 0">' +
    details
      .map(
        ([k, v]) =>
          '<tr><td style="padding:2px 8px 2px 0;color:#444"><strong>' +
          esc_(k) +
          '</strong></td><td>' +
          esc_(v != null ? v : '') +
          '</td></tr>'
      )
      .join('') +
    '</table>' +
    (w.WorkOrderText
      ? '<p style="margin:10px 0 4px"><strong>Work Requested:</strong></p><div>' +
        esc_(norm_(w.WorkOrderText)) +
        '</div>'
      : '') +
    // Label child rows as Tasks in HTML emails
    '<p style="margin:8px 0 2px"><strong>Tasks:</strong></p>' +
    '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin:0 0 8px 0">' +
    recordRows +
    '</table>' +

    '<p style="color:#555">Offerdahl Broadcast Service, Inc.</p>' +
    '</div>';

  return { subject, text, html };
}

// ---------- ST Core ----------
function getInitial(){
  const email = getEmailControls_();

  const ticketsRaw    = tickets_();
  const workordersRaw = workOrders_();

  return {
    build: BUILD, url: ASP_URL,
    clients:        read_(SH_CLIENTS),
    markets:        read_(SH_MARKETS),
    callsigns:      read_(SH_CALLSIGNS),
    engineers:      read_(SH_ENG),
    contacts:       tryRead_(SH_CONTACTS),
    clientContacts: tryRead_(SH_CLIENTCONTACTS),

    tickets:        _sortTicketsNewest_(ticketsRaw),
    supportRecords: read_(SH_TREC),

    workorders:       _sortWorkOrdersNewest_(workordersRaw),
    workOrderRecords: read_(SH_WREC),

    email
  };
}
function refreshInitial(){ return getInitial(); }
function doGet(){
  const t=HtmlService.createTemplateFromFile('index');
  t.initial=getInitial(); t.build=BUILD; t.asp_url=ASP_URL;
  return t.evaluate().setTitle('OBSI Support Console')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- Requester resolvers ----------
function getContactById_(contactId){
  const id = String(contactId||'').trim(); if(!id) return null;
  const legacy = tryRead_(SH_CONTACTS);
  if (legacy.length) return legacy.find(c => String(c.ContactID||'').trim() === id) || null;
  const cc = tryRead_(SH_CLIENTCONTACTS);
  return cc.find(c => String(c.ContactID||'').trim() === id) || null;
}
function resolveRequester_(p){
  if (p && p.RequestedByType && p.RequestedByID){
    const type = String(p.RequestedByType).toUpperCase();
    if (type === 'CONTACT'){
      const rec = getContactById_(p.RequestedByID);
      return {
        type: 'CONTACT',
        id  : (rec && rec.ContactID) || String(p.RequestedByID||''),
        name: (rec && rec.ContactName) || p.RequestedByName || '',
        email: (rec && (rec.Email || rec.ContactEmail)) || p.RequestedByEmail || ''
      };
    } else if (type === 'ENGINEER'){
      const eng = getEngineerById_(p.RequestedByID);
      return {
        type: 'ENGINEER',
        id  : (eng && eng.EngineerID) || String(p.RequestedByID||''),
        name: (eng && eng.EngineerName) || p.RequestedByName || '',
        email: (eng && (eng.EngineerEmail || eng.Email)) || p.RequestedByEmail || ''
      };
    }
  }
  if (p && (p.ContactID || p.ContactEmail || p.RequestedBy)){
    const rec = p.ContactID ? getContactById_(p.ContactID) : null;
    const name = (rec && rec.ContactName) || p.RequestedBy || '';
    const email = (rec && (rec.Email || rec.ContactEmail)) || p.ContactEmail || '';
    const id = (rec && rec.ContactID) || String(p.ContactID||'');
    return { type:'CONTACT', id:id, name:name, email:email };
  }
  return {
    type :'CONTACT',
    id   :'',
    name : String(p && p.RequestedBy || ''),
    email: String(p && p.ContactEmail || '')
  };
}

// ---------- ST Core ----------
function createSupportTicket_OBSI(p){
  const c=p||{},when=now_(),clients=read_(SH_CLIENTS),markets=read_(SH_MARKETS),
  nm=(a,k,id,ret)=>{const f=a.find(x=>String(x[k]||'')===String(id||''));return f?String(f[ret]||''):''};

  const req = resolveRequester_(c);

  const row={
    Timestamp     : when,
    TicketID      : String(c.TicketID||('T'+Date.now())),
    ClientID      : c.ClientID||'',
    ClientName    : nm(clients,'ClientID',c.ClientID,'ClientName'),
    MarketID      : c.MarketID||'',
    MarketName    : nm(markets,'MarketID',c.MarketID,'MarketName'),
    CallSign      : String(c.CallSign||'').toUpperCase(),
    RequestedBy   : req.name || c.RequestedBy || '',
    ContactEmail  : req.email || c.ContactEmail || '',
    RequestedByType : req.type || '',
    RequestedByID   : req.id   || '',
    RequestedByName : req.name || '',
    RequestedByEmail: req.email|| '',
    SupportRequest  : norm_(c.SupportRequest||''),
    Status          : 'NEW',
    EmailSent       : '',
    LastUpdate      : when,
    ContactID       : c.ContactID||'',
    CreatedBy       : (Session.getActiveUser()&&Session.getActiveUser().getEmail())||'OBSI Console',
    Time            : when,
    EngineerID      : c.EngineerID||'',
    EngineerName    : c.EngineerName||'',
    Priority        : c.Priority||'Low'
  };

  const miss=[];
  if(!row.ClientID)     miss.push('ClientID');
  if(!row.MarketID)     miss.push('MarketID');
  if(!row.CallSign)     miss.push('CallSign');
  if(!row.SupportRequest) miss.push('SupportRequest');
  if(miss.length) throw new Error('Missing: '+miss.join(', ')+' — STRICT CANON');

  appendByHdr_(SH_TKT,row);

  try{
    const bundle = getTicketBundle_(row.TicketID);
    if(bundle){
      const mail = composeTicketEmailBundle_(bundle,'Open','We will follow up with updates as we work your request.');
      const routed = routeEmailForTicket_(bundle.ticket);
      const sent=sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok){
        writeByKey_(SH_TKT,'TicketID',row.TicketID,{EmailSent:now_(),LastUpdate:now_()});
      }
    }
  }catch(_){}
  return {ok:true, TicketID:row.TicketID, notice:'Created Ticket '+row.TicketID};
}

function setTicketPending(ticketId){
  const tidNorm = String(ticketId||'').trim();
  const t=tickets_().find(x=>eqTicketId_(x.TicketID, tidNorm)); if(!t) throw new Error('Ticket not found');
  writeByKey_(SH_TKT,'TicketID',t.TicketID,{Status:'PENDING',LastUpdate:now_()});
  try{
    const bundle = getTicketBundle_(t.TicketID);
    if(bundle){
      const mail = composeTicketEmailBundle_(bundle,'Pending','We\'ll keep you updated.');
      const routed = routeEmailForTicket_(bundle.ticket);
      const sent=sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok){
        writeByKey_(SH_TKT,'TicketID',t.TicketID,{EmailSent:now_(),LastUpdate:now_()});
        return {ok:true, notice:'Ticket '+t.TicketID+' marked PENDING (email sent)'};
      }else if(sent.reason==='no-to'){
        return {ok:true, notice:'Ticket '+t.TicketID+' marked PENDING (no recipient email found)'};
      }
    }
  }catch(_){}
  return {ok:true, notice:'Ticket '+t.TicketID+' marked PENDING'};
}

/**
 * promoteTicketToWorkOrder — ESCALATE Ticket -> Task-based WorkOrder
 * If workOrderIdOpt is provided, attaches Ticket as a new Task to that
 * existing master WorkOrder. Otherwise, creates a new master WorkOrder
 * and a first Task from the Ticket's SupportRequest.
 *
 * NOTE:
 * - WorkOrderRecords are treated as Tasks.
 * - PerfDate is used as Task Due Date (MM/dd/yyyy).
 */
function promoteTicketToWorkOrder(ticketId, workOrderIdOpt){
  const tidNorm = String(ticketId||'').trim();
  const t = tickets_().find(x=>eqTicketId_(x.TicketID, tidNorm));
  if(!t) throw new Error('Ticket not found');

  const existingWO = String(workOrderIdOpt||'').trim();
  let woId = existingWO;
  let createdNew = false;

  if(!woId){
    // Create NEW master WorkOrder
    woId = 'WO-'+fmt_(new Date(),'yyyyMMdd-HHmmss');
    const s = sh_(SH_WO),
          h = hdr_(s),
          row = h.map(()=>''), 
          set = (k,v)=>{ const i=h.indexOf(k); if(i>=0) row[i]=v; };

    set('WorkOrderID',woId);
    set('TicketID',t.TicketID);
    set('ClientID',t.ClientID);
    set('ClientName',t.ClientName);
    set('MarketID',t.MarketID);
    set('MarketName',t.MarketName);
    set('CallSign',String(t.CallSign||'').toUpperCase());
    set('WorkOrderText',t.SupportRequest||'');
    set('Status','New');
    set('LastUpdate',now_());
    s.appendRow(row);
    createdNew = true;
  } else {
    // Attach to EXISTING master WorkOrder; verify it exists and is eligible
    const all = workOrders_();
    const w = all.find(x => String(x.WorkOrderID||'').trim() === woId);
    if(!w) throw new Error('WorkOrder not found: '+woId);

    const st = String(w.Status||'').toUpperCase();
    if(st === 'CLOSED' || st === 'DELETED'){
      throw new Error('Cannot attach to closed/deleted WorkOrder: '+woId);
    }
    if(String(w.ClientID) !== String(t.ClientID) || String(w.MarketID) !== String(t.MarketID)){
      throw new Error('WorkOrder '+woId+' does not match Ticket client/market.');
    }
  }

  // Mark Ticket as ESCALATED
  try{
    writeByKey_(SH_TKT,'TicketID',t.TicketID,{Status:'ESCALATED',LastUpdate:now_()});
  }catch(_){}

  // Create initial Task (WO child record) from Ticket's SupportRequest
  let taskId = null;
  try{
    const due = fmtDateMDY_(new Date());  // Task Due Date = today (MM/dd/yyyy)
    taskId = 'TASK-'+fmt_(new Date(),'yyyyMMdd-HHmmss');

    const rec = {
      Timestamp    : now_(),
      WorkRecordID : taskId,                 // treated as TaskID
      WorkOrderID  : woId,
      TicketID     : t.TicketID,
      EngineerID   : t.EngineerID || '',
      EngineerName : t.EngineerName || '',
      WorkPerformed: norm_(t.SupportRequest || ''), // Task description
      StartTime    : '',                     // legacy time fields left empty
      EndTime      : '',
      TotalTime    : '',
      PerfDate     : due,                    // Due Date
      EmailSent    : '',
      Status       : 'New'
    };

    appendByHdr_(SH_WREC, rec);
  }catch(_){}

  // Send WorkOrder email (bundle will now include the Task we just created)
  try{
    const wBundle = getWorkOrderBundle_(woId);
    if(wBundle){
      const mail   = composeWorkOrderEmailBundle_(wBundle,'Escalated','We will proceed with scheduling and updates.');
      const routed = routeEmailForTicket_(t);
      const sent   = sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok){
        writeByKey_(SH_TKT,'TicketID',t.TicketID,{EmailSent:now_(),LastUpdate:now_()});
      }
    }
  }catch(_){}

  const msg = createdNew
    ? ('Escalated '+t.TicketID+' -> '+woId)
    : ('Attached '+t.TicketID+' to '+woId);

  return { ok:true, WorkOrderID:woId, TaskID:taskId, createdNew:createdNew, notice:msg };
}


// ---------- WO Core ----------
function assignWorkOrder(woId,engId){
  const e=getEngineerById_((engId||'').trim()); if(!e) throw new Error('Engineer not found');
  writeByKey_(SH_WO,'WorkOrderID',woId,{EngineerID:engId,EngineerName:e.EngineerName||'',LastUpdate:now_()});

  try{
    const wBundle = getWorkOrderBundle_(woId);
    if(!wBundle) throw new Error('WorkOrder not found');
    const t = wBundle.parentTicket;
    // Prepare to[] as a flat array (avoid nested-array)
    const to = (t && requesterEmailFromTicket_(t)) ? [ requesterEmailFromTicket_(t) ] : [];
    const cc=[ engineerEmailById_((engId)), ...corpEmails_(wBundle.workOrder.ClientID), ...notifyEngineersEmails_() ];
    const mail = composeWorkOrderEmailBundle_(wBundle,'Assigned','We will follow up with scheduling and updates.');
    const sent=sendEmail_(to.filter(Boolean),cc,mail.subject,mail.text,mail.html);
    if(sent.ok) writeByKey_(SH_WO,'WorkOrderID',woId,{EmailSent:now_(),LastUpdate:now_()});
  }catch(_){}

  return {ok:true, notice:'Assigned '+woId+' to '+(e.EngineerName||engId)};
}

function setWorkOrderScheduledDate(woId,iso){
  if(!woId) throw new Error('WorkOrderID required');
  const s=sh_(SH_WO),h=hdr_(s),v=s.getDataRange().getValues(),
        cWO=h.indexOf('WorkOrderID'),cSD=h.indexOf('ScheduledDate'),cLU=h.indexOf('LastUpdate'),cST=h.indexOf('Status');
  const r=v.findIndex((row,i)=>i>0&&String(row[cWO])===String(woId)); if(r<1) throw new Error('WO not found');
  const R=r+1,val=iso?dateOnlyFromISO_(iso):'';
  s.getRange(R,cSD+1).setValue(val).setNumberFormat('yyyy-mm-dd');
  if(cLU>=0) s.getRange(R,cLU+1).setValue(now_());
  if(cST>=0) s.getRange(R,cST+1).setValue('Scheduled');

  try{
    const wBundle = getWorkOrderBundle_(woId);
    if(wBundle){
      const mail = composeWorkOrderEmailBundle_(wBundle,'Scheduled','We will follow up with additional details as needed.');
      const t = wBundle.parentTicket || {
        MarketID    : wBundle.workOrder.MarketID,
        ContactEmail: requesterEmailFromTicket_(wBundle.parentTicket||{})
      };
      const routed = routeEmailForTicket_(t);
      const sent=sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok) writeByKey_(SH_WO,'WorkOrderID',woId,{EmailSent:now_(),LastUpdate:now_()});
    }
  }catch(_){}

  return {ok:true, notice:'Scheduled '+woId+' for '+fmtDateMDY_(val||new Date())};
}

function updateTicketText(id,txt){
  writeByKey_(SH_TKT,'TicketID',id,{SupportRequest:norm_(txt),LastUpdate:now_()});
  return {ok:true, notice:'Saved Ticket '+id};
}
function updateWorkOrderText(id,txt){
  writeByKey_(SH_WO,'WorkOrderID',id,{WorkOrderText:norm_(txt),LastUpdate:now_()});
  return {ok:true, notice:'Saved Work Order '+id};
}

// Records
const pStatusT_ = tid=>{
  const tidNorm = String(tid||'').trim();
  const t=tickets_().find(x=>eqTicketId_(x.TicketID, tidNorm));
  return t?String(t.Status||''):'';
};

// SAFE REFACTOR: avoid const `w` to prevent "Cannot access 'w' before initialization"
const pStatusW_ = wid=>{
  const widNorm = String(wid||'').trim();
  const all = workOrders_();
  const found = all.find(x => String(x.WorkOrderID||'').trim() === widNorm);
  return found ? String(found.Status||'') : '';
};

/**
 * addSupportRecord — FORCE Ticket + SupportRecord to IN PROGRESS
 * on every child record, regardless of previous status.
 */
function addSupportRecord(tid,eid,note,st,en,perfIso){
  const tidNorm = String(tid||'').trim();

  try {
    writeByKey_(SH_TKT,'TicketID',tidNorm,{
      Status:'IN PROGRESS',
      LastUpdate:now_()
    });
  } catch(_) {}

  const start = hhmm_(st);
  const end   = hhmm_(en) || '';
  const tot   = (start && end) ? totalHHMM_(start,end) : '';
  const perf  = perfFromIso_(perfIso);
  const eng   = getEngineerById_(eid) || {};

  const rec = {
    Timestamp      : now_(),
    TicketID       : tidNorm,
    EngineerID     : eid || '',
    EngineerName   : eng.EngineerName || '',
    SupportProvided: norm_(note || ''),
    StartTime      : start,
    EndTime        : end,
    TotalTime      : tot,
    PerFdate       : perf,
    EmailSent      : ''
  };

  if(hasCol_(SH_TREC,'Status')) rec.Status = 'IN PROGRESS';

  appendByHdr_(SH_TREC,rec);

  try{ writeByKey_(SH_TKT,'TicketID',tidNorm,{LastUpdate:now_()}); }catch(_){}

  try{
    const bundle = getTicketBundle_(tidNorm);
    if(bundle){
      const mail   = composeTicketEmailBundle_(bundle,'Update','We\'ll keep you updated.');
      const routed = routeEmailForTicket_(bundle.ticket);
      const sent   = sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok) writeByKey_(SH_TKT,'TicketID',tidNorm,{EmailSent:now_(),LastUpdate:now_()});
    }
  }catch(_){}

  return {
    ok    : true,
    notice: 'Added ST record (' + (tot || '-') + ')'
  };
}

/**
 * updateSupportRecord — Update an existing support record (by Timestamp)
 */
function updateSupportRecord(ts, eid, note, st, en, perfIso){
  const recs = read_(SH_TREC);
  const rec = recs.find(r => String(r.Timestamp) === String(ts));
  if(!rec) throw new Error('Support record not found: ' + ts);

  const start = hhmm_(st);
  const end = hhmm_(en) || '';
  const tot = (start && end) ? totalHHMM_(start,end) : '';
  const perf = perfFromIso_(perfIso);
  const eng = getEngineerById_(eid) || {};

  const updates = {
    SupportProvided: norm_(note || ''),
    EngineerID: eid || '',
    EngineerName: eng.EngineerName || '',
    StartTime: start,
    EndTime: end,
    TotalTime: tot,
    PerFdate: perf
  };
  if(hasCol_(SH_TREC,'Status')) updates.Status = 'IN PROGRESS';

  writeByKey_(SH_TREC,'Timestamp',ts, updates);

  // Update parent Ticket last update
  try{
    const tid = rec.TicketID;
    if(tid) writeByKey_(SH_TKT,'TicketID',tid,{LastUpdate:now_()});
  }catch(_){}

  // Optionally send update email (preserve original behavior)
  try{
    const bundle = getTicketBundle_(rec.TicketID);
    if(bundle){
      const mail   = composeTicketEmailBundle_(bundle,'Update','We\'ll keep you updated.');
      const routed = routeEmailForTicket_(bundle.ticket);
      const sent   = sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok) writeByKey_(SH_TKT,'TicketID',rec.TicketID,{EmailSent:now_(),LastUpdate:now_()});
    }
  }catch(_){}

  return { ok:true, notice:'Support record updated' };
}

/**
 * addWorkOrderRecord — ADD TASK (child row) for a WorkOrder.
 *
 * WorkOrderRecords are treated as Tasks. PerfDate is used as Task
 * Due Date for new entries. Start/End/TotalTime are kept for
 * legacy support only and are NOT defaulted if omitted.
 */
function addWorkOrderRecord(woid,tid,eid,note,st,en,perfIso){
  const start = hhmm_(st) || '';            // no default "now" time
  const end   = hhmm_(en) || '';
  const tot   = (start && end) ? totalHHMM_(start,end) : '';
  const perf  = perfFromIso_(perfIso);      // Due Date (MM/dd/yyyy)
  const eng   = getEngineerById_(eid) || {};
  const id    = 'WR-'+fmt_(new Date(),'yyyyMMdd-HHmmss');

  const rec = {
    Timestamp    : now_(),
    WorkRecordID : id,                      // TaskID
    WorkOrderID  : woid,
    TicketID     : tid || '',
    EngineerID   : eid || '',
    EngineerName : eng.EngineerName || '',
    WorkPerformed: norm_(note || ''),      // Task description
    StartTime    : start,                  // legacy, may be empty
    EndTime      : end,
    TotalTime    : tot,
    PerfDate     : perf,                   // Due Date
    EmailSent    : ''
  };

  if(hasCol_(SH_WREC,'Status')) rec.Status = pStatusW_(woid) || '';

  appendByHdr_(SH_WREC,rec);

  try{ writeByKey_(SH_WO,'WorkOrderID',woid,{LastUpdate:now_()}); }catch(_){}

  try{
    const wBundle = getWorkOrderBundle_(woid);
    if(wBundle){
      const mail   = composeWorkOrderEmailBundle_(wBundle,'Update','We\'ll keep you updated.');
      const t      = wBundle.parentTicket || {
        MarketID    : wBundle.workOrder.MarketID,
        ContactEmail: requesterEmailFromTicket_(wBundle.parentTicket||{})
      };
      const routed = routeEmailForTicket_(t);
      const sent   = sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok) writeByKey_(SH_WO,'WorkOrderID',woid,{EmailSent:now_(),LastUpdate:now_()});
    }
  }catch(_){}

  return {
    ok         : true,
    WorkRecordID: id,
    notice     : 'Added WO record (' + (tot || '-') + ')'
  };
}

/**
 * updateWorkOrderRecord — Update an existing WorkOrder record (by WorkRecordID)
 */
function updateWorkOrderRecord(wrid, eid, note, st, en, perfIso){
  const recs = read_(SH_WREC);
  const rec = recs.find(r => String(r.WorkRecordID) === String(wrid));
  if(!rec) throw new Error('Work order record not found: ' + wrid);

  const start = hhmm_(st) || '';
  const end = hhmm_(en) || '';
  const tot = (start && end) ? totalHHMM_(start,end) : '';
  const perf = perfFromIso_(perfIso);
  const eng = getEngineerById_(eid) || {};

  const updates = {
    WorkPerformed: norm_(note || ''),
    EngineerID: eid || '',
    EngineerName: eng.EngineerName || '',
    StartTime: start,
    EndTime: end,
    TotalTime: tot,
    PerfDate: perf
  };
  if(hasCol_(SH_WREC,'Status')) updates.Status = pStatusW_(rec.WorkOrderID) || '';

  writeByKey_(SH_WREC,'WorkRecordID',wrid, updates);

  // Update parent WorkOrder last update
  try{
    const wid = rec.WorkOrderID;
    if(wid) writeByKey_(SH_WO,'WorkOrderID',wid,{LastUpdate:now_()});
  }catch(_){}

  // Optionally send update email
  try{
    const wBundle = getWorkOrderBundle_(rec.WorkOrderID);
    if(wBundle){
      const mail   = composeWorkOrderEmailBundle_(wBundle,'Update','We\'ll keep you updated.');
      const t      = wBundle.parentTicket || {
        MarketID    : wBundle.workOrder.MarketID,
        ContactEmail: requesterEmailFromTicket_(wBundle.parentTicket||{})
      };
      const routed = routeEmailForTicket_(t);
      const sent   = sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok) writeByKey_(SH_WO,'WorkOrderID',rec.WorkOrderID,{EmailSent:now_(),LastUpdate:now_()});
    }
  }catch(_){}

  return { ok:true, notice:'Work order record updated' };
}


// Close events
function closeTicket(id){
  writeByKey_(SH_TKT,'TicketID',id,{Status:'CLOSED',LastUpdate:now_()});
  cascadeT_(id,'CLOSED');
  try{
    const bundle = getTicketBundle_(id);
    if(bundle){
      const mail   = composeTicketEmailBundle_(bundle,'Closed','Thank you for working with OBSI.');
      const routed = routeEmailForTicket_(bundle.ticket);
      const sent   = sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok) writeByKey_(SH_TKT,'TicketID',id,{EmailSent:now_(),LastUpdate:now_()});
    }
  }catch(_){}
  return {ok:true, notice:'Closed Ticket '+id};
}

function closeWorkOrder(id){
  writeByKey_(SH_WO,'WorkOrderID',id,{Status:'CLOSED',LastUpdate:now_()});
  cascadeW_(id,'CLOSED');

  try{
    const wBundle = getWorkOrderBundle_(id);
    if(wBundle){
      const mail   = composeWorkOrderEmailBundle_(wBundle,'Closed','Thank you for working with OBSI.');
      const tCtx   = wBundle.parentTicket || {
        MarketID    : wBundle.workOrder.MarketID,
        ContactEmail: requesterEmailFromTicket_(wBundle.parentTicket||{})
      };
      const routed = routeEmailForTicket_(tCtx);
      const sent   = sendEmail_(routed.to, routed.cc, mail.subject, mail.text, mail.html);
      if(sent.ok) writeByKey_(SH_WO,'WorkOrderID',id,{EmailSent:now_(),LastUpdate:now_()});
    }
  }catch(_){}

  try{
    const allWOs = workOrders_();
    const woRow = allWOs.find(x => String(x.WorkOrderID||'').trim() === String(id||'').trim());
    if (woRow && woRow.TicketID){
      const t = tickets_().find(x => eqTicketId_(x.TicketID, woRow.TicketID));
      if (t && String(t.Status||'').toUpperCase() !== 'CLOSED'){
        writeByKey_(SH_TKT,'TicketID',t.TicketID,{Status:'CLOSED',LastUpdate:now_()});
        cascadeT_(t.TicketID,'CLOSED');
      }
    }
  }catch(_){}

  return {ok:true, notice:'Closed Work Order '+id};
}

// Cascades / soft delete
function cascadeT_(tid,status){
  const s=tryGetSheet_(SH_TREC); if(!s) return;
  const h=hdr_(s); if(!h.includes('Status')) return;
  const v=s.getDataRange().getValues(),cT=h.indexOf('TicketID'),cS=h.indexOf('Status');
  v.forEach((r,i)=>{
    if(i && String(r[cT])===String(tid)){
      const cur=String(r[cS]||'').toUpperCase();
      if(cur!=='DELETED') s.getRange(i+1,cS+1).setValue(status);
    }
  });
}
function cascadeW_(wid,status){
  const s=tryGetSheet_(SH_WREC); if(!s) return;
  const h=hdr_(s); if(!h.includes('Status')) return;
  const v=s.getDataRange().getValues(),cW=h.indexOf('WorkOrderID'),cS=h.indexOf('Status');
  v.forEach((r,i)=>{
    if(i && String(r[cW])===String(wid)){
      const cur=String(r[cS]||'').toUpperCase();
      if(cur!=='DELETED') s.getRange(i+1,cS+1).setValue(status);
    }
  });
}

function deleteSupportRecord(ts){
  const s=tryGetSheet_(SH_TREC); if(!s) return {ok:true,notice:'No-op'};
  const h=hdr_(s),v=s.getDataRange().getValues(); if(v.length<2)return{ok:true,notice:'No-op'};
  const cTS=h.indexOf('Timestamp'); if(cTS<0)return{ok:true,notice:'No-op'};
  const r0=v.findIndex((row,i)=>i>0&&String(row[cTS])===String(ts)); if(r0<1)return{ok:true,notice:'No-op'};
  const cS=h.indexOf('Status'),cN=h.indexOf('SupportProvided');
  if(cS>=0)s.getRange(r0+1,cS+1).setValue('DELETED');
  else if(cN>=0){
    const old=String(v[r0][cN]||'');
    if(old.indexOf('[DELETED]')!==0)s.getRange(r0+1,cN+1).setValue('[DELETED] '+old);
  }
  return {ok:true, notice:'Deleted Ticket record'};
}
function deleteWorkOrderRecord(wrid){
  const s=tryGetSheet_(SH_WREC); if(!s) return {ok:true,notice:'No-op'};
  const h=hdr_(s),v=s.getDataRange().getValues(); if(v.length<2)return{ok:true,notice:'No-op'};
  const cID=h.indexOf('WorkRecordID'); if(cID<0)return{ok:true,notice:'No-op'};
  const r0=v.findIndex((row,i)=>i>0&&String(row[cID])===String(wrid)); if(r0<1)return{ok:true,notice:'No-op'};
  const cS=h.indexOf('Status'),cN=h.indexOf('WorkPerformed');
  if(cS>=0)s.getRange(r0+1,cS+1).setValue('DELETED');
  else if(cN>=0){
    const old=String(v[r0][cN]||'');
    if(old.indexOf('[DELETED]')!==0)s.getRange(r0+1,cN+1).setValue('[DELETED] '+old);
  }
  return {ok:true, notice:'Deleted Work Order record'};
}

// ===== Ticket & Work Order "hard delete" (for ASP toolbar) =====
function softDeleteTicket(ticketId){
  const tidNorm = String(ticketId||'').trim();
  if(!tidNorm) throw new Error('TicketID required.');

  const sT = sh_(SH_TKT);
  const vT = sT.getDataRange().getValues();
  if(vT.length<2) return {ok:true, notice:'No tickets to delete'};
  const hT = vT[0].map(String);
  const cTid = hT.indexOf('TicketID');
  if(cTid<0) throw new Error('SupportTickets missing TicketID column');

  let deleted = false;
  for(let r=vT.length-1;r>=1;r--){
    if(String(vT[r][cTid]) === tidNorm){
      sT.deleteRow(r+1);
      deleted = true;
      break;
    }
  }

  try { cascadeT_(tidNorm,'DELETED'); } catch(_){}

  try{
    const sWO = sh_(SH_WO);
    const vWO = sWO.getDataRange().getValues();
    if(vWO.length>1){
      const hWO = vWO[0].map(String);
      const cTT = hWO.indexOf('TicketID');
      const cWO = hWO.indexOf('WorkOrderID');
      if(cTT>=0 && cWO>=0){
        const woIds = [];
        for(let r=1;r<vWO.length;r++){
          if(String(vWO[r][cTT])===tidNorm){
            const wid = String(vWO[r][cWO]||'');
            if(wid) woIds.push(wid);
          }
        }
        for(let r=vWO.length-1;r>=1;r--){
          if(String(vWO[r][cTT])===tidNorm){
            sWO.deleteRow(r+1);
          }
        }
        woIds.forEach(function(wid){ cascadeW_(wid,'DELETED'); });
      }
    }
  }catch(_){}

  return {ok:true, notice: deleted ? ('Ticket '+tidNorm+' deleted') : 'Ticket not found'};
}

function softDeleteWorkOrder(woId){
  const widNorm = String(woId||'').trim();
  if(!widNorm) throw new Error('WorkOrderID required.');

  const sWO = sh_(SH_WO);
  const v   = sWO.getDataRange().getValues();
  if(v.length<2) return {ok:true, notice:'No work orders to delete'};
  const h   = v[0].map(String);
  const cWO = h.indexOf('WorkOrderID');
  if(cWO<0) throw new Error('WorkOrders missing WorkOrderID column');

  let deleted = false;
  for(let r=v.length-1;r>=1;r--){
    if(String(v[r][cWO]) === widNorm){
      sWO.deleteRow(r+1);
      deleted = true;
      break;
    }
  }

  try { cascadeW_(widNorm,'DELETED'); } catch(_){}

  return {ok:true, notice: deleted ? ('Work Order '+widNorm+' deleted') : 'Work Order not found'};
}

// ---------- Manual engineer email ----------
function getEngineersPicker(excludeEngineerID){ return getEngineersPickerCore_(excludeEngineerID); }
function sendEngineerEmail(woId){ return sendEngineerEmailWithExtras(woId, []); }
function sendEngineerEmailWithExtras(woId, extraEngineerIds){
  const wBundle = getWorkOrderBundle_(woId); if(!wBundle) throw new Error('WorkOrder not found');
  const w = wBundle.workOrder;
  const assigned = w.EngineerID ? getEngineerById_(w.EngineerID) : null;
  const to = (assigned&&(assigned.EngineerEmail||assigned.Email))?String(assigned.EngineerEmail||assigned.Email):'';
  if(!to) throw new Error('Assigned engineer email not found');
  const ccExtras = (extraEngineerIds||[]).map(id=>engineerEmailById_(id)).filter(Boolean);

  const pairs = [
    ['Work Order', w.WorkOrderID],
    w.TicketID?['From Ticket', w.TicketID]:null,
    ['Client', w.ClientName], ['Market', w.MarketName], ['Station', String(w.CallSign||'').toUpperCase()],
    ['Engineer', assigned?assigned.EngineerName:w.EngineerName],
    ['Scheduled', w.ScheduledDate?fmtDateMDY_(new Date(w.ScheduledDate)):'(not scheduled)']
  ].filter(Boolean);

  const mail = (function friendlyEmailPlus(subjectCore, heading, introLines, detailsPairs, outro){
    const lines=[];
    lines.push(norm_(heading),'');
    (introLines||[]).forEach(x=>{ if(x) lines.push(norm_(x)); });
    if(detailsPairs && detailsPairs.length){
      lines.push('');
      detailsPairs.forEach(([k,v])=>{
        if(v!=null && v!=='') lines.push(k+': '+norm_(String(v)));
      });
    }
    lines.push('', norm_(outro||'-- Offerdahl Broadcast Service, Inc.'));
    const text=lines.join('\n');

    const rows=(detailsPairs||[]).filter(([_,v])=>v!=null&&v!=='')
      .map(([k,v])=>'<tr><td style="padding:2px 8px 2px 0;color:#444"><strong>'+esc_(k)+'</strong></td><td style="padding:2px 0">'+esc_(String(v))+'</td></tr>').join('');
    const html='<div style="font:14px/1.35 -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Arial,sans-serif">'
      +'<h3 style="margin:0 0 8px 0">'+esc_(heading)+'</h3>'
      +(introLines && introLines.length ? '<p>'+introLines.map(esc_).join('<br/>')+'</p>' : '')
      +(rows ? '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin:8px 0">'+rows+'</table>' : '')
      +'<p style="color:#555">Offerdahl Broadcast Service, Inc.</p>'
      +'</div>';
    return { subject: norm_(subjectCore), text, html };
  })(
    'Work Order Assigned — '+String(w.CallSign||'').toUpperCase()+' / '+String(w.ClientName||''),
    'Hello '+(assigned?assigned.EngineerName:'Engineer')+',',
    ['You have been assigned a new work order.','', 'Work Description:', norm_(w.WorkOrderText||'')],
    pairs,
    'Thank you.'
  );

  const res = sendEmail_([to], ccExtras, mail.subject, mail.text, mail.html);
  if(res.ok){ try{ writeByKey_(SH_WO,'WorkOrderID',woId,{EmailSent:now_(),LastUpdate:now_()}); }catch(_){} }
  return {ok:true, sentTo:to, cc:res.cc||[], notice:'Engineer email sent for '+woId};
}

// ---------- Itinerary ----------
function getItineraryEngineers(){
  return engineers_().filter(e=>String(e.EngineerID||'')&&String(e.EngineerName||''))
    .map(e=>({EngineerID:String(e.EngineerID),EngineerName:String(e.EngineerName),Email:String(e.EngineerEmail||e.Email||'')}))
    .sort((a,b)=>a.EngineerName.localeCompare(b.EngineerName));
}
function emailEngineerItinerary(p){
  const o=p||{}, engineerId=o.engineerId, startDateISO=o.startDateISO, endDateISO=o.endDateISO,
        toEmail=o.toEmail, ccEmail=o.ccEmail, bccEmail=o.bccEmail, extraEngineerIds=o.extraEngineerIds;
  if(!engineerId) throw new Error('EngineerID required.');
  if(!startDateISO||!endDateISO) throw new Error('Start and end dates required.');
  if(!toEmail) throw new Error('Recipient (To) required.');

  const s=sh_(SH_WO),v=s.getDataRange().getValues(); if(!v.length) throw new Error('WorkOrders empty');
  const h=v[0].map(x=>String(x||'').trim()), cWO=h.indexOf('WorkOrderID'),cCN=h.indexOf('ClientName'),cMN=h.indexOf('MarketName'),
        cCS=h.indexOf('CallSign'),cEID=h.indexOf('EngineerID'),cEN=h.indexOf('EngineerName'),cTXT=h.indexOf('WorkOrderText'),
        cST=h.indexOf('Status'),cSD=h.indexOf('ScheduledDate');
  if([cWO,cCN,cMN,cCS,cEID,cEN,cTXT,cST,cSD].some(i=>i<0)) throw new Error('WorkOrders headers are not canonical.');

  const start=dateOnlyFromISO_(startDateISO), end=new Date(dateOnlyFromISO_(endDateISO).getTime()+86399999);
  const items=[]; let engName='';
  for(let r=1;r<v.length;r++){
    const row=v[r]; if(String(row[cEID]).trim()!==String(engineerId).trim()) continue;
    let d=row[cSD]; if(!d) continue;
    if(!(d instanceof Date)){
      const tryD=new Date(d);
      if(!isNaN(tryD)) d=tryD;
      else if(typeof d==='number') d=new Date(Math.round((d-25569)*86400*1000));
      else continue;
    }
    if(d<start||d>end) continue;
    engName=engName||String(row[cEN]||'').trim();
    items.push({
      Scheduled  : fmtDateMDY_(d),
      WorkOrderID: row[cWO],
      ClientName : row[cCN],
      MarketName : row[cMN],
      CallSign   : row[cCS],
      Status     : row[cST],
      WorkOrderText: row[cTXT]
    });
  }
  items.sort((a,b)=>a.Scheduled.localeCompare(b.Scheduled)
    || String(a.ClientName||'').localeCompare(String(b.ClientName||''))
    || String(a.CallSign||'').localeCompare(String(b.CallSign||'')));

  const lines = [];
  lines.push('Engineer: '+(engName||engineerId));
  lines.push('Date Range: '+fmtIsoMDY_(startDateISO)+' – '+fmtIsoMDY_(endDateISO),'');
  if(!items.length){
    lines.push('No scheduled work orders in this range.');
  } else {
    lines.push('Schedule:');
    items.forEach(it=>{
      lines.push('- '+it.Scheduled+' · WO '+it.WorkOrderID+' · '+String(it.CallSign||'').toUpperCase()+' · '+it.ClientName+' ('+it.MarketName+')');
      const summary = norm_(String(it.WorkOrderText||'')).slice(0,220);
      if(summary) lines.push('  '+summary);
    });
  }
  const subject = 'Itinerary — '+(engName||engineerId)+' — '+fmtIsoMDY_(startDateISO)+' – '+fmtIsoMDY_(endDateISO);
  const body    = lines.join('\n');

  const ccMergedArr = [ (ccEmail||'').trim(), ...((extraEngineerIds||[]).map(id=>engineerEmailById_(id)).filter(Boolean)) ].filter(Boolean);
  const bccArr      = (bccEmail && String(bccEmail).trim()) ? [String(bccEmail).trim()] : [];

  const res = sendEmail_([toEmail], ccMergedArr, subject, body, null, bccArr);
  return {ok: !!res.ok, count: items.length, subject: subject, engineerName:(engName||engineerId), notice:'Itinerary sent ('+items.length+')'};
}