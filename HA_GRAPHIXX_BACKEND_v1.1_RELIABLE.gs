/**
 * HA.GRAPHIXX ERP - Google Apps Script Backend API
 * 
 * Arahan pemasangan:
 * 1. Buka https://script.google.com -> New Project
 * 2. Copy semua kod ini ke dalam Code.gs
 * 3. Buka Google Sheet anda (atau buat baru)
 * 4. Dapatkan Spreadsheet ID dari URL
 * 5. Tukar SPREADSHEET_ID di bawah dengan ID anda
 * 6. Save -> Run setupSheets() (authorize bila diminta)
 * 7. Deploy -> New deployment -> Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 8. Copy Web App URL dan letak di dalam HTML (APPS_SCRIPT_URL)
 * 9. URL live: https://script.google.com/macros/s/AKfycbzBF7v4xeP0CYS9BLmXRorV9g3fUr_nLuDIbYRpVAzd9e-wIY163P5TPjYZIqSA5HEowA/exec
 */

const SPREADSHEET_ID = '1EI9gpXc6bKu-KxlJ3UlqAShizH3yEhwfWIcJ1uYRPVM';
const APP_VERSION = '1.1.0-reliable';
const SESSION_DAYS = 7;
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
const REQUIRED_SHEETS = ['Users','Jobs','Orders','Settings','Sessions','PaymentHistory','Files','AuditLog','ErrorLog'];

// ==========================================
// SETUP - Cipta sheet automatik
// ==========================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  let usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
    usersSheet.appendRow(['User ID','Password','Role','Name','Agent Code','Phone','Email','Address','Active','Created At','Updated At','Must Change Password']);
  } else {
    ensureColumns_(usersSheet, ['User ID','Password','Role','Name','Agent Code','Phone','Email','Address','Active','Created At','Updated At','Must Change Password']);
  }

  const usersData = usersSheet.getDataRange().getValues();
  let hasAdmin = false;
  for (let i = 1; i < usersData.length; i++) {
    if (String(usersData[i][2]).toLowerCase() === 'admin' && String(usersData[i][8]).toUpperCase() === 'TRUE') { hasAdmin = true; break; }
  }
  let adminMsg = 'Admin sedia ada dikekalkan.';
  if (!hasAdmin) {
    const randomPw = generateSecureToken_('pw').replace(/[^A-Za-z0-9]/g,'').slice(0,16);
    usersSheet.appendRow(['admin', hashPW(randomPw), 'admin', 'Administrator', '', '', 'admin@hagraphixx.local', '', 'TRUE', new Date(), new Date(), 'TRUE']);
    adminMsg = 'ADMIN BARU DICIPTA. Password sementara: ' + randomPw + ' — tukar selepas login pertama.';
  }

  let jobsSheet = ss.getSheetByName('Jobs');
  const jobHeaders = ['Job ID','Nama Job','No. Invois','Status','Jumlah Helai','Jumlah (RM)','Konfigurasi Harga JSON','Cost JSON','Dicipta Pada','Dikemaskini Pada','Dicipta Oleh','Role Pencipta','Agent Code','Customer Name','Customer Contact','Brand JSON','Notes','Deposit (RM)','Paid Amount (RM)','Balance (RM)','Payment Status','Currency','FX Rate','Locale','Discount (RM)','Tax Rate (%)','Shipping (RM)'];
  if (!jobsSheet) { jobsSheet = ss.insertSheet('Jobs'); jobsSheet.appendRow(jobHeaders); }
  else ensureColumns_(jobsSheet, jobHeaders);

  let ordersSheet = ss.getSheetByName('Orders');
  const orderHeaders = ['Job ID','Bil','Lengan','Saiz','Nama','Kolar','No/Remarks','Harga Unit (RM)','Base Price (RM)','Tarikh Disimpan','Status Item'];
  if (!ordersSheet) { ordersSheet = ss.insertSheet('Orders'); ordersSheet.appendRow(orderHeaders); }
  else ensureColumns_(ordersSheet, orderHeaders);

  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) { settingsSheet = ss.insertSheet('Settings'); settingsSheet.appendRow(['Kunci','Nilai','Dikemaskini Pada']); }
  updateSettingRow(settingsSheet, 'app_version', APP_VERSION, new Date().toLocaleString('en-GB'));
  const existingSettings = settingsSheet.getDataRange().getValues();
  if (!existingSettings.some(r => r[0] === 'currency_rates')) {
    updateSettingRow(settingsSheet, 'currency_rates', JSON.stringify({RM:1,SGD:0.30,USD:0.22}), new Date().toLocaleString('en-GB'));
  }

  ensureSheet_(ss, 'Sessions', ['Token','User ID','Role','Name','Created At','Expires At']);
  ensureSheet_(ss, 'PaymentHistory', ['Date','Job ID','Amount','Type','Method','Reference','Note','Recorded By','Transaction ID']);
  ensureSheet_(ss, 'Files', ['Uploaded At','Job ID','File Type','Original Name','Drive File ID','Drive URL','MIME Type','Size Bytes','Uploaded By']);
  ensureSheet_(ss, 'AuditLog', ['Timestamp','User ID','Action','Job ID','Detail']);
  ensureSheet_(ss, 'ErrorLog', ['Timestamp','Action','User ID','Message','Stack']);

  return 'Setup selesai untuk HA.GRAPHIXX ERP ' + APP_VERSION + '. ' + adminMsg;
}

// HELPER: Cipta user baru dengan password hashed.
// Run manual dari editor: adminCreateUser('agen001', 'PasswordKuat123', 'agen', 'Agen Ali', 'AG001')
function adminCreateUser(userId, password, role, name, agentCode) {
  if (!userId || !password || password.length < 8) return 'Gagal: password mesti sekurang-kurangnya 8 aksara';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase() === String(userId).toLowerCase()) return 'Gagal: user ID sudah wujud';
  }
  sheet.appendRow([userId, hashPW(password), role || 'customer', name || userId, agentCode || '', '', '', '', 'TRUE', new Date(), new Date(), 'TRUE']);
  return 'User dicipta: ' + userId + ' (' + (role || 'customer') + ')';
}

// HELPER: Disable semua demo user lama (admin123/agen123/cus123) yang mungkin
// masih ada dalam sheet dari setup versi lama. Run SEKALI dari editor sebelum live.
function disableDemoUsers() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  const data = sheet.getDataRange().getValues();
  const demoPWs = [sha256Hex_('admin123'), sha256Hex_('agen123'), sha256Hex_('cus123'), 'admin123', 'agen123', 'cus123'];
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (demoPWs.indexOf(String(data[i][1])) !== -1) {
      sheet.getRange(i+1, 9).setValue('FALSE'); // Active = FALSE
      count++;
    }
  }
  return count + ' demo user di-disable. Cipta user sebenar guna adminCreateUser().';
}

// ==========================================
// WEB APP ENTRY POINTS
// ==========================================
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let action = 'ping';
  let params = {};
  
  // Handle GET and POST form-urlencoded (params come in e.parameter)
  if (e.parameter.action) {
    action = e.parameter.action;
    params = e.parameter;
  } else if (e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      action = body.action || action;
      params = body;
    } catch(err) {
      params = e.parameter;
    }
  }
  
  // Endpoints that DON'T require auth
  const publicEndpoints = ['ping', 'login', 'portalLogin'];
  const adminOnlyEndpoints = ['deleteJob','updateJobStatus','updateOrderStatus','saveQuote','saveSettings','updatePayment','getAuditLog','createBackup','installDailyBackup','healthCheck','exportData','adminResetPassword','updateInvoiceAdjustments','exportReport'];
  
  // Check auth for non-public endpoints
  if (publicEndpoints.indexOf(action) === -1) {
    const token = String(params.token || '');
    if (!token) {
      return jsonOut({ ok: false, error: 'Authentication required' });
    }
    
    // Verify token from Sessions sheet (or legacy portal token)
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sessionUser = null;
    
    // Fast path: short-lived script cache. Sessions sheet remains source of truth.
    const sessionCache = CacheService.getScriptCache();
    const cachedSession = sessionCache.get('session:' + token);
    if (cachedSession) {
      try { sessionUser = JSON.parse(cachedSession); } catch(ignore) { sessionUser = null; }
    }

    // Sessions sheet fallback
    let sessionsSheet = ss.getSheetByName('Sessions');
    if (!sessionsSheet) {
      sessionsSheet = ss.insertSheet('Sessions');
      sessionsSheet.appendRow(['Token', 'User ID', 'Role', 'Name', 'Created At', 'Expires At']);
    }
    const sessionsData = sessionsSheet.getDataRange().getValues();
    const now = new Date();
    for (let i = 1; !sessionUser && i < sessionsData.length; i++) {
      if (String(sessionsData[i][0]) === token) {
        const expires = new Date(sessionsData[i][5]);
        if (expires > now) {
          sessionUser = { id: String(sessionsData[i][1]), role: String(sessionsData[i][2]), name: String(sessionsData[i][3] || '') };
          sessionCache.put('session:' + token, JSON.stringify(sessionUser), 21600);
        }
        break;
      }
    }
    
    // NOTE: Legacy portal token (portal-USERID-TIMESTAMP) TELAH DIBUANG.
    // Semua token mesti datang dari Sessions sheet (secure random, 7-day expiry).
    // User portal lama yang masih pegang token legacy akan dipaksa login semula.
    
    if (!sessionUser) {
      return jsonOut({ ok: false, error: 'Invalid or expired session' });
    }
    
    // Inject verified user info into params
    params.userId = sessionUser.id;
    params.userRole = sessionUser.role;
    params.userName = sessionUser.name;

    // Enforce temporary-password change at API level, not UI only.
    const passwordChangeAllowed = ['changePassword','portalChangePassword','logout'];
    if (passwordChangeAllowed.indexOf(action) === -1) {
      const usersSheetForPolicy = ss.getSheetByName('Users');
      if (usersSheetForPolicy) {
        const policyRows = usersSheetForPolicy.getDataRange().getValues();
        for (let p = 1; p < policyRows.length; p++) {
          if (String(policyRows[p][0]) === String(sessionUser.id)) {
            if (String(policyRows[p][11] || '').toUpperCase() === 'TRUE') {
              return jsonOut({ ok:false, error:'PASSWORD_CHANGE_REQUIRED', mustChangePassword:true });
            }
            break;
          }
        }
      }
    }
    
    // Admin-only check
    if (adminOnlyEndpoints.indexOf(action) !== -1 && sessionUser.role !== 'admin') {
      return jsonOut({ ok: false, error: 'Admin access required' });
    }
  }
  
  let result;
  try {
    switch(action) {
      case 'ping':         result = { ok: true, message: 'HA.GRAPHIXX ERP API Active' }; break;
      case 'login':        result = apiLogin(params); break;
      case 'getDashboard': result = apiGetDashboard(params); break;
      case 'getJobs':      result = apiGetJobs(params); break;
      case 'getJob':       result = apiGetJob(params); break;
      case 'saveJob':      result = apiSaveJob(params); break;
      case 'deleteJob':    result = apiDeleteJob(params); break;
      case 'updateOrderStatus': result = apiUpdateOrderStatus(params); break;
      case 'updateJobStatus':   result = apiUpdateJobStatus(params); break;
      case 'saveQuote':    result = apiSaveQuote(params); break;
      case 'getSettings':  result = apiGetSettings(params); break;
      case 'saveSettings': result = apiSaveSettings(params); break;
      // PORTAL endpoints (customer-facing)
      case 'portalLogin':         result = apiPortalLogin(params); break;
      case 'portalBootstrap':     result = apiPortalBootstrap(params); break;
      case 'portalSubmitRequest': result = apiPortalSubmitRequest(params); break;
      case 'portalUpdateProfile': result = apiPortalUpdateProfile(params); break;
      case 'portalChangePassword':result = apiPortalChangePassword(params); break;
      case 'logout':       invalidateSession(String(params.token||'')); result = { ok: true, message: 'Logged out' }; break;
      case 'getAuditLog':  result = apiGetAuditLog(params); break;
      case 'updatePayment': result = apiUpdatePayment(params); break;
      case 'quoteReply':   result = apiQuoteReply(params); break;
      case 'getPaymentHistory': result = apiGetPaymentHistory(params); break;
      case 'changePassword': result = apiChangePassword(params); break;
      case 'uploadFile': result = apiUploadFile(params); break;
      case 'getFiles': result = apiGetFiles(params); break;
      case 'downloadFile': result = apiDownloadFile(params); break;
      case 'createBackup': result = apiCreateBackup(params); break;
      case 'installDailyBackup': result = apiInstallDailyBackup(params); break;
      case 'healthCheck': result = apiHealthCheck(params); break;
      case 'exportData': result = apiExportData(params); break;
      case 'getJobOptions': result = apiGetJobOptions(params); break;
      case 'adminResetPassword': result = apiAdminResetPassword(params); break;
      case 'updateInvoiceAdjustments': result = apiUpdateInvoiceAdjustments(params); break;
      case 'exportReport': result = apiExportReport(params); break;
      default:             result = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch(err) {
    logError_(action, String(params.userId || ''), err);
    result = { ok: false, error: cleanError_(err) };
  }
  
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// CORS preflight handler
function doGetOptions(e) {
  return ContentService
    .createTextOutput('')
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper: output JSON
function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Helper: create session token (7-day expiry)
function createSession(userId, role, name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sessionsSheet = ss.getSheetByName('Sessions');
  if (!sessionsSheet) {
    sessionsSheet = ss.insertSheet('Sessions');
    sessionsSheet.appendRow(['Token', 'User ID', 'Role', 'Name', 'Created At', 'Expires At']);
  }
  // Generate secure token
  const randBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, userId + Date.now() + Math.random());
  const token = 'sess-' + randBytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').slice(0, 32);
  const now = new Date();
  const expires = new Date(now.getTime() + SESSION_DAYS * 24 * 60 * 60 * 1000);
  sessionsSheet.appendRow([token, userId, role, name, now, expires]);
  CacheService.getScriptCache().put('session:' + token, JSON.stringify({id:String(userId),role:String(role),name:String(name||'')}), 21600);
  // Clean expired sessions (keep last 100 rows)
  const data = sessionsSheet.getDataRange().getValues();
  if (data.length > 100) {
    for (let i = data.length - 1; i >= 1; i--) {
      const exp = new Date(data[i][5]);
      if (exp < now) sessionsSheet.deleteRow(i + 1);
    }
  }
  return token;
}

// Helper: invalidate session (logout)
function invalidateSession(token) {
  CacheService.getScriptCache().remove('session:' + String(token||''));
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sessionsSheet = ss.getSheetByName('Sessions');
  if (!sessionsSheet) return;
  const data = sessionsSheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]) === token) {
      sessionsSheet.deleteRow(i + 1);
      break;
    }
  }
}

// Helper: sanitize text to prevent XSS
function sanitizeText(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// ==========================================
// SECURITY - password hashing (SHA-256)
// New passwords are stored as hashes. Comparison also accepts legacy
// plain-text values so existing sheets keep working.
// ==========================================
function hashPW(pw) {
  try {
    const salt = generateSecureToken_('salt').slice(0, 24);
    let value = salt + '|' + String(pw);
    for (let i = 0; i < 1200; i++) value = sha256Hex_(value);
    return 'v2$' + salt + '$' + value;
  } catch (e) {
    return sha256Hex_(String(pw));
  }
}
function pwMatches(stored, input) {
  stored = String(stored || '');
  input = String(input || '');
  if (!stored || !input) return false;
  if (stored.indexOf('v2$') === 0) {
    const parts = stored.split('$');
    if (parts.length !== 3) return false;
    let value = parts[1] + '|' + input;
    for (let i = 0; i < 1200; i++) value = sha256Hex_(value);
    return timingSafeEqual_(parts[2], value);
  }
  const legacy = sha256Hex_(input);
  return timingSafeEqual_(stored, legacy) || timingSafeEqual_(stored, input);
}

// ==========================================
// API FUNCTIONS
// =========================================

// LOGIN
function apiLogin(params) {
  const userId = String(params.userId || '').toLowerCase().trim();
  const password = String(params.password || '');
  const role = String(params.role || '');
  
  if (!userId || !password) return { ok: false, error: 'ID and password required' };
  if (isLoginLocked_(userId)) return {ok:false,error:'Too many failed attempts. Try again in 15 minutes.'};
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { ok: false, error: 'Users sheet not found' };
  
  const data = sheet.getDataRange().getValues();
  // Columns: User ID(0), Password(1), Role(2), Name(3), Agent Code(4), Phone(5), Email(6), Address(7), Active(8)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = String(row[0]).toLowerCase();
    const rowPw = String(row[1]);
    const userRole = String(row[2]);
    
    // Match by User ID or Email
    if ((rowId === userId || String(row[6] || '').toLowerCase() === userId) && pwMatches(rowPw, password)) {
      if (role && role !== userRole) continue;
      const active = String(row[8]).toUpperCase() === 'TRUE';
      if (!active) return { ok: false, error: 'Account not active' };
      // Generate session token (7-day expiry)
      // Transparently migrate legacy SHA/plain passwords to v2 salted hash.
      if (String(rowPw).indexOf('v2$') !== 0) sheet.getRange(i+1, 2).setValue(hashPW(password));
      clearLoginFailures_(userId);
      const token = createSession(row[0], userRole, row[3] || '');
      return {
        ok: true,
        token: token,
        appVersion: APP_VERSION,
        mustChangePassword: String(row[11] || '').toUpperCase() === 'TRUE',
        permissions: getPermissions_(userRole),
        user: {
          id: row[0], role: userRole, name: row[3], agentCode: row[4],
          phone: row[5] || '', email: row[6] || '', address: row[7] || ''
        }
      };
    }
  }
  recordLoginFailure_(userId);
  return { ok: false, error: 'Wrong ID, password or role' };
}

// DASHBOARD - ringkasan statistik (role-filtered)
function apiGetDashboard(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const ordersSheet = ss.getSheetByName('Orders');
  const userId = String(params.userId || '');
  const userRole = String(params.userRole || '');
  const isAdmin = userRole === 'admin';
  const canSeeRevenue = isAdmin || userRole === 'agen';
  const jobsData = jobsSheet ? jobsSheet.getDataRange().getValues() : [];
  const ordersData = ordersSheet ? ordersSheet.getDataRange().getValues() : [];

  let totalJobs = 0, totalHelai = 0, totalRevenue = 0, activeJobs = 0, totalOrders = 0;
  const statusCounts = {};
  const closed = ['completed','delivered','cancelled','rejected'];
  const ownedJobIds = {};

  for (let i = 1; i < jobsData.length; i++) {
    if (!isAdmin && String(jobsData[i][10] || '') !== userId) continue;
    const jobId = String(jobsData[i][0] || '');
    ownedJobIds[jobId] = true;
    totalJobs++;
    totalHelai += parseInt(jobsData[i][4],10) || 0;
    if (canSeeRevenue) totalRevenue += parseFloat(jobsData[i][5]) || 0;
    const status = String(jobsData[i][3] || 'submitted').toLowerCase();
    if (closed.indexOf(status) === -1) activeJobs++;
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  }
  for (let i = 1; i < ordersData.length; i++) {
    if (isAdmin || ownedJobIds[String(ordersData[i][0] || '')]) totalOrders++;
  }

  const recentJobs = [];
  const recentCount = isAdmin ? 5 : 10;
  for (let i = jobsData.length - 1; i >= 1 && recentJobs.length < recentCount; i--) {
    if (!isAdmin && String(jobsData[i][10] || '') !== userId) continue;
    recentJobs.push({
      jobId: jobsData[i][0], namaJob: jobsData[i][1], noInvois: jobsData[i][2],
      status: jobsData[i][3], helai: jobsData[i][4],
      jumlah: canSeeRevenue ? (parseFloat(jobsData[i][5]) || 0) : 0,
      dicipta: formatDateSafe_(jobsData[i][8]), currency: jobsData[i][21] || 'RM', fxRate: parseFloat(jobsData[i][22]) || 1
    });
  }
  return { ok:true, appVersion:APP_VERSION, permissions:getPermissions_(userRole), stats:{totalJobs,totalHelai,totalRevenue,activeJobs,statusCounts,totalOrders}, recentJobs };
}

// GET JOBS - senarai semua job (role-filtered)
function apiGetJobs(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Jobs');
  if (!sheet) return { ok:false, error:'Sheet Jobs tidak wujud' };
  const userId = String(params.userId || '');
  const userRole = String(params.userRole || '');
  const isAdmin = userRole === 'admin';
  const search = String(params.search || '').trim().toLowerCase();
  const statusFilter = String(params.status || 'all').toLowerCase();
  const sort = String(params.sort || 'newest').toLowerCase();
  const page = Math.max(1, parseInt(params.page,10) || 1);
  const pageSize = Math.min(100, Math.max(5, parseInt(params.pageSize,10) || 25));
  const data = sheet.getDataRange().getValues();
  const jobs = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (!isAdmin && String(data[i][10] || '') !== userId) continue;
    const status = String(data[i][3] || 'submitted').toLowerCase();
    if (statusFilter !== 'all' && status !== statusFilter) continue;
    const haystack = [data[i][0],data[i][1],data[i][2],data[i][10],data[i][13],data[i][14]].join(' ').toLowerCase();
    if (search && haystack.indexOf(search) === -1) continue;
    jobs.push({
      jobId:data[i][0], namaJob:data[i][1], noInvois:data[i][2], status:data[i][3],
      helai:parseInt(data[i][4],10)||0, jumlah:parseFloat(data[i][5])||0, config:data[i][6],
      dicipta:formatDateTimeSafe_(data[i][8]), dikemaskini:formatDateTimeSafe_(data[i][9]),
      diciptaOleh:data[i][10]||'', rolePencipta:data[i][11]||'', customerName:data[i][13]||'', customerContact:data[i][14]||'',
      deposit:parseFloat(data[i][17])||0, paidAmount:parseFloat(data[i][18])||0, balance:parseFloat(data[i][19])||0,
      payStatus:data[i][20]||'unpaid', currency:data[i][21]||'RM', fxRate:parseFloat(data[i][22])||1, locale:data[i][23]||'en-MY', discount:parseFloat(data[i][24])||0, taxRate:parseFloat(data[i][25])||0, shipping:parseFloat(data[i][26])||0
    });
  }
  if(sort==='oldest') jobs.reverse();
  else if(sort==='amount_desc') jobs.sort(function(a,b){return b.jumlah-a.jumlah;});
  else if(sort==='amount_asc') jobs.sort(function(a,b){return a.jumlah-b.jumlah;});
  else if(sort==='name') jobs.sort(function(a,b){return String(a.namaJob).localeCompare(String(b.namaJob));});
  const total = jobs.length;
  const totalPages = Math.max(1, Math.ceil(total/pageSize));
  const safePage = Math.min(page,totalPages);
  const start = (safePage-1)*pageSize;
  return { ok:true, jobs:jobs.slice(start,start+pageSize), pagination:{page:safePage,pageSize,total,totalPages}, permissions:getPermissions_(userRole) };
}

// GET JOB - dapatkan satu job + orders (ownership-checked)
function apiGetJob(params) {
  const jobId = String(params.jobId || '');
  const userId = String(params.userId || '');
  const userRole = String(params.userRole || '');
  const isAdmin = userRole === 'admin';
  if (!jobId) return { ok: false, error: 'jobId diperlukan' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const ordersSheet = ss.getSheetByName('Orders');
  
  // Cari job
  const jobsData = jobsSheet.getDataRange().getValues();
  let jobInfo = null;
  let jobOwner = '';
  for (let i = 1; i < jobsData.length; i++) {
    if (jobsData[i][0] === jobId) {
      jobOwner = String(jobsData[i][10] || '');
      // Ownership check: non-admin can only access own jobs
      if (!isAdmin && jobOwner !== userId) {
        return { ok: false, error: 'Access denied: not your job' };
      }
      jobInfo = {
        jobId: jobsData[i][0],
        namaJob: jobsData[i][1],
        noInvois: jobsData[i][2],
        status: jobsData[i][3],
        helai: jobsData[i][4],
        // Non-admin: don't expose total amount
      jumlah: jobsData[i][5],
        config: jobsData[i][6],
        dicipta: jobsData[i][8] ? new Date(jobsData[i][8]).toLocaleString('en-GB') : '',
        dikemaskini: jobsData[i][9] ? new Date(jobsData[i][9]).toLocaleString('en-GB') : '',
        diciptaOleh: jobsData[i][10] || '',
        rolePencipta: jobsData[i][11] || '',
        // Payment fields (columns 18-21, 0-indexed 17-20)
        deposit: parseFloat(jobsData[i][17]) || 0,
        paidAmount: parseFloat(jobsData[i][18]) || 0,
        balance: parseFloat(jobsData[i][19]) || 0,
        payStatus: String(jobsData[i][20] || 'unpaid'),
        currency: jobsData[i][21] || 'RM',
        fxRate: parseFloat(jobsData[i][22]) || 1,
        locale: jobsData[i][23] || 'en-MY',
        discount: parseFloat(jobsData[i][24]) || 0,
        taxRate: parseFloat(jobsData[i][25]) || 0,
        shipping: parseFloat(jobsData[i][26]) || 0
      };
      break;
    }
  }
  if (!jobInfo) return { ok: false, error: 'Job tidak dijumpai' };
  
  // Cari orders untuk job ini
  const ordersData = ordersSheet.getDataRange().getValues();
  const orders = [];
  for (let i = 1; i < ordersData.length; i++) {
    if (ordersData[i][0] === jobId) {
      orders.push({
        bil: ordersData[i][1],
        lengan: ordersData[i][2],
        saiz: ordersData[i][3],
        nama: ordersData[i][4],
        kolar: ordersData[i][5],
        remarks: ordersData[i][6],
        hargaUnit: ordersData[i][7],
        basePrice: ordersData[i][8] || ordersData[i][7],
        status: ordersData[i][10] || 'pending'
      });
    }
  }
  
  return { ok: true, job: jobInfo, orders };
}

// SAVE JOB - simpan job baru atau kemaskini
function apiSaveJob(params) {
  return withScriptLock_(function() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const jobsSheet = ss.getSheetByName('Jobs');
    const ordersSheet = ss.getSheetByName('Orders');
    const jobData = typeof params.job === 'string' ? JSON.parse(params.job) : (params.job || {});
    const ordersData = typeof params.orders === 'string' ? JSON.parse(params.orders) : (params.orders || []);
    const userId = String(params.userId || 'unknown');
    const userRole = String(params.userRole || '');
    if (!Array.isArray(ordersData) || ordersData.length < 1) return {ok:false,error:'At least one order item is required'};
    if (ordersData.length > 5000) return {ok:false,error:'Maximum 5000 order items per job'};

    let jobId = String(jobData.jobId || '');
    const isNew = !jobId;
    const jobsData = jobsSheet.getDataRange().getValues();
    let existingRow = -1;
    for (let i=1;i<jobsData.length;i++) if (String(jobsData[i][0])===jobId) { existingRow=i+1; break; }
    if (!isNew && existingRow < 0) return {ok:false,error:'Job not found'};
    if (!isNew && userRole !== 'admin' && String(jobsData[existingRow-1][10]||'').toLowerCase() !== userId.toLowerCase()) return {ok:false,error:'Access denied: anda hanya boleh edit job sendiri'};

    const existingStatus = !isNew ? String(jobsData[existingRow-1][3]||'submitted').toLowerCase() : 'submitted';
    const editableByOwnerStatuses=['submitted','reviewing','rejected'];
    if(!isNew && userRole!=='admin' && editableByOwnerStatuses.indexOf(existingStatus)===-1) return {ok:false,error:'This job is locked after quotation. Contact admin for changes.'};
    const preservePrices = !isNew && userRole !== 'admin' && existingStatus !== 'submitted';
    const existingPrices = {};
    if (preservePrices) {
      const oldOrders = ordersSheet.getDataRange().getValues();
      for (let i=1;i<oldOrders.length;i++) if (String(oldOrders[i][0])===jobId) existingPrices[parseInt(oldOrders[i][1],10)] = parseFloat(oldOrders[i][7])||0;
    }

    if (isNew) jobId = generateJobId_('JOB');
    const totalHelai = ordersData.length;
    let totalRM = 0;
    let finalTotal = 0;
    const now = new Date();
    const nowStr = now.toLocaleString('en-GB');
    const currency = ['RM','SGD','USD'].indexOf(String(jobData.currency||'')) >= 0 ? String(jobData.currency) : 'RM';
    const fxRate = Math.max(0.000001, parseFloat(jobData.fxRate)||1);
    const locale = String(jobData.locale||'en-MY').slice(0,20);
    const config = jobData.config ? (typeof jobData.config==='string' ? jobData.config : JSON.stringify(jobData.config)) : '';

    const orderRows = [];
    ordersData.forEach(function(o,idx){
      const oldBil = parseInt(o.bil,10)||0;
      const itemPrice = preservePrices && oldBil > 0 && existingPrices[oldBil] !== undefined ? existingPrices[oldBil] : Math.max(0,parseFloat(o.hargaUnit)||0);
      totalRM += itemPrice;
      orderRows.push([jobId,idx+1,limitText_(o.lengan||'Pendek',50),limitText_(o.saiz||'M',20),limitText_(String(o.nama||'').toUpperCase(),120),limitText_(o.kolar||'Round neck (RN)',80),limitText_(o.remarks||'-',250),itemPrice,Math.max(0,parseFloat(o.basePrice)||itemPrice),nowStr,limitText_(o.status||'pending',30)]);
    });

    if (isNew) {
      const invoice = limitText_(jobData.noInvois || ('INV-' + Utilities.formatDate(now,'Asia/Kuala_Lumpur','yyyyMMdd-HHmmss')),80);
      finalTotal = totalRM;
      jobsSheet.appendRow([jobId,limitText_(jobData.namaJob||'Untitled',150),invoice,userRole==='admin'?limitText_(jobData.status||'submitted',30):'submitted',totalHelai,totalRM,config,'',now,nowStr,userId,userRole,limitText_(params.agentCode||'',50),limitText_(jobData.customerName||'',150),limitText_(jobData.customerContact||'',150),jobData.brand?JSON.stringify(jobData.brand):'',limitText_(jobData.notes||'',500),0,0,totalRM,'unpaid',currency,fxRate,locale,0,0,0]);
    } else {
      const r=existingRow;
      jobsSheet.getRange(r,2).setValue(limitText_(jobData.namaJob||jobsData[r-1][1],150));
      jobsSheet.getRange(r,3).setValue(limitText_(jobData.noInvois||jobsData[r-1][2],80));
      const discount=Math.max(0,parseFloat(jobsData[r-1][24])||0),taxRate=Math.max(0,parseFloat(jobsData[r-1][25])||0),shipping=Math.max(0,parseFloat(jobsData[r-1][26])||0);
      const taxable=Math.max(0,totalRM-discount), grandTotal=taxable+(taxable*taxRate/100)+shipping;
      finalTotal = grandTotal;
      const alreadyPaid=(parseFloat(jobsData[r-1][17])||0)+(parseFloat(jobsData[r-1][18])||0);
      if(alreadyPaid>grandTotal+0.01)return {ok:false,error:'Updated total cannot be lower than amount already paid'};
      jobsSheet.getRange(r,5).setValue(totalHelai);
      jobsSheet.getRange(r,6).setValue(grandTotal);
      jobsSheet.getRange(r,7).setValue(config);
      jobsSheet.getRange(r,10).setValue(nowStr);
      jobsSheet.getRange(r,22).setValue(currency);
      jobsSheet.getRange(r,23).setValue(fxRate);
      jobsSheet.getRange(r,24).setValue(locale);
      const dep=parseFloat(jobsData[r-1][17])||0, paid=parseFloat(jobsData[r-1][18])||0;
      const bal=Math.max(0,grandTotal-dep-paid);
      jobsSheet.getRange(r,20).setValue(bal);
      jobsSheet.getRange(r,21).setValue(bal<=0?'paid':(dep+paid>0?'partial':'unpaid'));
    }

    const allOrders = ordersSheet.getDataRange().getValues();
    for (let i=allOrders.length-1;i>=1;i--) if (String(allOrders[i][0])===jobId) ordersSheet.deleteRow(i+1);
    if (orderRows.length) ordersSheet.getRange(ordersSheet.getLastRow()+1,1,orderRows.length,orderRows[0].length).setValues(orderRows);

    writeAuditLog(userId,isNew?'saveJob':'updateJob',jobId,(isNew?'Created':'Updated')+' '+totalHelai+' items, RM'+finalTotal.toFixed(2));
    return {ok:true,jobId,totalHelai,totalRM:finalTotal,subtotal:totalRM,status:isNew?'submitted':existingStatus,currency,fxRate};
  });
}

// DELETE JOB
function apiDeleteJobUnlocked_(params) {
  const jobId=String(params.jobId||'');if(!jobId)return {ok:false,error:'jobId diperlukan'};
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID),jobsSheet=ss.getSheetByName('Jobs'),ordersSheet=ss.getSheetByName('Orders');
  let found=false;const jobsData=jobsSheet.getDataRange().getValues();for(let i=jobsData.length-1;i>=1;i--)if(String(jobsData[i][0])===jobId){jobsSheet.deleteRow(i+1);found=true;}
  const ordersData=ordersSheet.getDataRange().getValues();for(let i=ordersData.length-1;i>=1;i--)if(String(ordersData[i][0])===jobId)ordersSheet.deleteRow(i+1);
  const pay=ss.getSheetByName('PaymentHistory');if(pay){const d=pay.getDataRange().getValues();for(let i=d.length-1;i>=1;i--)if(String(d[i][1])===jobId)pay.deleteRow(i+1);}
  const files=ss.getSheetByName('Files');if(files){const d=files.getDataRange().getValues();for(let i=d.length-1;i>=1;i--)if(String(d[i][1])===jobId){try{DriveApp.getFileById(String(d[i][4])).setTrashed(true);}catch(ignore){}files.deleteRow(i+1);}}
  if(!found)return {ok:false,error:'Job not found'};
  writeAuditLog(String(params.userId||''),'deleteJob',jobId,'Deleted job and related records');return {ok:true,message:'Job dipadam: '+jobId};
}

// UPDATE ORDER STATUS (ownership-checked)
function apiUpdateOrderStatusUnlocked_(params) {
  const jobId = String(params.jobId || '');
  const bil = parseInt(params.bil) || 0;
  const newStatus = String(params.status || '');
  const userId = String(params.userId || '');
  const userRole = String(params.userRole || '');
  const isAdmin = userRole === 'admin';
  
  if (!jobId || !bil) return { ok: false, error: 'jobId dan bil diperlukan' };
  const allowedItemStatuses=['pending','production','ready','delivered'];
  if(allowedItemStatuses.indexOf(newStatus)===-1)return {ok:false,error:'Invalid item status'};
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  // Ownership check: non-admin can only update own job's orders
  if (!isAdmin) {
    const jobsData = jobsSheet.getDataRange().getValues();
    let isOwner = false;
    for (let i = 1; i < jobsData.length; i++) {
      if (String(jobsData[i][0]) === jobId && String(jobsData[i][10]||'') === userId) { isOwner = true; break; }
    }
    if (!isOwner) return { ok: false, error: 'Access denied: not your job' };
  }
  const ordersSheet = ss.getSheetByName('Orders');
  const data = ordersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId && parseInt(data[i][1]) === bil) {
      ordersSheet.getRange(i+1, 11).setValue(newStatus); // Status Item column
      writeAuditLog(userId, 'updateOrderStatus', jobId, 'Item '+bil+' → '+newStatus);
      return { ok: true, message: 'Status dikemaskini' };
    }
  }
  return { ok: false, error: 'Order tidak dijumpai' };
}

// UPDATE JOB STATUS (workflow: submitted → reviewing → quoted → confirmed → production → ready → delivered → completed)
function apiUpdateJobStatusUnlocked_(params) {
  const jobId = String(params.jobId || '');
  const newStatus = String(params.status || '');
  const reason = String(params.reason || '');
  
  if (!jobId || !newStatus) return { ok: false, error: 'jobId and status required' };
  const allowedJobStatuses=['submitted','reviewing','quoted','revised-quote','counter-offer','confirmed','production','ready','delivered','completed','rejected'];
  if(allowedJobStatuses.indexOf(newStatus)===-1)return {ok:false,error:'Invalid job status'};
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const usersSheet = ss.getSheetByName('Users');
  const data = jobsSheet.getDataRange().getValues();
  const now = new Date().toLocaleString('en-GB');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId) {
      const currentStatus=String(data[i][3]||'submitted').toLowerCase();
      const transitions={submitted:['reviewing','rejected'],reviewing:['quoted','rejected'],quoted:['revised-quote','confirmed','rejected'], 'revised-quote':['revised-quote','confirmed','rejected'], 'counter-offer':['revised-quote','confirmed','rejected'],confirmed:['production','reviewing','rejected'],production:['ready','confirmed'],ready:['delivered','production'],delivered:['completed','ready'],completed:['reviewing'],rejected:['reviewing']};
      if(newStatus!==currentStatus && (!transitions[currentStatus] || transitions[currentStatus].indexOf(newStatus)===-1)) return {ok:false,error:'Invalid status transition: '+currentStatus+' → '+newStatus};
      jobsSheet.getRange(i+1, 4).setValue(newStatus); // Status column
      jobsSheet.getRange(i+1, 10).setValue(now); // Dikemaskini Pada column
      if (reason) jobsSheet.getRange(i+1, 17).setValue('[' + newStatus + '] ' + reason); // Notes column
      
      // Send email notification to job creator
      try {
        const creatorId = String(data[i][10] || ''); // diciptaOleh column
        const jobName = String(data[i][1] || '');
        if (usersSheet && creatorId) {
          const usersData = usersSheet.getDataRange().getValues();
          for (let u = 1; u < usersData.length; u++) {
            if (String(usersData[u][0]) === creatorId) {
              const userEmail = String(usersData[u][6] || ''); // email column
              if (userEmail && userEmail.indexOf('@') > -1 && !/\.local$/i.test(userEmail)) {
                const businessConfig = getSettingJSON_('business_config', {emailNotifications:true,portalUrl:'',statusEmailSubject:'HA.GRAPHIXX - Job {jobId} - {status}'});
                if (businessConfig.emailNotifications === false) break;
                const statusLabels = {
                  submitted:'Submitted', reviewing:'Under Review', quoted:'Quoted - Price Set',
                  'revised-quote':'Quote Revised - Please Review', 'counter-offer':'Counter Offer Received',
                  confirmed:'Confirmed', production:'In Production', ready:'Ready for Pickup',
                  delivered:'Delivered', completed:'Completed', rejected:'Rejected'
                };
                const statusLabel = statusLabels[newStatus] || newStatus;
                const subjectTemplate = String(businessConfig.statusEmailSubject || 'HA.GRAPHIXX - Job {jobId} - {status}');
                const subject = subjectTemplate.replace(/\{jobId\}/g, jobId).replace(/\{status\}/g, statusLabel).slice(0,200);
                const portalUrl = /^https?:\/\//i.test(String(businessConfig.portalUrl||'')) ? String(businessConfig.portalUrl) : '';
                const body = 'Hi ' + (usersData[u][3] || 'Customer') + ',\n\n' +
                  'Your job status has been updated:\n\n' +
                  'Job ID: ' + jobId + '\n' +
                  'Job Name: ' + jobName + '\n' +
                  'New Status: ' + statusLabel + '\n' +
                  (reason ? 'Note: ' + reason + '\n' : '') +
                  '\nPlease log in to your portal to view details.\n' +
                  (portalUrl ? ('\n' + portalUrl + '\n') : '') +
                  '\nThank you,\nHA.GRAPHIXX Team';
                MailApp.sendEmail(userEmail, subject, body);
              }
              break;
            }
          }
        }
      } catch(emailErr) {
        // Email sending failed - don't block the status update
      }
      
      writeAuditLog(String(params.userId||''), 'updateJobStatus', jobId, newStatus + (reason?' ('+reason+')':''));
      return { ok: true, message: 'Job status updated to: ' + newStatus };
    }
  }
  return { ok: false, error: 'Job not found' };
}

// SAVE QUOTE (admin set harga jualan untuk setiap item dalam job)
function apiSaveQuoteUnlocked_(params) {
  const jobId=String(params.jobId||'');let quotes=params.quotes||[];
  if(typeof quotes==='string'){try{quotes=JSON.parse(quotes);}catch(e){quotes=[];}}
  if(!jobId)return {ok:false,error:'jobId required'};
  if(!Array.isArray(quotes)||!quotes.length)return {ok:false,error:'No quotes provided'};
  for(let q=0;q<quotes.length;q++){
    const bil=parseInt(quotes[q].bil,10),price=parseFloat(quotes[q].price);
    if(!bil||!isFinite(price)||price<0)return {ok:false,error:'Invalid quote at item '+(q+1)};
  }
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID),jobsSheet=ss.getSheetByName('Jobs'),ordersSheet=ss.getSheetByName('Orders');
  const jobsData=jobsSheet.getDataRange().getValues(),now=new Date().toLocaleString('en-GB');let jobRow=-1;
  for(let i=1;i<jobsData.length;i++)if(String(jobsData[i][0])===jobId){jobRow=i+1;const cur=String(jobsData[i][3]||'').toLowerCase();if(['submitted','reviewing','quoted','revised-quote','counter-offer',''].indexOf(cur)===-1)return {ok:false,error:'Cannot revise quote after job status '+cur};if(cur==='submitted'||cur==='reviewing'||cur==='')jobsSheet.getRange(i+1,4).setValue('quoted');jobsSheet.getRange(i+1,10).setValue(now);break;}
  if(jobRow<0)return {ok:false,error:'Job not found'};
  const ordersData=ordersSheet.getDataRange().getValues();let subtotal=0,matched=0;
  for(let i=1;i<ordersData.length;i++)if(String(ordersData[i][0])===jobId){const bil=parseInt(ordersData[i][1],10),quote=quotes.find(q=>parseInt(q.bil,10)===bil);if(quote){const price=parseFloat(quote.price);ordersSheet.getRange(i+1,8).setValue(price);subtotal+=price;matched++;}else subtotal+=parseFloat(ordersData[i][7])||0;}
  if(!matched)return {ok:false,error:'No matching order items found for quote'};
  const jd=jobsData[jobRow-1],discount=Math.max(0,parseFloat(jd[24])||0),taxRate=Math.max(0,parseFloat(jd[25])||0),shipping=Math.max(0,parseFloat(jd[26])||0),taxable=Math.max(0,subtotal-discount),grandTotal=taxable+(taxable*taxRate/100)+shipping;
  const paid=(parseFloat(jd[17])||0)+(parseFloat(jd[18])||0);if(paid>grandTotal+0.01)return {ok:false,error:'Quoted total cannot be lower than amount already paid'};
  const balance=Math.max(0,grandTotal-paid),payStatus=balance<=0?'paid':(paid>0?'partial':'unpaid');
  jobsSheet.getRange(jobRow,6).setValue(grandTotal);jobsSheet.getRange(jobRow,20).setValue(balance);jobsSheet.getRange(jobRow,21).setValue(payStatus);
  writeAuditLog(String(params.userId||''),'saveQuote',jobId,'Subtotal RM'+subtotal.toFixed(2)+' Grand RM'+grandTotal.toFixed(2));
  return {ok:true,message:'Quotation saved',subtotal,totalQuoted:grandTotal,balance,payStatus};
}

// UPDATE PAYMENT (admin set deposit/paid amount, auto-calc balance + status)
function apiUpdatePayment(params) {
  return withScriptLock_(function(){
    const jobId=String(params.jobId||'');
    const deposit=Math.max(0,parseFloat(params.deposit)||0);
    const paidAmount=Math.max(0,parseFloat(params.paidAmount)||0);
    const userId=String(params.userId||'');
    const method=limitText_(params.method||'bank_transfer',40);
    const reference=limitText_(params.reference||'',120);
    const note=limitText_(params.note||'',300);
    if(!jobId) return {ok:false,error:'jobId required'};
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), jobsSheet=ss.getSheetByName('Jobs');
    const data=jobsSheet.getDataRange().getValues();
    for(let i=1;i<data.length;i++) if(String(data[i][0])===jobId){
      const totalAmount=parseFloat(data[i][5])||0;
      const oldDeposit=parseFloat(data[i][17])||0, oldPaid=parseFloat(data[i][18])||0;
      if(deposit+paidAmount > totalAmount+0.01) return {ok:false,error:'Total payment cannot exceed grand total'};
      const balance=Math.max(0,totalAmount-deposit-paidAmount);
      const payStatus=balance<=0?'paid':(deposit+paidAmount>0?'partial':'unpaid');
      jobsSheet.getRange(i+1,18,1,4).setValues([[deposit,paidAmount,balance,payStatus]]);
      const paySheet=ensureSheet_(ss,'PaymentHistory',['Date','Job ID','Amount','Type','Method','Reference','Note','Recorded By','Transaction ID']);
      const rows=[];
      const depDelta=deposit-oldDeposit, paidDelta=paidAmount-oldPaid;
      if(Math.abs(depDelta)>0.0001) rows.push([new Date(),jobId,Math.abs(depDelta),depDelta>0?'Deposit':'Deposit Adjustment',method,reference,note,userId,generateSecureToken_('pay')]);
      if(Math.abs(paidDelta)>0.0001) rows.push([new Date(),jobId,Math.abs(paidDelta),paidDelta>0?'Payment':'Payment Adjustment',method,reference,note,userId,generateSecureToken_('pay')]);
      if(rows.length) paySheet.getRange(paySheet.getLastRow()+1,1,rows.length,9).setValues(rows);
      writeAuditLog(userId,'updatePayment',jobId,'Deposit:'+deposit+' Paid:'+paidAmount+' Balance:'+balance+' Status:'+payStatus+' Method:'+method+' Ref:'+reference);
      return {ok:true,deposit,paidAmount,balance,payStatus,transactionsAdded:rows.length};
    }
    return {ok:false,error:'Job not found'};
  });
}

// GET PAYMENT HISTORY for a job
function apiGetPaymentHistory(params) {
  const jobId=String(params.jobId||''), userId=String(params.userId||''), userRole=String(params.userRole||'');
  if(!jobId) return {ok:false,error:'jobId required'};
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID);
  if(userRole!=='admin'){
    const d=ss.getSheetByName('Jobs').getDataRange().getValues(); let owner=false;
    for(let i=1;i<d.length;i++) if(String(d[i][0])===jobId && String(d[i][10]||'')===userId){owner=true;break;}
    if(!owner) return {ok:false,error:'Access denied'};
  }
  const sh=ss.getSheetByName('PaymentHistory'); if(!sh) return {ok:true,history:[]};
  const d=sh.getDataRange().getValues(), history=[];
  for(let i=d.length-1;i>=1;i--) if(String(d[i][1])===jobId) history.push({date:formatDateTimeSafe_(d[i][0]),amount:parseFloat(d[i][2])||0,type:String(d[i][3]||''),method:String(d[i][4]||''),reference:String(d[i][5]||''),note:String(d[i][6]||''),recordedBy:String(d[i][7]||d[i][6]||''),transactionId:String(d[i][8]||'')});
  return {ok:true,history};
}

// GET SETTINGS
function apiGetSettings(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return { ok: false, error: 'Sheet Settings tidak wujud' };
  const userRole = String(params.userRole || '');
  const isAdmin = userRole === 'admin';
  
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  // SECURITY: Non-admin cannot access cost_config (modal/supplier cost)
  if (!isAdmin) {
    delete settings.cost_config;
  }
  return { ok: true, settings };
}

// SAVE SETTINGS
function apiSaveSettings(params) {
  return withScriptLock_(function(){
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), sheet=ss.getSheetByName('Settings');
    const nowStr=new Date().toLocaleString('en-GB');
    const pairs=[
      ['pricing_config',params.pricingConfig],['cost_config',params.costConfig],['currency_rates',params.currencyRates],
      ['notification_templates',params.notificationTemplates],['business_config',params.businessConfig]
    ];
    pairs.forEach(function(pair){ if(pair[1]!==undefined && pair[1]!==null && pair[1]!==''){ const v=typeof pair[1]==='string'?pair[1]:JSON.stringify(pair[1]); updateSettingRow(sheet,pair[0],v,nowStr); }});
    updateSettingRow(sheet,'app_version',APP_VERSION,nowStr);
    writeAuditLog(String(params.userId||''),'saveSettings','',pairs.filter(p=>p[1]!==undefined).map(p=>p[0]).join(','));
    return {ok:true,message:'Settings disimpan',appVersion:APP_VERSION};
  });
}

function updateSettingRow(sheet, key, value, timestamp) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i+1, 2, 1, 2).setValues([[value, timestamp]]);
      return;
    }
  }
  sheet.appendRow([key, value, timestamp]);
}

// ==========================================
// PORTAL API (customer-facing)
// Map ERP data to portal format
// ==========================================

// Map job status to production stage (0-5)
function statusToStage(status) {
  const map = {
    'submitted': 0, 'reviewing': 0,
    'quoted': 1, 'confirmed': 1, 'revised-quote': 1, 'counter-offer': 1,
    'production': 2, 'ready': 4,
    'delivered': 5, 'completed': 5,
    'rejected': 0, 'pending': 0, 'Aktif': 2
  };
  return map[status] !== undefined ? map[status] : 0;
}

// Map job status to portal status label
function statusToLabel(status) {
  const map = {
    'submitted': 'Arranging Artwork', 'reviewing': 'Arranging Artwork',
    'quoted': 'Quote Ready - Please Review', 'revised-quote': 'Quote Revised - Please Review',
    'counter-offer': 'Counter Offer Sent', 'confirmed': 'Awaiting Approval',
    'production': 'In Printing', 'ready': 'Ready for Pickup',
    'delivered': 'Completed', 'completed': 'Completed',
    'rejected': 'On Hold', 'pending': 'Arranging Artwork', 'Aktif': 'In Printing'
  };
  return map[status] || 'Arranging Artwork';
}

// PORTAL LOGIN - login with email + password
function apiPortalLogin(params) {
  const email = String(params.email || '').toLowerCase().trim();
  const password = String(params.password || '');
  
  if (!email || !password) return { ok: false, error: 'Email and password required' };
  if (isLoginLocked_('portal:'+email)) return {ok:false,error:'Too many failed attempts. Try again in 15 minutes.'};
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { ok: false, error: 'Users sheet not found' };
  
  // Columns: User ID(0), Password(1), Role(2), Name(3), Agent Code(4), Phone(5), Email(6), Address(7), Active(8)
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const userId = String(row[0]).toLowerCase();
    const userEmail = String(row[6] || '').toLowerCase();
    const userPw = String(row[1]);
    const role = String(row[2]);
    const active = String(row[8]).toUpperCase() === 'TRUE';
    
    // Match by User ID or Email column
    if ((userId === email || userEmail === email) && pwMatches(userPw, password) && active) {
      if (role === 'customer' || role === 'agen') {
        const profile = {
          id: String(row[0]),
          name: String(row[3] || ''),
          email: userEmail || String(row[0]),
          phone: String(row[5] || ''),
          address: String(row[7] || ''),
          postcode: '',
          city: '',
          state: ''
        };
        // Use real session token (same system as ERP login)
        if (String(userPw).indexOf('v2$') !== 0) sheet.getRange(i+1,2).setValue(hashPW(password));
        clearLoginFailures_('portal:'+email);
        const token = createSession(String(row[0]), role, String(row[3]||''));
        return { ok:true, token, profile, mustChangePassword:String(row[11]||'').toUpperCase()==='TRUE' };
      }
    }
  }
  recordLoginFailure_('portal:'+email);
  return { ok: false, error: 'Wrong email or password' };
}

function getPortalProfile(sheet, rowIdx, row) {
  // Columns: User ID(0), Password(1), Role(2), Name(3), Agent Code(4), Phone(5), Email(6), Address(7), Active(8)
  return {
    id: String(row[0]),
    name: String(row[3] || ''),
    email: String(row[6] || '') || String(row[0]),
    phone: String(row[5] || ''),
    address: String(row[7] || ''),
    postcode: '',
    city: '',
    state: ''
  };
}

// PORTAL BOOTSTRAP - load all data for logged-in customer
function apiPortalBootstrap(params) {
  // userId is injected by session token verification
  const userId = String(params.userId || '');
  if (!userId) return { ok: false, error: 'Token required' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const usersSheet = ss.getSheetByName('Users');
  
  // Get profile
  let profile = null;
  if (usersSheet) {
    const usersData = usersSheet.getDataRange().getValues();
    for (let i = 1; i < usersData.length; i++) {
      if (String(usersData[i][0]) === userId) {
        profile = getPortalProfile(usersSheet, i, usersData[i]);
        break;
      }
    }
  }
  
  if (!profile) return { ok: false, error: 'User not found' };
  
  // Get jobs for this user (diciptaOleh = userId)
  const jobsData = jobsSheet ? jobsSheet.getDataRange().getValues() : [];
  const orders = [];
  const invoices = [];
  const requests = [];
  
  for (let i = 1; i < jobsData.length; i++) {
    const row = jobsData[i];
    const jobUserId = String(row[10] || ''); // diciptaOleh column
    if (jobUserId !== userId) continue;
    
    const jobId = String(row[0]);
    const namaJob = String(row[1]);
    const noInvois = String(row[2] || '');
    const status = String(row[3] || 'submitted');
    const helai = parseInt(row[4]) || 0;
    const jumlah = parseFloat(row[5]) || 0;
    const dicipta = row[8] ? new Date(row[8]).toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' }) : '';
    
    const stage = statusToStage(status);
    const statusLabel = statusToLabel(status);
    
    // Determine due date (14 days from created as default)
    let dueDate = dicipta;
    if (row[8]) {
      const due = new Date(row[8]);
      due.setDate(due.getDate() + 14);
      dueDate = due.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });
    }
    
    // Map to portal order format
    orders.push({
      id: jobId,
      service: namaJob,
      qty: helai,
      stage: stage,
      status: statusLabel,
      due: dueDate,
      invoice: noInvois,
      amount: jumlah
    });
    
    // Map to invoice (if invois exists)
    if (noInvois) {
      const balance = (status === 'completed' || status === 'delivered') ? 0 : jumlah;
      invoices.push({
        id: noInvois,
        order: jobId,
        date: dicipta,
        amount: jumlah,
        balance: balance,
        status: balance === 0 ? 'Paid' : 'Sent'
      });
    }
    
    // Map to request (if submitted/reviewing)
    if (status === 'submitted' || status === 'reviewing' || status === 'quoted') {
      requests.push({
        id: 'REQ-' + jobId.slice(-8),
        service: namaJob,
        status: status === 'quoted' ? 'Approved' : 'Pending',
        date: dicipta
      });
    }
  }
  
  // If no data, return empty arrays (not seed data)
  return { ok: true, profile, orders, invoices, requests };
}

// PORTAL SUBMIT REQUEST - customer submit new quote request
function apiPortalSubmitRequestUnlocked_(params) {
  const userId = String(params.userId || '');
  const service = String(params.service || 'Custom Jersey');
  const qty = parseInt(params.qty) || 0;
  const notes = String(params.notes || '');
  
  // Create a job with status 'submitted'
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  
  const now = new Date();
  const ts = Utilities.formatDate(now, 'Asia/Kuala_Lumpur', 'yyyyMMdd-HHmmss');
  const rand = Math.floor(Math.random() * 16777215).toString(16).toUpperCase().padStart(6, '0');
  const jobId = 'SUB-' + ts + '-' + rand;
  const reqId = 'REQ-' + ts.slice(-8);
  
  jobsSheet.appendRow([
    jobId,
    service,
    'REQ-' + now.getFullYear() + String(now.getMonth()+1).padStart(2,'0') + String(now.getDate()).padStart(2,'0'),
    'submitted',
    qty,
    0, // amount TBD - admin will quote
    '', // Konfigurasi Harga JSON
    '', // Cost JSON
    now, // Dicipta Pada
    now.toLocaleString('en-GB'), // Dikemaskini Pada
    userId, // Dicipta Oleh
    'customer', // Role Pencipta
    '', // Agent Code
    '', // Customer Name
    '', // Customer Contact
    '', // Brand JSON
    ''  // Notes
  ]);
  
  return {
    ok: true,
    request: {
      id: reqId,
      service: service,
      qty: qty,
      status: 'Pending',
      date: now.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
    }
  };
}

// PORTAL UPDATE PROFILE
function apiPortalUpdateProfileUnlocked_(params) {
  const userId = String(params.userId || '');
  const profile = params.profile || {};
  if (typeof profile === 'string') { try { profile = JSON.parse(profile); } catch(e) { profile = {}; } }
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { ok: false, error: 'Users sheet not found' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      // Update name (col 4 = index 3), phone (col 6 = index 5), address (col 8 = index 7)
      if (profile.name)    sheet.getRange(i+1, 4).setValue(profile.name);
      if (profile.phone)   sheet.getRange(i+1, 6).setValue(profile.phone);
      if (profile.address) sheet.getRange(i+1, 8).setValue(profile.address);
      return { ok: true };
    }
  }
  return { ok: false, error: 'User not found' };
}

// PORTAL CHANGE PASSWORD
function apiPortalChangePasswordUnlocked_(params) {
  const userId = String(params.userId || '');
  const current = String(params.current || '');
  const next = String(params.next || '');
  
  if(next.length<10 || !/[A-Z]/.test(next) || !/[a-z]/.test(next) || !/[0-9]/.test(next)) return {ok:false,error:'Password must be 10+ characters with uppercase, lowercase and number'};
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { ok: false, error: 'Users sheet not found' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      if (!pwMatches(String(data[i][1]), current)) return { ok: false, error: 'Current password is incorrect' };
      sheet.getRange(i+1, 2).setValue(hashPW(next));
      if (sheet.getLastColumn() >= 12) sheet.getRange(i+1, 12).setValue('FALSE');
      sheet.getRange(i+1, 11).setValue(new Date());
      invalidateAllUserSessions_(userId,String(params.token||''));
      writeAuditLog(userId,'portalChangePassword','', 'Password changed');
      return { ok:true, message:'Password changed successfully' };
    }
  }
  return { ok: false, error: 'User not found' };
}

// ==========================================
// AUDIT LOG
// ==========================================
function writeAuditLog(userId, action, jobId, detail) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let auditSheet = ss.getSheetByName('AuditLog');
  if (!auditSheet) {
    auditSheet = ss.insertSheet('AuditLog');
    auditSheet.appendRow(['Timestamp', 'User ID', 'Action', 'Job ID', 'Detail']);
  }
  auditSheet.appendRow([new Date().toLocaleString('en-GB'), userId, action, jobId, String(detail || '')]);
}

function apiGetAuditLog(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const auditSheet = ss.getSheetByName('AuditLog');
  if (!auditSheet) return { ok: true, logs: [] };
  const data = auditSheet.getDataRange().getValues();
  const logs = [];
  for (let i = data.length - 1; i >= 1 && logs.length < 50; i--) {
    logs.push({
      timestamp: data[i][0],
      userId: data[i][1],
      action: data[i][2],
      jobId: data[i][3],
      detail: data[i][4]
    });
  }
  return { ok: true, logs };
}

// ==========================================
// QUOTE REPLY (customer/agen accept/reject/counter — NOT admin-only)
// ==========================================
function apiQuoteReplyUnlocked_(params) {
  const jobId = String(params.jobId || '');
  const replyType = String(params.replyType || ''); // accept, reject, counter
  const counterAmount = parseFloat(params.counterAmount) || 0;
  const reason = String(params.reason || '');
  const userId = String(params.userId || '');
  const userRole = String(params.userRole || '');
  const isAdmin = userRole === 'admin';
  
  if (!jobId || !replyType) return { ok: false, error: 'jobId and replyType required' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const data = jobsSheet.getDataRange().getValues();
  const now = new Date().toLocaleString('en-GB');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId) {
      // Ownership check: non-admin can only reply to own jobs
      if (!isAdmin && String(data[i][10] || '') !== userId) {
        return { ok: false, error: 'Access denied: not your job' };
      }
      // Quote reply ONLY valid bila job sudah di-quote (quoted / revised-quote).
      // Elak customer/agen accept/reject/counter job yang belum di-quote atau
      // yang sudah di-confirm/production (transaksi telah berkunci).
      const curStatus = String(data[i][3] || '').toLowerCase();
      if (['quoted', 'revised-quote'].indexOf(curStatus) === -1) {
        return { ok: false, error: 'Quote reply hanya dibenarkan bila job berstatus quoted atau revised-quote (status semasa: ' + curStatus + ')' };
      }
      let newStatus = '';
      let detail = '';
      if (replyType === 'accept') {
        newStatus = 'confirmed';
        detail = 'Quote accepted by ' + userId;
      } else if (replyType === 'reject') {
        newStatus = 'rejected';
        detail = 'Quote rejected by ' + userId + (reason ? ': ' + reason : '');
      } else if (replyType === 'counter') {
        if(counterAmount<=0)return {ok:false,error:'Counter amount must be greater than zero'};
        newStatus = 'counter-offer';
        detail = 'Counter offer RM' + counterAmount.toFixed(2) + ' by ' + userId;
      } else {
        return { ok: false, error: 'Invalid replyType' };
      }
      jobsSheet.getRange(i+1, 4).setValue(newStatus);
      jobsSheet.getRange(i+1, 10).setValue(now);
      if (reason || counterAmount > 0) jobsSheet.getRange(i+1, 17).setValue('[' + newStatus + '] ' + detail);
      writeAuditLog(userId, 'quoteReply', jobId, detail);
      return { ok: true, message: 'Quote reply: ' + newStatus, newStatus };
    }
  }
  return { ok: false, error: 'Job not found' };
}

// ==========================================
// RELIABILITY / SECURITY / MAINTENANCE v1.1
// ==========================================
function apiDeleteJob(params){ return withScriptLock_(function(){ return apiDeleteJobUnlocked_(params); }); }
function apiUpdateOrderStatus(params){ return withScriptLock_(function(){ return apiUpdateOrderStatusUnlocked_(params); }); }
function apiUpdateJobStatus(params){ return withScriptLock_(function(){ return apiUpdateJobStatusUnlocked_(params); }); }
function apiSaveQuote(params){ return withScriptLock_(function(){ return apiSaveQuoteUnlocked_(params); }); }
function apiQuoteReply(params){ return withScriptLock_(function(){ return apiQuoteReplyUnlocked_(params); }); }
function apiPortalSubmitRequest(params){ return withScriptLock_(function(){ return apiPortalSubmitRequestUnlocked_(params); }); }
function apiPortalUpdateProfile(params){ return withScriptLock_(function(){ return apiPortalUpdateProfileUnlocked_(params); }); }
function apiPortalChangePassword(params){ return withScriptLock_(function(){ return apiPortalChangePasswordUnlocked_(params); }); }

function apiChangePassword(params){
  return withScriptLock_(function(){
    const userId=String(params.userId||''), current=String(params.current||''), next=String(params.next||'');
    if(next.length<10 || !/[A-Z]/.test(next) || !/[a-z]/.test(next) || !/[0-9]/.test(next)) return {ok:false,error:'Password mesti minimum 10 aksara serta ada huruf besar, huruf kecil dan nombor'};
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=ss.getSheetByName('Users'), d=sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++) if(String(d[i][0])===userId){
      if(!pwMatches(String(d[i][1]),current)) return {ok:false,error:'Current password is incorrect'};
      sh.getRange(i+1,2).setValue(hashPW(next)); sh.getRange(i+1,11).setValue(new Date()); sh.getRange(i+1,12).setValue('FALSE');
      invalidateAllUserSessions_(userId,String(params.token||''));
      writeAuditLog(userId,'changePassword','', 'Password changed');
      return {ok:true,message:'Password changed successfully'};
    }
    return {ok:false,error:'User not found'};
  });
}

function apiUploadFile(params){
  return withScriptLock_(function(){
    const jobId=String(params.jobId||''), fileName=limitText_(params.fileName||'file',150), fileType=limitText_(params.fileType||'artwork',40), userId=String(params.userId||''), userRole=String(params.userRole||'');
    let mimeType=String(params.mimeType||'application/octet-stream');
    if(mimeType==='application/octet-stream'){const ext=(fileName.split('.').pop()||'').toLowerCase(),map={png:'image/png',jpg:'image/jpeg',jpeg:'image/jpeg',webp:'image/webp',pdf:'application/pdf',csv:'text/csv',xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',xls:'application/vnd.ms-excel'};mimeType=map[ext]||mimeType;}
    if(!jobId) return {ok:false,error:'Save/load a job before uploading files'};
    if(!canAccessJob_(jobId,userId,userRole)) return {ok:false,error:'Access denied'};
    const base64=String(params.base64||'').replace(/^data:[^;]+;base64,/,'');
    if(!base64) return {ok:false,error:'No file data received'};
    const bytes=Utilities.base64Decode(base64);
    if(bytes.length>MAX_UPLOAD_BYTES) return {ok:false,error:'Maximum file size is 5MB'};
    const allowed=['image/png','image/jpeg','image/webp','application/pdf','text/csv','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/vnd.ms-excel'];
    if(allowed.indexOf(mimeType)===-1) return {ok:false,error:'Unsupported file type'};
    const folder=getOrCreateSubfolder_(getOrCreateFolder_('HA.GRAPHIXX ERP Files'),jobId);
    const blob=Utilities.newBlob(bytes,mimeType,fileName);
    const file=folder.createFile(blob);
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=ensureSheet_(ss,'Files',['Uploaded At','Job ID','File Type','Original Name','Drive File ID','Drive URL','MIME Type','Size Bytes','Uploaded By']);
    sh.appendRow([new Date(),jobId,fileType,fileName,file.getId(),file.getUrl(),mimeType,bytes.length,userId]);
    writeAuditLog(userId,'uploadFile',jobId,fileType+': '+fileName);
    return {ok:true,file:{name:fileName,type:fileType,url:file.getUrl(),size:bytes.length,uploadedAt:formatDateTimeSafe_(new Date())}};
  });
}

function apiGetFiles(params){
  const jobId=String(params.jobId||''), userId=String(params.userId||''), userRole=String(params.userRole||'');
  if(!jobId || !canAccessJob_(jobId,userId,userRole)) return {ok:false,error:'Access denied'};
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Files'); if(!sh) return {ok:true,files:[]};
  const d=sh.getDataRange().getValues(), files=[];
  for(let i=d.length-1;i>=1;i--) if(String(d[i][1])===jobId) files.push({uploadedAt:formatDateTimeSafe_(d[i][0]),type:d[i][2],name:d[i][3],fileId:d[i][4],mimeType:d[i][6],size:parseInt(d[i][7],10)||0,uploadedBy:d[i][8]});
  return {ok:true,files:files};
}

function apiDownloadFile(params){
  const jobId=String(params.jobId||''),fileId=String(params.fileId||''),userId=String(params.userId||''),userRole=String(params.userRole||'');
  if(!jobId||!fileId||!canAccessJob_(jobId,userId,userRole))return {ok:false,error:'Access denied'};
  const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Files');if(!sh)return {ok:false,error:'File record not found'};
  const d=sh.getDataRange().getValues();let record=null;
  for(let i=1;i<d.length;i++)if(String(d[i][1])===jobId&&String(d[i][4])===fileId){record=d[i];break;}
  if(!record)return {ok:false,error:'File record not found'};
  try{const file=DriveApp.getFileById(fileId),blob=file.getBlob(),bytes=blob.getBytes();if(bytes.length>MAX_UPLOAD_BYTES)return {ok:false,error:'File exceeds secure download limit'};return {ok:true,fileName:String(record[3]||file.getName()),mimeType:String(record[6]||blob.getContentType()||'application/octet-stream'),base64:Utilities.base64Encode(bytes)};}catch(e){return {ok:false,error:'File is unavailable or was deleted'};}
}

function apiCreateBackup(params){
  const ssFile=DriveApp.getFileById(SPREADSHEET_ID), folder=getOrCreateFolder_('HA.GRAPHIXX ERP Backups');
  const stamp=Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','yyyyMMdd-HHmmss');
  const copy=ssFile.makeCopy('HA_GRAPHIXX_ERP_BACKUP_'+stamp,folder);
  pruneOldBackups_(folder,30);
  writeAuditLog(String(params.userId||''),'createBackup','',copy.getName());
  return {ok:true,name:copy.getName(),url:copy.getUrl()};
}
function scheduledDailyBackup(){ apiCreateBackup({userId:'SYSTEM'}); }
function apiInstallDailyBackup(params){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='scheduledDailyBackup') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('scheduledDailyBackup').timeBased().everyDays(1).atHour(2).create();
  writeAuditLog(String(params.userId||''),'installDailyBackup','','Daily 2AM trigger installed');
  return {ok:true,message:'Daily backup trigger installed (around 2:00 AM project timezone)'};
}

function apiHealthCheck(params){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID), checks=[];
  REQUIRED_SHEETS.forEach(function(name){ const sh=ss.getSheetByName(name); checks.push({name:'Sheet '+name,ok:!!sh,detail:sh?(sh.getLastRow()+' rows'):'Missing'}); });
  const jobs=ss.getSheetByName('Jobs'), users=ss.getSheetByName('Users');
  checks.push({name:'Jobs columns',ok:jobs&&jobs.getLastColumn()>=27,detail:jobs?(jobs.getLastColumn()+'/27 columns'):'No Jobs sheet'});
  checks.push({name:'Users columns',ok:users&&users.getLastColumn()>=12,detail:users?(users.getLastColumn()+'/12 columns'):'No Users sheet'});
  checks.push({name:'Spreadsheet access',ok:true,detail:ss.getName()});
  checks.push({name:'App version',ok:true,detail:APP_VERSION});
  const failed=checks.filter(function(c){return !c.ok;}).length;
  return {ok:failed===0,healthy:failed===0,failed:failed,checks:checks,appVersion:APP_VERSION,timestamp:formatDateTimeSafe_(new Date())};
}

function apiExportData(params){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID), output={version:APP_VERSION,exportedAt:new Date().toISOString(),sheets:{}};
  ['Users','Jobs','Orders','Settings','PaymentHistory','Files','AuditLog'].forEach(function(name){ const sh=ss.getSheetByName(name); if(sh){ const rows=sh.getDataRange().getDisplayValues(); if(name==='Users') for(let i=1;i<rows.length;i++) rows[i][1]='[REDACTED]'; output.sheets[name]=rows; } });
  writeAuditLog(String(params.userId||''),'exportData','','Full JSON export');
  return {ok:true,data:output};
}




function apiExportReport(params){
  const type=String(params.reportType||'sales').toLowerCase();
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID),jobsSh=ss.getSheetByName('Jobs'),ordersSh=ss.getSheetByName('Orders');
  const jobs=jobsSh.getDataRange().getValues(),orders=ordersSh.getDataRange().getValues();
  const settings=getSettingJSON_('cost_config',{rn_pendek:18,rn_panjang:24,rn_muslimah:27,rn_kid:0,polo:23,retro:25,surcharge_start:'3XL',surcharge_step:2});
  const ordersByJob={};for(let i=1;i<orders.length;i++){const id=String(orders[i][0]);(ordersByJob[id]||(ordersByJob[id]=[])).push(orders[i]);}
  let headers=[],rows=[];
  if(type==='sales'){
    headers=['Job ID','Job Name','Invoice','Status','Pieces','Grand Total RM','Created','Owner'];
    for(let i=1;i<jobs.length;i++)rows.push([jobs[i][0],jobs[i][1],jobs[i][2],jobs[i][3],jobs[i][4],jobs[i][5],formatDateTimeSafe_(jobs[i][8]),jobs[i][10]]);
  }else if(type==='unpaid'){
    headers=['Job ID','Job Name','Invoice','Grand Total RM','Paid RM','Balance RM','Payment Status','Owner'];
    for(let i=1;i<jobs.length;i++){const bal=parseFloat(jobs[i][19])||0;if(bal>0)rows.push([jobs[i][0],jobs[i][1],jobs[i][2],jobs[i][5],(parseFloat(jobs[i][17])||0)+(parseFloat(jobs[i][18])||0),bal,jobs[i][20]||'unpaid',jobs[i][10]]);}
  }else if(type==='production'){
    headers=['Job ID','Job Name','Status','Pieces','Updated','Owner'];const allowed=['confirmed','production','ready'];
    for(let i=1;i<jobs.length;i++)if(allowed.indexOf(String(jobs[i][3]).toLowerCase())>=0)rows.push([jobs[i][0],jobs[i][1],jobs[i][3],jobs[i][4],formatDateTimeSafe_(jobs[i][9]),jobs[i][10]]);
  }else if(type==='agent'){
    headers=['Agent/User','Jobs','Pieces','Sales RM'];const agg={};
    for(let i=1;i<jobs.length;i++){const owner=String(jobs[i][10]||'unknown'),a=agg[owner]||(agg[owner]={jobs:0,pieces:0,sales:0});a.jobs++;a.pieces+=parseInt(jobs[i][4],10)||0;a.sales+=parseFloat(jobs[i][5])||0;}
    Object.keys(agg).sort().forEach(k=>rows.push([k,agg[k].jobs,agg[k].pieces,agg[k].sales]));
  }else if(type==='monthly'){
    headers=['Month','Jobs','Pieces','Sales RM'];const agg={};
    for(let i=1;i<jobs.length;i++){let d;try{d=new Date(jobs[i][8]);}catch(e){continue;}if(isNaN(d.getTime()))continue;const key=Utilities.formatDate(d,'Asia/Kuala_Lumpur','yyyy-MM'),a=agg[key]||(agg[key]={jobs:0,pieces:0,sales:0});a.jobs++;a.pieces+=parseInt(jobs[i][4],10)||0;a.sales+=parseFloat(jobs[i][5])||0;}
    Object.keys(agg).sort().forEach(k=>rows.push([k,agg[k].jobs,agg[k].pieces,agg[k].sales]));
  }else if(type==='profit'){
    headers=['Job ID','Job Name','Pieces','Revenue RM','Estimated Cost RM','Estimated Profit RM','Margin %'];
    for(let i=1;i<jobs.length;i++){const id=String(jobs[i][0]),revenue=parseFloat(jobs[i][5])||0;let cost=0;(ordersByJob[id]||[]).forEach(function(o){cost+=estimateOrderCost_(o,settings);});const profit=revenue-cost,margin=revenue>0?profit/revenue*100:0;rows.push([id,jobs[i][1],jobs[i][4],revenue,cost,profit,margin.toFixed(2)]);}
  }else return {ok:false,error:'Unknown report type'};
  writeAuditLog(String(params.userId||''),'exportReport','',type+' '+rows.length+' rows');
  return {ok:true,reportType:type,headers:headers,rows:rows,fileName:'ha_graphixx_'+type+'_'+Utilities.formatDate(new Date(),'Asia/Kuala_Lumpur','yyyyMMdd-HHmmss')+'.csv'};
}
function estimateOrderCost_(orderRow,cfg){
  const sleeve=String(orderRow[2]||'').toLowerCase(),size=String(orderRow[3]||'').toUpperCase(),collar=String(orderRow[5]||'').toLowerCase();let key='rn_pendek';
  if(collar.indexOf('polo')>=0)key='polo';else if(collar.indexOf('retro')>=0)key='retro';else if(sleeve.indexOf('panjang')>=0)key='rn_panjang';else if(sleeve.indexOf('muslimah')>=0)key='rn_muslimah';else if(sleeve.indexOf('kid')>=0)key='rn_kid';
  let cost=parseFloat(cfg[key])||0;const sizes=['XS','S','M','L','XL','2XL','3XL','4XL','5XL','6XL','7XL','8XL','9XL','10XL'],start=sizes.indexOf(String(cfg.surcharge_start||'3XL').toUpperCase()),idx=sizes.indexOf(size);if(start>=0&&idx>=start)cost+=(idx-start+1)*(parseFloat(cfg.surcharge_step)||0);return cost;
}
function apiUpdateInvoiceAdjustments(params){
  return withScriptLock_(function(){
    const jobId=String(params.jobId||''),discount=Math.max(0,parseFloat(params.discount)||0),taxRate=Math.max(0,parseFloat(params.taxRate)||0),shipping=Math.max(0,parseFloat(params.shipping)||0);
    if(!jobId)return {ok:false,error:'jobId required'};
    if(taxRate>100)return {ok:false,error:'Tax rate cannot exceed 100%'};
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),jobs=ss.getSheetByName('Jobs'),orders=ss.getSheetByName('Orders'),jd=jobs.getDataRange().getValues(),od=orders.getDataRange().getValues();
    let subtotal=0;for(let i=1;i<od.length;i++)if(String(od[i][0])===jobId)subtotal+=parseFloat(od[i][7])||0;
    if(discount>subtotal)return {ok:false,error:'Discount cannot exceed subtotal'};
    const taxable=Math.max(0,subtotal-discount),taxAmount=taxable*taxRate/100,grandTotal=taxable+taxAmount+shipping;
    for(let i=1;i<jd.length;i++)if(String(jd[i][0])===jobId){const dep=parseFloat(jd[i][17])||0,paid=parseFloat(jd[i][18])||0;if(dep+paid>grandTotal+0.01)return {ok:false,error:'Grand total cannot be lower than amount already paid'};const balance=Math.max(0,grandTotal-dep-paid),payStatus=balance<=0?'paid':(dep+paid>0?'partial':'unpaid');jobs.getRange(i+1,6).setValue(grandTotal);jobs.getRange(i+1,20).setValue(balance);jobs.getRange(i+1,21).setValue(payStatus);jobs.getRange(i+1,25,1,3).setValues([[discount,taxRate,shipping]]);writeAuditLog(String(params.userId||''),'updateInvoiceAdjustments',jobId,'Discount:'+discount+' Tax:'+taxRate+' Shipping:'+shipping+' Grand:'+grandTotal);return {ok:true,subtotal,discount,taxRate,taxAmount,shipping,grandTotal,balance,payStatus};}
    return {ok:false,error:'Job not found'};
  });
}
function apiGetJobOptions(params){
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName('Jobs'),d=sh.getDataRange().getValues();
  const userId=String(params.userId||''),isAdmin=String(params.userRole||'')==='admin',options=[];
  for(let i=d.length-1;i>=1&&options.length<1000;i--){if(!isAdmin&&String(d[i][10]||'')!==userId)continue;options.push({jobId:d[i][0],namaJob:d[i][1],noInvois:d[i][2],status:d[i][3]});}
  return {ok:true,jobs:options};
}
function apiAdminResetPassword(params){
  return withScriptLock_(function(){
    const target=String(params.targetUserId||'').trim(),newPassword=String(params.newPassword||'');
    if(!target)return {ok:false,error:'Target user ID required'};
    if(newPassword.length<10||!/[A-Z]/.test(newPassword)||!/[a-z]/.test(newPassword)||!/[0-9]/.test(newPassword))return {ok:false,error:'Temporary password must be 10+ chars with uppercase, lowercase and number'};
    const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ss.getSheetByName('Users'),d=sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++)if(String(d[i][0]).toLowerCase()===target.toLowerCase()){
      sh.getRange(i+1,2).setValue(hashPW(newPassword));sh.getRange(i+1,11).setValue(new Date());sh.getRange(i+1,12).setValue('TRUE');invalidateAllUserSessions_(String(d[i][0]),'');
      writeAuditLog(String(params.userId||''),'adminResetPassword','',target);return {ok:true,message:'Password reset. User must change it at next login.'};
    }
    return {ok:false,error:'User not found'};
  });
}

function getSettingJSON_(key,fallback){try{const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Settings'),d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++)if(String(d[i][0])===String(key))return typeof d[i][1]==='object'?d[i][1]:JSON.parse(String(d[i][1]||'{}'));}catch(e){}return fallback;}
function isLoginLocked_(key){const c=CacheService.getScriptCache(),n=parseInt(c.get('loginfail:'+String(key).toLowerCase()),10)||0;return n>=5;}
function recordLoginFailure_(key){const c=CacheService.getScriptCache(),k='loginfail:'+String(key).toLowerCase(),n=(parseInt(c.get(k),10)||0)+1;c.put(k,String(n),900);}
function clearLoginFailures_(key){CacheService.getScriptCache().remove('loginfail:'+String(key).toLowerCase());}

function withScriptLock_(fn){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(30000)) throw new Error('System busy. Please try again in a few seconds.');
  try{return fn();}finally{lock.releaseLock();}
}
function ensureSheet_(ss,name,headers){ let sh=ss.getSheetByName(name); if(!sh){sh=ss.insertSheet(name);sh.appendRow(headers);}else ensureColumns_(sh,headers);return sh; }
function ensureColumns_(sheet,headers){
  while(sheet.getMaxColumns()<headers.length) sheet.insertColumnAfter(sheet.getMaxColumns());
  if(sheet.getLastRow()===0) sheet.getRange(1,1,1,headers.length).setValues([headers]);
  else { const current=sheet.getRange(1,1,1,headers.length).getValues()[0]; for(let i=0;i<headers.length;i++) if(!current[i]) sheet.getRange(1,i+1).setValue(headers[i]); }
}
function sha256Hex_(value){ const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(value)); return bytes.map(function(b){return (b<0?b+256:b).toString(16).padStart(2,'0');}).join(''); }
function generateSecureToken_(prefix){ return String(prefix||'tok')+'-'+sha256Hex_(Date.now()+'|'+Math.random()+'|'+Utilities.getUuid()).slice(0,40); }
function timingSafeEqual_(a,b){ a=String(a);b=String(b); if(a.length!==b.length)return false; let x=0;for(let i=0;i<a.length;i++)x|=a.charCodeAt(i)^b.charCodeAt(i);return x===0; }
function generateJobId_(prefix){ const now=new Date(),ts=Utilities.formatDate(now,'Asia/Kuala_Lumpur','yyyyMMdd-HHmmss');return String(prefix||'JOB')+'-'+ts+'-'+generateSecureToken_('id').slice(-8).toUpperCase(); }
function limitText_(v,max){ return String(v===undefined||v===null?'':v).trim().slice(0,max||500); }
function formatDateSafe_(v){ try{return v?Utilities.formatDate(new Date(v),'Asia/Kuala_Lumpur','dd/MM/yyyy'):'';}catch(e){return String(v||'');} }
function formatDateTimeSafe_(v){ try{return v?Utilities.formatDate(new Date(v),'Asia/Kuala_Lumpur','dd/MM/yyyy HH:mm:ss'):'';}catch(e){return String(v||'');} }
function cleanError_(err){ const msg=String(err&&err.message?err.message:err); return msg.replace(/Exception:\s*/g,'').slice(0,500); }
function logError_(action,userId,err){ try{ const ss=SpreadsheetApp.openById(SPREADSHEET_ID),sh=ensureSheet_(ss,'ErrorLog',['Timestamp','Action','User ID','Message','Stack']); sh.appendRow([new Date(),action,userId,cleanError_(err),String(err&&err.stack||'').slice(0,2000)]); }catch(ignore){} }
function getOrCreateFolder_(name){ const it=DriveApp.getFoldersByName(name); return it.hasNext()?it.next():DriveApp.createFolder(name); }
function getOrCreateSubfolder_(parent,name){const it=parent.getFoldersByName(name);return it.hasNext()?it.next():parent.createFolder(name);}
function pruneOldBackups_(folder,keep){const files=[],it=folder.getFiles();while(it.hasNext())files.push(it.next());files.sort(function(a,b){return b.getDateCreated()-a.getDateCreated();});for(let i=keep;i<files.length;i++)files[i].setTrashed(true);}
function getPermissions_(role){
  const map={admin:{viewAllJobs:true,editOwnJobs:true,editAnyJob:true,managePricing:true,viewCost:true,manageStatus:true,managePayment:true,backup:true,upload:true},agen:{viewAllJobs:false,editOwnJobs:true,editAnyJob:false,managePricing:false,viewCost:false,manageStatus:false,managePayment:false,backup:false,upload:true},customer:{viewAllJobs:false,editOwnJobs:true,editAnyJob:false,managePricing:false,viewCost:false,manageStatus:false,managePayment:false,backup:false,upload:true}};
  return map[String(role||'customer')]||map.customer;
}
function canAccessJob_(jobId,userId,userRole){ if(userRole==='admin')return true; const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Jobs'),d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++)if(String(d[i][0])===String(jobId))return String(d[i][10]||'')===String(userId);return false; }
function invalidateAllUserSessions_(userId,exceptToken){ const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Sessions');if(!sh)return;const d=sh.getDataRange().getValues();for(let i=d.length-1;i>=1;i--)if(String(d[i][1])===userId&&String(d[i][0])!==exceptToken)sh.deleteRow(i+1); }
