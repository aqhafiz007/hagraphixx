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

// ==========================================
// SETUP - Cipta sheet automatik
// ==========================================
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // Sheet: Users (GID 1895047476) - dah ada data
  let usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    usersSheet = ss.insertSheet('Users');
    usersSheet.appendRow(['User ID', 'Password', 'Role', 'Name', 'Agent Code', 'Phone', 'Email', 'Address', 'Active', 'Created At', 'Updated At']);
  }
  // Pastikan ada user baru jika belum ada
  const usersData = usersSheet.getDataRange().getValues();
  const existingIds = usersData.map(r => String(r[0]));
  const newUsers = [
    ['admin', hashPW('admin123'), 'admin', 'Administrator', '', '', 'admin@hagraphixx.local', '', 'TRUE', new Date(), new Date()],
    ['agen001', hashPW('agen123'), 'agen', 'Agen Ali', 'AG001', '012-345 6789', 'agen001@hagraphixx.local', '', 'TRUE', new Date(), new Date()],
    ['cus001', hashPW('cus123'), 'customer', 'Siti Aishah', '', '011-1115 0199', 'cus001@hagraphixx.local', 'Asrama SMKA Al-Mashoor (L), Jalan Air Itam', 'TRUE', new Date(), new Date()],
    ['agen002', hashPW('agen123'), 'agen', 'Agen Bakar', 'AG002', '', 'agen002@hagraphixx.local', '', 'TRUE', new Date(), new Date()],
    ['agen003', hashPW('agen123'), 'agen', 'Agen Siti Nur', 'AG003', '', 'agen003@hagraphixx.local', '', 'TRUE', new Date(), new Date()],
    ['cus002', hashPW('cus123'), 'customer', 'Tan Wei Jie', '', '', 'cus002@hagraphixx.local', '', 'TRUE', new Date(), new Date()],
    ['cus003', hashPW('cus123'), 'customer', 'Nurul Ain', '', '', 'cus003@hagraphixx.local', '', 'TRUE', new Date(), new Date()],
  ];
  newUsers.forEach(u => {
    if (!existingIds.includes(u[0])) usersSheet.appendRow(u);
  });
  
  // Sheet: Jobs - dah ada data
  if (!ss.getSheetByName('Jobs')) {
    const jobsSheet = ss.insertSheet('Jobs');
    jobsSheet.appendRow(['Job ID', 'Nama Job', 'No. Invois', 'Status', 'Jumlah Helai', 'Jumlah (RM)', 'Konfigurasi Harga JSON', 'Cost JSON', 'Dicipta Pada', 'Dikemaskini Pada', 'Dicipta Oleh', 'Role Pencipta', 'Agent Code', 'Customer Name', 'Customer Contact', 'Brand JSON', 'Notes']);
  }
  
  // Sheet: Orders - dah ada data
  if (!ss.getSheetByName('Orders')) {
    const ordersSheet = ss.insertSheet('Orders');
    ordersSheet.appendRow(['Job ID', 'Bil', 'Lengan', 'Saiz', 'Nama', 'Kolar', 'No/Remarks', 'Harga Unit (RM)', 'Base Price (RM)', 'Tarikh Disimpan', 'Status Item']);
  }
  
  // Sheet: Settings - dah ada data
  if (!ss.getSheetByName('Settings')) {
    const settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Kunci', 'Nilai', 'Dikemaskini Pada']);
  }
  
  return 'Setup selesai! Sheets: Users, Jobs, Orders, Settings - semua dah ada data.';
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
      default:             result = { ok: false, error: 'Unknown action: ' + action };
    }
  } catch(err) {
    result = { ok: false, error: String(err) };
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

// ==========================================
// SECURITY - password hashing (SHA-256)
// New passwords are stored as hashes. Comparison also accepts legacy
// plain-text values so existing sheets keep working.
// ==========================================
function hashPW(pw) {
  try {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw));
    return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
  } catch (e) { return String(pw); }
}
function pwMatches(stored, input) {
  if (!stored) return false;
  if (stored === hashPW(input)) return true;
  if (stored === String(input)) return true; // legacy plain-text
  return false;
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
      return {
        ok: true,
        user: {
          id: row[0],
          role: userRole,
          name: row[3],
          agentCode: row[4],
          phone: row[5] || '',
          email: row[6] || '',
          address: row[7] || ''
        }
      };
    }
  }
  return { ok: false, error: 'Wrong ID, password or role' };
}

// DASHBOARD - ringkasan statistik
function apiGetDashboard(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const ordersSheet = ss.getSheetByName('Orders');
  
  const jobsData = jobsSheet ? jobsSheet.getDataRange().getValues() : [];
  const ordersData = ordersSheet ? ordersSheet.getDataRange().getValues() : [];
  
  let totalJobs = 0, totalHelai = 0, totalRevenue = 0, activeJobs = 0;
  const statusCounts = {};
  const closed = ['completed', 'delivered', 'cancelled', 'rejected'];
  
  for (let i = 1; i < jobsData.length; i++) {
    totalJobs++;
    totalHelai += parseInt(jobsData[i][4]) || 0;
    totalRevenue += parseFloat(jobsData[i][5]) || 0;
    const status = String(jobsData[i][3] || 'submitted').toLowerCase();
    if (closed.indexOf(status) === -1) activeJobs++;
    statusCounts[status] = (statusCounts[status] || 0) + 1;
  }
  
  const recentJobs = [];
  for (let i = Math.max(1, jobsData.length - 5); i < jobsData.length; i++) {
    if (i < 1) continue;
    recentJobs.push({
      jobId: jobsData[i][0],
      namaJob: jobsData[i][1],
      noInvois: jobsData[i][2],
      status: jobsData[i][3],
      helai: jobsData[i][4],
      jumlah: jobsData[i][5],
      dicipta: jobsData[i][8] ? new Date(jobsData[i][8]).toLocaleDateString('en-GB') : ''
    });
  }
  recentJobs.reverse();
  
  return {
    ok: true,
    stats: {
      totalJobs,
      totalHelai,
      totalRevenue,
      activeJobs,
      statusCounts,
      totalOrders: ordersData.length - 1
    },
    recentJobs
  };
}

// GET JOBS - senarai semua job
function apiGetJobs(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Jobs');
  if (!sheet) return { ok: false, error: 'Sheet Jobs tidak wujud' };
  
  const data = sheet.getDataRange().getValues();
  const jobs = [];
  for (let i = 1; i < data.length; i++) {
    jobs.push({
      jobId: data[i][0],
      namaJob: data[i][1],
      noInvois: data[i][2],
      status: data[i][3],
      helai: data[i][4],
      jumlah: data[i][5],
      config: data[i][6],
      dicipta: data[i][8] ? new Date(data[i][8]).toLocaleString('en-GB') : '',
      dikemaskini: data[i][9] ? new Date(data[i][9]).toLocaleString('en-GB') : '',
      diciptaOleh: data[i][10] || '',
      rolePencipta: data[i][11] || ''
    });
  }
  return { ok: true, jobs };
}

// GET JOB - dapatkan satu job + orders
function apiGetJob(params) {
  const jobId = String(params.jobId || '');
  if (!jobId) return { ok: false, error: 'jobId diperlukan' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const ordersSheet = ss.getSheetByName('Orders');
  
  // Cari job
  const jobsData = jobsSheet.getDataRange().getValues();
  let jobInfo = null;
  for (let i = 1; i < jobsData.length; i++) {
    if (jobsData[i][0] === jobId) {
      jobInfo = {
        jobId: jobsData[i][0],
        namaJob: jobsData[i][1],
        noInvois: jobsData[i][2],
        status: jobsData[i][3],
        helai: jobsData[i][4],
        jumlah: jobsData[i][5],
        config: jobsData[i][6],
        dicipta: jobsData[i][8] ? new Date(jobsData[i][8]).toLocaleString('en-GB') : '',
        dikemaskini: jobsData[i][9] ? new Date(jobsData[i][9]).toLocaleString('en-GB') : '',
        diciptaOleh: jobsData[i][10] || '',
        rolePencipta: jobsData[i][11] || ''
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const ordersSheet = ss.getSheetByName('Orders');
  
  const jobData = typeof params.job === 'string' ? JSON.parse(params.job) : params.job;
  const ordersData = typeof params.orders === 'string' ? JSON.parse(params.orders) : (params.orders || []);
  const userId = String(params.userId || 'unknown');
  const userRole = String(params.userRole || '');
  
  // Jana Job ID jika baru
  let jobId = jobData.jobId || '';
  const isNew = !jobId;
  
  if (isNew) {
    const now = new Date();
    const ts = Utilities.formatDate(now, 'Asia/Kuala_Lumpur', 'yyyyMMdd-HHmmss');
    const rand = Math.floor(Math.random() * 16777215).toString(16).toUpperCase().padStart(6, '0');
    jobId = 'JOB-' + ts + '-' + rand;
  }
  
  // Kira jumlah
  let totalHelai = ordersData.length;
  let totalRM = 0;
  ordersData.forEach(o => totalRM += parseFloat(o.hargaUnit) || 0);
  
  const config = jobData.config ? (typeof jobData.config === 'string' ? jobData.config : JSON.stringify(jobData.config)) : '';
  const nowStr = new Date().toLocaleString('en-GB');
  
  if (isNew) {
    // Tambah row baru ke Jobs
    jobsSheet.appendRow([
      jobId,
      jobData.namaJob || 'Untitled',
      jobData.noInvois || ('INV-' + new Date().getFullYear() + String(new Date().getMonth()+1).padStart(2,'0') + String(new Date().getDate()).padStart(2,'0')),
      jobData.status || 'Aktif',
      totalHelai,
      totalRM,
      config,
      '', // Cost JSON (column 8)
      new Date(),
      nowStr,
      userId,
      userRole,
      '', // Agent Code
      '', // Customer Name
      '', // Customer Contact
      '', // Brand JSON
      ''  // Notes
    ]);
    
    // Tambah orders - Columns: Job ID, Bil, Lengan, Saiz, Nama, Kolar, No/Remarks, Harga Unit, Base Price, Tarikh, Status
    let bil = 1;
    ordersData.forEach(o => {
      ordersSheet.appendRow([
        jobId,
        bil++,
        o.lengan || 'Pendek',
        o.saiz || 'M',
        (o.nama || '').toUpperCase(),
        o.kolar || 'Round neck (RN)',
        o.remarks || '-',
        parseFloat(o.hargaUnit) || 0,
        parseFloat(o.basePrice) || parseFloat(o.hargaUnit) || 0,
        nowStr,
        o.status || 'pending'
      ]);
    });
  } else {
    // Kemaskini job sedia ada
    const jobsData = jobsSheet.getDataRange().getValues();
    for (let i = 1; i < jobsData.length; i++) {
      if (jobsData[i][0] === jobId) {
        jobsSheet.getRange(i+1, 2, 1, 9).setValues([[
          jobData.namaJob || jobsData[i][1],
          jobData.noInvois || jobsData[i][2],
          jobData.status || jobsData[i][3],
          totalHelai,
          totalRM,
          config,
          jobsData[i][7],   // Cost JSON - keep
          jobsData[i][8],   // Dicipta Pada - keep
          nowStr            // Dikemaskini Pada
        ]]);
        break;
      }
    }
    
    // Padam orders lama untuk job ini, kemudian tambah baru
    const ordersAll = ordersSheet.getDataRange().getValues();
    const rowsToDelete = [];
    for (let i = ordersAll.length - 1; i >= 1; i--) {
      if (ordersAll[i][0] === jobId) rowsToDelete.push(i+1);
    }
    rowsToDelete.forEach(r => ordersSheet.deleteRow(r));
    
    // Tambah orders baru
    let bil = 1;
    ordersData.forEach(o => {
      ordersSheet.appendRow([
        jobId,
        bil++,
        o.lengan || 'Pendek',
        o.saiz || 'M',
        (o.nama || '').toUpperCase(),
        o.kolar || 'Round neck (RN)',
        o.remarks || '-',
        parseFloat(o.hargaUnit) || 0,
        parseFloat(o.basePrice) || parseFloat(o.hargaUnit) || 0,
        nowStr,
        o.status || 'pending'
      ]);
    });
  }
  
  return { ok: true, jobId, totalHelai, totalRM };
}

// DELETE JOB
function apiDeleteJob(params) {
  const jobId = String(params.jobId || '');
  if (!jobId) return { ok: false, error: 'jobId diperlukan' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const ordersSheet = ss.getSheetByName('Orders');
  
  // Padam dari Jobs
  const jobsData = jobsSheet.getDataRange().getValues();
  for (let i = jobsData.length - 1; i >= 1; i--) {
    if (jobsData[i][0] === jobId) jobsSheet.deleteRow(i+1);
  }
  
  // Padam dari Orders
  const ordersData = ordersSheet.getDataRange().getValues();
  for (let i = ordersData.length - 1; i >= 1; i--) {
    if (ordersData[i][0] === jobId) ordersSheet.deleteRow(i+1);
  }
  
  return { ok: true, message: 'Job dipadam: ' + jobId };
}

// UPDATE ORDER STATUS
function apiUpdateOrderStatus(params) {
  const jobId = String(params.jobId || '');
  const bil = parseInt(params.bil) || 0;
  const newStatus = String(params.status || '');
  
  if (!jobId || !bil) return { ok: false, error: 'jobId dan bil diperlukan' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const ordersSheet = ss.getSheetByName('Orders');
  const data = ordersSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId && parseInt(data[i][1]) === bil) {
      ordersSheet.getRange(i+1, 11).setValue(newStatus); // Status Item column
      return { ok: true, message: 'Status dikemaskini' };
    }
  }
  return { ok: false, error: 'Order tidak dijumpai' };
}

// UPDATE JOB STATUS (workflow: submitted → reviewing → quoted → confirmed → production → ready → delivered → completed)
function apiUpdateJobStatus(params) {
  const jobId = String(params.jobId || '');
  const newStatus = String(params.status || '');
  const reason = String(params.reason || '');
  
  if (!jobId || !newStatus) return { ok: false, error: 'jobId and status required' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const usersSheet = ss.getSheetByName('Users');
  const data = jobsSheet.getDataRange().getValues();
  const now = new Date().toLocaleString('en-GB');
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === jobId) {
      jobsSheet.getRange(i+1, 4).setValue(newStatus); // Status column
      jobsSheet.getRange(i+1, 10).setValue(now); // Dikemaskini Pada column
      if (reason) jobsSheet.getRange(i+1, 4).setValue(newStatus + ' (rejected: ' + reason + ')');
      
      // Send email notification to job creator
      try {
        const creatorId = String(data[i][10] || ''); // diciptaOleh column
        const jobName = String(data[i][1] || '');
        if (usersSheet && creatorId) {
          const usersData = usersSheet.getDataRange().getValues();
          for (let u = 1; u < usersData.length; u++) {
            if (String(usersData[u][0]) === creatorId) {
              const userEmail = String(usersData[u][6] || ''); // email column
              if (userEmail && userEmail.indexOf('@') > -1) {
                const statusLabels = {
                  submitted:'Submitted', reviewing:'Under Review', quoted:'Quoted - Price Set',
                  confirmed:'Confirmed', production:'In Production', ready:'Ready for Pickup',
                  delivered:'Delivered', completed:'Completed', rejected:'Rejected'
                };
                const statusLabel = statusLabels[newStatus] || newStatus;
                const subject = 'HA.GRAPHIXX - Job ' + jobId + ' - ' + statusLabel;
                const body = 'Hi ' + (usersData[u][3] || 'Customer') + ',\n\n' +
                  'Your job status has been updated:\n\n' +
                  'Job ID: ' + jobId + '\n' +
                  'Job Name: ' + jobName + '\n' +
                  'New Status: ' + statusLabel + '\n' +
                  (reason ? 'Note: ' + reason + '\n' : '') +
                  '\nPlease log in to your portal to view details.\n\n' +
                  'https://ha.graphixx\n\n' +
                  'Thank you,\nHA.GRAPHIXX Team';
                MailApp.sendEmail(userEmail, subject, body);
              }
              break;
            }
          }
        }
      } catch(emailErr) {
        // Email sending failed - don't block the status update
      }
      
      return { ok: true, message: 'Job status updated to: ' + newStatus };
    }
  }
  return { ok: false, error: 'Job not found' };
}

// SAVE QUOTE (admin set harga jualan untuk setiap item dalam job)
function apiSaveQuote(params) {
  const jobId = String(params.jobId || '');
  const quotes = params.quotes || [];
  
  if (!jobId) return { ok: false, error: 'jobId required' };
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const jobsSheet = ss.getSheetByName('Jobs');
  const ordersSheet = ss.getSheetByName('Orders');
  
  // Update job status to quoted
  const jobsData = jobsSheet.getDataRange().getValues();
  const now = new Date().toLocaleString('en-GB');
  let jobFound = false;
  for (let i = 1; i < jobsData.length; i++) {
    if (jobsData[i][0] === jobId) {
      jobsSheet.getRange(i+1, 4).setValue('quoted');
      jobsSheet.getRange(i+1, 9).setValue(now);
      jobFound = true;
      break;
    }
  }
  if (!jobFound) return { ok: false, error: 'Job not found' };
  
  // Update order prices with admin quotes
  const ordersData = ordersSheet.getDataRange().getValues();
  let totalQuoted = 0;
  for (let i = 1; i < ordersData.length; i++) {
    if (ordersData[i][0] === jobId) {
      const bil = parseInt(ordersData[i][1]);
      const quote = quotes.find(q => parseInt(q.bil) === bil);
      if (quote) {
        ordersSheet.getRange(i+1, 8).setValue(parseFloat(quote.price) || 0);
        totalQuoted += parseFloat(quote.price) || 0;
      } else {
        totalQuoted += parseFloat(ordersData[i][7]) || 0;
      }
    }
  }
  
  // Update job total
  for (let i = 1; i < jobsData.length; i++) {
    if (jobsData[i][0] === jobId) {
      jobsSheet.getRange(i+1, 6).setValue(totalQuoted);
      break;
    }
  }
  
  return { ok: true, message: 'Quotation saved', totalQuoted };
}

// GET SETTINGS
function apiGetSettings(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) return { ok: false, error: 'Sheet Settings tidak wujud' };
  
  const data = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) settings[data[i][0]] = data[i][1];
  }
  return { ok: true, settings };
}

// SAVE SETTINGS
function apiSaveSettings(params) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Settings');
  
  const pricingConfig = params.pricingConfig ? (typeof params.pricingConfig === 'string' ? params.pricingConfig : JSON.stringify(params.pricingConfig)) : '';
  const costConfig = params.costConfig ? (typeof params.costConfig === 'string' ? params.costConfig : JSON.stringify(params.costConfig)) : '';
  
  const nowStr = new Date().toLocaleString('en-GB');
  
  if (pricingConfig) updateSettingRow(sheet, 'pricing_config', pricingConfig, nowStr);
  if (costConfig) updateSettingRow(sheet, 'cost_config', costConfig, nowStr);
  
  return { ok: true, message: 'Settings disimpan' };
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
    'quoted': 1, 'confirmed': 1,
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
    'quoted': 'Awaiting Approval', 'confirmed': 'Awaiting Approval',
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
        return { ok: true, token: 'portal-' + row[0] + '-' + Date.now(), profile };
      }
    }
  }
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
  const token = String(params.token || '');
  if (!token) return { ok: false, error: 'Token required' };
  
  // Extract user ID from token: portal-USERID-TIMESTAMP
  const parts = token.split('-');
  const userId = parts.length >= 3 ? parts.slice(1, -1).join('-') : '';
  
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
function apiPortalSubmitRequest(params) {
  const token = String(params.token || '');
  const service = String(params.service || 'Custom Jersey');
  const qty = parseInt(params.qty) || 0;
  const notes = String(params.notes || '');
  
  const parts = token.split('-');
  const userId = parts.length >= 3 ? parts.slice(1, -1).join('-') : '';
  
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
function apiPortalUpdateProfile(params) {
  const token = String(params.token || '');
  const profile = params.profile || {};
  
  const parts = token.split('-');
  const userId = parts.length >= 3 ? parts.slice(1, -1).join('-') : '';
  
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
function apiPortalChangePassword(params) {
  const token = String(params.token || '');
  const current = String(params.current || '');
  const next = String(params.next || '');
  
  if (!next || next.length < 8) return { ok: false, error: 'New password must be at least 8 characters' };
  
  const parts = token.split('-');
  const userId = parts.length >= 3 ? parts.slice(1, -1).join('-') : '';
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return { ok: false, error: 'Users sheet not found' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === userId) {
      if (!pwMatches(String(data[i][1]), current)) return { ok: false, error: 'Current password is incorrect' };
      sheet.getRange(i+1, 2).setValue(hashPW(next)); // Update password column (hashed)
      return { ok: true };
    }
  }
  return { ok: false, error: 'User not found' };
}
