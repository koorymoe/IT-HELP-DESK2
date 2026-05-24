// ═══════════════════════════════════════════════════════════════
//  IT HELP DESK — Code.gs الكامل v10.3 (مع تعيين المراقب من الإيميل)
// ═══════════════════════════════════════════════════════════════

var SS           = SpreadsheetApp.getActiveSpreadsheet();
var SH_USERS     = 'الموظفون';
var SH_TICKETS   = 'البلاغات';
var SH_DEVICES   = 'الأجهزة';
var SH_GUIDELINES= 'التوجيهات';
var SH_MANUAL    = 'كتيب_التعليمات';
var SH_DEPTS     = 'الأقسام';
var SH_NOTIF     = 'السجل';
var SH_IUSERS    = 'InternetUsers';

var COL_USERS = {
  id:          'كود داخلي',
  empId:       'رقم البصمة',
  firstName:   'الاسم الأول',
  lastName:    'الاسم الأخير',
  phone:       'رقم الهاتف',
  dept:        'القسم',
  role:        'الصلاحية',
  pwHash:      'كلمة المرور Hash',
  active:      'نشط',
  createdAt:   'تاريخ الإنشاء',
  lastLogin:   'آخر دخول',
  notifyEmail: 'إيميل الإشعارات'
};

// البحث المرن عن عمود — يدعم الأسماء البديلة وإزالة # من بداية/نهاية النص
var COL_USERS_ALIASES = {
  id:          ['كود داخلي','id','ID','كود'],
  empId:       ['رقم البصمة','#رقم البصمة','رقم البصمة#','empId','رقم_البصمة','الرقم الوظيفي'],
  firstName:   ['الاسم الأول','firstName','الاسم'],
  lastName:    ['الاسم الأخير','lastName','الكنية'],
  phone:       ['رقم الهاتف','رقم الهاتف #','#رقم الهاتف','phone','الهاتف'],
  dept:        ['القسم','dept','department','الإدارة','الدائرة'],
  role:        ['الصلاحية','role','الدور','الوظيفة'],
  pwHash:      ['كلمة المرور Hash','كلمة المرور','pwHash','password','الرمز'],
  active:      ['نشط','active','مفعل','الحالة'],
  createdAt:   ['تاريخ الإنشاء','createdAt','التاريخ','تاريخ'],
  lastLogin:   ['آخر دخول','lastLogin','آخر_دخول'],
  notifyEmail: ['إيميل الإشعارات','البريد الإلكتروني','الإيميل','email','notifyEmail','البريد']
};

function findColIndex(headers, key) {
  var aliases = COL_USERS_ALIASES[key] || [COL_USERS[key]];
  for (var a = 0; a < aliases.length; a++) {
    for (var h = 0; h < headers.length; h++) {
      var clean = String(headers[h]).trim().replace(/^#+|#+$/g, '').trim();
      if (clean === aliases[a] || String(headers[h]).trim() === aliases[a]) return h;
    }
  }
  return -1;
}

var COL_TICKETS = {
  id:0, createdAt:1, requesterName:2, requesterId:3,
  requesterDept:4, deviceId:5, problemType:7, priority:8,
  description:9, status:10, assignedId:11, assignedName:12
};

var EMAIL_SENDER_NAME = 'IT Help Desk';
var SYSTEM_URL        = 'https://script.google.com/macros/s/AKfycbyUJjIi77Q3DMeusfZuj4gv7HXWLgx9Zu4YzCeHfwyvigajWbzNEvGSueTDCqyE3bgK/exec';
var ROLES_IT    = ['it','it_manager','admin'];
var ROLES_MGR   = ['manager','it_manager','admin'];
var ROLES_TECH  = ['tech','it','it_manager','admin'];
var ROLES_ADMIN = ['admin'];
var ROLES_STAFF = ['it','it_manager','manager','admin'];

// ══════════════════════════════════════════════════════
//  doGet — فتح النظام + معالجة روابط الإيميل
// ══════════════════════════════════════════════════════
function doGet(e) {
  var action   = e && e.parameter ? e.parameter.action : '';
  var ticketId = e && e.parameter ? e.parameter.tid    : '';
  var userId   = e && e.parameter ? e.parameter.uid    : '';
  var token    = e && e.parameter ? e.parameter.token  : '';

  if (action === 'claim' && ticketId && userId)
    return handleEmailClaim(ticketId, userId, token);
  if (action === 'solve' && ticketId && userId)
    return handleEmailSolve(ticketId, userId, token);
  if (action === 'assignForm' && ticketId && userId)
    return handleEmailAssignForm(ticketId, userId, token);
  if (action === 'doAssign' && ticketId && userId) {
    var assignTo = e && e.parameter ? e.parameter.assignTo : '';
    return handleEmailDoAssign(ticketId, userId, token, assignTo);
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle('IT Help Desk');
}

function generateEmailToken(ticketId, userId) {
  return sha256(ticketId + userId + TOKEN_SECRET).slice(0, 16);
}

// ── استلام من الإيميل (IT) ──
function handleEmailClaim(ticketId, userId, token) {
  var html = '';
  try {
    if (token !== generateEmailToken(ticketId, userId)) {
      html = buildResultPage('❌ رابط غير صحيح', 'هذا الرابط غير صحيح أو منتهي الصلاحية.', '#ef4444');
      return HtmlService.createHtmlOutput(html);
    }
    var empData = getEmployeeByEmpId(userId);
    if (!empData) { html = buildResultPage('❌ موظف غير موجود', 'رقم البصمة غير موجود في النظام.', '#ef4444'); return HtmlService.createHtmlOutput(html); }
    if (ROLES_IT.indexOf(empData.role) < 0) { html = buildResultPage('❌ غير مصرح', 'ليس لديك صلاحية لاستلام البلاغات.', '#ef4444'); return HtmlService.createHtmlOutput(html); }
    var sh = getSheet(SH_TICKETS), allData = sh.getDataRange().getValues();
    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][COL_TICKETS.id]).trim() === ticketId) {
        var currentAssigned = String(allData[i][COL_TICKETS.assignedName] || '').trim();
        var currentStatus   = String(allData[i][COL_TICKETS.status] || '').trim();
        if (currentAssigned && currentStatus !== 'جديدة') {
          html = buildResultPage('⚠️ تم الاستلام مسبقاً', 'هذا البلاغ تم استلامه مسبقاً من قبل: <strong>' + currentAssigned + '</strong><br>الحالة الحالية: ' + currentStatus, '#f59e0b');
          return HtmlService.createHtmlOutput(html);
        }
        sh.getRange(i+1, COL_TICKETS.assignedId+1).setValue(empData.empId);
        sh.getRange(i+1, COL_TICKETS.assignedName+1).setValue(empData.firstName + ' ' + empData.lastName);
        sh.getRange(i+1, COL_TICKETS.status+1).setValue('معينة');
        notifyTeamAboutClaim(ticketId, empData, allData[i]);
        html = buildResultPage('✅ تم الاستلام بنجاح', 'تم تسجيل استلامك للبلاغ <strong>' + ticketId + '</strong><br>سيصل إشعار لباقي الفريق.', '#059669', ticketId);
        return HtmlService.createHtmlOutput(html);
      }
    }
    html = buildResultPage('❌ البلاغ غير موجود', 'لم يتم العثور على البلاغ رقم ' + ticketId, '#ef4444');
  } catch(e) { html = buildResultPage('❌ خطأ', e.message, '#ef4444'); }
  return HtmlService.createHtmlOutput(html);
}

// ── حل من الإيميل (IT) ──
function handleEmailSolve(ticketId, userId, token) {
  var html = '';
  try {
    if (token !== generateEmailToken(ticketId, userId)) {
      html = buildResultPage('❌ رابط غير صحيح', 'هذا الرابط غير صحيح أو منتهي الصلاحية.', '#ef4444');
      return HtmlService.createHtmlOutput(html);
    }
    var empData = getEmployeeByEmpId(userId);
    if (!empData || ROLES_IT.indexOf(empData.role) < 0) { html = buildResultPage('❌ غير مصرح', 'ليس لديك صلاحية لإغلاق البلاغات.', '#ef4444'); return HtmlService.createHtmlOutput(html); }
    var sh = getSheet(SH_TICKETS), allData = sh.getDataRange().getValues();
    for (var i = 0; i < allData.length; i++) {
      if (String(allData[i][COL_TICKETS.id]).trim() === ticketId) {
        var currentStatus = String(allData[i][COL_TICKETS.status] || '').trim();
        if (currentStatus === 'تم حل البلاغ') { html = buildResultPage('✅ البلاغ محلول مسبقاً', 'هذا البلاغ تم حله مسبقاً.', '#059669'); return HtmlService.createHtmlOutput(html); }
        sh.getRange(i+1, COL_TICKETS.status+1).setValue('تم حل البلاغ');
        if (!String(allData[i][COL_TICKETS.assignedName]||'').trim()) {
          sh.getRange(i+1, COL_TICKETS.assignedId+1).setValue(empData.empId);
          sh.getRange(i+1, COL_TICKETS.assignedName+1).setValue(empData.firstName+' '+empData.lastName);
        }
        var reqEmail = getEmailByEmpId(String(allData[i][COL_TICKETS.requesterId]));
        if (reqEmail) sendSolvedEmail(reqEmail, ticketId, String(allData[i][COL_TICKETS.requesterName]||''), 'تم حل البلاغ من قبل '+empData.firstName+' '+empData.lastName);
        html = buildResultPage('✅ تم تسجيل الحل بنجاح', 'تم إغلاق البلاغ <strong>' + ticketId + '</strong> وإرسال إشعار لصاحب البلاغ.', '#059669', ticketId);
        return HtmlService.createHtmlOutput(html);
      }
    }
    html = buildResultPage('❌ البلاغ غير موجود', 'لم يتم العثور على البلاغ رقم ' + ticketId, '#ef4444');
  } catch(e) { html = buildResultPage('❌ خطأ', e.message, '#ef4444'); }
  return HtmlService.createHtmlOutput(html);
}

// ── تعيين من الإيميل (المراقب) — صفحة الاختيار ──
function handleEmailAssignForm(ticketId, mgrEmpId, token) {
  try {
    if (token !== generateEmailToken(ticketId, mgrEmpId))
      return HtmlService.createHtmlOutput(buildResultPage('❌ رابط غير صحيح', 'هذا الرابط غير صحيح أو منتهي الصلاحية.', '#ef4444'));
    var mgrData = getEmployeeByEmpId(mgrEmpId);
    if (!mgrData || ROLES_MGR.indexOf(mgrData.role) < 0)
      return HtmlService.createHtmlOutput(buildResultPage('❌ غير مصرح', 'ليس لديك صلاحية تعيين البلاغات.', '#ef4444'));
    var tSh = getSheet(SH_TICKETS), tAll = tSh.getDataRange().getValues(), ticketRow = null;
    for (var i = 1; i < tAll.length; i++) { if (String(tAll[i][COL_TICKETS.id]).trim() === ticketId) { ticketRow = tAll[i]; break; } }
    if (!ticketRow) return HtmlService.createHtmlOutput(buildResultPage('❌ البلاغ غير موجود', 'رقم البلاغ: ' + ticketId, '#ef4444'));
    var assignedName = String(ticketRow[COL_TICKETS.assignedName] || '').trim();
    var currentStatus = String(ticketRow[COL_TICKETS.status] || '').trim();
    if (assignedName && currentStatus !== 'جديدة')
      return HtmlService.createHtmlOutput(buildResultPage('⚠️ تم التعيين مسبقاً', 'هذا البلاغ تم تعيينه إلى: <strong>' + assignedName + '</strong>', '#f59e0b'));
    var uSh = getSheet(SH_USERS), uAll = uSh.getDataRange().getValues(), uH = uAll[0];
    var idCol=uH.indexOf(COL_USERS.id), empIdCol=uH.indexOf(COL_USERS.empId);
    var fnCol=uH.indexOf(COL_USERS.firstName), lnCol=uH.indexOf(COL_USERS.lastName);
    var roleCol=uH.indexOf(COL_USERS.role), activeCol=uH.indexOf(COL_USERS.active);
    var itStaff = [];
    for (var j = 1; j < uAll.length; j++) {
      var r = uAll[j], role = String(r[roleCol]||'').trim().toLowerCase(), active = r[activeCol];
      if (ROLES_IT.indexOf(role) >= 0 && (active===true||String(active).toLowerCase()==='true'))
        itStaff.push({id:String(r[idCol]||'').trim(), empId:String(r[empIdCol]||'').trim(), name:(String(r[fnCol]||'')+' '+String(r[lnCol]||'')).trim(), role:role});
    }
    var ticketInfo = {id:String(ticketRow[COL_TICKETS.id]||''), problemType:String(ticketRow[COL_TICKETS.problemType]||''), requesterName:String(ticketRow[COL_TICKETS.requesterName]||''), requesterDept:String(ticketRow[COL_TICKETS.requesterDept]||''), priority:String(ticketRow[COL_TICKETS.priority]||''), status:currentStatus};
    return buildAssignFormPage(ticketInfo, mgrData, itStaff, token);
  } catch(e) { return HtmlService.createHtmlOutput(buildResultPage('❌ خطأ', e.message, '#ef4444')); }
}

// ── تنفيذ التعيين من الإيميل (المراقب) ──
function handleEmailDoAssign(ticketId, mgrEmpId, token, assignToEmpId) {
  try {
    if (token !== generateEmailToken(ticketId, mgrEmpId))
      return HtmlService.createHtmlOutput(buildResultPage('❌ رابط غير صحيح', 'هذا الرابط غير صحيح أو منتهي الصلاحية.', '#ef4444'));
    var mgrUser = getEmployeeByEmpId(mgrEmpId);
    if (!mgrUser || ROLES_MGR.indexOf(mgrUser.role) < 0)
      return HtmlService.createHtmlOutput(buildResultPage('❌ غير مصرح', 'ليس لديك صلاحية لتعيين البلاغات.', '#ef4444'));
    if (!assignToEmpId)
      return HtmlService.createHtmlOutput(buildResultPage('❌ لم يتم الاختيار', 'يرجى اختيار موظف IT.', '#ef4444'));
    var assigneeData = getEmployeeByEmpId(assignToEmpId);
    if (!assigneeData)
      return HtmlService.createHtmlOutput(buildResultPage('❌ الموظف غير موجود', 'رقم البصمة: ' + assignToEmpId, '#ef4444'));
    var fullName = assigneeData.firstName + ' ' + assigneeData.lastName;
    var result = assignTicketDirect(ticketId, assigneeData.empId, fullName, mgrUser);
    if (!result.success)
      return HtmlService.createHtmlOutput(buildResultPage('❌ خطأ', result.message, '#ef4444'));
    try {
      if (assigneeData.email) {
        var tSh = getSheet(SH_TICKETS), tAll = tSh.getDataRange().getValues();
        for (var k = 0; k < tAll.length; k++) { if (String(tAll[k][COL_TICKETS.id]).trim() === ticketId) { notifyAssignee(ticketId, assigneeData, tAll[k]); break; } }
      }
    } catch(e2) { Logger.log('notifyAssignee: '+e2.toString()); }
    try {
      var tSh2 = getSheet(SH_TICKETS), tAll2 = tSh2.getDataRange().getValues();
      for (var m = 0; m < tAll2.length; m++) { if (String(tAll2[m][COL_TICKETS.id]).trim() === ticketId) { notifyTeamAboutAssign(ticketId, mgrUser, assigneeData, tAll2[m]); break; } }
    } catch(e3) { Logger.log('notifyTeamAboutAssign: '+e3.toString()); }
    return HtmlService.createHtmlOutput(buildResultPage('✅ تم التعيين بنجاح', 'تم تعيين البلاغ <strong>' + ticketId + '</strong> إلى <strong>' + fullName + '</strong><br>تم إرسال إشعار للموظف وبقية الفريق.', '#059669', ticketId));
  } catch(e) { return HtmlService.createHtmlOutput(buildResultPage('❌ خطأ', e.message, '#ef4444')); }
}

// ── صفحة اختيار موظف IT (HTML جميل) ──
function buildAssignFormPage(ticketInfo, mgrData, itStaff, token) {
  var ROLE_LABELS = {it:'موظف IT', it_manager:'مدير IT', admin:'مدير النظام'};
  var staffCards = itStaff.length ? itStaff.map(function(s) {
    return '<div class="staff-card" onclick="pick(this,\''+s.empId+'\')">'
      +'<div class="av">'+(s.name?s.name.charAt(0):'؟')+'</div>'
      +'<div class="sn">'+s.name+'</div>'
      +'<div class="sr">'+(ROLE_LABELS[s.role]||'موظف IT')+'</div>'
      +'</div>';
  }).join('') : '<p style="text-align:center;color:#f87171;grid-column:1/-1">لا يوجد موظفو IT نشطون</p>';
  var pc = ticketInfo.priority==='عاجلة'?'#ef4444':ticketInfo.priority==='عالية'?'#f59e0b':ticketInfo.priority==='متوسطة'?'#3b82f6':'#10b981';
  var html = '<!DOCTYPE html><html dir="rtl" lang="ar"><head>'
    +'<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    +'<title>تعيين البلاغ '+ticketInfo.id+'</title>'
    +'<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">'
    +'<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Cairo,Tahoma,Arial,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;padding:24px 16px}'
    +'.wrap{max-width:680px;margin:0 auto}.top{text-align:center;margin-bottom:28px}.top h1{font-size:24px;font-weight:900;color:#f8fafc}'
    +'.top p{font-size:13px;color:#64748b;margin-top:6px}'
    +'.tkt{background:#1e293b;border-radius:14px;padding:20px;margin-bottom:24px;border:1px solid #334155}'
    +'.tkt h2{font-size:14px;font-weight:700;color:#8b5cf6;margin-bottom:14px}'
    +'.tkt-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}'
    +'.tkt-item label{display:block;font-size:11px;color:#64748b;margin-bottom:3px}'
    +'.tkt-item span{font-size:14px;font-weight:600;color:#f1f5f9}'
    +'.pri{display:inline-block;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:700;background:'+pc+'20;color:'+pc+'}'
    +'.sec{font-size:14px;font-weight:700;color:#8b5cf6;margin-bottom:14px}'
    +'.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:12px;margin-bottom:28px}'
    +'.staff-card{background:#1e293b;border:2px solid #334155;border-radius:12px;padding:18px 12px;text-align:center;cursor:pointer;transition:all .18s}'
    +'.staff-card:hover{border-color:#6366f1;transform:translateY(-2px)}'
    +'.staff-card.sel{border-color:#6366f1;background:#1e2040;box-shadow:0 0 0 3px rgba(99,102,241,.25)}'
    +'.av{width:52px;height:52px;border-radius:50%;background:linear-gradient(135deg,#312e81,#4338ca);color:#a5b4fc;font-size:22px;font-weight:900;display:flex;align-items:center;justify-content:center;margin:0 auto 10px}'
    +'.sn{font-size:13px;font-weight:700;color:#f1f5f9;margin-bottom:4px}.sr{font-size:11px;color:#64748b}'
    +'.btn-wrap{text-align:center}.btn{display:inline-block;padding:14px 48px;background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;border:none;border-radius:10px;font-size:16px;font-weight:700;cursor:pointer;font-family:Cairo,Tahoma,Arial,sans-serif;transition:opacity .2s}'
    +'.btn:disabled{opacity:.4;cursor:not-allowed}.msg{color:#f87171;font-size:13px;margin-top:10px;text-align:center}'
    +'@media(max-width:480px){.tkt-grid{grid-template-columns:1fr}}</style></head><body><div class="wrap">'
    +'<div class="top"><h1>🎯 تعيين البلاغ</h1><p>مرحباً '+mgrData.firstName+' — اختر موظف IT لتعيين البلاغ إليه</p></div>'
    +'<div class="tkt"><h2>📋 تفاصيل البلاغ</h2><div class="tkt-grid">'
    +'<div class="tkt-item"><label>رقم البلاغ</label><span>'+ticketInfo.id+'</span></div>'
    +'<div class="tkt-item"><label>نوع المشكلة</label><span>'+ticketInfo.problemType+'</span></div>'
    +'<div class="tkt-item"><label>المُبلِّغ</label><span>'+ticketInfo.requesterName+'</span></div>'
    +'<div class="tkt-item"><label>القسم</label><span>'+ticketInfo.requesterDept+'</span></div>'
    +'<div class="tkt-item"><label>الأولوية</label><span><span class="pri">'+ticketInfo.priority+'</span></span></div>'
    +'<div class="tkt-item"><label>الحالة</label><span>'+ticketInfo.status+'</span></div>'
    +'</div></div>'
    +'<div class="sec">👥 فريق تقنية المعلومات</div>'
    +'<div class="grid">'+staffCards+'</div>'
    +'<div class="btn-wrap"><button class="btn" id="assignBtn" disabled onclick="doAssign()">تعيين البلاغ ✓</button>'
    +'<div class="msg" id="msg"></div></div>'
    +'</div><script>'
    +'var sel="";'
    +'function pick(el,empId){document.querySelectorAll(".staff-card").forEach(function(c){c.classList.remove("sel");});el.classList.add("sel");sel=empId;document.getElementById("assignBtn").disabled=false;document.getElementById("msg").textContent="";}'
    +'function doAssign(){if(!sel){document.getElementById("msg").textContent="يرجى اختيار موظف أولاً";return;}'
    +'var btn=document.getElementById("assignBtn");btn.disabled=true;btn.textContent="جاري التعيين...";'
    +'window.location.href="'+SYSTEM_URL+'?action=doAssign"'
    +'+"&tid="+encodeURIComponent("'+ticketInfo.id+'")'
    +'+"&uid="+encodeURIComponent("'+mgrData.empId+'")'
    +'+"&token="+encodeURIComponent("'+token+'")'
    +'+"&assignTo="+encodeURIComponent(sel);}'
    +'<\/script></body></html>';
  return HtmlService.createHtmlOutput(html).setTitle('تعيين البلاغ '+ticketInfo.id).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── إشعار فريق IT عند تعيين المراقب ──
function notifyTeamAboutAssign(ticketId, assignedBy, assignee, ticketRow) {
  try {
    var uD=sheetData(SH_USERS),uH=uD.headers;
    var problemType=String(ticketRow[COL_TICKETS.problemType]||''),requesterName=String(ticketRow[COL_TICKETS.requesterName]||'');
    uD.rows.forEach(function(r){
      var role=String(r[uH.indexOf('role')]||''),email=String(r[uH.indexOf('notifyEmail')]||'').trim();
      var empId=String(r[uH.indexOf('empId')]||'').trim(),active=r[uH.indexOf('active')];
      if(ROLES_STAFF.indexOf(role)<0)return;
      if(empId===assignee.empId)return;
      if(!email||!(active===true||String(active).toLowerCase()==='true'))return;
      var html='<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body dir="rtl" style="font-family:Arial;padding:20px;background:#f8fafc">'
        +'<div style="max-width:500px;margin:0 auto;background:#fff;border-radius:12px;padding:24px;box-shadow:0 4px 20px rgba(0,0,0,.1)">'
        +'<div style="background:linear-gradient(135deg,#f59e0b,#d97706);padding:16px 20px;border-radius:8px;margin-bottom:16px">'
        +'<h3 style="margin:0;color:#fff;font-size:15px">📌 تم تعيين بلاغ</h3></div>'
        +'<p style="font-size:13px;color:#374151">قام <strong>'+assignedBy.firstName+' '+assignedBy.lastName+'</strong> بتعيين البلاغ <strong>'+ticketId+'</strong> إلى <strong>'+assignee.firstName+' '+assignee.lastName+'</strong></p>'
        +'<div style="background:#f8fafc;border-radius:8px;padding:12px;margin:12px 0;font-size:13px">'
        +'<div><strong>رقم البلاغ:</strong> '+ticketId+'</div>'
        +'<div><strong>نوع المشكلة:</strong> '+problemType+'</div>'
        +'<div><strong>المُبلِّغ:</strong> '+requesterName+'</div></div>'
        +'<a href="'+SYSTEM_URL+'" style="background:#6366f1;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-size:13px;display:inline-block">فتح النظام</a>'
        +'</div></body></html>';
      try{MailApp.sendEmail({to:email,subject:'📌 تم تعيين البلاغ ['+ticketId+'] إلى '+assignee.firstName,htmlBody:html});}catch(e2){}
    });
  }catch(e){Logger.log('notifyTeamAboutAssign: '+e.toString());}
}

// ── إيميل المراقب (زر تعيين) ──
function buildManagerEmail(ticket, assignUrl) {
  var pc=ticket.priority==='عاجلة'?'#ef4444':ticket.priority==='عالية'?'#f59e0b':ticket.priority==='متوسطة'?'#3b82f6':'#10b981';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body dir="rtl" style="margin:0;padding:0;background:#f8fafc;font-family:Arial,sans-serif">'
    +'<div style="max-width:600px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1)">'
    +'<div style="background:linear-gradient(135deg,#f59e0b,#d97706);padding:24px 28px">'
    +'<h2 style="margin:0;color:#fff;font-size:18px">🔔 بلاغ جديد يحتاج تعيين</h2>'
    +'<p style="margin:6px 0 0;color:rgba(255,255,255,.8);font-size:13px">رقم البلاغ: '+ticket.id+'</p>'
    +'</div><div style="padding:24px 28px">'
    +'<table style="width:100%;border-collapse:collapse;margin-bottom:20px">'
    +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px;width:35%">نوع المشكلة</td><td style="padding:9px 14px;font-size:13px">'+ticket.problemType+'</td></tr>'
    +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">المُبلِّغ</td><td style="padding:9px 14px;font-size:13px">'+ticket.requesterName+'</td></tr>'
    +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">القسم</td><td style="padding:9px 14px;font-size:13px">'+ticket.requesterDept+'</td></tr>'
    +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">الأولوية</td><td style="padding:9px 14px"><span style="background:'+pc+'20;color:'+pc+';padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold">'+ticket.priority+'</span></td></tr>'
    +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">الوصف</td><td style="padding:9px 14px;font-size:13px">'+ticket.description+'</td></tr>'
    +'</table>'
    +'<div style="display:flex;gap:10px;flex-wrap:wrap">'
    +'<a href="'+assignUrl+'" style="background:linear-gradient(135deg,#f59e0b,#d97706);color:#fff;padding:13px 28px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block">👥 تعيين البلاغ لموظف IT</a>'
    +'<a href="'+SYSTEM_URL+'" style="background:#f8fafc;color:#374151;padding:13px 22px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block;border:2px solid #e2e8f0">🔍 فتح النظام</a>'
    +'</div>'
    +'<p style="font-size:11px;color:#94a3b8;margin-top:12px">اضغط "تعيين البلاغ" لاختيار موظف IT مباشرة بدون تسجيل دخول</p>'
    +'</div>'
    +'<div style="padding:14px 28px;background:#f8fafc;border-top:1px solid #e2e8f0;text-align:center;font-size:11px;color:#94a3b8">IT Help Desk — نظام الدعم التقني</div>'
    +'</div></body></html>';
}

function getEmployeeByEmpId(empId) {
  var sh=getSheet(SH_USERS),allData=sh.getDataRange().getValues(),rh=allData[0];
  var empCol=findColIndex(rh,'empId'),roleCol=findColIndex(rh,'role');
  var fnCol=findColIndex(rh,'firstName'),lnCol=findColIndex(rh,'lastName');
  var emailCol=findColIndex(rh,'notifyEmail'),idCol=findColIndex(rh,'id');
  for(var i=1;i<allData.length;i++){
    var rEmpId=String(allData[i][empCol]||'').trim();
    if(rEmpId===String(empId).trim()||(parseInt(rEmpId)&&parseInt(rEmpId)===parseInt(empId))){
      return{empId:rEmpId,id:String(allData[i][idCol]||'').trim(),firstName:String(allData[i][fnCol]||'').trim(),lastName:String(allData[i][lnCol]||'').trim(),role:String(allData[i][roleCol]||'user').trim().toLowerCase(),email:emailCol>=0?String(allData[i][emailCol]||'').trim():''};
    }
  }
  return null;
}

function notifyTeamAboutClaim(ticketId, claimant, ticketRow) {
  try {
    var uD=sheetData(SH_USERS),uH=uD.headers;
    var problemType=String(ticketRow[COL_TICKETS.problemType]||''),requesterName=String(ticketRow[COL_TICKETS.requesterName]||'');
    uD.rows.forEach(function(r){
      var role=String(r[uH.indexOf('role')]||''),email=String(r[uH.indexOf('notifyEmail')]||'').trim();
      var empId=String(r[uH.indexOf('empId')]||'').trim(),active=r[uH.indexOf('active')];
      if(ROLES_STAFF.indexOf(role)<0)return;
      if(empId===claimant.empId)return;
      if(!email||!(active===true||String(active).toLowerCase()==='true'))return;
      var html='<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body dir="rtl" style="font-family:Arial;padding:20px;background:#f8fafc">'
        +'<div style="max-width:500px;margin:0 auto;background:#fff;border-radius:12px;padding:24px;box-shadow:0 4px 20px rgba(0,0,0,.1)">'
        +'<div style="background:linear-gradient(135deg,#06b6d4,#3b82f6);padding:16px 20px;border-radius:8px;margin-bottom:16px">'
        +'<h3 style="margin:0;color:#fff;font-size:15px">📋 تم استلام بلاغ</h3></div>'
        +'<p style="font-size:13px;color:#374151">قام <strong>'+claimant.firstName+' '+claimant.lastName+'</strong> باستلام البلاغ:</p>'
        +'<div style="background:#f8fafc;border-radius:8px;padding:12px;margin:12px 0;font-size:13px">'
        +'<div><strong>رقم البلاغ:</strong> '+ticketId+'</div>'
        +'<div><strong>نوع المشكلة:</strong> '+problemType+'</div>'
        +'<div><strong>المُبلِّغ:</strong> '+requesterName+'</div></div>'
        +'<a href="'+SYSTEM_URL+'" style="background:#6366f1;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-size:13px;display:inline-block">فتح النظام</a>'
        +'</div></body></html>';
      try{MailApp.sendEmail({to:email,subject:'📋 تم استلام البلاغ ['+ticketId+'] من قبل '+claimant.firstName+' '+claimant.lastName,htmlBody:html});}catch(e2){}
    });
  }catch(e){Logger.log('notifyTeamAboutClaim: '+e.toString());}
}

function notifyAssignee(ticketId, assignee, ticketRow) {
  if(!assignee.email)return;
  try{
    var problemType=String(ticketRow[COL_TICKETS.problemType]||''),requesterName=String(ticketRow[COL_TICKETS.requesterName]||'');
    var requesterDept=String(ticketRow[COL_TICKETS.requesterDept]||''),description=String(ticketRow[COL_TICKETS.description]||''),priority=String(ticketRow[COL_TICKETS.priority]||'');
    var token=generateEmailToken(ticketId,assignee.empId);
    var claimUrl=SYSTEM_URL+'?action=claim&tid='+encodeURIComponent(ticketId)+'&uid='+encodeURIComponent(assignee.empId)+'&token='+token;
    var html='<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body dir="rtl" style="margin:0;padding:0;background:#f8fafc;font-family:Arial">'
      +'<div style="max-width:600px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1)">'
      +'<div style="background:linear-gradient(135deg,#6366f1,#8b5cf6);padding:24px 28px">'
      +'<h2 style="margin:0;color:#fff;font-size:18px">📋 تم تعيينك لبلاغ جديد</h2>'
      +'<p style="margin:6px 0 0;color:rgba(255,255,255,.7);font-size:13px">'+ticketId+'</p></div>'
      +'<div style="padding:24px 28px">'
      +'<p style="font-size:14px">عزيزي <strong>'+assignee.firstName+' '+assignee.lastName+'</strong>،</p>'
      +'<table style="width:100%;border-collapse:collapse;margin-bottom:20px;font-size:13px">'
      +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;width:35%">نوع المشكلة</td><td style="padding:9px 14px">'+problemType+'</td></tr>'
      +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold">المُبلِّغ</td><td style="padding:9px 14px">'+requesterName+'</td></tr>'
      +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold">القسم</td><td style="padding:9px 14px">'+requesterDept+'</td></tr>'
      +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold">الأولوية</td><td style="padding:9px 14px">'+priority+'</td></tr>'
      +'<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold">الوصف</td><td style="padding:9px 14px">'+description+'</td></tr>'
      +'</table>'
      +'<a href="'+claimUrl+'" style="background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;padding:13px 22px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block">✋ استلام البلاغ مباشرة</a>'
      +'</div><div style="padding:14px 28px;background:#f8fafc;border-top:1px solid #e2e8f0;text-align:center;font-size:11px;color:#94a3b8">IT Help Desk</div>'
      +'</div></body></html>';
    MailApp.sendEmail({to:assignee.email,subject:'📋 تم تعيينك لبلاغ ['+ticketId+'] — '+problemType,htmlBody:html});
  }catch(e){Logger.log('notifyAssignee: '+e.toString());}
}

function buildResultPage(title, message, color, ticketId) {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;background:#f8fafc;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:20px}.card{background:#fff;border-radius:16px;padding:36px 32px;max-width:420px;width:100%;text-align:center;box-shadow:0 8px 32px rgba(0,0,0,.12)}.icon{width:70px;height:70px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:32px;margin:0 auto 20px}h2{font-size:20px;margin-bottom:12px;color:#0f172a}p{font-size:14px;color:#64748b;line-height:1.7;margin-bottom:24px}a.btn{display:inline-block;padding:12px 28px;border-radius:10px;text-decoration:none;font-size:14px;font-weight:bold;color:#fff;margin:4px}</style></head>'
    +'<body><div class="card"><div class="icon" style="background:'+color+'20;color:'+color+'">'+(title.includes('✅')?'✅':title.includes('⚠️')?'⚠️':'❌')+'</div>'
    +'<h2 dir="rtl">'+title+'</h2><p dir="rtl">'+message+'</p>'
    +(ticketId?'<a class="btn" href="'+SYSTEM_URL+'" style="background:linear-gradient(135deg,#6366f1,#8b5cf6)">فتح النظام</a>':'')
    +'</div></body></html>';
}

function doPost_proxy(req) {
  try {
    var action=req.action,data=req.data||{},token=req.token||'';
    if(action==='auth.login')return handleLogin(data);
    var user=verifyToken(token);
    if(!user)return fail('انتهت الجلسة — أعد تسجيل الدخول');
    updateLastLogin(user.id);
    return route(action,data,user);
  }catch(e){Logger.log('doPost_proxy error: '+e.toString());return fail('خطأ داخلي: '+e.message);}
}

function route(action,data,user){
  switch(action){
    case 'auth.check':              return handleAuthCheck(user);
    case 'stats.dashboard':         return handleDashboard(user);
    case 'stats.user':              return handleUserStats(user);
    case 'tickets.list':            return handleTicketsList(data,user);
    case 'tickets.create':          return handleTicketCreate(data,user);
    case 'tickets.update':          return handleTicketUpdate(data,user);
    case 'tickets.claim':           return handleTicketClaim(data,user);
    case 'tickets.unassign':        return handleTicketUnassign(data,user);
    case 'tickets.assign':          return handleTicketAssign(data,user);
    case 'tickets.help':            return handleTicketHelp(data,user);
    case 'tickets.cancelHelp':      return handleTicketCancelHelp(data,user);
    case 'tickets.attach':          return handleTicketAttach(data,user);
    case 'tickets.delete':          return handleTicketDelete(data,user);
    case 'devices.checkin':         return handleDeviceCheckin(data,user);
    case 'devices.sendToTech':      return handleDeviceSendToTech(data,user);
    case 'devices.techRepair':      return handleDeviceTechRepair(data,user);
    case 'devices.receiveFromTech': return handleDeviceReceiveFromTech(data,user);
    case 'devices.finalDeliver':    return handleDeviceFinalDeliver(data,user);
    case 'devices.list':            return handleDevicesList(data,user);
    case 'users.list':              return handleUsersList(user);
    case 'users.add':               return handleUsersAdd(data,user);
    case 'users.toggle':            return handleUsersToggle(data,user);
    case 'users.delete':            return handleUsersDelete(data,user);
    case 'users.resetPw':           return handleUsersResetPw(data,user);
    case 'user.myInfo':             return handleUserMyInfo(user);
    case 'internetUsers.list':      return handleInternetUsersList(user);
    case 'internetUsers.set':       return handleInternetUsersSet(data,user);
    case 'internetUsers.search':    return handleInternetUsersSearch(data,user);
    case 'guidelines.list':         return handleGuidelinesList();
    case 'guidelines.add':          return handleGuidelinesAdd(data,user);
    case 'guidelines.update':       return handleGuidelinesUpdate(data,user);
    case 'guidelines.delete':       return handleGuidelinesDelete(data,user);
    case 'manual.list':             return handleManualList(data);
    case 'manual.add':              return handleManualAdd(data,user);
    case 'manual.update':           return handleManualUpdate(data,user);
    case 'manual.delete':           return handleManualDelete(data,user);
    case 'depts.list':              return handleDeptsList();
    case 'depts.add':               return handleDeptsAdd(data,user);
    case 'depts.delete':            return handleDeptsDelete(data,user);
    // aliases للفرونت
    case 'departments.list':        return handleDeptsList();
    case 'departments.add':         return handleDeptsAdd(data,user);
    case 'departments.delete':      return handleDeptsDelete(data,user);
    case 'internet.users.list':     return handleInternetUsersList(user);
    case 'internet.users.set':      return handleInternetUsersSet(data,user);
    case 'internet.users.update':   return handleInternetUsersSet(data,user);
    case 'internet.users.search':   return handleInternetUsersSearch(data,user);
    case 'tickets.myList':          return handleTicketsList({requesterId:user.empId},user);
    case 'knowledge.list':          return ok({items:[]});
    case 'notifications.list':      return handleNotificationsList(user);
    case 'notifications.markAllRead': return ok({message:'تم'});
    case 'notifications.broadcastClaim':  return handleBroadcastClaim(data,user);
    case 'notifications.broadcastAssign': return handleBroadcastAssign(data,user);
    case 'users.create':            return handleUsersAdd(data,user);
    case 'users.setActive':         return handleUsersToggle(data,user);
    case 'users.resetPassword':     return handleUsersResetPw(data,user);
    case 'devices.repairDone':      return handleDeviceTechRepair(data,user);
    case 'devices.receiveBack':     return handleDeviceReceiveFromTech(data,user);
    case 'devices.deliver':         return handleDeviceFinalDeliver(data,user);
    case 'manual.delete':           return handleManualDelete(data,user);
    case 'notifications.check':     return handleNotificationsCheck(data,user);
    case 'ai.chat':                 return handleAIChat(data,user);
    case 'stats.tech':              return handleTechStats(user);
    default: return fail('action غير معروف: '+action);
  }
}

// ══════ UTILITIES ══════
function ok(d){return Object.assign({success:true},d||{});}
function fail(m){return{success:false,message:m||'خطأ'};}
function genId(p){return p+'-'+Utilities.formatDate(new Date(),'Asia/Baghdad','yyyyMMdd')+'-'+Math.random().toString(36).substr(2,5).toUpperCase();}
function now(){return Utilities.formatDate(new Date(),'Asia/Baghdad','dd/MM/yyyy HH:mm');}
function nowIso(){return new Date().toISOString();}
function sha256(str){var raw=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,str,Utilities.Charset.UTF_8);return raw.map(function(b){return('0'+(b<0?b+256:b).toString(16)).slice(-2);}).join('');}
function getSheet(name){var sh=SS.getSheetByName(name);if(!sh){var allowCreate=['InternetUsers','السجل'];if(allowCreate.indexOf(name)>=0){sh=SS.insertSheet(name);initSheet(sh,name);}else throw new Error('شيت "'+name+'" غير موجود');}return sh;}
function sheetData(name){
  var sh=getSheet(name),vals=sh.getDataRange().getValues();
  if(vals.length<2)return{headers:vals[0]||[],rows:[]};
  if(name===SH_USERS){
    var rawHeaders=vals[0];
    var engHeaders=rawHeaders.map(function(h){
      var clean=String(h).trim().replace(/^#+|#+$/g,'').trim();
      for(var key in COL_USERS_ALIASES){
        var aliases=COL_USERS_ALIASES[key];
        for(var a=0;a<aliases.length;a++){if(clean===aliases[a]||String(h).trim()===aliases[a])return key;}
      }
      return clean||String(h);
    });
    return{headers:engHeaders,rows:vals.slice(1)};
  }
  return{headers:vals[0],rows:vals.slice(1)};
}

function getUsersColIdxFlex(sh,key){
  var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  return findColIndex(headers,key);
}
function appendRow(sheetName,obj,headers){
  var sh=getSheet(sheetName),row;
  if(sheetName===SH_USERS){var rh=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];row=rh.map(function(h){var k=Object.keys(COL_USERS).find(function(x){return COL_USERS[x]===h;});return(k&&obj[k]!==undefined)?obj[k]:(obj[h]!==undefined?obj[h]:'');});}
  else{row=headers.map(function(h){return obj[h]!==undefined?obj[h]:'';});}
  sh.appendRow(row);
}
function updateRow(sheetName,rowIndex,obj,headers){var sh=getSheet(sheetName),row=headers.map(function(h){return obj[h]!==undefined?obj[h]:'';});sh.getRange(rowIndex+2,1,1,row.length).setValues([row]);}
function getUsersColIdx(sh,engKey){var headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];var ar=COL_USERS[engKey];if(ar){var i=headers.indexOf(ar);if(i>=0)return i;}return headers.indexOf(engKey);}
function arrToObj(arr,headers){var obj={};headers.forEach(function(h,i){obj[h]=arr[i]!==undefined?arr[i]:'';});return obj;}
function initSheet(sh,name){var h={InternetUsers:['fullName','internetUser','dept'],السجل:['id','type','message','createdAt']};if(h[name])sh.appendRow(h[name]);}

// ══════ AUTH ══════
function handleLogin(data){
  var empId=(data.empId||'').toString().trim(),pwHash=(data.pwHash||'').toString().trim().toLowerCase();
  if(!empId||!pwHash)return fail('البيانات ناقصة');
  var empIdNum=parseInt(empId),sh=getSheet(SH_USERS),allData=sh.getDataRange().getValues(),rh=allData[0];
  var idCol=findColIndex(rh,'id'),empIdCol=findColIndex(rh,'empId'),pwCol=findColIndex(rh,'pwHash');
  var activeCol=findColIndex(rh,'active'),roleCol=findColIndex(rh,'role');
  var firstCol=findColIndex(rh,'firstName'),lastCol=findColIndex(rh,'lastName');
  var deptCol=findColIndex(rh,'dept'),phoneCol=findColIndex(rh,'phone'),emailCol=findColIndex(rh,'notifyEmail');
  for(var i=1;i<allData.length;i++){
    var r=allData[i],rEmpId=String(r[empIdCol]||'').trim();
    if(rEmpId!==empId&&rEmpId!==empId.toUpperCase()&&(isNaN(empIdNum)||parseInt(rEmpId)!==empIdNum))continue;
    var rPw=String(r[pwCol]||'').trim().toLowerCase(),rAct=r[activeCol];
    if(!(rAct===true||String(rAct).toUpperCase()==='TRUE'))return fail('الحساب معطّل');
    if(rPw!==pwHash)return fail('الرمز السري غير صحيح');
    var userId=String(r[idCol]||'').trim()||('USR-'+rEmpId);
    var user={id:userId,empId:rEmpId,firstName:String(r[firstCol]||'').trim(),lastName:String(r[lastCol]||'').trim(),dept:String(r[deptCol]||'').trim(),role:String(r[roleCol]||'user').trim().toLowerCase(),phone:String(r[phoneCol]||'').trim(),notifyEmail:emailCol>=0?String(r[emailCol]||'').trim():'',isAdmin:String(r[roleCol]||'').trim().toLowerCase()==='admin'};
    return ok({token:createToken(user.id,user.empId,user.role),user:user});
  }
  return fail('رقم البصمة غير موجود');
}
function handleAuthCheck(user){return ok({user:user});}
var TOKEN_SECRET='ITHelpDesk_SecretKey_2025';
function createToken(userId,empId,role){var p=userId+'|'+empId+'|'+role+'|'+Date.now(),sig=sha256(p+TOKEN_SECRET);return Utilities.base64Encode(p+'||'+sig);}
function verifyToken(token){
  try{
    if(!token)return null;
    var dec=Utilities.newBlob(Utilities.base64Decode(token)).getDataAsString(),parts=dec.split('||');
    if(parts.length!==2)return null;
    var payload=parts[0],sig=parts[1];
    if(sha256(payload+TOKEN_SECRET)!==sig)return null;
    var p=payload.split('|');if(p.length<4)return null;
    var userId=p[0],empId=p[1],role=p[2],ts=parseInt(p[3]);
    if(Date.now()-ts>12*3600*1000)return null;
    var sh=getSheet(SH_USERS),allD=sh.getDataRange().getValues(),rh=allD[0];
    var idC=findColIndex(rh,'id'),acC=findColIndex(rh,'active'),empC=findColIndex(rh,'empId');
    var rlC=findColIndex(rh,'role'),fnC=findColIndex(rh,'firstName'),lnC=findColIndex(rh,'lastName');
    var dpC=findColIndex(rh,'dept'),phC=findColIndex(rh,'phone'),emC=findColIndex(rh,'notifyEmail');
    for(var i=1;i<allD.length;i++){
      var r=allD[i],rId=String(r[idC]||'').trim(),rEmp=String(r[empC]||'').trim();
      if(rId!==userId&&rEmp!==empId)continue;
      if(!(r[acC]===true||String(r[acC]).toUpperCase()==='TRUE'))return null;
      return{id:rId||('USR-'+rEmp),empId:rEmp,firstName:String(r[fnC]||'').trim(),lastName:String(r[lnC]||'').trim(),dept:String(r[dpC]||'').trim(),role:String(r[rlC]||'user').trim().toLowerCase(),phone:String(r[phC]||'').trim(),notifyEmail:emC>=0?String(r[emC]||'').trim():'',isAdmin:String(r[rlC]||'').trim().toLowerCase()==='admin'};
    }
    return null;
  }catch(e){return null;}
}
function updateLastLogin(userId){
  try{var sh=getSheet(SH_USERS),d=sh.getDataRange().getValues(),hi=d[0];
  var idCol=findColIndex(hi,'id'),llCol=findColIndex(hi,'lastLogin');
  for(var i=1;i<d.length;i++){if(String(d[i][idCol])===String(userId)){if(llCol>=0)sh.getRange(i+1,llCol+1).setValue(now());break;}}}catch(e){}
}

// ══════ TICKETS ══════
function handleTicketsList(data,user){
  var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues();
  var q=(data.q||'').toLowerCase(),statusF=data.status||'all',priorityF=data.priority||'all';
  var overdueOnly=data.overdueOnly||false,limitN=parseInt(data.limit)||0;
  var isIT=ROLES_IT.indexOf(user.role)>=0,isMgr=ROLES_MGR.indexOf(user.role)>=0;
  var tickets=[],now48=new Date(Date.now()-48*3600*1000);
  for(var i=allData.length-1;i>=0;i--){
    var r=allData[i];
    var tid=r[COL_TICKETS.id]!==undefined?String(r[COL_TICKETS.id]||'').trim():'';
    if(!tid||tid.indexOf('TKT')<0)continue;
    var reqId=String(r[COL_TICKETS.requesterId]||'').trim();
    if(!isIT&&!isMgr){if(reqId!==user.empId&&reqId!==user.id&&String(reqId)!==String(user.empId))continue;}
    var status=String(r[COL_TICKETS.status]||'').trim(),priority=String(r[COL_TICKETS.priority]||'').trim(),created=String(r[COL_TICKETS.createdAt]||'').trim();
    if(statusF!=='all'&&status!==statusF)continue;
    if(priorityF!=='all'&&priority!==priorityF)continue;
    var isOverdue=false;
    if(status!=='تم حل البلاغ'&&created){try{var pp=created.split(/[\/\s:]/),d2=new Date(pp[2],pp[1]-1,pp[0]);if(d2<now48)isOverdue=true;}catch(e2){}}
    if(overdueOnly&&!isOverdue)continue;
    var desc=String(r[COL_TICKETS.description]||'').trim();
    var probType=String(r[COL_TICKETS.problemType]||'').trim();
    var reqName=String(r[COL_TICKETS.requesterName]||'').trim();
    var reqDept=String(r[COL_TICKETS.requesterDept]||'').trim();
    if(q&&!(tid+probType+desc+reqName+reqDept).toLowerCase().includes(q))continue;
    tickets.push({id:tid,problemType:probType,priority:priority||'متوسطة',deviceId:String(r[COL_TICKETS.deviceId]||'').trim(),desc:desc,status:status||'جديدة',requesterId:reqId,requesterName:reqName,requesterDept:reqDept,assignedId:String(r[COL_TICKETS.assignedId]||'').trim(),assignedName:String(r[COL_TICKETS.assignedName]||'').trim(),solution:'',notes:'',createdAt:created,updatedAt:created,solvedAt:status==='تم حل البلاغ'?created:'',history:[],isOverdue:isOverdue,helpRequested:false});
    if(limitN>0&&tickets.length>=limitN)break;
  }
  return ok({tickets:tickets});
}

function handleTicketCreate(data,user){
  var id=genId('TKT'),sh=getSheet(SH_TICKETS),row=new Array(13).fill('');
  row[COL_TICKETS.id]=id;row[COL_TICKETS.createdAt]=now();
  row[COL_TICKETS.requesterName]=user.firstName+' '+user.lastName;
  row[COL_TICKETS.requesterId]=user.empId;row[COL_TICKETS.requesterDept]=user.dept;
  row[COL_TICKETS.deviceId]=data.deviceId||'';row[COL_TICKETS.problemType]=data.problemType||'';
  row[COL_TICKETS.priority]=data.priority||'متوسطة';row[COL_TICKETS.description]=data.description||'';
  row[COL_TICKETS.status]='جديدة';sh.appendRow(row);
  notifyITTeam({id:id,problemType:data.problemType||'',priority:data.priority||'متوسطة',description:data.description||'',requesterName:user.firstName+' '+user.lastName,requesterDept:user.dept});
  return ok({message:'تم إرسال البلاغ بنجاح',ticketId:id});
}

function handleTicketUpdate(data,user){
  if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');
  var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues();
  for(var i=0;i<allData.length;i++){
    if(String(allData[i][COL_TICKETS.id]).trim()===data.ticketId){
      if(data.status)sh.getRange(i+1,COL_TICKETS.status+1).setValue(data.status);
      if(data.status==='تم حل البلاغ'){var em=getEmailByEmpId(String(allData[i][COL_TICKETS.requesterId]));if(em)sendSolvedEmail(em,data.ticketId,String(allData[i][COL_TICKETS.requesterName]),data.solution||'');}
      return ok({message:'تم تحديث البلاغ'});
    }
  }
  return fail('البلاغ غير موجود');
}

function handleTicketClaim(data,user){if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');return assignTicketDirect(data.ticketId,user.empId,user.firstName+' '+user.lastName,user);}

function handleTicketUnassign(data,user){
  if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');
  var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues();
  for(var i=0;i<allData.length;i++){if(String(allData[i][COL_TICKETS.id]).trim()===data.ticketId){sh.getRange(i+1,COL_TICKETS.assignedId+1).setValue('');sh.getRange(i+1,COL_TICKETS.assignedName+1).setValue('');sh.getRange(i+1,COL_TICKETS.status+1).setValue('جديدة');return ok({message:'تم إلغاء التعيين'});}}
  return fail('البلاغ غير موجود');
}

function handleTicketAssign(data,user){
  if(ROLES_IT.indexOf(user.role)<0&&ROLES_MGR.indexOf(user.role)<0)return fail('غير مصرح');
  var sh2=getSheet(SH_USERS),uData=sh2.getDataRange().getValues(),an='';
  for(var i=1;i<uData.length;i++){if(String(uData[i][getUsersColIdx(sh2,'id')]||'').trim()===data.assigneeId){an=(String(uData[i][getUsersColIdx(sh2,'firstName')]||'')+' '+String(uData[i][getUsersColIdx(sh2,'lastName')]||'')).trim();break;}}
  var result=assignTicketDirect(data.ticketId,data.assigneeId,an,user);
  if(result.success&&an){
    try{
      var assigneeData=getEmployeeByEmpId(data.assigneeId);
      if(!assigneeData){for(var j=1;j<uData.length;j++){if(String(uData[j][getUsersColIdx(sh2,'id')]||'').trim()===data.assigneeId){assigneeData={empId:String(uData[j][getUsersColIdx(sh2,'empId')]||'').trim(),firstName:String(uData[j][getUsersColIdx(sh2,'firstName')]||'').trim(),lastName:String(uData[j][getUsersColIdx(sh2,'lastName')]||'').trim(),email:String(uData[j][getUsersColIdx(sh2,'notifyEmail')]||'').trim()};break;}}}
      if(assigneeData&&assigneeData.email){var tSh=getSheet(SH_TICKETS),tAll=tSh.getDataRange().getValues();for(var k=0;k<tAll.length;k++){if(String(tAll[k][COL_TICKETS.id]).trim()===data.ticketId){notifyAssignee(data.ticketId,assigneeData,tAll[k]);break;}}}
    }catch(e2){Logger.log('notifyAssignee: '+e2.toString());}
  }
  return result;
}

function assignTicketDirect(ticketId,assigneeId,assigneeName,byUser){
  var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues();
  for(var i=0;i<allData.length;i++){if(String(allData[i][COL_TICKETS.id]).trim()===ticketId){sh.getRange(i+1,COL_TICKETS.assignedId+1).setValue(assigneeId);sh.getRange(i+1,COL_TICKETS.assignedName+1).setValue(assigneeName);sh.getRange(i+1,COL_TICKETS.status+1).setValue('معينة');return ok({message:'تم التعيين إلى '+assigneeName});}}
  return fail('البلاغ غير موجود');
}

function handleTicketHelp(data,user){var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues();for(var i=0;i<allData.length;i++){if(String(allData[i][COL_TICKETS.id]).trim()===data.ticketId){return ok({message:'تم إرسال طلب المساعدة'});}}return fail('البلاغ غير موجود');}
function handleTicketCancelHelp(data,user){return ok({message:'تم إلغاء طلب المساعدة'});}

function handleTicketAttach(data,user){
  try{var folders=DriveApp.getFoldersByName('HelpDesk_Attachments'),folder;if(folders.hasNext())folder=folders.next();else folder=DriveApp.createFolder('HelpDesk_Attachments');var bytes=Utilities.base64Decode(data.base64),blob=Utilities.newBlob(bytes,data.mimeType||'image/jpeg',data.fileName||'attachment'),file=folder.createFile(blob);file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);return ok({message:'تم رفع الملف',url:'https://drive.google.com/file/d/'+file.getId()+'/view'});}catch(e){return fail('خطأ في رفع الملف: '+e.message);}
}

function handleTicketDelete(data,user){
  if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح — أدمن فقط');
  var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues();
  for(var i=0;i<allData.length;i++){if(String(allData[i][COL_TICKETS.id]).trim()===data.ticketId){sh.deleteRow(i+1);return ok({message:'تم حذف البلاغ'});}}
  return fail('البلاغ غير موجود');
}

// ══════ STATS ══════
function handleDashboard(user){
  var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues(),now48=new Date(Date.now()-48*3600*1000);
  var stats={total:0,open:0,assigned:0,inProgress:0,closed:0,overdue:0,help:0};
  var byPriority={'عاجلة':0,'عالية':0,'متوسطة':0,'منخفضة':0},byDate={},byIT={},byProb={};
  allData.forEach(function(r){
    var tid=String(r[COL_TICKETS.id]||'').trim();if(!tid||tid.indexOf('TKT')<0)return;
    var status=String(r[COL_TICKETS.status]||''),priority=String(r[COL_TICKETS.priority]||''),created=String(r[COL_TICKETS.createdAt]||''),assigned=String(r[COL_TICKETS.assignedId]||''),aName=String(r[COL_TICKETS.assignedName]||''),problem=String(r[COL_TICKETS.problemType]||'');
    stats.total++;if(status==='جديدة')stats.open++;if(status==='معينة')stats.assigned++;if(status==='قيد المعالجة')stats.inProgress++;if(status==='تم حل البلاغ')stats.closed++;
    if(byPriority[priority]!==undefined)byPriority[priority]++;
    if(status!=='تم حل البلاغ'&&created){try{var p=created.split(/[\/\s:]/),cd=new Date(p[2],p[1]-1,p[0]);if(cd<now48)stats.overdue++;}catch(e2){}}
    var day=created.split(' ')[0]||'';if(day)byDate[day]=(byDate[day]||0)+1;
    if(assigned&&aName){if(!byIT[assigned])byIT[assigned]={name:aName,total:0,solved:0,open:0};byIT[assigned].total++;if(status==='تم حل البلاغ')byIT[assigned].solved++;else byIT[assigned].open++;}
    if(problem)byProb[problem]=(byProb[problem]||0)+1;
  });
  var itPerf=Object.values(byIT).map(function(p){return{name:p.name,total:p.total,solved:p.solved,open:p.open,rate:p.total?Math.round(p.solved/p.total*100):0,avg:0};}).sort(function(a,b){return b.total-a.total;}).slice(0,8);
  var topProblems=Object.entries(byProb).sort(function(a,b){return b[1]-a[1];}).slice(0,6).map(function(e){return{name:e[0],count:e[1]};});
  var daily=Object.entries(byDate).slice(-14).map(function(e){return{d:e[0],c:e[1]};});
  return ok({stats:stats,byPriority:byPriority,itPerformance:itPerf,topProblems:topProblems,daily:daily});
}

function handleUserStats(user){
  var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues(),stats={total:0,open:0,assigned:0,inProgress:0,closed:0},myT=[];
  allData.forEach(function(r){var tid=String(r[COL_TICKETS.id]||'').trim();if(!tid||tid.indexOf('TKT')<0)return;var reqId=String(r[COL_TICKETS.requesterId]||'').trim();if(reqId!==user.empId&&reqId!==user.id)return;stats.total++;var st=String(r[COL_TICKETS.status]||'');if(st==='جديدة')stats.open++;if(st==='معينة')stats.assigned++;if(st==='قيد المعالجة')stats.inProgress++;if(st==='تم حل البلاغ')stats.closed++;myT.push(r);});
  var recent=[];myT.slice(-5).reverse().forEach(function(r){recent.push({id:String(r[COL_TICKETS.id]||''),problemType:String(r[COL_TICKETS.problemType]||''),status:String(r[COL_TICKETS.status]||''),createdAt:String(r[COL_TICKETS.createdAt]||'')});});
  return ok({stats:stats,recent:recent});
}

// ══════ NOTIFICATIONS ══════
function notifyITTeam(ticket){
  try{
    var uD=sheetData(SH_USERS),uH=uD.headers;
    uD.rows.forEach(function(r){
      var role=String(r[uH.indexOf('role')]||''),email=String(r[uH.indexOf('notifyEmail')]||'').trim();
      var empId=String(r[uH.indexOf('empId')]||'').trim(),active=r[uH.indexOf('active')];
      if(ROLES_STAFF.indexOf(role)<0)return;
      if(!email||!(active===true||String(active).toLowerCase()==='true'))return;
      var token=generateEmailToken(ticket.id,empId);
      if(role==='manager'){
        var assignUrl=SYSTEM_URL+'?action=assignForm&tid='+encodeURIComponent(ticket.id)+'&uid='+encodeURIComponent(empId)+'&token='+token;
        try{MailApp.sendEmail({to:email,subject:'🔔 بلاغ جديد يحتاج تعيين ['+ticket.id+'] — '+ticket.problemType,htmlBody:buildManagerEmail(ticket,assignUrl)});}catch(e2){}
      }else{
        var claimUrl=SYSTEM_URL+'?action=claim&tid='+encodeURIComponent(ticket.id)+'&uid='+encodeURIComponent(empId)+'&token='+token;
        var solveUrl=SYSTEM_URL+'?action=solve&tid='+encodeURIComponent(ticket.id)+'&uid='+encodeURIComponent(empId)+'&token='+token;
        try{MailApp.sendEmail({to:email,subject:'🔔 بلاغ جديد ['+ticket.id+'] — '+ticket.problemType,htmlBody:buildTicketEmail(ticket,claimUrl,solveUrl)});}catch(e2){}
      }
    });
  }catch(e){}
}

function buildTicketEmail(ticket,claimUrl,solveUrl){
  claimUrl=claimUrl||SYSTEM_URL;solveUrl=solveUrl||SYSTEM_URL;
  var pc=ticket.priority==='عاجلة'?'#ef4444':ticket.priority==='عالية'?'#f59e0b':ticket.priority==='متوسطة'?'#3b82f6':'#10b981';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body dir="rtl" style="margin:0;padding:0;background:#f8fafc;font-family:Arial,sans-serif">'
    +'<div style="max-width:600px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1)">'
    +'<div style="background:linear-gradient(135deg,#6366f1,#8b5cf6);padding:24px 28px">'
    +'<h2 style="margin:0;color:#fff;font-size:18px">🔔 بلاغ جديد — '+ticket.problemType+'</h2>'
    +'<p style="margin:6px 0 0;color:rgba(255,255,255,.7);font-size:13px">رقم البلاغ: '+ticket.id+'</p>'
    +'</div><div style="padding:24px 28px">'
    +'<table style="width:100%;border-collapse:collapse;margin-bottom:20px">'
    +'<tr><td style="padding:10px 14px;background:#f8fafc;font-weight:bold;font-size:13px;width:35%">المُبلِّغ</td><td style="padding:10px 14px;font-size:13px">'+ticket.requesterName+'</td></tr>'
    +'<tr><td style="padding:10px 14px;background:#f8fafc;font-weight:bold;font-size:13px">القسم</td><td style="padding:10px 14px;font-size:13px">'+ticket.requesterDept+'</td></tr>'
    +'<tr><td style="padding:10px 14px;background:#f8fafc;font-weight:bold;font-size:13px">الأولوية</td><td style="padding:10px 14px"><span style="background:'+pc+'20;color:'+pc+';padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold">'+ticket.priority+'</span></td></tr>'
    +'<tr><td style="padding:10px 14px;background:#f8fafc;font-weight:bold;font-size:13px">وصف المشكلة</td><td style="padding:10px 14px;font-size:13px">'+ticket.description+'</td></tr>'
    +'</table>'
    +'<div style="display:flex;gap:10px;flex-wrap:wrap">'
    +'<a href="'+claimUrl+'" style="background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;padding:12px 24px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block">✋ استلام البلاغ مباشرة</a>'
    +'<a href="'+solveUrl+'" style="background:#ecfdf5;color:#059669;padding:12px 24px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block;border:2px solid #a7f3d0">✅ تم حل البلاغ</a>'
    +'<a href="'+SYSTEM_URL+'" style="background:#f8fafc;color:#374151;padding:12px 24px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block;border:2px solid #e2e8f0">🔍 فتح النظام</a>'
    +'</div></div>'
    +'<div style="padding:14px 28px;background:#f8fafc;border-top:1px solid #e2e8f0;text-align:center;font-size:11px;color:#94a3b8">IT Help Desk — نظام الدعم التقني</div>'
    +'</div></body></html>';
}

function sendSolvedEmail(to,ticketId,userName,solution){try{MailApp.sendEmail({to:to,subject:'✅ تم حل بلاغك ['+ticketId+']',htmlBody:'<div dir="rtl" style="font-family:Arial;padding:20px"><h2 style="color:#059669">✅ تم حل بلاغك</h2><p>عزيزي '+userName+'</p><p>البلاغ رقم <strong>'+ticketId+'</strong> تم حله</p><p>'+(solution||'تم المعالجة')+'</p></div>'});}catch(e){}}
function getEmailByEmpId(empId){var sh=getSheet(SH_USERS),d=sh.getDataRange().getValues(),rh=d[0],empCol=findColIndex(rh,'empId'),emailCol=findColIndex(rh,'notifyEmail');for(var i=1;i<d.length;i++){if(String(d[i][empCol]).trim()===String(empId).trim())return emailCol>=0?String(d[i][emailCol]||'').trim()||null:null;}return null;}
function getEmailForUser(userId){var d=sheetData(SH_USERS),hi=d.headers;for(var i=0;i<d.rows.length;i++){if(String(d.rows[i][hi.indexOf('id')])===userId)return String(d.rows[i][hi.indexOf('notifyEmail')]||'').trim()||null;}return null;}

function handleNotificationsCheck(data,user){
  if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');
  var since=data.since?new Date(data.since):new Date(Date.now()-600000),sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues(),newTickets=[];
  allData.forEach(function(r){var tid=String(r[COL_TICKETS.id]||'').trim();if(!tid||tid.indexOf('TKT')<0)return;if(String(r[COL_TICKETS.status]||'')!=='جديدة')return;var created=String(r[COL_TICKETS.createdAt]||'');try{var p=created.split(/[\/\s:]/),cd=new Date(p[2],p[1]-1,p[0],p[3]||0,p[4]||0);if(cd>since)newTickets.push({id:tid,problemType:String(r[COL_TICKETS.problemType]||''),requesterName:String(r[COL_TICKETS.requesterName]||''),requesterDept:String(r[COL_TICKETS.requesterDept]||''),priority:String(r[COL_TICKETS.priority]||'')});}catch(e2){}});
  return ok({newTickets:newTickets,timestamp:nowIso()});
}

// ══════ DEVICES ══════
var DEVICE_HEADERS=['devId','ticketId','deviceType','deviceDesc','ownerName','ownerDept','ownerEmpId','status','checkinBy','checkinAt','techId','techName','sentToTechAt','techResult','techResultLabel','workDone','failReason','techDoneAt','itReceivedAt','itNotes','deliveredAt','finalCondition','finalNotes'];
function ensureDeviceSheet(){var sh=SS.getSheetByName(SH_DEVICES);if(!sh){sh=SS.insertSheet(SH_DEVICES);sh.appendRow(DEVICE_HEADERS);return sh;}var firstRow=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];var hasHeaders=firstRow.some(function(h){return String(h).trim()==='devId';});if(!hasHeaders){sh.insertRowBefore(1);sh.getRange(1,1,1,DEVICE_HEADERS.length).setValues([DEVICE_HEADERS]);}return sh;}
function handleDeviceCheckin(data,user){if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');var devId=genId('DEV'),sh=ensureDeviceSheet();var hi=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),DEVICE_HEADERS.length)).getValues()[0];var row=new Array(Math.max(hi.length,DEVICE_HEADERS.length)).fill('');function setD(k,v){var i=hi.indexOf(k);if(i>=0)row[i]=v;}setD('devId',devId);setD('ticketId',data.ticketId||'');setD('deviceType',data.deviceType||'');setD('deviceDesc',data.deviceDesc||'');setD('ownerName',data.ownerName||'');setD('ownerDept',data.ownerDept||'');setD('status','عند IT');setD('checkinBy',user.firstName+' '+user.lastName);setD('checkinAt',now());setD('finalNotes',data.notes||'');sh.appendRow(row);return ok({message:'تم تسجيل استلام الجهاز',devId:devId});}
function handleDeviceSendToTech(data,user){if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');var uD=sheetData(SH_USERS),uH=uD.headers,tn='';for(var i=0;i<uD.rows.length;i++){if(String(uD.rows[i][uH.indexOf('id')])===data.techId){tn=String(uD.rows[i][uH.indexOf('firstName')])+' '+String(uD.rows[i][uH.indexOf('lastName')]);break;}}return updateDeviceField(data.devId,{status:'عند الفني',techId:data.techId||'',techName:tn,sentToTechAt:now(),finalNotes:data.notes||''});}
function handleDeviceTechRepair(data,user){if(ROLES_TECH.indexOf(user.role)<0)return fail('غير مصرح');var labels={fixed:'تم الإصلاح الكامل',partial:'إصلاح جزئي',failed:'تعذّر الإصلاح',needs_parts:'يحتاج قطع غيار'};return updateDeviceField(data.devId,{status:data.readyForPickup?'أُنجز — بانتظار IT':'عند الفني',techResult:data.result||'fixed',techResultLabel:labels[data.result]||data.result,workDone:data.workDone||'',failReason:data.failReason||'',techDoneAt:now()});}
function handleDeviceReceiveFromTech(data,user){if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');return updateDeviceField(data.devId,{status:'جاهز للتسليم',itReceivedAt:now(),itNotes:data.itNotes||''});}
function handleDeviceFinalDeliver(data,user){
  if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');
  var res=updateDeviceField(data.devId,{status:'مُسلَّم',deliveredAt:now(),finalCondition:data.condition||'',finalNotes:data.notes||''});if(!res.success)return res;
  var sh=SS.getSheetByName(SH_DEVICES),allD=sh?sh.getDataRange().getValues():[],hi=allD[0]||[],tid='';
  for(var i=1;i<allD.length;i++){if(String(allD[i][hi.indexOf('devId')]||'').trim()===data.devId){tid=String(allD[i][hi.indexOf('ticketId')]||'');break;}}
  if(tid)handleTicketUpdate({ticketId:tid,status:'تم حل البلاغ',solution:'تم تسليم الجهاز للموظف — '+(data.condition||'تم الإصلاح')},user);
  return ok({message:'تم التسليم وإغلاق البلاغ'});
}
function handleDevicesList(data,user){
  var sh=ensureDeviceSheet(),allData=sh.getDataRange().getValues();if(allData.length<2)return ok({devices:[]});
  var hi=allData[0],statusF=data.status||'all',devices=[];
  for(var i=allData.length-1;i>=1;i--){var r=allData[i];function gd(c){var ci=hi.indexOf(c);return ci>=0?String(r[ci]||'').trim():'';}var devId=gd('devId');if(!devId)continue;var status=gd('status');if(statusF!=='all'&&status!==statusF)continue;if(user.role==='tech'&&gd('techId')!==user.id&&gd('techId')!==user.empId)continue;devices.push({devId:devId,ticketId:gd('ticketId'),deviceType:gd('deviceType'),deviceDesc:gd('deviceDesc'),ownerName:gd('ownerName'),ownerDept:gd('ownerDept'),status:status,checkinBy:gd('checkinBy'),checkinAt:gd('checkinAt'),techId:gd('techId'),techName:gd('techName'),sentToTechAt:gd('sentToTechAt'),techResult:gd('techResult'),techResultLabel:gd('techResultLabel'),workDone:gd('workDone'),failReason:gd('failReason'),techDoneAt:gd('techDoneAt'),itReceivedAt:gd('itReceivedAt'),itNotes:gd('itNotes'),deliveredAt:gd('deliveredAt'),finalCondition:gd('finalCondition')});}
  return ok({devices:devices});
}
function updateDeviceField(devId,updates){var sh=ensureDeviceSheet(),d=sh.getDataRange().getValues();if(d.length<1)return fail('الجهاز غير موجود');var hi=d[0],devIdCol=hi.indexOf('devId');if(devIdCol<0)devIdCol=0;for(var i=1;i<d.length;i++){if(String(d[i][devIdCol]).trim()===devId){Object.keys(updates).forEach(function(k){var col=hi.indexOf(k);if(col>=0)sh.getRange(i+1,col+1).setValue(updates[k]);});return ok({message:'تم تحديث الجهاز'});}}return fail('الجهاز غير موجود: '+devId);}

// ══════ USERS ══════
function handleUsersList(user){
  if(ROLES_ADMIN.indexOf(user.role)<0&&ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');
  var sh=getSheet(SH_USERS),allData=sh.getDataRange().getValues();
  if(allData.length<2)return ok({users:[]});
  var rh=allData[0];
  // نحسب الـ indexes مرة واحدة خارج الـ loop
  var idIdx    = findColIndex(rh,'id');
  var empIdx   = findColIndex(rh,'empId');
  var fnIdx    = findColIndex(rh,'firstName');
  var lnIdx    = findColIndex(rh,'lastName');
  var deptIdx  = findColIndex(rh,'dept');
  var roleIdx  = findColIndex(rh,'role');
  var phoneIdx = findColIndex(rh,'phone');
  var acIdx    = findColIndex(rh,'active');
  var llIdx    = findColIndex(rh,'lastLogin');
  var caIdx    = findColIndex(rh,'createdAt');

  function getCell(row, idx){ return idx>=0?String(row[idx]||'').trim():''; }

  var users=[];
  for(var i=1;i<allData.length;i++){
    var r=allData[i];
    var empId=getCell(r,empIdx);
    if(!empId)continue;
    var av=acIdx>=0?r[acIdx]:true;
    users.push({
      id:       getCell(r,idIdx)||('USR-'+empId),
      empId:    empId,
      firstName:getCell(r,fnIdx),
      lastName: getCell(r,lnIdx),
      dept:     getCell(r,deptIdx),
      department:getCell(r,deptIdx),
      role:     getCell(r,roleIdx)||'user',
      phone:    getCell(r,phoneIdx),
      email:    getCell(r, findColIndex(rh,'notifyEmail')),
      active:   av===true||String(av).toUpperCase()==='TRUE',
      lastLogin:getCell(r,llIdx),
      createdAt:getCell(r,caIdx)
    });
  }
  return ok({users:users});
}
function handleUsersAdd(data,user){
  if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');
  var empId=(data.empId||'').toString().trim().toUpperCase();if(!empId)return fail('رقم البصمة مطلوب');
  var d=sheetData(SH_USERS),hi=d.headers;for(var i=0;i<d.rows.length;i++){if(String(d.rows[i][hi.indexOf('empId')]).toUpperCase()===empId)return fail('رقم البصمة موجود مسبقاً');}
  var id='USR-'+Utilities.getUuid().replace(/-/g,'').slice(0,10).toUpperCase();
  var deptVal=data.dept||data.department||'';
  var pwVal=data.pwHash||data.password||'';
  var emailVal=data.notifyEmail||data.email||'';
  appendRow(SH_USERS,{id:id,empId:empId,firstName:data.firstName||'',lastName:data.lastName||'',dept:deptVal,role:data.role||'user',pwHash:pwVal,phone:data.phone||'',notifyEmail:emailVal,active:true,lastLogin:'',createdAt:now()},d.headers);
  var fullName=(data.firstName+' '+data.lastName).trim(),mr=getInternetUserForEmployee(data.firstName,data.lastName,data.dept||''),extraMsg='',autoMatch=null;
  if(mr&&mr.internetUser){saveInternetUser(empId,fullName,mr.internetUser,data.dept||'','نظام تلقائي');autoMatch={matched:true,internetUser:mr.internetUser,matchedName:mr.matchedName,confidence:mr.confidence};extraMsg=' — تم ربط يوزر الإنترنت تلقائياً: '+mr.internetUser+' ('+mr.confidence+'% تطابق)';}
  else{autoMatch={matched:false};extraMsg=' — لم يُعثر على يوزر إنترنت';}
  return ok({message:'تمت إضافة الموظف'+extraMsg,autoMatch:autoMatch});
}
function handleUsersToggle(data,user){if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');var uid=data.id||data.userId||'';var sh=getSheet(SH_USERS),d=sh.getDataRange().getValues(),rh=d[0],idCol=findColIndex(rh,'id'),acCol=findColIndex(rh,'active');for(var i=1;i<d.length;i++){if(String(d[i][idCol])===uid){var newActive=data.active!==undefined?data.active:!(d[i][acCol]===true||String(d[i][acCol]).toLowerCase()==='true');sh.getRange(i+1,acCol+1).setValue(newActive);return ok({message:(newActive?'تم تفعيل':'تم تعطيل')+' الحساب'});}}return fail('الموظف غير موجود');}
function handleUsersDelete(data,user){if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');var uid=data.id||data.userId||'';var sh=getSheet(SH_USERS),d=sh.getDataRange().getValues(),rh=d[0],idCol=findColIndex(rh,'id');for(var i=1;i<d.length;i++){if(String(d[i][idCol])===uid){sh.deleteRow(i+1);return ok({message:'تم حذف الموظف'});}}return fail('الموظف غير موجود');}
function handleUsersResetPw(data,user){if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');var uid=data.id||data.userId||'';var pw=data.pwHash||data.newPassword||data.password||'';var sh=getSheet(SH_USERS),d=sh.getDataRange().getValues(),rh=d[0],idCol=findColIndex(rh,'id'),pwCol=findColIndex(rh,'pwHash');for(var i=1;i<d.length;i++){if(String(d[i][idCol])===uid){sh.getRange(i+1,pwCol+1).setValue(pw);return ok({message:'تم تغيير الرمز السري'});}}return fail('الموظف غير موجود');}
function handleUserMyInfo(user){var result=getInternetUserForEmployee(user.firstName,user.lastName,user.dept);var internetUser=result?result.internetUser:'';var tSh=getSheet(SH_TICKETS),tData=tSh.getDataRange().getValues();var ticketCount=tData.filter(function(r){var tid=String(r[COL_TICKETS.id]||'').trim();if(!tid||tid.indexOf('TKT')<0)return false;var rid=String(r[COL_TICKETS.requesterId]||'').trim();return rid===user.empId||rid===user.id;}).length;return ok({internetUser:internetUser||null,ticketCount:ticketCount,phone:user.phone||''});}

// ══════ INTERNET USERS ══════
function smartNameMatch(a,b){if(!a||!b)return 0;function norm(n){return n.trim().replace(/\s+/g,' ').replace(/[أإآ]/g,'ا').replace(/[ةه]/g,'ه').replace(/[يى]/g,'ي').toLowerCase();}var na=norm(a),nb=norm(b);if(na===nb)return 1.0;var ap=na.split(' ').filter(function(p){return p.length>0;}),bp=nb.split(' ').filter(function(p){return p.length>0;});var shorter=ap.length<=bp.length?ap:bp,longer=ap.length<=bp.length?bp:ap;var li=-1,mc=0,af=true;for(var k=0;k<shorter.length;k++){var f=false;for(var j=li+1;j<longer.length;j++){if(longer[j]===shorter[k]||longer[j].indexOf(shorter[k])===0||shorter[k].indexOf(longer[j])===0){li=j;mc++;f=true;break;}}if(!f){af=false;break;}}if(af&&mc===shorter.length)return 0.85+0.15*mc/longer.length;var aSet={},bSet={};ap.forEach(function(p){aSet[p]=true;});bp.forEach(function(p){bSet[p]=true;});var inter=Object.keys(aSet).filter(function(k){return bSet[k];}).length,us=Object.keys(Object.assign({},aSet,bSet)).length;return us>0?inter/us:0;}
function findInternetUserByName(fullName){try{var sh=SS.getSheetByName(SH_IUSERS);if(!sh)return{status:'not_found'};var allData=sh.getDataRange().getValues();if(allData.length<2)return{status:'not_found'};var candidates=[],THRESHOLD=0.72;for(var i=1;i<allData.length;i++){var sName=String(allData[i][0]||'').trim(),sUser=String(allData[i][1]||'').trim(),sDept=String(allData[i][2]||'').trim();if(!sName||!sUser)continue;var score=smartNameMatch(fullName,sName);if(score>=THRESHOLD)candidates.push({sheetName:sName,internetUser:sUser,dept:sDept,score:score});}if(!candidates.length)return{status:'not_found'};candidates.sort(function(a,b){return b.score-a.score;});var top=candidates[0],runner=candidates[1];if(runner&&(top.score-runner.score)<0.05&&top.score<0.95)return{status:'ambiguous',candidates:candidates.slice(0,3).map(function(c){return{name:c.sheetName,user:c.internetUser,score:Math.round(c.score*100)};})};return{status:'found',internetUser:top.internetUser,matchedName:top.sheetName,dept:top.dept,confidence:Math.round(top.score*100)};}catch(e){return{status:'not_found'};}}
function getInternetUserForEmployee(firstName,lastName,dept){var fullName=(firstName+' '+lastName).trim(),sh=SS.getSheetByName(SH_IUSERS);if(!sh)return null;var allData=sh.getDataRange().getValues().slice(1);for(var i=0;i<allData.length;i++){var sName=String(allData[i][0]||'').trim(),sUser=String(allData[i][1]||'').trim(),sDept=String(allData[i][2]||'').trim();if(!sName||!sUser)continue;if(sName===fullName)return{internetUser:sUser,confidence:100,matchedName:sName};var ns=smartNameMatch(fullName,sName);if(ns>=0.80){var dm=dept&&sDept&&(dept.includes(sDept)||sDept.includes(dept)),fs=dm?Math.min(ns+0.10,1.0):ns;if(fs>=0.80)return{internetUser:sUser,confidence:Math.round(fs*100),matchedName:sName};}}return null;}
function saveInternetUser(empId,fullName,internetUser,dept,updatedBy){var sh=SS.getSheetByName(SH_IUSERS);if(!sh){sh=SS.insertSheet(SH_IUSERS);sh.appendRow(['fullName','internetUser','dept']);}var allData=sh.getDataRange().getValues();for(var i=1;i<allData.length;i++){if(smartNameMatch(fullName,String(allData[i][0]||'').trim())>=0.9){sh.getRange(i+1,2).setValue(internetUser);return;}}sh.appendRow([fullName,internetUser,dept||'']);}
function handleInternetUsersList(user){if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');var iSh=SS.getSheetByName(SH_IUSERS),iData=iSh?iSh.getDataRange().getValues().slice(1):[];var nameMap={};iData.forEach(function(r){var n=String(r[0]||'').trim(),u=String(r[1]||'').trim(),d=String(r[2]||'').trim();if(n)nameMap[n]={internetUser:u,dept:d};});var uD=sheetData(SH_USERS),uH=uD.headers;var users=uD.rows.map(function(r){function g(c){return String(r[uH.indexOf(c)]!==undefined?r[uH.indexOf(c)]:'');}var empId=g('empId').toUpperCase(),fullName=(g('firstName')+' '+g('lastName')).trim();var info=nameMap[fullName]||{};if(!info.internetUser){var m=findInternetUserByName(fullName);if(m&&m.status==='found')info={internetUser:m.internetUser,dept:m.dept||''};}return{empId:empId,name:fullName,dept:g('dept'),internetUser:info.internetUser||'',updatedAt:''};}).filter(function(u){return u.empId;});return ok({users:users});}
function handleInternetUsersSet(data,user){if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');var empId=(data.empId||'').toString().trim().toUpperCase(),iUser=(data.internetUser||'').toString().trim().toLowerCase();if(!empId||!iUser)return fail('البيانات ناقصة');var uD=sheetData(SH_USERS),uH=uD.headers,fullName='',dept='';for(var i=0;i<uD.rows.length;i++){if(String(uD.rows[i][uH.indexOf('empId')]).toUpperCase()===empId){fullName=(String(uD.rows[i][uH.indexOf('firstName')])+' '+String(uD.rows[i][uH.indexOf('lastName')])).trim();dept=String(uD.rows[i][uH.indexOf('dept')]||'');break;}}if(!fullName)return fail('رقم البصمة '+empId+' غير موجود');saveInternetUser(empId,fullName,iUser,dept,user.firstName+' '+user.lastName);return ok({message:'تم حفظ يوزر الإنترنت: '+iUser+' للموظف '+fullName});}
function handleInternetUsersSearch(data,user){if(ROLES_IT.indexOf(user.role)<0)return fail('غير مصرح');var empId=(data.empId||'').toString().trim().toUpperCase();if(!empId)return fail('ادخل رقم البصمة');var uD=sheetData(SH_USERS),uH=uD.headers,fullName='',dept='',found=false;for(var i=0;i<uD.rows.length;i++){if(String(uD.rows[i][uH.indexOf('empId')]).toUpperCase()===empId){fullName=(String(uD.rows[i][uH.indexOf('firstName')])+' '+String(uD.rows[i][uH.indexOf('lastName')])).trim();dept=String(uD.rows[i][uH.indexOf('dept')]||'');found=true;break;}}if(!found)return fail('رقم البصمة '+empId+' غير موجود بالنظام');var iSh=SS.getSheetByName(SH_IUSERS),internetUser='',matchedName='',confidence=0;if(iSh){var iData=iSh.getDataRange().getValues().slice(1);for(var j=0;j<iData.length;j++){var sN=String(iData[j][0]||'').trim(),sU=String(iData[j][1]||'').trim();if(sN===fullName&&sU){internetUser=sU;matchedName=sN;confidence=100;break;}}if(!internetUser){var m=findInternetUserByName(fullName);if(m&&m.status==='found'){internetUser=m.internetUser;matchedName=m.matchedName;confidence=m.confidence;}}}return ok({empId:empId,fullName:fullName,dept:dept,internetUser:internetUser||null,matchedName:matchedName||null,confidence:confidence});}

// ══════ GUIDELINES ══════
function handleGuidelinesList(){var sh=SS.getSheetByName(SH_GUIDELINES);if(!sh)return ok({guidelines:[]});var allData=sh.getDataRange().getValues();if(allData.length<1)return ok({guidelines:[]});var headers=allData[0].map(function(h){return String(h).trim();}),list=[];for(var i=1;i<allData.length;i++){var r=allData[i];function gf(keys){for(var k=0;k<keys.length;k++){var idx=headers.indexOf(keys[k]);if(idx>=0&&r[idx]!==undefined&&String(r[idx]).trim())return String(r[idx]).trim();}return '';}var id=gf(['id','رقم','ID'])||String(i),title=gf(['title','عنوان','العنوان','التوجيه']),content=gf(['content','محتوى','المحتوى','التفاصيل','الوصف']);if(!title&&!content)continue;list.push({id:id,title:title,content:content,icon:gf(['icon','أيقونة','ايقونة'])||'⚠️',priority:gf(['priority','أهمية','الأهمية','الأولوية'])||'مهم',createdAt:gf(['createdAt','تاريخ','التاريخ'])});}return ok({guidelines:list});}
function handleGuidelinesAdd(data,user){if(!data.title||!data.content)return fail('العنوان والمحتوى مطلوبان');if(['it','it_manager','admin'].indexOf(user.role)<0)return fail('غير مصرح');var sh=SS.getSheetByName(SH_GUIDELINES);if(!sh)return fail('شيت التوجيهات غير موجود');var headers=sh.getDataRange().getValues()[0].map(function(h){return String(h).trim();}),row=new Array(headers.length).fill('');function sc(keys,val){for(var k=0;k<keys.length;k++){var i=headers.indexOf(keys[k]);if(i>=0){row[i]=val;return;}}}sc(['id','رقم'],genId('GD'));sc(['title','عنوان','العنوان','التوجيه'],data.title);sc(['content','محتوى','المحتوى','التفاصيل'],data.content);sc(['icon','أيقونة'],data.icon||'⚠️');sc(['priority','أهمية','الأهمية'],data.priority||'مهم');sc(['createdAt','تاريخ','التاريخ'],now());sh.appendRow(row);return ok({message:'تمت إضافة التوجيه'});}
function handleGuidelinesUpdate(data,user){if(['it','it_manager','admin'].indexOf(user.role)<0)return fail('غير مصرح');var sh=SS.getSheetByName(SH_GUIDELINES);if(!sh)return fail('غير موجود');var d=sh.getDataRange().getValues(),hi=d[0];function fi(keys){for(var k=0;k<keys.length;k++){var i=hi.indexOf(keys[k]);if(i>=0)return i;}return -1;}var idCol=fi(['id','رقم','ID']);for(var i=1;i<d.length;i++){if(String(d[i][idCol])===data.id){var tc=fi(['title','عنوان','العنوان','التوجيه']),cc=fi(['content','محتوى','المحتوى']),ic=fi(['icon','أيقونة']),pc=fi(['priority','أهمية']);if(data.title&&tc>=0)sh.getRange(i+1,tc+1).setValue(data.title);if(data.content&&cc>=0)sh.getRange(i+1,cc+1).setValue(data.content);if(data.icon&&ic>=0)sh.getRange(i+1,ic+1).setValue(data.icon);if(data.priority&&pc>=0)sh.getRange(i+1,pc+1).setValue(data.priority);return ok({message:'تم التعديل'});}}return fail('غير موجود');}
function handleGuidelinesDelete(data,user){if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');var sh=SS.getSheetByName(SH_GUIDELINES);if(!sh)return fail('غير موجود');var d=sh.getDataRange().getValues(),hi=d[0];function fi(keys){for(var k=0;k<keys.length;k++){var i=hi.indexOf(keys[k]);if(i>=0)return i;}return -1;}var idCol=fi(['id','رقم','ID']);for(var i=1;i<d.length;i++){if(String(d[i][idCol])===data.id){sh.deleteRow(i+1);return ok({message:'تم الحذف'});}}return fail('غير موجود');}

// ══════ MANUAL ══════
function handleManualList(data){var sh=SS.getSheetByName(SH_MANUAL);if(!sh)return ok({entries:[]});var allData=sh.getDataRange().getValues();if(allData.length<1)return ok({entries:[]});var headers=allData[0].map(function(h){return String(h).trim();}),device=(data.device||'').toLowerCase(),entries=[];for(var i=1;i<allData.length;i++){var r=allData[i];function gm(keys){for(var k=0;k<keys.length;k++){var idx=headers.indexOf(keys[k]);if(idx>=0&&r[idx]!==undefined)return String(r[idx]||'').trim();}return '';}var id=gm(['id','رقم'])||String(i),dev=gm(['device','جهاز','نوع_الجهاز','نوع الجهاز']),title=gm(['title','عنوان','العنوان']),sr=gm(['steps','خطوات','الخطوات']);var steps=[];try{steps=JSON.parse(sr);}catch(e){steps=sr?sr.split('\n').filter(Boolean):[];}if(!title&&!steps.length)continue;if(device&&dev.toLowerCase()!==device)continue;entries.push({id:id,device:dev,problemType:gm(['problemType','مشكلة','نوع_المشكلة']),title:title,description:gm(['description','وصف','الوصف']),icon:gm(['icon','أيقونة'])||'💻',steps:steps});}return ok({entries:entries});}
function handleManualAdd(data,user){if(['it','it_manager','admin'].indexOf(user.role)<0)return fail('غير مصرح');if(!data.title||!data.device||!data.steps||!data.steps.length)return fail('البيانات ناقصة');var sh=SS.getSheetByName(SH_MANUAL);if(!sh)return fail('شيت كتيب التعليمات غير موجود');var headers=sh.getDataRange().getValues()[0].map(function(h){return String(h).trim();}),row=new Array(headers.length).fill('');function sc(keys,val){for(var k=0;k<keys.length;k++){var i=headers.indexOf(keys[k]);if(i>=0){row[i]=val;return;}}}sc(['id','رقم'],genId('MN'));sc(['device','جهاز','نوع_الجهاز'],data.device);sc(['title','عنوان','العنوان'],data.title);sc(['steps','خطوات','الخطوات'],JSON.stringify(data.steps));sc(['problemType','مشكلة'],data.problemType||'');sc(['description','وصف'],data.description||'');sc(['icon','أيقونة'],data.icon||'💻');sc(['createdAt','تاريخ'],now());sh.appendRow(row);return ok({message:'تمت إضافة الدليل'});}
function handleManualUpdate(data,user){if(['it','it_manager','admin'].indexOf(user.role)<0)return fail('غير مصرح');var sh=SS.getSheetByName(SH_MANUAL);if(!sh)return fail('غير موجود');var d=sh.getDataRange().getValues(),hi=d[0];function fi(keys){for(var k=0;k<keys.length;k++){var i=hi.indexOf(keys[k]);if(i>=0)return i;}return -1;}var idCol=fi(['id','رقم']);for(var i=1;i<d.length;i++){if(String(d[i][idCol])===data.id){function sv(keys,val){if(val===undefined||val===null)return;var c=fi(keys);if(c>=0)sh.getRange(i+1,c+1).setValue(val);}sv(['title','عنوان'],data.title);sv(['device','جهاز'],data.device);sv(['problemType','مشكلة'],data.problemType);sv(['description','وصف'],data.description);sv(['icon','أيقونة'],data.icon);if(data.steps){var sc2=fi(['steps','خطوات']);if(sc2>=0)sh.getRange(i+1,sc2+1).setValue(JSON.stringify(data.steps));}return ok({message:'تم التعديل'});}}return fail('غير موجود');}
function handleManualDelete(data,user){if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');var sh=SS.getSheetByName(SH_MANUAL);if(!sh)return fail('غير موجود');var d=sh.getDataRange().getValues(),hi=d[0];function fi(keys){for(var k=0;k<keys.length;k++){var i=hi.indexOf(keys[k]);if(i>=0)return i;}return -1;}var idCol=fi(['id','رقم']);for(var i=1;i<d.length;i++){if(String(d[i][idCol])===data.id){sh.deleteRow(i+1);return ok({message:'تم الحذف'});}}return fail('غير موجود');}

// ══════ DEPTS ══════
function handleDeptsList(){
  var sh=getSheet(SH_DEPTS),allData=sh.getDataRange().getValues();
  if(!allData.length)return ok({depts:[],departments:[]});
  var headers=allData[0],nameCol=-1;
  // يبحث عن أي عمود فيه كلمة "قسم" أو "اسم" أو "name"
  for(var h=0;h<headers.length;h++){
    var hs=String(headers[h]).trim();
    if(hs.includes('قسم')||hs.includes('اسم')||hs.toLowerCase().includes('name')){nameCol=h;break;}
  }
  if(nameCol<0)nameCol=0;
  var depts=[];
  for(var i=1;i<allData.length;i++){var val=String(allData[i][nameCol]||'').trim();if(val)depts.push(val);}
  return ok({depts:depts,departments:depts});
}
function getDeptNameCol(headers){for(var h=0;h<headers.length;h++){var hs=String(headers[h]).trim();if(hs.includes('قسم')||hs.includes('اسم')||hs.toLowerCase().includes('name'))return h;}return 0;}
function handleDeptsAdd(data,user){
  if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');
  if(!data.name)return fail('اسم القسم مطلوب');
  var sh=getSheet(SH_DEPTS),allData=sh.getDataRange().getValues(),headers=allData[0];
  var nameCol=getDeptNameCol(headers);
  for(var i=1;i<allData.length;i++){if(String(allData[i][nameCol]).trim()===data.name.trim())return fail('القسم موجود مسبقاً');}
  var row=new Array(headers.length).fill('');row[nameCol]=data.name.trim();
  for(var h2=0;h2<headers.length;h2++){if(String(headers[h2]).includes('تاريخ')||String(headers[h2]).toLowerCase().includes('date')){row[h2]=now();break;}}
  sh.appendRow(row);return ok({message:'تمت إضافة القسم'});
}
function handleDeptsDelete(data,user){
  if(ROLES_ADMIN.indexOf(user.role)<0)return fail('غير مصرح');
  var sh=getSheet(SH_DEPTS),allData=sh.getDataRange().getValues(),headers=allData[0];
  var nameCol=getDeptNameCol(headers);
  var delName=data.name||data.id||'';
  for(var i=1;i<allData.length;i++){if(String(allData[i][nameCol]).trim()===delName.trim()){sh.deleteRow(i+1);return ok({message:'تم حذف القسم'});}}
  return fail('القسم غير موجود');
}

// ══════ AI CHAT ══════
var AI_MODEL='claude-sonnet-4-20250514';
function getITKnowledge(){try{var sh=getSheet(SH_TICKETS),allData=sh.getDataRange().getValues(),byType={};allData.forEach(function(r){var tid=String(r[COL_TICKETS.id]||'').trim();if(!tid||tid.indexOf('TKT')<0)return;if(String(r[COL_TICKETS.status]||'')!=='تم حل البلاغ')return;var desc=String(r[COL_TICKETS.description]||'').trim(),prob=String(r[COL_TICKETS.problemType]||'').trim(),dev=String(r[COL_TICKETS.deviceId]||'').trim();if(!prob||!desc)return;var key=prob+(dev?'|'+dev:'');if(!byType[key])byType[key]={count:0,desc:[],type:prob,device:dev};byType[key].count++;if(byType[key].desc.length<2)byType[key].desc.push(desc.slice(0,100));});var lines=Object.values(byType).sort(function(a,b){return b.count-a.count;}).slice(0,10).map(function(g){return '['+g.type+(g.device?' | '+g.device:'')+'] حدث '+g.count+' مرة';});return lines.length?lines.join('\n'):'';}catch(e){return '';}}
function handleAIChat(data,user){
  var messages=data.messages||[],systemOverride=data.systemOverride||'';
  var internetUser='';try{var iuResult=getInternetUserForEmployee(user.firstName,user.lastName,user.dept);if(iuResult)internetUser=iuResult.internetUser;}catch(e){}
  var knowledge='';try{knowledge=getITKnowledge();}catch(e){}
  var systemPrompt=systemOverride||buildAISystemPrompt(user,internetUser,knowledge);
  var options={method:'post',contentType:'application/json',headers:{'x-api-key':getAIKey(),'anthropic-version':'2023-06-01'},payload:JSON.stringify({model:AI_MODEL,max_tokens:800,system:systemPrompt,messages:messages.slice(-10)}),muteHttpExceptions:true};
  try{var resp=UrlFetchApp.fetch('https://api.anthropic.com/v1/messages',options);var body=JSON.parse(resp.getContentText());if(body.error)return fail('خطأ من الذكاء الاصطناعي: '+body.error.message);var reply=body.content&&body.content[0]?body.content[0].text:'';var submitTicket=reply.indexOf('[TICKET]')>=0;reply=reply.replace('[TICKET]','').trim();return ok({reply:reply,submitTicket:submitTicket});}catch(e){return fail('تعذر الاتصال بالذكاء الاصطناعي: '+e.message);}
}
function buildAISystemPrompt(user,internetUser,knowledge){var iuSection=internetUser?'يوزر الإنترنت للموظف الحالي: '+internetUser+' — إذا سأل عن يوزره أعطه هذا مباشرة.':'الموظف الحالي ليس عنده يوزر إنترنت مسجل — إذا سأل عنه قله يتواصل مع فريق IT.';var knowledgeSection=knowledge?'## أكثر المشاكل تكراراً وحلولها من سجل IT:\n'+knowledge:'## لا توجد حلول مسجلة بعد — استخدم خبرتك التقنية العامة.';return 'أنت مساعد تقني ذكي لشركة عراقية. اسم الموظف: '+user.firstName+' '+user.lastName+' | القسم: '+user.dept+'\n\n'+iuSection+'\n\n'+knowledgeSection+'\n\n## قواعد صارمة:\n- لا تذكر أي IP أو كلمة مرور أو معلومات شبكية داخلية\n- أجب باللهجة العراقية البسيطة\n- كل رد: تشخيص مختصر + خطوة عملية واحدة\n- بعد 5 خطوات فاشلة: اقترح بلاغ بكلمة [TICKET]';}
function getAIKey(){var key=PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');if(!key)throw new Error('مفتاح Anthropic API غير موجود — أضفه في Script Properties');return key;}

// ══════ NOTIFICATIONS LIST (للفرونت) ══════
function handleNotificationsList(user) {
  if(ROLES_IT.indexOf(user.role)<0 && ROLES_MGR.indexOf(user.role)<0) return fail('غير مصرح');
  try {
    var sh=getSheet(SH_TICKETS), allData=sh.getDataRange().getValues();
    var notifs=[];
    for(var i=allData.length-1;i>=1;i--){
      var r=allData[i];
      var tid=String(r[COL_TICKETS.id]||'').trim();
      if(!tid||tid.indexOf('TKT')<0) continue;
      var status=String(r[COL_TICKETS.status]||'').trim();
      if(status==='تم حل البلاغ') continue;
      notifs.push({
        id: tid,
        subject: String(r[COL_TICKETS.problemType]||''),
        problem: String(r[COL_TICKETS.problemType]||''),
        userName: String(r[COL_TICKETS.requesterName]||''),
        department: String(r[COL_TICKETS.requesterDept]||''),
        priority: String(r[COL_TICKETS.priority]||'متوسطة'),
        status: status,
        assignedName: String(r[COL_TICKETS.assignedName]||''),
        claimedBy: String(r[COL_TICKETS.assignedId]||''),
        claimerName: String(r[COL_TICKETS.assignedName]||''),
        read: !!(r[COL_TICKETS.assignedName])
      });
      if(notifs.length>=20) break;
    }
    return ok({notifications:notifs});
  } catch(e) { return ok({notifications:[]}); }
}

function handleBroadcastClaim(data, user) {
  try {
    var tSh=getSheet(SH_TICKETS), tAll=tSh.getDataRange().getValues();
    for(var k=0;k<tAll.length;k++){
      if(String(tAll[k][COL_TICKETS.id]).trim()===data.ticketId){
        notifyTeamAboutClaim(data.ticketId, user, tAll[k]);
        break;
      }
    }
    return ok({message:'تم الإرسال'});
  } catch(e) { return ok({message:'تم'}); }
}

function handleBroadcastAssign(data, user) {
  try {
    var tSh=getSheet(SH_TICKETS), tAll=tSh.getDataRange().getValues();
    for(var k=0;k<tAll.length;k++){
      if(String(tAll[k][COL_TICKETS.id]).trim()===data.ticketId){
        var assignedName=String(tAll[k][COL_TICKETS.assignedName]||'');
        var assigneeData=assignedName?{firstName:assignedName,lastName:'',empId:String(tAll[k][COL_TICKETS.assignedId]||''),email:''}:null;
        if(assigneeData) notifyTeamAboutAssign(data.ticketId, user, assigneeData, tAll[k]);
        break;
      }
    }
    return ok({message:'تم الإرسال'});
  } catch(e) { return ok({message:'تم'}); }
}

// ══════ TECH STATS ══════
function handleTechStats(user){
  if(user.role!=='tech')return fail('غير مصرح');
  var sh=SS.getSheetByName(SH_DEVICES);if(!sh)return ok({stats:{total:0,fixed:0,pending:0,delivered:0},recent:[]});
  var allData=sh.getDataRange().getValues();if(allData.length<2)return ok({stats:{total:0,fixed:0,pending:0,delivered:0},recent:[]});
  var hi=allData[0];function ci(k){return hi.indexOf(k);}
  var stats={total:0,fixed:0,pending:0,delivered:0},myDevices=[];
  for(var i=1;i<allData.length;i++){var r=allData[i];function gv(k){var c=ci(k);return c>=0?String(r[c]||'').trim():'';}var techId=gv('techId');if(techId!==user.id&&techId!==user.empId)continue;stats.total++;var status=gv('status');if(status==='مُسلَّم')stats.delivered++;else if(status==='أُنجز — بانتظار IT'||status==='جاهز للتسليم')stats.fixed++;else stats.pending++;myDevices.push({devId:gv('devId'),deviceType:gv('deviceType'),deviceDesc:gv('deviceDesc'),ownerName:gv('ownerName'),ownerDept:gv('ownerDept'),status:status,techResult:gv('techResult'),techResultLabel:gv('techResultLabel'),workDone:gv('workDone'),techDoneAt:gv('techDoneAt'),checkinAt:gv('checkinAt')});}
  return ok({stats:stats,recent:myDevices.slice(-10).reverse()});
}
