// ══════════════════════════════════════════════════════════════════════
//  ADDITIONS TO Code.gs — Manager Email Assign Flow
//  أضف هذا الكود في نهاية Code.gs الموجود
// ══════════════════════════════════════════════════════════════════════


// ── 1. في دالة doGet الموجودة، أضف هذين الـ if بعد سطر handleEmailSolve ──
//
//   if (action === 'assignForm' && ticketId && userId) {
//     return handleEmailAssignForm(ticketId, userId, token);
//   }
//   if (action === 'doAssign' && ticketId && userId) {
//     var assignTo = e && e.parameter ? e.parameter.assignTo : '';
//     return handleEmailDoAssign(ticketId, userId, token, assignTo);
//   }


// ── 2. دالة صفحة التعيين — المراقب يختار موظف IT ──
function handleEmailAssignForm(ticketId, mgrEmpId, token) {
  var html = '';
  try {
    if (token !== generateEmailToken(ticketId, mgrEmpId)) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ رابط غير صحيح', 'هذا الرابط غير صحيح أو منتهي الصلاحية.', '#ef4444'));
    }
    var mgrData = getEmployeeByEmpId(mgrEmpId);
    if (!mgrData || ROLES_MGR.indexOf(mgrData.role) < 0) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ غير مصرح', 'ليس لديك صلاحية تعيين البلاغات.', '#ef4444'));
    }

    // جلب بيانات البلاغ
    var tSh = getSheet(SH_TICKETS), tAll = tSh.getDataRange().getValues();
    var ticketRow = null;
    for (var i = 1; i < tAll.length; i++) {
      if (String(tAll[i][COL_TICKETS.id]).trim() === ticketId) { ticketRow = tAll[i]; break; }
    }
    if (!ticketRow) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ البلاغ غير موجود', 'رقم البلاغ: ' + ticketId, '#ef4444'));
    }

    var assignedName = String(ticketRow[COL_TICKETS.assignedName] || '').trim();
    var currentStatus = String(ticketRow[COL_TICKETS.status] || '').trim();
    if (assignedName && currentStatus !== 'جديدة') {
      return HtmlService.createHtmlOutput(buildResultPage('⚠️ تم التعيين مسبقاً', 'هذا البلاغ تم تعيينه إلى: <strong>' + assignedName + '</strong>', '#f59e0b'));
    }

    // جلب موظفي IT النشطين
    var uSh = getSheet(SH_USERS), uAll = uSh.getDataRange().getValues(), uH = uAll[0];
    var idCol = uH.indexOf(COL_USERS.id), empIdCol = uH.indexOf(COL_USERS.empId);
    var fnCol = uH.indexOf(COL_USERS.firstName), lnCol = uH.indexOf(COL_USERS.lastName);
    var roleCol = uH.indexOf(COL_USERS.role), activeCol = uH.indexOf(COL_USERS.active);
    var itStaff = [];
    for (var j = 1; j < uAll.length; j++) {
      var r = uAll[j];
      var role = String(r[roleCol] || '').trim().toLowerCase();
      var active = r[activeCol];
      if (ROLES_IT.indexOf(role) >= 0 && (active === true || String(active).toLowerCase() === 'true')) {
        itStaff.push({
          id: String(r[idCol] || '').trim(),
          empId: String(r[empIdCol] || '').trim(),
          name: (String(r[fnCol] || '') + ' ' + String(r[lnCol] || '')).trim(),
          role: role
        });
      }
    }

    var ticketInfo = {
      id: String(ticketRow[COL_TICKETS.id] || ''),
      problemType: String(ticketRow[COL_TICKETS.problemType] || ''),
      requesterName: String(ticketRow[COL_TICKETS.requesterName] || ''),
      requesterDept: String(ticketRow[COL_TICKETS.requesterDept] || ''),
      priority: String(ticketRow[COL_TICKETS.priority] || ''),
      status: currentStatus
    };

    return buildAssignFormPage(ticketInfo, mgrData, itStaff, token);
  } catch(e) {
    return HtmlService.createHtmlOutput(buildResultPage('❌ خطأ', e.message, '#ef4444'));
  }
}


// ── 3. دالة تنفيذ التعيين من صفحة الإيميل ──
function handleEmailDoAssign(ticketId, mgrEmpId, token, assignToEmpId) {
  try {
    if (token !== generateEmailToken(ticketId, mgrEmpId)) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ رابط غير صحيح', 'هذا الرابط غير صحيح أو منتهي الصلاحية.', '#ef4444'));
    }
    var mgrUser = getEmployeeByEmpId(mgrEmpId);
    if (!mgrUser || ROLES_MGR.indexOf(mgrUser.role) < 0) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ غير مصرح', 'ليس لديك صلاحية لتعيين البلاغات.', '#ef4444'));
    }
    if (!assignToEmpId) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ لم يتم الاختيار', 'يرجى اختيار موظف IT.', '#ef4444'));
    }

    var assigneeData = getEmployeeByEmpId(assignToEmpId);
    if (!assigneeData) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ الموظف غير موجود', 'رقم البصمة: ' + assignToEmpId, '#ef4444'));
    }

    var fullName = assigneeData.firstName + ' ' + assigneeData.lastName;
    var result = assignTicketDirect(ticketId, assigneeData.empId, fullName, mgrUser);
    if (!result.success) {
      return HtmlService.createHtmlOutput(buildResultPage('❌ خطأ', result.message, '#ef4444'));
    }

    // إشعار للموظف المعين
    try {
      if (assigneeData.email) {
        var tSh = getSheet(SH_TICKETS), tAll = tSh.getDataRange().getValues();
        for (var k = 0; k < tAll.length; k++) {
          if (String(tAll[k][COL_TICKETS.id]).trim() === ticketId) {
            notifyAssignee(ticketId, assigneeData, tAll[k]); break;
          }
        }
      }
    } catch(e2) { Logger.log('notifyAssignee error: ' + e2.toString()); }

    // إشعار لباقي فريق IT
    try {
      var tSh2 = getSheet(SH_TICKETS), tAll2 = tSh2.getDataRange().getValues();
      for (var m = 0; m < tAll2.length; m++) {
        if (String(tAll2[m][COL_TICKETS.id]).trim() === ticketId) {
          notifyTeamAboutAssign(ticketId, mgrUser, assigneeData, tAll2[m]); break;
        }
      }
    } catch(e3) { Logger.log('notifyTeamAboutAssign error: ' + e3.toString()); }

    return HtmlService.createHtmlOutput(buildResultPage(
      '✅ تم التعيين بنجاح',
      'تم تعيين البلاغ <strong>' + ticketId + '</strong> إلى <strong>' + fullName + '</strong><br>تم إرسال إشعار للموظف وبقية الفريق.',
      '#059669', ticketId
    ));
  } catch(e) {
    return HtmlService.createHtmlOutput(buildResultPage('❌ خطأ', e.message, '#ef4444'));
  }
}


// ── 4. إشعار فريق IT عند التعيين من المراقب ──
function notifyTeamAboutAssign(ticketId, assignedBy, assignee, ticketRow) {
  try {
    var uD = sheetData(SH_USERS), uH = uD.headers;
    var problemType = String(ticketRow[COL_TICKETS.problemType] || '');
    var requesterName = String(ticketRow[COL_TICKETS.requesterName] || '');
    uD.rows.forEach(function(r) {
      var role   = String(r[uH.indexOf('role')] || '');
      var email  = String(r[uH.indexOf('notifyEmail')] || '').trim();
      var empId  = String(r[uH.indexOf('empId')] || '').trim();
      var active = r[uH.indexOf('active')];
      if (ROLES_STAFF.indexOf(role) < 0) return;
      if (empId === assignee.empId) return; // لا ترسل للموظف المعين (راح يوصله إيميل منفصل)
      if (!email || !(active === true || String(active).toLowerCase() === 'true')) return;
      var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body dir="rtl" style="font-family:Arial;padding:20px;background:#f8fafc">'
        + '<div style="max-width:500px;margin:0 auto;background:#fff;border-radius:12px;padding:24px;box-shadow:0 4px 20px rgba(0,0,0,.1)">'
        + '<div style="background:linear-gradient(135deg,#f59e0b,#d97706);padding:16px 20px;border-radius:8px;margin-bottom:16px">'
        + '<h3 style="margin:0;color:#fff;font-size:15px">📌 تم تعيين بلاغ</h3>'
        + '</div>'
        + '<p style="font-size:13px;color:#374151">قام <strong>' + assignedBy.firstName + ' ' + assignedBy.lastName + '</strong> بتعيين البلاغ <strong>' + ticketId + '</strong> إلى <strong>' + assignee.firstName + ' ' + assignee.lastName + '</strong></p>'
        + '<div style="background:#f8fafc;border-radius:8px;padding:12px;margin:12px 0;font-size:13px">'
        + '<div><strong>رقم البلاغ:</strong> ' + ticketId + '</div>'
        + '<div><strong>نوع المشكلة:</strong> ' + problemType + '</div>'
        + '<div><strong>المُبلِّغ:</strong> ' + requesterName + '</div>'
        + '</div>'
        + '<a href="' + SYSTEM_URL + '" style="background:#6366f1;color:#fff;padding:10px 20px;border-radius:8px;text-decoration:none;font-size:13px;display:inline-block">فتح النظام</a>'
        + '</div></body></html>';
      try { MailApp.sendEmail({ to: email, subject: '📌 تم تعيين البلاغ [' + ticketId + '] إلى ' + assignee.firstName, htmlBody: html }); } catch(e2) {}
    });
  } catch(e) { Logger.log('notifyTeamAboutAssign: ' + e.toString()); }
}


// ── 5. صفحة اختيار موظف IT (HTML جميل) ──
function buildAssignFormPage(ticketInfo, mgrData, itStaff, token) {
  var ROLE_LABELS = { it: 'موظف IT', it_manager: 'مدير IT', admin: 'مدير النظام' };
  var staffCards = itStaff.length ? itStaff.map(function(s) {
    var initials = s.name ? s.name.charAt(0) : '؟';
    return '<div class="staff-card" onclick="pick(this,\'' + s.empId + '\')">'
      + '<div class="av">' + initials + '</div>'
      + '<div class="sn">' + s.name + '</div>'
      + '<div class="sr">' + (ROLE_LABELS[s.role] || 'موظف IT') + '</div>'
      + '</div>';
  }).join('') : '<p style="text-align:center;color:#f87171;grid-column:1/-1">لا يوجد موظفو IT نشطون</p>';

  var priorityColor = ticketInfo.priority === 'عاجلة' ? '#ef4444'
    : ticketInfo.priority === 'عالية' ? '#f59e0b'
    : ticketInfo.priority === 'متوسطة' ? '#3b82f6' : '#10b981';

  var html = '<!DOCTYPE html><html dir="rtl" lang="ar"><head>'
    + '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>تعيين البلاغ ' + ticketInfo.id + '</title>'
    + '<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">'
    + '<style>'
    + '*{box-sizing:border-box;margin:0;padding:0}'
    + 'body{font-family:Cairo,Tahoma,Arial,sans-serif;background:#0f172a;color:#e2e8f0;min-height:100vh;padding:24px 16px}'
    + '.wrap{max-width:680px;margin:0 auto}'
    + '.top{text-align:center;margin-bottom:28px}'
    + '.top h1{font-size:24px;font-weight:900;color:#f8fafc}'
    + '.top p{font-size:13px;color:#64748b;margin-top:6px}'
    + '.tkt{background:#1e293b;border-radius:14px;padding:20px;margin-bottom:24px;border:1px solid #334155}'
    + '.tkt h2{font-size:14px;font-weight:700;color:#8b5cf6;margin-bottom:14px;letter-spacing:.5px}'
    + '.tkt-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px}'
    + '.tkt-item label{display:block;font-size:11px;color:#64748b;margin-bottom:3px}'
    + '.tkt-item span{font-size:14px;font-weight:600;color:#f1f5f9}'
    + '.pri-badge{display:inline-block;padding:2px 10px;border-radius:99px;font-size:12px;font-weight:700;background:' + priorityColor + '20;color:' + priorityColor + '}'
    + '.sec-title{font-size:14px;font-weight:700;color:#8b5cf6;margin-bottom:14px;letter-spacing:.5px}'
    + '.grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(140px,1fr));gap:12px;margin-bottom:28px}'
    + '.staff-card{background:#1e293b;border:2px solid #334155;border-radius:12px;padding:18px 12px;text-align:center;cursor:pointer;transition:all .18s}'
    + '.staff-card:hover{border-color:#6366f1;background:#1e2040;transform:translateY(-2px)}'
    + '.staff-card.sel{border-color:#6366f1;background:#1e2040;box-shadow:0 0 0 3px rgba(99,102,241,.25)}'
    + '.av{width:52px;height:52px;border-radius:50%;background:linear-gradient(135deg,#312e81,#4338ca);color:#a5b4fc;font-size:22px;font-weight:900;display:flex;align-items:center;justify-content:center;margin:0 auto 10px}'
    + '.sn{font-size:13px;font-weight:700;color:#f1f5f9;margin-bottom:4px}'
    + '.sr{font-size:11px;color:#64748b}'
    + '.btn-wrap{text-align:center}'
    + '.btn{display:inline-block;padding:14px 48px;background:linear-gradient(135deg,#6366f1,#8b5cf6);color:#fff;border:none;border-radius:10px;font-size:16px;font-weight:700;cursor:pointer;font-family:Cairo,Tahoma,Arial,sans-serif;transition:opacity .2s}'
    + '.btn:disabled{opacity:.4;cursor:not-allowed}'
    + '.msg{color:#f87171;font-size:13px;margin-top:10px;min-height:18px;text-align:center}'
    + '@media(max-width:480px){.tkt-grid{grid-template-columns:1fr}}'
    + '</style></head><body><div class="wrap">'
    + '<div class="top"><h1>🎯 تعيين البلاغ</h1><p>مرحباً ' + mgrData.firstName + ' — اختر موظف IT لتعيين البلاغ إليه</p></div>'
    + '<div class="tkt"><h2>📋 تفاصيل البلاغ</h2><div class="tkt-grid">'
    + '<div class="tkt-item"><label>رقم البلاغ</label><span>' + ticketInfo.id + '</span></div>'
    + '<div class="tkt-item"><label>نوع المشكلة</label><span>' + ticketInfo.problemType + '</span></div>'
    + '<div class="tkt-item"><label>المُبلِّغ</label><span>' + ticketInfo.requesterName + '</span></div>'
    + '<div class="tkt-item"><label>القسم</label><span>' + ticketInfo.requesterDept + '</span></div>'
    + '<div class="tkt-item"><label>الأولوية</label><span><span class="pri-badge">' + ticketInfo.priority + '</span></span></div>'
    + '<div class="tkt-item"><label>الحالة</label><span>' + ticketInfo.status + '</span></div>'
    + '</div></div>'
    + '<div class="sec-title">👥 فريق تقنية المعلومات</div>'
    + '<div class="grid">' + staffCards + '</div>'
    + '<div class="btn-wrap">'
    + '<button class="btn" id="assignBtn" disabled onclick="doAssign()">تعيين البلاغ ✓</button>'
    + '<div class="msg" id="msg"></div>'
    + '</div>'
    + '</div>'
    + '<script>'
    + 'var sel="";'
    + 'function pick(el,empId){'
    + '  document.querySelectorAll(".staff-card").forEach(function(c){c.classList.remove("sel");});'
    + '  el.classList.add("sel");sel=empId;'
    + '  document.getElementById("assignBtn").disabled=false;'
    + '  document.getElementById("msg").textContent="";'
    + '}'
    + 'function doAssign(){'
    + '  if(!sel){document.getElementById("msg").textContent="يرجى اختيار موظف أولاً";return;}'
    + '  var btn=document.getElementById("assignBtn");btn.disabled=true;btn.textContent="جاري التعيين...";'
    + '  window.location.href="' + SYSTEM_URL + '?action=doAssign"'
    + '    +"&tid="+encodeURIComponent("' + ticketInfo.id + '")'
    + '    +"&uid="+encodeURIComponent("' + mgrData.empId + '")'
    + '    +"&token="+encodeURIComponent("' + token + '")'
    + '    +"&assignTo="+encodeURIComponent(sel);'
    + '}'
    + '<\/script>'
    + '</body></html>';

  return HtmlService.createHtmlOutput(html)
    .setTitle('تعيين البلاغ ' + ticketInfo.id)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ── 6. تعديل notifyITTeam — المراقب يوصله إيميل مختلف بزر "تعيين" ──
// احذف دالة notifyITTeam القديمة وضع هذي بدلها:
//
// function notifyITTeam(ticket){
//   try{
//     var uD=sheetData(SH_USERS),uH=uD.headers;
//     uD.rows.forEach(function(r){
//       var role   = String(r[uH.indexOf('role')]||'');
//       var email  = String(r[uH.indexOf('notifyEmail')]||'').trim();
//       var empId  = String(r[uH.indexOf('empId')]||'').trim();
//       var active = r[uH.indexOf('active')];
//       if(ROLES_STAFF.indexOf(role)<0)return;
//       if(!email||!(active===true||String(active).toLowerCase()==='true'))return;
//       var token = generateEmailToken(ticket.id, empId);
//
//       if(role === 'manager'){
//         // المراقب — زر تعيين
//         var assignUrl = SYSTEM_URL+'?action=assignForm&tid='+encodeURIComponent(ticket.id)+'&uid='+encodeURIComponent(empId)+'&token='+token;
//         try{MailApp.sendEmail({to:email,subject:'🔔 بلاغ جديد يحتاج تعيين ['+ticket.id+'] — '+ticket.problemType,htmlBody:buildManagerEmail(ticket,assignUrl)});}catch(e2){}
//       } else {
//         // موظف IT — زر استلام وحل
//         var claimUrl = SYSTEM_URL+'?action=claim&tid='+encodeURIComponent(ticket.id)+'&uid='+encodeURIComponent(empId)+'&token='+token;
//         var solveUrl = SYSTEM_URL+'?action=solve&tid='+encodeURIComponent(ticket.id)+'&uid='+encodeURIComponent(empId)+'&token='+token;
//         try{MailApp.sendEmail({to:email,subject:'🔔 بلاغ جديد ['+ticket.id+'] — '+ticket.problemType,htmlBody:buildTicketEmail(ticket,claimUrl,solveUrl)});}catch(e2){}
//       }
//     });
//   }catch(e){}
// }


// ── 7. إيميل المراقب (HTML) ──
function buildManagerEmail(ticket, assignUrl) {
  var priorityColor = ticket.priority==='عاجلة'?'#ef4444':ticket.priority==='عالية'?'#f59e0b':ticket.priority==='متوسطة'?'#3b82f6':'#10b981';
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body dir="rtl" style="margin:0;padding:0;background:#f8fafc;font-family:Arial,sans-serif">'
    + '<div style="max-width:600px;margin:20px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,0,.1)">'
    + '<div style="background:linear-gradient(135deg,#f59e0b,#d97706);padding:24px 28px">'
    + '<h2 style="margin:0;color:#fff;font-size:18px">🔔 بلاغ جديد يحتاج تعيين</h2>'
    + '<p style="margin:6px 0 0;color:rgba(255,255,255,.8);font-size:13px">رقم البلاغ: ' + ticket.id + '</p>'
    + '</div>'
    + '<div style="padding:24px 28px">'
    + '<table style="width:100%;border-collapse:collapse;margin-bottom:20px">'
    + '<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px;width:35%">نوع المشكلة</td><td style="padding:9px 14px;font-size:13px">' + ticket.problemType + '</td></tr>'
    + '<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">المُبلِّغ</td><td style="padding:9px 14px;font-size:13px">' + ticket.requesterName + '</td></tr>'
    + '<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">القسم</td><td style="padding:9px 14px;font-size:13px">' + ticket.requesterDept + '</td></tr>'
    + '<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">الأولوية</td><td style="padding:9px 14px"><span style="background:' + priorityColor + '20;color:' + priorityColor + ';padding:3px 10px;border-radius:12px;font-size:12px;font-weight:bold">' + ticket.priority + '</span></td></tr>'
    + '<tr><td style="padding:9px 14px;background:#f8fafc;font-weight:bold;font-size:13px">الوصف</td><td style="padding:9px 14px;font-size:13px">' + ticket.description + '</td></tr>'
    + '</table>'
    + '<div style="display:flex;gap:10px;flex-wrap:wrap">'
    + '<a href="' + assignUrl + '" style="background:linear-gradient(135deg,#f59e0b,#d97706);color:#fff;padding:13px 28px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block">👥 تعيين البلاغ لموظف IT</a>'
    + '<a href="' + SYSTEM_URL + '" style="background:#f8fafc;color:#374151;padding:13px 22px;border-radius:8px;text-decoration:none;font-size:14px;font-weight:bold;display:inline-block;border:2px solid #e2e8f0">🔍 فتح النظام</a>'
    + '</div>'
    + '<p style="font-size:11px;color:#94a3b8;margin-top:12px">اضغط "تعيين البلاغ" لاختيار موظف IT مباشرة من الإيميل بدون تسجيل دخول</p>'
    + '</div>'
    + '<div style="padding:14px 28px;background:#f8fafc;border-top:1px solid #e2e8f0;text-align:center;font-size:11px;color:#94a3b8">IT Help Desk — نظام الدعم التقني</div>'
    + '</div></body></html>';
}

