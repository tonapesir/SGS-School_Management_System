var MAIN_HEADERS = [
  "Timestamp","Acad Year","Reg No","Student ID","Book No","Aadhar",
  "Class","Division","PEN","Roll No","Full Name",
  "Mother Name","Gender","Religion","Caste","Sub-caste","DOB","DOB Words",
  "Nationality","Mother Tongue","Birth Place",
  "Previous School","Admission Date","Admission Class",
  "Contact","Address","Photo URL","WhatsApp Mobile No","Alternate Mobile No",
  "Category","Minority"
];
var LC_HEADERS = [
  "Timestamp","Serial No","Student ID","Full Name","Gender",
  "DOB","DOB Words","Class","Religion","Caste","Nationality","Mother Tongue",
  "Admission Date","Admission Class","Leave Date","Leave Class","Class Start Date",
  "LC No","LC Date","Conduct","Progress","Stxt75","Stxt76","Remarks",
  "Medium","Next School","LC Count","Fee Status","Stxt60","Stxt61",
  "Stxt62","Stxt63","Stxt67","Stxt68","Stxt69","Stxt70","Nationality Type"
];
var BF_HEADERS = [
  "Timestamp","Serial No","Student ID","Full Name",
  "DOB","DOB Words","Class","Tukdi","Religion","Caste","Nationality",
  "Acad Year","BF Date","Purpose","Admission Date","Admission Class",
  "Remarks","Address","Contact"
];
var AT_HEADERS = [
  "Timestamp","Serial No","Student ID","Full Name","Gender",
  "DOB","Class","Tukdi","Acad Year","AT Date","Purpose",
  "Total Days","Present Days","Percentage","From Date","To Date","Remarks"
];
var PHOTO_FOLDER_ID = "1urIudSUS7U0ClMTjvp__uZ0ZgTQ7GrbU";

// ===== CLASS TEACHER MODULE (V19.32) — sheet headers =====
var DAILY_ATT_HEADERS = ["Timestamp","Date","Class","Division","RegNo","StudentId","FullName","Reason","MarkedBy"];
var NOTICE_HEADERS   = ["Timestamp","Date","Title","Message","PostedBy","TargetClass"];
var DIARY_HEADERS    = ["Timestamp","Date","RegNo","StudentId","FullName","Class","Division","Type","Remark","EnteredBy"];
var TRANSPORT_HEADERS= ["Timestamp","RegNo","StudentId","FullName","Class","Division","NativeVillage","TravelMode","TransportContact","UpdatedBy"];
// ===== FEES (V19.33) — प्रत्येक जमा installment ची स्वतंत्र row, पण त्या क्षणापर्यंतची cumulative FeePaid/PendingFee सोबत =====
var FEES_HEADERS     = ["Timestamp","RegNo","StudentId","FullName","Class","Division","AcYear","TotalFee","FeePaid","PendingFee","UpdatedBy"];
var TOTAL_FEE_PER_STUDENT = 1000; // उदा. एकूण फी — गरजेनुसार येथे बदला
var CONTACTS_HEADERS = ["Timestamp","RegNo","StudentId","FullName","Class","Division","WhatsAppMobile","OtherMobile","UpdatedBy"];
var CATEGORY_LIST_SRV = ["SC","ST","VJA","NT B","NT C","NT D","SEBC","SBC","OBC","Gen"];

function getOrCreateSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  } else if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
  } else {
    var existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    for (var i = 0; i < headers.length; i++) {
      if (!existing[i]) sh.getRange(1, i + 1).setValue(headers[i]);
    }
  }
  return sh;
}

function findRowByKey(sh, col, key) {
  key = (key || "").toString().trim().toLowerCase();
  if (!key || !sh || sh.getLastRow() < 2) return 0;
  var vals = sh.getRange(2, col, sh.getLastRow()-1, 1).getValues();
  for (var i=0; i<vals.length; i++) {
    if ((vals[i][0] || "").toString().trim().toLowerCase() === key) return i + 2;
  }
  return 0;
}

function writeRow(sh, rowNum, row) {
  if (rowNum) {
    sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
    return {rowIndex: rowNum, mode: "updated"};
  }
  sh.appendRow(row);
  return {rowIndex: sh.getLastRow(), mode: "created"};
}

function doPost(e) {
  var p = e.parameter || {};
  if (p.action === "uploadPhoto") {
    var photoResult = handlePhotoUpload(p);
    return ContentService.createTextOutput(JSON.stringify(photoResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return handleAction(p, ""); // "" = JSON response — serial परत frontend ला येतो
}

function doGet(e) {
  var p = e.parameter || {};
  var cb = p.callback || "";
  if (!p.action && (p.regNo || p.studentId || p.firstName)) p.action = "upsert";
  if (p.action === "verify") {
    return doVerifyPage(p);
  }
  if (p.action === "ping") {
    try {
      var ssName = SpreadsheetApp.getActiveSpreadsheet().getName();
      return wrap(cb, {status:"ok", message:"Connected! Sheet: " + ssName});
    } catch(ex) {
      return wrap(cb, {status:"error", message:"Script is deployed but NOT BOUND to a Sheet! Re-deploy from inside the Sheet."});
    }
  }
  if (p.action === "getAll") {
    return doGetAllAction(p, cb);
  }
  if (p.action === "testPhotoFolder") {
    return wrap(cb, testPhotoFolderAccess());
  }
  if (p.action === "photoChunkStart") {
    return wrap(cb, startPhotoChunkUpload(p));
  }
  if (p.action === "photoChunk") {
    return wrap(cb, savePhotoChunk(p));
  }
  if (p.action === "photoChunkFinish") {
    return wrap(cb, finishPhotoChunkUpload(p));
  }
  if (p.action === "getPhotoUrl" && p.regNo) {
    return wrap(cb, getPhotoUrlByRegNo(p.regNo));
  }
  if (p.action === "search" && p.q) {
    return doSearchAction(p, cb);
  }
  if (p.action === "getDashboardStats") {
    return doGetDashboardStats(p, cb);
  }
  if (p.action === "getStudentHistory") {
    return doGetStudentHistory(p, cb);
  }
  if (p.action === "getAllCertificates") {
    return doGetAllCertificates(p, cb);
  }
  if (p.action === "getUsers") {
    return doGetUsers(p, cb);
  }
  if (p.action === "saveUser") {
    return doSaveUser(p, cb);
  }
  if (p.action === "deleteUser") {
    return doDeleteUser(p, cb);
  }
  if (p.action === "changePassword") {
    return doChangePassword(p, cb);
  }
  if (p.action === "getAnalyticsData") {
    return doGetAnalyticsData(p, cb);
  }
  if (p.action === "verify") {
    return doVerifyPage(p);
  }
  // ===== CLASS TEACHER MODULE (V19.32) =====
  if (p.action === "getClassStudents") {
    return doGetClassStudents(p, cb);
  }
  if (p.action === "getClassList") {
    return doGetClassList(p, cb);
  }
  if (p.action === "getTodayBirthdays") {
    return doGetTodayBirthdays(p, cb);
  }
  if (p.action === "getAttendance") {
    return doGetAttendance(p, cb);
  }
  if (p.action === "getNotices") {
    return doGetNotices(p, cb);
  }
  if (p.action === "getDiary") {
    return doGetDiary(p, cb);
  }
  if (p.action === "getTransport") {
    return doGetTransport(p, cb);
  }
  if (p.action === "getFees") {
    return doGetFees(p, cb);
  }
  if (p.action === "getTeacherDashboard") {
    return doGetTeacherDashboard(p, cb);
  }
  if (p.action === "getMyLog") {
    return doGetMyLog(p, cb);
  }
  if (p.action === "getStudentContacts") {
    return doGetStudentContacts(p, cb);
  }
  if (p.action === "getStatsReport") {
    return doGetStatsReport(p, cb);
  }
  if (p.action === "deleteStudent") {
    return doDeleteStudent(p, cb);
  }
  var a = p.action || "";
  if (a.indexOf("save") === 0 || a.indexOf("upsert") === 0 ||
      a.indexOf("update") === 0) {
    return handleAction(p, cb);
  }
  return wrap(cb, {status:"ok", message:"Running!"});
}

function getNextSerial(sh) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  var vals = sh.getRange(2, 2, lastRow - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < vals.length; i++) {
    var n = parseInt(vals[i][0], 10);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function handleAction(d, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return wrap(cb, {status:"error", message:"Script NOT bound to a Sheet. Open Sheet → Extensions → Apps Script."});
    var action = d.action || "upsert";
    var result = {rowIndex:0, mode:"created"};

    if (action === "save" || action === "update" || action === "upsert") {
      var sh = getOrCreateSheet("Students", MAIN_HEADERS);
      var rowToWrite = 0;
      if (action === "update" && d.rowIndex) rowToWrite = parseInt(d.rowIndex,10);
      if (action === "upsert") rowToWrite = findRowByKey(sh, 3, d.regNo) || findRowByKey(sh, 4, d.studentId);
      // "संपर्क / Mobile" फील्ड फॉर्म मधून काढले आहे — जुनी नोंद असल्यास ती जागच्या जागी सुरक्षित ठेवा
      var contactVal = d.contact;
      if (!contactVal && rowToWrite) {
        contactVal = sh.getRange(rowToWrite, 25).getValue() || "";
      }
      // Category/Minority फॉर्म मधून रिकामे आल्यास (जुनी नोंद edit करताना चुकून रिकामे राहिल्यास) आधीचीच माहिती सुरक्षित ठेवा
      var categoryVal = d.category;
      var minorityVal = d.minority;
      if (!categoryVal && rowToWrite) categoryVal = sh.getRange(rowToWrite, 30).getValue() || "";
      if (!minorityVal && rowToWrite) minorityVal = sh.getRange(rowToWrite, 31).getValue() || "";
      var row = [new Date().toLocaleString("en-IN"),
        d.acYear,d.regNo,d.studentId,d.bookNo,d.aadhar,d.iyatta,d.tukdi,
        d.pen,d.rollNo,d.firstName,d.motherName,d.gender,
        d.religion,d.caste,d.subcaste,d.dob,d.dobWords,d.nationality,d.motherTongue,
        d.birthVillage,d.prevSchool,
        d.admissionDate,d.admissionClass,contactVal,d.address,d.photoUrl,
        d.whatsappMobile||"", d.alternateMobile||"",
        categoryVal||"", minorityVal||""];
      result = writeRow(sh, rowToWrite, row);
    } else if (action === "save_lc" || action === "update_lc" || action === "upsert_lc") {
      var sh = getOrCreateSheet("LC", LC_HEADERS);
      // ✅ सदैव नवीन row — प्रत्येक LC ला स्वतंत्र serial नंबर
      var rowToWrite = 0;
      var lcSerial = getNextSerial(sh);
      var row = [new Date().toLocaleString("en-IN"),
        lcSerial,d.stxt1,d.firstName,d.gender,d.dob,d.dobWords,
        d.iyatta,d.religion,d.caste,d.nationality,d.motherTongue,
        d.admissionDate,d.admissionClass,d.lcLeaveDate,d.lcLeaveClass,d.classStartDate,
        d.lcNo,d.lcDate,d.conduct,d.progress,d.stxt75,d.stxt76,d.remarks,
        d.medium,d.nextSchool,d.lcCount,d.feeStatus,
        d.stxt60,d.stxt61,d.stxt62,d.stxt63,d.stxt67,d.stxt68,d.stxt69,d.stxt70,d.nationalityType];
      result = writeRow(sh, rowToWrite, row);
      result.serial = lcSerial;
    } else if (action === "save_bf" || action === "upsert_bf") {
      var sh = getOrCreateSheet("Bonafide", BF_HEADERS);
      // ✅ सदैव नवीन row — प्रत्येक Bonafide ला स्वतंत्र serial नंबर
      var rowToWrite = 0;
      var bfSerial = getNextSerial(sh);
      var row = [new Date().toLocaleString("en-IN"),
        bfSerial,d.stxt1,d.firstName,d.dob,d.dobWords,
        d.iyatta,d.tukdi,d.religion,d.caste,d.nationality,
        d.acYear,d.bfDate,d.purpose,d.admissionDate,d.admissionClass,
        d.remarks,d.address,d.contact];
      result = writeRow(sh, rowToWrite, row);
      result.serial = bfSerial;
    } else if (action === "save_at" || action === "upsert_at") {
      var sh = getOrCreateSheet("Attendance", AT_HEADERS);
      // ✅ सदैव नवीन row — प्रत्येक Attendance certificate ला स्वतंत्र serial नंबर
      var rowToWrite = 0;
      var atSerial = getNextSerial(sh);
      var row = [new Date().toLocaleString("en-IN"),
        atSerial,d.stxt1,d.firstName,d.gender,d.dob,
        d.iyatta,d.tukdi,d.acYear,d.atDate,d.purpose,
        d.totalDays,d.presentDays,d.pct,d.fromDate,d.toDate,d.remarks];
      result = writeRow(sh, rowToWrite, row);
      result.serial = atSerial;
    } else if (action === "saveAttendance") {
      result = doSaveAttendance(d);
    } else if (action === "saveNotice") {
      result = doSaveNotice(d);
    } else if (action === "saveDiary") {
      result = doSaveDiary(d);
    } else if (action === "saveTransport") {
      result = doSaveTransport(d);
    } else if (action === "saveFees") {
      result = doSaveFees(d);
    } else if (action === "saveStudentContact") {
      result = doSaveStudentContact(d);
    } else if (action === "updateClassDivision") {
      result = doUpdateClassDivision(d);
    } else {
      return wrap(cb, {status:"error", message:"Unknown action: "+action});
    }
    if (result && result.status === "error") return wrap(cb, result);
    logAudit(d.auditUser, d.auditRole, action, result.serial || d.regNo || d.stxt1 || result.ref || "");
    return wrap(cb, {status:"ok", rowIndex:result.rowIndex, action:action, mode:result.mode, serial:result.serial||"", count:result.count||0,
      oldIyatta:result.oldIyatta||"", oldTukdi:result.oldTukdi||""});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ===== AUDIT LOG (V19.8) — कोणत्या User ने कधी काय Save केले =====
function logAudit(user, role, action, refInfo) {
  try {
    var sh = getOrCreateSheet("Log", ["Timestamp","User","Role","Action","Reference"]);
    sh.appendRow([new Date().toLocaleString("en-IN"), user || "unknown", role || "", action || "", refInfo || ""]);
  } catch(e) {
    // Logging failure should never block the main save operation
  }
}

// ===== DASHBOARD LIVE STATS (V19.8) =====
function doGetDashboardStats(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var todayStr = fmt(new Date());
    var now = new Date();
    var monthKey = now.getFullYear() + "-" + ("0"+(now.getMonth()+1)).slice(-2);

    var stSh = ss.getSheetByName("Students");
    var totalStudents = 0;
    var classCounts = {};
    if (stSh && stSh.getLastRow() > 1) {
      var stRows = stSh.getRange(2, 1, stSh.getLastRow()-1, 11).getValues();
      for (var i=0;i<stRows.length;i++) {
        if (!stRows[i][2] && !stRows[i][10]) continue;
        totalStudents++;
        var cls = (stRows[i][6]||"").toString().trim();
        if (cls) classCounts[cls] = (classCounts[cls]||0) + 1;
      }
    }

    function countTodayAndMonth(sheetName) {
      var sh = ss.getSheetByName(sheetName);
      var today = 0, month = 0;
      if (sh && sh.getLastRow() > 1) {
        var rows = sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues();
        for (var i=0;i<rows.length;i++) {
          var ts = rows[i][0];
          if (!ts) continue;
          var d = new Date(ts);
          if (isNaN(d.getTime())) continue;
          var dKey = fmt(d);
          var mKey = d.getFullYear() + "-" + ("0"+(d.getMonth()+1)).slice(-2);
          if (dKey === todayStr) today++;
          if (mKey === monthKey) month++;
        }
      }
      return {today:today, month:month};
    }

    var lcCounts = countTodayAndMonth("LC");
    var bfCounts = countTodayAndMonth("Bonafide");
    var atCounts = countTodayAndMonth("Attendance");
    var monthCerts = lcCounts.month + bfCounts.month + atCounts.month;

    return wrap(cb, {
      status:"ok",
      totalStudents: totalStudents,
      todayLC: lcCounts.today,
      todayBF: bfCounts.today,
      monthCerts: monthCerts,
      classCounts: classCounts
    });
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ===== LC/BF/AT HISTORY FOR ONE STUDENT (V19.8) =====
function doGetStudentHistory(p, cb) {
  try {
    var regNo = (p.regNo || "").toString().trim().toLowerCase();
    if (!regNo) return wrap(cb, {status:"error", message:"regNo required"});
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = [];

    function scan(sheetName, type, detailFn) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if ((r[2]||"").toString().trim().toLowerCase() !== regNo) continue;
        out.push({
          type: type,
          serial: r[1],
          date: fmt(r[0]),
          detail: detailFn(r)
        });
      }
    }
    scan("LC", "LC", function(r){ return "LC No: " + (r[17]||"—"); });
    scan("Bonafide", "BF", function(r){ return "Purpose: " + (r[13]||"—"); });
    scan("Attendance", "AT", function(r){ return "Purpose: " + (r[10]||"—"); });

    out.sort(function(a,b){ return (a.date < b.date) ? 1 : -1; });
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ===== BULK EXPORT — ALL LC/BF/AT (V19.8) =====
function doGetAllCertificates(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var out = [];

    function scan(sheetName, type, detailFn) {
      var sh = ss.getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[2]) continue;
        out.push({
          type: type,
          serial: r[1],
          date: fmt(r[0]),
          regNo: r[2],
          firstName: r[3],
          detail: detailFn(r)
        });
      }
    }
    scan("LC", "LC", function(r){ return "LC No: " + (r[17]||"—"); });
    scan("Bonafide", "BF", function(r){ return "Purpose: " + (r[13]||"—"); });
    scan("Attendance", "AT", function(r){ return "Purpose: " + (r[10]||"—"); });

    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================
// 👥 USER MANAGEMENT PRO (V19.9)
// =====================================================
var USER_HEADERS = ["Username","Password","Role","Label","AssignedClass"];
var DEFAULT_USERS_SEED = [
  ["user_1","Pass@1234","master","Master User",""],
  ["user_2","Pass@1234","deo","DEO User",""],
  ["user_3","Pass@1234","cert","Certificate User",""]
];

function getUsersSheet() {
  var sh = getOrCreateSheet("Users", USER_HEADERS);
  if (sh.getLastRow() < 2) {
    for (var i=0;i<DEFAULT_USERS_SEED.length;i++) sh.appendRow(DEFAULT_USERS_SEED[i]);
  }
  return sh;
}

function doGetUsers(p, cb) {
  try {
    var sh = getUsersSheet();
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 5).getValues();
    var out = [];
    for (var i=0;i<rows.length;i++) {
      if (!rows[i][0]) continue;
      out.push({ rowIndex:i+2, username:rows[i][0], password:rows[i][1], role:rows[i][2], label:rows[i][3], assignedClass:rows[i][4]||"" });
    }
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doSaveUser(p, cb) {
  try {
    // फक्त Super Master role असलेल्या logged-in user कडूनच call यायला हवा — frontend कडून व्हॅलिडेशन,
    // तसेच backend मध्येही requesterRole तपासतो
    if (p.requesterRole !== "super") {
      return wrap(cb, {status:"error", message:"User Management अधिकार फक्त Super Master User ला आहे."});
    }
    var sh = getUsersSheet();
    var username = (p.username||"").toString().trim();
    if (!username) return wrap(cb, {status:"error", message:"Username आवश्यक आहे."});
    var rowToWrite = findRowByKey(sh, 1, username);
    var row = [username, p.password||"", p.role||"cert", p.label||username, p.assignedClass||""];
    if (rowToWrite) {
      sh.getRange(rowToWrite, 1, 1, 5).setValues([row]);
    } else {
      sh.appendRow(row);
    }
    logAudit(p.requesterUser, p.requesterRole, "saveUser", username);
    return wrap(cb, {status:"ok"});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doDeleteUser(p, cb) {
  try {
    if (p.requesterRole !== "super") {
      return wrap(cb, {status:"error", message:"User Management अधिकार फक्त Super Master User ला आहे."});
    }
    var sh = getUsersSheet();
    var username = (p.username||"").toString().trim();
    var rowToWrite = findRowByKey(sh, 1, username);
    if (!rowToWrite) return wrap(cb, {status:"error", message:"User सापडला नाही."});
    sh.deleteRow(rowToWrite);
    logAudit(p.requesterUser, p.requesterRole, "deleteUser", username);
    return wrap(cb, {status:"ok"});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doChangePassword(p, cb) {
  try {
    var sh = getUsersSheet();
    var username = (p.username||"").toString().trim();
    var rowToWrite = findRowByKey(sh, 1, username);
    if (!rowToWrite) return wrap(cb, {status:"error", message:"User सापडला नाही."});
    var current = sh.getRange(rowToWrite, 2).getValue();
    if ((current||"").toString() !== (p.oldPassword||"").toString()) {
      return wrap(cb, {status:"error", message:"जुना Password चुकीचा आहे."});
    }
    sh.getRange(rowToWrite, 2).setValue(p.newPassword||"");
    logAudit(username, p.requesterRole, "changePassword", username);
    return wrap(cb, {status:"ok"});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================
// ✅ CERTIFICATE VERIFY PAGE — QR कोड scan केल्यावर उघडणारे पान (V19.9)
// =====================================================
function doVerifyPage(p) {
  var type = (p.type||"").toString().toUpperCase();
  var serial = (p.serial||"").toString().trim();
  var sheetMap = { LC:"LC", BF:"Bonafide", AT:"Attendance" };
  var sheetName = sheetMap[type];
  var found = null;
  try {
    if (sheetName && serial) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sh = ss.getSheetByName(sheetName);
      if (sh && sh.getLastRow() > 1) {
        var rows = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
        for (var i=0;i<rows.length;i++) {
          if ((rows[i][1]||"").toString() === serial) {
            found = { serial:rows[i][1], regNo:rows[i][2], firstName:rows[i][3], date:fmt(rows[i][0]) };
            break;
          }
        }
      }
    }
  } catch(e) {}

  var html;
  if (found) {
    html = "<div style='font-family:sans-serif;max-width:420px;margin:40px auto;padding:24px;border:2px solid #1a7a3a;border-radius:10px;text-align:center;background:#f4fff4'>"
      + "<div style='font-size:40px'>✅</div>"
      + "<h2 style='color:#1a7a3a'>Certificate Verified</h2>"
      + "<p><b>Type:</b> " + type + "</p>"
      + "<p><b>Serial No:</b> " + found.serial + "</p>"
      + "<p><b>Reg No:</b> " + found.regNo + "</p>"
      + "<p><b>Name:</b> " + found.firstName + "</p>"
      + "<p><b>Date:</b> " + found.date + "</p>"
      + "<p style='font-size:12px;color:#666;margin-top:16px'>Shri Govindram Seksaria High School, Pachora</p>"
      + "</div>";
  } else {
    html = "<div style='font-family:sans-serif;max-width:420px;margin:40px auto;padding:24px;border:2px solid #7a1a1a;border-radius:10px;text-align:center;background:#fff4f4'>"
      + "<div style='font-size:40px'>❌</div>"
      + "<h2 style='color:#7a1a1a'>Not Verified</h2>"
      + "<p>सदर Certificate आमच्या records मध्ये आढळले नाही.</p>"
      + "</div>";
  }
  return HtmlService.createHtmlOutput(html);
}

// =====================================================
// 📊 ANALYTICS & REPORTING (V19.9)
// =====================================================
function classifyLcReason(remarks) {
  var t = (remarks||"").toString().toLowerCase();
  if (t.indexOf("बदली") !== -1 || t.indexOf("transfer") !== -1) return "Transfer";
  if (t.indexOf("प्रगती") !== -1 || t.indexOf("progress") !== -1) return "Progress";
  return "Other";
}

function doGetAnalyticsData(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var admissionsByYear = {};
    var genderCounts = {};
    var religionCounts = {};
    var casteCounts = {};
    var lcReasonCounts = { Transfer:0, Progress:0, Other:0 };

    var stSh = ss.getSheetByName("Students");
    if (stSh && stSh.getLastRow() > 1) {
      var rows = stSh.getRange(2, 1, stSh.getLastRow()-1, 26).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[2] && !r[10]) continue;
        var admDate = r[22];
        if (admDate) {
          var yr = (admDate instanceof Date) ? admDate.getFullYear() : new Date(admDate).getFullYear();
          if (!isNaN(yr)) admissionsByYear[yr] = (admissionsByYear[yr]||0) + 1;
        }
        var gender = (r[12]||"Unknown").toString();
        genderCounts[gender] = (genderCounts[gender]||0) + 1;
        var religion = (r[13]||"Unknown").toString().trim();
        if (religion) religionCounts[religion] = (religionCounts[religion]||0) + 1;
        var caste = (r[14]||"Unknown").toString().trim();
        if (caste) casteCounts[caste] = (casteCounts[caste]||0) + 1;
      }
    }

    var lcSh = ss.getSheetByName("LC");
    if (lcSh && lcSh.getLastRow() > 1) {
      var lcRows = lcSh.getRange(2, 1, lcSh.getLastRow()-1, 24).getValues();
      for (var j=0;j<lcRows.length;j++) {
        if (!lcRows[j][2]) continue;
        var reason = classifyLcReason(lcRows[j][23]);
        lcReasonCounts[reason] = (lcReasonCounts[reason]||0) + 1;
      }
    }

    return wrap(cb, {
      status:"ok",
      admissionsByYear: admissionsByYear,
      genderCounts: genderCounts,
      religionCounts: religionCounts,
      casteCounts: casteCounts,
      lcReasonCounts: lcReasonCounts
    });
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// =====================================================
// 👩‍🏫 CLASS TEACHER MODULE (V19.32)
// =====================================================
function ckey(iyatta, tukdi) {
  return (iyatta||"").toString().trim() + "|" + (tukdi||"").toString().trim();
}
function todayStr() { return fmt(new Date()); }
function safeParseJson(s) {
  try { return JSON.parse(s || "[]"); } catch(e) { return []; }
}

// ----- Attendance (Daily register — only absentees are stored) -----
function doSaveAttendance(d) {
  try {
    var iyatta = (d.iyatta||"").toString().trim();
    var tukdi = (d.tukdi||"").toString().trim();
    var date = (d.date||todayStr()).toString().trim();
    if (!iyatta) return {status:"error", message:"Class आवश्यक आहे."};
    var sh = getOrCreateSheet("DailyAttendance", DAILY_ATT_HEADERS);
    // आधीचे त्या दिवसाचे/वर्गाचे रेकॉर्ड्स काढून टाका, मग नवीन यादी लिहा
    if (sh.getLastRow() > 1) {
      var vals = sh.getRange(2,1,sh.getLastRow()-1,4).getValues();
      for (var i = vals.length-1; i>=0; i--) {
        var rDate = fmt(vals[i][1]);
        if (rDate === date && (vals[i][2]||"").toString().trim() === iyatta && (vals[i][3]||"").toString().trim() === tukdi) {
          sh.deleteRow(i+2);
        }
      }
    }
    var absentList = safeParseJson(d.absentJson);
    var ts = new Date().toLocaleString("en-IN");
    for (var j=0;j<absentList.length;j++) {
      var a = absentList[j];
      sh.appendRow([ts, date, iyatta, tukdi, a.regNo||"", a.studentId||"", a.fullName||"", a.reason||"", d.markedBy||""]);
    }
    return {rowIndex: sh.getLastRow(), mode:"updated", count: absentList.length, ref: iyatta+"-"+tukdi+" "+date};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetAttendance(p, cb) {
  try {
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var date = (p.date||"").toString().trim();
    var dateFrom = (p.dateFrom||"").toString().trim();
    var dateTo = (p.dateTo||"").toString().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DailyAttendance");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,9).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        var rDate = fmt(r[1]);
        if (iyatta && (r[2]||"").toString().trim() !== iyatta) continue;
        if (tukdi && (r[3]||"").toString().trim() !== tukdi) continue;
        if (date && rDate !== date) continue;
        if (dateFrom && rDate < dateFrom) continue;
        if (dateTo && rDate > dateTo) continue;
        out.push({date:rDate, iyatta:r[2], tukdi:r[3], regNo:r[4], studentId:r[5], fullName:r[6], reason:r[7], markedBy:r[8]});
      }
    }
    out.sort(function(a,b){ return (a.date < b.date) ? 1 : -1; });
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Notices (posted by Super Master / Master, visible to Class Teachers) -----
function doSaveNotice(d) {
  if (d.requesterRole !== "super" && d.requesterRole !== "master") {
    return {status:"error", message:"सूचना पोस्ट करण्याचा अधिकार फक्त Super Master / Master User ला आहे."};
  }
  try {
    var sh = getOrCreateSheet("Notices", NOTICE_HEADERS);
    var row = [new Date().toLocaleString("en-IN"), d.date||todayStr(), d.title||"", d.message||"", d.postedBy||d.requesterUser||"", d.targetClass||""];
    sh.appendRow(row);
    return {rowIndex: sh.getLastRow(), mode:"created", ref: d.title||""};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetNotices(p, cb) {
  try {
    var targetClass = (p.targetClass||"").toString().trim();
    var limit = parseInt(p.limit||"20",10);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Notices");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,6).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        var tc = (r[5]||"").toString().trim();
        if (tc && targetClass && tc !== targetClass) continue;
        out.push({timestamp:r[0], date:fmt(r[1])||r[1], title:r[2], message:r[3], postedBy:r[4], targetClass:tc});
      }
    }
    out.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });
    return wrap(cb, {status:"ok", data: out.slice(0, limit)});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Student Diary (Updown/transport info + Positive/Warning/Scholarship remarks) -----
function doSaveDiary(d) {
  try {
    var sh = getOrCreateSheet("DiaryRemarks", DIARY_HEADERS);
    var row = [new Date().toLocaleString("en-IN"), d.date||todayStr(), d.regNo||"", d.studentId||"", d.fullName||"",
      d.iyatta||"", d.tukdi||"", d.type||"", d.remark||"", d.enteredBy||""];
    sh.appendRow(row);
    return {rowIndex: sh.getLastRow(), mode:"created", ref: d.fullName||d.regNo||""};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetDiary(p, cb) {
  try {
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var type = (p.type||"").toString().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("DiaryRemarks");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (regNo && (r[2]||"").toString().trim().toLowerCase() !== regNo) continue;
        if (!regNo && iyatta && (r[5]||"").toString().trim() !== iyatta) continue;
        if (!regNo && tukdi && (r[6]||"").toString().trim() !== tukdi) continue;
        if (type && (r[7]||"").toString().trim() !== type) continue;
        out.push({timestamp:r[0], date:fmt(r[1])||r[1], regNo:r[2], studentId:r[3], fullName:r[4], iyatta:r[5], tukdi:r[6], type:r[7], remark:r[8], enteredBy:r[9]});
      }
    }
    out.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Transport / Updown info (one record per student, upsert) -----
function doSaveTransport(d) {
  try {
    var sh = getOrCreateSheet("StudentTransport", TRANSPORT_HEADERS);
    var rowToWrite = findRowByKey(sh, 2, d.regNo);
    var row = [new Date().toLocaleString("en-IN"), d.regNo||"", d.studentId||"", d.fullName||"",
      d.iyatta||"", d.tukdi||"", d.nativeVillage||"", d.travelMode||"", d.transportContact||"", d.updatedBy||""];
    var res = writeRow(sh, rowToWrite, row);
    res.ref = d.fullName||d.regNo||"";
    return res;
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetTransport(p, cb) {
  try {
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StudentTransport");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,10).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (regNo && (r[1]||"").toString().trim().toLowerCase() !== regNo) continue;
        if (!regNo && iyatta && (r[4]||"").toString().trim() !== iyatta) continue;
        if (!regNo && tukdi && (r[5]||"").toString().trim() !== tukdi) continue;
        out.push({regNo:r[1], studentId:r[2], fullName:r[3], iyatta:r[4], tukdi:r[5], nativeVillage:r[6], travelMode:r[7], transportContact:r[8], updatedBy:r[9]});
      }
    }
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Fees (टप्प्याटप्प्याने जमा होणारी रक्कम — प्रत्येक एन्ट्री ही एक जमा नोंद, त्या क्षणापर्यंतच्या
//        cumulative TotalFee/FeePaid/PendingFee सह लॉग केली जाते — V19.33) -----
function doSaveFees(d) {
  try {
    if (d.requesterRole !== "super" && d.requesterRole !== "master" && d.requesterRole !== "teacher") {
      return {status:"error", message:"Fees माहिती Save करण्याचा अधिकार Super Master / Master / Class Teacher ला आहे."};
    }
    // Class Teacher फक्त स्वतःच्या वर्गाच्याच विद्यार्थ्यांची Fees नोंद करू शकतो
    if (d.requesterRole === "teacher") {
      var tIyatta = (d.teacherIyatta||"").toString().trim();
      var tTukdi = (d.teacherTukdi||"").toString().trim();
      if ((d.iyatta||"").toString().trim() !== tIyatta || (d.tukdi||"").toString().trim() !== tTukdi) {
        return {status:"error", message:"आपण फक्त स्वतःच्या वर्गाच्याच विद्यार्थ्यांची Fees नोंद करू शकता."};
      }
    }
    var amount = parseFloat(d.amountPaid)||0;
    if (amount <= 0) return {status:"error", message:"जमा रक्कम बरोबर टाका."};
    var sh = getOrCreateSheet("Fees", FEES_HEADERS);
    var regNoKey = (d.regNo||"").toString().trim().toLowerCase();

    // या विद्यार्थ्याने आधी किती जमा केले आहे ते शोधा (cumulative FeePaid, sheet मध्ये सर्वात शेवटची त्याची नोंद)
    var prevPaid = 0;
    if (sh.getLastRow() > 1) {
      var lastRow = sh.getLastRow();
      var vals = sh.getRange(2, 1, lastRow - 1, 11).getValues();
      for (var i = vals.length - 1; i >= 0; i--) {
        if ((vals[i][1]||"").toString().trim().toLowerCase() === regNoKey) {
          prevPaid = parseFloat(vals[i][8]) || 0; // column 9 = FeePaid (cumulative)
          break;
        }
      }
    }

    var totalFee = TOTAL_FEE_PER_STUDENT;
    var feePaid = prevPaid + amount;
    var pendingFee = Math.max(0, totalFee - feePaid);
    var now = new Date();
    var row = [now.toLocaleString("en-IN"), d.regNo||"", d.studentId||"", d.fullName||"",
      d.iyatta||"", d.tukdi||"", d.acYear||"", totalFee, feePaid, pendingFee, d.enteredBy||""];
    sh.appendRow(row);
    return {rowIndex: sh.getLastRow(), mode:"created", ref: d.fullName||d.regNo||""};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetFees(p, cb) {
  try {
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Fees");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,11).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (regNo && (r[1]||"").toString().trim().toLowerCase() !== regNo) continue;
        if (!regNo && iyatta && (r[4]||"").toString().trim() !== iyatta) continue;
        if (!regNo && tukdi && (r[5]||"").toString().trim() !== tukdi) continue;
        out.push({date:fmt(r[0])||r[0], regNo:r[1], studentId:r[2], fullName:r[3], iyatta:r[4], tukdi:r[5],
          acYear:r[6], totalFee:r[7], feePaid:r[8], pendingFee:r[9], updatedBy:r[10]});
      }
    }
    // नवीन-ते-जुने क्रमाने — entries[0] म्हणजे प्रत्येक विद्यार्थ्याची सर्वात अलीकडची (latest cumulative) नोंद
    out.sort(function(a,b){ return (a.date < b.date) ? 1 : -1; });
    return wrap(cb, {status:"ok", data: out, totalFeePerStudent: TOTAL_FEE_PER_STUDENT});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- विद्यार्थी Delete करणे — पूर्ण माहिती आधी DeletedStudents Sheet मध्ये जतन करूनच (V19.32) -----
function doDeleteStudent(p, cb) {
  try {
    if (p.requesterRole !== "super" && p.requesterRole !== "master") {
      return wrap(cb, {status:"error", message:"विद्यार्थी Delete करण्याचा अधिकार फक्त Super Master / Master User ला आहे."});
    }
    var regNo = (p.regNo||"").toString().trim();
    if (!regNo) return wrap(cb, {status:"error", message:"Reg No आवश्यक आहे."});
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh) return wrap(cb, {status:"error", message:"Students sheet सापडली नाही."});
    var rowToWrite = findRowByKey(sh, 3, regNo); // कॉलम C = Reg No
    if (!rowToWrite) return wrap(cb, {status:"error", message:"हा Reg No असलेला विद्यार्थी सापडला नाही."});

    var rowData = sh.getRange(rowToWrite, 1, 1, 31).getValues()[0];
    var delHeaders = MAIN_HEADERS.concat(["Deleted By", "Deleted Date"]);
    var delSh = getOrCreateSheet("DeletedStudents", delHeaders);
    var newRow = rowData.concat([p.requesterUser || "", new Date().toLocaleString("en-IN")]);
    delSh.appendRow(newRow);

    sh.deleteRow(rowToWrite);
    logAudit(p.requesterUser, p.requesterRole, "deleteStudent", regNo);
    return wrap(cb, {status:"ok", regNo: regNo});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- वर्ग/तुकडी बदल — मेंटेनन्स (V19.20) -----
function doUpdateClassDivision(d) {
  try {
    if (d.requesterRole !== "super" && d.requesterRole !== "master" && d.requesterRole !== "teacher") {
      return {status:"error", message:"अधिकार नाही."};
    }
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh) return {status:"error", message:"Students sheet सापडली नाही."};
    var rowToWrite = findRowByKey(sh, 3, d.regNo); // कॉलम C = Reg No
    if (!rowToWrite) return {status:"error", message:"विद्यार्थी सापडला नाही."};
    var curIyatta = sh.getRange(rowToWrite, 7).getValue().toString().trim();
    var curTukdi = sh.getRange(rowToWrite, 8).getValue().toString().trim();
    if (d.requesterRole === "teacher") {
      var tIyatta = (d.teacherIyatta||"").toString().trim();
      if (curIyatta !== tIyatta) {
        return {status:"error", message:"हा विद्यार्थी आपल्या वर्गाचा नाही."};
      }
      if ((d.iyatta||"").toString().trim() !== curIyatta) {
        return {status:"error", message:"Class Teacher फक्त तुकडी बदलू शकतो, वर्ग नाही."};
      }
    }
    sh.getRange(rowToWrite, 7, 1, 2).setValues([[d.iyatta||"", d.tukdi||""]]);
    return {rowIndex: rowToWrite, mode:"updated", ref: d.regNo||"",
      oldIyatta: curIyatta, oldTukdi: curTukdi};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

// ----- सांख्यिकी अहवाल — एकाच call मध्ये सर्व माहिती एकत्र (Performance — V19.21) -----
function statsAge(dobStr, asOfStr) {
  if (!dobStr) return null;
  var d = new Date(dobStr);
  if (isNaN(d.getTime())) return null;
  var asOf = new Date(asOfStr);
  var age = asOf.getFullYear() - d.getFullYear();
  var mo = asOf.getMonth() - d.getMonth();
  if (mo < 0 || (mo === 0 && asOf.getDate() < d.getDate())) age--;
  return age;
}

function doGetStatsReport(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var today = fmt(new Date());
    var asOf = (p.asOfDate || "2026-09-30").toString();
    var classMap = {};

    var stSh = ss.getSheetByName("Students");
    if (stSh && stSh.getLastRow() > 1) {
      var rows = stSh.getRange(2, 1, stSh.getLastRow()-1, 31).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[10] && !r[2]) continue;
        var iyatta = (r[6]||"").toString().trim();
        var tukdi = (r[7]||"").toString().trim();
        var key = iyatta + "|" + tukdi;
        if (!classMap[key]) classMap[key] = {
          iyatta: iyatta, tukdi: tukdi, total: 0, absent: 0, fees: 0,
          female: {}, male: {}, ageCounts: {}
        };
        var c = classMap[key];
        c.total++;
        var gender = r[12];
        var cat = (r[29]||"").toString().trim();
        var minority = (r[30]||"").toString().trim().toLowerCase();
        var bucket = gender === "Female" ? c.female : (gender === "Male" ? c.male : null);
        if (bucket) {
          if (CATEGORY_LIST_SRV.indexOf(cat) !== -1) bucket[cat] = (bucket[cat]||0) + 1;
          if (minority === "yes") bucket.Minority = (bucket.Minority||0) + 1;
        }
        var age = statsAge(fmt(r[16]), asOf);
        if (age !== null) c.ageCounts[age] = (c.ageCounts[age]||0) + 1;
      }
    }

    var attSh = ss.getSheetByName("DailyAttendance");
    if (attSh && attSh.getLastRow() > 1) {
      var attRows = attSh.getRange(2, 1, attSh.getLastRow()-1, 9).getValues();
      for (var j=0;j<attRows.length;j++) {
        var ar = attRows[j];
        if (fmt(ar[1]) !== today) continue;
        var key2 = (ar[2]||"").toString().trim() + "|" + (ar[3]||"").toString().trim();
        if (classMap[key2]) classMap[key2].absent++;
      }
    }

    // Fees sheet मध्ये आता प्रत्येक installment ला cumulative FeePaid असतो, त्यामुळे सरळ बेरीज न करता
    // प्रत्येक विद्यार्थ्याची फक्त सर्वात अलीकडची (शेवटची) नोंद घ्यावी लागते, मगच वर्गनिहाय बेरीज करायची
    var feeSh = ss.getSheetByName("Fees");
    if (feeSh && feeSh.getLastRow() > 1) {
      var feeRows = feeSh.getRange(2, 1, feeSh.getLastRow()-1, 11).getValues();
      var latestByReg = {};
      for (var k=0;k<feeRows.length;k++) {
        var fr = feeRows[k];
        var rg = (fr[1]||"").toString().trim().toLowerCase();
        if (!rg) continue;
        // rows नेहमी वेळेनुसार appended असतात, त्यामुळे शेवटी सापडलेली नोंद = सर्वात अलीकडची
        latestByReg[rg] = { iyatta:(fr[4]||"").toString().trim(), tukdi:(fr[5]||"").toString().trim(), feePaid: parseFloat(fr[8])||0 };
      }
      for (var rgKey in latestByReg) {
        var e = latestByReg[rgKey];
        var key3 = e.iyatta + "|" + e.tukdi;
        if (classMap[key3]) classMap[key3].fees += e.feePaid;
      }
    }

    var classes = Object.keys(classMap).map(function(k){ return classMap[k]; });
    classes.sort(function(a,b){
      var ka = a.iyatta+"|"+a.tukdi, kb = b.iyatta+"|"+b.tukdi;
      return ka < kb ? -1 : (ka > kb ? 1 : 0);
    });
    return wrap(cb, {status:"ok", classes: classes, today: today, asOfDate: asOf, categoryList: CATEGORY_LIST_SRV});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- विद्यार्थ्याचे WhatsApp Mobile + इतर Mobile (Class Teacher entry) -----
function doSaveStudentContact(d) {
  try {
    var sh = getOrCreateSheet("StudentContacts", CONTACTS_HEADERS);
    var rowToWrite = findRowByKey(sh, 2, d.regNo);
    var row = [new Date().toLocaleString("en-IN"), d.regNo||"", d.studentId||"", d.fullName||"",
      d.iyatta||"", d.tukdi||"", d.whatsappMobile||"", d.otherMobile||"", d.updatedBy||""];
    var res = writeRow(sh, rowToWrite, row);
    res.ref = d.fullName||d.regNo||"";
    return res;
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function doGetStudentContacts(p, cb) {
  try {
    var regNo = (p.regNo||"").toString().trim().toLowerCase();
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("StudentContacts");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,9).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (regNo && (r[1]||"").toString().trim().toLowerCase() !== regNo) continue;
        if (!regNo && iyatta && (r[4]||"").toString().trim() !== iyatta) continue;
        if (!regNo && tukdi && (r[5]||"").toString().trim() !== tukdi) continue;
        out.push({regNo:r[1], studentId:r[2], fullName:r[3], iyatta:r[4], tukdi:r[5], whatsappMobile:r[6], otherMobile:r[7], updatedBy:r[8]});
      }
    }
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Class Teacher's student list (filtered by their assigned class) -----
function doGetClassStudents(p, cb) {
  try {
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,31).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[10] && !r[2]) continue;
        if (iyatta && (r[6]||"").toString().trim() !== iyatta) continue;
        if (tukdi && (r[7]||"").toString().trim() !== tukdi) continue;
        out.push({
          regNo:r[2], studentId:r[3], acYear:r[1], pen:r[8], rollNo:r[9], fullName:r[10],
          motherName:r[11], gender:r[12], dob:fmt(r[16]), contact:r[24], address:r[25], photoUrl:r[26],
          whatsappMobile:r[27]||"", alternateMobile:r[28]||"",
          category:(r[29]||"").toString().trim(), minority:(r[30]||"").toString().trim(),
          iyatta:r[6], tukdi:r[7]
        });
      }
    }
    out.sort(function(a,b){ return (parseInt(a.rollNo,10)||0) - (parseInt(b.rollNo,10)||0); });
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Notice Board साठी वर्ग-तुकडी यादी — Students sheet मध्ये प्रत्यक्षात असलेल्या इयत्ता-तुकडी जोड्या (V19.34) -----
function doGetClassList(p, cb) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    var out = [];
    var seen = {};
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 7, sh.getLastRow()-1, 2).getValues(); // कॉलम G=Class, H=Division
      for (var i=0;i<rows.length;i++) {
        var iyatta = (rows[i][0]||"").toString().trim();
        var tukdi = (rows[i][1]||"").toString().trim();
        if (!iyatta) continue;
        var key = iyatta + "|" + tukdi;
        if (seen[key]) continue;
        seen[key] = true;
        out.push({iyatta: iyatta, tukdi: tukdi});
      }
    }
    out.sort(function(a,b){
      var ka = a.iyatta+"|"+a.tukdi, kb = b.iyatta+"|"+b.tukdi;
      return ka < kb ? -1 : (ka > kb ? 1 : 0);
    });
    return wrap(cb, {status:"ok", data: out});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- आज वाढदिवस असलेले विद्यार्थी — संपूर्ण शाळेतील (Super Master / Master Dashboard — V19.34) -----
function doGetTodayBirthdays(p, cb) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var today = fmt(new Date());
    var todayMD = today.slice(5); // MM-DD
    var todayYear = new Date().getFullYear();
    var sh = ss.getSheetByName("Students");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2, 1, sh.getLastRow()-1, 31).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (!r[10] && !r[2]) continue; // रिकामी row वगळा
        var dobStr = fmt(r[16]);
        if (!dobStr || dobStr.slice(5) !== todayMD) continue;
        var birthYear = parseInt(dobStr.slice(0,4), 10);
        var age = !isNaN(birthYear) ? (todayYear - birthYear) : null;
        out.push({
          iyatta: r[6], tukdi: r[7], regNo: r[2], studentId: r[3], rollNo: r[9],
          fullName: r[10], dob: dobStr, age: age,
          whatsappMobile: r[27]||"", alternateMobile: r[28]||"", contact: r[24]||"",
          photoUrl: r[26]||""
        });
      }
    }
    out.sort(function(a,b){
      var ka = (a.iyatta||"").toString()+"|"+(a.tukdi||"").toString();
      var kb = (b.iyatta||"").toString()+"|"+(b.tukdi||"").toString();
      return ka < kb ? -1 : (ka > kb ? 1 : 0);
    });
    return wrap(cb, {status:"ok", data: out, today: today});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

// ----- Combined Teacher Dashboard -----
function doGetTeacherDashboard(p, cb) {
  try {
    var iyatta = (p.iyatta||"").toString().trim();
    var tukdi = (p.tukdi||"").toString().trim();
    var date = (p.date||todayStr()).toString().trim();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var totalStudents = 0;
    var birthdayStudents = [];
    var todayMD = date.slice(5); // MM-DD
    var stSh = ss.getSheetByName("Students");
    if (stSh && stSh.getLastRow() > 1) {
      var stRows = stSh.getRange(2,1,stSh.getLastRow()-1,27).getValues();
      for (var i=0;i<stRows.length;i++) {
        var r = stRows[i];
        if (!r[10] && !r[2]) continue;
        if (iyatta && (r[6]||"").toString().trim() !== iyatta) continue;
        if (tukdi && (r[7]||"").toString().trim() !== tukdi) continue;
        totalStudents++;
        var dobStr = fmt(r[16]);
        if (dobStr && dobStr.slice(5) === todayMD) {
          birthdayStudents.push({regNo:r[2], fullName:r[10], dob:dobStr, contact:r[24]});
        }
      }
    }

    var absentList = [];
    var attSh = ss.getSheetByName("DailyAttendance");
    if (attSh && attSh.getLastRow() > 1) {
      var attRows = attSh.getRange(2,1,attSh.getLastRow()-1,9).getValues();
      for (var j=0;j<attRows.length;j++) {
        var ar = attRows[j];
        if (fmt(ar[1]) !== date) continue;
        if (iyatta && (ar[2]||"").toString().trim() !== iyatta) continue;
        if (tukdi && (ar[3]||"").toString().trim() !== tukdi) continue;
        absentList.push({regNo:ar[4], studentId:ar[5], fullName:ar[6], reason:ar[7]});
      }
    }

    var notices = [];
    var noticeSh = ss.getSheetByName("Notices");
    if (noticeSh && noticeSh.getLastRow() > 1) {
      var ck = ckey(iyatta, tukdi);
      var nRows = noticeSh.getRange(2,1,noticeSh.getLastRow()-1,6).getValues();
      for (var k=0;k<nRows.length;k++) {
        var nr = nRows[k];
        var tc = (nr[5]||"").toString().trim();
        if (tc && tc !== iyatta+"|"+tukdi && tc !== iyatta) continue;
        notices.push({timestamp:nr[0], date:fmt(nr[1])||nr[1], title:nr[2], message:nr[3], postedBy:nr[4]});
      }
      notices.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });
      notices = notices.slice(0,5);
    }

    return wrap(cb, {
      status:"ok",
      totalStudents: totalStudents,
      absentCount: absentList.length,
      presentCount: totalStudents - absentList.length,
      absentList: absentList,
      birthdayStudents: birthdayStudents,
      notices: notices
    });
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doGetMyLog(p, cb) {
  try {
    var username = (p.username||"").toString().trim().toLowerCase();
    var limit = parseInt(p.limit||"100",10);
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
    var out = [];
    if (sh && sh.getLastRow() > 1) {
      var rows = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
      for (var i=0;i<rows.length;i++) {
        var r = rows[i];
        if (username && (r[1]||"").toString().trim().toLowerCase() !== username) continue;
        out.push({timestamp:r[0], user:r[1], role:r[2], action:r[3], reference:r[4]});
      }
    }
    out.sort(function(a,b){ return new Date(b.timestamp) - new Date(a.timestamp); });
    return wrap(cb, {status:"ok", data: out.slice(0, limit)});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function wrap(cb, obj) {
  var json = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function doGetAllAction(p, cb) {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh || sh.getLastRow() < 2) return wrap(cb, {status:"ok", data:[]});
    var rows = sh.getRange(2, 1, sh.getLastRow()-1, 31).getValues();
    var all = [];
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if (!r[10] && !r[2]) continue; // रिकामी row वगळा
      all.push({
        _rowIndex: i + 2,
        acYear:r[1], regNo:r[2], studentId:r[3], bookNo:r[4], aadhar:r[5],
        iyatta:r[6], tukdi:r[7], pen:r[8], rollNo:r[9],
        firstName:r[10], motherName:r[11], gender:r[12],
        religion:r[13], caste:r[14], subcaste:r[15], dob:fmt(r[16]), dobWords:r[17],
        nationality:r[18], motherTongue:r[19], birthVillage:r[20],
        prevSchool:r[21], admissionDate:fmt(r[22]), admissionClass:r[23],
        contact:r[24], address:r[25], photoUrl:r[26],
        whatsappMobile:r[27]||"", alternateMobile:r[28]||"",
        category:r[29]||"", minority:r[30]||""
      });
    }
    return wrap(cb, {status:"ok", data:all});
  } catch(err) {
    return wrap(cb, {status:"error", message:err.toString()});
  }
}

function doSearchAction(p, cb) {
  try {
    var q = (p.q || "").toString().trim().toLowerCase();
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Students");
    if (!sh || sh.getLastRow() < 2) return wrap(cb,{status:"notfound"});
    var rows = sh.getRange(2,1,sh.getLastRow()-1,31).getValues();
    for (var i=0;i<rows.length;i++) {
      var r=rows[i];
      var cands=[
        (r[2]||"").toString().trim().toLowerCase(),
        (r[3]||"").toString().trim().toLowerCase(),
        (r[5]||"").toString().trim().toLowerCase(),
        (r[10]||"").toString().trim().toLowerCase()
      ];
      if (cands.indexOf(q)!==-1) {
        return wrap(cb,{status:"found",rowIndex:i+2,data:{
          acYear:r[1],regNo:r[2],studentId:r[3],bookNo:r[4],aadhar:r[5],
          iyatta:r[6],tukdi:r[7],pen:r[8],rollNo:r[9],
          firstName:r[10],motherName:r[11],gender:r[12],
          religion:r[13],caste:r[14],subcaste:r[15],dob:fmt(r[16]),dobWords:r[17],nationality:r[18],
          motherTongue:r[19],birthVillage:r[20],
          prevSchool:r[21],admissionDate:fmt(r[22]),admissionClass:r[23],
          contact:r[24],address:r[25],photoUrl:r[26],
          whatsappMobile:r[27]||"",alternateMobile:r[28]||"",
          category:r[29]||"",minority:r[30]||""
        }});
      }
    }
    return wrap(cb,{status:"notfound"});
  } catch(err) {
    return wrap(cb,{status:"error",message:err.toString()});
  }
}
function fmt(v){
  if(!v) return "";
  if(v instanceof Date){
    return v.getFullYear()+"-"+("0"+(v.getMonth()+1)).slice(-2)+"-"+("0"+v.getDate()).slice(-2);
  }
  return v.toString();
}

function cleanRegFileName(regNo) {
  var text = (regNo || "student").toString().trim();
  var bad = '\\/:*?"<>|#%&{}$!\'@+=~';
  var out = "";
  for (var i = 0; i < text.length; i++) {
    var ch = text.charAt(i);
    out += bad.indexOf(ch) >= 0 ? "_" : ch;
  }
  return out || "student";
}

function publicDriveUrl(fileId) {
  return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
}

function testPhotoFolderAccess() {
  try {
    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var blob = Utilities.newBlob("ok", "text/plain", "_sgs_upload_test.txt");
    var file = folder.createFile(blob);
    file.setTrashed(true);
    return {status:"ok", folderName:folder.getName(), folderId:PHOTO_FOLDER_ID};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function savePhotoData(regNo, photoData, mimeType) {
  try {
    regNo = (regNo || "").toString().trim();
    if (!regNo) return {status:"error", message:"regNo required"};
    if (!photoData) return {status:"error", message:"photoData required"};

    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var fileName = cleanRegFileName(regNo) + ".jpg";
    var existing = folder.getFilesByName(fileName);
    while (existing.hasNext()) {
      existing.next().setTrashed(true);
    }
    var bytes = Utilities.base64Decode(photoData);
    if (bytes.length > 10240) return {status:"error", message:"Photo size is " + bytes.length + " bytes. Please compress below 10 KB."};
    var blob = Utilities.newBlob(bytes, mimeType || "image/jpeg", fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var photoUrl = publicDriveUrl(file.getId());
    updateStudentPhotoUrl(regNo, photoUrl);
    return {status:"ok", regNo:regNo, photoUrl:photoUrl, fileId:file.getId()};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function handlePhotoUpload(p) {
  return savePhotoData(p.regNo, p.photoData, p.mimeType);
}

function startPhotoChunkUpload(p) {
  try {
    if (!p.uploadId) return {status:"error", message:"uploadId required"};
    var cache = CacheService.getScriptCache();
    cache.put("photo_" + p.uploadId + "_meta", JSON.stringify({
      regNo:p.regNo || "",
      mimeType:p.mimeType || "image/jpeg",
      count:parseInt(p.count || "0", 10)
    }), 600);
    return {status:"ok", uploadId:p.uploadId};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function savePhotoChunk(p) {
  try {
    if (!p.uploadId) return {status:"error", message:"uploadId required"};
    if (p.idx === undefined) return {status:"error", message:"idx required"};
    if (p.chunk === undefined) return {status:"error", message:"chunk required"};
    CacheService.getScriptCache().put("photo_" + p.uploadId + "_" + p.idx, p.chunk, 600);
    return {status:"ok", idx:p.idx};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function finishPhotoChunkUpload(p) {
  try {
    var uploadId = p.uploadId || "";
    var count = parseInt(p.count || "0", 10);
    if (!uploadId) return {status:"error", message:"uploadId required"};
    if (!count || count < 1) return {status:"error", message:"count required"};
    var cache = CacheService.getScriptCache();
    var keys = [];
    for (var i = 0; i < count; i++) keys.push("photo_" + uploadId + "_" + i);
    var got = cache.getAll(keys);
    var parts = [];
    for (var j = 0; j < count; j++) {
      var part = got["photo_" + uploadId + "_" + j];
      if (part === null || part === undefined) return {status:"error", message:"Missing photo chunk " + (j + 1) + "/" + count};
      parts.push(part);
    }
    return savePhotoData(p.regNo, parts.join(""), p.mimeType || "image/jpeg");
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}

function updateStudentPhotoUrl(regNo, photoUrl) {
  var sh = getOrCreateSheet("Students", MAIN_HEADERS);
  var row = findRowByKey(sh, 3, regNo) || findRowByKey(sh, 4, regNo);
  if (row) {
    sh.getRange(row, 27).setValue(photoUrl);
    SpreadsheetApp.flush();
  }
}

function getPhotoUrlByRegNo(regNo) {
  try {
    regNo = (regNo || "").toString().trim();
    if (!regNo) return {status:"error", message:"regNo required"};
    var sh = getOrCreateSheet("Students", MAIN_HEADERS);
    var row = findRowByKey(sh, 3, regNo) || findRowByKey(sh, 4, regNo);
    var folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
    var files = folder.getFilesByName(cleanRegFileName(regNo) + ".jpg");
    if (files.hasNext()) {
      var file = files.next();
      var photoUrl = publicDriveUrl(file.getId());
      updateStudentPhotoUrl(regNo, photoUrl);
      return {status:"ok", photoUrl:photoUrl};
    }
    if (row) {
      var savedUrl = sh.getRange(row, 33).getValue();
      if (savedUrl) return {status:"ok", photoUrl:savedUrl};
    }
    return {status:"notfound", photoUrl:""};
  } catch(err) {
    return {status:"error", message:err.toString()};
  }
}
