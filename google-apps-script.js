// Google Apps Script for Chat App Backend
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Main sheet: Messages, backup sheet: MessagesBackup
// Images stored as base64 data URLs in the sheet

const MESSAGES_SHEET = 'Messages';
const BACKUP_SHEET = 'MessagesBackup';
const HEADERS = ['From', 'To', 'Message', 'Timestamp', 'Status', 'ImageData'];

// Hardcoded users
const USERS = {
  '9236647910': '0197',
  '9198433007': '7003'
};

const ONLINE_THRESHOLD_MS = 90 * 1000;
const TYPING_THRESHOLD_MS = 10 * 1000;

var lastSeen = {};
var typingUntil = {};

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(name) {
  name = name || MESSAGES_SHEET;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);

  if (sheet) {
    var headers = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    var match = true;
    for (var h = 0; h < HEADERS.length; h++) {
      if (headers[h] !== HEADERS[h]) { match = false; break; }
    }
    if (!match) {
      sheet.clear();
      sheet.appendRow(HEADERS);
      sheet.getRange('A:B').setNumberFormat('@');
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADERS);
    sheet.getRange('A:B').setNumberFormat('@');
  }

  return sheet;
}

function otherUser_(me) {
  var keys = Object.keys(USERS);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i] !== me) return keys[i];
  }
  return null;
}

function doGet(e) {
  try {
    var p = e.parameter;
    if (p.action === 'getMessages' && p.user) {
      return getMessages(String(p.user));
    }
    return jsonResponse({ success: false, message: 'Invalid request' });
  } catch (err) {
    return jsonResponse({ success: false, message: 'Error: ' + err.toString() });
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'login':
        return login(String(data.pin));
      case 'sendMessage':
        return sendMessage(String(data.from), String(data.to), String(data.message));
      case 'sendImage':
        return sendImage(String(data.from), String(data.to), String(data.imageData));
      case 'heartbeat':
        return heartbeat(String(data.mobile));
      case 'setTyping':
        return setTyping(String(data.mobile), data.isTyping);
      case 'markRead':
        return markRead(String(data.mobile));
      case 'clearChat':
        return clearChat_(String(data.user), data.before);
      default:
        return jsonResponse({ success: false, message: 'Invalid action' });
    }
  } catch (err) {
    return jsonResponse({ success: false, message: 'Error: ' + err.toString() });
  }
}

function login(pin) {
  var keys = Object.keys(USERS);
  for (var i = 0; i < keys.length; i++) {
    if (USERS[keys[i]] === pin) {
      lastSeen[keys[i]] = Date.now();
      return jsonResponse({
        success: true,
        user: keys[i],
        other: otherUser_(keys[i])
      });
    }
  }
  return jsonResponse({ success: false, message: 'Invalid PIN' });
}

function heartbeat(mobile) {
  if (USERS[mobile]) {
    lastSeen[mobile] = Date.now();
    return jsonResponse({ success: true });
  }
  return jsonResponse({ success: false });
}

function sendMessage(from, to, message) {
  if (!USERS[from] || !USERS[to] || !message) {
    return jsonResponse({ success: false, message: 'Invalid' });
  }
  var row = [from, to, message, new Date(), 'sent', ''];
  getSheet_().appendRow(row);
  backupRow_(row);
  typingUntil[from] = 0;
  return jsonResponse({ success: true });
}

function sendImage(from, to, dataUrl) {
  if (!USERS[from] || !USERS[to] || !dataUrl) {
    return jsonResponse({ success: false, message: 'Invalid' });
  }

  try {
    var row = [from, to, '', new Date(), 'sent', dataUrl];
    getSheet_().appendRow(row);
    backupRow_(row);
    typingUntil[from] = 0;
    return jsonResponse({ success: true });
  } catch (err) {
    return jsonResponse({ success: false, message: 'Error: ' + err.toString() });
  }
}

function getMessages(user) {
  var other = otherUser_(user);
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return jsonResponse({ success: true, messages: [], otherOnline: false, typing: false });
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var messages = [];

  for (var i = 0; i < rows.length; i++) {
    var from = String(rows[i][0]);
    var to = String(rows[i][1]);

    if ((from === user && to === other) || (from === other && to === user)) {
      var status = rows[i][4] || 'sent';

      if (from === other && to === user && status === 'sent') {
        sheet.getRange(i + 2, 5).setValue('delivered');
        status = 'delivered';
      }

      var ts = rows[i][3];
      var imageData = rows[i][5] ? String(rows[i][5]) : '';

      messages.push({
        from: from,
        to: to,
        message: String(rows[i][2]),
        timestamp: (ts instanceof Date) ? ts.toISOString() : String(ts),
        status: status,
        imageData: imageData
      });
    }
  }

  var now = Date.now();
  var otherOnline = false;
  if (lastSeen[other] && (now - lastSeen[other]) < ONLINE_THRESHOLD_MS) {
    otherOnline = true;
  }

  var typing = false;
  if (typingUntil[other] && typingUntil[other] > now) {
    typing = true;
  }

  return jsonResponse({
    success: true,
    messages: messages,
    otherOnline: otherOnline,
    typing: typing
  });
}

function markRead(user) {
  var other = otherUser_(user);
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) return jsonResponse({ success: true });

  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === other && String(rows[i][1]) === user && rows[i][4] !== 'read') {
      sheet.getRange(i + 2, 5).setValue('read');
    }
  }

  return jsonResponse({ success: true });
}

function setTyping(mobile, isTyping) {
  typingUntil[mobile] = isTyping ? Date.now() + TYPING_THRESHOLD_MS : 0;
  return jsonResponse({ success: true });
}

function backupRow_(row) {
  try {
    getSheet_(BACKUP_SHEET).appendRow(row);
  } catch (err) {}
}

function rowKey_(r) {
  var ts = (r[3] instanceof Date) ? r[3].getTime() : new Date(r[3]).getTime();
  return String(r[0]) + '|' + String(r[1]) + '|' + String(r[2]) + '|' + ts + '|' + String(r[5] || '').length;
}

function clearChat_(user, before) {
  if (!USERS[user]) return jsonResponse({ success: false, message: 'Invalid' });
  var other = otherUser_(user);
  var beforeMs = before ? new Date(before).getTime() : Date.now();
  if (isNaN(beforeMs)) beforeMs = Date.now();

  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ success: true, cleared: 0 });

  var rows = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var backup = getSheet_(BACKUP_SHEET);

  var backupKeys = {};
  var backupLast = backup.getLastRow();
  if (backupLast >= 2) {
    var brows = backup.getRange(2, 1, backupLast - 1, 6).getValues();
    for (var b = 0; b < brows.length; b++) backupKeys[rowKey_(brows[b])] = true;
  }

  var toBackup = [];
  var toDelete = [];

  for (var i = 0; i < rows.length; i++) {
    var from = String(rows[i][0]);
    var to = String(rows[i][1]);
    if (!((from === user && to === other) || (from === other && to === user))) continue;

    var ts = rows[i][3];
    var tsMs = (ts instanceof Date) ? ts.getTime() : new Date(ts).getTime();
    if (isNaN(tsMs) || tsMs > beforeMs) continue;

    var key = rowKey_(rows[i]);
    if (!backupKeys[key]) {
      backupKeys[key] = true;
      toBackup.push(rows[i]);
    }
    toDelete.push(i + 2);
  }

  if (toBackup.length > 0) {
    backup.getRange(backup.getLastRow() + 1, 1, toBackup.length, 6).setValues(toBackup);
  }

  if (toDelete.length > 0) {
    var first = toDelete[0];
    var last = toDelete[toDelete.length - 1];
    if (toDelete.length === last - first + 1) {
      sheet.deleteRange(first, 1, toDelete.length, 6);
    } else {
      for (var d = toDelete.length - 1; d >= 0; d--) sheet.deleteRow(toDelete[d]);
    }
  }

  return jsonResponse({ success: true, cleared: toDelete.length });
}
