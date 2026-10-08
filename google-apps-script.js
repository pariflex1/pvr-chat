// Google Apps Script for Chat App Backend
// Deploy as Web App: Execute as "Me", Access "Anyone"
// Single sheet: Messages
// Images stored in Google Drive folder: PVRChat_Images

const MESSAGES_SHEET = 'Messages';
const HEADERS = ['From', 'To', 'Message', 'Timestamp', 'Status', 'ImageUrl'];
const DRIVE_FOLDER = 'PVRChat_Images';

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

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MESSAGES_SHEET);

  if (sheet) {
    var headers = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    var match = true;
    for (var h = 0; h < HEADERS.length; h++) {
      if (headers[h] !== HEADERS[h]) { match = false; break; }
    }
    if (!match) {
      ss.deleteSheet(sheet);
      sheet = null;
    }
  }

  if (!sheet) {
    sheet = ss.insertSheet(MESSAGES_SHEET);
    sheet.appendRow(HEADERS);
    sheet.getRange('A:B').setNumberFormat('@');
  }

  return sheet;
}

function getDriveFolder_() {
  var folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(DRIVE_FOLDER);
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
    if (p.action === 'getImage' && p.id) {
      return getImage(String(p.id));
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
        return sendImage(String(data.from), String(data.to), String(data.imageData), String(data.mimeType));
      case 'heartbeat':
        return heartbeat(String(data.mobile));
      case 'setTyping':
        return setTyping(String(data.mobile), data.isTyping);
      case 'markRead':
        return markRead(String(data.mobile));
      case 'clearChat':
        return clearChat(String(data.mobile));
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
  var sheet = getSheet_();
  sheet.appendRow([from, to, message, new Date(), 'sent', '']);
  typingUntil[from] = 0;
  return jsonResponse({ success: true });
}

function sendImage(from, to, base64Data, mimeType) {
  if (!USERS[from] || !USERS[to] || !base64Data) {
    return jsonResponse({ success: false, message: 'Invalid' });
  }

  try {
    var folder = getDriveFolder_();
    var decoded = Utilities.base64Decode(base64Data);
    var ext = 'webp';
    if (mimeType.indexOf('png') !== -1) ext = 'png';
    else if (mimeType.indexOf('jpeg') !== -1 || mimeType.indexOf('jpg') !== -1) ext = 'jpg';

    var blob = Utilities.newBlob(decoded, mimeType, 'img_' + new Date().getTime() + '.' + ext);
    var file = folder.createFile(blob);

    // make publicly viewable
    file.setSharing(DriftApp.Access.ANYONE_WITH_LINK, DriftApp.Permission.VIEW);

    var fileId = file.getId();
    var url = 'https://drive.google.com/uc?export=view&id=' + fileId;

    var sheet = getSheet_();
    sheet.appendRow([from, to, '', new Date(), 'sent', url]);
    typingUntil[from] = 0;

    return jsonResponse({ success: true, url: url });
  } catch (err) {
    return jsonResponse({ success: false, message: 'Upload error: ' + err.toString() });
  }
}

function getImage(fileId) {
  try {
    var file = DriveApp.getFileById(fileId);
    var blob = file.getBlob();
    return ContentService
      .createOutput(blob.getBytes(), blob.getContentType())
      .setMimeType(blob.getContentType());
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
      messages.push({
        from: from,
        to: to,
        message: String(rows[i][2]),
        timestamp: (ts instanceof Date) ? ts.toISOString() : String(ts),
        status: status,
        imageUrl: rows[i][5] ? String(rows[i][5]) : ''
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

function clearChat(mobile) {
  if (!USERS[mobile]) return jsonResponse({ success: false, message: 'Invalid' });
  var sheet = getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }
  return jsonResponse({ success: true });
}
