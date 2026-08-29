/**
 * 우리들의 마음 우체국 (Rolling Letters) - Google Apps Script Backend
 *
 * ponytail: Minimal Google Sheets CRUD & SHA-256 password protection + Admin Master PIN.
 * No external dependencies, works directly in Google Apps Script V8 runtime.
 */

var SHEET_NAME = 'Letters';
var HEADERS = ['id', 'authorName', 'relationship', 'message', 'color', 'passwordHash', 'likes', 'createdAt'];

// 선생님/관리자 전용 마스터 PIN (원하시는 4자리 번호로 변경 가능합니다)
var ADMIN_MASTER_PIN = '0000';

function _getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
  }
  return sheet;
}

function _hashPassword(password) {
  if (!password || String(password).trim().length < 4) {
    throw new Error('비밀번호는 4자리 이상이어야 합니다.');
  }
  var rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password).trim());
  return Utilities.base64Encode(rawBytes);
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('우리들의 마음 우체국 💌')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getLetters() {
  var sheet = _getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  return data.map(function(row) {
    return {
      id: String(row[0]),
      authorName: String(row[1]),
      relationship: String(row[2]),
      message: String(row[3]),
      color: String(row[4] || '#fff6b8'),
      likes: Number(row[6] || 0),
      createdAt: String(row[7])
    };
  });
}

function addLetter(payload) {
  var authorName = String(payload.authorName || '').trim();
  var relationship = String(payload.relationship || '').trim();
  var message = String(payload.message || '').trim();
  var color = String(payload.color || '#fff6b8').trim();
  var password = String(payload.password || '').trim();

  if (!authorName || authorName.length > 20) throw new Error('이름은 1자 이상 20자 이하이어야 합니다.');
  if (!relationship || relationship.length > 20) throw new Error('관계는 1자 이상 20자 이하이어야 합니다.');
  if (!message || message.length > 500) throw new Error('편지 내용은 1자 이상 500자 이하이어야 합니다.');
  if (password.length !== 4) throw new Error('비밀번호는 4자리여야 합니다.');

  var passwordHash = _hashPassword(password);
  var id = 'letter_' + new Date().getTime() + '_' + Math.random().toString(36).substring(2, 7);
  var createdAt = new Date().toISOString();
  var likes = 0;

  var sheet = _getOrCreateSheet();
  sheet.appendRow([id, authorName, relationship, message, color, passwordHash, likes, createdAt]);

  return {
    id: id,
    authorName: authorName,
    relationship: relationship,
    message: message,
    color: color,
    likes: likes,
    createdAt: createdAt
  };
}

function updateLetter(id, password, updateData) {
  var sheet = _getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) throw new Error('편지를 찾을 수 없습니다.');

  var trimmedPassword = String(password).trim();
  var isAdmin = (trimmedPassword === ADMIN_MASTER_PIN);

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var inputHash = _hashPassword(trimmedPassword);

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var storedHash = String(data[i][5]);
      if (!isAdmin && storedHash !== inputHash) {
        throw new Error('비밀번호가 일치하지 않습니다.');
      }

      var relationship = String(updateData.relationship || data[i][2]).trim().slice(0, 20);
      var message = String(updateData.message || data[i][3]).trim().slice(0, 500);
      var color = String(updateData.color || data[i][4]).trim();

      var rowIdx = i + 2;
      sheet.getRange(rowIdx, 3).setValue(relationship);
      sheet.getRange(rowIdx, 4).setValue(message);
      sheet.getRange(rowIdx, 5).setValue(color);

      return {
        id: id,
        authorName: String(data[i][1]),
        relationship: relationship,
        message: message,
        color: color,
        likes: Number(data[i][6] || 0),
        createdAt: String(data[i][7]),
        isAdmin: isAdmin
      };
    }
  }
  throw new Error('해당 ID의 편지를 찾을 수 없습니다.');
}

function deleteLetter(id, password) {
  var sheet = _getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) throw new Error('편지를 찾을 수 없습니다.');

  var trimmedPassword = String(password).trim();
  var isAdmin = (trimmedPassword === ADMIN_MASTER_PIN);

  var data = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
  var inputHash = _hashPassword(trimmedPassword);

  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var storedHash = String(data[i][5]);
      if (!isAdmin && storedHash !== inputHash) {
        throw new Error('비밀번호가 일치하지 않습니다.');
      }
      sheet.deleteRow(i + 2);
      return { success: true, id: id, isAdmin: isAdmin };
    }
  }
  throw new Error('해당 ID의 편지를 찾을 수 없습니다.');
}

function likeLetter(id) {
  var sheet = _getOrCreateSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) throw new Error('편지를 찾을 수 없습니다.');

  var data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      var rowIdx = i + 2;
      var likesRange = sheet.getRange(rowIdx, 7);
      var currentLikes = Number(likesRange.getValue() || 0);
      var newLikes = currentLikes + 1;
      likesRange.setValue(newLikes);
      return { success: true, id: id, likes: newLikes };
    }
  }
  throw new Error('해당 ID의 편지를 찾을 수 없습니다.');
}
