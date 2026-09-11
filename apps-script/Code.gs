const CONFIG = {
  SHEET_ID: '1d2I68enDasYwc1EtzYV6gUxhizc22o6rCnFyyYat8JQ',
  SHEET_NAME: '제출현황',
  ROOT_FOLDER_ID: '185rf4cqBBUvTfkbQpZBKsU72QLOoghq-',
  ADMIN_SHA256: '802e1ebd816154c22ed82d4bdbddab71a0e4495715f4644243bcaf44e37ca20c',
  TIMEZONE: 'Asia/Seoul'
};

function doGet(e) {
  const p = (e && e.parameter) || {};
  const action = p.action || 'ping';
  let data;

  try {
    if (action === 'ping') {
      data = { ok: true, service: 'papyrus-mangwon', time: nowText_() };
    } else if (action === 'progress') {
      data = getProgress_(String(p.teamCode || '').trim().toUpperCase());
    } else if (action === 'admin') {
      if (!isAdminCode_(String(p.adminCode || ''))) {
        data = { ok: false, error: '관리자 코드가 올바르지 않습니다.' };
      } else {
        data = getAdminData_();
      }
    } else {
      data = { ok: false, error: '지원하지 않는 요청입니다.' };
    }
  } catch (err) {
    data = { ok: false, error: String(err && err.message ? err.message : err) };
  }

  return output_(data, p.callback);
}

function doPost(e) {
  let body = {};
  try {
    body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return output_({ ok: false, error: '요청 형식을 읽을 수 없습니다.' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (body.action === 'createTeam') return output_(createTeam_(body));
    if (body.action === 'visit') return output_(saveVisit_(body));
    return output_({ ok: false, error: '지원하지 않는 요청입니다.' });
  } catch (err) {
    return output_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function createTeam_(body) {
  const teamCode = normalizeTeamCode_(body.teamCode);
  const members = Array.isArray(body.members) ? body.members : [];
  if (!teamCode) throw new Error('유효한 모둠 코드가 아닙니다.');
  if (members.length < 2 || members.length > 5) throw new Error('모둠원은 2~5명이어야 합니다.');

  const cleanMembers = members.map((m, i) => {
    const sid = String(m.sid || '').trim();
    const name = String(m.name || '').trim();
    if (!sid || !name) throw new Error((i + 1) + '번 모둠원의 학번과 이름을 확인하세요.');
    return { sid, name };
  });

  const sheet = sheet_();
  const existing = findTeamRow_(sheet, teamCode);
  if (existing) {
    return { ok: true, teamCode, row: existing, alreadyExists: true, progress: progressFromRow_(sheet, existing) };
  }

  const now = nowText_();
  sheet.appendRow([
    teamCode,
    cleanMembers.map(m => m.sid + ' ' + m.name).join(' · '),
    '0/3',
    '', '', '',
    '', '', '',
    '', '', '',
    '진행 전',
    now,
    now
  ]);

  const folder = getOrCreateTeamFolder_(teamCode);
  return { ok: true, teamCode, row: sheet.getLastRow(), folderId: folder.getId(), progress: 0 };
}

function saveVisit_(body) {
  const teamCode = normalizeTeamCode_(body.teamCode);
  const round = Number(body.round);
  const shopName = String(body.shopName || '').trim();
  const photoData = String(body.photoData || '');

  if (!teamCode) throw new Error('유효한 모둠 코드가 아닙니다.');
  if (![1, 2, 3].includes(round)) throw new Error('인증 차수는 1~3차만 가능합니다.');
  if (!shopName) throw new Error('책방 이름이 없습니다.');
  if (!photoData.startsWith('data:image/')) throw new Error('인증사진 형식이 올바르지 않습니다.');

  const sheet = sheet_();
  const row = findTeamRow_(sheet, teamCode);
  if (!row) throw new Error('먼저 모둠을 생성하세요.');

  const shopCol = 4 + (round - 1) * 3;
  const photoCol = shopCol + 1;
  const timeCol = shopCol + 2;
  const currentShop = String(sheet.getRange(row, shopCol).getValue() || '');
  const currentPhoto = String(sheet.getRange(row, photoCol).getValue() || '');

  if (currentShop || currentPhoto) {
    if (currentShop === shopName && currentPhoto) {
      return { ok: true, alreadyExists: true, teamCode, round, progress: progressFromRow_(sheet, row) };
    }
    throw new Error(round + '차 인증은 이미 제출되었습니다.');
  }

  if (round > 1) {
    const previousShopCol = 4 + (round - 2) * 3;
    if (!String(sheet.getRange(row, previousShopCol).getValue() || '')) {
      throw new Error((round - 1) + '차 인증을 먼저 완료하세요.');
    }
  }

  const parsed = parseDataUrl_(photoData);
  const folder = getOrCreateTeamFolder_(teamCode);
  const ext = extensionForMime_(parsed.mimeType);
  const safeShop = shopName.replace(/[\\/:*?"<>|]/g, '_').slice(0, 40);
  const stamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
  const file = folder.createFile(Utilities.newBlob(parsed.bytes, parsed.mimeType, round + '차_' + safeShop + '_' + stamp + '.' + ext));
  const fileUrl = file.getUrl();
  const now = nowText_();

  sheet.getRange(row, shopCol, 1, 3).setValues([[shopName, fileUrl, now]]);
  const progress = progressFromRow_(sheet, row);
  sheet.getRange(row, 3).setValue(progress + '/3');
  sheet.getRange(row, 13).setValue(progress === 3 ? '완료' : '진행 중');
  sheet.getRange(row, 15).setValue(now);

  return {
    ok: true,
    teamCode,
    round,
    progress,
    fileId: file.getId(),
    fileUrl
  };
}

function getProgress_(teamCode) {
  teamCode = normalizeTeamCode_(teamCode);
  if (!teamCode) return { ok: false, error: '유효한 모둠 코드가 아닙니다.' };
  const sheet = sheet_();
  const row = findTeamRow_(sheet, teamCode);
  if (!row) return { ok: false, error: '등록되지 않은 모둠입니다.' };
  const values = sheet.getRange(row, 1, 1, 15).getDisplayValues()[0];
  const visits = [
    { round: 1, shop: values[3], submittedAt: values[5] },
    { round: 2, shop: values[6], submittedAt: values[8] },
    { round: 3, shop: values[9], submittedAt: values[11] }
  ].filter(v => v.shop);
  return { ok: true, teamCode, progress: visits.length, visits };
}

function getAdminData_() {
  const sheet = sheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, teams: [], totals: { teams: 0, inProgress: 0, done: 0 } };

  const rows = sheet.getRange(2, 1, lastRow - 1, 15).getDisplayValues();
  const teams = rows.filter(r => r[0]).map(r => ({
    teamCode: r[0],
    members: r[1],
    progressText: r[2],
    progress: Number(String(r[2]).split('/')[0]) || 0,
    visits: [
      photoRecord_(1, r[3], r[4], r[5]),
      photoRecord_(2, r[6], r[7], r[8]),
      photoRecord_(3, r[9], r[10], r[11])
    ].filter(Boolean),
    status: r[12],
    createdAt: r[13],
    updatedAt: r[14]
  }));

  return {
    ok: true,
    teams,
    totals: {
      teams: teams.length,
      inProgress: teams.filter(t => t.progress > 0 && t.progress < 3).length,
      done: teams.filter(t => t.progress === 3).length
    }
  };
}

function photoRecord_(round, shop, fileUrl, submittedAt) {
  if (!shop) return null;
  const fileId = extractDriveFileId_(fileUrl);
  return {
    round,
    shop,
    fileUrl,
    fileId,
    thumbnailUrl: fileId ? 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(fileId) + '&sz=w800' : '',
    submittedAt
  };
}

function sheet_() {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) throw new Error('제출현황 시트를 찾을 수 없습니다.');
  return sheet;
}

function findTeamRow_(sheet, teamCode) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const finder = sheet.getRange(2, 1, lastRow - 1, 1).createTextFinder(teamCode).matchEntireCell(true);
  const cell = finder.findNext();
  return cell ? cell.getRow() : 0;
}

function progressFromRow_(sheet, row) {
  const shops = sheet.getRange(row, 4, 1, 9).getDisplayValues()[0];
  return [shops[0], shops[3], shops[6]].filter(Boolean).length;
}

function getOrCreateTeamFolder_(teamCode) {
  const root = DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const existing = root.getFoldersByName(teamCode);
  return existing.hasNext() ? existing.next() : root.createFolder(teamCode);
}

function parseDataUrl_(dataUrl) {
  const match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) throw new Error('사진 데이터를 읽을 수 없습니다.');
  return { mimeType: match[1], bytes: Utilities.base64Decode(match[2]) };
}

function extensionForMime_(mime) {
  if (mime === 'image/png') return 'png';
  if (mime === 'image/webp') return 'webp';
  if (mime === 'image/heic' || mime === 'image/heif') return 'heic';
  return 'jpg';
}

function extractDriveFileId_(url) {
  const m = String(url || '').match(/\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : '';
}

function normalizeTeamCode_(code) {
  const v = String(code || '').trim().toUpperCase();
  return /^PAP-[A-HJ-NP-Z2-9]{4}$/.test(v) ? v : '';
}

function isAdminCode_(code) {
  return sha256Hex_(String(code || '')) === CONFIG.ADMIN_SHA256;
}

function sha256Hex_(text) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2))
    .join('');
}

function nowText_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function output_(data, callback) {
  const json = JSON.stringify(data);
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
