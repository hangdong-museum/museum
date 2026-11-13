/**
 * 국립중앙박물관 보물찾기 출석체크 — 데이터 서버 (Google Apps Script)
 * - 화면은 GitHub Pages(docs/), 이 스크립트는 시트·드라이브 저장만 담당
 * - 스프레드시트에 바인딩된 스크립트로 사용
 * - 최초 1회 setup() 실행 → 시트/폴더 자동 생성
 */

const SHEET_SETTINGS = '설정';
const SHEET_STUDENTS = '학생';
const SHEET_TREASURES = '보물';
const SHEET_RECORDS = '기록';

const REC_HEADER = ['학번', '이름', '보물', '정답시각', '사진시각', '사진파일ID', '사진링크', '썸네일ID', '저장시각'];

/* ───────── API 진입 (GitHub Pages 화면 → fetch POST) ─────────
 * 요청: { fn: '함수이름', args: [...] }   (Content-Type: text/plain)
 * 응답: { ok: true, data } 또는 { ok: false, error }
 */
const API = {
  login: login,
  checkAnswer: checkAnswer,
  markSaved: markSaved,
  uploadPhoto: uploadPhoto,
  getTreasureImage: getTreasureImage,
  adminData: adminData,
  adminPhoto: adminPhoto,
  adminThumbs: adminThumbs
};

function doPost(e) {
  let res;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API[req.fn];
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    res = { ok: true, data: fn.apply(null, req.args || []) };
  } catch (err) {
    res = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, msg: '박물관 보물찾기 API 동작 중' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────── 초기 설정 ───────── */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sh = ss.getSheetByName(SHEET_SETTINGS);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SETTINGS);
    sh.getRange(1, 1, 4, 2).setValues([
      ['항목', '값'],
      ['공용코드', '1234'],
      ['관리자코드', 'admin1234'],
      ['사진폴더ID', '']
    ]);
    sh.setColumnWidth(1, 120); sh.setColumnWidth(2, 320);
  }

  if (!ss.getSheetByName(SHEET_STUDENTS)) {
    const s = ss.insertSheet(SHEET_STUDENTS);
    s.getRange(1, 1, 1, 4).setValues([['학년', '반', '번호', '이름']]);
  }

  if (!ss.getSheetByName(SHEET_TREASURES)) {
    const s = ss.insertSheet(SHEET_TREASURES);
    const rows = [
      ['번호', '위치', '힌트1', '힌트2', '힌트3', '힌트4', '질문', '정답(|로 구분)', '이미지(드라이브ID/URL, 비우면 기본)'],
      [1, '상설전시관 2층 「사유의 방」',
        '나는 약 1,400년 전 삼국 시대에 태어났어요. 청동으로 몸을 만들고 그 위에 금을 입혔지요.',
        '한쪽 다리를 다른 쪽 무릎 위에 걸치고 앉아, 손가락 끝을 살짝 뺨에 대고 있어요.',
        '입가에는 은은한 미소가 번져 있어요. 먼 미래에 부처가 되어 사람들을 구할 방법을 곰곰이 생각하는 중이랍니다.',
        '',
        '이 유물의 이름은? (금동미륵보살 OOOOO)', '반가사유상|금동미륵보살반가사유상', ''],
      [2, '상설전시관 1층 「역사의 길」 끝',
        '나는 고려 시대인 1348년에 세워진 탑이에요. 높이가 13m가 넘는 거대한 탑이지요.',
        '다른 돌탑들은 대부분 화강암인데, 나는 하얗고 부드러운 대리석으로 만들어졌어요.',
        '몸에는 부처와 보살, 불교 이야기 장면이 빼곡하게 새겨져 있어요.',
        '1907년 일본으로 무단으로 반출되었다가, 다행히 돌아왔어요.',
        '이 유물의 이름은? (OOOO 10층석탑)', '경천사지|경천사지10층석탑', ''],
      [3, '상설전시관 1층 「역사의 길」',
        '나는 진짜 돌이 아니에요. 최신 디지털 기술로 실제 크기 그대로 되살려 낸 비석이에요.',
        '진짜 비석은 지금 중국 지린성 지안(집안)에 서 있어요. 높이 약 6.4m, 무게 약 37톤이나 되는 거대한 돌이지요.',
        '414년, 아들 장수왕이 돌아가신 아버지의 업적을 기리기 위해 세웠어요. 네 면에 1,775자의 글자가 새겨져 있어요.',
        '비석의 주인공은 영토를 크게 넓혀 이름에도 그 뜻이 담긴 고구려의 왕이에요.',
        '이 유물의 이름은? (디지털 OOOOOOO)', '광개토대왕릉비|디지털광개토대왕릉비', ''],
      [4, '상설전시관 1층 선사·고대관 「신라실」',
        '나는 경주의 커다란 신라 무덤 속에서 발견되었어요. 온몸이 황금으로 되어 있지요.',
        '머리에 쓰는 물건이에요. 둥근 테 위로 한자 \'出(출)\' 모양 장식과 사슴뿔 모양 장식이 솟아 있어요.',
        '초록빛 곱은옥(굽은 옥)과 작고 동그란 금 조각(달개)이 잔뜩 매달려 있어서, 움직이면 반짝반짝 흔들려요.',
        '신라가 \'황금의 나라\'라고 불린 이유를 보여 주는 대표 유물이에요.',
        '이 유물의 이름은? (신라 OO)', '금관|신라금관', ''],
      [5, '상설전시관 3층 도자공예관 「청자실」',
        '나는 고려 시대(12세기)에 만들어졌어요. 비취처럼 맑은 푸른빛, \'비색\'을 띠고 있지요.',
        '향을 피우는 도구예요. 뚜껑의 구멍 사이로 향기로운 연기가 솔솔 빠져나와요.',
        '뚜껑에는 흙을 뚫어서 무늬를 만드는 \'투각\' 기법으로 동그란 고리가 이어진 무늬를 새겼어요.',
        '몸통은 연꽃잎이 감싸고 있고, 맨 아래에서는 귀여운 토끼 세 마리가 나를 받치고 있어요.',
        '이 유물의 이름은? (청자 투각 OOOOOO)', '칠보무늬향로|칠보문향로|청자투각칠보무늬향로', ''],
      [6, '상설전시관 2층 서화관',
        '나는 책이에요. 조선 왕실의 결혼식, 장례식, 잔치 같은 중요한 행사를 글과 그림으로 꼼꼼하게 기록했지요.',
        '행사에 참여한 사람들의 행렬을 그린 그림(반차도)이 특히 유명해요.',
        '정조가 강화도에 세운 왕실 도서관에 보관되어 있었는데, 1866년 병인양요 때 프랑스군이 가져갔어요.',
        '오랜 노력 끝에 2011년, 145년 만에 우리나라로 돌아왔어요.',
        '이 유물의 이름은? (외규장각 OO)', '의궤|외규장각의궤', '']
    ];
    s.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    s.setColumnWidths(3, 4, 260);
  }

  if (!ss.getSheetByName(SHEET_RECORDS)) {
    const s = ss.insertSheet(SHEET_RECORDS);
    s.getRange(1, 1, 1, REC_HEADER.length).setValues([REC_HEADER]);
    s.getRange('A:A').setNumberFormat('@');
    s.setFrozenRows(1);
  }

  getPhotoFolder_();
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('시트1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);
}

/* ───────── 내부 유틸 ───────── */
function getSettings_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  const v = sh.getDataRange().getValues();
  const o = {};
  v.slice(1).forEach(r => { o[String(r[0]).trim()] = String(r[1]).trim(); });
  return o;
}

function getPhotoFolder_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_SETTINGS);
  const v = sh.getDataRange().getValues();
  let row = v.findIndex(r => String(r[0]).trim() === '사진폴더ID');
  const id = row >= 0 ? String(v[row][1]).trim() : '';
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  const folder = DriveApp.createFolder('박물관 보물찾기 사진_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyyMMdd'));
  if (row < 0) { sh.appendRow(['사진폴더ID', folder.getId()]); }
  else { sh.getRange(row + 1, 2).setValue(folder.getId()); }
  return folder;
}

function getSubFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function norm_(s) {
  return String(s || '').replace(/\s+/g, '').trim();
}

/** 학생 탭(학년, 반, 번호, 이름) → [{sid, name, cls}]  학번 = 학년 + 반2자리 + 번호2자리 */
function getStudents_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_STUDENTS);
  const num = x => parseInt(String(x).replace(/[^0-9]/g, ''), 10);
  const pad = n => (n < 10 ? '0' : '') + n;
  return sh.getDataRange().getDisplayValues().slice(1)
    .map(r => ({ g: num(r[0]), c: num(r[1]), n: num(r[2]), name: norm_(r[3]) }))
    .filter(r => r.g && r.c && r.n && r.name)
    .map(r => ({ sid: r.g + pad(r.c) + pad(r.n), name: r.name, cls: r.g + '-' + r.c }));
}

function auth_(sid, name, code) {
  const st = getSettings_();
  if (norm_(code) !== norm_(st['공용코드'])) throw new Error('코드가 올바르지 않습니다.');
  const hit = getStudents_().find(s => s.sid === norm_(sid) && s.name === norm_(name));
  if (!hit) throw new Error('학번과 이름을 확인해 주세요.');
  return { sid: hit.sid, name: hit.name };
}

function getTreasureRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_TREASURES);
  return sh.getDataRange().getDisplayValues().slice(1).filter(r => String(r[0]).trim() !== '');
}

function readRecords_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_RECORDS);
  if (sh.getMaxColumns() < REC_HEADER.length) sh.insertColumnsAfter(sh.getMaxColumns(), REC_HEADER.length - sh.getMaxColumns());
  if (!sh.getRange(1, REC_HEADER.length).getValue()) sh.getRange(1, 1, 1, REC_HEADER.length).setValues([REC_HEADER]);
  const last = sh.getLastRow();
  if (last < 2) return { sh, rows: [] };
  return { sh, rows: sh.getRange(2, 1, last - 1, REC_HEADER.length).getDisplayValues() };
}

function studentStatus_(sid) {
  const { rows } = readRecords_();
  const o = {};
  rows.forEach(r => {
    if (norm_(r[0]) !== sid) return;
    o[String(r[2])] = { answered: !!r[3], photo: !!r[5], photoTime: r[4] };
  });
  return o;
}

/** 기록 upsert (학번+보물 기준 한 행) */
function upsertRecord_(sid, name, no, patch) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const { sh, rows } = readRecords_();
    let idx = rows.findIndex(r => norm_(r[0]) === sid && String(r[2]) === String(no));
    let rec;
    if (idx < 0) {
      rec = [sid, name, String(no), '', '', '', '', '', ''];
    } else {
      rec = rows[idx].slice();
    }
    const old = { fileId: rec[5], thumbId: rec[7] };
    if (patch.answerTime) rec[3] = patch.answerTime;
    if (patch.photoTime) { rec[4] = patch.photoTime; if (!rec[8]) rec[8] = patch.photoTime; }
    if (patch.fileId) rec[5] = patch.fileId;
    if (patch.url) rec[6] = patch.url;
    if (patch.thumbId !== undefined) rec[7] = patch.thumbId;
    if (patch.savedTime && !rec[8]) rec[8] = patch.savedTime;
    if (idx < 0) {
      sh.appendRow(rec);
      sh.getRange(sh.getLastRow(), 1).setNumberFormat('@').setValue(sid);
    } else {
      sh.getRange(idx + 2, 1, 1, rec.length).setValues([rec]);
    }
    return old;
  } finally {
    lock.releaseLock();
  }
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

/* ───────── 학생 API ───────── */
function login(sid, name, code) {
  const u = auth_(sid, name, code);
  const treasures = getTreasureRows_().map(r => ({
    no: String(r[0]),
    location: r[1],
    hints: [r[2], r[3], r[4], r[5]].filter(h => String(h).trim() !== ''),
    question: r[6],
    image: String(r[8] || '').trim()
  }));
  return { user: u, treasures, status: studentStatus_(u.sid) };
}

function checkAnswer(sid, name, code, no, answer) {
  const u = auth_(sid, name, code);
  const row = getTreasureRows_().find(r => String(r[0]) === String(no));
  if (!row) throw new Error('보물 정보를 찾을 수 없습니다.');
  const answers = String(row[7]).split('|').map(norm_).filter(Boolean);
  const ok = answers.indexOf(norm_(answer)) >= 0;
  if (ok) upsertRecord_(u.sid, u.name, no, { answerTime: now_() });
  return { ok, status: studentStatus_(u.sid) };
}

function uploadPhoto(sid, name, code, no, dataUrl, thumbUrl, savedAt) {
  const u = auth_(sid, name, code);
  const st = studentStatus_(u.sid)[String(no)];
  if (!st || !st.answered) throw new Error('먼저 유물 이름을 맞혀 주세요.');

  const m = String(dataUrl).match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (!m) throw new Error('사진 형식이 올바르지 않습니다.');
  // 사진시각 = 학생 기기에서 사진을 저장한 시각 (없거나 형식이 다르면 업로드 시각)
  const time = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(String(savedAt || '')) ? String(savedAt) : now_();
  const stamp = time.replace(/[-: ]/g, '').slice(4);
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1],
    u.sid + '_' + u.name + '_보물' + no + '_' + stamp + '.jpg');

  const root = getPhotoFolder_();
  const file = getSubFolder_(root, '보물' + no).createFile(blob);

  let thumbId = '';
  const tm = String(thumbUrl || '').match(/^data:(image\/[a-z]+);base64,(.+)$/);
  if (tm) {
    const tb = Utilities.newBlob(Utilities.base64Decode(tm[2]), tm[1], 't_' + blob.getName());
    thumbId = getSubFolder_(root, '_썸네일').createFile(tb).getId();
  }

  const old = upsertRecord_(u.sid, u.name, no, { photoTime: time, fileId: file.getId(), url: file.getUrl(), thumbId });
  [old.fileId, old.thumbId].forEach(id => {
    if (id && id !== file.getId() && id !== thumbId) {
      try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {}
    }
  });
  return { ok: true, status: studentStatus_(u.sid) };
}

/** 사진을 휴대폰에 저장했다는 사실만 기록 (사진은 보내지 않음) */
function markSaved(sid, name, code, no, savedAt) {
  const u = auth_(sid, name, code);
  const st = studentStatus_(u.sid)[String(no)];
  if (!st || !st.answered) throw new Error('먼저 유물 이름을 맞혀 주세요.');
  const time = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(String(savedAt || '')) ? String(savedAt) : now_();
  upsertRecord_(u.sid, u.name, no, { savedTime: time });
  return { ok: true };
}

/** 보물 이미지(드라이브 파일ID) → dataURL */
function getTreasureImage(fileId) {
  const cache = CacheService.getScriptCache();
  const key = 'timg_' + fileId;
  const c = cache.get(key);
  if (c) return c;
  const blob = DriveApp.getFileById(fileId).getBlob();
  const url = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  if (url.length < 95000) cache.put(key, url, 21600);
  return url;
}

/* ───────── 관리자 API ───────── */
function adminAuth_(code) {
  if (norm_(code) !== norm_(getSettings_()['관리자코드'])) throw new Error('관리자 코드가 올바르지 않습니다.');
}

function adminData(code) {
  adminAuth_(code);
  const students = getStudents_();
  const treasures = getTreasureRows_().map(r => String(r[0]));
  const rec = {};
  readRecords_().rows.forEach(r => {
    rec[norm_(r[0]) + '_' + r[2]] = { a: r[3], p: r[4], f: r[5], t: r[7], s: r[8] };
  });
  return { students, treasures, rec };
}

function adminPhoto(code, fileId, thumb) {
  adminAuth_(code);
  const file = DriveApp.getFileById(fileId);
  let blob = thumb ? file.getThumbnail() : null;
  if (!blob) blob = file.getBlob();
  return 'data:' + (blob.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(blob.getBytes());
}

/** 썸네일 여러 장 한 번에 (최대 10장) → {fileId: dataURL 또는 ''}
 *  - 업로드 때 만든 썸네일 파일(작은 파일)은 그대로 전송
 *  - 썸네일 파일이 없는 예전 기록은 드라이브 썸네일 → 없으면 원본 사진 전송 */
function adminThumbs(code, ids) {
  adminAuth_(code);
  const out = {};
  (ids || []).slice(0, 10).forEach(id => {
    try {
      const f = DriveApp.getFileById(id);
      let b = null;
      if (f.getSize() < 150000) {
        b = f.getBlob();
      } else {
        try { b = f.getThumbnail(); } catch (e) { b = null; }
        if (!b || b.getBytes().length < 2000) b = f.getBlob();
      }
      out[id] = 'data:' + (b.getContentType() || 'image/jpeg') + ';base64,' + Utilities.base64Encode(b.getBytes());
    } catch (e) { out[id] = ''; }
  });
  return out;
}

/* ───────── 활동 종료 후 정리 (편집기에서 직접 실행) ───────── */
function deleteAllPhotos() {
  const folder = getPhotoFolder_();
  folder.setTrashed(true);
  const { sh } = readRecords_();
  const last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 5, last - 1, 5).clearContent();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const set = ss.getSheetByName(SHEET_SETTINGS);
  const v = set.getDataRange().getValues();
  const row = v.findIndex(r => String(r[0]).trim() === '사진폴더ID');
  if (row >= 0) set.getRange(row + 1, 2).setValue('');
}
