/**
 * 경일고 교무실 현황판 - 변동 시간표 중계
 *
 * 구글 드라이브 「현황판 시간표」 폴더에서 날짜 이름의 시간표 파일을 찾아 현황판에 넘겨 줍니다.
 *   - 파일 이름에 날짜가 들어 있으면 됩니다: 11-23.xlsx, 1123.xlsx, 11.23, 2026-11-23, 11월23일 …
 *   - 양식은 「전체시간표(2026).xlsx」의 '전체시간표' 시트와 같아야 합니다.
 *     (기본 시간표를 복사해 바뀐 칸만 고친 뒤 날짜 이름으로 저장하면 됩니다)
 *   - 엑셀(.xlsx)과 구글 시트 모두 됩니다. 같은 날짜 파일이 여럿이면 가장 나중에 고친 파일을 씁니다.
 *
 * 설치
 *   1. 아래 FOLDER_ID 에 드라이브 폴더 주소의 마지막 부분을 넣습니다.
 *      (https://drive.google.com/drive/folders/XXXX  ← XXXX 부분)
 *   2. 왼쪽 [서비스] + 를 눌러 "Drive API" 를 추가합니다. (엑셀 파일을 읽는 데 필요)
 *   3. 함수 목록에서 test 를 골라 실행 → 권한 승인 → 실행 로그 확인
 *   4. [배포] > [새 배포] > 유형 "웹 앱", 실행 계정 "나", 액세스 권한 "모든 사용자" → .../exec 주소 복사
 *   5. index.html 의 RELAY_URL 에 그 주소를 붙여넣습니다.
 *   코드를 고친 뒤에는 [배포 관리] > 수정 > 버전 "새 버전" 으로 다시 배포해야 반영됩니다.
 */
var FOLDER_ID = '';
var DOW = ['일', '월', '화', '수', '목', '금', '토'];

function doGet(e) {
  var out;
  try {
    var date = e && e.parameter && e.parameter.date;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date || '')) date = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
    out = getChange_(date);
  } catch (err) {
    out = { fatal: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// 편집기에서 실행해 결과를 확인합니다. 날짜를 바꿔 시험할 수 있습니다.
function test() {
  var date = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var r = getChange_(date);
  Logger.log('날짜: ' + date);
  Logger.log('폴더의 날짜 파일: ' + JSON.stringify(r.upcoming));
  if (r.error) Logger.log('오류: ' + r.error);
  if (r.cells) Logger.log(r.name + ' → ' + r.cells.length + '칸 읽음. 예: ' + JSON.stringify(r.cells.slice(0, 3)));
  else if (!r.error) Logger.log('오늘 날짜 파일 없음 → 현황판은 기본 시간표를 표시');
}

/* ---------- 폴더 살피기 ---------- */
function listDated_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get('list');
  if (hit) return JSON.parse(hit);
  if (!FOLDER_ID) throw new Error('relay.gs 의 FOLDER_ID 가 비어 있습니다');
  var files = DriveApp.getFolderById(FOLDER_ID).getFiles(), byDate = {};
  var year = +Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy');
  while (files.hasNext()) {
    var f = files.next();
    if (f.isTrashed()) continue;
    var mime = f.getMimeType();
    if (mime !== MimeType.GOOGLE_SHEETS && !/spreadsheetml|ms-excel/.test(mime)) continue;
    var d = dateFromName_(f.getName(), year);
    if (!d) continue;
    var t = f.getLastUpdated().getTime();
    if (!byDate[d] || byDate[d].updated < t) byDate[d] = { id: f.getId(), name: f.getName(), mime: mime, updated: t };
  }
  cache.put('list', JSON.stringify(byDate), 50);
  return byDate;
}

function dateFromName_(name, year) {
  var s = name.replace(/\.(xlsx|xlsm|xls)$/i, ''), m, y, mo, d;
  if ((m = /(20\d{2})\s*[-.\/년]?\s*(\d{1,2})\s*[-.\/월]\s*(\d{1,2})/.exec(s)) || (m = /(20\d{2})(\d{2})(\d{2})/.exec(s))) {
    y = +m[1]; mo = +m[2]; d = +m[3];
  } else if ((m = /(\d{1,2})\s*[-.\/월]\s*(\d{1,2})/.exec(s)) || (m = /(?:^|\D)(\d{2})(\d{2})(?:\D|$)/.exec(s))) {
    mo = +m[1]; d = +m[2]; y = year;
    var now = new Date(), cand = new Date(y, mo - 1, d);
    if (cand - now > 180 * 864e5) y--;            // 연말·연초 보정
    if (now - cand > 180 * 864e5) y++;
  } else return null;
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  return y + '-' + ('0' + mo).slice(-2) + '-' + ('0' + d).slice(-2);
}

/* ---------- 변동 시간표 읽기 ---------- */
function getChange_(date) {
  var byDate = listDated_();
  var today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');
  var upcoming = Object.keys(byDate).filter(function (d) { return d >= today; }).sort();
  var f = byDate[date];
  if (!f) return { cells: null, upcoming: upcoming };

  var cache = CacheService.getScriptCache(), key = 'g:' + f.id + ':' + f.updated;
  var hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  var out;
  try {
    var values = readValues_(f);
    var wd = DOW[new Date(date + 'T12:00:00+09:00').getDay()];
    var parsed = parseGrid_(values, wd);
    out = parsed.error
      ? { cells: null, name: f.name, updated: f.updated, upcoming: upcoming, error: f.name + ': ' + parsed.error }
      : { cells: parsed.cells, name: f.name, updated: f.updated, upcoming: upcoming };
  } catch (err) {
    out = { cells: null, name: f.name, updated: f.updated, upcoming: upcoming, error: f.name + ' 파일을 읽지 못했습니다' };
    Logger.log(err);
  }
  cache.put(key, JSON.stringify(out), 21600);
  return out;
}

function readValues_(f) {
  if (f.mime === MimeType.GOOGLE_SHEETS) return pickSheet_(SpreadsheetApp.openById(f.id));
  // 엑셀은 잠시 구글 시트로 바꿔 읽고 바로 지웁니다
  if (typeof Drive === 'undefined') throw new Error('Drive API 서비스가 추가되지 않았습니다');
  var tmp = Drive.Files.create({ name: '_현황판_임시_' + f.name, mimeType: MimeType.GOOGLE_SHEETS },
                               DriveApp.getFileById(f.id).getBlob());
  try { return pickSheet_(SpreadsheetApp.openById(tmp.id)); }
  finally { DriveApp.getFileById(tmp.id).setTrashed(true); }
}

function pickSheet_(ss) {
  var sh = ss.getSheetByName('전체시간표') || ss.getSheets()[0];
  return sh.getDataRange().getDisplayValues();
}

/**
 * 표 읽기. values: 2차원 글자 배열(병합 칸은 맨 왼쪽 위에만 값이 있음)
 * 반환: { cells: [[학년, 반, 교시, 과목, 교사], ...] } 또는 { error }
 */
function parseGrid_(values, weekday) {
  var norm = function (v) { return String(v == null ? '' : v).replace(/\r/g, '').trim(); };
  var hdr = -1, i, j;
  for (i = 0; i < values.length && hdr < 0; i++) {
    var n = 0;
    for (j = 0; j < values[i].length; j++) if (/^[1-3]\s*학년$/.test(norm(values[i][j]))) n++;
    if (n >= 2) hdr = i;
  }
  if (hdr < 0 || hdr + 1 >= values.length) return { error: '학년 머리줄을 찾지 못했습니다' };
  var gRow = values[hdr], cRow = values[hdr + 1];
  var dayCol = -1, perCol = -1;
  for (j = 0; j < gRow.length; j++) {
    if (norm(gRow[j]) === '요일' || norm(cRow[j]) === '요일') dayCol = j;
    if (norm(gRow[j]) === '교시' || norm(cRow[j]) === '교시') perCol = j;
  }
  if (perCol < 0) return { error: '교시 열을 찾지 못했습니다' };
  var cols = [], grade = null;
  for (j = 0; j < gRow.length; j++) {
    var gm = /^([1-3])\s*학년$/.exec(norm(gRow[j]));
    if (gm) grade = +gm[1];
    var cn = parseInt(norm(cRow[j]), 10);
    if (grade && j !== perCol && j !== dayCol && cn > 0) cols.push({ j: j, g: grade, c: cn });
  }
  if (!cols.length) return { error: '반 번호 줄을 찾지 못했습니다' };

  var rows = [], day = '', anyDay = false;
  for (i = hdr + 2; i < values.length; i++) {
    if (dayCol >= 0) {
      var dv = norm(values[i][dayCol]).charAt(0);
      if (DOW.indexOf(dv) >= 0) { day = dv; anyDay = true; }
    }
    var p = parseInt(norm(values[i][perCol]), 10);
    if (!(p >= 1 && p <= 10)) continue;
    rows.push({ day: day, p: p, r: values[i] });
  }
  if (anyDay) rows = rows.filter(function (x) { return x.day === weekday; });
  if (!rows.length) return { error: anyDay ? weekday + '요일 줄이 없습니다' : '교시 줄이 없습니다' };

  var cells = [];
  rows.forEach(function (x) {
    cols.forEach(function (c) {
      var sp = splitCell_(norm(x.r[c.j]));
      cells.push([c.g, c.c, x.p, sp[0], sp[1]]);
    });
  });
  return { cells: cells };
}

function splitCell_(t) {
  if (!t) return ['', ''];
  var lines = t.split(/\s*\n\s*/);
  if (lines.length >= 2) return [lines[0], lines.slice(1).join(' ')];
  var m = /^(.*\S)\s+([가-힣]{2,4})$/.exec(t);
  return m ? [m[1], m[2]] : [t, ''];
}
