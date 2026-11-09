/**
 * AI 체험 뽑기판 — 구글 시트 저장소 (Apps Script 웹 앱)
 *
 * 설치: 구글 시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여 넣고 저장
 *       → 배포 → 새 배포 → 유형: 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자 → 배포
 *       → 나온 웹 앱 URL(…/exec)을 appscript-config.js에 붙여 넣기
 *
 * 'data' 시트에 기록이 한 줄에 하나씩 저장돼요. 직접 고치지 마세요.
 * 시트 메뉴 [뽑기판 → 결과표 만들기]로 방문자·부스별 표를 볼 수 있어요.
 */

var SHEET = 'data';
var PATH_RE = /^(config|state|visitors|draws)\/[A-Za-z0-9_-]{1,40}$/;

function doGet(e) {
  var p = (e && e.parameter) || {};
  var ver = getVer_();
  if (p.action === 'get') {
    if (!PATH_RE.test(p.path || '')) return out_({ ok: false, code: 'bad-path' });
    var one = findRow_(p.path);
    return out_({ ok: true, ver: ver, data: one ? one.data : null });
  }
  if (p.since && String(ver) === String(p.since)) return out_({ ok: true, unchanged: true, ver: ver });
  return out_({ ok: true, ver: ver, docs: loadAll_() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return out_({ ok: false, code: 'busy' }); }
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var sh = sheet_();

    if (b.op === 'clear') {
      var cols = (b.collections || []).filter(function (c) { return /^(visitors|draws)$/.test(c); });
      var vals = sh.getDataRange().getValues();
      for (var i = vals.length - 1; i >= 1; i--) {
        var col = String(vals[i][0]).split('/')[0];
        if (cols.indexOf(col) >= 0) sh.deleteRow(i + 1);
      }
      return out_({ ok: true, ver: bump_() });
    }

    if (!PATH_RE.test(b.path || '')) return out_({ ok: false, code: 'bad-path' });
    var row = findRow_(b.path);
    var now = new Date();

    if (b.op === 'create' && row) return out_({ ok: false, code: 'already-exists' });
    if (b.op === 'set' || b.op === 'create') {
      var json = JSON.stringify(b.data || {});
      if (json.length > 45000) return out_({ ok: false, code: 'too-large' });
      if (row) sh.getRange(row.row, 2, 1, 2).setValues([[json, now]]);
      else sh.appendRow([b.path, json, now]);
    } else if (b.op === 'update') {
      if (!row) return out_({ ok: false, code: 'not-found' });
      var merged = row.data || {};
      var patch = b.data || {};
      for (var k in patch) merged[k] = patch[k];
      sh.getRange(row.row, 2, 1, 2).setValues([[JSON.stringify(merged), now]]);
    } else if (b.op === 'delete') {
      if (row) sh.deleteRow(row.row);
    } else {
      return out_({ ok: false, code: 'bad-op' });
    }
    return out_({ ok: true, ver: bump_() });
  } catch (err) {
    return out_({ ok: false, code: 'error', error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 내부 함수 ---------- */
function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(['path', 'json', 'updatedAt']); sh.setFrozenRows(1); }
  return sh;
}
function loadAll_() {
  var vals = sheet_().getDataRange().getValues(), docs = {};
  for (var i = 1; i < vals.length; i++) {
    if (!vals[i][0]) continue;
    try { docs[vals[i][0]] = JSON.parse(vals[i][1]); } catch (err) {}
  }
  return docs;
}
function findRow_(path) {
  var vals = sheet_().getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] === path) {
      var d = null; try { d = JSON.parse(vals[i][1]); } catch (err) {}
      return { row: i + 1, data: d };
    }
  }
  return null;
}
function getVer_() { return Number(PropertiesService.getScriptProperties().getProperty('ver') || 0); }
function bump_() {
  var v = Math.max(Date.now(), getVer_() + 1);
  PropertiesService.getScriptProperties().setProperty('ver', String(v));
  return v;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 시트 메뉴: 결과표 ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('뽑기판').addItem('결과표 만들기', 'makeReport').addToUi();
}
function makeReport() {
  var docs = loadAll_(), visitors = [], draws = [];
  for (var p in docs) {
    if (p.indexOf('visitors/') === 0) visitors.push(docs[p]);
    if (p.indexOf('draws/') === 0) draws.push(docs[p]);
  }
  visitors.sort(function (a, b) { return (a.no || 0) - (b.no || 0) || (a.createdAt || 0) - (b.createdAt || 0); });
  var LV = { m: '중', h: '고', k: '유', e: '초' };
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var vs = ss.getSheetByName('방문자') || ss.insertSheet('방문자');
  vs.clear();
  var rows = [['참가번호', '학교급', '학교', '학년', '이름', '등록 시각', '뽑은 체험', '완료', '체험 목록']];
  visitors.forEach(function (v) {
    var mine = draws.filter(function (d) { return d.code === v.code; });
    rows.push([v.no || '', LV[v.level] || '', v.school || '', v.grade || '', v.name || '',
      v.createdAt ? new Date(v.createdAt) : '', mine.length,
      mine.filter(function (d) { return d.done; }).length,
      mine.map(function (d) { return d.exp + (d.done ? '(완료)' : ''); }).join(', ')]);
  });
  vs.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  vs.setFrozenRows(1);

  var bs = ss.getSheetByName('부스별') || ss.insertSheet('부스별');
  bs.clear();
  var count = {};
  draws.forEach(function (d) {
    var c = count[d.exp] = count[d.exp] || { n: 0, done: 0, m: 0, h: 0 };
    c.n++; if (d.done) c.done++;
    var v = visitors.filter(function (x) { return x.code === d.code; })[0];
    if (v && (v.level === 'm' || v.level === 'h')) c[v.level]++;
  });
  var brows = [['체험(부스)', '뽑힘', '체험 완료', '중', '고']];
  Object.keys(count).sort(function (a, b) { return count[b].n - count[a].n; }).forEach(function (k) {
    brows.push([k, count[k].n, count[k].done, count[k].m, count[k].h]);
  });
  bs.getRange(1, 1, brows.length, 5).setValues(brows);
  bs.setFrozenRows(1);
  SpreadsheetApp.getActive().toast('방문자 ' + visitors.length + '명, 뽑기 ' + draws.length + '건으로 결과표를 만들었어요.');
}
