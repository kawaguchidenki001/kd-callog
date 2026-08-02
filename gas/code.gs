/*******************************************************
 * 通話記録システム v1  (Google Apps Script)
 * Galaxy自動録音 → Autosync → Drive受信フォルダ →
 * Gemini文字起こし → スプレッドシートDB → 検索API(doGet)
 *
 * 使い方: README_設定手順.md を参照
 *  1) CONFIG の2箇所を書き換える
 *  2) 関数「setup」を1回実行（フォルダ・シート・トリガー自動作成）
 *******************************************************/

const CONFIG = {
  GEMINI_API_KEY: '',        // ← GASエディタ側で入力する。ここには絶対に書かない
  API_TOKEN: '',             // ← GASエディタ側で入力する。ここには絶対に書かない        // 検索アプリ用の合言葉（好きな英数字に変更）
  MODEL: 'gemini-3.5-flash',       // 無料枠で使えるモデル。エラーが出たらAI Studioの一覧の名前に変更
  MODEL_FALLBACK: 'gemini-2.5-flash', // 上が使えない場合に自動で試すモデル
  MAX_PER_RUN: 5,                  // 1回のトリガーで処理する最大件数
  MAX_INLINE_MB: 18,               // これを超える音声はDrive経由(Files API不要の分割)へ回さずスキップ
  TRIGGER_MINUTES: 5,              // 何分おきに新着チェックするか
  AUDIO_KEEP_DAYS: 90              // 処理済み音声を何日後に自動削除するか（0で削除しない）
};

const PROP = PropertiesService.getScriptProperties();
const HEADERS = ['日時', '名前', '電話番号', '通話時間(目安)', '要約', '全文', '音声リンク', 'ファイル名', '状態'];

/* ============ 初期セットアップ（1回だけ実行） ============ */
function setup() {
  // フォルダ作成
  const root = DriveApp.createFolder('通話記録システム');
  const inbox = root.createFolder('受信');
  const done = root.createFolder('処理済み');
  const err = root.createFolder('エラー');

  // スプレッドシート作成
  const ss = SpreadsheetApp.create('通話記録DB');
  ss.setSpreadsheetTimeZone('Asia/Tokyo');
  DriveApp.getFileById(ss.getId()).moveTo(root);

  const rec = ss.getSheets()[0];
  rec.setName('通話記録');
  rec.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
  rec.setFrozenRows(1);
  rec.setColumnWidth(5, 300); // 要約
  rec.setColumnWidth(6, 500); // 全文

  const master = ss.insertSheet('顧客マスタ');
  master.getRange(1, 1, 1, 4).setValues([['名前', '会社名', '電話番号', 'メモ']]).setFontWeight('bold');
  master.setFrozenRows(1);

  // ID保存
  PROP.setProperties({
    INBOX_ID: inbox.getId(),
    DONE_ID: done.getId(),
    ERR_ID: err.getId(),
    SS_ID: ss.getId()
  });

  // トリガー作成（既存があれば削除して作り直し）
  ScriptApp.getProjectTriggers().forEach(function (t) {
    const fn = t.getHandlerFunction();
    if (fn === 'processNewRecordings' || fn === 'cleanupOldAudio') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('processNewRecordings').timeBased().everyMinutes(CONFIG.TRIGGER_MINUTES).create();
  if (CONFIG.AUDIO_KEEP_DAYS) {
    ScriptApp.newTrigger('cleanupOldAudio').timeBased().everyDays(1).atHour(3).create();
  }

  Logger.log('セットアップ完了！');
  Logger.log('受信フォルダ: ' + inbox.getUrl());
  Logger.log('スプレッドシート: ' + ss.getUrl());
  Logger.log('Autosyncの同期先には「通話記録システム/受信」を指定してください。');
}

/* ============ メイン処理（5分毎に自動実行） ============ */
function processNewRecordings() {
  if (CONFIG.GEMINI_API_KEY.indexOf('ここに') === 0) return; // キー未設定なら何もしない

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(0)) return; // 前回実行中なら重複起動しない

  try {
    const startMs = Date.now();
    const inbox = DriveApp.getFolderById(PROP.getProperty('INBOX_ID'));
    const done = DriveApp.getFolderById(PROP.getProperty('DONE_ID'));
    const errF = DriveApp.getFolderById(PROP.getProperty('ERR_ID'));
    const sheet = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('通話記録');
    const master = loadMaster();

    // 受信フォルダのファイルを収集して古い順に並べる
    const items = [];
    const it = inbox.getFiles();
    while (it.hasNext()) {
      const f = it.next();
      const p = parseFileName(f.getName());
      if (!p) {
        f.moveTo(errF);
        appendRecord(sheet, ['', '', '', '', '', '', '', f.getName(), 'ファイル名形式外']);
        continue;
      }
      items.push({ file: f, p: p });
    }
    items.sort(function (a, b) { return a.p.date - b.p.date; });

    let count = 0;
    for (let i = 0; i < items.length; i++) {
      if (count >= CONFIG.MAX_PER_RUN) break;
      if (Date.now() - startMs > 270 * 1000) break; // 実行時間ガード(4.5分)

      const f = items[i].file;
      const p = items[i].p;
      const sizeMB = f.getSize() / (1024 * 1024);

      if (sizeMB > CONFIG.MAX_INLINE_MB) {
        f.moveTo(errF);
        appendRecord(sheet, [p.date, p.name, p.tel, '', '', '', '', f.getName(),
          '大容量スキップ(' + sizeMB.toFixed(1) + 'MB)']);
        continue;
      }

      try {
        const g = transcribe(f.getBlob());
        // 顧客マスタで名前⇔番号を補完
        let name = p.name, tel = p.tel;
        if (tel && !name) name = master.nameByTel[normTel(tel)] || '';
        if (name && !tel) tel = master.telByName[name] || '';

        f.moveTo(done);
        appendRecord(sheet, [
          p.date, name, tel,
          Math.max(1, Math.round(sizeMB)) + '分程度',
          g.summary, g.transcript,
          f.getUrl(), f.getName(), '完了'
        ]);
        count++;
      } catch (e) {
        const msg = String(e && e.message || e);
        if (msg.indexOf('RATE_LIMIT') >= 0) break; // 無料枠上限 → 次回トリガーで続きから
        f.moveTo(errF);
        appendRecord(sheet, [p.date, p.name, p.tel, '', '', '', '', f.getName(),
          'エラー: ' + msg.substring(0, 180)]);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

/* ============ エラー分をやり直す ============
 * 関数「retryErrors」を実行すると、エラーフォルダの音声を受信フォルダへ戻し、
 * シートのエラー行を削除します。次のトリガーで自動的に再処理されます。 */
function retryErrors() {
  const inbox = DriveApp.getFolderById(PROP.getProperty('INBOX_ID'));
  const errF = DriveApp.getFolderById(PROP.getProperty('ERR_ID'));
  const sheet = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('通話記録');

  let moved = 0;
  const it = errF.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!parseFileName(f.getName())) continue; // 名前が読めないものは戻さない
    f.moveTo(inbox);
    moved++;
  }

  // シートのエラー行を下から削除
  let removed = 0;
  if (sheet.getLastRow() >= 2) {
    const n = sheet.getLastRow() - 1;
    const st = sheet.getRange(2, 9, n, 1).getValues();
    for (let i = n - 1; i >= 0; i--) {
      if (String(st[i][0]).indexOf('エラー') === 0) { sheet.deleteRow(i + 2); removed++; }
    }
  }
  Logger.log('やり直し対象 ' + moved + '件を受信へ戻し、エラー行 ' + removed + '件を削除しました');
}

/* ============ 古い音声の自動削除（Drive容量対策・毎日1回） ============
 * テキスト（要約・全文）はシートに残るので検索閲覧に支障はありません */
function cleanupOldAudio() {
  if (!CONFIG.AUDIO_KEEP_DAYS) return;
  const limit = new Date(Date.now() - CONFIG.AUDIO_KEEP_DAYS * 86400 * 1000);
  const done = DriveApp.getFolderById(PROP.getProperty('DONE_ID'));
  const sheet = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('通話記録');

  // 削除対象ファイル名を集める
  const removed = {};
  const it = done.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (f.getDateCreated() < limit) { removed[f.getName()] = true; f.setTrashed(true); }
  }
  if (!Object.keys(removed).length || sheet.getLastRow() < 2) return;

  // シート側のリンクを消して状態を更新
  const n = sheet.getLastRow() - 1;
  const names = sheet.getRange(2, 8, n, 1).getValues();
  const links = sheet.getRange(2, 7, n, 1).getValues();
  const states = sheet.getRange(2, 9, n, 1).getValues();
  let changed = false;
  for (let i = 0; i < n; i++) {
    if (removed[String(names[i][0])] && links[i][0]) {
      links[i][0] = ''; states[i][0] = '完了(音声削除済)'; changed = true;
    }
  }
  if (changed) {
    sheet.getRange(2, 7, n, 1).setValues(links);
    sheet.getRange(2, 9, n, 1).setValues(states);
  }
}

/* ============ 用語辞書（聞き間違いの自動修正） ============
 * スプレッドシートの「用語辞書」シートに 誤→正 を書くと、
 * ①AIへのヒントとして渡し ②文字起こし後に自動置換します */
function getDictSheet() {
  const ss = SpreadsheetApp.openById(PROP.getProperty('SS_ID'));
  let sh = ss.getSheetByName('用語辞書');
  if (!sh) {
    sh = ss.insertSheet('用語辞書');
    sh.getRange(1, 1, 1, 2).setValues([['誤（聞き間違い）', '正（正しい表記）']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, 4, 2).setValues([
      ['川口電気', '河口電機'],
      ['川口電機', '河口電機'],
      ['川口', '河口'],
      ['皿谷', '更谷']
    ]);
    sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 200);
  }
  return sh;
}

function loadDict() {
  const sh = getDictSheet();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .map(function (r) { return { from: String(r[0]).trim(), to: String(r[1]).trim() }; })
    .filter(function (d) { return d.from && d.to && d.from !== d.to; })
    .sort(function (a, b) { return b.from.length - a.from.length; }); // 長い語から置換
}

function applyDict(text, dict) {
  let t = String(text || '');
  dict.forEach(function (d) { t = t.split(d.from).join(d.to); });
  return t;
}

/* AIに渡す固有名詞ヒント（辞書の正しい表記＋顧客マスタの名前） */
function buildHints() {
  const words = {};
  loadDict().forEach(function (d) { words[d.to] = true; });
  const sh = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('顧客マスタ');
  if (sh && sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      [r[0], r[1]].forEach(function (v) { v = String(v).trim(); if (v) words[v] = true; });
    });
  }
  const list = Object.keys(words).slice(0, 80);
  return list.length ? list.join('、') : '';
}

/* ============ 既存の記録に辞書を適用（過去分の一括修正） ============
 * 関数「fixExisting」を実行すると、要約と全文の表記を辞書どおりに直します */
function fixExisting() {
  const dict = loadDict();
  if (!dict.length) { Logger.log('用語辞書が空です'); return; }
  const sheet = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('通話記録');
  if (sheet.getLastRow() < 2) { Logger.log('記録がありません'); return; }

  const n = sheet.getLastRow() - 1;
  const rg = sheet.getRange(2, 2, n, 5); // 名前・電話番号・通話時間・要約・全文
  const v = rg.getValues();
  let hit = 0;
  for (let i = 0; i < n; i++) {
    const name = applyDict(v[i][0], dict);
    const sum = applyDict(v[i][3], dict);
    const full = applyDict(v[i][4], dict);
    if (name !== v[i][0] || sum !== v[i][3] || full !== v[i][4]) {
      v[i][0] = name; v[i][3] = sum; v[i][4] = full; hit++;
    }
  }
  if (hit) rg.setValues(v);
  Logger.log(hit + '件を修正しました');
}

/* ============ ファイル名の解析 ============
 * 例: 「通話記録 恵里_260103_063841.m4a」
 *     「通話記録 09012345678_260105_090307.m4a」 */
function parseFileName(fname) {
  const m = fname.match(/^通話記録[ 　]+(.+)_(\d{6})_(\d{6})\.(m4a|3ga|amr|mp4)$/i);
  if (!m) return null;
  const d = m[2], t = m[3];
  const date = new Date(2000 + Number(d.slice(0, 2)), Number(d.slice(2, 4)) - 1, Number(d.slice(4, 6)),
    Number(t.slice(0, 2)), Number(t.slice(2, 4)), Number(t.slice(4, 6)));
  const label = m[1].trim();
  const isTel = /^[0-9+\-() ]+$/.test(label);
  return { date: date, name: isTel ? '' : label, tel: isTel ? label : '' };
}

/* ============ Gemini文字起こし ============ */
function transcribe(blob) {
  const model = PROP.getProperty('ACTIVE_MODEL') || CONFIG.MODEL;
  const hints = buildHints();
  const dict = loadDict();
  let r;
  try {
    r = callGemini(blob, 'audio/mp4', model, hints);
  } catch (e) {
    const msg = String(e);
    if (msg.indexOf('HTTP 400') >= 0) r = callGemini(blob, 'audio/aac', model, hints); // MIME違いの保険
    else if (msg.indexOf('HTTP 404') >= 0 && model !== CONFIG.MODEL_FALLBACK) {
      r = callGemini(blob, 'audio/mp4', CONFIG.MODEL_FALLBACK, hints);
      PROP.setProperty('ACTIVE_MODEL', CONFIG.MODEL_FALLBACK);
    } else throw e;
  }
  return { transcript: applyDict(r.transcript, dict), summary: applyDict(r.summary, dict) };
}

function callGemini(blob, mime, model, hints) {
  const prompt =
    'これは業務電話の録音です。日本語で正確に文字起こしし、話者を「A:」「B:」で区別して1発言ごとに改行してください。' +
    '聞き取れない箇所は（不明瞭）としてください。加えて内容の要約を2文以内で作成してください。' +
    (hints ? '登場する可能性のある固有名詞（この表記を優先して使うこと）: ' + hints + '。' : '') +
    '次の形式のJSONのみを返してください: {"transcript":"...","summary":"..."}';

  const base = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';

  const payload = {
    contents: [{
      parts: [
        { inline_data: { mime_type: mime, data: Utilities.base64Encode(blob.getBytes()) } },
        { text: prompt }
      ]
    }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.2 }
  };
  const opt = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  // 認証方式：①x-goog-api-keyヘッダー ②?key= クエリ の順に試す
  // （新形式AQ.キー／旧形式AIzaキーのどちらでも通るようにするため）
  const mode = PROP.getProperty('AUTH_MODE') || 'header';
  let res = fetchGemini(base, opt, mode);
  if (res.getResponseCode() === 401 || res.getResponseCode() === 403) {
    const other = mode === 'header' ? 'query' : 'header';
    const res2 = fetchGemini(base, opt, other);
    if (res2.getResponseCode() === 200) { PROP.setProperty('AUTH_MODE', other); res = res2; }
    else res = res2.getResponseCode() === 401 ? res : res2;
  }

  const code = res.getResponseCode();
  if (code === 429 || code === 503) throw new Error('RATE_LIMIT');
  if (code !== 200) throw new Error('HTTP ' + code + ': ' + res.getContentText().substring(0, 200));

  const body = JSON.parse(res.getContentText());
  const part = body.candidates && body.candidates[0] &&
    body.candidates[0].content && body.candidates[0].content.parts &&
    body.candidates[0].content.parts[0];
  if (!part || !part.text) throw new Error('応答が空(' +
    (body.candidates && body.candidates[0] && body.candidates[0].finishReason || '不明') + ')');

  const raw = part.text;
  return parseResult(raw);
}

/* Geminiの応答を解析。余計な文字が付いていても本文を取りこぼさない */
function parseResult(raw) {
  var t = String(raw).replace(/^\uFEFF/, '').trim();
  t = t.replace(/^```[a-z]*\s*/i, '').replace(/```\s*$/, '').trim();

  // ① そのまま解析
  try { return pick(JSON.parse(t)); } catch (e) {}

  // ② 最初の { から対応する } までを切り出して解析（後ろの余計な文字を無視）
  var s = t.indexOf('{');
  if (s >= 0) {
    var depth = 0, inStr = false, escaped = false;
    for (var i = s; i < t.length; i++) {
      var c = t.charAt(i);
      if (inStr) {
        if (escaped) escaped = false;
        else if (c === '\\') escaped = true;
        else if (c === '"') inStr = false;
      } else if (c === '"') inStr = true;
      else if (c === '{') depth++;
      else if (c === '}') {
        depth--;
        if (depth === 0) {
          try { return pick(JSON.parse(t.substring(s, i + 1))); } catch (e2) {}
          break;
        }
      }
    }
  }

  // ③ それでもダメなら中身を正規表現で拾う
  var mt = t.match(/"transcript"\s*:\s*"((?:[^"\\]|\\.)*)"/);
  var ms = t.match(/"summary"\s*:\s*"((?:[^"\\]|\\.)*)"/);
  if (mt) return { transcript: unesc(mt[1]), summary: ms ? unesc(ms[1]) : '' };

  // ④ 最後の手段：応答テキストをそのまま全文として保存（捨てない）
  if (t.length > 20) return { transcript: t, summary: '' };
  throw new Error('解析不能な応答');
}

function pick(j) {
  return { transcript: String(j.transcript || ''), summary: String(j.summary || '') };
}
function unesc(s) {
  return String(s).replace(/\\n/g, '\n').replace(/\\t/g, '\t')
    .replace(/\\"/g, '"').replace(/\\\\/g, '\\');
}

function fetchGemini(base, opt, mode) {
  const o = Object.assign({}, opt);
  if (mode === 'header') {
    o.headers = { 'x-goog-api-key': CONFIG.GEMINI_API_KEY };
    return UrlFetchApp.fetch(base, o);
  }
  return UrlFetchApp.fetch(base + '?key=' + encodeURIComponent(CONFIG.GEMINI_API_KEY), o);
}

/* ============ 接続テスト（キーが有効か確認する用） ============
 * 関数「testApiKey」を実行 → ログに「OK」が出れば準備完了 */
function testApiKey() {
  const model = CONFIG.MODEL;
  const payload = { contents: [{ parts: [{ text: 'こんにちは、と一言返してください。' }] }] };
  const opt = { method: 'post', contentType: 'application/json',
    payload: JSON.stringify(payload), muteHttpExceptions: true };
  const base = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';

  ['header', 'query'].forEach(function (mode) {
    const res = fetchGemini(base, opt, mode);
    if (res.getResponseCode() === 200) {
      PROP.setProperty('AUTH_MODE', mode);
      Logger.log('OK：認証方式 ' + mode + ' / モデル ' + model + ' で接続成功しました');
    } else {
      Logger.log('NG(' + mode + ') ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 300));
    }
  });
}

/* ============ 顧客マスタ読込 ============ */
function loadMaster() {
  const sh = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('顧客マスタ');
  const telByName = {}, nameByTel = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
      const name = String(r[0]).trim(), tel = normTel(String(r[2]));
      if (name && tel) { telByName[name] = String(r[2]); nameByTel[tel] = name; }
    });
  }
  return { telByName: telByName, nameByTel: nameByTel };
}

function normTel(s) { return String(s).replace(/[^0-9]/g, ''); }

function appendRecord(sheet, row) {
  sheet.appendRow(row.map(function (v) {
    return (typeof v === 'string' && /^[=+@]/.test(v)) ? "'" + v : v; // 数式化防止
  }));
}

/* ============ 検索API（将来の検索PWA用） ============
 * デプロイ > 新しいデプロイ > ウェブアプリ で公開して使う
 * 例: <URL>?token=kensaku2026&q=見積&from=2026-07-01 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cb = p.callback || '';
  if (p.token !== CONFIG.API_TOKEN) return jsonOut({ error: 'unauthorized' }, cb);

  const sheet = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('通話記録');
  if (sheet.getLastRow() < 2) return jsonOut({ count: 0, rows: [] }, cb);

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADERS.length).getValues();
  let rows = values.map(function (r, i) {
    return {
      id: i + 2,
      datetime: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : String(r[0]),
      name: String(r[1]), tel: String(r[2]), length: String(r[3]),
      summary: String(r[4]), transcript: String(r[5]),
      audio: String(r[6]), status: String(r[8])
    };
  });

  if (p.id) { // 1件詳細
    const one = rows.filter(function (r) { return String(r.id) === String(p.id); });
    return jsonOut({ count: one.length, rows: one }, cb);
  }

  if (p.name) rows = rows.filter(function (r) { return r.name.indexOf(p.name) >= 0; });
  if (p.tel) rows = rows.filter(function (r) { return normTel(r.tel).indexOf(normTel(p.tel)) >= 0; });
  if (p.from) rows = rows.filter(function (r) { return r.datetime >= p.from; });
  if (p.to) rows = rows.filter(function (r) { return r.datetime <= p.to + ' 99'; });
  if (p.q) rows = rows.filter(function (r) {
    return (r.name + r.tel + r.summary + r.transcript).indexOf(p.q) >= 0;
  });

  rows.sort(function (a, b) { return a.datetime < b.datetime ? 1 : -1; }); // 新しい順
  const limit = Math.min(Number(p.limit) || 100, 300);
  rows = rows.slice(0, limit);

  if (p.full !== '1') rows.forEach(function (r) { // 一覧は全文を先頭120字に
    if (r.transcript.length > 120) r.transcript = r.transcript.substring(0, 120) + '…';
  });

  return jsonOut({ count: rows.length, rows: rows }, cb);
}

/* JSONP対応：callback指定時はJavaScriptとして返す（CORS制約を受けない） */
function jsonOut(obj, cb) {
  const json = JSON.stringify(obj);
  if (cb && /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
