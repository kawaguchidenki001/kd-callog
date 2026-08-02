/*******************************************************
 * 通話記録システム v1  (Google Apps Script)
 * Galaxy自動録音 → Autosync → Drive受信フォルダ →
 * Gemini文字起こし → スプレッドシートDB → 検索API(doGet)
 *
 * 使い方: README_設定手順.md を参照
 *  1) CONFIG の2箇所を書き換える
 *  2) 関数「setup」を1回実行（フォルダ・シート・トリガー自動作成）
 *******************************************************/

/* APIキーと合言葉は「スクリプト プロパティ」に保存する（コードには書かない）。
 * GASエディタ左下の ⚙プロジェクトの設定 → スクリプト プロパティ で
 *   GEMINI_API_KEY … AI StudioのAPIキー
 *   API_TOKEN      … 検索アプリの合言葉（好きな英数字）
 * を登録する。こうしておけばコードを貼り直しても設定は消えない。
 * 登録できているかは checkSetup() を実行して確認する。 */
const CONFIG = {
  MODEL: 'gemini-3.5-flash',       // 無料枠で使えるモデル。エラーが出たらAI Studioの一覧の名前に変更
  MODEL_FALLBACK: 'gemini-2.5-flash', // 上が使えない場合に自動で試すモデル
  MAX_PER_RUN: 5,                  // 1回のトリガーで処理する最大件数
  MAX_INLINE_MB: 18,               // これ以下はリクエストに直接埋め込む。超えたらFiles API経由
  MAX_FILE_MB: 45,                 // Files API経由の上限(約45分の通話)。UrlFetchApp制限のためこれ超はスキップ
  TRIGGER_MINUTES: 5,              // 何分おきに新着チェックするか
  AUDIO_KEEP_DAYS: 90              // 処理済み音声を何日後に自動削除するか（0で削除しない）
};

const PROP = PropertiesService.getScriptProperties();
const HEADERS = ['日時', '名前', '電話番号', '通話時間(目安)', '要約', '全文', '音声リンク', 'ファイル名', '状態'];

function apiKey() { return String(PROP.getProperty('GEMINI_API_KEY') || '').trim(); }
function apiToken() { return String(PROP.getProperty('API_TOKEN') || '').trim(); }

/* ============ 設定の確認（困ったときに最初に実行する） ============
 * スクリプトプロパティ・フォルダ・シート・トリガーが揃っているかを点検する */
function checkSetup() {
  const mask = function (s) { return s ? s.substring(0, 4) + '…(' + s.length + '文字)' : '未設定！'; };
  Logger.log('GEMINI_API_KEY: ' + mask(apiKey()));
  Logger.log('API_TOKEN: ' + mask(apiToken()));
  ['INBOX_ID', 'DONE_ID', 'ERR_ID', 'SS_ID'].forEach(function (k) {
    Logger.log(k + ': ' + (PROP.getProperty(k) ? 'OK' : '未設定！ setup()を実行してください'));
  });
  const t = ScriptApp.getProjectTriggers().map(function (x) { return x.getHandlerFunction(); });
  Logger.log('トリガー: ' + (t.length ? t.join(', ') : 'なし！ setup()を実行してください'));
  try {
    const sh = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('通話記録');
    Logger.log('通話記録シート: ' + (sh ? (sh.getLastRow() - 1) + '件' : '見つかりません！'));
  } catch (e) {
    Logger.log('スプレッドシートを開けません: ' + e);
  }
}

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
  // キー未設定なら何もしない（空のまま動かすと全件エラー行になってしまうため）
  if (!apiKey()) {
    Logger.log('GEMINI_API_KEY が未設定です。プロジェクトの設定→スクリプト プロパティに登録してください');
    return;
  }

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
    items.sort(function (a, b) { return b.p.date - a.p.date; }); // 新しい順（直近の通話を優先。過去分は空き枠で消化）

    let count = 0;
    for (let i = 0; i < items.length; i++) {
      if (count >= CONFIG.MAX_PER_RUN) break;
      if (Date.now() - startMs > 270 * 1000) break; // 実行時間ガード(4.5分)

      const f = items[i].file;
      const p = items[i].p;
      const sizeMB = f.getSize() / (1024 * 1024);

      if (sizeMB > CONFIG.MAX_FILE_MB) {
        f.moveTo(errF);
        appendRecord(sheet, [p.date, p.name, p.tel, '', '', '', '', f.getName(),
          '大容量スキップ(' + sizeMB.toFixed(1) + 'MB)']);
        continue;
      }

      // 18MB超はアップロード＋処理待ちで数分かかるため、残り時間に余裕がある時だけ着手
      const isLarge = sizeMB > CONFIG.MAX_INLINE_MB;
      if (isLarge && Date.now() - startMs > 150 * 1000) continue;

      try {
        const g = isLarge ? transcribeLarge(f) : transcribe(f.getBlob());
        // 顧客マスタで名前⇔番号を補完
        let name = p.name, tel = p.tel;
        if (tel && !name) name = master.nameByTel[normTel(tel)] || '';
        if (name && !tel) tel = master.telByName[name] || '';
        // 未登録番号の名寄せ：通話中に相手が名乗っていたら名前を採用し、マスタへ候補行を追加
        if (tel && !name && g.name) {
          name = g.name;
          addMasterCandidate(g.name, g.company, tel, master);
        }

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
function inlinePart(blob, mime) {
  return { inline_data: { mime_type: mime, data: Utilities.base64Encode(blob.getBytes()) } };
}

function transcribe(blob) {
  const model = PROP.getProperty('ACTIVE_MODEL') || CONFIG.MODEL;
  const hints = buildHints();
  const dict = loadDict();
  let r;
  try {
    r = callGemini(inlinePart(blob, 'audio/mp4'), model, hints);
  } catch (e) {
    const msg = String(e);
    if (msg.indexOf('HTTP 400') >= 0) r = callGemini(inlinePart(blob, 'audio/aac'), model, hints); // MIME違いの保険
    else if (msg.indexOf('HTTP 404') >= 0 && model !== CONFIG.MODEL_FALLBACK) {
      r = callGemini(inlinePart(blob, 'audio/mp4'), CONFIG.MODEL_FALLBACK, hints);
      PROP.setProperty('ACTIVE_MODEL', CONFIG.MODEL_FALLBACK);
    } else throw e;
  }
  return withDict(r, dict);
}

/* 文字起こし結果の全フィールドに用語辞書を適用 */
function withDict(r, dict) {
  return {
    transcript: applyDict(r.transcript, dict),
    summary: applyDict(r.summary, dict),
    name: applyDict(r.name || '', dict),
    company: applyDict(r.company || '', dict)
  };
}

/* ============ 長時間通話（18MB超）: Files API経由 ============
 * インライン埋め込みの上限を超える音声は、いったんGeminiのFiles APIへ
 * アップロードし、処理完了(ACTIVE)を待ってから文字起こしする。
 * アップロードした一時ファイルは使用後すぐ削除（放置しても48時間で自動消滅） */
function transcribeLarge(file) {
  const mime = mimeFromName(file.getName());
  const model = PROP.getProperty('ACTIVE_MODEL') || CONFIG.MODEL;
  const hints = buildHints();
  const dict = loadDict();
  const up = filesApiUpload(file.getBlob(), mime);
  try {
    filesApiWaitActive(up.name);
    const part = { file_data: { mime_type: mime, file_uri: up.uri } };
    let r;
    try {
      r = callGemini(part, model, hints);
    } catch (e) {
      if (String(e).indexOf('HTTP 404') >= 0 && model !== CONFIG.MODEL_FALLBACK) {
        r = callGemini(part, CONFIG.MODEL_FALLBACK, hints);
        PROP.setProperty('ACTIVE_MODEL', CONFIG.MODEL_FALLBACK);
      } else throw e;
    }
    return withDict(r, dict);
  } finally {
    filesApiDelete(up.name);
  }
}

function mimeFromName(fname) {
  if (/\.(3ga|3gp)$/i.test(fname)) return 'audio/3gpp';
  if (/\.amr$/i.test(fname)) return 'audio/amr';
  return 'audio/mp4'; // m4a / mp4
}

/* Files APIへ再開可能アップロード（開始→本体送信の2段階） */
function filesApiUpload(blob, mime) {
  const bytes = blob.getBytes();
  const start = fetchWithAuth('https://generativelanguage.googleapis.com/upload/v1beta/files', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'X-Goog-Upload-Protocol': 'resumable',
      'X-Goog-Upload-Command': 'start',
      'X-Goog-Upload-Header-Content-Length': String(bytes.length),
      'X-Goog-Upload-Header-Content-Type': mime
    },
    payload: JSON.stringify({ file: { display_name: blob.getName() } }),
    muteHttpExceptions: true
  });
  const sc = start.getResponseCode();
  if (sc === 429 || sc === 503) throw new Error('RATE_LIMIT');
  if (sc !== 200) throw new Error('Files API開始失敗 HTTP ' + sc + ': ' + start.getContentText().substring(0, 200));
  const h = start.getHeaders();
  const uploadUrl = h['x-goog-upload-url'] || h['X-Goog-Upload-URL'] || h['X-Goog-Upload-Url'];
  if (!uploadUrl) throw new Error('Files APIのアップロードURLが取得できませんでした');

  const fin = UrlFetchApp.fetch(uploadUrl, {
    method: 'post',
    headers: { 'X-Goog-Upload-Command': 'upload, finalize', 'X-Goog-Upload-Offset': '0' },
    payload: bytes,
    muteHttpExceptions: true
  });
  const fc = fin.getResponseCode();
  if (fc !== 200) throw new Error('Files API送信失敗 HTTP ' + fc + ': ' + fin.getContentText().substring(0, 200));
  const j = JSON.parse(fin.getContentText());
  if (!j.file || !j.file.uri) throw new Error('Files API応答にfile.uriがありません');
  return { uri: j.file.uri, name: j.file.name };
}

/* アップロード後のサーバー側処理完了(ACTIVE)を待つ。最大約2分 */
function filesApiWaitActive(name) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/' + name;
  for (let i = 0; i < 24; i++) {
    const res = fetchWithAuth(url, { method: 'get', muteHttpExceptions: true });
    if (res.getResponseCode() === 200) {
      const st = JSON.parse(res.getContentText()).state;
      if (st === 'ACTIVE') return;
      if (st === 'FAILED') throw new Error('Files APIの音声処理に失敗しました');
    }
    Utilities.sleep(5000);
  }
  throw new Error('Files API処理待ちタイムアウト');
}

function filesApiDelete(name) {
  try {
    fetchWithAuth('https://generativelanguage.googleapis.com/v1beta/' + name,
      { method: 'delete', muteHttpExceptions: true });
  } catch (e) { /* 削除失敗は無視（48時間で自動消滅する） */ }
}

function callGemini(dataPart, model, hints) {
  const prompt =
    'これは業務電話の録音です。日本語で正確に文字起こしし、話者を「A:」「B:」で区別して1発言ごとに改行してください。' +
    '聞き取れない箇所は（不明瞭）としてください。加えて内容の要約を2文以内で作成してください。' +
    'また、電話の相手が会話中に名乗っていた場合は、その名前を"name"、会社名・屋号を"company"に入れてください（不明なら空文字）。' +
    (hints ? '登場する可能性のある固有名詞（この表記を優先して使うこと）: ' + hints + '。' : '') +
    '次の形式のJSONのみを返してください: {"transcript":"...","summary":"...","name":"...","company":"..."}';

  const base = 'https://generativelanguage.googleapis.com/v1beta/models/' + model + ':generateContent';

  const payload = {
    contents: [{
      parts: [dataPart, { text: prompt }]
    }],
    generationConfig: { responseMimeType: 'application/json', temperature: 0.2 }
  };
  const opt = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const res = fetchWithAuth(base, opt);
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
  var mn = t.match(/"name"\s*:\s*"((?:[^"\\]|\\.)*)"/);
  var mc = t.match(/"company"\s*:\s*"((?:[^"\\]|\\.)*)"/);
  if (mt) return { transcript: unesc(mt[1]), summary: ms ? unesc(ms[1]) : '',
    name: mn ? unesc(mn[1]) : '', company: mc ? unesc(mc[1]) : '' };

  // ④ 最後の手段：応答テキストをそのまま全文として保存（捨てない）
  if (t.length > 20) return { transcript: t, summary: '', name: '', company: '' };
  throw new Error('解析不能な応答');
}

function pick(j) {
  return { transcript: String(j.transcript || ''), summary: String(j.summary || ''),
    name: String(j.name || ''), company: String(j.company || '') };
}
function unesc(s) {
  return String(s).replace(/\\n/g, '\n').replace(/\\t/g, '\t')
    .replace(/\\"/g, '"').replace(/\\\\/g, '\\');
}

/* 認証方式：①x-goog-api-keyヘッダー ②?key= クエリ の順に試す
 * （新形式AQ.キー／旧形式AIzaキーのどちらでも通るようにするため）
 * 成功した方式を ScriptProperties に記憶して次回から使う */
function fetchWithAuth(url, opt) {
  const mode = PROP.getProperty('AUTH_MODE') || 'header';
  let res = fetchGemini(url, opt, mode);
  if (res.getResponseCode() === 401 || res.getResponseCode() === 403) {
    const other = mode === 'header' ? 'query' : 'header';
    const res2 = fetchGemini(url, opt, other);
    if (res2.getResponseCode() === 200) { PROP.setProperty('AUTH_MODE', other); res = res2; }
    else res = res2.getResponseCode() === 401 ? res : res2;
  }
  return res;
}

function fetchGemini(base, opt, mode) {
  const o = Object.assign({}, opt);
  if (mode === 'header') {
    o.headers = Object.assign({}, opt.headers, { 'x-goog-api-key': apiKey() });
    return UrlFetchApp.fetch(base, o);
  }
  return UrlFetchApp.fetch(base + (base.indexOf('?') >= 0 ? '&' : '?') +
    'key=' + encodeURIComponent(apiKey()), o);
}

/* ============ 接続テスト（キーが有効か確認する用） ============
 * 関数「testApiKey」を実行 → ログに「OK」が出れば準備完了 */
function testApiKey() {
  if (!apiKey()) {
    Logger.log('NG：GEMINI_API_KEY が未設定です。' +
      'プロジェクトの設定→スクリプト プロパティ に GEMINI_API_KEY を登録してください');
    return;
  }
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

/* ============ 顧客マスタの自動拡充（未登録番号の名寄せ） ============
 * 未登録番号の通話で相手が名乗っていた場合、マスタへ「要確認」付きの
 * 候補行を自動追加する。間違いがあればシート上で直せば以後はその表記が使われる */
function addMasterCandidate(name, company, tel, master) {
  name = String(name || '').trim();
  const key = normTel(tel);
  if (!name || name.length > 30 || !key || master.nameByTel[key]) return;
  const sh = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('顧客マスタ');
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  appendRecord(sh, [name, String(company || '').trim(), tel, '自動追加 ' + today + ' 要確認']);
  // 同一実行内の後続ファイルにも効かせる
  master.nameByTel[key] = name;
  if (!master.telByName[name]) master.telByName[name] = tel;
}

/* ============ 過去の記録の名前をマスタから一括補完 ============
 * 関数「fillNamesFromMaster」を実行すると、名前が空で電話番号がマスタに
 * ある行の名前を埋めます（自動追加分を確認・修正した後に実行すると便利） */
function fillNamesFromMaster() {
  const master = loadMaster();
  const sheet = SpreadsheetApp.openById(PROP.getProperty('SS_ID')).getSheetByName('通話記録');
  if (sheet.getLastRow() < 2) { Logger.log('記録がありません'); return; }
  const n = sheet.getLastRow() - 1;
  const rg = sheet.getRange(2, 2, n, 2); // 名前・電話番号
  const v = rg.getValues();
  let hit = 0;
  for (let i = 0; i < n; i++) {
    const name = String(v[i][0]).trim(), tel = normTel(String(v[i][1]));
    if (!name && tel && master.nameByTel[tel]) { v[i][0] = master.nameByTel[tel]; hit++; }
  }
  if (hit) rg.setValues(v);
  Logger.log(hit + '件の名前を補完しました');
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
  // 例外をそのまま投げるとGoogleのHTMLエラーページが返り、PWA側は「応答がありません」
  // としか分からなくなる。必ずJSON(P)で理由を返す
  try {
    return handleGet(p, cb);
  } catch (err) {
    return jsonOut({ error: 'server', detail: String(err && err.message || err).substring(0, 300) }, cb);
  }
}

function handleGet(p, cb) {
  const token = apiToken();
  if (!token) return jsonOut({ error: 'no_token', detail: 'スクリプトプロパティ API_TOKEN が未設定です' }, cb);
  if (p.token !== token) return jsonOut({ error: 'unauthorized' }, cb);

  if (p.ping === '1') return jsonOut({ ok: true }, cb); // 接続テスト用

  const ss = PROP.getProperty('SS_ID');
  if (!ss) return jsonOut({ error: 'server', detail: 'SS_IDが未設定です。setup()を実行してください' }, cb);
  const sheet = SpreadsheetApp.openById(ss).getSheetByName('通話記録');
  if (!sheet) return jsonOut({ error: 'server', detail: '「通話記録」シートが見つかりません' }, cb);
  const last = sheet.getLastRow();
  if (last < 2) return jsonOut({ count: 0, rows: [], last: last }, cb);

  // ---- 1件詳細（全文つき）。該当行だけを読むので速い ----
  if (p.id) {
    const id = Number(p.id);
    if (!(id >= 2 && id <= last)) return jsonOut({ count: 0, rows: [] }, cb);
    const r = sheet.getRange(id, 1, 1, HEADERS.length).getValues()[0];
    const one = rowObj(r, id);
    one.transcript = String(r[5]);
    return jsonOut({ count: 1, rows: [one] }, cb);
  }

  // ---- 差分取得：前回の続き（since行より後）だけを返す ----
  // PWAは取得済みをローカルに保存しているので、通常はここを通り一瞬で返る
  if (p.since) {
    const since = Number(p.since);
    if (!(since >= 2 && since <= last)) return jsonOut({ reload: true, last: last }, cb);
    // 行が削除されて番号がずれていないか、ファイル名で照合する
    if (p.sincekey && String(sheet.getRange(since, 8).getValue()) !== p.sincekey) {
      return jsonOut({ reload: true, last: last }, cb);
    }
    const add = readRows(sheet, since + 1, last - since, false);
    add.sort(byNewest);
    return jsonOut({ count: add.length, rows: add, last: last, added: true }, cb);
  }

  // ---- 一覧・検索 ----
  const needText = !!p.q; // 全文検索のときだけ全文を読む（重いので普段は読まない）
  let rows = readRows(sheet, 2, last - 1, needText);

  if (p.name) rows = rows.filter(function (r) { return r.name.indexOf(p.name) >= 0; });
  if (p.tel) rows = rows.filter(function (r) { return normTel(r.tel).indexOf(normTel(p.tel)) >= 0; });
  if (p.from) rows = rows.filter(function (r) { return r.datetime >= p.from; });
  if (p.to) rows = rows.filter(function (r) { return r.datetime <= p.to + ' 99'; });
  if (p.q) {
    rows = rows.filter(function (r) {
      return (r.name + r.tel + r.summary + r.transcript).indexOf(p.q) >= 0;
    });
    rows.forEach(function (r) { r.transcript = ''; }); // 一覧では使わないので返さない
  }

  rows.sort(byNewest);
  rows = rows.slice(0, Math.min(Number(p.limit) || 200, 1000));
  return jsonOut({ count: rows.length, rows: rows, last: last }, cb);
}

function byNewest(a, b) { return a.datetime < b.datetime ? 1 : -1; }

/* 必要な列だけ読む。全文(F列)は重いので withText のときだけ読む */
function readRows(sheet, start, n, withText) {
  if (n <= 0) return [];
  const a = sheet.getRange(start, 1, n, 5).getValues();  // 日時〜要約
  const b = sheet.getRange(start, 7, n, 3).getValues();  // 音声リンク・ファイル名・状態
  const t = withText ? sheet.getRange(start, 6, n, 1).getValues() : null;
  const out = [];
  for (let i = 0; i < n; i++) {
    const o = rowObj([a[i][0], a[i][1], a[i][2], a[i][3], a[i][4], '', b[i][0], b[i][1], b[i][2]], start + i);
    if (t) o.transcript = String(t[i][0]);
    out.push(o);
  }
  return out;
}

function rowObj(r, id) {
  return {
    id: id,
    datetime: r[0] instanceof Date ? Utilities.formatDate(r[0], 'Asia/Tokyo', 'yyyy-MM-dd HH:mm') : String(r[0]),
    name: String(r[1]), tel: String(r[2]), length: String(r[3]),
    summary: String(r[4]), transcript: '',
    audio: String(r[6]), file: String(r[7]), status: String(r[8])
  };
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
