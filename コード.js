/***** ── 設定 ─────────────────────────────────*****/
const SPREADSHEET_ID = '1wdHF0txuZtrkMrC128fwUSImyt320JhBVqXloS7FgpU'; // ←ご指定
const SHEET_NAME      = 'Monitoring'; // ケアマネ用モニタリング
const OPENAI_MODEL    = 'gpt-4o-mini';
const SHARE_SHEET_NAME = 'ExternalShares';
const SHARE_LOG_SHEET_NAME = 'ExternalShareAccessLog';
const MONITORING_REPORTS_SHEET_NAME = 'MonitoringReports';
const SHARE_QR_SIZE = '220x220';
const SHARE_QR_FOLDER_ID = '1QZTrJ0_c07ILdqLg1jelYhOQ7hSGAdmt';
const QR_FOLDER_ID = '1QZTrJ0_c07ILdqLg1jelYhOQ7hSGAdmt';
const HONOBONO_SHEET_NAME = 'ほのぼのID';
const HONOBONO_QR_URL_COL = 6; // F列
const HONOBONO_MEMBER_ID_HEADER   = 'ほのぼのID';
const HONOBONO_NAME_HEADER        = '氏名';
const HONOBONO_KANA_HEADER        = 'フリガナ';
const HONOBONO_CENTER_HEADER      = '地域包括支援センター';
const HONOBONO_STAFF_HEADER       = '担当者名';
const HONOBONO_QR_HEADER          = '共有QRコードURL'; // 参照は見出し優先。定数 HONOBONO_QR_URL_COL があればフォールバック

// 画像/動画/PDF の既定保存先（利用者IDごとにサブフォルダを自動作成）
const DEFAULT_FOLDER_ID         = '1glDniVONBBD8hIvRGMPPT1iLXdtHJpEC';
const MEDIA_ROOT_FOLDER_ID      = DEFAULT_FOLDER_ID;
const REPORT_FOLDER_ID_PROP     = DEFAULT_FOLDER_ID;
const ATTACHMENTS_FOLDER_ID_PROP= DEFAULT_FOLDER_ID;

// Docsテンプレ（任意）：プロパティで上書き可（なければ自動レイアウト）
const DOC_TEMPLATE_ID_PROP        = PropertiesService.getScriptProperties().getProperty('DOC_TEMPLATE_ID') || '';
const DOC_TEMPLATE_ID_FAMILY_PROP = PropertiesService.getScriptProperties().getProperty('DOC_TEMPLATE_ID_FAMILY') || '';

/** パラメータの余計な " を除去するユーティリティ */
function cleanParam_(value) {
  return String(value || "")
    .trim()
    .replace(/^"+|"+$/g, "");   // 先頭・末尾の " を削除
}

function doGet(e) {
  try {
    Logger.log("🟢 doGet called at " + new Date());
    Logger.log("raw event = " + JSON.stringify(e));

    const params = (e && e.parameter) ? e.parameter : {};
    Logger.log("params = " + JSON.stringify(params));

    // ============ JSON API（/exec?api=shareMeta ...） ============
    if (params.api === 'share') {
      const token = cleanParam_(params.shareId || params.share || params.token || '');
      const recordId = cleanParam_(params.recordId || params.record || '');
      const password = cleanParam_(params.password || params.pass || '');
      if (!token) {
        return ContentService.createTextOutput(
          JSON.stringify({ status: 'error', message: 'shareId is missing' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
      const result = enterExternalShare(token, password, recordId);
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (params.api === 'shareMeta') {
      Logger.log("🌐 API mode detected");
      // raw を優先して拾う（空白や改行だけは後段で除去）
      const rawToken  = params.shareId || params.share || params.token || "";
      const token     = cleanParam_(rawToken); // 先頭末尾の余計な " だけ除去
      const recordId  = cleanParam_(params.recordId || params.record || "");
      Logger.log(`API token(raw)="${rawToken}" token(clean)="${token}" recordId="${recordId}"`);

      if (!token) {
        return ContentService.createTextOutput(
          JSON.stringify({ status:'error', message:'shareId is missing' })
        ).setMimeType(ContentService.MimeType.JSON);
      }

      const result = getExternalShareMeta(token, recordId);
      Logger.log("API result = " + JSON.stringify(result));
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ============ HTML 表示モード ============
    // ★ ここ超重要：raw パラメータをまず拾う
    const rawToken  = params.shareId || params.share || params.token || "";
    const rawRecord = params.recordId || params.record || "";

    // 表側埋め込み値（JS でそのまま使わせる）
    const embedToken   = rawToken;    // 加工しない（フロントがそのまま使う）
    const embedRecordId= rawRecord;   // 同上
    const printParamRaw = params.print || params.mode || "";

    // 「見た目上の判定」は空白以外の文字があるかどうかで決定
    const hasToken   = String(embedToken).trim() !== "";
    const wantsPrint = hasToken && String(printParamRaw).trim() !== "" && String(printParamRaw).trim() !== "0";

    Logger.log(`HTML embedToken="${embedToken}" recordId="${embedRecordId}" wantsPrint=${wantsPrint}`);

    // ★ raw の token が1文字でもあれば share テンプレートを出す
    const templateName = wantsPrint ? "print" : (hasToken ? "share" : "member");
    Logger.log("templateName = " + templateName);

    const tmpl = HtmlService.createTemplateFromFile(templateName);

    // ★ 埋め込み（share.html 側の TEMPLATE_TOKEN / TEMPLATE_RECORD_ID に反映される）
    tmpl.shareToken   = embedToken || "";
    tmpl.shareRecordId= embedRecordId || "";

    let title = hasToken ? "モニタリング共有ビュー" : "ケアマネ・モニタリング";

    // ======= 印刷モード：share が確定している時だけサーバ側で事前解決 =======
    if (wantsPrint && hasToken) {
      Logger.log("👉 print mode detected (server-side prefetch)");
      // ここでは clean 済み token をサーバ関数に渡す（シート照合は厳密）
      const tokenClean = cleanParam_(embedToken);
      const recordIdClean = cleanParam_(embedRecordId);
      const meta = getExternalShareMeta(tokenClean, recordIdClean);
      tmpl.shareMeta = meta;

      let printMode = "record";
      let printRecords = [];
      let primaryRecord = null;
      let centerLabel = "";
      let staffLabel = "";
      let errorMessage = "";
      const requestedMode = String(params.mode || "").trim().toLowerCase();

      const context = shareFindByToken_(tokenClean);
      if (context && meta && meta.status === "success") {
        const shareState = context.share;
        const initialRecords = Array.isArray(meta.records) ? meta.records.slice() : [];
        primaryRecord = meta.primaryRecord || (initialRecords.length ? initialRecords[0] : null);
        printRecords = initialRecords;

        if (!printRecords.length) {
          const fallback = shareBuildResponse_(shareState, recordIdClean, true, true);
          printRecords = fallback.records.slice();
          primaryRecord = fallback.primaryRecord || primaryRecord;
        }

        if (requestedMode === "center" && primaryRecord && primaryRecord.center) {
          const centerRecords = getRecordsByCenter(primaryRecord.center);
          const payload = shareBuildCustomRecordSet_(shareState, centerRecords, recordIdClean);
          printRecords = payload.records;
          primaryRecord = payload.primaryRecord || primaryRecord;
          centerLabel = primaryRecord.center || (primaryRecord.fields && primaryRecord.fields.center) || "";
          printMode = "center";
        } else if (requestedMode === "staff" && primaryRecord && primaryRecord.staff) {
          const staffRecords = getRecordsByStaff(primaryRecord.staff);
          const payload = shareBuildCustomRecordSet_(shareState, staffRecords, recordIdClean);
          printRecords = payload.records;
          primaryRecord = payload.primaryRecord || primaryRecord;
          staffLabel = primaryRecord.staff || (primaryRecord.fields && primaryRecord.fields.staff) || "";
          printMode = "staff";
        }
      } else {
        errorMessage = meta && meta.message ? String(meta.message) : "共有情報を取得できませんでした。";
      }

      tmpl.printMode = printMode;
      tmpl.printRecords = printRecords;
      tmpl.printPrimaryRecord = primaryRecord;
      tmpl.printCenter = centerLabel;
      tmpl.printStaff = staffLabel;
      tmpl.printErrorMessage = errorMessage;
      tmpl.printRecordId = recordIdClean;

      const tz = Session.getScriptTimeZone ? (Session.getScriptTimeZone() || "Asia/Tokyo") : "Asia/Tokyo";
      tmpl.printedAtText = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd HH:mm");
      title = "モニタリング記録 印刷";
    }

    Logger.log("✅ doGet finished, returning template: " + templateName);

    return tmpl.evaluate()
      .setTitle(title)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");

  } catch (err) {
    Logger.log("❌ ERROR in doGet: " + (err && err.stack ? err.stack : err));
    return ContentService.createTextOutput(
      JSON.stringify({ status: 'error', message: String(err && err.message || err) })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}





/** クライアントから参照するための Web アプリURL（/exec） */
function getExecUrl(){ return ScriptApp.getService().getUrl(); }

/***** ── 保存（テキスト＋添付メタ） ─────────────────*****/
function saveRecordFromBrowser(memberId, content, isoTimestamp, attachmentsJson, kind) {
  if (!memberId) throw new Error('利用者IDが空です');
  if (!content && !attachmentsJson) throw new Error('保存する内容が空です');

  const sheet = ensureSheet_();
  const ts    = isoTimestamp ? new Date(isoTimestamp) : new Date();
  const kindSafe = String(kind || 'その他').trim();

  sheet.appendRow([
    ts,
    String(memberId).trim(),
    kindSafe,                       // 種別
    String(content || '').trim(),   // 記録内容
    String(attachmentsJson || '[]') // 添付（JSON）
  ]);
  return { status: 'success' };
}

/***** ── 汎用 doPost（外部共有 API 用分岐を追加） ──*****/
function doPost(e) {
  try {
    // 1) 生パラメータ／JSON ボディの解釈
    var params = e && e.parameter ? e.parameter : {};
    var action = params.action || params.shareApi || '';
    var jsonPayload = null;

    // try parse JSON body if present
    if (!action && e && e.postData && e.postData.contents) {
      var postType = (e.postData && e.postData.type) || '';
      if (postType.indexOf('application/json') === 0) {
        try {
          jsonPayload = JSON.parse(e.postData.contents);
          action = action || jsonPayload.action || jsonPayload.shareApi || '';
        } catch (_err) {
          jsonPayload = null;
        }
      }
    }

    action = String(action || '').trim();

    // normalize a few common aliases
    if (action === 'enter' ) action = 'shareEnter';
    if (action === 'meta') action = 'shareMeta';

    // 2) 外部共有：閲覧（enter）
    if (action === 'shareEnter') {
      var tokenParam = (params && (params.shareId || params.share || params.token)) || '';
      if (!tokenParam && jsonPayload) {
        tokenParam = jsonPayload.shareId || jsonPayload.share || jsonPayload.token || '';
      }
      var passwordParam = (params && params.password) || '';
      if (!passwordParam && jsonPayload) {
        passwordParam = jsonPayload.password || '';
      }
      var recordIdParam = (params && (params.recordId || params.record)) || '';
      if (!recordIdParam && jsonPayload) {
        recordIdParam = jsonPayload.recordId || jsonPayload.record || '';
      }
      var shareResult = enterExternalShare(tokenParam, passwordParam, recordIdParam);
      return ContentService.createTextOutput(JSON.stringify(shareResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3) 外部共有：メタ取得（POST 経由で来た場合のサポート）
    if (action === 'shareMeta') {
      var tokenParam2 = (params && (params.shareId || params.share || params.token)) || '';
      if (!tokenParam2 && jsonPayload) {
        tokenParam2 = jsonPayload.shareId || jsonPayload.share || jsonPayload.token || '';
      }
      var recordIdParam2 = (params && (params.recordId || params.record)) || '';
      if (!recordIdParam2 && jsonPayload) {
        recordIdParam2 = jsonPayload.recordId || jsonPayload.record || '';
      }
      var metaResult = getExternalShareMeta(tokenParam2, recordIdParam2);
      return ContentService.createTextOutput(JSON.stringify(metaResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 4) 既存のバイナリアップロード処理（action === 'upload' を期待）
    if (action === 'upload') {
      var memberId = (params && params.memberId) || '';
      var name = (params && params.name) || 'upload';
      if (!memberId) throw new Error('memberId is required');

      var up = e && e.files && (e.files.file || e.files['file']);
      if (!up) throw new Error('no file found (e.files.file is empty)');
      if (Array.isArray(up)) up = up[0];

      var blob = up;
      if (name) blob.setName(name);

      var root = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID_PROP || DEFAULT_FOLDER_ID);
      var folder = getOrCreateChildFolder_(root, String(memberId).trim());
      var file = folder.createFile(blob);
      if (name) file.setName(name);

      var fileId = file.getId();
      var url = 'https://drive.google.com/file/d/' + fileId + '/view';

      try { ensureSharingForMember_(file, memberId); } catch (_e) {}

      var out = { status: 'success', fileId: fileId, url: url, name: file.getName(), mimeType: file.getMimeType(), uploadedAt: new Date().toISOString() };
      return ContentService.createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 5) unknown action
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    var outErr = { status: 'error', message: String(err && err.message || err) };
    return ContentService.createTextOutput(JSON.stringify(outErr))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
// 地域包括支援センター・担当者名を保存
function saveCenterInfo(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("ほのぼのID"); // 保存先タブ
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues(); 
  // A=ほのぼのID, B=氏名, C=フリガナ, D=地域包括支援センター/担当者

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(data.memberId)) {
      // D列に地域包括支援センター + 担当者名を保存
      sheet.getRange(i + 2, 4).setValue(data.center + "／" + data.staff);
      return { ok: true };
    }
  }
  return { ok: false, message: "ID not found" };
}
/***** ── Base64アップロード（フロントから呼ばれる） ─────────────────*****/
function uploadAttachment_(memberId, fileName, mimeType, base64) {
  const where = [];
  try {
    where.push('start');
    if (!memberId) throw new Error('memberIdが未指定です');
    if (!fileName) throw new Error('fileNameが未指定です');
    if (!base64) throw new Error('base64が空です');

    where.push('folder');
    const rootId = ATTACHMENTS_FOLDER_ID_PROP || MEDIA_ROOT_FOLDER_ID;
    const root = DriveApp.getFolderById(rootId);
    if (!root) throw new Error('保存先フォルダIDが不正: ' + rootId);
    const folder = getOrCreateChildFolder_(root, String(memberId).trim());

    where.push('decode');
    let bytes;
    try { bytes = Utilities.base64Decode(base64); }
    catch (e) { throw new Error('base64デコードに失敗: ' + e); }

    const blob = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', fileName);

    where.push('createFile');
    const file = folder.createFile(blob);
    file.setName(fileName);

    where.push('share');
    try { ensureSharingForMember_(file, memberId); } catch (e) { Logger.log('share error: ' + e); }

    const fileId = file.getId();
    const url = 'https://drive.google.com/file/d/' + fileId + '/view';
    const uploadedAt = new Date().toISOString();

    where.push('done');
    return { status:'success', fileId, url, name:file.getName(), mimeType:file.getMimeType(), uploadedAt };

  } catch (err) {
    const msg = 'uploadAttachment_ 失敗 at [' + where.join(' > ') + ']: ' + (err && err.message || err);
    Logger.log(msg);
    return { status:'error', message: msg };
  }
}

/***** ── 取得（期間対応・行番号・添付付き） ─────────────────*****/
function getRecordsByMemberId_v3(memberId, days) {
  const dbg = { spreadsheetId: SPREADSHEET_ID, sheetName: SHEET_NAME, memberId: String(memberId), days };
  try {
    const data = fetchRecordsWithIndex_(memberId, days);
    return { status:'success', records: data, data, debug: { ...dbg, matched: data.length } };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e), debug:dbg };
  }
}

function fetchRecordsWithIndex_(memberId, days) { 
  if (!memberId) throw new Error('memberIdが未指定です');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);

  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length <= 1) return [];

  const header = vals[0].map(v => String(v || '').trim());
  const indexes = resolveRecordColumnIndexes_(header);
  if (indexes.date < 0 || indexes.memberId < 0 || indexes.kind < 0 || indexes.record < 0 || indexes.attachments < 0) {
    throw new Error(`ヘッダー不一致（必要: 日付/利用者ID/種別/記録内容/添付, 実際: ${JSON.stringify(header)}）`);
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  let limitDate = null;
  if (days && String(days) !== 'all') {
    const n = Number(days);
    if (!isNaN(n) && n > 0) limitDate = new Date(Date.now() - n * 24 * 3600 * 1000);
  }

  const out = [];
  const targetId = String(memberId).trim();
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const id  = String(row[indexes.memberId] || '').trim();
    const dateVal = row[indexes.date];

    // 🔎 デバッグ: 各行を確認
    Logger.log("行%d: ID=%s, targetId=%s, date=%s", i+1, id, targetId, dateVal);

    if (id !== targetId) {
      Logger.log("  → memberId不一致でスキップ");
      continue;
    }

    const record = buildRecordFromRow_(row, header, indexes, tz, i);

    if (limitDate && record.timestamp !== null && record.timestamp < limitDate.getTime()) {
      Logger.log("  → 日付制限によりスキップ: ts=%s, limit=%s", record.timestamp, limitDate);
      continue;
    }

    Logger.log("  → 採用: %s", JSON.stringify(record));
    out.push(record);
  }

  out.sort((a,b) => {
    const ta = (typeof a.timestamp === 'number') ? a.timestamp : 0;
    const tb = (typeof b.timestamp === 'number') ? b.timestamp : 0;
    if (tb !== ta) return tb - ta;
    return (b.rowIndex || 0) - (a.rowIndex || 0);
  });

  Logger.log("✅ fetchRecordsWithIndex_: memberId=%s, days=%s, found=%s", targetId, days, out.length);

  return out;
}





function resolveRecordColumnIndexes_(header){
  const trimmed = header.map(v => String(v || '').trim());
  const lower = trimmed.map(v => v.toLowerCase());
  const find = (...names) => {
    for (let i = 0; i < names.length; i++) {
      const candidate = String(names[i] || '').trim();
      if (!candidate) continue;
      const idxExact = trimmed.indexOf(candidate);
      if (idxExact >= 0) return idxExact;
      const idxLower = lower.indexOf(candidate.toLowerCase());
      if (idxLower >= 0) return idxLower;
    }
    return -1;
  };
  return {
    date: find('日付','date'),
    memberId: find('利用者ID','memberid','id'),
    kind: find('種別','区分','kind'),
    record: find('記録内容','本文','text','内容'),
    attachments: find('添付','attachments'),
    center: find('center','センター','地域包括支援センター'),
    staff: find('staff','担当者'),
    status: find('status','状態','状態・経過','経過'),
    special: find('special','特記事項','特記','特記事項・備考'),
    recordId: find('recordId','recordid','記録ID'),
    memberName: find('利用者名','氏名','名前','memberName')
  };
}

function formatRecordDate_(value, tz){
  const d = (value instanceof Date) ? value : new Date(value);
  if (d && d.getTime && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, tz, 'yyyy/MM/dd HH:mm');
  }
  return String(value ?? '');
}

function formatFieldValue_(value, tz){
  if (value == null) return '';
  if (value instanceof Date) {
    return formatRecordDate_(value, tz);
  }
  if (typeof value === 'number' && !isFinite(value)) return '';
  return String(value);
}

function buildRecordFromRow_(row, header, indexes, tz, rowIndex){
  let attachmentsRaw = [];
  if (indexes.attachments >= 0) {
    try {
      attachmentsRaw = JSON.parse(String(row[indexes.attachments] || '[]')) || [];
    } catch(_err) {
      attachmentsRaw = [];
    }
  }
  const normalizedAttachments = Array.isArray(attachmentsRaw)
    ? attachmentsRaw.map(att => {
        if (att && typeof att === 'object') {
          const fileId = String(att.fileId || att.id || '').trim();
          const url = String(att.url || (fileId ? `https://drive.google.com/file/d/${fileId}/view` : '') || '').trim();
          const name = String(att.name || att.fileName || '').trim();
          const mimeType = String(att.mimeType || att.type || '').trim();
          return { fileId, url, name, mimeType };
        }
        const label = String(att ?? '').trim();
        return { fileId: '', url: '', name: label, mimeType: '' };
      })
    : [];

  const attachmentsSummary = normalizedAttachments.map(att => att && att.name ? att.name : (att && att.url ? att.url : '')).filter(Boolean).join('\n');
  const rawDate = indexes.date >= 0 ? row[indexes.date] : '';
  const dateText = indexes.date >= 0 ? formatRecordDate_(rawDate, tz) : '';
  const timestamp = (() => {
    if (!(rawDate instanceof Date)) {
      const d = new Date(rawDate);
      if (d && !isNaN(d.getTime())) return d.getTime();
    }
    return (rawDate instanceof Date && !isNaN(rawDate.getTime())) ? rawDate.getTime() : null;
  })();

  const fields = {};
  for (let idx = 0; idx < header.length; idx++) {
    const key = String(header[idx] || '').trim();
    if (!key) continue;
    if (idx === indexes.attachments) {
      fields[key] = attachmentsSummary;
    } else if (idx === indexes.date) {
      fields[key] = dateText;
    } else {
      fields[key] = formatFieldValue_(row[idx], tz);
    }
  }

  const recordIdValue = indexes.recordId >= 0 ? String(row[indexes.recordId] || '').trim() : '';
  const centerValue = indexes.center >= 0 ? String(row[indexes.center] || '').trim() : '';
  const staffValue = indexes.staff >= 0 ? String(row[indexes.staff] || '').trim() : '';
  const statusValue = indexes.status >= 0 ? String(row[indexes.status] || '').trim() : '';
  const specialValue = indexes.special >= 0 ? String(row[indexes.special] || '').trim() : '';
  let memberNameValue = indexes.memberName >= 0 ? String(row[indexes.memberName] || '').trim() : '';
  if (recordIdValue && !('recordId' in fields)) {
    fields.recordId = recordIdValue;
  }
  if (centerValue && !('center' in fields)) {
    fields.center = centerValue;
  }
  if (staffValue && !('staff' in fields)) {
    fields.staff = staffValue;
  }
  if (statusValue && !('status' in fields)) {
    fields.status = statusValue;
  }
  if (specialValue && !('special' in fields)) {
    fields.special = specialValue;
  }
  if (!memberNameValue) {
    memberNameValue = fields['利用者名'] || fields['氏名'] || fields['名前'] || '';
  }

  const textValue = indexes.record >= 0 ? String(row[indexes.record] ?? '') : '';
  const kindValue = indexes.kind >= 0 ? String(row[indexes.kind] ?? '') : '';
  const memberIdValue = indexes.memberId >= 0 ? String(row[indexes.memberId] || '').trim() : '';
  const sheetRowIndex = rowIndex + 1;

  return {
    rowIndex: sheetRowIndex,
    recordId: recordIdValue || String(sheetRowIndex),
    memberId: memberIdValue,
    memberName: memberNameValue,
    dateText,
    kind: kindValue,
    text: textValue,
    attachments: normalizedAttachments,
    timestamp,
    center: centerValue,
    staff: staffValue,
    status: statusValue,
    special: specialValue,
    fields
  };
}

function fetchRecordsByColumn_(columnKey, value){
  const target = String(value || '').trim();
  if (!target) return [];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length <= 1) return [];
  const header = vals[0].map(v => String(v || '').trim());
  const indexes = resolveRecordColumnIndexes_(header);
  const targetIndex = columnKey === 'center' ? indexes.center : indexes.staff;
  if (targetIndex < 0) return [];
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const lowerTarget = target.toLowerCase();
  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const cell = String(row[targetIndex] || '').trim();
    if (cell.toLowerCase() !== lowerTarget) continue;
    out.push(buildRecordFromRow_(row, header, indexes, tz, i));
  }
  out.sort((a,b) => {
    const ta = (typeof a.timestamp === 'number') ? a.timestamp : 0;
    const tb = (typeof b.timestamp === 'number') ? b.timestamp : 0;
    if (tb !== ta) return tb - ta;
    return (b.rowIndex || 0) - (a.rowIndex || 0);
  });
  return out;
}

function getRecordsByCenter(centerName){
  return fetchRecordsByColumn_('center', centerName);
}

function getRecordsByStaff(staffName){
  return fetchRecordsByColumn_('staff', staffName);
}

/***** ── ダッシュボード要約 ─────────────────*****/
function normalizeMemberId_(value) {
  if (value == null) return '';
  const normalized = String(value).normalize('NFKC').replace(/[^0-9]/g, '');
  if (!normalized) return '';
  if (normalized.length >= 4) return normalized;
  return ('0000' + normalized).slice(-4);
}

function normalizeMemberHeaderLabel_(label) {
  if (label == null) return '';
  return toHiragana_(String(label))
    .replace(/[\s　]+/g, '')
    .replace(/[()（）]/g, '')
    .toLowerCase();
}

function findMemberSheetColumnIndex_(headerNormalized, candidates) {
  if (!Array.isArray(headerNormalized)) return -1;
  for (const candidate of candidates) {
    const normalizedCandidate = normalizeMemberHeaderLabel_(candidate);
    if (!normalizedCandidate) continue;
    const idx = headerNormalized.findIndex(label => label === normalizedCandidate || label.includes(normalizedCandidate));
    if (idx >= 0) return idx;
  }
  return -1;
}

function getMemberSheetColumnInfo_(values) {
  const info = {
    header: [],
    headerNormalized: [],
    width: 0,
    idCol: -1,
    nameCol: -1,
    yomiCol: -1,
    careCol: -1,
    centerCol: -1,
    qrCol: -1
  };
  if (!Array.isArray(values) || !values.length) return info;

  const header = Array.isArray(values[0]) ? values[0].map(v => String(v || '').trim()) : [];
  const headerNormalized = header.map(normalizeMemberHeaderLabel_);
  const width = header.length;

  info.header = header;
  info.headerNormalized = headerNormalized;
  info.width = width;

  const idCandidates = ['id', '利用者id', 'りようしゃid', 'ご利用者id', 'モニタリングid'];
  let idCol = findMemberSheetColumnIndex_(headerNormalized, idCandidates);
  if (idCol < 0) {
    idCol = width > 0 ? 0 : -1;
  }

  const nameCandidates = ['氏名', '利用者名', '名前', '氏名漢字', 'しめい', 'なまえ'];
  let nameCol = findMemberSheetColumnIndex_(headerNormalized, nameCandidates);
  if (nameCol < 0) {
    if (width > 1) {
      nameCol = 1;
      if (nameCol === idCol && width > nameCol + 1) {
        nameCol = nameCol + 1;
      }
    } else if (width > 0) {
      nameCol = 0;
    }
  }

  const yomiCandidates = [
    'ふりがな', 'よみ', 'よみがな', 'しめいふりがな', 'しめいよみ', 'しめいかな',
    'かな', 'かなめい', 'ふりかな', 'めいかな', '氏名かな', '氏名ｶﾅ', '氏名カナ', 'しめいかな'
  ];
  const careCandidates = ['担当ケアマネ', '担当けあまね', 'ケアマネ', 'けあまね', '担当者', 'たんとうしゃ', '担当', 'たんとう'];
  const centerCandidates = ['包括支援センター', '地域包括支援センター', '包括', '地域包括'];
  const qrCandidates = [
    '共有qrコードurl',
    '共有qrこーどurl',
    '共有きゅーあーるこーどurl',
    'qrコードurl',
    'qrこーどurl',
    'きゅーあーるこーどurl',
    'qrurl',
    'qrコードリンク',
    'qrコード',
    'qr'
  ];

  const yomiCol = findMemberSheetColumnIndex_(headerNormalized, yomiCandidates);
  const careCol = findMemberSheetColumnIndex_(headerNormalized, careCandidates);
  const centerCol = findMemberSheetColumnIndex_(headerNormalized, centerCandidates);
  const qrCol = findMemberSheetColumnIndex_(headerNormalized, qrCandidates);

  info.idCol = idCol;
  info.nameCol = nameCol;
  info.yomiCol = yomiCol;
  info.careCol = careCol;
  info.centerCol = centerCol;
  info.qrCol = qrCol;
  return info;
}

const SMALL_KANA_MAP_ = {
  'ぁ':'あ','ぃ':'い','ぅ':'う','ぇ':'え','ぉ':'お',
  'っ':'つ','ゃ':'や','ゅ':'ゆ','ょ':'よ','ゎ':'わ','ゕ':'か','ゖ':'け'
};

function toHiragana_(value) {
  return String(value || '')
    .normalize('NFKC')
    .replace(/[ァ-ン]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0x60));
}

function buildDashboardSortKey_(entry) {
  if (!entry || typeof entry !== 'object') return '';
  const primary = entry.yomi || entry.name || '';
  const fallback = primary || entry.id || '';
  const base = primary || fallback;
  if (!base) return '';
  return toHiragana_(base)
    .replace(/[ぁ-ん]/g, ch => SMALL_KANA_MAP_[ch] || ch)
    .replace(/[\s　]+/g, '');
}

function hasFurigana_(entry) {
  if (!entry || typeof entry !== 'object') return false;
  const yomi = entry.yomi == null ? '' : String(entry.yomi).trim();
  return yomi !== '';
}

function getDashboardSummary() {
  const dbg = { spreadsheetId: SPREADSHEET_ID, sheetName: SHEET_NAME };
  try {
    const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const now = new Date();
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
    monthStart.setHours(0, 0, 0, 0);
    const monthLabel = Utilities.formatDate(monthStart, tz, 'yyyy/MM');

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    const memberMap = {};
    const memberSheet = ss.getSheetByName('ほのぼのID');
    if (memberSheet) {
      const mVals = memberSheet.getDataRange().getValues();
      const layout = getMemberSheetColumnInfo_(mVals);
      for (let i = 1; i < mVals.length; i++) {
        const row = mVals[i];
        const idValue = (layout.idCol >= 0 && layout.idCol < row.length) ? row[layout.idCol] : '';
        const id = normalizeMemberId_(idValue);
        if (!id) continue;
        const name = (layout.nameCol >= 0 && layout.nameCol < row.length) ? String(row[layout.nameCol] || '').trim() : '';
        const rawYomi = (layout.yomiCol >= 0 && layout.yomiCol < row.length) ? row[layout.yomiCol] : '';
        const yomi = rawYomi == null ? '' : String(rawYomi).normalize('NFKC').trim();
        const careRaw = (layout.careCol >= 0 && layout.careCol < row.length) ? row[layout.careCol] : '';
        const careManager = careRaw == null ? '' : String(careRaw).trim();
        memberMap[id] = { name, yomi, careManager };
      }
    }

    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
    const vals = sh.getDataRange().getValues();
    if (!vals || vals.length === 0) {
      const emptyData = Object.keys(memberMap).map(id => {
        const info = memberMap[id] || {};
        return {
          id,
          name: info.name || '',
          careManager: info.careManager || '',
          countThisMonth: 0,
          latestTimestamp: null,
          latestDateText: '',
          monitoringStatus: 'pending'
        };
      });
      return { status: 'success', data: emptyData, monthLabel, debug: dbg };
    }

    const header = vals[0].map(v => String(v || '').trim());
    const colDate = header.indexOf('日付');
    const colId   = header.indexOf('利用者ID');
    if (colDate < 0 || colId < 0) {
      throw new Error(`ヘッダー不一致（必要: 日付/利用者ID, 実際: ${JSON.stringify(header)}）`);
    }

    const summaryMap = new Map();
    const ensureEntry = (id) => {
      const info = memberMap[id] || {};
      if (!summaryMap.has(id)) {
        summaryMap.set(id, {
          id,
          name: info.name || '',
          yomi: info.yomi || '',
          careManager: info.careManager || '',
          countThisMonth: 0,
          latestTimestamp: null,
        });
      } else {
        const entry = summaryMap.get(id);
        if (info.name && !entry.name) entry.name = info.name;
        if (info.yomi && !entry.yomi) entry.yomi = info.yomi;
        if (info.careManager) entry.careManager = info.careManager;
      }
      return summaryMap.get(id);
    };

    for (let i = 1; i < vals.length; i++) {
      const row = vals[i];
      const rawId = String(row[colId] || '').trim();
      if (!rawId) continue;
      const half = rawId.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/[^0-9]/g, '');
      const id = ('0000' + half).slice(-4);
      if (!id) continue;
      const entry = ensureEntry(id);

      const rawDate = row[colDate];
      const d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      if (!(d instanceof Date) || isNaN(d.getTime())) continue;
      const ts = d.getTime();
      if (!entry.latestTimestamp || ts > entry.latestTimestamp) {
        entry.latestTimestamp = ts;
      }
      if (ts >= monthStart.getTime()) {
        entry.countThisMonth += 1;
      }
    }

    Object.keys(memberMap).forEach(id => ensureEntry(id));

    const data = Array.from(summaryMap.values()).map(entry => {
      const info = memberMap[entry.id] || {};
      const name = entry.name || info.name || '';
      const yomi = entry.yomi || info.yomi || '';
      const careManager = entry.careManager || info.careManager || '';
      const latestTimestamp = entry.latestTimestamp || null;
      return {
        id: entry.id,
        name,
        yomi,
        careManager,
        countThisMonth: entry.countThisMonth,
        latestTimestamp,
        latestDateText: latestTimestamp
          ? Utilities.formatDate(new Date(latestTimestamp), tz, 'yyyy/MM/dd HH:mm')
          : '',
        monitoringStatus: entry.countThisMonth > 0 ? 'completed' : 'pending'
      };
    });

    data.sort((a, b) => {
      const aHas = hasFurigana_(a);
      const bHas = hasFurigana_(b);
      if (aHas !== bHas) return aHas ? -1 : 1;
      const keyA = buildDashboardSortKey_(a);
      const keyB = buildDashboardSortKey_(b);
      const cmpKey = keyA.localeCompare(keyB, 'ja');
      if (cmpKey !== 0) return cmpKey;
      const nameA = String(a.name || '');
      const nameB = String(b.name || '');
      const cmpName = nameA.localeCompare(nameB, 'ja');
      if (cmpName !== 0) return cmpName;
      return String(a.id || '').localeCompare(String(b.id || ''), 'ja');
    });

    return { status: 'success', data, monthLabel, debug: dbg };
  } catch (e) {
    return { status: 'error', message: String(e && e.message || e), debug: dbg };
  }
}

/***** ── AI要約／アドバイス（ケアマネ視点） ─────────────────*****/
function generateAISummaryForDays(memberId, format, days) {
  try {
    const records = fetchRecordsWithIndex_(memberId, days);
    if (records.length === 0) return { status:'success', summary:'記録がありません。' };

    const lines = records
      .map(r => `【${r.dateText}｜${r.kind}】${oneLine_(r.text, 140)}`)
      .join('\n');

    const system = `あなたは介護支援専門員（ケアマネジャー）のモニタリング記録要約アシスタントです。
- 介護保険法に沿ったモニタリング視点（アセスメント/生活状況/ADL/IADL/リスク/医療的配慮/家族支援/多職種連携/サービス実施状況/課題/支援方針/次回予定）で簡潔に。
- 個人情報はぼかし、断定的な医療判断は避け、観察事実と助言を分ける。`;

    let user;
    switch (format) {
      case 'icf':
        user = `以下をICF視点（心身機能/活動/参加/環境因子/個人因子）で要約し、最後に「総合評価/次回までの支援方針」を添えて200～250字で。\n\n${lines}`;
        break;
      case 'soap':
        user = `以下をSOAP（S/O/A/P）で要約。Pでは「支援方針・連携依頼・次回予定」を具体的に。200～250字。\n\n${lines}`;
        break;
      case 'doctor':
        user = `以下を医療連携向けに、事実（Vitals/服薬/症状変化/転倒等/通院・受診調整）を中心に200～250字で要約。受診判断材料を簡潔に。\n\n${lines}`;
        break;
      case 'family':
        user = `以下を家族向けにやさしい表現で、安心材料/見守りのコツ/受診目安/次回までのお願いを含め200～250字でまとめてください。\n\n${lines}`;
        break;
      default:
        user = `以下のモニタリングから、生活状況/課題/リスク/サービス実施状況/支援方針/次回予定の順で200～250字に要約。\n\n${lines}`;
    }

    const text = openaiChat_(OPENAI_MODEL, system, user, 500, 0.3);
    const periodLabel = (!days || days==='all') ? '全期間' : `直近${days}日`;
    saveSummaryLog_(memberId, 'summary', periodLabel, text);

    return { status:'success', summary:text };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}

// 置き換え：期間固定だった generateCareAdviceForDays を汎用化
function generateCareAdviceForDays(memberId, days) {
  // 下位互換：既存呼び出しは「3か月」をデフォルトに
  return generateCareAdviceWithHorizon(memberId, days, '3m');
}

/**
 * 追加：緊急度（horizon）を指定して提案を生成
 * horizon: 'now' | '2w' | '1m' | '3m'
 */
function generateCareAdviceWithHorizon(memberId, days, horizon) {
  try {
    const records = fetchRecordsWithIndex_(memberId, days);
    if (records.length === 0) return { status:'success', advice:'記録がありません。' };

    const lines = records
      .map(r => `【${r.dateText}｜${r.kind}】${oneLine_(r.text, 140)}`)
      .join('\n');

    const horizonMap = {
      'now': { label:'すぐに対応', word:'直ちに着手する', limit:'200～250字', extras:'優先順位・責任者・期限を必ず明記。' },
      '2w' : { label:'2週間',     word:'今後2週間で',        limit:'250～300字', extras:'短期で達成可能なマイルストーンを設定。' },
      '1m' : { label:'1か月',     word:'今後1か月間で',      limit:'300～350字', extras:'週次の確認ポイントを含める。' },
      '3m' : { label:'3か月',     word:'今後3か月間で',      limit:'350～400字', extras:'月次ゴールと見直し時期を示す。' }
    };
    const hv = horizonMap[horizon] || horizonMap['3m'];

    const system = `あなたはケアマネ視点の多職種連携コーディネーターです。
- 安全第一、在宅生活の継続を支える具体策を短文で。
- 「サービス」「家族」「環境調整」「リスク対応」「医療連携」「次回アクション」に分けて出力。
- 数値・担当・期限をできる範囲で明記し、曖昧さを避ける。`;

    const user = `以下のモニタリングを踏まえ、${hv.word}実行する具体策を、各見出しごとに1～3行で提案してください。
制約：合計${hv.limit}、専門用語は避け、家庭や事業所でも実行しやすい内容。${hv.extras}
見出しは「サービス／家族／環境調整／リスク対応／医療連携／次回アクション」。
${lines}`;

    const text = openaiChat_(OPENAI_MODEL, system, user, 700, 0.4);
    const periodLabel = (!days || days==='all') ? '全期間' : `直近${days}日`;
    const label = hv.label;

    saveSummaryLog_(memberId, `advice-${horizon}`, `${periodLabel}｜${label}`, text);
    return { status:'success', advice:text, horizon: label };

  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}


/***** ── PDF（Docs→PDF化） ─────────────────*****/
function generatePdfReport(memberId, days, audience) {
  const dbg = {
    memberId: String(memberId),
    days,
    audience: audience || 'doctor',
    templateId: DOC_TEMPLATE_ID_PROP,
    templateIdFamily: DOC_TEMPLATE_ID_FAMILY_PROP,
    folderId: REPORT_FOLDER_ID_PROP
  };
  try {
    if (!memberId) return { status:'error', message:'利用者IDが未指定です', debug:dbg };

    const periodLabel = (!days || days === 'all') ? '全期間' : `直近${days}日`;
    const records = fetchRecordsWithIndex_(memberId, days);
    const formatForAudience =
      audience === 'family' ? 'family' :
      audience === 'doctor' ? 'doctor' : 'normal';

    const summaryRes = generateAISummaryForDays(memberId, formatForAudience, days);
    const summaryText = (summaryRes && summaryRes.status === 'success') ? summaryRes.summary : '';

    const now = new Date();
    const tz  = Session.getScriptTimeZone() || 'Asia/Tokyo';
    const ymd = Utilities.formatDate(now, tz, 'yyyyMMdd_HHmm');

    const audMap = { family:'家族向け', doctor:'医療連携', normal:'事業者向け' };
    const audienceTag = audMap[audience] || '事業者向け';

    const docName = `モニタリング報告書_${memberId}_${audienceTag}_${ymd}`;

    let docId;
    if (DOC_TEMPLATE_ID_FAMILY_PROP && audience === 'family') {
      docId = fillTemplateDoc_(DOC_TEMPLATE_ID_FAMILY_PROP, docName, memberId, periodLabel, audienceTag, summaryText, records);
    } else if (DOC_TEMPLATE_ID_PROP) {
      docId = fillTemplateDoc_(DOC_TEMPLATE_ID_PROP, docName, memberId, periodLabel, audienceTag, summaryText, records);
    } else {
      docId = buildDocFallback_(docName, memberId, periodLabel, audienceTag, summaryText, records);
    }

    const docFile = DriveApp.getFileById(docId);
    const pdfBlob = docFile.getAs('application/pdf').setName(docName + '.pdf');
    const pdfFile = REPORT_FOLDER_ID_PROP ? DriveApp.getFolderById(REPORT_FOLDER_ID_PROP).createFile(pdfBlob)
                                          : DriveApp.createFile(pdfBlob);

    ensureSharingForMember_(pdfFile, memberId);

    return { status:'success', fileId: pdfFile.getId(), fileName: pdfFile.getName(), debug:dbg };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e), debug:dbg };
  }
}

function fillTemplateDoc_(templateId, docName, memberId, periodLabel, audienceTag, summaryText, records){
  const copy = DriveApp.getFileById(templateId).makeCopy(docName);
  const docId = copy.getId();
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now = new Date();
  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  body.replaceText('{{MEMBER_ID}}', String(memberId));
  body.replaceText('{{PERIOD}}', String(periodLabel));
  body.replaceText('{{GENERATED_AT}}', Utilities.formatDate(now, tz, 'yyyy/MM/dd HH:mm'));
  body.replaceText('{{AUDIENCE}}', String(audienceTag));
  body.replaceText('{{SUMMARY}}', summaryText || '（要約なし）');

  const recordsText = (records.length
    ? records.map(r => `・${r.dateText}【${r.kind}】 ${r.text}`).join('\n')
    : '（該当期間の記録なし）');
  body.replaceText('{{RECORDS}}', recordsText);

  doc.saveAndClose();
  return docId;
}

function buildDocFallback_(docName, memberId, periodLabel, audienceTag, summaryText, records){
  const doc = DocumentApp.create(docName);
  const tz  = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now = new Date();
  const body = doc.getBody();

  body.appendParagraph('モニタリング報告書').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendParagraph(`利用者ID：${memberId}　期間：${periodLabel}　宛先：${audienceTag}`);
  body.appendParagraph(Utilities.formatDate(now, tz, '作成日時：yyyy/MM/dd HH:mm')).setForegroundColor('#666666');

  body.appendParagraph('要約（ケアマネ視点）').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  body.appendParagraph(summaryText || '（要約なし）');

  body.appendParagraph('記録（時系列）').setHeading(DocumentApp.ParagraphHeading.HEADING2);
  if (records.length) {
    records.forEach(r => body.appendListItem(`${r.dateText}【${r.kind}】 ${r.text}`));
  } else {
    body.appendParagraph('（該当期間の記録なし）');
  }

  doc.saveAndClose();
  return doc.getId();
}

/***** ── 編集／削除 ─────────────────*****/
function findMonitoringRowIndex_(identifier, values, indexes){
  if (!identifier || typeof identifier !== 'object') return 0;
  if (!Array.isArray(values) || values.length <= 1) return 0;
  const safeString = (value) => String(value == null ? '' : value).trim();
  const maxRow = values.length;

  let candidate = Number(identifier.rowIndex || identifier.row || identifier.sheetRow);
  if (!candidate && identifier.recordId && /^\d+$/.test(String(identifier.recordId))) {
    candidate = Number(identifier.recordId);
  }
  if (candidate && candidate >= 2 && candidate <= maxRow) {
    return candidate;
  }

  const recordId = safeString(identifier.recordId || identifier.id || '');
  if (recordId && indexes.recordId >= 0) {
    for (let i = 1; i < values.length; i++) {
      const cell = safeString(values[i][indexes.recordId]);
      if (cell === recordId) {
        return i + 1;
      }
    }
  }

  const memberId = safeString(identifier.memberId || '');
  if (memberId && indexes.memberId >= 0) {
    let fallback = 0;
    const targetTs = Number(identifier.timestamp || 0);
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowMember = safeString(row[indexes.memberId]);
      if (rowMember !== memberId) continue;
      if (indexes.date >= 0 && targetTs) {
        const rawDate = row[indexes.date];
        let rowTs = NaN;
        if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
          rowTs = rawDate.getTime();
        } else {
          const parsed = new Date(rawDate);
          if (!isNaN(parsed.getTime())) rowTs = parsed.getTime();
        }
        if (!isNaN(rowTs) && Math.abs(rowTs - targetTs) <= 1000) {
          return i + 1;
        }
      }
      if (!fallback) fallback = i + 1;
    }
    if (fallback) return fallback;
  }
  return 0;
}

function updateMonitoringRecord(data){
  try {
    const payload = data && typeof data === 'object' ? data : {};
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
    const values = sheet.getDataRange().getValues();
    if (!values || values.length <= 1) throw new Error('記録が存在しません');
    const header = values[0].map(v => String(v || '').trim());
    const indexes = resolveRecordColumnIndexes_(header);
    const rowIndex = findMonitoringRowIndex_(payload, values, indexes);
    if (!rowIndex || rowIndex < 2) throw new Error('対象の記録が見つかりません');
    if (payload.memberId && indexes.memberId >= 0) {
      const currentMember = String(values[rowIndex - 1][indexes.memberId] || '').trim();
      if (currentMember && String(payload.memberId).trim() && currentMember !== String(payload.memberId).trim()) {
        throw new Error('対象の記録が見つかりません');
      }
    }
    const sanitize = (value) => String(value == null ? '' : value).trim();
    if (indexes.center >= 0) {
      sheet.getRange(rowIndex, indexes.center + 1).setValue(sanitize(payload.center));
    }
    if (indexes.staff >= 0) {
      sheet.getRange(rowIndex, indexes.staff + 1).setValue(sanitize(payload.staff));
    }
    if (indexes.status >= 0) {
      sheet.getRange(rowIndex, indexes.status + 1).setValue(sanitize(payload.status));
    }
    if (indexes.special >= 0) {
      sheet.getRange(rowIndex, indexes.special + 1).setValue(sanitize(payload.special));
    }
    return { status:'success', rowIndex };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}

function deleteMonitoringRecord(identifier){
  try {
    const payload = (identifier && typeof identifier === 'object') ? identifier : { recordId: identifier };
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
    const values = sheet.getDataRange().getValues();
    if (!values || values.length <= 1) throw new Error('記録が存在しません');
    const header = values[0].map(v => String(v || '').trim());
    const indexes = resolveRecordColumnIndexes_(header);
    const rowIndex = findMonitoringRowIndex_(payload, values, indexes);
    if (!rowIndex || rowIndex < 2) throw new Error('対象の記録が見つかりません');
    sheet.deleteRow(rowIndex);
    return { status:'success' };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}

/***** ── 権限管理（Accessシート：利用者ID/氏名/メール） ─────────────────*****/
function ensureAccessSheet_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Access');
  if (!sh){
    sh = ss.insertSheet('Access');
    sh.appendRow(['利用者ID','氏名','メール']);
  }
  return sh;
}
function readAccessEmails_(memberId){
  const sh = ensureAccessSheet_();
  const vals = sh.getDataRange().getValues();
  const out = [];
  for (let i=1; i<vals.length; i++){
    if (String(vals[i][0]).trim() !== String(memberId).trim()) continue;
    const email = String(vals[i][2] || '').trim();
    if (email) out.push(email);
  }
  return out;
}
function ensureSharingForMember_(file, memberId){
  const emails = readAccessEmails_(memberId);
  if (emails && emails.length) {
    try { file.addViewers(emails); }
    catch(e){ Logger.log('share error: '+e); }
  }
}

/***** ── ログ保存（要約/アドバイス） ─────────────────*****/
function ensureLogSheet_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName('Logs');
  if (!sh){
    sh = ss.insertSheet('Logs');
    sh.appendRow(['日時','利用者ID','種別','期間','内容']);
  }
  return sh;
}
function saveSummaryLog_(memberId, kind, periodLabel, text){
  const sh = ensureLogSheet_();
  sh.appendRow([new Date(), String(memberId), String(kind), String(periodLabel), String(text || '')]);
}

/***** ── ユーティリティ ─────────────────*****/
function ensureSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const header = ['日付','利用者ID','種別','記録内容','添付'];
  const lr = sheet.getLastRow();
  sheet.getRange(1,1,1,header.length).setValues([header]);
  // フィルタや保護はお好みで追加可能

  return sheet;
}
function getOrCreateChildFolder_(rootFolder, childName){
  var it = rootFolder.getFoldersByName(childName);
  if (it.hasNext()) return it.next();
  return rootFolder.createFolder(childName);
}
function oneLine_(s, maxLen) {
  const t = String(s || '').replace(/\s+/g,' ').trim();
  return (maxLen && t.length > maxLen) ? t.slice(0, maxLen) + '…' : t;
}
function openaiChat_(model, system, user, maxTokens, temperature) {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('OPENAI_API_KEY が未設定です（スクリプトプロパティに保存してください）');

  const payload = {
    model,
    messages: [
      ...(system ? [{ role: 'system', content: system }] : []),
      { role: 'user', content: user }
    ],
    temperature: temperature ?? 0.3,
    max_tokens: maxTokens ?? 400
  };

  const res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) throw new Error(`OpenAI API エラー (${code}): ${body}`);

  const json = JSON.parse(body);
  const text = (json && json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || '';
  return String(text).trim();
}

/***** ── テスト ─────────────────*****/
function quickTestGet_v3() {
  const res = getRecordsByMemberId_v3('1','30');
  Logger.log(JSON.stringify(res, null, 2));
  return res;
}
function diagnoseUploadPath(){
  var root = DriveApp.getFolderById(ATTACHMENTS_FOLDER_ID_PROP || DEFAULT_FOLDER_ID);
  var testFolder = getOrCreateChildFolder_(root, 'DIAGNOSE_TEST');
  var file = testFolder.createFile(Utilities.newBlob('ok','text/plain','diag.txt'));
  Logger.log('created: %s in %s', file.getId(), testFolder.getName());
}
/** 利用者一覧を取得（ほのぼのIDシートから） */
function getMemberList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('ほのぼのID');
  if (!sh) throw new Error('シート「ほのぼのID」が見つかりません');
  const vals = sh.getDataRange().getValues();
  const layout = getMemberSheetColumnInfo_(vals);
  const out = [];

  for (let i=1; i<vals.length; i++) {
    const row = vals[i];
    const idValue = (layout.idCol >= 0 && layout.idCol < row.length) ? row[layout.idCol] : '';
    const id = normalizeMemberId_(idValue);
    if (!id) continue;
    const name = (layout.nameCol >= 0 && layout.nameCol < row.length) ? String(row[layout.nameCol] || '').trim() : '';
    const rawYomi = (layout.yomiCol >= 0 && layout.yomiCol < row.length) ? row[layout.yomiCol] : '';
    const yomi = rawYomi == null ? '' : String(rawYomi).normalize('NFKC').trim();
    const kana = yomi;
    const careRaw = (layout.careCol >= 0 && layout.careCol < row.length) ? row[layout.careCol] : '';
    const careManager = careRaw == null ? '' : String(careRaw).trim();
    out.push({ id, name, yomi, kana, careManager });
  }

  out.sort((a, b) => {
    const aHas = hasFurigana_(a);
    const bHas = hasFurigana_(b);
    if (aHas !== bHas) return aHas ? -1 : 1;
    const keyA = buildDashboardSortKey_(a);
    const keyB = buildDashboardSortKey_(b);
    const cmpKey = keyA.localeCompare(keyB, 'ja');
    if (cmpKey !== 0) return cmpKey;
    const nameA = String(a.name || '');
    const nameB = String(b.name || '');
    const cmpName = nameA.localeCompare(nameB, 'ja');
    if (cmpName !== 0) return cmpName;
    return String(a.id || '').localeCompare(String(b.id || ''), 'ja');
  });

  return out;
}

  /** 新規利用者を登録 */
function addMember(id, name) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('ほのぼのID');
  if (!sh) throw new Error('シート「ほのぼのID」が見つかりません');

  // IDフォーマット修正
  id = String(id || '').replace(/[^0-9]/g,'');
  id = ('0000' + id).slice(-4);

  // 氏名フォーマット修正
  name = String(name || '').trim().replace(/\s+/g,' ');
  
  // 重複チェック
  const vals = sh.getDataRange().getValues();
  for (let i=1; i<vals.length; i++){
    if (String(vals[i][0]) === id){
      throw new Error('同じIDがすでに存在します: ' + id);
    }
  }

  sh.appendRow([id, name]);
  return { status:'success', id, name };
}

/** 既存利用者の氏名を更新 */
function updateMemberName(id, newName){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('ほのぼのID');
  const vals = sh.getDataRange().getValues();

  id = String(id).replace(/[^0-9]/g,'');
  id = ('0000' + id).slice(-4);
  newName = String(newName).trim().replace(/\s+/g,' ');

  for (let i=1; i<vals.length; i++){
    if (String(vals[i][0]) === id){
      sh.getRange(i+1,2).setValue(newName); // B列（氏名）更新
      return { status:'success', id, newName };
    }
  }
  return { status:'error', message:'IDが見つかりません: '+id };
}

/***** ── 外部共有リンク ─────────────────*****/
function getExecUrlSafe_(){
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) return url;
  } catch(_e) {}
  try {
    const prop = PropertiesService.getScriptProperties().getProperty('EXEC_URL_FALLBACK');
    if (prop) return prop;
  } catch(_e) {}
  return '';
}

function buildExternalShareUrl_(token){
  const base = getExecUrlSafe_();
  if (!base) return '';
  const tok = String(token || '').trim();
  if (!tok) return '';
  return `${base}?shareId=${encodeURIComponent(tok)}`;
}

function parseQrDimensions_(value){
  const defaultMatch = String(SHARE_QR_SIZE || '220x220').match(/(\d+)x(\d+)/);
  let defaultWidth = 220;
  let defaultHeight = 220;
  if (defaultMatch) {
    defaultWidth = parseInt(defaultMatch[1], 10) || 220;
    defaultHeight = parseInt(defaultMatch[2], 10) || defaultWidth;
  }
  if (!value) return { width: defaultWidth, height: defaultHeight };
  const raw = String(value).trim();
  const match = raw.match(/(\d+)(?:x(\d+))?/i);
  if (match) {
    const width = parseInt(match[1], 10);
    const height = match[2] ? parseInt(match[2], 10) : width;
    if (!isNaN(width) && width > 0 && !isNaN(height) && height > 0) {
      return { width, height };
    }
  }
  const numeric = parseInt(raw, 10);
  if (!isNaN(numeric) && numeric > 0) {
    return { width: numeric, height: numeric };
  }
  return { width: defaultWidth, height: defaultHeight };
}

function buildExternalShareQrDataUrl_(url, size) {
  try {
    const trimmed = String(url || '').trim();
    if (!trimmed) return '';
    const dims = parseQrDimensions_(size || SHARE_QR_SIZE);
    const width = Math.max(1, Number(dims.width || 0) || 220);
    const height = Math.max(1, Number(dims.height || 0) || width);
    const encoded = encodeURIComponent(trimmed);
    return `https://chart.googleapis.com/chart?cht=qr&chs=${width}x${height}&choe=UTF-8&chl=${encoded}`;
  } catch (e) {
    Logger.log("buildExternalShareQrDataUrl_ error: " + e);
    return '';
  }
}


// 共有メタに加えて records を返す強化版
function getMemberRecords_(memberId, limit) {
  const SHEET_NAME = 'Monitoring'; 
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    Logger.log('❌ Records sheet "%s" not found', SHEET_NAME);
    return [];
  }

  const lr = sh.getLastRow(), lc = sh.getLastColumn();
  if (lr < 2) return [];

  const values = sh.getRange(1, 1, lr, lc).getValues();
  const header = values[0].map(v => String(v || '').trim());
  const data   = values.slice(1);

  const iDate   = header.indexOf('日付');
  const iMember = header.indexOf('利用者ID');
  const iKind   = header.indexOf('種別');
  const iText   = header.indexOf('記録内容');
  const iAtt    = header.indexOf('添付');

  Logger.log("🔎 header=%s", JSON.stringify(header));
  Logger.log("🔎 index: 日付=%s 利用者ID=%s", iDate, iMember);

  if (iDate < 0 || iMember < 0) {
    Logger.log("❌ 必須列が見つかりません");
    return [];
  }

  const wantId = String(memberId).trim();
  Logger.log("🔎 search memberId=%s", wantId);

  const out = [];
  for (let r = data.length - 1; r >= 0; r--) {
    const row = data[r];
    const got = String(row[iMember]).trim();
    if (got === wantId) {
      const rawDate = row[iDate];
      let ts = 0;
      if (rawDate instanceof Date) {
        ts = rawDate.getTime();
      } else {
        const d = new Date(rawDate);
        if (!isNaN(d.getTime())) ts = d.getTime();
      }

      // 添付
      let attachments = [];
      try {
        const raw = row[iAtt];
        if (raw && typeof raw === 'string') {
          const a = JSON.parse(raw);
          if (Array.isArray(a)) attachments = a;
        }
      } catch (_) {}

      out.push({
        recordId: String(r + 2), // 行番号
        timestamp: ts,
        dateText: ts ? Utilities.formatDate(new Date(ts), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd') : '',
        kind: row[iKind] || '',
        text: row[iText] || '',
        attachments
      });

      if (limit && out.length >= limit) break;
    }
  }

  out.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  Logger.log("📥 getMemberRecords_ returned count=%s", out.length);
  if (out.length) Logger.log("sample record=%s", JSON.stringify(out[0]));
  return out;
}






// Webアプリのリクエストパラメータを取得するユーティリティ
function getRequestParameters_() {
  try {
    return JSON.parse(HtmlService.createHtmlOutputFromFile('dummy')
      .getContent()); // ダミー: 実際は doGet(e) の e.parameter をグローバルに保持する設計が必要
  } catch (e) {
    return {};
  }
}





// ✅ 全角→半角にして「数字だけ」を返す（ゼロ埋めなし）
function toHalfWidthDigits(str) {
  if (str == null) return '';
  return String(str)
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)) // 全角数字→半角
    .replace(/[^0-9]/g, '') // 数字以外を除去
    .trim();
}

/** 呼び出し側の後方互換 */
function buildExternalShareQrUrl_(shareUrl, size){
  return buildExternalShareQrDataUrl_(shareUrl, size);
}

/***** QRコードをDriveに保存する *****/
function saveQrCodeToDrive_(memberId, shareUrl) {
  Logger.log("▶ saveQrCodeToDrive_ START: memberId=%s, shareUrl=%s", memberId, shareUrl);

  try {
    if (!memberId || !shareUrl) {
      Logger.log("❌ saveQrCodeToDrive_: 引数不足");
      return { ok: false };
    }

    // QRコード生成API
    const qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=220x220&data=' + encodeURIComponent(shareUrl);
    const resp  = UrlFetchApp.fetch(qrUrl);
    const blob  = resp.getBlob().setName(`QR_${toHalfWidthDigits(memberId)}.png`);

    const folder = DriveApp.getFolderById(QR_FOLDER_ID);

    // 同じファイル名を削除
    const existing = folder.getFilesByName(blob.getName());
    while (existing.hasNext()) existing.next().setTrashed(true);

    // 新規保存
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // 埋め込み用URL（<img src="...">で使える）
    // 通常のGoogle DriveファイルURL（クリック用）
    const embedUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w300";
    const viewUrl  = file.getUrl();

    Logger.log("✅ QR保存完了: fileId=%s", file.getId());
    Logger.log("✅ viewUrl=%s", viewUrl);
    Logger.log("✅ embedUrl=%s", embedUrl);

    return {
      ok: true,
      fileId: file.getId(),
      embedUrl,
      viewUrl
    };

  } catch (e) {
    Logger.log("❌ ERROR in saveQrCodeToDrive_: %s", e.stack || e);
    return { ok: false, error: e.message };
  }
}






// A列から memberId の行を探す（全角/半角差を吸収）
function findMemberRowById_(memberId, sh) {
  const want = toHalfWidthDigits(memberId);
  const last = sh.getLastRow();
  if (last < 1) return null;
  const vals = sh.getRange(1, 1, last, 1).getValues(); // A列
  for (let i = 0; i < vals.length; i++) {
    const got = toHalfWidthDigits(vals[i][0]);
    if (got && got === want) return i + 1; // 行番号
  }
  return null;
}

function getMemberQrDriveUrl_(memberId) {
  if (!memberId) return '';
  try {
    const row = findMemberRowById_(memberId);
    if (!row) return '';
    const sh = ensureMemberCenterHeaders_();
    const value = sh.getRange(row, 6).getValue();
    return String(value || '').trim();
  } catch (err) {
    Logger.log('⚠️ getMemberQrDriveUrl_ failed: ' + (err && err.message ? err.message : err));
    return '';
  }
}
/***** 共有リンクを発行する *****/
function createExternalShare(memberId, options) { 
  Logger.log("▶ createExternalShare START: memberId=%s options=%s", memberId, JSON.stringify(options));

  try {
    const normalizedId = normalizeMemberId_(memberId);
    const rawId = toHalfWidthDigits(memberId);
    const resolvedId = normalizedId || rawId;
    if (!resolvedId) throw new Error("利用者IDが未指定です");

    const shareSheet = shareGetSheet_();

    const config = options && typeof options === 'object' ? options : {};
    const audienceRaw = String(config.audience || '').trim().toLowerCase();
    const audienceList = ['family','center','medical','service','caremanager'];
    const audience = audienceList.includes(audienceRaw) ? audienceRaw : 'family';

    const maskMode = (config.maskMode === 'none') ? 'none' : 'simple';
    const passwordHash = hashSharePassword_(config.password);

    const token = Utilities.getUuid().replace(/-/g, '');
    const url = buildExternalShareUrl_(token);

    // 有効期限（例：10日後）
    let expiresAt = '';
    if (config.expiresInDays) {
      const expires = new Date();
      expires.setDate(expires.getDate() + Number(config.expiresInDays));
      expiresAt = expires.toISOString();
    } else if (config.expiresAt) {
      expiresAt = new Date(config.expiresAt).toISOString();
    }

    // 共有範囲
    const rangeSpec = shareNormalizeRangeInput_(config.rangeSpec || config.range || '30');

    const nowIso = new Date().toISOString();

    // 🔹 ExternalShares に必ず記録
    const appendRowIndex = shareSheet.getLastRow() + 1;
    shareSheet.appendRow([
      token,
      resolvedId,
      passwordHash,
      expiresAt,
      maskMode,
      JSON.stringify(config.allowedAttachments || []),
      nowIso,
      '',
      '',
      audience,
      0,
      rangeSpec,
      ''
    ]);

    // 🔹 QR保存（Google Driveに保存）
    let qrInfo = { ok: false };
    try {
      qrInfo = saveQrCodeToDrive_(resolvedId, url);
    } catch (err) {
      Logger.log("⚠️ saveQrCodeToDrive_ failed: %s", err.stack || err);
    }

    const qrRawUrl = qrInfo && (qrInfo.embedUrl || qrInfo.viewUrl)
      ? (qrInfo.embedUrl || qrInfo.viewUrl)
      : '';
    const qrEmbedUrl = shareNormalizeQrEmbedUrl_(qrRawUrl || url) || '';
    const qrColumnIndex = SHARE_SHEET_HEADERS.indexOf('QrUrl') + 1;
    if (qrColumnIndex > 0) {
      try {
        shareSheet.getRange(appendRowIndex, qrColumnIndex).setValue(qrEmbedUrl);
      } catch (err) {
        Logger.log("⚠️ failed to update share QR url: %s", err && err.message ? err.message : err);
      }
    }

    // 🔹 ほのぼのIDシートにも QRコードURL を反映
    try {
      const storedUrl = qrEmbedUrl || (qrInfo && qrInfo.viewUrl) || '';
      if (storedUrl) {
        updateHonobonoQrUrl_(resolvedId, storedUrl);
      }
    } catch (err) {
      Logger.log("⚠️ ほのぼのID への書き込み失敗: " + err);
    }

    return {
      status: 'success',
      shareLink: url,
      qrDriveUrl: qrEmbedUrl || (qrInfo && qrInfo.embedUrl) || "",
      qrViewUrl: qrInfo.viewUrl || ""
    };

  } catch (e) {
    Logger.log("❌ ERROR in createExternalShare: %s", e.stack || e);
    return { status: 'error', message: String(e) };
  }
}





function getExternalShares(memberId) {
  Logger.log('🟢 getExternalShares called with memberId=' + memberId);
  try {
    const id = String(memberId || '').trim();
    if (!id) throw new Error('利用者IDが未指定です');

    const sheet = shareGetSheet_();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { status: 'success', shares: [] };
    }

    const values = sheet.getRange(2, 1, lastRow - 1, SHARE_SHEET_HEADERS.length).getValues();
    const shares = [];
    const now = Date.now();
    const memberQrDriveUrl = getMemberQrDriveUrl_(id);

    values.forEach(row => {
      const share = shareParseShareRow_(row);
      if (!share.token || share.memberId !== id) return;
      if (share.revokedAt) return;

      const url = buildExternalShareUrl_(share.token);
      const qrDataUrl = buildExternalShareQrDataUrl_(url);

      shares.push({
        token: share.token,
        url,
        shareLink: url,
        qrDriveUrl: memberQrDriveUrl,
        qrUrl: qrDataUrl,
        qrDataUrl,
        qrCode: qrDataUrl,
        createdAtText: shareFormatDateTime_(share.createdAt),
        createdAtMs: share.createdAt ? share.createdAt.getTime() : 0,
        expiresAtText: shareFormatDateTime_(share.expiresAt),
        expired: share.expiresAt ? share.expiresAt.getTime() < now : false,
        audience: share.audience,
        passwordProtected: !!share.passwordHash,
        maskMode: share.maskMode,
        allowAllAttachments: share.allowAllAttachments,
        allowedCount: share.allowAllAttachments ? 0 : share.allowedAttachmentIds.length,
        lastAccessText: shareFormatDateTime_(share.lastAccessAt),
        remainingLabel: shareRemainingLabel_(share.expiresAt),
        accessCount: share.accessCount,
        rangeLabel: share.rangeLabel
      });
    });

    shares.sort((a, b) => (b.createdAtMs || 0) - (a.createdAtMs || 0));
    shares.forEach(s => { delete s.createdAtMs; });

    return { status: 'success', shares };
  } catch (err) {
    return { status: 'error', message: String(err && err.message || err) };
  }
}

function revokeExternalShare(token) {
  try {
    const context = shareFindByToken_(token);
    if (!context) throw new Error('対象の共有リンクが見つかりません');
    const revokedCol = SHARE_SHEET_HEADERS.indexOf('RevokedAt') + 1;
    context.sheet.getRange(context.rowIndex, revokedCol).setValue(new Date());
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: String(err && err.message || err) };
  }
}

const SHARE_SHEET_HEADERS = [
  'Token',
  'MemberID',
  'PasswordHash',
  'ExpiresAt',
  'MaskMode',
  'AllowedAttachments',
  'CreatedAt',
  'RevokedAt',
  'LastAccessAt',
  'Audience',
  'AccessCount',
  'RangeSpec',
  'QrUrl'
];

const SHARE_ALLOWED_AUDIENCES = ['family', 'center', 'medical', 'service', 'caremanager'];

function shareGetSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHARE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHARE_SHEET_NAME);
  }

  if (sheet.getMaxColumns() < SHARE_SHEET_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), SHARE_SHEET_HEADERS.length - sheet.getMaxColumns());
  }

  sheet.getRange(1, 1, 1, SHARE_SHEET_HEADERS.length).setValues([SHARE_SHEET_HEADERS]);
  return sheet;
}

function shareNormalizeRangeInput_(value) {
  const raw = String(value == null ? '' : value).trim().toLowerCase();
  if (!raw) return '30';
  if (raw === 'all' || raw === 'full' || raw === '0' || raw === 'alltime') return 'all';
  if (raw === 'month' || raw === 'monthly' || raw === 'latest-month') return 'month';
  if (raw === '90' || raw === '90d' || raw === '90days') return '90';
  if (raw === '30' || raw === '30d' || raw === '30days') return '30';
  const num = Number(raw);
  if (Number.isFinite(num)) {
    if (num <= 0) return '30';
    if (num >= 90) return '90';
    if (num >= 30) return '30';
  }
  return raw === '7' || raw === '7d' || raw === '7days' ? '7' : '30';
}

function shareNormalizeToken_(value) {
  return String(value == null ? '' : value)
    .replace(/[​‌‍﻿]/g, '')
    .replace(/^"+|"+$/g, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function shareParseDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const parsed = new Date(value);
  return isNaN(parsed.getTime()) ? null : parsed;
}

function shareParseAllowedAttachmentInfo_(value) {
  let raw = value;
  if (typeof raw === 'string') {
    try {
      raw = JSON.parse(raw);
    } catch (_err) {
      raw = [raw];
    }
  }

  const list = Array.isArray(raw) ? raw : [];
  const normalized = list.map(v => String(v || '').trim()).filter(Boolean);
  const allowAll = normalized.includes('__ALL__');
  const ids = normalized.filter(v => v && v !== '__ALL__');
  return { allowAll, ids };
}

function shareNormalizeAudience_(value) {
  const raw = String(value || '').trim().toLowerCase();
  return SHARE_ALLOWED_AUDIENCES.includes(raw) ? raw : 'family';
}

function shareParseRangeSpec_(spec) {
  const normalized = shareNormalizeRangeInput_(spec);
  if (normalized === 'all') {
    return { type: 'all' };
  }
  if (normalized === 'month') {
    return { type: 'month' };
  }
  const days = Number(normalized);
  return { type: 'days', days: Number.isFinite(days) && days > 0 ? days : 30 };
}

function shareRangeLabel_(range) {
  if (!range) return '直近30日';
  if (range.type === 'all') return '全期間';
  if (range.type === 'month') return '月次モニタリング';
  if (range.days >= 90) return '直近90日';
  if (range.days <= 7) return '直近7日';
  return '直近30日';
}

function shareNormalizeQrEmbedUrl_(url) {
  const raw = String(url || '').trim();
  if (!raw) return '';
  if (/^https?:\/\//i.test(raw) === false) return raw;
  if (raw.includes('drive.google.com/thumbnail')) return raw;
  const idMatch = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/) || raw.match(/\/d\/([a-zA-Z0-9_-]+)\//);
  if (idMatch && idMatch[1]) {
    return `https://drive.google.com/thumbnail?id=${idMatch[1]}&sz=w300`;
  }
  return raw;
}

function monitoringReportsGetSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(MONITORING_REPORTS_SHEET_NAME) || null;
}

function monitoringReportsBuildHeaderIndex_(headerRow) {
  const map = {};
  headerRow.forEach((label, idx) => {
    const key = String(label || '').trim().toLowerCase();
    if (key) map[key] = idx;
  });
  return map;
}

function monitoringReportsParseMonth_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), 1);
  }
  const raw = String(value).trim();
  if (!raw) return null;
  const replaced = raw
    .replace(/年|\.|\//g, '-')
    .replace(/月/g, '')
    .replace(/[^0-9-]/g, '-');
  const match = replaced.match(/(\d{4})-(\d{1,2})/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return null;
  return new Date(year, month - 1, 1);
}

function monitoringReportsFormatMonthLabel_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return `${date.getFullYear()}年${date.getMonth() + 1}月`;
}

function monitoringReportsFindLatest_(memberId) {
  const normalizedId = normalizeMemberId_(memberId);
  if (!normalizedId) return null;
  const sheet = monitoringReportsGetSheet_();
  if (!sheet) return null;
  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return null;
  const header = values[0];
  const idx = monitoringReportsBuildHeaderIndex_(header);
  const idxMember = idx.memberid;
  const idxMonth = idx.month;
  const idxText = idx.reporttext;
  const idxGenerated = idx.generatedat;
  const idxStatus = idx.status;
  const idxSpecial = idx.special != null ? idx.special
    : (idx['ai要約'] != null ? idx['ai要約']
      : (idx['aisummary'] != null ? idx['aisummary']
        : (idx['ai_summary'] != null ? idx['ai_summary'] : null)));
  if (idxMember == null || idxText == null) return null;

  let latest = null;
  let latestKey = -Infinity;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const idRaw = normalizeMemberId_(row[idxMember]);
    if (!idRaw || idRaw !== normalizedId) continue;
    const monthDate = idxMonth != null ? monitoringReportsParseMonth_(row[idxMonth]) : null;
    const generatedAt = idxGenerated != null ? shareParseDate_(row[idxGenerated]) : null;
    const fallbackKey = generatedAt ? generatedAt.getTime() : r;
    const candidateKey = monthDate ? monthDate.getTime() : fallbackKey;
    if (candidateKey > latestKey) {
      latestKey = candidateKey;
      latest = {
        memberId: normalizedId,
        monthDate,
        monthLabel: monitoringReportsFormatMonthLabel_(monthDate),
        monthRaw: idxMonth != null ? row[idxMonth] : '',
        reportText: String(row[idxText] || ''),
        generatedAt,
        generatedAtText: shareFormatDateTime_(generatedAt),
        status: idxStatus != null ? String(row[idxStatus] || '') : '',
        special: idxSpecial != null ? String(row[idxSpecial] || '') : ''
      };
    }
  }
  return latest;
}

function shareLoadMonitoringReport_(share) {
  if (!share || !share.memberId) return null;
  const audience = String(share.audience || '').toLowerCase();
  if (audience !== 'caremanager') return null;
  if (!share.range || share.range.type !== 'month') return null;
  const report = monitoringReportsFindLatest_(share.memberId);
  if (!report) return null;
  const text = String(report.reportText || '');
  return {
    monthLabel: report.monthLabel || String(report.monthRaw || ''),
    generatedAtText: report.generatedAtText || '',
    status: report.status || '',
    special: report.special || '',
    reportText: text,
  };
}

function shareFormatDateTime_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  const tz = Session.getScriptTimeZone ? (Session.getScriptTimeZone() || 'Asia/Tokyo') : 'Asia/Tokyo';
  return Utilities.formatDate(date, tz, 'yyyy/MM/dd HH:mm');
}

function shareRemainingLabel_(expiresAt) {
  if (!(expiresAt instanceof Date) || isNaN(expiresAt.getTime())) return '';
  const diff = expiresAt.getTime() - Date.now();
  if (diff <= 0) return '';
  const minutes = Math.floor(diff / (60 * 1000));
  if (minutes < 60) return `残り約${minutes}分`;
  const hours = Math.floor(minutes / 60);
  if (hours < 48) return `残り約${hours}時間`;
  const days = Math.floor(hours / 24);
  return `残り約${days}日`;
}

function shareNormalizeAttachment_(attachment) {
  if (!attachment || typeof attachment !== 'object') return null;
  const fileId = String(attachment.fileId || attachment.id || '').trim();
  const url = String(attachment.url || (fileId ? `https://drive.google.com/file/d/${fileId}/view` : '')).trim();
  const name = String(attachment.name || attachment.fileName || attachment.title || '添付ファイル');
  if (!fileId && !url) return null;
  return {
    fileId,
    url,
    name,
    mimeType: String(attachment.mimeType || '')
  };
}

function shareFilterAttachmentsForShare_(attachments, share) {
  const list = Array.isArray(attachments) ? attachments : [];
  if (share.allowAllAttachments) {
    return list.map(shareNormalizeAttachment_).filter(Boolean);
  }
  if (!share.allowedAttachmentIds.length) return [];
  const allowedSet = new Set(share.allowedAttachmentIds);
  return list
    .map(shareNormalizeAttachment_)
    .filter(att => att && att.fileId && allowedSet.has(att.fileId));
}

function shareMaskText_(text, mode) {
  if (mode === 'none') return String(text || '');
  const value = String(text || '');
  return value
    .replace(/[0-9０-９]/g, '＊')
    .replace(/([A-Za-z぀-ヿ一-鿿]{2,})/g, match => match.charAt(0) + '＊'.repeat(match.length - 1));
}

function shareRangeWindow_(range) {
  if (!range || range.type === 'all') {
    return { since: null };
  }
  const days = range.days || 30;
  const now = new Date();
  const since = now.getTime() - days * 24 * 60 * 60 * 1000;
  return { since };
}

function shareLoadRecords_(share, recordId, range) {
  const normalizedRecordId = String(recordId || '').trim();
  const allRecords = getMemberRecords_(share.memberId, 200) || [];
  const { since } = shareRangeWindow_(range);
  const sanitized = [];

  allRecords.forEach(rec => {
    const ts = Number(rec.timestamp || 0) || 0;
    if (since && ts && ts < since) return;
    if (normalizedRecordId && String(rec.recordId || '').trim() !== normalizedRecordId) return;

    const attachments = shareFilterAttachmentsForShare_(rec.attachments, share);
    const maskedText = shareMaskText_(rec.text || '', share.maskMode);

    sanitized.push({
      recordId: rec.recordId,
      memberId: share.memberId,
      memberName: rec.memberName || '',
      dateText: rec.dateText || '',
      kind: rec.kind || '',
      audience: share.audience,
      text: maskedText,
      attachments,
      timestamp: ts || null,
      center: rec.center || (rec.fields && rec.fields.center) || '',
      staff: rec.staff || (rec.fields && rec.fields.staff) || '',
      status: rec.status || '',
      special: rec.special || '',
      fields: rec.fields || {}
    });
  });

  sanitized.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  let primaryRecord = sanitized.length ? sanitized[0] : null;
  if (normalizedRecordId) {
    const found = sanitized.find(rec => String(rec.recordId || '').trim() === normalizedRecordId);
    if (found) primaryRecord = found;
  }

  return { records: sanitized, primaryRecord };
}

function shareResolveProfile_(memberId) {
  try {
    if (typeof honobonoFindById_ === 'function') {
      const info = honobonoFindById_(memberId);
      if (info) {
        return {
          name: String(info.name || ''),
          center: String(info.center || ''),
          staff: String(info.staff || ''),
          qrUrl: String(info.qrUrl || '')
        };
      }
    }
  } catch (err) {
    Logger.log('⚠️ shareResolveProfile_ error: ' + (err && err.message ? err.message : err));
  }
  return { name: '', center: '', staff: '', qrUrl: '' };
}

function sharePickFirstString_(...values) {
  for (const value of values) {
    if (typeof value !== 'string') continue;
    const trimmed = value.trim();
    if (trimmed) return trimmed;
  }
  return '';
}

function shareNormalizeQrCandidate_(value) {
  if (typeof value !== 'string') return '';
  const trimmed = value.trim();
  if (!trimmed) return '';
  return shareNormalizeQrEmbedUrl_(trimmed) || trimmed;
}

function shareBuildQrPayload_(share, profile) {
  const safeShare = share || {};
  const safeProfile = profile || {};
  const baseUrl = buildExternalShareUrl_(safeShare.token);
  const shareLink = sharePickFirstString_(safeShare.shareLink, safeShare.url, baseUrl);
  const computedDriveUrl = getMemberQrDriveUrl_(safeShare.memberId);
  const qrDriveUrl = sharePickFirstString_(
    safeShare.qrDriveUrl,
    safeShare.qrUrl,
    safeProfile.qrUrl,
    computedDriveUrl
  );

  const embedCandidates = [
    safeShare.qrEmbedUrl,
    safeShare.qrDataUrl,
    safeShare.qrCode,
    qrDriveUrl
  ];

  let qrEmbedUrl = '';
  for (const candidate of embedCandidates) {
    const normalized = shareNormalizeQrCandidate_(candidate);
    if (normalized) {
      qrEmbedUrl = normalized;
      break;
    }
  }

  const qrDataUrl = shareLink ? buildExternalShareQrDataUrl_(shareLink) : '';
  const finalEmbed = sharePickFirstString_(qrEmbedUrl, qrDataUrl);
  const resolvedQrUrl = sharePickFirstString_(finalEmbed, qrDriveUrl, qrDataUrl, shareLink);
  const qrCode = sharePickFirstString_(qrDataUrl, finalEmbed, resolvedQrUrl);

  return {
    url: sharePickFirstString_(shareLink, baseUrl),
    shareLink: sharePickFirstString_(shareLink, baseUrl),
    qrDriveUrl: qrDriveUrl,
    qrEmbedUrl: finalEmbed,
    qrDataUrl,
    qrUrl: resolvedQrUrl,
    qrCode
  };
}

function shareBuildSummary_(share, profile, hasRecords) {
  const qrPayload = shareBuildQrPayload_(share, profile);
  const url = sharePickFirstString_(qrPayload.url, buildExternalShareUrl_(share.token));
  const shareLink = sharePickFirstString_(qrPayload.shareLink, url);
  const expired = share.expiresAt ? share.expiresAt.getTime() < Date.now() : false;

  const memberName = profile.name || share.memberName || '';
  const memberCenter = profile.center || profile.centerName || share.memberCenter || share.centerName || '';
  const memberStaff = profile.staff || profile.staffName || share.memberStaff || share.staffName || '';

  return {
    token: share.token,
    memberId: share.memberId,
    memberName,
    memberCenter,
    memberStaff,
    expiresAtText: shareFormatDateTime_(share.expiresAt),
    expired,
    audience: share.audience,
    rangeSpec: share.rangeSpec,
    rangeType: share.range ? share.range.type : '',
    requirePassword: !!share.passwordHash,
    maskMode: share.maskMode,
    allowAllAttachments: share.allowAllAttachments,
    allowedCount: share.allowAllAttachments ? 0 : share.allowedAttachmentIds.length,
    remainingLabel: shareRemainingLabel_(share.expiresAt),
    rangeLabel: share.rangeLabel,
    url,
    shareLink,
    qrDriveUrl: qrPayload.qrDriveUrl,
    qrEmbedUrl: sharePickFirstString_(qrPayload.qrEmbedUrl, qrPayload.qrDataUrl),
    qrUrl: sharePickFirstString_(qrPayload.qrUrl, qrPayload.qrEmbedUrl, qrPayload.qrDataUrl, shareLink),
    qrDataUrl: qrPayload.qrDataUrl,
    qrCode: sharePickFirstString_(qrPayload.qrCode, qrPayload.qrDataUrl),
    hasRecords: hasRecords
  };
}

function shareParseShareRow_(row) {
  const tokenRaw = String(row[0] || '').trim();
  const attachmentInfo = shareParseAllowedAttachmentInfo_(row[5]);
  const rangeSpec = shareNormalizeRangeInput_(row[11]);
  const range = shareParseRangeSpec_(rangeSpec);
  return {
    token: tokenRaw,
    normalizedToken: shareNormalizeToken_(tokenRaw),
    memberId: String(row[1] || '').trim(),
    passwordHash: String(row[2] || '').trim(),
    expiresAt: shareParseDate_(row[3]),
    maskMode: String(row[4] || 'simple').trim() === 'none' ? 'none' : 'simple',
    allowedAttachmentIds: attachmentInfo.ids,
    allowAllAttachments: attachmentInfo.allowAll,
    createdAt: shareParseDate_(row[6]),
    revokedAt: shareParseDate_(row[7]),
    lastAccessAt: shareParseDate_(row[8]),
    audience: shareNormalizeAudience_(row[9]),
    accessCount: Number(row[10] || 0) || 0,
    rangeSpec,
    range,
    rangeLabel: shareRangeLabel_(range),
    qrUrl: String(row[12] || '').trim()
  };
}

function shareFindByToken_(token) {
  const normalized = shareNormalizeToken_(token);
  if (!normalized) return null;
  const sheet = shareGetSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, SHARE_SHEET_HEADERS.length).getValues();
  for (let i = 0; i < values.length; i++) {
    const share = shareParseShareRow_(values[i]);
    if (share.normalizedToken && share.normalizedToken === normalized) {
      return { sheet, rowIndex: i + 2, share };
    }
  }
  return null;
}

function shareGetLogSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const name = SHARE_LOG_SHEET_NAME || 'ExternalShareAccessLog';
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  const header = ['AccessedAt', 'Token', 'MemberID', 'Result'];
  if (sheet.getMaxColumns() < header.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), header.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  return sheet;
}

function shareLogAccess_(token, memberId, result) {
  try {
    if (!token) return;
    const sheet = shareGetLogSheet_();
    sheet.appendRow([new Date(), token, String(memberId || ''), String(result || '')]);
  } catch (err) {
    Logger.log('⚠️ shareLogAccess_ failed: ' + (err && err.message ? err.message : err));
  }
}

function shareUpdateAccessStats_(sheet, rowIndex) {
  try {
    const lastAccessCol = SHARE_SHEET_HEADERS.indexOf('LastAccessAt') + 1;
    const accessCountCol = SHARE_SHEET_HEADERS.indexOf('AccessCount') + 1;
    sheet.getRange(rowIndex, lastAccessCol).setValue(new Date());
    const range = sheet.getRange(rowIndex, accessCountCol);
    const current = Number(range.getValue() || 0);
    range.setValue(current + 1);
  } catch (err) {
    Logger.log('⚠️ shareUpdateAccessStats_ failed: ' + (err && err.message ? err.message : err));
  }
}


function shareBuildCustomRecordSet_(share, rawRecords, recordId) {
  const normalizedRecordId = String(recordId || '').trim();
  const base = Array.isArray(rawRecords) ? rawRecords : [];
  const sanitized = base.map(rec => {
    const attachments = shareFilterAttachmentsForShare_(rec.attachments, share);
    const masked = shareMaskText_(rec.text || '', share.maskMode);
    return {
      recordId: rec.recordId,
      memberId: rec.memberId || share.memberId,
      memberName: rec.memberName || '',
      dateText: rec.dateText || '',
      kind: rec.kind || '',
      audience: share.audience,
      text: masked,
      attachments,
      timestamp: Number(rec.timestamp || 0) || null,
      center: rec.center || (rec.fields && rec.fields.center) || '',
      staff: rec.staff || (rec.fields && rec.fields.staff) || '',
      status: rec.status || '',
      special: rec.special || '',
      fields: rec.fields || {}
    };
  });

  sanitized.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  let primaryRecord = sanitized.length ? sanitized[0] : null;
  if (normalizedRecordId) {
    const found = sanitized.find(item => String(item.recordId || '').trim() === normalizedRecordId);
    if (found) primaryRecord = found;
  }

  return { records: sanitized, primaryRecord };
}
function shareBuildResponse_(share, recordId, includeRecords, includeReport) {
  const recordResult = shareLoadRecords_(share, recordId, share.range);
  const profile = shareResolveProfile_(share.memberId);
  const summary = shareBuildSummary_(share, profile, recordResult.records.length > 0);
  const message = recordResult.records.length ? '' : '記録が存在しません';
  const report = includeReport ? shareLoadMonitoringReport_(share) : null;
  const records = includeRecords ? recordResult.records : [];
  const primaryRecord = includeRecords ? recordResult.primaryRecord : null;
  return {
    summary,
    records,
    rawRecords: records.slice(),
    primaryRecord,
    message,
    recordCount: recordResult.records.length,
    report,
    url: summary.url,
    shareLink: summary.shareLink,
    qrUrl: summary.qrUrl,
    qrEmbedUrl: summary.qrEmbedUrl,
    qrDriveUrl: summary.qrDriveUrl,
    qrDataUrl: summary.qrDataUrl,
    qrCode: summary.qrCode
  };
}

function getExternalShareMeta(token, recordId) {
  Logger.log('🟦 getExternalShareMeta called token="%s" recordId="%s"', token, recordId);
  try {
    const context = shareFindByToken_(token);
    if (!context) {
      shareLogAccess_(shareNormalizeToken_(token), '', 'invalid');
      return { status: 'error', message: '共有リンクが存在しません' };
    }

    const share = context.share;
    if (share.revokedAt) {
      shareLogAccess_(share.token, share.memberId, 'invalid');
      return { status: 'error', message: 'この共有リンクは無効化されています。' };
    }

    const includeRecords = !share.passwordHash;
    const includeReport = includeRecords;
    const response = shareBuildResponse_(share, recordId, includeRecords, includeReport);

    return {
      status: 'success',
      share: response.summary,
      records: response.records,
      rawRecords: response.rawRecords || response.records,
      primaryRecord: response.primaryRecord,
      report: response.report,
      message: includeRecords ? response.message : '',
      url: response.url,
      qrUrl: response.qrUrl,
      qrEmbedUrl: response.qrEmbedUrl,
      qrDriveUrl: response.qrDriveUrl,
      qrDataUrl: response.qrDataUrl,
      qrCode: response.qrCode,
      shareLink: response.shareLink
    };
  } catch (err) {
    Logger.log('❌ getExternalShareMeta failed: ' + (err && err.stack ? err.stack : err));
    return { status: 'error', message: String(err && err.message || err) };
  }
}

function enterExternalShare(token, password, recordId) {
  Logger.log('🟦 enterExternalShare called token="%s" recordId="%s"', token, recordId);
  try {
    const context = shareFindByToken_(token);
    if (!context) {
      shareLogAccess_(shareNormalizeToken_(token), '', 'invalid');
      return { status: 'error', message: '共有リンクが存在しません' };
    }

    const share = context.share;
    if (share.revokedAt) {
      shareLogAccess_(share.token, share.memberId, 'invalid');
      return { status: 'error', message: '共有リンクは無効化されています' };
    }

    if (share.passwordHash) {
      const hash = hashSharePassword_(password);
      if (!hash || hash !== share.passwordHash) {
        shareLogAccess_(share.token, share.memberId, 'invalid');
        return { status: 'error', message: 'パスワードが一致しません。' };
      }
    }

    const response = shareBuildResponse_(share, recordId, true, true);
    shareUpdateAccessStats_(context.sheet, context.rowIndex);
    shareLogAccess_(share.token, share.memberId, 'success');

    return {
      status: 'success',
      share: response.summary,
      records: response.records,
      rawRecords: response.rawRecords || response.records,
      primaryRecord: response.primaryRecord,
      report: response.report,
      message: response.message,
      url: response.url,
      qrUrl: response.qrUrl,
      qrEmbedUrl: response.qrEmbedUrl,
      qrDriveUrl: response.qrDriveUrl,
      qrDataUrl: response.qrDataUrl,
      qrCode: response.qrCode,
      shareLink: response.shareLink
    };
  } catch (err) {
    Logger.log('❌ enterExternalShare failed: ' + (err && err.stack ? err.stack : err));
    return { status: 'error', message: String(err && err.message || err) };
  }
}

function hashSharePassword_(password){
  const value = String(password || '');
  if (!value) return '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function updateRecord(rowIndex, newText){
  try {
    const payload = { rowIndex: Number(rowIndex), record: String(newText || '') };
    // 既存の updateMonitoringRecord は center/staff/status/special を想定しているので
    // 本文だけは Monitoring シートの「記録内容」列を書き換える軽量版を用意
    return updateMonitoringRecordBodyOnly_(payload);
  } catch (e) {
    throw new Error('更新に失敗しました: ' + (e && e.message ? e.message : e));
  }
}

/** 既存UI互換：行番号のみで削除 */
function deleteRecord(rowIndex){
  try {
    return deleteMonitoringRecord({ rowIndex: Number(rowIndex) });
  } catch (e) {
    throw new Error('削除に失敗しました: ' + (e && e.message ? e.message : e));
  }
}

/** 本文（記録内容）だけを書き換える内部用：Monitoring シートの列検出を使う */
function updateMonitoringRecordBodyOnly_(data){
  const payload = data && typeof data === 'object' ? data : {};
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`シートが見つかりません: ${SHEET_NAME}`);
  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) throw new Error('記録が存在しません');

  const header = values[0].map(v => String(v || '').trim());
  const indexes = resolveRecordColumnIndexes_(header);
  const rowIndex = Number(payload.rowIndex || 0);
  if (!rowIndex || rowIndex < 2) throw new Error('対象の記録が見つかりません');

  const bodyCol = indexes.record >= 0 ? (indexes.record + 1) : 0;
  if (!bodyCol) throw new Error('「記録内容」列が見つかりません');

  sheet.getRange(rowIndex, bodyCol).setValue(String(payload.record || ''));
  return { status:'success', rowIndex };
}
/** 指定した memberId でレコードが取れるか確認 */
function test_fetchRecords() {
  const memberId = "5745";   // ← 問題のIDに差し替え
  const recs = fetchRecordsWithIndex_(memberId, 30); // 直近30日
  Logger.log("✅ fetchRecords length = " + recs.length);
  if (recs.length) {
    Logger.log("📄 first record = " + JSON.stringify(recs[0], null, 2));
  }
}

/** 取得・保存・削除 */
function getMemberCenterInfo(memberIdRaw) {
  const safeId = normalizeMemberId_(memberIdRaw);
  const row = findMemberRowById_(safeId);
  if (!row) return { ok:false, message:'対象のIDが見つかりません: ' + safeId };
  const sh = ensureMemberCenterHeaders_();
  return {
    ok: true,
    id: safeId,
    center: String(sh.getRange(row, 4).getValue() || '').trim(), // D
    staff:  String(sh.getRange(row, 5).getValue() || '').trim()  // E
  };
}

function saveMemberCenterInfo(memberIdRaw, center, staff) {
  const safeId = normalizeMemberId_(memberIdRaw);
  const row = findMemberRowById_(safeId);
  if (!row) return { ok:false, message:'対象のIDが見つかりません: ' + safeId };
  const sh = ensureMemberCenterHeaders_();
  const centerSafe = String(center || '').trim();
  const staffSafe = String(staff  || '').trim();
  sh.getRange(row, 4).setValue(centerSafe); // D=センター
  sh.getRange(row, 5).setValue(staffSafe); // E=担当者
  return { ok:true, id: safeId, center: centerSafe, staff: staffSafe };
}
/** ほのぼのIDシート: D=センター, E=担当者 を保証 */
function ensureMemberCenterHeaders_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName('ほのぼのID');
  if (!sh) throw new Error('シート「ほのぼのID」が見つかりません');
  // 少なくともF列まで用意
  if (sh.getMaxColumns() < 6) sh.insertColumnsAfter(sh.getMaxColumns(), 6 - sh.getMaxColumns());
  // ヘッダを確定（A=ID, B=氏名 は触らない）
  sh.getRange(1, 4).setValue('包括支援センター'); // D1
  sh.getRange(1, 5).setValue('担当者名');         // E1
  sh.getRange(1, 6).setValue('共有QRコードURL');  // F1
  return sh;
}
/** 行番号を A列（ほのぼのID）だけで厳密に探す */
function findMemberRowById_(memberIdRaw) {
  const id = normalizeMemberId_(memberIdRaw);  // "5767" などに正規化
  if (!id) return 0;
  const sh = ensureMemberCenterHeaders_();
  const vals = sh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    const cellId = normalizeMemberId_(vals[i][0]); // A列のみを見る
    if (cellId && cellId === id) return i + 1;     // 1-based
  }
  return 0;
}
function clearMemberCenterInfo(memberIdRaw) {
  const safeId = normalizeMemberId_(memberIdRaw);
  const row = findMemberRowById_(safeId);
  if (!row) return { ok:false, message:'対象のIDが見つかりません: ' + safeId };
  const sh = ensureMemberCenterHeaders_();
  sh.getRange(row, 4, 1, 2).clearContent(); // D,E を空に
  return { ok:true, id: safeId };
}

// 「ほのぼのID / ほのぼのＩＤ」どちらでも取得
function getHonobonoSheet_(ss) {
  const candidates = ['ほのぼのID', 'ほのぼのＩＤ'];
  for (const name of candidates) {
    const sh = ss.getSheetByName(name);
    if (sh) return sh;
  }
  throw new Error('シート「ほのぼのID（全角/半角）」が見つかりません');
}
// ほのぼのIDシートにQRコードURLを保存
function updateHonobonoQrUrl_(memberId, qrUrl){
  if (!memberId || !qrUrl) return;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(HONOBONO_SHEET_NAME);
  if (!sh) return;

  const last = sh.getLastRow();
  if (last < 2) return;

  const target = normalizeMemberId_(memberId);
  const ids = sh.getRange(2, 1, last - 1, 1).getValues(); // A列: 利用者ID
  for (let i = 0; i < ids.length; i++){
    if (normalizeMemberId_(ids[i][0]) === target){
      sh.getRange(i + 2, HONOBONO_QR_URL_COL).setValue(qrUrl);
      break;
    }
  }
}
/** '90' | '30' | '7' | 'all' を安全に解釈（空はデフォ90） */
function parseRangeSpec_(val) {
  const s = String(val == null ? '' : val).trim().toLowerCase();
  if (!s) return { type: 'days', days: 90 };
  if (s === 'all' || s === 'full' || s === '0' || s === 'alltime') return { type: 'all' };
  const n = Number(s);
  if (Number.isFinite(n) && n > 0) return { type: 'days', days: Math.floor(n) };
  return { type: 'days', days: 90 };
}

/** 共有の表示範囲から [sinceTs, untilTs] を返す（JSTの「日」境界で丸め） */
function getDateRangeForShare_(rangeSpec) {
  const tz = Session.getScriptTimeZone ? (Session.getScriptTimeZone() || 'Asia/Tokyo') : 'Asia/Tokyo';
  const now = new Date();
  const untilLocal = new Date(Utilities.formatDate(now, tz, 'yyyy/MM/dd 23:59:59')); // きょうの終端
  if (rangeSpec.type === 'all') {
    return { sinceTs: 0, untilTs: untilLocal.getTime() };
  }
  const days = rangeSpec.days || 90;
  const sinceLocal = new Date(untilLocal.getTime() - (days - 1) * 24 * 3600 * 1000); // 例：90日なら今日を含めて過去89日分
  // 始端は 00:00:00 に丸め
  const sinceText = Utilities.formatDate(sinceLocal, tz, 'yyyy/MM/dd 00:00:00');
  const sinceTs = new Date(sinceText).getTime();
  return { sinceTs, untilTs: untilLocal.getTime() };
}
/**
 * 記録シート（例: MonitoringRecords）から指定 MemberID の最近200件を取得。
 * 必要に応じて列名・シート名を実環境に合わせてください。
 *
 * 期待する列:
 * - MemberID
 * - Date        : yyyy/MM/dd あるいは ISO 文字列
 * - Kind        : 任意（「訪問」「電話」など）
 * - Center      : 地域包括支援センター名
 * - Staff       : 担当者名
 * - Text        : 本文
 * - Status      : 状態・経過（任意）
 * - Special     : 特記事項（任意）
 * - Attachments : JSON配列文字列 [{"name":"xxx","url":"https://..."}]
 */
function getMemberRecords_(memberId, limit) {
  const SHEET_NAME = 'Monitoring';
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    Logger.log('❌ Records sheet "%s" not found', SHEET_NAME);
    return [];
  }

  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 2) return [];

  const values = sh.getRange(1, 1, lr, lc).getValues();
  const header = values[0].map(v => String(v || '').trim());
  const data = values.slice(1);

  const iDate = header.indexOf('日付');
  const iMember = header.indexOf('利用者ID');
  const iKind = header.indexOf('種別');
  const iText = header.indexOf('記録内容');
  const iAtt = header.indexOf('添付');

  Logger.log("🔎 header=%s", JSON.stringify(header));
  Logger.log("🔎 index: 日付=%s 利用者ID=%s", iDate, iMember);
  Logger.log("📥 getMemberRecords_ scanning rows=%s memberId=%s", data.length, memberId);

  if (iDate < 0 || iMember < 0) {
    Logger.log('❌ 必須列が見つかりません');
    return [];
  }

  const wantId = String(memberId).trim();
  Logger.log("🔎 search memberId=%s", wantId);

  const out = [];
  for (let r = data.length - 1; r >= 0; r--) {
    const row = data[r];
    const got = String(row[iMember]).trim();
    if (got) {
      Logger.log("… row %s 利用者ID=%s", r + 2, got);
    }
    if (String(row[iMember]).trim() === wantId) {
      Logger.log("✅ HIT row %s", r + 2);

      const rawDate = row[iDate];
      let timestamp = 0;
      if (rawDate instanceof Date) {
        timestamp = rawDate.getTime();
      } else {
        const parsed = new Date(rawDate);
        if (!isNaN(parsed.getTime())) {
          timestamp = parsed.getTime();
        }
      }

      let attachments = [];
      try {
        const rawAttachments = row[iAtt];
        if (rawAttachments && typeof rawAttachments === 'string') {
          const parsedAttachments = JSON.parse(rawAttachments);
          if (Array.isArray(parsedAttachments)) {
            attachments = parsedAttachments;
          }
        }
      } catch (e) {
        const message = e && e.message ? e.message : e;
        Logger.log('⚠️ attachments parse error row %s: %s', r + 2, message);
      }

      out.push({
        recordId: String(r + 2),
        timestamp,
        dateText: timestamp
          ? Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy/MM/dd')
          : '',
        kind: iKind >= 0 ? row[iKind] || '' : '',
        text: iText >= 0 ? row[iText] || '' : '',
        attachments
      });

      if (limit && out.length >= limit) {
        break;
      }
    }
  }

  out.sort((a, b) => (b.timestamp || 0) - (a.timestamp || 0));
  Logger.log("✅ getMemberRecords_: memberId=%s hit=%s", memberId, out.length);
  if (out.length) Logger.log("sample record=%s", JSON.stringify(out[0]));
  return out;
}



if (typeof __honobonoCacheMap === 'undefined') {
  var __honobonoCacheMap = null;
}

function honobonoOpenSheet_() {
  // 既存の SPREADSHEET_ID / HONOBONO_SHEET_NAME をそのまま使用
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(HONOBONO_SHEET_NAME);
}

function honobonoBuildHeaderIndex_(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => {
    const key = String(h || '').trim();
    if (key) map[key] = i;
  });
  return map;
}

function honobonoAt_(row, headerIndexMap, headerName) {
  if (headerIndexMap[headerName] != null) {
    return String(row[headerIndexMap[headerName]] || '').trim();
  }
  // 見出しが無い場合のみ、指定列番号でフォールバック（例：QRは F列 = 6）
  if (headerName === HONOBONO_QR_HEADER && typeof HONOBONO_QR_URL_COL === 'number') {
    const idx = HONOBONO_QR_URL_COL - 1; // 1-based -> 0-based
    return String(row[idx] || '').trim();
  }
  return '';
}

/** ほのぼのIDマスタ全件を Map(id -> info) で取得（キャッシュ付） */
function honobonoGetMasterMap_() {
  if (__honobonoCacheMap) return __honobonoCacheMap;

  const sh = honobonoOpenSheet_();
  if (!sh) {
    console.warn('⚠ ほのぼのIDシートが見つかりません:', HONOBONO_SHEET_NAME);
    __honobonoCacheMap = new Map();
    return __honobonoCacheMap;
  }

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    __honobonoCacheMap = new Map();
    return __honobonoCacheMap;
  }

  const header = values[0];
  const idx = honobonoBuildHeaderIndex_(header);

  const map = new Map();
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const id = honobonoAt_(row, idx, HONOBONO_MEMBER_ID_HEADER);
    if (!id) continue;
    map.set(id, {
      memberId: id,
      name:  honobonoAt_(row, idx, HONOBONO_NAME_HEADER),
      kana:  honobonoAt_(row, idx, HONOBONO_KANA_HEADER),
      center:honobonoAt_(row, idx, HONOBONO_CENTER_HEADER),
      staff: honobonoAt_(row, idx, HONOBONO_STAFF_HEADER),
      qrUrl: honobonoAt_(row, idx, HONOBONO_QR_HEADER),
    });
  }
  __honobonoCacheMap = map;
  return __honobonoCacheMap;
}

/** IDで1件取得（無ければ null） */
function honobonoFindById_(memberId) {
  const map = honobonoGetMasterMap_();
  return map.get(String(memberId)) || null;
}

/**
 * 共有オブジェクトにマスタ情報を上書き注入する（破壊的）
 * - 既に share に値があればそれを優先し、空の場合のみマスタで補完
 */
function honobonoEnrichShare_(share, memberId) {
  try {
    const m = honobonoFindById_(memberId);
    if (!m) return share;
    share.memberId     = share.memberId     || m.memberId;
    share.memberName   = share.memberName   || m.name;
    share.memberKana   = share.memberKana   || m.kana;
    share.memberCenter = share.memberCenter || m.center;
    share.memberStaff  = share.memberStaff  || m.staff;
    share.qrUrl        = share.qrUrl        || m.qrUrl;
    return share;
  } catch (e) {
    console.warn('honobonoEnrichShare_ error:', e);
    return share;
  }

///半角数字に変換///
}
function convertFullWidthToHalfWidth() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("ほのぼのID"); // 対象シートを指定
  if (!sheet) {
    throw new Error("シート『ほのぼのID』が見つかりません。");
  }
  
  const range = sheet.getRange("A:A"); // A列全体
  const values = range.getValues();

  const converted = values.map(row => {
    let v = row[0];
    if (typeof v === "string") {
      v = v.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
    }
    return [v];
  });

  range.setValues(converted);
}
