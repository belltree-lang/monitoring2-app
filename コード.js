/***** ── 設定 ─────────────────────────────────*****/
const SPREADSHEET_ID = '1wdHF0txuZtrkMrC128fwUSImyt320JhBVqXloS7FgpU'; // ←ご指定
const SHEET_NAME      = 'Monitoring'; // ケアマネ用モニタリング
const OPENAI_MODEL    = 'gpt-4o-mini';
const SHARE_SHEET_NAME = 'ExternalShares';
const SHARE_LOG_SHEET_NAME = 'ExternalShareAccessLog';
const SHARE_QR_SIZE = '220x220';

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

/** Webアプリのエントリーポイント */
function doGet(e) {
  try {
    Logger.log("🟢 doGet called at " + new Date());
    Logger.log("raw event = " + JSON.stringify(e));

    const params = e && e.parameter ? e.parameter : {};
    Logger.log("params = " + JSON.stringify(params));

    // --- JSON API モード（フロントから fetch された時用） ---
    if (params.api === 'shareMeta') {
      const tokenParam = cleanParam_(params.shareId || params.share || params.token || "");
      const recordId   = cleanParam_(params.recordId || params.record || "");
      if (!tokenParam) {
        return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'shareId is missing' }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      const shareResult = getExternalShareMeta(tokenParam, recordId);
      return ContentService.createTextOutput(JSON.stringify(shareResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // --- 通常のHTML表示モード ---
    const shareToken = cleanParam_(params.shareId || params.share || params.token || "");
    const recordId   = cleanParam_(params.recordId || params.record || "");
    const printParamRaw = params.print || params.mode;
    const wantsPrint = shareToken && String(printParamRaw || "").trim() !== "" && String(printParamRaw || "").trim() !== "0";

    Logger.log("shareToken=" + shareToken + ", recordId=" + recordId + ", wantsPrint=" + wantsPrint);

    const templateName = wantsPrint ? "print" : (shareToken ? "share" : "member");
    Logger.log("templateName = " + templateName);

    const tmpl = HtmlService.createTemplateFromFile(templateName);

    if (shareToken) {
      tmpl.shareToken = shareToken;
      if (recordId) tmpl.shareRecordId = recordId;
    }

    let title = shareToken ? "モニタリング共有ビュー" : "ケアマネ・モニタリング";

    if (wantsPrint && shareToken) {
      const meta = getExternalShareMeta(shareToken, recordId);
      tmpl.shareMeta = meta;
      let printMode = "record";
      let printRecords = [];
      let primaryRecord = null;
      let centerLabel = "";
      let staffLabel = "";
      let errorMessage = "";
      const requestedMode = String(params.mode || "").trim().toLowerCase();

      if (meta && meta.status === "success" && meta.share) {
        const share = meta.share;
        const initialRecords = Array.isArray(meta.records) ? meta.records.slice() : [];
        primaryRecord = meta.primaryRecord || (initialRecords.length ? initialRecords[0] : null);
        printRecords = initialRecords;

        if (requestedMode === "center" && primaryRecord && primaryRecord.center) {
          const centerRecords = getRecordsByCenter(primaryRecord.center);
          const payload = buildExternalSharePayload_(share, { records: centerRecords, center: primaryRecord.center, recordId });
          printRecords = payload.records;
          primaryRecord = payload.primaryRecord || primaryRecord;
          centerLabel = primaryRecord.center || (primaryRecord.fields && primaryRecord.fields.center) || "";
          printMode = "center";
        } else if (requestedMode === "staff" && primaryRecord && primaryRecord.staff) {
          const staffRecords = getRecordsByStaff(primaryRecord.staff);
          const payload = buildExternalSharePayload_(share, { records: staffRecords, staff: primaryRecord.staff, recordId });
          printRecords = payload.records;
          primaryRecord = payload.primaryRecord || primaryRecord;
          staffLabel = primaryRecord.staff || (primaryRecord.fields && primaryRecord.fields.staff) || "";
          printMode = "staff";
        } else {
          const payload = buildExternalSharePayload_(share, { recordId });
          printRecords = payload.records;
          primaryRecord = payload.primaryRecord || primaryRecord;
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
      tmpl.printRecordId = recordId;

      const tz = Session.getScriptTimeZone ? (Session.getScriptTimeZone() || "Asia/Tokyo") : "Asia/Tokyo";
      tmpl.printedAtText = Utilities.formatDate(new Date(), tz, "yyyy/MM/dd HH:mm");
      title = "モニタリング記録 印刷";
    }

    Logger.log("✅ doGet finished, returning template: " + templateName);

    return tmpl.evaluate()
      .setTitle(title)
      .addMetaTag("viewport", "width=device-width, initial-scale=1.0");

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
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
  const info = { header: [], headerNormalized: [], width: 0, idCol: -1, nameCol: -1, yomiCol: -1, careCol: -1, centerCol: -1 };
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

  const yomiCol = findMemberSheetColumnIndex_(headerNormalized, yomiCandidates);
  const careCol = findMemberSheetColumnIndex_(headerNormalized, careCandidates);
  const centerCol = findMemberSheetColumnIndex_(headerNormalized, centerCandidates);

  info.idCol = idCol;
  info.nameCol = nameCol;
  info.yomiCol = yomiCol;
  info.careCol = careCol;
  info.centerCol = centerCol;
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

function buildExternalShareQrDataUrl_(url) {
  try {
    if (!url) return '';
    // Chart API の URL をそのまま返す
    const encoded = encodeURIComponent(url);
    return `https://chart.googleapis.com/chart?cht=qr&chs=220x220&choe=UTF-8&chl=${encoded}`;
  } catch (e) {
    Logger.log("buildExternalShareQrDataUrl_ error: " + e);
    return '';
  }
}


/** 呼び出し側の後方互換 */
function buildExternalShareQrUrl_(shareUrl, size){
  return buildExternalShareQrDataUrl_(shareUrl, size);
}
/** tokenを確実に保存し、列順固定でappend */
function createExternalShare(memberId, options){
  Logger.log("🟢 createExternalShare called with", memberId, JSON.stringify(options));
  Logger.log("🔎 createExternalShare called");
  Logger.log("memberId = " + memberId);
  Logger.log("options = " + JSON.stringify(options)); // ここは options を出す
  const baseUrl = getExecUrlSafe_();
  Logger.log("🟩 createExternalShare: serviceUrl=%s memberId=%s", baseUrl, memberId);

  try {
    const normalizedId = normalizeMemberId_(memberId);
    const rawId = String(memberId || '').trim();
    const resolvedId = normalizedId || rawId;
    if (!resolvedId) throw new Error('利用者IDが未指定です');

    const shareSheet = ensureShareSheet_();
    const config = options && typeof options === 'object' ? options : {};
    Logger.log("config = " + JSON.stringify(config)); // ← 定義後にログ出す

    const audienceRaw = String(config.audience || '').trim().toLowerCase();
    const audience = ['family','center','medical','service'].includes(audienceRaw)
      ? audienceRaw
      : 'family';

    const maskMode = (config.maskMode === 'none') ? 'none' : 'simple';
    const passwordHash = hashSharePassword_(config.password);

    const allowedRaw = Array.isArray(config.allowedAttachmentIds) ? config.allowedAttachmentIds : [];
    const allowAll = allowedRaw.includes('__ALL__');
    const allowedAttachmentIds = allowAll
      ? ['__ALL__']
      : Array.from(new Set(allowedRaw.filter(v => v && v !== '__ALL__').map(String)));

    const rangeSpec = normalizeShareRangeSpec_(config.range || config.rangeSpec || config.recordRange);
    const rangeLabel = formatShareRangeLabel_(rangeSpec);

    // 期限
    let expiresAtIso = '';
    if (config.expiresAt) {
      const expires = new Date(config.expiresAt);
      if (!isNaN(expires.getTime())) expiresAtIso = expires.toISOString();
    } else if (config.expiresInDays) {
      const days = Number(config.expiresInDays);
      if (!isNaN(days) && days > 0) {
        const expires = new Date(Date.now() + days * 24 * 3600 * 1000);
        expiresAtIso = expires.toISOString();
      }
    }

    const token = Utilities.getUuid().replace(/-/g, '');
    const nowIso = new Date().toISOString();

    // 列順に合わせて書き込む
    shareSheet.appendRow([
      token,                         // Token
      resolvedId,                    // MemberID
      passwordHash,                  // PasswordHash
      expiresAtIso,                  // ExpiresAt
      maskMode,                      // MaskMode
      JSON.stringify(allowedAttachmentIds), // AllowedAttachments
      nowIso,                        // CreatedAt
      '',                            // RevokedAt
      '',                            // LastAccessAt
      audience,                      // Audience
      0,                             // AccessCount
      rangeSpec                      // RangeSpec
    ]);

    const url = buildExternalShareUrl_(token);
    const qrDataUrl = buildExternalShareQrDataUrl_(url);

    return {
      status:'success',
      token,
      shareId: token,
      memberId: resolvedId,
      url,
      shareLink: url,
      qrUrl: qrDataUrl,
      qrDataUrl,
      qrCode: qrDataUrl,
      audience,
      expiresAt: expiresAtIso,
      maskMode,
      allowAllAttachments: allowAll,
      range: rangeSpec,
      rangeLabel
    };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}

function getExternalShares(memberId){
  Logger.log("🟢 getExternalShares called with memberId=" + memberId);
  try {
    const id = String(memberId || '').trim();
    if (!id) throw new Error('利用者IDが未指定です');

    const sheet = ensureShareSheet_();
    const values = sheet.getDataRange().getValues();
    if (!values || values.length <= 1) return { status:'success', shares: [] };

    const now = Date.now();
    const shares = [];

    for (let i = 1; i < values.length; i++) {
      const share = parseShareRow_(values[i]);
      if (!share.token || share.memberId !== id) continue;
      if (share.revokedAt) continue;

      const allowAll = share.allowedAttachmentIds.includes('__ALL__');
      const allowedCount = allowAll ? 0 : share.allowedAttachmentIds.filter(v => v && v !== '__ALL__').length;
      const expired = !!(share.expiresAt && share.expiresAt.getTime() < now);
      const url = buildExternalShareUrl_(share.token);
      const qrDataUrl = buildExternalShareQrDataUrl_(url);

      shares.push({
        token: share.token,
        url,
        shareLink: url,
        qrUrl: qrDataUrl,
        qrDataUrl,
        qrCode: qrDataUrl,
        createdAtText: formatShareDate_(share.createdAt),
        createdAtMs: share.createdAt ? share.createdAt.getTime() : 0,
        expiresAtText: formatShareDate_(share.expiresAt),
        expired,
        audience: share.audience,
        passwordProtected: !!share.passwordHash,
        maskMode: share.maskMode || 'simple',
        allowAllAttachments: allowAll,
        allowedCount,
        lastAccessText: formatShareDate_(share.lastAccessAt),
        remainingLabel: computeRemainingLabel_(share.expiresAt),
        accessCount: share.accessCount || 0,
        rangeLabel: share.rangeLabel
      });
    }

    shares.sort((a, b) => (b.createdAtMs || 0) - (a.createdAtMs || 0));
    shares.forEach(s => { if ('createdAtMs' in s) delete s.createdAtMs; });

    return { status:'success', shares };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}

function revokeExternalShare(token){
  try {
    const info = findShareRowByToken_(token);
    if (!info) throw new Error('対象の共有リンクが見つかりません');
    const { sheet, rowIndex } = info;
    sheet.getRange(rowIndex, 8).setValue(new Date()); // RevokedAt
    return { status:'success' };
  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}

function getExternalShareMeta(token, recordId) {
  Logger.log("🟦 getExternalShareMeta: token=%s recordId=%s", token, recordId);
  try {
    const info = findShareRowByToken_(token);
    if (!info) throw new Error('無効な共有リンクです');
    const { sheet, rowIndex, share } = info;
    if (share.revokedAt) throw new Error('共有リンクは停止されています');

    const now = Date.now();
    const expired = !!(share.expiresAt && share.expiresAt.getTime() < now);
    const allowAll = share.allowedAttachmentIds.includes('__ALL__');
    const allowedCount = allowAll ? 0 : share.allowedAttachmentIds.filter(v => v && v !== '__ALL__').length;
    const url = buildExternalShareUrl_(share.token);
    const qrDataUrl = buildExternalShareQrDataUrl_(url);
    const audienceInfo = getShareAudienceInfo_(share.audience);
    const profile = lookupMemberProfile_(share.memberId);
    if (!profile.found) throw new Error('利用者情報が見つかりません');

    const summary = {
      token: share.token,
      memberId: profile.id || share.memberId,
      memberName: profile.name || lookupMemberName_(share.memberId),
      memberCenter: profile.center || '',
      memberStaff: profile.staff || '',
      expiresAtText: formatShareDate_(share.expiresAt),
      expired,
      audience: share.audience,
      requirePassword: !!share.passwordHash,
      maskMode: share.maskMode || 'simple',
      allowAllAttachments: allowAll,
      allowedCount,
      remainingLabel: computeRemainingLabel_(share.expiresAt),
      rangeLabel: share.rangeLabel,
      url,
      shareLink: url,
      qrUrl: qrDataUrl,
      qrDataUrl,
      qrCode: qrDataUrl,
      audienceInfo
    };

    const recordIdSafe = String(recordId || '').trim();

    // パスワード不要なら即payload同梱
    let payload = { records: [], primaryRecord: null };
    if (!summary.requirePassword) {
      try {
        const days = shareRangeToFetchArg_(share.rangeSpec);   // 30/90/all を解釈
        const records = fetchRecordsWithIndex_(share.memberId, days);
        payload = { records, primaryRecord: records[0] || null };
      } catch (pErr) {
        Logger.log('getExternalShareMeta payload error: ' + pErr);
        payload = { records: [], primaryRecord: null };
      }

      // 🔎 デバッグ用ログ
      Logger.log('📊 payload.records count = ' + (payload.records ? payload.records.length : 0));
      if (payload.records && payload.records.length) {
        try {
          Logger.log('📊 first record sample = ' + JSON.stringify(payload.records[0]));
        } catch (_e) {}
      }

      // アクセスログ更新
      try {
        sheet.getRange(rowIndex, 9).setValue(new Date()); // LastAccessAt
        const nextCount = (share.accessCount || 0) + 1;
        sheet.getRange(rowIndex, 11).setValue(nextCount);
        logExternalShareAccess_(share);
      } catch (logErr) {
        Logger.log('getExternalShareMeta log error: ' + logErr);
      }
    }

    // 最終レスポンス
    return {
      status: 'success',
      share: summary,
      records: payload.records,
      primaryRecord: payload.primaryRecord
    };

  } catch (e) {
    return { status: 'error', message: String(e && e.message || e) };
  }
}

function enterExternalShare(token, password, recordId){
  try {
    const info = findShareRowByToken_(token);
    if (!info) throw new Error('無効な共有リンクです');
    const { sheet, rowIndex, share } = info;
    if (share.revokedAt) throw new Error('共有リンクは停止されています');

    // 期限判定
    const now = Date.now();
    if (share.expiresAt && share.expiresAt.getTime() < now) {
      return { status:'error', message:'この共有リンクは期限切れです。' };
    }

    // パスワード判定
    if (share.passwordHash) {
      const hash = hashSharePassword_(password);
      if (!hash || hash !== share.passwordHash) {
        return { status:'error', message:'パスワードが一致しません。' };
      }
    }

    const recordIdSafe = String(recordId || '').trim();
    let payload = { records: [], primaryRecord: null };
    try {
      payload = buildExternalSharePayload_(share, { recordId: recordIdSafe }) || { records: [], primaryRecord: null };
    } catch (pErr) {
      Logger.log('enterExternalShare payload error: ' + pErr);
    }

    // 🔎 デバッグ出力
    Logger.log('📊 enter payload.records count = ' + (payload.records ? payload.records.length : 0));
    if (payload.records && payload.records.length && !payload.primaryRecord) {
      payload.primaryRecord = payload.records[0];
    }

    // recordId が指定されているのに 0 件ならエラー
    if (recordIdSafe && recordIdSafe.length > 0 && (!payload.records || !payload.records.length)) {
      return { status:'error', message:'対象の記録が見つかりません。' };
    }

    // アクセスログ
    sheet.getRange(rowIndex, 9).setValue(new Date()); // LastAccessAt
    const nextCount = (share.accessCount || 0) + 1;
    sheet.getRange(rowIndex, 11).setValue(nextCount);
    logExternalShareAccess_(share);

    // サマリー（パス後の確定版）
    const allowAll = share.allowedAttachmentIds.includes('__ALL__');
    const allowedCount = allowAll ? 0 : share.allowedAttachmentIds.filter(v => v && v !== '__ALL__').length;
    const url = buildExternalShareUrl_(share.token);
    const qrDataUrl = buildExternalShareQrDataUrl_(url);
    const audienceInfo = getShareAudienceInfo_(share.audience);
    const profile = lookupMemberProfile_(share.memberId);
    if (!profile.found) throw new Error('利用者情報が見つかりません');

    const summary = {
      token: share.token,
      memberId: profile.id || share.memberId,
      memberName: profile.name || lookupMemberName_(share.memberId),
      memberCenter: profile.center || '',
      memberStaff: profile.staff || '',
      expiresAtText: formatShareDate_(share.expiresAt),
      expired: false,
      audience: share.audience,
      requirePassword: !!share.passwordHash,
      maskMode: share.maskMode || 'simple',
      allowAllAttachments: allowAll,
      allowedCount,
      remainingLabel: computeRemainingLabel_(share.expiresAt),
      rangeLabel: share.rangeLabel,
      url,
      shareLink: url,
      qrUrl: qrDataUrl,
      qrDataUrl,
      qrCode: qrDataUrl,
      audienceInfo
    };
    summary.hasRecords = !!(payload.records && payload.records.length);

    const response = {
      status:'success',
      share: summary,
      records: payload.records,
      primaryRecord: payload.primaryRecord
    };
    if (!summary.hasRecords) {
      response.message = '記録が存在しません';
    }
    return response;

  } catch (e) {
    return { status:'error', message:String(e && e.message || e) };
  }
}


/** 外部共有シートを必ずヘッダー揃えて用意する */
function ensureShareSheet_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("📄 ensureShareSheet_: spreadsheetId=%s", ss.getId());

  let sheet = ss.getSheetByName(SHARE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHARE_SHEET_NAME);
    Logger.log("📄 ensureShareSheet_: sheet created = %s", sheet.getName());
  } else {
    Logger.log("📄 ensureShareSheet_: sheet exists = %s", sheet.getName());
  }

  const header = [
    'Token',           // A1
    'MemberID',        // B1
    'PasswordHash',    // C1
    'ExpiresAt',       // D1 (ISO or Date)
    'MaskMode',        // E1 ('simple' | 'none')
    'AllowedAttachments', // F1 (JSON array or ['__ALL__'])
    'CreatedAt',       // G1 (Date)
    'RevokedAt',       // H1 (Date or blank)
    'LastAccessAt',    // I1 (Date)
    'Audience',        // J1 ('family'|'center'|'medical'|'service')
    'AccessCount',     // K1 (number)
    'RangeSpec'        // L1 ('all'|'90'|'30'|number)
  ];

  // 列数保証
  if (sheet.getMaxColumns() < header.length) {
    sheet.insertColumnsAfter(
      sheet.getMaxColumns(),
      header.length - sheet.getMaxColumns()
    );
  }

  // 1行目にヘッダーを常にセット
  sheet.getRange(1, 1, 1, header.length).setValues([header]);

  return sheet;
}

/** ExternalSharesの1行を安全にパース */
function parseShareRow_(row){
  const safeJson = (value) => {
    try { return JSON.parse(value); } catch(_e){ return []; }
  };
  const toDate = (value) => {
    if (!value) return null;
    const d = value instanceof Date ? value : new Date(value);
    return (d && !isNaN(d.getTime())) ? d : null;
  };
  const toNumber = (value) => {
    const num = Number(value);
    return isNaN(num) ? 0 : num;
  };
  const audienceRaw = String(row[9] || '').trim().toLowerCase();
  const audience = ['family','center','medical','service'].includes(audienceRaw) ? audienceRaw : 'family';
  const rangeSpec = normalizeShareRangeSpec_(row[11]);

  return {
    token: String(row[0] || '').trim(),
    memberId: (() => {
      const normalized = normalizeMemberId_(row[1]);
      return normalized || String(row[1] || '').trim();
    })(),
    passwordHash: String(row[2] || '').trim(),
    expiresAt: toDate(row[3]),
    maskMode: String(row[4] || 'simple').trim() || 'simple',
    allowedAttachmentIds: Array.isArray(row[5]) ? row[5] : safeJson(String(row[5] || '[]')),
    createdAt: toDate(row[6]),
    revokedAt: toDate(row[7]),
    lastAccessAt: toDate(row[8]),
    audience,
    accessCount: toNumber(row[10]),
    rangeSpec,
    rangeLabel: formatShareRangeLabel_(rangeSpec)
  };
}

function normalizeToken_(tok){
  return String(tok || "")
    .trim()
    .replace(/\s+/g, "")       // 改行やスペースを全部削除
    .toLowerCase();            // 念のため小文字化
}

function findShareRowByToken_(token){
  const tok = cleanParam_(token);   // ← ここでクリーンアップ
  if (!tok) return null;

  const sheet = ensureShareSheet_();
  const values = sheet.getDataRange().getValues();
  Logger.log("findShareRowByToken_: total rows=" + values.length);

  for (let i = 1; i < values.length; i++) {
    const share = parseShareRow_(values[i]);
    Logger.log(`🔍 compare row ${i+1}: sheet="${share.token}" vs request="${tok}"`);
    if (share.token === tok) {
      Logger.log(`✅ MATCH at row ${i+1}`);
      return { sheet, rowIndex: i + 1, share };
    }
  }
  Logger.log(`❌ no match after ${values.length-1} rows`);
  return null;
}



function buildExternalSharePayload_(share, options){
  const opts = options || {};
  const rangeArg = shareRangeToFetchArg_(share && (share.rangeSpec || share.rangeLabel || share.range));
  const allowAll = share.allowedAttachmentIds.includes('__ALL__');
  const allowedSet = new Set(allowAll ? [] : share.allowedAttachmentIds.filter(v => v && v !== '__ALL__'));
  const audience = share.audience || 'family';

  // レコード取得
  const recordsSource = Array.isArray(opts.records) && opts.records.length
    ? opts.records.slice()
    : fetchRecordsWithIndex_(share.memberId, rangeArg);

  // デバッグ用ログ
  if (recordsSource && recordsSource.length) {
    try {
      Logger.log("🔎 recordsSource sample = " + JSON.stringify(recordsSource[0]));
    } catch(_e){}
  }

  const recordIdFilter = String(opts.recordId || '').trim();
  const centerFilter = String(opts.center || '').trim();
  const staffFilter = String(opts.staff || '').trim();

  let filtered = recordsSource;

  // center 指定フィルタ
  if (centerFilter) {
    filtered = filtered.filter(rec => String(rec.center || '').trim().toLowerCase() === centerFilter.toLowerCase());
  }
  // staff 指定フィルタ
  if (staffFilter) {
    filtered = filtered.filter(rec => String(rec.staff || '').trim().toLowerCase() === staffFilter.toLowerCase());
  }
  // recordId 指定フィルタ（空や "0" "undefined" は無視）
  if (recordIdFilter && recordIdFilter !== "0" && recordIdFilter.toLowerCase() !== "undefined") {
    const matched = filtered.filter(rec => String(rec.recordId || rec.rowIndex || '').trim() === recordIdFilter);
    filtered = matched.length ? matched : [];
  }

  const results = [];
  let primaryRecord = null;

  filtered.forEach(rec => {
    const attachments = filterAttachmentsForShare_(rec.attachments, { allowAll, allowedSet });
    const maskedText = maskTextForExternal_(rec.text || '', share.maskMode);
    const timestamp = (typeof rec.timestamp === 'number') ? rec.timestamp : null;
    const fields = rec.fields ? Object.assign({}, rec.fields) : {};

    // マスク処理を fields に反映
    if ('記録内容' in fields) {
      fields['記録内容'] = maskedText;
    }
    if ('text' in fields && fields.text === rec.text) {
      fields.text = maskedText;
    }
    if ('添付' in fields) {
      fields['添付'] = attachments
        .map(att => att && att.name ? att.name : (att && att.url ? att.url : ''))
        .filter(Boolean)
        .join('\n');
    }
    if (!('center' in fields) && rec.center) {
      fields.center = rec.center;
    }
    if (!('staff' in fields) && rec.staff) {
      fields.staff = rec.staff;
    }

    const item = {
      recordId: rec.recordId || String(rec.rowIndex || ''),
      rowIndex: rec.rowIndex,
      memberId: rec.memberId || '',
      memberName: rec.memberName || lookupMemberName_(rec.memberId),
      dateText: rec.dateText || '',
      kind: rec.kind || '',
      audience,
      text: maskedText,
      attachments,
      timestamp,
      center: rec.center || '',
      staff: rec.staff || '',
      status: rec.status || fields.status || '',
      special: rec.special || fields.special || '',
      fields
    };

    results.push(item);

    const isPrimary = recordIdFilter
      ? String(item.recordId || '').trim() === recordIdFilter
      : !primaryRecord;
    if (isPrimary) {
      primaryRecord = item;
    }
  });

  // text か 添付が無いレコードは削除。ただし recordId 指定がある場合は例外
  const filteredResults = results.filter(rec => {
    if (recordIdFilter && String(rec.recordId || '').trim() === recordIdFilter) {
      return true;
    }
    return rec.text || (rec.attachments && rec.attachments.length);
  });

  if (!primaryRecord && filteredResults.length) {
    primaryRecord = filteredResults[0];
  }

  return { records: filteredResults, primaryRecord };
}



/** アクセスログ（追加シートに追記） */
function logExternalShareAccess_(share){
  try {
    const sheet = ensureShareLogSheet_();
    sheet.appendRow([ new Date(), share.token, share.memberId, share.audience || 'family' ]);
  } catch (e) {
    Logger.log('logExternalShareAccess_ error: ' + e);
  }
}

function ensureShareLogSheet_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHARE_LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHARE_LOG_SHEET_NAME);
  const header = ['AccessedAt','Token','MemberID','Audience'];
  if (sheet.getMaxColumns() < header.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), header.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, header.length).setValues([header]);
  return sheet;
}

function normalizeShareRangeSpec_(value){
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return '30';
  if (raw === 'all' || raw === 'full' || raw === 'unlimited') return 'all';
  if (raw === '90' || raw === '90d' || raw === '90days') return '90';
  if (raw === '30' || raw === '30d' || raw === '30days') return '30';
  const num = parseInt(raw, 10);
  if (!isNaN(num)) {
    if (num >= 90) return '90';
    if (num >= 30) return '30';
    if (num <= 0) return '30';
    return String(num);
  }
  return '30';
}

function formatShareRangeLabel_(spec){
  const normalized = normalizeShareRangeSpec_(spec);
  if (normalized === 'all') return '全期間';
  if (normalized === '90') return '直近90日';
  return '直近30日';
}

function shareRangeToFetchArg_(spec){
  const normalized = normalizeShareRangeSpec_(spec);
  if (normalized === 'all') return 'all';
  const days = Number(normalized);
  return (!isNaN(days) && days > 0) ? days : 'all';
}

function filterAttachmentsForShare_(attachments, option){
  const arr = Array.isArray(attachments) ? attachments : [];
  if (option.allowAll) {
    return arr.map(normalizeAttachmentForShare_).filter(Boolean);
  }
  if (!option.allowedSet || !option.allowedSet.size) return [];
  return arr.map(normalizeAttachmentForShare_).filter(att => att && att.fileId && option.allowedSet.has(att.fileId));
}

function normalizeAttachmentForShare_(att){
  if (!att || typeof att !== 'object') return null;
  const fileId = String(att.fileId || att.id || '').trim();
  const url = String(att.url || (fileId ? `https://drive.google.com/file/d/${fileId}/view` : '')).trim();
  const name = String(att.name || att.fileName || att.title || '添付ファイル');
  if (!fileId && !url) return null;
  return { fileId, url, name, mimeType: String(att.mimeType || '') };
}

function hashSharePassword_(password){
  const value = String(password || '');
  if (!value) return '';
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function maskTextForExternal_(text, mode){
  if (mode === 'none') return String(text || '');
  const value = String(text || '');
  return value
    .replace(/[0-9０-９]/g, '＊')
    .replace(/([A-Za-z\u3040-\u30FF\u4E00-\u9FFF]{2,})/g, (m) => m.charAt(0) + '＊'.repeat(m.length - 1));
}

function formatShareDate_(date){
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  return Utilities.formatDate(date, tz, 'yyyy/MM/dd HH:mm');
}

function computeRemainingLabel_(expiresAt){
  if (!(expiresAt instanceof Date) || isNaN(expiresAt.getTime())) return '';
  const diff = expiresAt.getTime() - Date.now();
  if (diff <= 0) return '';
  const hours = Math.floor(diff / (3600 * 1000));
  if (hours >= 48) {
    const days = Math.floor(hours / 24);
    return `残り約${days}日`;
  }
  if (hours >= 1) {
    return `残り約${hours}時間`;
  }
  const minutes = Math.floor(diff / (60 * 1000));
  return minutes > 0 ? `残り約${minutes}分` : '';
}

function lookupMemberProfile_(memberId){
  const empty = { id: String(memberId || ''), name: '', center: '', staff: '', found: false };
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName('ほのぼのID');
    if (!sh) return empty;
    const values = sh.getDataRange().getValues();
    if (!values || values.length <= 1) return empty;
    const layout = getMemberSheetColumnInfo_(values);
    const targetId = normalizeMemberId_(memberId);
    if (!targetId) return empty;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rawId = (layout.idCol >= 0 && layout.idCol < row.length) ? row[layout.idCol] : '';
      const normalized = normalizeMemberId_(rawId);
      if (!normalized || normalized !== targetId) continue;
      const nameRaw = (layout.nameCol >= 0 && layout.nameCol < row.length) ? row[layout.nameCol] : '';
      const centerRaw = (layout.centerCol >= 0 && layout.centerCol < row.length) ? row[layout.centerCol] : '';
      const staffRaw = (layout.careCol >= 0 && layout.careCol < row.length) ? row[layout.careCol] : '';
      return {
        id: targetId,
        name: String(nameRaw || '').trim(),
        center: String(centerRaw || '').trim(),
        staff: String(staffRaw || '').trim(),
        found: true
      };
    }
  } catch (_e) {}
  return empty;
}

function lookupMemberName_(memberId){
  const profile = lookupMemberProfile_(memberId);
  return profile && profile.name ? profile.name : '';
}

function getShareAudienceInfo_(audience){
  const map = {
    family: {
      label: 'ご家族向け共有',
      description: 'ご家族の皆さまが状況を把握しやすいよう、本文を簡潔にまとめています。',
      intro: 'ご家族とのコミュニケーションにご活用ください。',
      manualTips: [
        'QRコードからアクセスし、スマートフォンやパソコンで最新の記録をご覧いただけます。',
        '閲覧後のご感想や気づきがあれば、担当ケアマネジャーまでお知らせください。'
      ]
    },
    center: {
      label: '地域包括支援センター向け共有',
      description: '日付や種別を含めて記録を確認しやすいレイアウトです。',
      intro: '地域包括支援センターの職員さまとの情報共有にご利用ください。',
      manualTips: [
        'QRコードからアクセスし、閲覧専用のページで記録をご確認ください。',
        '気づいた点があればケアマネジャーへフィードバックをお願いします。'
      ]
    },
    medical: {
      label: '医療連携向け共有',
      description: '医師・看護師が経過を把握しやすいよう、必要事項を抜粋しています。',
      intro: '診察や訪問時の参考情報としてご活用ください。',
      manualTips: [
        'QRコードを読み取り、モニタリング記録を時系列で確認できます。',
        '必要に応じて担当ケアマネジャーへご連絡ください。'
      ]
    },
    service: {
      label: 'サービス事業者向け共有',
      description: 'ケア実務者が把握しやすいよう、現場目線で構成しています。',
      intro: 'サービス提供に関する情報共有にご利用ください。',
      manualTips: [
        'QRコードでアクセスし、必要な記録をいつでも確認できます。',
        'サービス提供に関する気づきはケアマネジャーまでご連絡ください。'
      ]
    }
  };
  const key = String(audience || 'family').toLowerCase();
  return map[key] || map.family;
}

/** 既存UI互換：本文だけ差し替える簡易更新 */
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
  // 少なくともE列まで用意
  if (sh.getMaxColumns() < 5) sh.insertColumnsAfter(sh.getMaxColumns(), 5 - sh.getMaxColumns());
  // ヘッダを確定（A=ID, B=氏名 は触らない）
  sh.getRange(1, 4).setValue('包括支援センター'); // D1
  sh.getRange(1, 5).setValue('担当者名');         // E1
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

