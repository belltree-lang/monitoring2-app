/***** ── 設定 ─────────────────────────────────*****/
const SPREADSHEET_ID = '1wdHF0txuZtrkMrC128fwUSImyt320JhBVqXloS7FgpU'; // ←ご指定
const SHEET_NAME      = 'Monitoring'; // ケアマネ用モニタリング
const OPENAI_MODEL    = 'gpt-4o-mini';

// 画像/動画/PDF の既定保存先（利用者IDごとにサブフォルダを自動作成）
const DEFAULT_FOLDER_ID         = '1glDniVONBBD8hIvRGMPPT1iLXdtHJpEC';
const MEDIA_ROOT_FOLDER_ID      = DEFAULT_FOLDER_ID;
const REPORT_FOLDER_ID_PROP     = DEFAULT_FOLDER_ID;
const ATTACHMENTS_FOLDER_ID_PROP= DEFAULT_FOLDER_ID;

// Docsテンプレ（任意）：プロパティで上書き可（なければ自動レイアウト）
const DOC_TEMPLATE_ID_PROP        = PropertiesService.getScriptProperties().getProperty('DOC_TEMPLATE_ID') || '';
const DOC_TEMPLATE_ID_FAMILY_PROP = PropertiesService.getScriptProperties().getProperty('DOC_TEMPLATE_ID_FAMILY') || '';

/***** ── Webエントリ ───────────────────────────*****/
function doGet(e) {
  const tmpl = HtmlService.createTemplateFromFile('member'); // ファイル名: member.html
  return tmpl.evaluate()
    .setTitle('ケアマネ・モニタリング')
    .addMetaTag('viewport','width=device-width, initial-scale=1.0');
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

/***** ── バイナリアップロード受付（fetch(FormData) → doPost）※未使用でも残置 ──*****/
function doPost(e) {
  try {
    var action = (e.parameter && e.parameter.action) || '';
    if (action !== 'upload') {
      return ContentService.createTextOutput(JSON.stringify({ status:'error', message:'unknown action' }))
        .setMimeType(ContentService.MimeType.JSON)
        .setHeader('Access-Control-Allow-Origin','*');
    }
    var memberId = (e.parameter && e.parameter.memberId) || '';
    var name     = (e.parameter && e.parameter.name) || 'upload';
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

    try { ensureSharingForMember_(file, memberId); } catch(_e){}

    var out = { status:'success', fileId:fileId, url:url, name:file.getName(), mimeType:file.getMimeType() };
    return ContentService.createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin','*');

  } catch (err) {
    var outErr = { status:'error', message: String(err && err.message || err) };
    return ContentService.createTextOutput(JSON.stringify(outErr))
      .setMimeType(ContentService.MimeType.JSON)
      .setHeader('Access-Control-Allow-Origin','*');
  }
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

    where.push('done');
    return { status:'success', fileId, url, name:file.getName(), mimeType:file.getMimeType() };

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
  const colDate = header.indexOf('日付');
  const colId   = header.indexOf('利用者ID');
  const colKind = header.indexOf('種別');
  const colRec  = header.indexOf('記録内容');
  const colAtt  = header.indexOf('添付');

  if (colDate < 0 || colId < 0 || colKind < 0 || colRec < 0 || colAtt < 0) {
    throw new Error(`ヘッダー不一致（必要: 日付/利用者ID/種別/記録内容/添付, 実際: ${JSON.stringify(header)}）`);
  }

  const tz = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const toDateText = (v) => {
    const d = (v instanceof Date) ? v : new Date(v);
    if (d && d.getTime && !isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, 'yyyy/MM/dd HH:mm');
    }
    return String(v ?? '');
  };

  let limitDate = null;
  if (days && String(days) !== 'all') {
    const n = Number(days);
    if (!isNaN(n) && n > 0) limitDate = new Date(Date.now() - n * 24 * 3600 * 1000);
  }

  const out = [];
  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const id  = String(row[colId] || '').trim();
    if (id !== String(memberId).trim()) continue;

    const rawDate = row[colDate];
    const d = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (limitDate && d instanceof Date && !isNaN(d) && d < limitDate) continue;

    let attachments = [];
    try { attachments = JSON.parse(String(row[colAtt] || '[]')) || []; }
    catch(_e){ attachments = []; }

    out.push({
      rowIndex : i + 1,
      dateText : toDateText(rawDate),
      kind     : String(row[colKind] ?? ''),
      text     : String(row[colRec]  ?? ''),
      attachments
    });
  }

  out.sort((a,b) => new Date(b.dateText).getTime() - new Date(a.dateText).getTime());
  return out;
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
function updateRecord(rowIndex, newText){
  if (!rowIndex) throw new Error('rowIndex未指定');
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  sh.getRange(Number(rowIndex), 4).setValue(String(newText||'')); // 4列目=記録内容
  return { status:'success' };
}
function deleteRecord(rowIndex){
  if (!rowIndex) throw new Error('rowIndex未指定');
  const sh = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  sh.deleteRow(Number(rowIndex));
  return { status:'success' };
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('ほのぼのID');
  const vals = sh.getDataRange().getValues();
  const out = [];

  for (let i=1; i<vals.length; i++) {
    let id   = String(vals[i][0] || '').trim();
    let name = String(vals[i][1] || '').trim();

    // ✅ 全角数字 → 半角数字に変換
    id = id.replace(/[０-９]/g, function(s) {
      return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });

    // 4桁ゼロ埋め
    id = ('0000' + id).slice(-4);

    if (id) out.push({ id, name });
  }
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
function helloWorld() {
  Logger.log("Hello from VS Code!");
}
