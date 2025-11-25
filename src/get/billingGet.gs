/**
 * Get layer: data retrieval utilities for billing pipeline.
 */

const BILLING_SPREADSHEET_ID = (() => {
  const props = PropertiesService.getScriptProperties();
  const propValue = props && props.getProperty('BILLING_SPREADSHEET_ID');
  if (propValue) return propValue.trim();
  return (typeof SPREADSHEET_ID !== 'undefined' && SPREADSHEET_ID) ? SPREADSHEET_ID : '';
})();

const TREATMENT_SHEET_PREFIX = '施術録_';
const PATIENT_INFO_SHEET_NAME = '患者情報';
const PAYMENT_RESULT_PREFIX = '入金結果_';
const PATIENT_INFO_COLUMNS_WIDTH = 62; // A~BJ

/**
 * Normalize YYYYMM text and derive year/month.
 * @param {string|number} billingMonth
 * @returns {{text: string, year: number, monthIndex: number}}
 */
function normalizeBillingMonth_(billingMonth) {
  const text = String(billingMonth || '').trim();
  if (!/^\d{6}$/.test(text)) {
    throw new Error('billingMonth は YYYYMM 形式で指定してください');
  }
  const year = Number(text.slice(0, 4));
  const monthIndex = Number(text.slice(4, 6)) - 1;
  if (!(year > 1900) || monthIndex < 0 || monthIndex > 11) {
    throw new Error('billingMonth の値が不正です');
  }
  return { text, year, monthIndex };
}

/**
 * Build sheet name like "施術録_YYYYMM".
 */
function buildTreatmentSheetName_(billingMonth) {
  const normalized = normalizeBillingMonth_(billingMonth);
  return `${TREATMENT_SHEET_PREFIX}${normalized.text}`;
}

function getBillingSpreadsheet_() {
  if (!BILLING_SPREADSHEET_ID) {
    throw new Error('BILLING_SPREADSHEET_ID が未設定です');
  }
  return SpreadsheetApp.openById(BILLING_SPREADSHEET_ID);
}

function normalizeHeaderLabel_(label) {
  return String(label || '')
    .replace(/\s+/g, '')
    .replace(/[()（）]/g, '')
    .toLowerCase();
}

function findColumnIndex_(headerRow, candidates, fallbackIndex) {
  const normalizedHeader = headerRow.map(normalizeHeaderLabel_);
  for (const candidate of candidates) {
    const target = normalizeHeaderLabel_(candidate);
    const idx = normalizedHeader.indexOf(target);
    if (idx >= 0) return idx;
  }
  return typeof fallbackIndex === 'number' ? fallbackIndex : -1;
}

/**
 * 1-1. 施術録（月別）取得。
 * 対象月シートの患者ID別の施術回数を返す。
 * @param {string|number} billingMonth YYYYMM
 * @returns {Object<string, {visitCount: number}>}
 */
function getTreatmentVisitCounts(billingMonth) {
  const { text, year, monthIndex } = normalizeBillingMonth_(billingMonth);
  const sheetName = buildTreatmentSheetName_(text);
  const ss = getBillingSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`施術録シートが見つかりません: ${sheetName}`);
  }

  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return {};

  const header = values[0].map(v => String(v || '').trim());
  const dateCol = findColumnIndex_(header, ['日付', 'date'], 0);
  const patientIdCol = findColumnIndex_(header, ['患者id', '患者ID', 'patientid', 'id'], 1);
  const result = {};

  const rangeStart = new Date(year, monthIndex, 1).getTime();
  const rangeEnd = new Date(year, monthIndex + 1, 1).getTime();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const patientId = String(row[patientIdCol] || '').trim();
    if (!patientId) continue;

    const dateValue = row[dateCol];
    let ts = null;
    if (dateValue instanceof Date && !isNaN(dateValue.getTime())) {
      ts = dateValue.getTime();
    } else {
      const parsed = new Date(dateValue);
      if (parsed && !isNaN(parsed.getTime())) ts = parsed.getTime();
    }
    if (ts === null || ts < rangeStart || ts >= rangeEnd) {
      continue;
    }

    if (!result[patientId]) {
      result[patientId] = { visitCount: 0 };
    }
    result[patientId].visitCount += 1;
  }

  return result;
}

function columnLetter_(index) {
  let n = index + 1;
  let letters = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

function buildRawRecord_(header, row) {
  const raw = {};
  for (let i = 0; i < PATIENT_INFO_COLUMNS_WIDTH; i++) {
    const key = header[i] || columnLetter_(i);
    raw[key] = row[i];
  }
  return raw;
}

function pickNumeric_(value, defaultValue) {
  const num = Number(value);
  return isFinite(num) ? num : defaultValue;
}

function pickBooleanInt_(value, defaultValue) {
  if (value === 1 || value === '1' || value === true) return 1;
  if (value === 0 || value === '0' || value === false) return 0;
  return defaultValue;
}

function pickByHeaderCandidates_(header, row, candidates, fallbackIndex, defaultValue) {
  const idx = findColumnIndex_(header, candidates, fallbackIndex);
  if (idx < 0 || idx >= row.length) return defaultValue;
  return row[idx];
}

/**
 * 1-2. 患者情報（A〜BJ列の全情報）を読み込み。
 * @returns {Array<Object>}
 */
function getPatientInfoAll() {
  const ss = getBillingSpreadsheet_();
  const sheet = ss.getSheetByName(PATIENT_INFO_SHEET_NAME);
  if (!sheet) {
    throw new Error(`患者情報シートが見つかりません: ${PATIENT_INFO_SHEET_NAME}`);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(1, 1, lastRow, PATIENT_INFO_COLUMNS_WIDTH).getValues();
  const header = (values[0] || []).map(v => String(v || '').trim());

  const records = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const raw = buildRawRecord_(header, row);

    const patientId = String(pickByHeaderCandidates_(header, row, ['患者id', '患者ID', 'id', 'patientid'], 0, '') || '').trim();
    const nameKanji = String(pickByHeaderCandidates_(header, row, ['氏名', '名前', '漢字氏名', 'name'], 1, '') || '').trim();
    const nameKana = String(pickByHeaderCandidates_(header, row, ['カナ', 'フリガナ', 'かな', 'ふりがな', 'namekana'], 2, '') || '').trim();
    const insuranceType = String(pickByHeaderCandidates_(header, row, ['保険区分', '保険種別', 'insurance', '保険'], -1, '') || '').trim();
    const burdenRate = pickNumeric_(pickByHeaderCandidates_(header, row, ['負担割合', '負担率', 'burdenrate'], -1, ''), 0);
    const bankCode = String(pickByHeaderCandidates_(header, row, ['金融機関コード', '銀行コード', 'bankcode'], 13, '') || '').trim();
    const branchCode = String(pickByHeaderCandidates_(header, row, ['支店コード', 'branchcode'], 14, '') || '').trim();
    const accountNumber = String(pickByHeaderCandidates_(header, row, ['口座番号', 'accountnumber'], 16, '') || '').trim();
    const isNew = pickBooleanInt_(pickByHeaderCandidates_(header, row, ['新規', '新患', 'isnew'], 20, 0), 0);
    const carryOverAmount = pickNumeric_(pickByHeaderCandidates_(header, row, ['未入金額', '繰越金額', 'carryoveramount'], -1, 0), 0);

    records.push({
      patientId,
      raw,
      nameKanji,
      nameKana,
      insuranceType,
      burdenRate,
      bankCode,
      branchCode,
      accountNumber,
      isNew,
      carryOverAmount
    });
  }

  return records;
}

/**
 * 1-3. 入金結果（PDF → スタッフ手入力）取得。
 * @param {string|number} billingMonth YYYYMM
 * @returns {Object<string, {patientId: string, bankStatus: string}>}
 */
function getPaymentResults(billingMonth) {
  const { text } = normalizeBillingMonth_(billingMonth);
  const sheetName = `${PAYMENT_RESULT_PREFIX}${text}`;
  const ss = getBillingSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`入金結果シートが見つかりません: ${sheetName}`);
  }

  const values = sheet.getDataRange().getValues();
  if (!values || values.length <= 1) return {};

  const header = values[0].map(v => String(v || '').trim());
  const patientIdCol = findColumnIndex_(header, ['患者id', '患者ID', 'id', 'patientid'], 0);
  const statusCol = findColumnIndex_(header, ['bankstatus', 'ステータス', '結果', '入金結果'], 1);

  const map = {};
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const patientId = String(row[patientIdCol] || '').trim();
    if (!patientId) continue;
    const bankStatus = String(row[statusCol] || '').trim();
    map[patientId] = { patientId, bankStatus };
  }

  return map;
}
