// ┌──────────────────────────────────────────────────┐
// │  🔒 관리자(결산 담당자) 이메일 목록 설정          │
// └──────────────────────────────────────────────────┘
const SYSTEM_ADMINS = [
  'hwijunjang@koreanair.com'
];

const ADMIN_SHEET_NAME = '__ADMIN_CONFIG__';
const SNAPSHOT_SHEET_NAME = '__DASHBOARD_DB__';
const MASTER_SHEET_NAME = '20_차입금마스터';
const PAGE_SIZE = 1000;
const MAX_ROWS_PER_REQUEST = 5000;

// ┌──────────────────────────────────────────────────┐
// │  커스텀 에러 클래스                               │
// └──────────────────────────────────────────────────┘

class AppError {
  constructor(message, code, details) {
    this.message = message;
    this.code = code;
    this.details = details || {};
    this.timestamp = new Date().toISOString();
  }

  toJSON() {
    return {
      error: this.message,
      code: this.code,
      details: this.details,
      timestamp: this.timestamp
    };
  }
}

// ┌──────────────────────────────────────────────────┐
// │  웹앱 진입점                                      │
// └──────────────────────────────────────────────────┘

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('항공기 금융 결산 시스템 (Enhanced)')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ┌──────────────────────────────────────────────────┐
// │  권한 관리                                        │
// └──────────────────────────────────────────────────┘

function getAppConfig() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(ADMIN_SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(ADMIN_SHEET_NAME);
      sheet.getRange("A1").setValue("관리자 이메일 (A열)");
      sheet.getRange("A2").setValue(SYSTEM_ADMINS[0]);
      sheet.getRange("B1").setValue("조회 허용 이메일 (B열)");
      sheet.getRange("C1").setValue("뷰어 허용 필터 (C열)");
      sheet.getRange("C2").setValue("DEAL");
      sheet.getRange("T1").setValue("버전 히스토리 (T열)");

      sheet.setColumnWidth(1, 250);
      sheet.setColumnWidth(2, 250);
      sheet.setColumnWidth(3, 200);
      sheet.setColumnWidth(20, 300);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e6b8af");
      sheet.getRange("T1").setFontWeight("bold").setBackground("#d9ead3");
    }

    // A열: 관리자
    const adminRange = sheet.getRange(2, 1, 39, 1).getValues();
    const admins = adminRange
      .flat()
      .filter(v => v && String(v).trim() !== '')
      .map(e => String(e).trim().toLowerCase());

    // B열: 조회 허용 이메일 (비어있으면 전체 허용)
    const viewerRange = sheet.getRange(2, 2, 39, 1).getValues();
    const allowedViewers = viewerRange
      .flat()
      .filter(v => v && String(v).trim() !== '')
      .map(e => String(e).trim().toLowerCase());

    // C열: 필터 허용 컬럼
    const filterRange = sheet.getRange(2, 3, 39, 1).getValues();
    const filters = filterRange
      .flat()
      .filter(v => v && String(v).trim() !== '')
      .map(s => String(s).trim());

    const uniqueAdmins = [...new Set([
      ...SYSTEM_ADMINS.map(e => e.toLowerCase()),
      ...admins
    ])];

    return {
      admins: uniqueAdmins,
      allowedViewers,   // 비어있으면 전체 허용
      allowedFilters: filters
    };
  } catch (e) {
    Logger.log('getAppConfig error: ' + e.toString());
    return {
      admins: SYSTEM_ADMINS.map(e => e.toLowerCase()),
      allowedViewers: [],
      allowedFilters: []
    };
  }
}

function getUserPermission() {
  try {
    let email = Session.getActiveUser().getEmail();
    if (!email) {
      email = Session.getEffectiveUser().getEmail();
    }

    const currentEmail = String(email || '').toLowerCase().trim();
    const conf = getAppConfig();
    const isAdmin = conf.admins.includes(currentEmail);

    // B열이 비어있으면 전체 허용, 값이 있으면 목록에 있어야 접근 가능
    const isAllowed = isAdmin ||
      conf.allowedViewers.length === 0 ||
      conf.allowedViewers.includes(currentEmail);

    logActivity('getUserPermission', { email: currentEmail, isAdmin, isAllowed });

    return { email: email || 'unknown', isAdmin, isAllowed };
  } catch (e) {
    Logger.log('getUserPermission error: ' + e.toString());
    return { email: 'unknown', isAdmin: false, isAllowed: false };
  }
}

function verifyAdminPermission() {
  const perm = getUserPermission();
  if (!perm.isAdmin) {
    throw new AppError(
      "관리자 권한이 없습니다.",
      "PERMISSION_DENIED",
      { email: perm.email }
    );
  }
  return perm;
}

// ┌──────────────────────────────────────────────────┐
// │  데이터 관리                                      │
// └──────────────────────────────────────────────────┘

function validateData(data) {
  const errors   = [];
  const warnings = [];
  const dealSet  = new Set();

  const DATE_RE  = /^\d{4}-\d{2}-\d{2}$/;
  const today    = new Date(); today.setHours(0,0,0,0);

  const numVal = (v) => parseFloat(String(v ?? '').replace(/,/g, '').replace(/%/g, ''));

  data.forEach((row, index) => {
    const rowNum = index + 1;
    const label  = row['DEAL'] ? `[${row['DEAL']}]` : `[${rowNum}행]`;

    // ── DEAL ──────────────────────────────────────────
    if (!row['DEAL'] || String(row['DEAL']).trim() === '') {
      errors.push({ row: rowNum, field: 'DEAL', message: 'DEAL 값이 비어있습니다.' });
    } else {
      const dealId = String(row['DEAL']).trim();
      if (dealSet.has(dealId)) {
        warnings.push({ row: rowNum, field: 'DEAL', message: `${label} 중복된 DEAL` });
      }
      dealSet.add(dealId);
    }

    // ── 필수값 누락 ────────────────────────────────────
    if (!row['차입처'] || String(row['차입처']).trim() === '' || String(row['차입처']).trim() === '-') {
      warnings.push({ row: rowNum, field: '차입처', message: `${label} 차입처 누락` });
    }
    if (!row['통화'] || String(row['통화']).trim() === '') {
      warnings.push({ row: rowNum, field: '통화', message: `${label} 통화 누락` });
    }
    if (!row['금리유형'] || String(row['금리유형']).trim() === '') {
      warnings.push({ row: rowNum, field: '금리유형', message: `${label} 금리유형 누락` });
    }

    // ── 날짜 ──────────────────────────────────────────
    const fromStr = String(row['FROM'] || '').trim();
    const toStr   = String(row['TO']   || '').trim();

    if (!fromStr) {
      warnings.push({ row: rowNum, field: 'FROM', message: `${label} FROM 날짜 누락` });
    } else if (!DATE_RE.test(fromStr)) {
      warnings.push({ row: rowNum, field: 'FROM', message: `${label} FROM 날짜 형식 오류 (YYYY-MM-DD 필요)` });
    }

    if (!toStr) {
      warnings.push({ row: rowNum, field: 'TO', message: `${label} TO 날짜 누락` });
    } else if (!DATE_RE.test(toStr)) {
      warnings.push({ row: rowNum, field: 'TO', message: `${label} TO 날짜 형식 오류 (YYYY-MM-DD 필요)` });
    }

    // FROM > TO
    if (DATE_RE.test(fromStr) && DATE_RE.test(toStr)) {
      if (new Date(fromStr) > new Date(toStr)) {
        errors.push({ row: rowNum, field: 'FROM/TO', message: `${label} FROM(${fromStr})이 TO(${toStr})보다 늦습니다` });
      }
    }

    // 만기 초과 잔액 (TO가 오늘 이전인데 잔액 > 0)
    const balance = numVal(row['당월잔액'] ?? row['원화환산잔액']);
    if (DATE_RE.test(toStr) && new Date(toStr) < today && !isNaN(balance) && balance > 0) {
      warnings.push({ row: rowNum, field: '잔액', message: `${label} 만기(${toStr}) 경과했으나 잔액 존재 (${balance.toLocaleString()})` });
    }
    // ── 금액 ──────────────────────────────────────────
    const krwBal    = numVal(row['원화환산잔액']);
    const monthBal  = numVal(row['당월잔액']);
    const loanAmt   = numVal(row['차입금액']);

    if (row['원화환산잔액'] !== undefined && row['원화환산잔액'] !== '') {
      if (isNaN(krwBal)) {
        warnings.push({ row: rowNum, field: '원화환산잔액', message: `${label} 원화환산잔액이 숫자가 아닙니다` });
      } else if (krwBal < 0) {
        warnings.push({ row: rowNum, field: '원화환산잔액', message: `${label} 원화환산잔액이 음수입니다` });
      }
    }

    if (!isNaN(monthBal) && !isNaN(loanAmt) && loanAmt > 0 && monthBal > loanAmt) {
      warnings.push({ row: rowNum, field: '당월잔액', message: `${label} 당월잔액(${monthBal.toLocaleString()})이 차입금액(${loanAmt.toLocaleString()}) 초과` });
    }

    // 원화환산잔액이 0인데 당월잔액이 양수인 경우만 경고 (0은 정상값)
    if (!isNaN(krwBal) && krwBal === 0 && !isNaN(monthBal) && monthBal > 0) {
      warnings.push({ row: rowNum, field: '잔액', message: `${label} 원화환산잔액이 0인데 당월잔액 존재` });
    }

    if (!isNaN(monthBal) && monthBal < 0) {
      warnings.push({ row: rowNum, field: '당월잔액', message: `${label} 당월잔액이 음수입니다` });
    }

    // ── 금리 ──────────────────────────────────────────
    const rate     = numVal(row['적용율']);
    const baseRate = numVal(row['기준금리']);
    const spread   = numVal(row['스프레드']);
    const rateType = String(row['금리유형'] || '').trim();

    if (row['적용율'] !== undefined && row['적용율'] !== '') {
      if (isNaN(rate)) {
        warnings.push({ row: rowNum, field: '적용율', message: `${label} 적용율이 숫자가 아닙니다` });
      } else if (rate < 0 || rate > 1) {
        warnings.push({ row: rowNum, field: '적용율', message: `${label} 적용율(${(rate*100).toFixed(2)}%)이 비정상 범위입니다` });
      }
    }

    // 변동금리인데 기준금리 누락 (숫자든 텍스트든 값이 있으면 통과)
    if (rateType.includes('변동') && (!row['기준금리'] || String(row['기준금리']).trim() === '')) {
      warnings.push({ row: rowNum, field: '기준금리', message: `${label} 변동금리인데 기준금리 누락` });
    }

    // 적용율이 있는데 스프레드 누락 (고정금리 제외)
    if (!isNaN(rate) && rate > 0 && !rateType.includes('고정') && (!row['스프레드'] || isNaN(spread))) {
      warnings.push({ row: rowNum, field: '스프레드', message: `${label} 적용율 있으나 스프레드 누락` });
    }
  });

  return {
    errors,
    warnings,
    isValid: errors.length === 0,
    summary: {
      totalRows:    data.length,
      errorCount:   errors.length,
      warningCount: warnings.length
    }
  };
}

function getData(startRow, maxRows) {
  startRow = startRow || 0;
  maxRows = Math.min(maxRows || PAGE_SIZE, MAX_ROWS_PER_REQUEST);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const conf = getAppConfig();

    let sheet = ss.getSheetByName(SNAPSHOT_SHEET_NAME);
    let dataSource = 'snapshot';

    if (!sheet) {
      sheet = ss.getSheetByName(MASTER_SHEET_NAME);
      dataSource = 'master';
    }

    if (!sheet) {
      throw new AppError(
        "데이터 시트를 찾을 수 없습니다.",
        "SHEET_NOT_FOUND",
        { attemptedSheets: [SNAPSHOT_SHEET_NAME, MASTER_SHEET_NAME] }
      );
    }

    const totalRows = sheet.getLastRow();

    if (totalRows === 0) {
      throw new AppError(
        "시트에 데이터가 없습니다.",
        "EMPTY_SHEET",
        { sheetName: sheet.getName() }
      );
    }

    let headerRowIndex = -1;
    const searchRows = Math.min(20, totalRows);
    const searchRange = sheet.getRange(1, 1, searchRows, sheet.getLastColumn());
    const searchValues = searchRange.getDisplayValues();

    for (let i = 0; i < searchRows; i++) {
      if (searchValues[i] && searchValues[i].some(c =>
        String(c).toUpperCase().replace(/\s/g, '') === 'DEAL'
      )) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new AppError(
        "'DEAL' 컬럼을 찾을 수 없습니다.",
        "HEADER_NOT_FOUND",
        { searchedRows: searchRows }
      );
    }

    const headers = searchValues[headerRowIndex].map(h => String(h).replace(/\s/g, ''));

    const dataStartRow = headerRowIndex + 2;
    const actualStartRow = dataStartRow + startRow;
    const endRow = Math.min(actualStartRow + maxRows - 1, totalRows);
    const rowsToFetch = endRow - actualStartRow + 1;

    if (rowsToFetch <= 0) {
      return {
        columns: headers.filter(h => h && h !== ''),
        data: [],
        hasMore: false,
        totalRows: 0,
        currentPage: { start: startRow, end: startRow },
        dataSource: dataSource
      };
    }

    const dataRange = sheet.getRange(actualStartRow, 1, rowsToFetch, headers.length);
    const dataValues = dataRange.getDisplayValues();

    let lastUpdate = "실시간 데이터";
    let metadata = {};

    try {
      const note = sheet.getRange("A1").getNote();
      if (note) {
        metadata = JSON.parse(note);
        if (metadata.lastUpdate) {
          lastUpdate = new Date(metadata.lastUpdate).toLocaleString('ko-KR');
        }
      }
    } catch (e) {
      Logger.log('Failed to parse metadata: ' + e.toString());
    }

    const cleanData = [];

    for (const row of dataValues) {
      const obj = {};
      let hasValidDeal = false;

      headers.forEach((h, i) => {
        if (!h) return;

        let val = row[i];

        if (h.includes('잔액') || h.includes('금액') || h.includes('율')) {
          const strVal = String(val);
          const hasPercent = strVal.includes('%');

          val = parseFloat(strVal.replace(/,/g, '').replace(/%/g, '')) || 0;

          if (h.includes('율') && (hasPercent || Math.abs(val) > 1)) {
            val /= 100;
          }
        }

        obj[h] = val;

        if (h === 'DEAL' && val && String(val).trim() !== '' && String(val).trim() !== '-') {
          hasValidDeal = true;
        }
      });

      if (hasValidDeal) {
        cleanData.push(obj);
      }
    }

    const validation = validateData(cleanData);
    const validColumns = headers.filter(h => h && h !== '');

    return {
      columns: validColumns,
      data: cleanData,
      hasMore: endRow < totalRows,
      totalRows: totalRows - (headerRowIndex + 1),
      currentPage: {
        start: startRow,
        end: startRow + cleanData.length
      },
      lastUpdate: lastUpdate,
      allowedFilters: conf.allowedFilters,
      validation: validation,
      dataSource: dataSource,
      metadata: metadata
    };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('getData error: ' + e.toString());
    return new AppError(
      "데이터 로드 중 오류 발생: " + e.message,
      "DATA_LOAD_FAILED",
      { originalError: e.toString() }
    ).toJSON();
  }
}

// ┌──────────────────────────────────────────────────┐
// │  상환스케줄 데이터                                │
// └──────────────────────────────────────────────────┘

/**
 * '13_상환스케줄' 시트에서 D열(날짜)과 I열(금액)을 읽어 연도별로 합산
 * @returns {Object} { yearlyData: Array<{year, amount}> }
 */
function getScheduleData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('13_상환스케줄');

    if (!sheet) {
      throw new AppError(
        "'13_상환스케줄' 시트를 찾을 수 없습니다.",
        "SHEET_NOT_FOUND",
        { sheetName: '13_상환스케줄' }
      );
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { yearlyData: [] };
    }

    // D열(4번째)과 I열(9번째) 읽기 — 4행부터 데이터 시작 (1~3행 헤더)
    const dataStartRow = 4;
    const rowCount = lastRow - dataStartRow + 1;
    if (rowCount < 1) return { yearlyData: [] };

    const dateRange   = sheet.getRange(dataStartRow, 4, rowCount, 1).getDisplayValues(); // D열
    const amountRange = sheet.getRange(dataStartRow, 9, rowCount, 1).getDisplayValues(); // I열

    // 연도별 합산
    const yearMap = {};

    for (let i = 0; i < dateRange.length; i++) {
      const dateStr  = String(dateRange[i][0] || '').trim();
      const amountStr = String(amountRange[i][0] || '').trim();

      if (!dateStr || !amountStr) continue;

      // 날짜에서 연도 추출 (YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD 등 대응)
      const yearMatch = dateStr.match(/^(\d{4})/);
      if (!yearMatch) continue;

      const year = yearMatch[1];
      const amount = parseFloat(amountStr.replace(/,/g, '')) || 0;

      if (amount === 0) continue;

      yearMap[year] = (yearMap[year] || 0) + amount;
    }

    // 연도 오름차순 정렬
    const years = Object.keys(yearMap).map(Number).sort((a, b) => a - b);

    // 빈 연도 없이 연속된 범위로 채우기
    const yearlyData = [];
    if (years.length > 0) {
      const minYear = years[0];
      const maxYear = years[years.length - 1];
      for (let y = minYear; y <= maxYear; y++) {
        yearlyData.push({ year: String(y), amount: yearMap[String(y)] || 0 });
      }
    }

    logActivity('getScheduleData', { rows: dateRange.length, years: yearlyData.length });

    return { yearlyData };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('getScheduleData error: ' + e.toString());
    return new AppError(
      "상환스케줄 데이터 로드 실패: " + e.message,
      "SCHEDULE_DATA_FAILED",
      { originalError: e.toString() }
    ).toJSON();
  }
}

// ┌──────────────────────────────────────────────────┐
// │  버전 히스토리                                    │
// └──────────────────────────────────────────────────┘

/**
 * __ADMIN_CONFIG__ T열에서 버전 히스토리 읽기
 * T1: 헤더명 (없으면 "버전 히스토리" 기본값)
 * T2~: 각 버전 설명 (한 줄에 하나씩)
 * @returns {Object} { header: string, versions: Array<{index, description}> }
 */
function getVersionHistory() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ADMIN_SHEET_NAME);

    if (!sheet) {
      throw new AppError(
        "설정 시트를 찾을 수 없습니다.",
        "SHEET_NOT_FOUND",
        { sheetName: ADMIN_SHEET_NAME }
      );
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return { header: '버전 히스토리', versions: [] };
    }

    // S열(19번째, 카테고리)과 T열(20번째, 설명) 동시 읽기
    const range = sheet.getRange(1, 19, lastRow, 2);
    const values = range.getDisplayValues();

    // T1 셀을 헤더명으로 사용 (비어있으면 기본값)
    const header = String(values[0][1] || '').trim() || '버전 히스토리';

    // T2부터 내용, 빈 행 제외
    const versions = values
      .slice(1)
      .map((row, i) => ({
        index: i + 1,
        category: String(row[0] || '').trim(),   // S열: [기능추가] 등
        description: String(row[1] || '').trim() // T열: 설명
      }))
      .filter(v => v.description !== '');

    logActivity('getVersionHistory', { count: versions.length });

    return { header, versions };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('getVersionHistory error: ' + e.toString());
    return new AppError(
      "버전 히스토리 로드 실패: " + e.message,
      "VERSION_HISTORY_FAILED",
      { originalError: e.toString() }
    ).toJSON();
  }
}

// ┌──────────────────────────────────────────────────┐
// │  피드백                                           │
// └──────────────────────────────────────────────────┘

/**
 * 피드백 제출 (Admin 전용)
 * __ADMIN_CONFIG__ V열(날짜), W열(작성자), X열(내용)
 */
function submitFeedback(content) {
  try {
    const text = String(content || '').trim();
    if (!text) {
      return { success: false, message: '내용을 입력해주세요.' };
    }
    if (text.length > 500) {
      return { success: false, message: '500자 이내로 입력해주세요.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!sheet) {
      throw new AppError('설정 시트를 찾을 수 없습니다.', 'SHEET_NOT_FOUND', { sheetName: ADMIN_SHEET_NAME });
    }

    const now = new Date();
    const email = Session.getEffectiveUser().getEmail();

    // V1 헤더가 없으면 먼저 세팅
    if (sheet.getRange(1, 22).getValue() === '') {
      sheet.getRange(1, 22).setValue('피드백_날짜').setFontWeight('bold').setBackground('#cfe2f3');
      sheet.getRange(1, 23).setValue('피드백_작성자').setFontWeight('bold').setBackground('#cfe2f3');
      sheet.getRange(1, 24).setValue('피드백_내용').setFontWeight('bold').setBackground('#cfe2f3');
      sheet.setColumnWidth(22, 160);
      sheet.setColumnWidth(23, 200);
      sheet.setColumnWidth(24, 400);
    }

    // V열(22번째) 기준으로 마지막 데이터 행 탐색 → 항상 V2부터 쌓임
    const vValues = sheet.getRange(1, 22, sheet.getMaxRows(), 1).getValues();
    let vLastRow = 1;
    for (let i = 0; i < vValues.length; i++) {
      if (String(vValues[i][0]).trim() !== '') vLastRow = i + 1;
    }

    // V열 다음 빈 행에 데이터 추가
    sheet.getRange(vLastRow + 1, 22, 1, 3).setValues([[
      now.toLocaleString('ko-KR'),
      email,
      text
    ]]);

    logActivity('submitFeedback', { email, length: text.length });

    return {
      success: true,
      feedback: {
        date: now.toLocaleString('ko-KR'),
        author: email,
        content: text
      }
    };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('submitFeedback error: ' + e.toString());
    return { success: false, message: '피드백 저장 실패: ' + e.message };
  }
}

/**
 * 피드백 목록 조회 (Admin 전용)
 * @returns {Object} { feedbacks: Array<{date, author, content}> }
 */
function getFeedback() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ADMIN_SHEET_NAME);
    if (!sheet) {
      throw new AppError('설정 시트를 찾을 수 없습니다.', 'SHEET_NOT_FOUND', { sheetName: ADMIN_SHEET_NAME });
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { feedbacks: [] };

    // V열(22번째) 기준으로 실제 마지막 행 탐색
    const vValues = sheet.getRange(1, 22, sheet.getMaxRows(), 1).getValues();
    let vLastRow = 1;
    for (let i = 0; i < vValues.length; i++) {
      if (String(vValues[i][0]).trim() !== '') vLastRow = i + 1;
    }
    if (vLastRow < 2) return { feedbacks: [] };

    const values = sheet.getRange(2, 22, vLastRow - 1, 3).getDisplayValues();

    const feedbacks = values
      .map((row, i) => ({
        index: i + 1,
        date: String(row[0] || '').trim(),
        author: String(row[1] || '').trim(),
        content: String(row[2] || '').trim()
      }))
      .filter(f => f.content !== '');

    return { feedbacks };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('getFeedback error: ' + e.toString());
    return new AppError(
      '피드백 조회 실패: ' + e.message,
      'FEEDBACK_LOAD_FAILED',
      { originalError: e.toString() }
    ).toJSON();
  }
}

// ┌──────────────────────────────────────────────────┐
// │  히스토리 데이터                                  │
// └──────────────────────────────────────────────────┘

/**
 * 22_차입금History 시트에서 시계열 데이터 읽기
 * @returns {Object} { columns, rows, deals }
 */
function getHistoryData() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('22_차입금History');

    if (!sheet) {
      throw new AppError(
        "'22_차입금History' 시트를 찾을 수 없습니다.",
        'SHEET_NOT_FOUND',
        { sheetName: '22_차입금History' }
      );
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { columns: [], rows: [], deals: [] };

    const rawData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = rawData[0].map(h => String(h).trim());

    const rows    = [];
    const dealSet = new Set();

    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];
      const obj = {};
      headers.forEach((h, j) => { obj[h] = String(row[j] || '').trim(); });

      if (!obj['기준월'] || obj['기준월'] === '') continue;

      rows.push(obj);
      if (obj['DEAL']) dealSet.add(obj['DEAL']);
    }

    const deals = [...dealSet].sort();

    logActivity('getHistoryData', { rows: rows.length, deals: deals.length });

    return { columns: headers, rows, deals };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('getHistoryData error: ' + e.toString());
    return new AppError(
      '히스토리 데이터 로드 실패: ' + e.message,
      'HISTORY_LOAD_FAILED',
      { originalError: e.toString() }
    ).toJSON();
  }
}

// ┌──────────────────────────────────────────────────┐
// │  접속자 현황 (Presence)                           │
// └──────────────────────────────────────────────────┘

/**
 * 접속 등록 / 핑 갱신
 * ScriptProperties에 JSON으로 저장: { email: { name, ts } }
 */
function presencePing() {
  try {
    const email = Session.getEffectiveUser().getEmail();
    if (!email) return { success: false };

    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty('presence') || '{}';
    const map   = JSON.parse(raw);

    map[email] = { name: email.split('@')[0], ts: Date.now() };
    props.setProperty('presence', JSON.stringify(map));

    // 현재 접속자 목록 반환
    const users = Object.entries(map).map(([e, v]) => ({
      email: e,
      name:  v.name,
      initials: v.name.slice(0, 2).toUpperCase()
    }));

    return { success: true, users };
  } catch (e) {
    Logger.log('presencePing error: ' + e.toString());
    return { success: false, users: [] };
  }
}

/**
 * 접속 해제 (탭 닫을 때)
 */
function presenceLeave() {
  try {
    const email = Session.getEffectiveUser().getEmail();
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty('presence') || '{}';
    const map   = JSON.parse(raw);
    delete map[email];
    props.setProperty('presence', JSON.stringify(map));
    return { success: true };
  } catch (e) {
    return { success: false };
  }
}

// ┌──────────────────────────────────────────────────┐
// │  활동 로그                                        │
// └──────────────────────────────────────────────────┘

function logActivity(action, details) {
  try {
    const user = Session.getEffectiveUser().getEmail();
    const timestamp = new Date().toISOString();

    Logger.log(JSON.stringify({
      timestamp: timestamp,
      user: user,
      action: action,
      details: details
    }));
  } catch (e) {
    Logger.log('logActivity error: ' + e.toString());
  }
}

// ┌──────────────────────────────────────────────────┐
// │  히스토리 데이터                                  │
// └──────────────────────────────────────────────────┘

/**
 * 22_차입금History 시트에서 시계열 데이터 읽기
 * @returns {Object} { columns, rows, deals }
 *   - columns: 헤더 배열
 *   - rows: 전체 데이터 배열 (각 행은 {기준월, DEAL, ...} 객체)
 *   - deals: DEAL 목록 (중복 제거, 정렬)
 */
function getHistoryData() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('22_차입금History');

    if (!sheet) {
      throw new AppError(
        "'22_차입금History' 시트를 찾을 수 없습니다.",
        'SHEET_NOT_FOUND',
        { sheetName: '22_차입금History' }
      );
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return { columns: [], rows: [], deals: [] };

    const rawData = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
    const headers = rawData[0].map(h => String(h).trim());

    const rows = [];
    const dealSet = new Set();

    for (let i = 1; i < rawData.length; i++) {
      const row = rawData[i];
      const obj = {};
      headers.forEach((h, j) => { obj[h] = String(row[j] || '').trim(); });

      // 기준월이 없는 행 스킵
      if (!obj['기준월'] || obj['기준월'] === '') continue;

      rows.push(obj);
      if (obj['DEAL']) dealSet.add(obj['DEAL']);
    }

    const deals = [...dealSet].sort();

    logActivity('getHistoryData', { rows: rows.length, deals: deals.length });

    return { columns: headers, rows, deals };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('getHistoryData error: ' + e.toString());
    return new AppError(
      '히스토리 데이터 로드 실패: ' + e.message,
      'HISTORY_LOAD_FAILED',
      { originalError: e.toString() }
    ).toJSON();
  }
}