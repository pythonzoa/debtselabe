// ┌──────────────────────────────────────────────────┐
// │  🔒 관리자(결산 담당자) 이메일 목록 설정          │
// └──────────────────────────────────────────────────┘
const SYSTEM_ADMINS = [
  'hwijunjang@koreanair.com'
];

const ADMIN_SHEET_NAME = '__ADMIN_CONFIG__';
const SNAPSHOT_SHEET_NAME = '__DASHBOARD_DB__';
const MASTER_SHEET_NAME = '20_차입금마스터';
const HISTORY_SHEET_NAME = '22_차입금History';

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
      sheet.getRange("C1").setValue("뷰어 허용 필터 (C열)");
      sheet.getRange("C2").setValue("DEAL");
      sheet.getRange("T1").setValue("버전 히스토리 (T열)");

      sheet.setColumnWidth(1, 250);
      sheet.setColumnWidth(2, 50);
      sheet.setColumnWidth(3, 200);
      sheet.setColumnWidth(20, 300);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#e6b8af");
      sheet.getRange("T1").setFontWeight("bold").setBackground("#d9ead3");
    }

    const adminRange = sheet.getRange("A2:A").getValues();
    const admins = adminRange
      .flat()
      .filter(v => v && String(v).trim() !== '')
      .map(e => String(e).trim().toLowerCase());

    const filterRange = sheet.getRange("C2:C").getValues();
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
      allowedFilters: filters
    };
  } catch (e) {
    Logger.log('getAppConfig error: ' + e.toString());
    return {
      admins: SYSTEM_ADMINS.map(e => e.toLowerCase()),
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

    logActivity('getUserPermission', { email: currentEmail, isAdmin: isAdmin });

    return { email: email || 'unknown', isAdmin: isAdmin };
  } catch (e) {
    Logger.log('getUserPermission error: ' + e.toString());
    return { email: 'unknown', isAdmin: false };
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

  const currentUser = Session.getEffectiveUser().getEmail().toLowerCase();
  const config = getAppConfig();
  if (!config.admins.includes(currentUser)) {
    throw new AppError(
      "권한 검증 실패",
      "PERMISSION_VERIFICATION_FAILED",
      { email: currentUser }
    );
  }

  return perm;
}

// ┌──────────────────────────────────────────────────┐
// │  데이터 관리                                      │
// └──────────────────────────────────────────────────┘

function validateData(data) {
  const errors = [];
  const warnings = [];
  const dealSet = new Set();

  data.forEach((row, index) => {
    const rowNum = index + 1;

    if (!row['DEAL'] || String(row['DEAL']).trim() === '') {
      errors.push({
        row: rowNum,
        field: 'DEAL',
        message: 'DEAL 값이 비어있습니다.'
      });
    } else {
      const dealId = String(row['DEAL']).trim();
      if (dealSet.has(dealId)) {
        warnings.push({
          row: rowNum,
          field: 'DEAL',
          message: `중복된 DEAL: ${dealId}`
        });
      }
      dealSet.add(dealId);
    }

    if (!row['차입처'] || String(row['차입처']).trim() === '' || String(row['차입처']).trim() === '-') {
      warnings.push({
        row: rowNum,
        field: '차입처',
        message: '차입처 정보가 없습니다.'
      });
    }

    if (row['원화환산잔액'] !== undefined && row['원화환산잔액'] !== '') {
      const balance = parseFloat(String(row['원화환산잔액']).replace(/,/g, ''));
      if (isNaN(balance)) {
        warnings.push({
          row: rowNum,
          field: '원화환산잔액',
          message: '숫자가 아닌 값이 포함되어 있습니다.'
        });
      }
    }

    if (row['TO']) {
      const dateStr = String(row['TO']).trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
        warnings.push({
          row: rowNum,
          field: 'TO',
          message: '날짜 형식이 올바르지 않습니다 (YYYY-MM-DD 형식 필요).'
        });
      }
    }

    if (row['적용율'] !== undefined && row['적용율'] !== '') {
      const rate = parseFloat(String(row['적용율']).replace(/,/g, '').replace(/%/g, ''));
      if (isNaN(rate)) {
        warnings.push({
          row: rowNum,
          field: '적용율',
          message: '금리 값이 올바르지 않습니다.'
        });
      } else if (rate < 0 || rate > 100) {
        warnings.push({
          row: rowNum,
          field: '적용율',
          message: '금리 값이 비정상적입니다 (0-100% 범위 초과).'
        });
      }
    }
  });

  return {
    errors,
    warnings,
    isValid: errors.length === 0,
    summary: {
      totalRows: data.length,
      errorCount: errors.length,
      warningCount: warnings.length
    }
  };
}

function publishData() {
  verifyAdminPermission();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(MASTER_SHEET_NAME);

    if (!sourceSheet) {
      throw new AppError(
        "원본 시트를 찾을 수 없습니다.",
        "SHEET_NOT_FOUND",
        { sheetName: MASTER_SHEET_NAME }
      );
    }

    let targetSheet = ss.getSheetByName(SNAPSHOT_SHEET_NAME);
    if (!targetSheet) {
      targetSheet = ss.insertSheet(SNAPSHOT_SHEET_NAME);
      targetSheet.hideSheet();
    }

    const sourceData = sourceSheet.getDataRange().getValues();

    if (sourceData.length === 0) {
      throw new AppError(
        "원본 시트에 데이터가 없습니다.",
        "EMPTY_SHEET",
        { sheetName: MASTER_SHEET_NAME }
      );
    }

    targetSheet.clear();
    targetSheet.getRange(1, 1, sourceData.length, sourceData[0].length).setValues(sourceData);

    const now = new Date();
    const metadata = {
      lastUpdate: now.toISOString(),
      updatedBy: Session.getEffectiveUser().getEmail(),
      rowCount: sourceData.length,
      columnCount: sourceData[0].length
    };

    targetSheet.getRange("A1").setNote(JSON.stringify(metadata, null, 2));

    logActivity('publishData', {
      rows: sourceData.length,
      timestamp: now.toISOString()
    });

    return {
      success: true,
      timestamp: now.toLocaleString('ko-KR'),
      metadata: metadata
    };
  } catch (e) {
    if (e instanceof AppError) throw e;
    throw new AppError(
      "데이터 발행 실패: " + e.message,
      "PUBLISH_FAILED",
      { originalError: e.toString() }
    );
  }
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

function getHistoryData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(HISTORY_SHEET_NAME);

    if (!sheet) {
      throw new AppError(
        "시계열 데이터 시트를 찾을 수 없습니다.",
        "SHEET_NOT_FOUND",
        { sheetName: HISTORY_SHEET_NAME }
      );
    }

    const dataRange = sheet.getDataRange();
    if (dataRange.getNumRows() === 0) {
      throw new AppError(
        "시트에 데이터가 없습니다.",
        "EMPTY_SHEET",
        { sheetName: HISTORY_SHEET_NAME }
      );
    }

    const values = dataRange.getDisplayValues();

    let headerRowIndex = -1;
    const maxSearchRows = Math.min(20, values.length);

    for (let i = 0; i < maxSearchRows; i++) {
      if (values[i] && values[i].some(c => {
        const normalized = String(c).toUpperCase().replace(/\s/g, '');
        return normalized === '기준월' || normalized === 'DEAL' || normalized === '기준일';
      })) {
        headerRowIndex = i;
        break;
      }
    }

    if (headerRowIndex === -1) {
      throw new AppError(
        "헤더 행을 찾을 수 없습니다.",
        "HEADER_NOT_FOUND",
        { requiredColumns: ['기준월', 'DEAL', '기준일'] }
      );
    }

    const headers = values[headerRowIndex].map(h => String(h).replace(/\s/g, ''));
    const dataRows = values.slice(headerRowIndex + 1);

    const cleanData = [];

    for (const row of dataRows) {
      const obj = {};
      let hasValidData = false;

      headers.forEach((h, i) => {
        if (!h) return;

        let val = row[i];

        if (h.includes('잔액') || h.includes('금액')) {
          const strVal = String(val);
          val = parseFloat(strVal.replace(/,/g, '').replace(/%/g, '')) || 0;
        }

        obj[h] = val;

        if ((h === '기준월' || h === '기준일' || h === 'DEAL') &&
            val && String(val).trim() !== '') {
          hasValidData = true;
        }
      });

      if (hasValidData) {
        cleanData.push(obj);
      }
    }

    const validColumns = headers.filter(h => h && h !== '');

    return {
      columns: validColumns,
      data: cleanData,
      totalRows: cleanData.length
    };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('getHistoryData error: ' + e.toString());
    return new AppError(
      "시계열 데이터 로드 중 오류 발생: " + e.message,
      "HISTORY_DATA_LOAD_FAILED",
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

    // T열 전체 읽기 (20번째 열 = T열)
    const range = sheet.getRange(1, 20, lastRow, 1);
    const values = range.getDisplayValues();

    // T1 셀을 헤더명으로 사용 (비어있으면 기본값)
    const header = String(values[0][0] || '').trim() || '버전 히스토리';

    // T2부터 내용, 빈 행 제외
    const versions = values
      .slice(1)
      .map((row, i) => ({
        index: i + 1,
        description: String(row[0] || '').trim()
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
// │  관리자 관리                                      │
// └──────────────────────────────────────────────────┘

function addAdmin(email) {
  verifyAdminPermission();

  try {
    const normalizedEmail = String(email).trim().toLowerCase();

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!normalizedEmail || !emailRegex.test(normalizedEmail)) {
      return {
        success: false,
        message: "유효한 이메일 주소를 입력해주세요."
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ADMIN_SHEET_NAME);

    if (!sheet) {
      throw new AppError(
        "설정 시트를 찾을 수 없습니다.",
        "SHEET_NOT_FOUND",
        { sheetName: ADMIN_SHEET_NAME }
      );
    }

    const conf = getAppConfig();
    if (conf.admins.includes(normalizedEmail)) {
      return {
        success: false,
        message: "이미 등록된 관리자입니다."
      };
    }

    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(normalizedEmail);

    logActivity('addAdmin', { email: normalizedEmail });

    return {
      success: true,
      message: normalizedEmail + "이(가) 관리자로 추가되었습니다."
    };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('addAdmin error: ' + e.toString());
    return {
      success: false,
      message: "관리자 추가 실패: " + e.message
    };
  }
}

function addAllowedFilter(filterName) {
  verifyAdminPermission();

  try {
    const normalizedFilter = String(filterName).trim();

    if (!normalizedFilter) {
      return {
        success: false,
        message: "필터 이름을 입력해주세요."
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ADMIN_SHEET_NAME);

    if (!sheet) {
      throw new AppError(
        "설정 시트를 찾을 수 없습니다.",
        "SHEET_NOT_FOUND",
        { sheetName: ADMIN_SHEET_NAME }
      );
    }

    const conf = getAppConfig();
    if (conf.allowedFilters.includes(normalizedFilter)) {
      return {
        success: false,
        message: "이미 등록된 필터입니다."
      };
    }

    const filterRange = sheet.getRange("C:C").getValues();
    let lastFilterRow = 1;
    for (let i = 0; i < filterRange.length; i++) {
      if (filterRange[i][0] && String(filterRange[i][0]).trim() !== '') {
        lastFilterRow = i + 1;
      }
    }

    sheet.getRange(lastFilterRow + 1, 3).setValue(normalizedFilter);

    logActivity('addAllowedFilter', { filter: normalizedFilter });

    return {
      success: true,
      message: normalizedFilter + " 필터가 추가되었습니다."
    };
  } catch (e) {
    if (e instanceof AppError) return e.toJSON();
    Logger.log('addAllowedFilter error: ' + e.toString());
    return {
      success: false,
      message: "필터 추가 실패: " + e.message
    };
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