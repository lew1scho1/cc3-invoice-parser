// ============================================================================
// CONFIG.GS - 설정 및 상수
// ============================================================================

var CONFIG = {
  SPREADSHEET_ID: '1d_QtwuD2MBt3rpG4atxizk3P3WJ2pA9LNzwFkpqgoSs',
  
  COMPANIES: {
    OUTRE: {
      name: 'OUTRE',
      dbSheet: 'Outre Active DB',
      orderSheet: 'ORDER_OUTRE',
      columns: {
        ITEM_NUMBER: 'ITEM NUMBER',
        ITEM_NAME: 'ITEM NAME',
        COLOR: 'COLOR',
        BARCODE: 'BARCODE'
      }
    },
    SNG: {
      name: 'SNG',
      dbSheet: 'SNG Active DB',
      orderSheet: 'ORDER_SNG',
      columns: {
        ITEM_NUMBER: 'Item Code',
        ITEM_NAME: 'Description',
        COLOR: 'Color',
        BARCODE: 'Barcode'
      }
    }
  },
  
  BARCODE_LENGTH: 12,
  MAX_COLORS_FOR_EXPANSION: 4,

  // 검색 인덱스 시트
  SEARCH_INDEX: {
    OUTRE: 'SEARCH_INDEX_OUTRE',
    SNG: 'SEARCH_INDEX_SNG'
  },

  // 히스토리/즐겨찾기 시트
  HISTORY_SHEET: 'USER_HISTORY',
  FAVORITES_SHEET: 'USER_FAVORITES',

  // 디버그 모드
  DEBUG: true,

  // ============================================================================
  // 상태 관리
  // ============================================================================

  // 상태 값 정의
  STATUS: {
    NEW: 'NEW',
    ACTIVE: 'ACTIVE',
    DISCONTINUE_UNKNOWN: 'DISCONTINUE_UNKNOWN',
    DISCONTINUED: 'DISCONTINUED'
  },

  // 상태별 UI 표시 설정
  STATUS_DISPLAY: {
    NEW: {
      label: '신규',
      color: '#4CAF50',
      badge: true,
      allowInput: true
    },
    ACTIVE: {
      label: '활성',
      color: 'transparent',
      badge: false,
      allowInput: true
    },
    DISCONTINUE_UNKNOWN: {
      label: '단종 예정',
      color: '#FFF9C4',
      badge: true,
      allowInput: true
    },
    DISCONTINUED: {
      label: '단종',
      color: '#E0E0E0',
      badge: true,
      allowInput: false
    }
  }
};

// ============================================================================
// 인보이스 설정
// ============================================================================

CONFIG.INVOICE = {
  // 시트 이름
  PARSING_SHEET: 'PARSING',      // 임시 파싱 결과 (레거시)
  PARSING_SHEETS: {
    SNG: 'PARSING_SNG',
    OUTRE: 'PARSING_OUTRE'
  },
  SNG_SHEET: 'INVOICE_SNG',      // 확정된 SNG 인보이스 (누적)
  OUTRE_SHEET: 'INVOICE_OUTRE',  // 확정된 OUTRE 인보이스 (누적)
  
  // Google Drive 폴더 ID (사용자가 설정)
  FOLDER_ID_PROPERTY: 'INVOICE_FOLDER_ID',
  
  // 헤더
  HEADERS: [
    'VENDOR',
    'Invoice No',
    'Invoice Date',
    'Total Amount',
    'Subtotal',
    'Discount',
    'Shipping',
    'Tax',
    'Line No',
    'Item ID',
    'UPC',
    'Description',
    'Brand',
    'Color',
    'Size/Length',
    'Qty Ordered',
    'Qty Shipped',
    'Unit Price',
    'Ext Price',
    'Memo'
  ],
  
  // 브랜드 매핑
  BRANDS: {
    SNG: 'Shake-N-Go',
    OUTRE: 'Outre'
  }
};

// ============================================================================
// 유틸리티 함수
// ============================================================================

/**
 * 월 차이 계산 (YYYYMM 형식)
 * @param {string} currentYm - 현재 월 (예: '202604')
 * @param {string} lastYm - 마지막 월 (예: '202601')
 * @return {number} 월 차이 (예: 3)
 */
function diffMonths(currentYm, lastYm) {
  var current = parseInt(currentYm, 10);
  var last = parseInt(lastYm, 10);

  if (isNaN(current) || isNaN(last)) {
    return 0;
  }

  var currentYear = Math.floor(current / 100);
  var currentMonth = current % 100;
  var lastYear = Math.floor(last / 100);
  var lastMonth = last % 100;

  return (currentYear - lastYear) * 12 + (currentMonth - lastMonth);
}

/**
 * 상태 계산 (파생값)
 * @param {string} firstSeenYm - 최초 등장 월 (YYYYMM)
 * @param {string} lastSeenYm - 마지막 등장 월 (YYYYMM)
 * @param {string} currentUploadYm - 현재 업로드 월 (YYYYMM)
 * @return {string} 상태 (NEW, ACTIVE, DISCONTINUE_UNKNOWN, DISCONTINUED)
 */
function calculateStatus(firstSeenYm, lastSeenYm, currentUploadYm) {
  // 신규 제품 (이번 달 처음 등장)
  if (firstSeenYm === currentUploadYm) {
    return CONFIG.STATUS.NEW;
  }

  // 월 차이 계산
  var monthsSince = diffMonths(currentUploadYm, lastSeenYm);

  // 활성 제품 (이번 달 등장)
  if (monthsSince === 0) {
    return CONFIG.STATUS.ACTIVE;
  }
  // 단종 예정 (1~2개월 미등장)
  else if (monthsSince >= 1 && monthsSince <= 2) {
    return CONFIG.STATUS.DISCONTINUE_UNKNOWN;
  }
  // 단종 확정 (3개월 이상 미등장)
  else {
    return CONFIG.STATUS.DISCONTINUED;
  }
}

function getDbSheet(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('DB sheet not found: ' + sheetName);
  }
  return sheet;
}
