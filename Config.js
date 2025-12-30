// ============================================================================
// CONFIG.GS - 설정 및 상수
// ============================================================================

var CONFIG = {
  SPREADSHEET_ID: '1d_QtwuD2MBt3rpG4atxizk3P3WJ2pA9LNzwFkpqgoSs',
  
  COMPANIES: {
    OUTRE: {
      name: 'OUTRE',
      dbSheet: 'DB_OUTRE',
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
      dbSheet: 'DB_SNG',
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
  
  // 디버그 모드
  DEBUG: true
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
