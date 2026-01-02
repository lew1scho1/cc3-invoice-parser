// ============================================================================
// CODE.GS - 메인 진입점 및 초기화
// ============================================================================

/**
 * 웹앱 진입점
 */
function doGet() {
  try {
    debugLog('doGet 호출');
    
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('CC3 HAIR ORDER APP')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  } catch (error) {
    debugLog('doGet 오류', { error: error.toString() });
    return HtmlService.createHtmlOutput('앱 로딩 오류: ' + error.toString());
  }
}

/**
 * 메뉴에 기능 추가
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('CC3 ORDER APP')
    .addItem('🔧 초기 설정 (제품 시트 생성)', 'initializeSheets')
    .addItem('📄 인보이스 시트 초기화', 'initializeInvoiceSheets')
    .addSeparator()
    .addItem('🧾 Upload -> Active DB 생성', 'processUploadToActiveDb')
    .addItem('Upload/Outre NEW 데이터 비우기', 'clearUploadSheetsData')
    .addSeparator()
    .addSubMenu(ui.createMenu('📄 인보이스')
      .addItem('📁 폴더 설정', 'setInvoiceFolder')
      .addItem('📄 파싱 시작', 'startParsing')
      .addItem('✅ 확정', 'confirmParsing')
      .addItem('❌ 취소', 'cancelParsing')
      .addSeparator()
      .addItem('🔍 Document AI 분석 저장', 'saveDocumentAIResponseToFiles'))
    .addSeparator()
    .addItem('🔌 연결 테스트', 'testConnectionUI')
    .addItem('🔍 검색 테스트', 'testSearch')
    .addItem('📊 시트 구조 확인', 'checkSheetStructure')
    .addSeparator()
    .addItem('🎨 컬러 정렬 테스트', 'testColorSorting')
    .addItem('📏 길이 정렬 테스트', 'testInchSorting')
    .addToUi();
  // Attach additional menus from other scripts (e.g. OUTRE UPC menu)
  if (typeof addOutreMenu === 'function') {
    addOutreMenu();
  }
}

/**
 * 초기 설정: 제품 DB 및 주문 시트 생성
 */
function initializeSheets() {
  try {
    debugLog('initializeSheets 시작');
    
    var ss = getSpreadsheet();
    
    var requiredSheets = [
      {
        name: 'Outre Active DB',
        headers: ['ITEM GROUP', 'ITEM NUMBER', 'ITEM NAME', 'COLOR', 'BARCODE']
      },
      {
        name: 'SNG Active DB',
        headers: ['Class', 'Old Item', 'Old Item Code', 'Item', 'Item Code', 'Color', 'Description', 'Barcode']
      },
      {
        name: 'ORDER_OUTRE',
        headers: ['ITEM NAME', 'COLOR', '수량']
      },
      {
        name: 'ORDER_SNG',
        headers: ['ITEM NAME', 'COLOR', '수량']
      }
    ];
    
    var results = [];
    
    for (var i = 0; i < requiredSheets.length; i++) {
      var sheetConfig = requiredSheets[i];
      var sheet = ss.getSheetByName(sheetConfig.name);
      
      if (sheet) {
        results.push('✓ ' + sheetConfig.name + ' - 이미 존재함');
      } else {
        sheet = ss.insertSheet(sheetConfig.name);
        sheet.appendRow(sheetConfig.headers);
        
        var headerRange = sheet.getRange(1, 1, 1, sheetConfig.headers.length);
        
        if (sheetConfig.name.indexOf('OUTRE') > -1) {
          headerRange.setBackground('#FF8C00');
        } else if (sheetConfig.name.indexOf('SNG') > -1) {
          headerRange.setBackground('#4169E1');
        } else {
          headerRange.setBackground('#666666');
        }
        
        headerRange.setFontColor('white');
        headerRange.setFontWeight('bold');
        headerRange.setHorizontalAlignment('center');
        
        for (var j = 1; j <= sheetConfig.headers.length; j++) {
          sheet.autoResizeColumn(j);
        }
        
        results.push('✅ ' + sheetConfig.name + ' - 생성 완료');
      }
    }
    
    debugLog('initializeSheets 완료', results);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '초기화 완료',
      results.join('\n'),
      ui.ButtonSet.OK
    );
    
    return {
      success: true,
      message: '초기화 완료',
      results: results
    };
    
  } catch (error) {
    debugLog('initializeSheets 오류', { error: error.toString() });
    logError('initializeSheets', error.toString());
    
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '오류 발생',
      '초기화 중 오류가 발생했습니다:\n' + error.toString(),
      ui.ButtonSet.OK
    );
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 인보이스 시트 초기화
 */
function initializeInvoiceSheets() {
  try {
    debugLog('initializeInvoiceSheets 시작');
    
    var ss = getSpreadsheet();
    var results = [];
    
    var sheets = [
      CONFIG.INVOICE.PARSING_SHEETS.SNG,
      CONFIG.INVOICE.PARSING_SHEETS.OUTRE,
      CONFIG.INVOICE.PARSING_SHEET,
      CONFIG.INVOICE.SNG_SHEET,
      CONFIG.INVOICE.OUTRE_SHEET
    ];
    
    for (var i = 0; i < sheets.length; i++) {
      var sheetName = sheets[i];
      var sheet = ss.getSheetByName(sheetName);
      
      if (sheet) {
        results.push('✓ ' + sheetName + ' - 이미 존재함');
      } else {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(CONFIG.INVOICE.HEADERS);
        
        var headerRange = sheet.getRange(1, 1, 1, CONFIG.INVOICE.HEADERS.length);
        
        if (sheetName === CONFIG.INVOICE.PARSING_SHEET ||
            sheetName === CONFIG.INVOICE.PARSING_SHEETS.SNG ||
            sheetName === CONFIG.INVOICE.PARSING_SHEETS.OUTRE) {
          headerRange.setBackground('#9370DB');
        } else if (sheetName.indexOf('SNG') > -1) {
          headerRange.setBackground('#4169E1');
        } else {
          headerRange.setBackground('#FF8C00');
        }
        
        headerRange.setFontColor('white');
        headerRange.setFontWeight('bold');
        headerRange.setHorizontalAlignment('center');
        
        for (var j = 1; j <= CONFIG.INVOICE.HEADERS.length; j++) {
          sheet.autoResizeColumn(j);
        }
        
        results.push('✅ ' + sheetName + ' - 생성 완료');
      }
    }
    
    debugLog('initializeInvoiceSheets 완료', results);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '인보이스 시트 초기화 완료',
      results.join('\n'),
      ui.ButtonSet.OK
    );
    
    return {
      success: true,
      results: results
    };
    
  } catch (error) {
    debugLog('initializeInvoiceSheets 오류', { error: error.toString() });
    logError('initializeInvoiceSheets', error.toString());
    
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '오류 발생',
      '초기화 중 오류가 발생했습니다:\n' + error.toString(),
      ui.ButtonSet.OK
    );
    
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 컬러 정렬 테스트 함수
 */
function testColorSorting() {
  var testColors = [
    '1', '1B', '2', '4', '27', '30', '530', '613', '130', '350', '33',
    '7', '99', '144',
    'T1', 'T1B', 'T27', 'T30', 'T99', 'T1B30', 'T2730',
    'P1', 'P27', 'P30', 'P99',
    'M1', 'M4', 'M27', 'M427', 'M2730613',
    'OT27', 'OT30',
    'DR27', 'FS30',
    'SUNFLOWER', 'UNICORN', 'AUBURN',
    '950+530+350+130S'
  ];
  
  var testProducts = [];
  for (var i = 0; i < testColors.length; i++) {
    testProducts.push({
      itemName: 'TEST PRODUCT 18"',
      color: testColors[i],
      quantity: 1
    });
  }
  
  var sorted = sortProducts(testProducts);
  
  debugLog('=== 컬러 정렬 테스트 결과 ===');
  for (var i = 0; i < sorted.length; i++) {
    debugLog((i + 1) + '. ' + sorted[i].color);
  }
  
  return sorted.map(function(p) { return p.color; });
}

/**
 * 길이별 정렬 테스트 함수
 */
function testInchSorting() {
  var testProducts = [
    { itemName: 'PRODUCT 20"', color: '1', quantity: 1 },
    { itemName: 'PRODUCT 10"', color: '1', quantity: 1 },
    { itemName: 'PRODUCT 14"', color: '1', quantity: 1 },
    { itemName: 'PRODUCT', color: '1', quantity: 1 },
    { itemName: 'PRODUCT 18"', color: '1', quantity: 1 },
    { itemName: 'PRODUCT 12"', color: '1', quantity: 1 }
  ];
  
  var sorted = sortProducts(testProducts);
  
  debugLog('=== 길이별 정렬 테스트 결과 ===');
  for (var i = 0; i < sorted.length; i++) {
    debugLog((i + 1) + '. ' + sorted[i].itemName);
  }
  
  return sorted.map(function(p) { return p.itemName; });
}
