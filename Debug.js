// ============================================================================
// DEBUG.GS - 디버그 및 로깅 함수
// ============================================================================

/**
 * 디버그 로그 (DEBUG 모드일 때만 출력)
 */
function debugLog(message, data) {
  if (!CONFIG.DEBUG) return;
  
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss.SSS');
  var logMessage = '[' + timestamp + '] ' + message;
  
  if (data !== undefined) {
    logMessage += ' | Data: ' + JSON.stringify(data);
  }
  
  Logger.log(logMessage);
  console.log(logMessage);
}

/**
 * 에러 로그 저장 (ERROR_LOG 시트에 기록)
 */
function logError(functionName, errorMessage, additionalData) {
  try {
    debugLog('ERROR in ' + functionName, { error: errorMessage, data: additionalData });
    
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var errorSheet = ss.getSheetByName('ERROR_LOG');
    
    // 에러 시트가 없으면 생성
    if (!errorSheet) {
      errorSheet = ss.insertSheet('ERROR_LOG');
      errorSheet.appendRow(['Timestamp', 'Function', 'Error', 'Data', 'User']);
      
      // 헤더 서식
      var headerRange = errorSheet.getRange(1, 1, 1, 5);
      headerRange.setBackground('#DC143C');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
    }
    
    // 에러 정보 저장
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var user = Session.getActiveUser().getEmail() || 'Unknown';
    var dataStr = additionalData ? JSON.stringify(additionalData) : '';
    
    errorSheet.appendRow([
      timestamp,
      functionName,
      errorMessage,
      dataStr,
      user
    ]);
    
  } catch (logError) {
    Logger.log('에러 로그 저장 실패: ' + logError.toString());
  }
}

/**
 * 함수 실행 시간 측정
 */
function measureTime(functionName, func) {
  var startTime = new Date().getTime();
  
  try {
    var result = func();
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    
    debugLog(functionName + ' 실행 완료', { duration: duration + 'ms' });
    
    return result;
  } catch (error) {
    var endTime = new Date().getTime();
    var duration = endTime - startTime;
    
    debugLog(functionName + ' 실행 실패', { duration: duration + 'ms', error: error.toString() });
    throw error;
  }
}

/**
 * 테스트용 검색 함수 (직접 실행 가능)
 */
function testSearch() {
  debugLog('=== 검색 테스트 시작 ===');
  
  // 바코드 검색 테스트
  var barcodeResult = searchProduct('827298817871');
  debugLog('바코드 검색 결과', barcodeResult);
  
  // 텍스트 검색 테스트
  var textResult = searchProduct('OG BULK');
  debugLog('텍스트 검색 결과', textResult);
  
  debugLog('=== 검색 테스트 완료 ===');
}

/**
 * DB 시트 구조 확인
 */
function checkSheetStructure() {
  debugLog('=== 시트 구조 확인 시작 ===');
  
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  // OUTRE 확인
  var outreSheet = ss.getSheetByName('DB_OUTRE');
  if (outreSheet) {
    var outreData = outreSheet.getDataRange().getValues();
    debugLog('DB_OUTRE 시트', {
      rows: outreData.length,
      headers: outreData[0]
    });
  } else {
    debugLog('DB_OUTRE 시트 없음');
  }
  
  // SNG 확인
  var sngSheet = ss.getSheetByName('DB_SNG');
  if (sngSheet) {
    var sngData = sngSheet.getDataRange().getValues();
    debugLog('DB_SNG 시트', {
      rows: sngData.length,
      headers: sngData[0]
    });
  } else {
    debugLog('DB_SNG 시트 없음');
  }
  
  debugLog('=== 시트 구조 확인 완료 ===');
}

/**
 * 연결 테스트
 */
function testConnection() {
  try {
    debugLog('연결 테스트 시작');
    
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var name = ss.getName();
    
    debugLog('연결 성공', { spreadsheet: name });
    
    return {
      success: true,
      message: '✅ 연결 성공!',
      spreadsheetName: name
    };
  } catch (error) {
    debugLog('연결 실패', { error: error.toString() });
    
    return {
      success: false,
      error: '❌ 연결 실패: ' + error.toString()
    };
  }
}

/**
 * UI용 연결 테스트
 */
function testConnectionUI() {
  var result = testConnection();
  var ui = SpreadsheetApp.getUi();

  if (result.success) {
    ui.alert(
      '연결 성공',
      '스프레드시트: ' + result.spreadsheetName,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(
      '연결 실패',
      result.error,
      ui.ButtonSet.OK
    );
  }
}

// ============================================================================
// 디버그 로그를 구글 시트에 출력
// ============================================================================

var DEBUG_LOGS = [];

/**
 * 디버그 로그 수집 (구글 시트 출력용)
 */
function collectDebugLog(message, data) {
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss.SSS');
  var logEntry = {
    timestamp: timestamp,
    message: message,
    data: data !== undefined ? JSON.stringify(data) : ''
  };
  DEBUG_LOGS.push(logEntry);
}

/**
 * 디버그 로그를 DEBUG_OUTPUT 시트에 출력 (배치 쓰기)
 * CRITICAL: appendRow() 대신 setValues()로 한 번에 쓰기 (100배 이상 빠름)
 */
function outputDebugLogsToSheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName('DEBUG_OUTPUT');

    // DEBUG_OUTPUT 시트가 없으면 생성
    if (!debugSheet) {
      debugSheet = ss.insertSheet('DEBUG_OUTPUT');

      // 헤더 설정
      debugSheet.appendRow(['Timestamp', 'Message', 'Data']);

      var headerRange = debugSheet.getRange(1, 1, 1, 3);
      headerRange.setBackground('#4285F4');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
    } else {
      // 기존 내용 삭제 (헤더 제외)
      var lastRow = debugSheet.getLastRow();
      if (lastRow > 1) {
        debugSheet.deleteRows(2, lastRow - 1);
      }
    }

    // CRITICAL: 배치 쓰기로 로그 출력 (100배 이상 빠름)
    if (DEBUG_LOGS.length > 0) {
      var rows = [];
      for (var i = 0; i < DEBUG_LOGS.length; i++) {
        var log = DEBUG_LOGS[i];
        rows.push([log.timestamp, log.message, log.data]);
      }

      // 한 번에 쓰기 (Single API call)
      var targetRange = debugSheet.getRange(2, 1, rows.length, 3);
      targetRange.setValues(rows);
    }

    // 컬럼 너비 자동 조정
    debugSheet.autoResizeColumn(1);
    debugSheet.autoResizeColumn(2);
    debugSheet.setColumnWidth(3, 600);

    // DEBUG_OUTPUT 시트로 이동
    ss.setActiveSheet(debugSheet);

  } catch (error) {
    Logger.log('디버그 로그 출력 실패: ' + error.toString());
  }
}

/**
 * Invoice Amount 파싱 디버그 (구글 시트 출력)
 */
function debugInvoiceAmountParsing() {
  DEBUG_LOGS = []; // 로그 초기화

  try {
    collectDebugLog('=== Invoice Amount 파싱 디버그 시작 ===');

    // 1. 폴더에서 파일 가져오기
    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      collectDebugLog('❌ 오류', '폴더가 설정되지 않았습니다.');
      outputDebugLogsToSheet();
      return;
    }

    collectDebugLog('폴더 ID', folderId);

    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();

    var fileList = [];
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      var mimeType = file.getMimeType();

      if (mimeType === MimeType.PDF ||
          mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        fileList.push({
          id: file.getId(),
          name: name,
          mimeType: mimeType
        });
      }
    }

    collectDebugLog('파일 목록', { count: fileList.length });

    if (fileList.length === 0) {
      collectDebugLog('❌ 오류', '파싱 가능한 파일이 없습니다.');
      outputDebugLogsToSheet();
      return;
    }

    // 첫 번째 파일 선택 (또는 특정 파일명 검색 가능)
    var selectedFile = fileList[0];
    collectDebugLog('선택된 파일', selectedFile);

    // 2. 파일에서 텍스트 추출
    var file = DriveApp.getFileById(selectedFile.id);
    var filename = file.getName();
    var mimeType = file.getMimeType();

    collectDebugLog('파일 정보', { filename: filename, mimeType: mimeType });

    var text = '';
    if (mimeType === MimeType.PDF) {
      text = extractTextFromPdf(file);
    } else {
      text = extractTextFromDocx(file.getBlob());
    }

    collectDebugLog('텍스트 추출 완료', { length: text.length });

    // 3. Vendor 감지
    var vendor = detectVendor(text, filename);
    collectDebugLog('=== Vendor Detection ===');
    collectDebugLog('Filename', filename);
    collectDebugLog('Detected vendor', vendor);

    if (vendor === 'UNKNOWN') {
      collectDebugLog('❌ 오류', '인보이스 회사를 감지할 수 없습니다.');
      outputDebugLogsToSheet();
      return;
    }

    // 4. parseHeaderInfo 진입 확인
    var data = {
      vendor: vendor,
      filename: filename,
      invoiceNo: '',
      invoiceDate: '',
      totalAmount: 0,
      subtotal: 0,
      discount: 0,
      shipping: 0,
      tax: 0,
      lineItems: []
    };

    var lines = text.split('\n');
    var fullText = lines.join('\n');

    collectDebugLog('=== parseHeaderInfo 시작 ===');
    collectDebugLog('data.vendor', data.vendor);
    collectDebugLog('fullText length', fullText.length);

    // 5. Vendor 체크 확인
    collectDebugLog('=== Total Amount 파싱 시작 ===');
    collectDebugLog('Checking vendor: data.vendor', '"' + data.vendor + '"');
    collectDebugLog('Is SNG?', (data.vendor === 'SNG'));

    // 6. Invoice Amount 파싱 시도
    if (data.vendor === 'SNG') {
      collectDebugLog('✅ SNG 조건 진입');

      var allMatches = [];
      var patterns = [
        /INVOICE\s+AMOUNT[:\s]*\$?\s*([\d,\.]+)/gi,
        /TOTAL\s+INVOICE\s+AMOUNT[:\s]*\$?\s*([\d,\.]+)/gi,
        /AMOUNT\s+DUE[:\s]*\$?\s*([\d,\.]+)/gi,
        /INVOICE\s+AMT[:\s]*\$?\s*([\d,\.]+)/gi,
        /TOTAL\s+AMOUNT\s+DUE[:\s]*\$?\s*([\d,\.]+)/gi,
        /AMOUNT[:\s]*\$?\s*([\d,\.]+)/gi
      ];

      collectDebugLog('=== Invoice Amount 파싱 시작 ===');
      collectDebugLog('텍스트 길이', fullText.length + ' 문자');

      for (var p = 0; p < patterns.length; p++) {
        var match;
        var patternMatches = 0;

        while ((match = patterns[p].exec(fullText)) !== null) {
          var amount = parseAmount(match[1]);
          patternMatches++;

          collectDebugLog('패턴 ' + p + ' 매치 #' + patternMatches, {
            match: match[0],
            amount: amount,
            position: match.index
          });

          if (amount >= 100) {
            allMatches.push({
              pattern: p,
              amount: amount,
              fullMatch: match[0],
              index: match.index
            });
          }
        }

        if (patternMatches > 0) {
          collectDebugLog('패턴 ' + p + ' 총 매치 수', patternMatches);
        }
      }

      collectDebugLog('=== 유효한 매치 (>=$100) ===', { count: allMatches.length });

      if (allMatches.length > 0) {
        allMatches.sort(function(a, b) { return b.index - a.index; });
        var bestMatch = allMatches[0];

        collectDebugLog('✅ Invoice Amount 파싱 성공', {
          amount: bestMatch.amount,
          pattern: bestMatch.fullMatch,
          position: bestMatch.index
        });
      } else {
        collectDebugLog('❌ Invoice Amount 파싱 실패', '유효한 금액을 찾을 수 없음');

        // 텍스트 샘플 출력
        var sample = fullText.substring(fullText.length - 500);
        collectDebugLog('텍스트 샘플 (마지막 500자)', sample);
      }

    } else {
      collectDebugLog('❌ SNG 조건 진입 실패', 'vendor가 SNG가 아닙니다: ' + vendor);
    }

    collectDebugLog('=== 디버그 완료 ===');

  } catch (error) {
    collectDebugLog('❌ 에러 발생', error.toString());
  }

  // 로그를 시트에 출력
  outputDebugLogsToSheet();
}

/**
 * Color 파싱 디버그 (구글 시트 출력)
 */
function debugColorParsing() {
  DEBUG_LOGS = []; // 로그 초기화

  try {
    collectDebugLog('=== Color 파싱 디버그 시작 ===');

    // 1. 폴더에서 파일 가져오기
    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      collectDebugLog('❌ 오류', '폴더가 설정되지 않았습니다.');
      outputDebugLogsToSheet();
      return;
    }

    collectDebugLog('폴더 ID', folderId);

    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();

    var fileList = [];
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      var mimeType = file.getMimeType();

      if (mimeType === MimeType.PDF ||
          mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        fileList.push({
          id: file.getId(),
          name: name,
          mimeType: mimeType
        });
      }
    }

    collectDebugLog('파일 목록', { count: fileList.length });

    if (fileList.length === 0) {
      collectDebugLog('❌ 오류', '파싱 가능한 파일이 없습니다.');
      outputDebugLogsToSheet();
      return;
    }

    // 첫 번째 파일 선택
    var selectedFile = fileList[0];
    collectDebugLog('선택된 파일', selectedFile);

    // 2. 파일에서 텍스트 추출
    var file = DriveApp.getFileById(selectedFile.id);
    var filename = file.getName();
    var mimeType = file.getMimeType();

    collectDebugLog('파일 정보', { filename: filename, mimeType: mimeType });

    var text = '';
    if (mimeType === MimeType.PDF) {
      text = extractTextFromPdf(file);
    } else {
      text = extractTextFromDocx(file.getBlob());
    }

    collectDebugLog('텍스트 추출 완료', { length: text.length });

    // 3. Vendor 감지
    var vendor = detectVendor(text, filename);
    collectDebugLog('Vendor 감지', vendor);

    if (vendor === 'UNKNOWN') {
      collectDebugLog('❌ 오류', '인보이스 회사를 감지할 수 없습니다.');
      outputDebugLogsToSheet();
      return;
    }

    // 4. 라인 파싱
    var lines = text.split('\n');
    collectDebugLog('총 라인 수', lines.length);

    // 5. 첫 번째 아이템 라인 찾기 및 컬러 라인 검색
    collectDebugLog('=== 아이템 라인 검색 시작 ===');

    var itemCount = 0;
    var maxItemsToDebug = 5; // 처음 5개 아이템만 디버깅

    for (var i = 0; i < lines.length && itemCount < maxItemsToDebug; i++) {
      var line = lines[i].trim();

      // SNG 아이템 라인 패턴: 문자+숫자 시작 + 탭
      var isSngItem = (vendor === 'SNG' && line.match(/^[A-Z]\d+\t/));
      // OUTRE 아이템 라인 패턴: 숫자 시작 + 탭 + 문자
      var isOutreItem = (vendor === 'OUTRE' && line.match(/^\d+\t[A-Z]/));

      if (!isSngItem && !isOutreItem) {
        continue;
      }

      itemCount++;

      var parts = line.split('\t');
      var itemId = '';
      var description = '';

      if (vendor === 'SNG') {
        itemId = parts[0] ? parts[0].trim() : '';
        description = parts[1] ? parts[1].trim() : '';
      } else if (vendor === 'OUTRE') {
        var descParts = [];
        for (var j = 1; j < parts.length - 3; j++) {
          if (parts[j] && parts[j].trim()) {
            descParts.push(parts[j].trim());
          }
        }
        description = descParts.join(' ');
      }

      collectDebugLog('=== 아이템 #' + itemCount + ' ===', {
        lineNo: i,
        itemId: itemId,
        description: description
      });

      // 6. 컬러 라인 검색
      var colorLines = [];
      var searchLog = {
        itemId: itemId,
        searchRange: Math.min(i + 50, lines.length) - (i + 1),
        linesChecked: 0,
        linesFiltered: [],
        linesCollected: []
      };

      for (var j = i + 1; j < Math.min(i + 50, lines.length); j++) {
        var nextLine = lines[j].trim();
        searchLog.linesChecked++;

        // 다음 아이템 라인을 만나면 중단
        if (vendor === 'SNG' && nextLine.match(/^[A-Z]\d+\t/)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: '다음 아이템 라인',
            text: nextLine.substring(0, 80)
          });
          break;
        }
        if (vendor === 'OUTRE' && nextLine.match(/^\d+\t[A-Z]/)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: '다음 아이템 라인',
            text: nextLine.substring(0, 80)
          });
          break;
        }

        if (!nextLine) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: '빈 라인',
            text: ''
          });
          continue;
        }

        // 페이지 헤더/푸터 패턴 무시
        if (nextLine.match(/^Page \d+/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'Page 번호',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/SHAKE-N-GO/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'SHAKE-N-GO',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^INVOICE/i) && nextLine.length < 50) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'INVOICE 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^[\-=]+$/)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: '구분선',
            text: nextLine
          });
          continue;
        }

        // 헤더 패턴 필터링
        if (nextLine.match(/^\s*QTY\s+.*\s+ITEM/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'QTY...ITEM 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*ORDERED\s+SHIPPED/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'ORDERED SHIPPED 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*ITEM\s+NUMBER/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'ITEM NUMBER 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*DESCRIPTION/i) && nextLine.length < 50) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'DESCRIPTION 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*UNIT\s+PRICE/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'UNIT PRICE 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*EXT\.?\s+PRICE/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'EXT PRICE 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*ORDER\s+NUMBER/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'ORDER NUMBER 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*CUSTOMER/i) && nextLine.length < 50) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'CUSTOMER 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*SHIP\s+TO/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'SHIP TO 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*SOLD\s+TO/i)) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'SOLD TO 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*DATE/i) && nextLine.length < 30) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'DATE 헤더',
            text: nextLine
          });
          continue;
        }
        if (nextLine.match(/^\s*TERMS/i) && nextLine.length < 30) {
          searchLog.linesFiltered.push({
            lineNo: j,
            reason: 'TERMS 헤더',
            text: nextLine
          });
          continue;
        }

        // 언더스코어가 있는 컬러 라인
        if (nextLine.indexOf('_') > -1) {
          colorLines.push(nextLine);
          searchLog.linesCollected.push({
            lineNo: j,
            type: '언더스코어',
            text: nextLine
          });
          continue;
        }

        // 컬러 패턴 매치
        if (nextLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/)) {
          colorLines.push(nextLine);
          searchLog.linesCollected.push({
            lineNo: j,
            type: '컬러 패턴',
            text: nextLine
          });
        } else {
          // 매치되지 않은 라인 기록
          if (nextLine.length > 0 && nextLine.length < 100) {
            searchLog.linesFiltered.push({
              lineNo: j,
              reason: '패턴 불일치',
              text: nextLine
            });
          }
        }
      }

      // 7. 검색 결과 로깅
      collectDebugLog('컬러 라인 검색 완료', {
        searchRange: searchLog.searchRange + ' 라인',
        linesChecked: searchLog.linesChecked,
        linesFiltered: searchLog.linesFiltered.length,
        linesCollected: searchLog.linesCollected.length
      });

      if (searchLog.linesCollected.length > 0) {
        collectDebugLog('✅ 수집된 컬러 라인 (' + searchLog.linesCollected.length + '개)');
        for (var logIdx = 0; logIdx < searchLog.linesCollected.length; logIdx++) {
          var collected = searchLog.linesCollected[logIdx];
          collectDebugLog('  [라인 ' + collected.lineNo + '] ' + collected.type, collected.text);
        }
      } else {
        collectDebugLog('❌ 컬러 라인 없음');
      }

      if (searchLog.linesFiltered.length > 0) {
        collectDebugLog('필터링된 라인 (처음 10개)', searchLog.linesFiltered.length + '개 중');
        for (var logIdx = 0; logIdx < Math.min(10, searchLog.linesFiltered.length); logIdx++) {
          var filtered = searchLog.linesFiltered[logIdx];
          collectDebugLog('  [라인 ' + filtered.lineNo + '] ' + filtered.reason, filtered.text.substring(0, 80));
        }
      }

      // 8. 컬러 파싱 시도
      if (colorLines.length > 0) {
        collectDebugLog('=== 컬러 파싱 시작 ===');
        collectDebugLog('원본 컬러 라인', colorLines);

        // 전처리
        var fullText = colorLines.join(' ');
        collectDebugLog('합친 텍스트', fullText);

        fullText = fullText.replace(/_+/g, ' ');
        fullText = fullText.replace(/\s+/g, ' ').trim();
        collectDebugLog('전처리 후 텍스트', fullText);

        // 정규식 매칭
        var regex = /([A-Z0-9\-\/]+)\s*-\s*(\d+)\s*(?:\((\d+)\))?/gi;
        var colorData = [];
        var match;

        while ((match = regex.exec(fullText)) !== null) {
          var color = match[1].trim();
          var shipped = parseInt(match[2]) || 0;
          var backordered = match[3] ? parseInt(match[3]) : 0;

          collectDebugLog('컬러 매치 발견', {
            fullMatch: match[0],
            color: color,
            shipped: shipped,
            backordered: backordered
          });

          if (color && color.length > 0 && (shipped > 0 || backordered > 0)) {
            colorData.push({
              color: color,
              shipped: shipped,
              backordered: backordered
            });
          }
        }

        if (colorData.length > 0) {
          collectDebugLog('✅ 컬러 파싱 성공', { count: colorData.length, data: colorData });
        } else {
          collectDebugLog('❌ 컬러 파싱 실패', '정규식 패턴에 매치되는 데이터 없음');
          collectDebugLog('정규식 패턴', '/([A-Z0-9\\-\\/]+)\\s*-\\s*(\\d+)\\s*(?:\\((\\d+)\\))?/gi');
        }
      } else {
        collectDebugLog('❌ 컬러 라인 없음', '파싱 시도 안함');
      }

      collectDebugLog(''); // 빈 줄로 구분
    }

    collectDebugLog('=== 디버그 완료 ===', '총 ' + itemCount + '개 아이템 검사');

  } catch (error) {
    collectDebugLog('❌ 에러 발생', error.toString());
  }

  // 로그를 시트에 출력
  outputDebugLogsToSheet();
}
