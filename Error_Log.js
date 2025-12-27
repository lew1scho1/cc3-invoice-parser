// ============================================================================
// ERROR_LOG.GS - 에러 로그 시스템
// ============================================================================

/**
 * 에러 로그 시트 이름
 */
var ERROR_LOG_SHEET = 'ERROR_LOG';

/**
 * 에러 로그 시트 초기화
 */
function initializeErrorLogSheet() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName(ERROR_LOG_SHEET);
    
    if (!sheet) {
      sheet = ss.insertSheet(ERROR_LOG_SHEET);
      
      var headers = [
        'Timestamp',
        'Function',
        'Error Message',
        'Error Details',
        'User',
        'Additional Data'
      ];
      
      sheet.appendRow(headers);
      
      // 헤더 서식
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#DC143C');
      headerRange.setFontColor('white');
      headerRange.setFontWeight('bold');
      headerRange.setHorizontalAlignment('center');
      
      // 열 너비 설정
      sheet.setColumnWidth(1, 150); // Timestamp
      sheet.setColumnWidth(2, 200); // Function
      sheet.setColumnWidth(3, 300); // Error Message
      sheet.setColumnWidth(4, 400); // Error Details
      sheet.setColumnWidth(5, 150); // User
      sheet.setColumnWidth(6, 300); // Additional Data
      
      // 시트를 맨 뒤로 이동
      sheet.activate();
      ss.moveActiveSheet(ss.getNumSheets());
      
      debugLog('ERROR_LOG 시트 생성 완료');
      
      return {
        success: true,
        message: '✅ ERROR_LOG 시트가 생성되었습니다.'
      };
    } else {
      return {
        success: true,
        message: '✓ ERROR_LOG 시트가 이미 존재합니다.'
      };
    }
    
  } catch (error) {
    Logger.log('initializeErrorLogSheet 오류: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 에러 로그 저장 (개선된 버전)
 */
function logError(functionName, errorMessage, additionalData) {
  try {
    var ss = getSpreadsheet();
    var errorSheet = ss.getSheetByName(ERROR_LOG_SHEET);
    
    // 에러 시트가 없으면 생성
    if (!errorSheet) {
      initializeErrorLogSheet();
      errorSheet = ss.getSheetByName(ERROR_LOG_SHEET);
    }
    
    // 현재 시간
    var now = new Date();
    var timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    
    // 사용자 정보
    var user = Session.getActiveUser().getEmail() || 'Unknown';
    
    // 에러 상세 정보 (스택 트레이스 포함)
    var errorDetails = '';
    if (errorMessage && errorMessage.stack) {
      errorDetails = errorMessage.stack;
    } else {
      errorDetails = String(errorMessage);
    }
    
    // 추가 데이터
    var dataStr = '';
    if (additionalData) {
      try {
        dataStr = JSON.stringify(additionalData, null, 2);
      } catch (e) {
        dataStr = String(additionalData);
      }
    }
    
    // 에러 메시지 (간단한 버전)
    var simpleMessage = String(errorMessage);
    if (errorMessage && errorMessage.message) {
      simpleMessage = errorMessage.message;
    }
    
    // 로그 저장
    errorSheet.appendRow([
      timestamp,
      functionName,
      simpleMessage,
      errorDetails,
      user,
      dataStr
    ]);
    
    // 최신 로그가 위로 오도록 정렬 (선택사항)
    // var range = errorSheet.getRange(2, 1, errorSheet.getLastRow() - 1, 6);
    // range.sort({column: 1, ascending: false});
    
    debugLog('에러 로그 저장 완료', { 
      function: functionName, 
      message: simpleMessage 
    });
    
  } catch (logError) {
    // 에러 로그 저장 중 오류가 발생해도 메인 기능은 계속 진행
    Logger.log('에러 로그 저장 실패: ' + logError.toString());
    console.error('에러 로그 저장 실패:', logError);
  }
}

/**
 * 에러 로그 조회 (최근 N개)
 */
function getRecentErrors(count) {
  try {
    count = count || 10;
    
    var ss = getSpreadsheet();
    var errorSheet = ss.getSheetByName(ERROR_LOG_SHEET);
    
    if (!errorSheet) {
      return {
        success: false,
        message: 'ERROR_LOG 시트가 없습니다.'
      };
    }
    
    var data = errorSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        errors: [],
        message: '에러 로그가 없습니다.'
      };
    }
    
    // 헤더 제외하고 최근 N개 가져오기
    var errors = data.slice(1, Math.min(count + 1, data.length));
    
    return {
      success: true,
      errors: errors
    };
    
  } catch (error) {
    Logger.log('getRecentErrors 오류: ' + error.toString());
    return {
      success: false,
      error: error.toString()
    };
  }
}

/**
 * 에러 로그 비우기
 */
function clearErrorLog() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var response = ui.alert(
      '에러 로그 삭제',
      'ERROR_LOG 시트의 모든 에러 로그를 삭제하시겠습니까?',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    var ss = getSpreadsheet();
    var errorSheet = ss.getSheetByName(ERROR_LOG_SHEET);
    
    if (!errorSheet) {
      ui.alert('ERROR_LOG 시트가 없습니다.');
      return;
    }
    
    // 헤더 제외하고 모든 데이터 삭제
    var lastRow = errorSheet.getLastRow();
    if (lastRow > 1) {
      errorSheet.deleteRows(2, lastRow - 1);
    }
    
    ui.alert('에러 로그가 삭제되었습니다.');
    
    debugLog('에러 로그 삭제 완료');
    
  } catch (error) {
    Logger.log('clearErrorLog 오류: ' + error.toString());
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('오류', '에러 로그 삭제 중 오류가 발생했습니다:\n' + error.toString(), ui.ButtonSet.OK);
  }
}

/**
 * 에러 로그 UI 표시
 */
function showErrorLog() {
  try {
    var result = getRecentErrors(20);
    
    if (!result.success) {
      var ui = SpreadsheetApp.getUi();
      ui.alert('오류', result.error || result.message, ui.ButtonSet.OK);
      return;
    }
    
    if (result.errors.length === 0) {
      var ui = SpreadsheetApp.getUi();
      ui.alert('에러 로그', '에러 로그가 없습니다.', ui.ButtonSet.OK);
      return;
    }
    
    // 에러 로그를 텍스트로 포맷팅
    var logText = '=== 최근 에러 로그 (최대 20개) ===\n\n';
    
    for (var i = 0; i < result.errors.length; i++) {
      var error = result.errors[i];
      logText += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n';
      logText += '시간: ' + error[0] + '\n';
      logText += '함수: ' + error[1] + '\n';
      logText += '에러: ' + error[2] + '\n';
      logText += '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n';
    }
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('에러 로그', logText, ui.ButtonSet.OK);
    
  } catch (error) {
    Logger.log('showErrorLog 오류: ' + error.toString());
  }
}

/**
 * 에러 로그 내보내기 (텍스트 파일)
 */
function exportErrorLog() {
  try {
    var ss = getSpreadsheet();
    var errorSheet = ss.getSheetByName(ERROR_LOG_SHEET);
    
    if (!errorSheet) {
      var ui = SpreadsheetApp.getUi();
      ui.alert('ERROR_LOG 시트가 없습니다.');
      return;
    }
    
    var data = errorSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      var ui = SpreadsheetApp.getUi();
      ui.alert('에러 로그가 없습니다.');
      return;
    }
    
    // CSV 형식으로 변환
    var csv = '';
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var csvRow = row.map(function(cell) {
        var cellStr = String(cell).replace(/"/g, '""');
        return '"' + cellStr + '"';
      }).join(',');
      csv += csvRow + '\n';
    }
    
    // 파일 생성
    var fileName = 'ERROR_LOG_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '.csv';
    var blob = Utilities.newBlob(csv, 'text/csv', fileName);
    
    // Google Drive에 저장
    var file = DriveApp.createFile(blob);
    
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '내보내기 완료',
      '에러 로그가 Google Drive에 저장되었습니다:\n' + fileName + '\n\n' +
      '파일 URL:\n' + file.getUrl(),
      ui.ButtonSet.OK
    );
    
    debugLog('에러 로그 내보내기 완료', { fileName: fileName });
    
  } catch (error) {
    Logger.log('exportErrorLog 오류: ' + error.toString());
    
    var ui = SpreadsheetApp.getUi();
    ui.alert('오류', '내보내기 중 오류가 발생했습니다:\n' + error.toString(), ui.ButtonSet.OK);
  }
}