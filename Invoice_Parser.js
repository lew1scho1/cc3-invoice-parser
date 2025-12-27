// ============================================================================
// INVOICE_PARSER.GS - 통합 인보이스 파서 v2
// ============================================================================

/**
 * 폴더 설정
 */
function setInvoiceFolder() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.prompt(
    '인보이스 폴더 설정',
    'Google Drive 폴더 URL을 입력하세요:\n(예: https://drive.google.com/drive/folders/[FOLDER_ID])',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var input = response.getResponseText().trim();
  
  // URL에서 폴더 ID 추출
  var folderId = extractFolderId(input);
  
  if (!folderId) {
    ui.alert('오류', '올바른 폴더 URL을 입력해주세요.', ui.ButtonSet.OK);
    return;
  }
  
  // 폴더 접근 확인
  try {
    var folder = DriveApp.getFolderById(folderId);
    var folderName = folder.getName();
    
    // 설정 저장
    PropertiesService.getDocumentProperties().setProperty(
      CONFIG.INVOICE.FOLDER_ID_PROPERTY, 
      folderId
    );
    
    ui.alert(
      '설정 완료',
      '폴더가 설정되었습니다:\n' + folderName,
      ui.ButtonSet.OK
    );
    
    debugLog('폴더 설정 완료', { folderId: folderId, name: folderName });
    
  } catch (error) {
    ui.alert(
      '오류',
      '폴더에 접근할 수 없습니다.\n권한을 확인해주세요.\n\n' + error.toString(),
      ui.ButtonSet.OK
    );
    
    debugLog('폴더 설정 실패', { error: error.toString() });
  }
}

/**
 * URL에서 폴더 ID 추출
 */
function extractFolderId(input) {
  if (!input) return null;
  
  // 이미 ID만 입력한 경우
  if (input.length > 20 && input.indexOf('/') === -1) {
    return input;
  }
  
  // URL에서 추출
  var match = input.match(/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

/**
 * 파싱 시작 (폴더의 파일 목록 표시)
 */
function startParsing() {
  var ui = SpreadsheetApp.getUi();
  
  // 폴더 ID 가져오기
  var folderId = PropertiesService.getDocumentProperties()
    .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);
  
  if (!folderId) {
    ui.alert(
      '폴더 미설정',
      '먼저 인보이스 폴더를 설정해주세요.\n\n메뉴: CC3 ORDER APP > 📄 인보이스 > 📁 폴더 설정',
      ui.ButtonSet.OK
    );
    return;
  }
  
  try {
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    
    var fileList = [];
    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType();
      
      // PDF 또는 DOCX만
      if (mimeType === MimeType.PDF || 
          mimeType === MimeType.MICROSOFT_WORD ||
          mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
        fileList.push({
          id: file.getId(),
          name: file.getName(),
          date: file.getDateCreated()
        });
      }
    }
    
    if (fileList.length === 0) {
      ui.alert(
        '파일 없음',
        '폴더에 PDF 또는 DOCX 파일이 없습니다.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // 날짜순 정렬 (최신순)
    fileList.sort(function(a, b) {
      return b.date - a.date;
    });
    
    // 파일 선택 UI
    var fileNames = fileList.map(function(f, i) {
      return (i + 1) + '. ' + f.name;
    }).join('\n');

    var response = ui.prompt(
      '파일 선택',
      '파싱할 파일 번호를 입력하세요.\n복수 선택 시 쉼표로 구분 (예: 1,3,5):\n\n' + fileNames,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    var input = response.getResponseText().trim();
    var selectedIndices = [];

    // 쉼표로 구분된 번호들 파싱
    var inputParts = input.split(',');
    for (var p = 0; p < inputParts.length; p++) {
      var num = parseInt(inputParts[p].trim());
      if (!isNaN(num) && num >= 1 && num <= fileList.length) {
        selectedIndices.push(num - 1);
      }
    }

    if (selectedIndices.length === 0) {
      ui.alert('오류', '올바른 번호를 입력해주세요.', ui.ButtonSet.OK);
      return;
    }

    // 중복 제거
    selectedIndices = selectedIndices.filter(function(value, index, self) {
      return self.indexOf(value) === index;
    });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var results = [];
    var successCount = 0;
    var failCount = 0;

    // 멀티 파일 모드: 첫 파일 전에 PARSING 탭 비우기
    if (selectedIndices.length > 1) {
      var parsingSheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);
      var existingData = parsingSheet.getDataRange().getValues();
      var hasExistingData = existingData.length > 1;

      if (hasExistingData) {
        var existingVendor = existingData[1][0];
        var existingLineCount = existingData.length - 1;

        var response = ui.alert(
          'PARSING 탭에 기존 데이터 있음',
          existingVendor + ' 인보이스 ' + existingLineCount + '개 라인이 있습니다.\n\n' +
          '예 = 삭제 후 진행\n' +
          '아니오 = DB로 이동 후 진행\n' +
          '취소 = 파싱 중단',
          ui.ButtonSet.YES_NO_CANCEL
        );

        if (response === ui.Button.YES) {
          ss.toast('기존 데이터 삭제 중...', '파싱 진행 중', -1);
          clearParsingSheet();
        } else if (response === ui.Button.NO) {
          ss.toast('기존 데이터 저장 중...', '파싱 진행 중', -1);
          var targetSheetName = existingVendor === 'SNG' ?
            CONFIG.INVOICE.SNG_SHEET :
            CONFIG.INVOICE.OUTRE_SHEET;
          moveDataToSheet(existingData, targetSheetName);
          clearParsingSheet();
        } else {
          return; // 취소
        }
      }
    }

    // 선택된 파일들 순차 파싱
    for (var idx = 0; idx < selectedIndices.length; idx++) {
      var selectedFile = fileList[selectedIndices[idx]];

      ss.toast('파일 ' + (idx + 1) + '/' + selectedIndices.length + ' 파싱 중: ' + selectedFile.name, '파싱 진행 중', -1);

      // 멀티 파일 모드 플래그 전달
      var result = parseAndSaveToParsing(selectedFile.id, selectedIndices.length > 1);

      if (result.success) {
        successCount++;
        results.push('✅ ' + selectedFile.name + '\n   ' + result.vendor + ', ' + result.lineCount + '개 라인');
      } else {
        failCount++;
        results.push('❌ ' + selectedFile.name + '\n   ' + result.error);
      }

      // 각 파일 파싱 후 잠시 대기 (API rate limit 방지)
      if (idx < selectedIndices.length - 1) {
        Utilities.sleep(1000);
      }
    }

    // 토스트 닫기
    ss.toast('', '', 1);

    // 결과 요약
    var summaryMessage = '파싱 완료!\n\n' +
                        '성공: ' + successCount + '개\n' +
                        '실패: ' + failCount + '개\n\n' +
                        results.join('\n\n') +
                        '\n\nPARSING 탭에서 확인 후 "✅ 확정" 버튼을 눌러주세요.';

    ui.alert('파싱 결과', summaryMessage, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert(
      '오류',
      '파일 목록을 가져오는 중 오류가 발생했습니다:\n' + error.toString(),
      ui.ButtonSet.OK
    );
    
    debugLog('startParsing 오류', { error: error.toString() });
  }
}

/**
 * 파싱 후 PARSING 탭에 저장
 * @param {string} fileId - 파일 ID
 * @param {boolean} isMultiFileMode - 멀티 파일 모드 여부 (true면 확인 창 스킵)
 */
function parseAndSaveToParsing(fileId, isMultiFileMode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  try {
    debugLog('parseAndSaveToParsing 시작', { fileId: fileId, isMultiFileMode: isMultiFileMode });

    // 0. PARSING 시트에 기존 데이터가 있는지 확인 (싱글 파일 모드만)
    if (!isMultiFileMode) {
      var parsingSheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);
      var existingData = parsingSheet.getDataRange().getValues();
      var hasExistingData = existingData.length > 1 && existingData[1][0]; // 헤더 제외하고 실제 데이터 있는지

      if (hasExistingData) {
        // 기존 데이터의 vendor 확인
        var existingVendor = existingData[1][0]; // 첫 번째 데이터 행의 VENDOR
        var existingLineCount = existingData.length - 1;

        var response = ui.alert(
          'PARSING 탭에 기존 데이터 있음',
          existingVendor + ' 인보이스 ' + existingLineCount + '개 라인이 있습니다.\n\n' +
          '예 = 삭제 후 새 파싱\n' +
          '아니오 = DB로 이동 후 새 파싱\n' +
          '취소 = 유지하고 아래 추가',
          ui.ButtonSet.YES_NO_CANCEL
        );

        if (response === ui.Button.YES) {
          // YES: 삭제 후 새 파싱
          ss.toast('기존 데이터 삭제 중...', '파싱 진행 중', -1);
          clearParsingSheet();
        } else if (response === ui.Button.NO) {
          // NO: DB로 이동 후 새 파싱
          ss.toast('기존 데이터 저장 중...', '파싱 진행 중', -1);
          var targetSheetName = existingVendor === 'SNG' ?
            CONFIG.INVOICE.SNG_SHEET :
            CONFIG.INVOICE.OUTRE_SHEET;

          moveDataToSheet(existingData, targetSheetName);

          debugLog('기존 데이터 이동 완료', {
            vendor: existingVendor,
            targetSheet: targetSheetName,
            rows: existingData.length - 1
          });

          clearParsingSheet();
          ss.toast(existingLineCount + '개 라인을 ' + targetSheetName + '로 이동했습니다.', '완료', 3);
        } else {
          // CANCEL: 그냥 놔두고 아래에 추가
          ss.toast('기존 데이터 유지, 아래에 추가...', '파싱 진행 중', -1);
          // clearParsingSheet() 호출 안 함
        }
      }
    }

    // 1. 파일 가져오기
    ss.toast('파일 정보 가져오는 중...', '파싱 진행 중', -1);
    var file = DriveApp.getFileById(fileId);
    var filename = file.getName();
    var mimeType = file.getMimeType();

    debugLog('파일 정보', { filename: filename, mimeType: mimeType });

    // 2. 파싱 (SNG는 Document AI, OUTRE는 기존 방식)
    var result;
    var text = '';
    var useDocumentAI = false;

    // 파일명으로 vendor 미리 감지 및 파싱 방식 결정
    var vendorHint = 'UNKNOWN';
    if (filename.indexOf('3000') === 0 || filename.match(/\d{10}/)) {
      vendorHint = 'SNG';
      // PDF는 Document AI, DOCX는 기존 방식
      useDocumentAI = (mimeType === MimeType.PDF);
    } else if (filename.indexOf('SINV') > -1) {
      vendorHint = 'OUTRE';
      // PDF는 Document AI, DOCX는 기존 방식
      useDocumentAI = (mimeType === MimeType.PDF);
    }

    if (useDocumentAI) {
      // Document AI 사용 (PDF 전용)
      ss.toast('Document AI로 인보이스 파싱 중 (' + vendorHint + ')...', '파싱 진행 중', -1);
      try {
        var aiResult = parseInvoiceWithDocumentAI(file);
        result = convertDocumentAIToInvoiceData(aiResult, filename);

        debugLog('Document AI 파싱 완료', {
          vendor: result.vendor,
          lineItems: result.lineItems.length
        });

      } catch (aiError) {
        // Document AI 실패 시 기존 방식으로 폴백
        debugLog('Document AI 실패, 기존 방식으로 폴백', { error: aiError.toString() });
        ss.toast('Document AI 실패, 기존 방식으로 파싱 중...', '파싱 진행 중', -1);

        if (mimeType === MimeType.PDF) {
          text = extractTextFromPdf(file);
        } else {
          text = extractTextFromDocx(file.getBlob());
        }

        if (!text || text.trim() === '') {
          throw new Error('파일에서 텍스트를 추출할 수 없습니다.');
        }

        debugLog('텍스트 추출 완료', { length: text.length });

        ss.toast('인보이스 데이터 파싱 중...', '파싱 진행 중', -1);
        result = parseInvoice(text, filename);
      }

    } else {
      // 기존 방식 사용 (DOCX 또는 PDF 폴백)
      ss.toast('텍스트 추출 중 (' + vendorHint + ')...', '파싱 진행 중', -1);

      if (mimeType === MimeType.PDF) {
        text = extractTextFromPdf(file);
      } else {
        text = extractTextFromDocx(file.getBlob());
      }

      if (!text || text.trim() === '') {
        throw new Error('파일에서 텍스트를 추출할 수 없습니다.');
      }

      debugLog('텍스트 추출 완료', { length: text.length });

      ss.toast('인보이스 데이터 파싱 중 (기존 방식)...', '파싱 진행 중', -1);
      result = parseInvoice(text, filename);
    }

    debugLog('파싱 완료', {
      vendor: result.vendor,
      lineItems: result.lineItems.length
    });

    // 4. PARSING 시트에 저장
    ss.toast(result.lineItems.length + '개 라인 저장 중...', '파싱 진행 중', -1);
    // clearParsingSheet()는 위에서 이미 처리됨
    var savedCount = saveToParsingSheet(result);

    // ExtPrice 합계 계산
    var extPriceSum = 0;
    for (var i = 0; i < result.lineItems.length; i++) {
      extPriceSum += result.lineItems[i].extPrice || 0;
    }
    extPriceSum = Number(extPriceSum.toFixed(2));

    // 검증 메시지
    var validationMsg = '';
    var priceDifference = Math.abs(result.totalAmount - extPriceSum);
    if (priceDifference > 1.0) {
      validationMsg = '\n\n⚠️ 경고: 금액 불일치\n' +
                      'Invoice Amount: $' + result.totalAmount.toFixed(2) + '\n' +
                      'ExtPrice 합계: $' + extPriceSum.toFixed(2) + '\n' +
                      '차이: $' + priceDifference.toFixed(2);
    } else {
      validationMsg = '\n\n✅ 금액 검증 통과\n' +
                      'Invoice Amount: $' + result.totalAmount.toFixed(2) + '\n' +
                      'ExtPrice 합계: $' + extPriceSum.toFixed(2);
    }

    return {
      success: true,
      message: '✅ 파싱 완료!\n\n' +
               '회사: ' + result.vendor + '\n' +
               '인보이스 번호: ' + result.invoiceNo + '\n' +
               '라인 수: ' + savedCount + '개' +
               validationMsg +
               '\n\nPARSING 탭에서 확인 후 "✅ 확정" 버튼을 눌러주세요.',
      vendor: result.vendor,
      invoiceNo: result.invoiceNo,
      lineCount: savedCount
    };
    
  } catch (error) {
    debugLog('parseAndSaveToParsing 오류', { error: error.toString() });
    logError('parseAndSaveToParsing', error, { fileId: fileId });
    
    return {
      success: false,
      error: '❌ 파싱 실패:\n' + error.toString()
    };
  }
}

/**
 * PDF에서 텍스트 추출
 */
function extractTextFromPdf(file) {
  try {
    // PDF를 임시 Google Doc으로 변환 (OCR 제거 - Text 기반 PDF용)
    var blob = file.getBlob();
    var resource = {
      title: 'temp_pdf_' + new Date().getTime(),
      mimeType: MimeType.GOOGLE_DOCS
    };

    var convertedFile = Drive.Files.insert(resource, blob, {
      convert: true
      // ocr: false - Text 기반 PDF는 OCR 불필요
    });
    
    var doc = DocumentApp.openById(convertedFile.id);
    var text = doc.getBody().getText();
    
    // 임시 파일 삭제
    DriveApp.getFileById(convertedFile.id).setTrashed(true);
    
    return text;
    
  } catch (error) {
    debugLog('extractTextFromPdf 오류', { error: error.toString() });
    throw new Error('PDF 텍스트 추출 실패: ' + error.toString());
  }
}

/**
 * DOCX에서 텍스트 추출
 */
function extractTextFromDocx(blob) {
  try {
    var resource = {
      title: 'temp_docx_' + new Date().getTime(),
      mimeType: MimeType.GOOGLE_DOCS
    };
    
    var file = Drive.Files.insert(resource, blob, { convert: true });
    var doc = DocumentApp.openById(file.id);
    var text = doc.getBody().getText();
    
    // 임시 파일 삭제
    DriveApp.getFileById(file.id).setTrashed(true);
    
    return text;
    
  } catch (error) {
    debugLog('extractTextFromDocx 오류', { error: error.toString() });
    throw new Error('DOCX 텍스트 추출 실패: ' + error.toString());
  }
}

/**
 * 데이터를 대상 시트로 이동 (성능 개선 버전)
 * @param {Array} existingData - 이동할 데이터 (헤더 포함)
 * @param {string} targetSheetName - 대상 시트 이름
 */
function moveDataToSheet(existingData, targetSheetName) {
  try {
    var targetSheet = getSheet(targetSheetName);
    var dataRows = existingData.slice(1); // 헤더 제외

    if (dataRows.length === 0) {
      return;
    }

    // 성능 개선: appendRow 대신 setValues 사용
    var lastRow = targetSheet.getLastRow();
    var targetRange = targetSheet.getRange(lastRow + 1, 1, dataRows.length, dataRows[0].length);
    targetRange.setValues(dataRows);

    debugLog('moveDataToSheet 완료', {
      targetSheet: targetSheetName,
      rows: dataRows.length
    });

  } catch (error) {
    debugLog('moveDataToSheet 오류', { error: error.toString() });
    throw error;
  }
}

/**
 * PARSING 시트 비우기
 */
function clearParsingSheet() {
  try {
    var sheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);

    // 헤더 제외하고 모든 데이터 삭제
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      // frozen rows 문제 방지: 범위를 먼저 지우고 행 삭제
      var numRows = lastRow - 1;
      var range = sheet.getRange(2, 1, numRows, sheet.getLastColumn());
      range.clearContent();

      // 빈 행이 많으면 삭제 (단, 최소 1개 데이터 행은 유지)
      if (numRows > 1) {
        sheet.deleteRows(3, numRows - 1);
      }
    }

    debugLog('PARSING 시트 비우기 완료');

  } catch (error) {
    debugLog('clearParsingSheet 오류', { error: error.toString() });
  }
}

/**
 * PARSING 시트에 저장
 */
function saveToParsingSheet(data) {
  try {
    var sheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);
    var savedCount = 0;
    
    for (var i = 0; i < data.lineItems.length; i++) {
      var item = data.lineItems[i];
      
      var row = [
        data.vendor,
        data.invoiceNo,
        data.invoiceDate,
        data.totalAmount,
        data.subtotal,
        data.discount,
        data.shipping,
        data.tax,
        item.lineNo,
        item.itemId,
        item.upc,
        item.description,
        item.brand,
        item.color,
        item.sizeLength,
        item.qtyOrdered,
        item.qtyShipped,
        item.unitPrice,
        item.extPrice,
        item.memo
      ];
      
      sheet.appendRow(row);
      savedCount++;
    }
    
    debugLog('PARSING 시트 저장 완료', { savedCount: savedCount });
    
    return savedCount;
    
  } catch (error) {
    debugLog('saveToParsingSheet 오류', { error: error.toString() });
    throw error;
  }
}

/**
 * 확정 (PARSING → INVOICE_SNG/OUTRE로 이동)
 */
function confirmParsing() {
  try {
    var ui = SpreadsheetApp.getUi();
    
    var parsingSheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);
    var data = parsingSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      ui.alert(
        '데이터 없음',
        'PARSING 탭에 데이터가 없습니다.\n먼저 "📄 파싱 시작"을 실행해주세요.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Vendor 확인 (첫 번째 데이터 행)
    var vendor = data[1][0]; // VENDOR 컬럼
    
    if (!vendor || (vendor !== 'SNG' && vendor !== 'OUTRE')) {
      ui.alert(
        '오류',
        'VENDOR 정보가 올바르지 않습니다: ' + vendor,
        ui.ButtonSet.OK
      );
      return;
    }
    
    var targetSheetName = vendor === 'SNG' ? 
      CONFIG.INVOICE.SNG_SHEET : 
      CONFIG.INVOICE.OUTRE_SHEET;
    
    var response = ui.alert(
      '확정 확인',
      vendor + ' 인보이스를 ' + targetSheetName + ' 탭에 추가하시겠습니까?\n\n' +
      '라인 수: ' + (data.length - 1) + '개',
      ui.ButtonSet.YES_NO
    );
    
    if (response !== ui.Button.YES) {
      return;
    }
    
    // 데이터 복사 (헤더 제외)
    var targetSheet = getSheet(targetSheetName);
    var dataRows = data.slice(1); // 헤더 제외
    
    for (var i = 0; i < dataRows.length; i++) {
      targetSheet.appendRow(dataRows[i]);
    }
    
    debugLog('확정 완료', { 
      vendor: vendor, 
      targetSheet: targetSheetName, 
      rows: dataRows.length 
    });
    
    // PARSING 시트 비우기
    clearParsingSheet();
    
    ui.alert(
      '확정 완료',
      dataRows.length + '개 라인이 ' + targetSheetName + ' 탭에 추가되었습니다.',
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      '오류',
      '확정 중 오류가 발생했습니다:\n' + error.toString(),
      ui.ButtonSet.OK
    );
    
    debugLog('confirmParsing 오류', { error: error.toString() });
    logError('confirmParsing', error);
  }
}

/**
 * 취소 (PARSING 시트 비우기)
 */
function cancelParsing() {
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    '취소 확인',
    'PARSING 탭의 데이터를 삭제하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  clearParsingSheet();
  
  ui.alert(
    '취소 완료',
    'PARSING 탭이 비워졌습니다.',
    ui.ButtonSet.OK
  );
}

// ============================================================================
// 파싱 로직
// ============================================================================

/**
 * 인보이스 파싱 (통합)
 */
function parseInvoice(text, filename) {
  debugLog('parseInvoice 시작');

  // 디버깅: 텍스트 샘플 로그
  Logger.log('=== 추출된 텍스트 샘플 (처음 500자) ===');
  Logger.log(text.substring(0, 500));
  Logger.log('=== 텍스트 길이: ' + text.length + ' ===');

  var vendor = detectVendor(text, filename);
  debugLog('회사 감지', { vendor: vendor });
  
  if (vendor === 'UNKNOWN') {
    throw new Error('인보이스 회사를 감지할 수 없습니다.');
  }
  
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
  
  data = parseCommonStructure(text, data);

  // ExtPrice 합계 계산 및 검증
  var extPriceSum = 0;
  for (var i = 0; i < data.lineItems.length; i++) {
    extPriceSum += data.lineItems[i].extPrice || 0;
  }
  extPriceSum = Number(extPriceSum.toFixed(2));

  debugLog('ExtPrice 합계 검증', {
    invoiceAmount: data.totalAmount,
    extPriceSum: extPriceSum,
    difference: Math.abs(data.totalAmount - extPriceSum)
  });

  // 차이가 $1 이상이면 경고
  var priceDifference = Math.abs(data.totalAmount - extPriceSum);
  if (priceDifference > 1.0) {
    debugLog('⚠️ 경고: Invoice Amount와 ExtPrice 합계 불일치', {
      invoiceAmount: data.totalAmount,
      extPriceSum: extPriceSum,
      difference: priceDifference
    });
  }

  debugLog('파싱 완료', {
    invoiceNo: data.invoiceNo,
    lineItems: data.lineItems.length,
    invoiceAmount: data.totalAmount,
    extPriceSum: extPriceSum,
    validated: priceDifference <= 1.0
  });

  return data;
}

/**
 * 회사 감지
 */
function detectVendor(text, filename) {
  var upperText = text.toUpperCase();
  var upperFilename = filename.toUpperCase();

  if (upperText.indexOf('SHAKE-N-GO') > -1 ||
      upperFilename.indexOf('3000') === 0 ||
      upperText.match(/\b3\d{9}\b/)) {
    return 'SNG';
  }

  if (upperText.indexOf('OUTRE') > -1 ||
      upperFilename.indexOf('SINV') > -1 ||
      upperText.match(/\bSINV\d+\b/) ||
      upperText.match(/BEAUTIFUL HAIR/i)) {
    return 'OUTRE';
  }

  return 'UNKNOWN';
}

/**
 * 공통 구조 파싱
 */
function parseCommonStructure(text, data) {
  var lines = text.split('\n');
  data = parseHeaderInfo(lines, data);
  data.lineItems = parseLineItems(lines, data.vendor);
  return data;
}

/**
 * 헤더 정보 파싱 (개선 버전)
 */
function parseHeaderInfo(lines, data) {
  var fullText = lines.join('\n');
  
  debugLog('헤더 파싱 시작', { vendor: data.vendor });
  
  // Invoice No
  if (data.vendor === 'SNG') {
    // SNG는 파일명에서 Invoice Number 추출 (10자리 숫자)
    var invoiceMatch = data.filename.match(/(\d{10})/);
    if (invoiceMatch) {
      data.invoiceNo = invoiceMatch[1];
      debugLog('SNG Invoice No 파일명에서 추출', { filename: data.filename, invoiceNo: data.invoiceNo });
    }
  } else if (data.vendor === 'OUTRE') {
    // OUTRE 인보이스 번호는 SINV 형식만 사용
    // 본문에서 찾기
    var invoiceMatch = fullText.match(/INVOICE#?\s*:?\s*(SINV\d+)/i);
    if (!invoiceMatch) {
      invoiceMatch = fullText.match(/\b(SINV\d+)\b/i);
    }
    // 본문에서 못 찾으면 파일명에서 찾기
    if (!invoiceMatch) {
      invoiceMatch = data.filename.match(/\b(SINV\d+)\b/i);
    }
    if (invoiceMatch) {
      data.invoiceNo = invoiceMatch[1].toUpperCase();
    }
  }
  
  debugLog('Invoice No 파싱', { invoiceNo: data.invoiceNo });
  
  // Invoice Date
  var dateMatch = fullText.match(/INVOICE DATE\s*:?\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/i);
  if (!dateMatch) {
    dateMatch = fullText.match(/DATE SHIPPED\s*:?\s*(\d{1,2}\/\d{1,2}\/\d{2,4})/i);
  }
  if (!dateMatch) {
    dateMatch = fullText.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})/);
  }
  
  if (dateMatch) {
    data.invoiceDate = parseDate(dateMatch[1]);
  }
  
  debugLog('Invoice Date 파싱', { invoiceDate: data.invoiceDate });
  
  // Total Amount
  if (data.vendor === 'SNG') {
    Logger.log('=== Invoice Amount 파싱 시작 (SNG) ===');
    Logger.log('텍스트 길이: ' + fullText.length + ' 문자');

    // 1. "INVOICE AMOUNT" 텍스트 위치 찾기
    var invoiceAmountPattern = /INVOICE\s+AMOUNT/gi;
    var invoiceAmountPositions = [];
    var match;

    while ((match = invoiceAmountPattern.exec(fullText)) !== null) {
      invoiceAmountPositions.push({
        text: match[0],
        index: match.index
      });
    }

    Logger.log('INVOICE AMOUNT 발견 개수: ' + invoiceAmountPositions.length);

    if (invoiceAmountPositions.length > 0) {
      // 가장 마지막 (하단) INVOICE AMOUNT 선택
      var lastInvoiceAmount = invoiceAmountPositions[invoiceAmountPositions.length - 1];
      Logger.log('마지막 INVOICE AMOUNT 위치: ' + lastInvoiceAmount.index);

      // 2. 해당 위치 이후 약 200자 범위 내에서 금액 찾기
      var searchStart = lastInvoiceAmount.index;
      var searchEnd = Math.min(searchStart + 200, fullText.length);
      var searchText = fullText.substring(searchStart, searchEnd);

      Logger.log('검색 범위: ' + searchStart + ' ~ ' + searchEnd + ' (' + searchText.length + '자)');

      // 3. 소수점 2자리 숫자 패턴 찾기 (49.99 ~ 100000.00)
      // 쉼표가 있을 수도 있고 없을 수도 있음: 4,660.00 또는 4660.00
      var amountPattern = /(\d{1,3}(?:,\d{3})*\.\d{2})/g;
      var amounts = [];

      while ((match = amountPattern.exec(searchText)) !== null) {
        var amount = parseAmount(match[1]);

        // 49.99 이상 100000.00 이하
        if (amount >= 49.99 && amount <= 100000.00) {
          amounts.push({
            amount: amount,
            text: match[1],
            relativePos: match.index,
            absolutePos: searchStart + match.index
          });

          Logger.log('유효 금액 발견: $' + amount + ' (위치: ' + (searchStart + match.index) + ', 텍스트: ' + match[1] + ')');
        }
      }

      Logger.log('유효한 금액 후보: ' + amounts.length + '개');

      if (amounts.length > 0) {
        // 가장 마지막 (가장 아래쪽) 금액 선택
        var bestAmount = amounts[amounts.length - 1];
        data.totalAmount = bestAmount.amount;

        Logger.log('✅ Invoice Amount 파싱 성공: $' + bestAmount.amount +
                   ' (위치: ' + bestAmount.absolutePos + ', 텍스트: ' + bestAmount.text + ')');
      } else {
        Logger.log('❌ Invoice Amount 파싱 실패: 유효한 금액을 찾을 수 없음');
        Logger.log('검색 텍스트 샘플:');
        Logger.log(searchText);
      }
    } else {
      Logger.log('❌ INVOICE AMOUNT 텍스트를 찾을 수 없음');
    }
  } else if (data.vendor === 'OUTRE') {
    // OUTRE는 SUBTOTAL을 먼저 파싱해야 TOTAL 검증에 사용 가능
    // Subtotal 먼저 파싱
    var subtotalMatch = fullText.match(/SUBTOTAL\s*:?\s*([\d,\.]+)/i);
    if (subtotalMatch) {
      data.subtotal = parseAmount(subtotalMatch[1]);
      Logger.log('SUBTOTAL 파싱 (TOTAL 검증용): $' + data.subtotal);
    }

    // OUTRE의 경우 인보이스 상단(첫 100-200줄)에서 TOTAL을 찾아야 함
    // "TOTAL US$2,292.75" 형식 또는 TOTAL과 금액이 다른 줄에 있을 수 있음
    // 전화번호(346-843-2709)와 혼동하지 않도록 주의

    Logger.log('=== OUTRE Total Amount 파싱 시작 ===');

    // 1. 먼저 전체 텍스트에서 "TOTAL US$" 또는 "TOTAL" 다음 줄 금액 패턴으로 시도
    var totalMatch = fullText.match(/\bTOTAL\s+US\$\s*([\d,]+\.?\d{0,2})/i);

    if (totalMatch) {
      data.totalAmount = parseAmount(totalMatch[1]);
      Logger.log('✅ TOTAL US$ 패턴 매치: $' + data.totalAmount);
    }

    // 2. 못 찾으면 첫 200줄 범위에서 "TOTAL" 근처의 금액 찾기
    if (!totalMatch) {
      var topLines = lines.slice(0, 200).join('\n');
      Logger.log('첫 200줄에서 검색 중... (길이: ' + topLines.length + ')');

      // TOTAL 근처 금액 (같은 줄 또는 다음 줄)
      totalMatch = topLines.match(/\bTOTAL[\s\S]{0,50}?([\d,]+\.\d{2})/i);

      if (totalMatch) {
        var amount = parseAmount(totalMatch[1]);
        // 전화번호(346, 843, 2709) 같은 작은 숫자 제외
        // 또한 SUBTOTAL보다 작으면 무시 (TOTAL은 SUBTOTAL보다 크거나 같아야 함)
        if (amount >= 100 && (data.subtotal === 0 || amount >= data.subtotal * 0.5)) {
          data.totalAmount = amount;
          Logger.log('✅ 첫 200줄에서 TOTAL 근처 금액 매치: $' + data.totalAmount);
        } else {
          Logger.log('⚠️ TOTAL 후보 금액이 너무 작음 (SUBTOTAL 대비): $' + amount);
        }
      }
    }

    // 3. 여전히 못 찾으면 SUBTOTAL 근처 찾기 (SUBTOTAL 바로 아래에 TOTAL이 있는 경우)
    if (!totalMatch || data.totalAmount === 0) {
      Logger.log('SUBTOTAL 근처에서 TOTAL 검색 중...');

      // SUBTOTAL 위치 찾기
      var subtotalIndex = fullText.indexOf('SUBTOTAL');
      if (subtotalIndex > -1) {
        // SUBTOTAL 이후 500자 범위
        var afterSubtotal = fullText.substring(subtotalIndex, subtotalIndex + 500);
        Logger.log('SUBTOTAL 이후 텍스트 샘플: ' + afterSubtotal.substring(0, 200));

        // SUBTOTAL 이후의 모든 금액 찾기 (SUBTOTAL 자체는 제외)
        var amountPattern = /([\d,]+\.\d{2})/g;
        var amounts = [];
        var match;

        while ((match = amountPattern.exec(afterSubtotal)) !== null) {
          var amount = parseAmount(match[1]);
          // $100 이상이고 SUBTOTAL과 다른 금액만 수집
          if (amount >= 100 && Math.abs(amount - data.subtotal) > 0.01) {
            amounts.push(amount);
            Logger.log('  금액 후보: $' + amount);
          } else if (Math.abs(amount - data.subtotal) <= 0.01) {
            Logger.log('  SUBTOTAL 값 제외: $' + amount);
          }
        }

        // SUBTOTAL보다 작거나 같은 첫 번째 큰 금액 선택 (TOTAL은 SUBTOTAL에서 할인/세금 적용 가능)
        // 70% ~ 100% 범위 내의 금액 찾기
        for (var ai = 0; ai < amounts.length; ai++) {
          if (amounts[ai] >= data.subtotal * 0.7 && amounts[ai] <= data.subtotal) {
            data.totalAmount = amounts[ai];
            Logger.log('✅ SUBTOTAL 근처 금액 선택: $' + data.totalAmount + ' (SUBTOTAL의 ' +
                       ((data.totalAmount / data.subtotal) * 100).toFixed(1) + '%)');
            break;
          }
        }

        // 못 찾으면 가장 큰 금액 선택
        if (data.totalAmount === 0 && amounts.length > 0) {
          var maxAmount = Math.max.apply(null, amounts);
          data.totalAmount = maxAmount;
          Logger.log('⚠️ SUBTOTAL 근처 최대 금액 선택: $' + data.totalAmount);
        }
      }
    }

    // 4. 그래도 못 찾으면 SUBTOTAL 값 사용 (최후 수단)
    if (data.totalAmount === 0 && data.subtotal > 0) {
      Logger.log('⚠️ TOTAL을 찾을 수 없어 SUBTOTAL 값 사용: $' + data.subtotal);
      data.totalAmount = data.subtotal;
    }
  }

  debugLog('Total Amount 파싱', { totalAmount: data.totalAmount });

  // Subtotal (OUTRE는 이미 파싱함)
  if (data.vendor !== 'OUTRE') {
    var subtotalMatch = fullText.match(/SUBTOTAL\s*:?\s*([\d,\.]+)/i);
    if (subtotalMatch) {
      data.subtotal = parseAmount(subtotalMatch[1]);
    }
  }

  debugLog('Subtotal 파싱', { subtotal: data.subtotal });
  
  // Discount
  var discountMatch = fullText.match(/DISCOUNT\s*:?\s*-?\s*([\d,\.]+)/i);
  if (discountMatch) {
    data.discount = parseAmount(discountMatch[1]);
  }
  
  // Shipping
  var shippingMatch = fullText.match(/(?:SHIPPING|S\s*&\s*H)(?:\s+CHARGE)?\s*:?\s*([\d,\.]+)/i);
  if (shippingMatch) {
    data.shipping = parseAmount(shippingMatch[1]);
  }
  
  // Tax
  var taxMatch = fullText.match(/\bTAX\s*:?\s*([\d,\.]+)/i);
  if (taxMatch) {
    data.tax = parseAmount(taxMatch[1]);
  }
  
  debugLog('헤더 파싱 완료', {
    invoiceNo: data.invoiceNo,
    invoiceDate: data.invoiceDate,
    totalAmount: data.totalAmount,
    subtotal: data.subtotal,
    discount: data.discount,
    shipping: data.shipping,
    tax: data.tax
  });
  
  return data;
}

/**
 * 라인 아이템 파싱 (SNG/OUTRE 통합, 개선 버전)
 */
function parseLineItems(lines, vendor) {
  var items = [];
  var lineNo = 1;

  debugLog('라인 아이템 파싱 시작', { vendor: vendor, totalLines: lines.length });

  // OUTRE의 경우: 테이블 헤더를 찾아서 그 이후부터만 파싱
  var startLine = 0;
  if (vendor === 'OUTRE') {
    // 1단계: "QTY SHIPPED" 패턴 찾기
    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];

      // "QTY SHIPPED" 또는 "QTY" + "SHIPPED" 패턴
      if (line.match(/QTY\s+SHIPPED/i) || line.match(/QTY.*SHIPPED/i) ||
          (line.match(/\bQTY\b/i) && i + 1 < lines.length && lines[i + 1].match(/SHIPPED/i))) {

        debugLog('QTY SHIPPED 헤더 후보 발견', { line: i, text: line.substring(0, 50) });

        // 2단계: 근처에 DESCRIPTION, UNIT PRICE 등 확인
        var foundHeader = false;
        for (var j = i; j < Math.min(i + 10, lines.length); j++) {
          if (lines[j].match(/DESCRIPTION|UNIT.*PRICE|DISC.*PRICE|EXT.*PRICE/i)) {
            foundHeader = true;
            debugLog('가격/설명 헤더 발견', { line: j, text: lines[j].substring(0, 50) });
            break;
          }
        }

        if (foundHeader) {
          // 3단계: 헤더 이후에서 실제 제품 라인 찾기
          // OUTRE는 여러 줄 형식: QTY만 있는 라인을 찾음
          Logger.log('=== 헤더 발견 후 첫 20줄 검사 시작 (라인 ' + i + ' 이후) ===');

          for (var k = i + 1; k < Math.min(i + 30, lines.length); k++) {
            var testLine = lines[k].trim();

            Logger.log('  [' + k + '] 길이=' + testLine.length + ' | ' + testLine.substring(0, 100));

            // OUTRE 다중 라인 형식: QTY만 있는 라인 찾기 (1~3자리 숫자만)
            if (testLine.match(/^\d{1,3}$/)) {
              var qty = parseInt(testLine);

              Logger.log('    QTY 전용 라인 발견: ' + qty);

              // 수량 범위 검증 (0-700)
              if (qty >= 0 && qty <= 700) {
                startLine = k;
                Logger.log('  ✅ 테이블 시작 라인 확정 (QTY 전용): ' + k);
                debugLog('OUTRE 테이블 시작 라인 찾음 (다중 라인 형식)', {
                  headerLine: i,
                  startLine: startLine,
                  firstItemQty: qty,
                  headerText: line.substring(0, 50)
                });
                break;
              }
            }
          }

          if (startLine > 0) {
            break; // 찾았으면 루프 종료
          }
        }
      }
    }

    // 못 찾았으면 경고 로그
    if (startLine === 0) {
      debugLog('⚠️ OUTRE 테이블 시작점을 찾지 못함 - 전체 텍스트에서 파싱 시도');
    }
  }

  for (var i = startLine; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

    var isItemLine = false;
    var parts = [];

    if (vendor === 'SNG') {
      var tabParts = line.split('\t');

      if (tabParts.length >= 7) {
        var firstCol = tabParts[0].trim();
        var secondCol = tabParts[1].trim();
        var thirdCol = tabParts[2].trim();

        if (firstCol.match(/^[A-Z]\d+$/) &&
            !isNaN(parseInt(secondCol)) &&
            !isNaN(parseInt(thirdCol))) {
          isItemLine = true;
          parts = tabParts;
        }
      }

    } else if (vendor === 'OUTRE') {
      // OUTRE 다중 라인 형식:
      // Line 1: QTY만 (예: "5")
      // Line 2: DESCRIPTION (예: "BIG BEAUTIFUL HAIR CLIP-IN- 9PCS - PERUVIAN WAVE 18" - HT")
      // Line 3: COLORS (예: "CBRN- 2   JBLK- 0 (2)   NBLK- 1 (1)   NBRN- 2")
      // Line 4-6: 빈 줄들
      // Line 7: UNIT PRICE (예: "18.00")
      // Line 8: DISC PRICE (예: "17.00")
      // Line 9: EXT PRICE (예: "85.00")

      // QTY만 있는 라인 감지 (1~3자리 숫자만)
      if (line.match(/^\d{1,3}$/)) {
        var qty = parseInt(line);

        // 수량 범위 검증 (0-700) + Description 검증
        if (qty >= 0 && qty <= 700 && i + 1 < lines.length) {
          var nextLine = lines[i + 1].trim();

          // 다음 줄이 유효한 제품 Description인지 확인
          // 제품명 패턴: 대문자로 시작, 제품 관련 키워드 포함
          // 긍정 키워드: HAIR, WIG, LACE, WEAVE, CLIP, REMI, BATIK, SUGARPUNCH, X-PRESSION, BEAUTIFUL, MELTED,
          //              BRAID, CLOSURE, WAVE, CURL, STRAIGHT, BUNDLE, PONYTAIL, TARA, QW, BIG, BOHEMIAN, HD, PERUVIAN, TWIST, FEED
          // 제외: "COD tag Fee", 메타데이터, 전화번호 등
          var hasProductKeywords = nextLine.match(/HAIR|WIG|LACE|WEAVE|CLIP|REMI|BATIK|SUGARPUNCH|X-PRESSION|BEAUTIFUL|MELTED|BRAID|CLOSURE|WAVE|CURL|STRAIGHT|BUNDLE|PONYTAIL|TARA|QW|BIG|BOHEMIAN|HD|PERUVIAN|TWIST|FEED|LOOKS|PASSION/i);
          var hasMetadata = nextLine.match(/\bSHIP\s+TO\b|\bSOLD\s+TO\b|\bWEIGHT\b|\bSUBTOTAL\b|\bRICHMOND\b|\bLLC\b|\bPKWAY\b|\bCOD\b|\bFee\b|\btag\b|\bDATE\s+SHIPPED\b|\bPAGE\b|\bSHIP\s+VIA\b|\bPAYMENT\b|\bTERMS\b/i);
          var startsWithUpperCase = nextLine.match(/^[A-Z]/);

          // 2개 이상의 연속된 대문자 단어가 있거나 제품 키워드가 있으면 유효
          var hasMultipleUpperWords = nextLine.match(/[A-Z]{2,}.*[A-Z]{2,}/);

          var isValidDescription = startsWithUpperCase &&
                                  (hasProductKeywords || hasMultipleUpperWords) &&
                                  !hasMetadata;

          // 디버깅: 왜 제외되었는지 로그
          if (!isValidDescription) {
            Logger.log('  🔍 Description 검증 실패: ' + nextLine.substring(0, 50));
            Logger.log('    startsWithUpperCase: ' + !!startsWithUpperCase);
            Logger.log('    hasProductKeywords: ' + !!hasProductKeywords);
            Logger.log('    hasMultipleUpperWords: ' + !!hasMultipleUpperWords);
            Logger.log('    hasMetadata: ' + !!hasMetadata);
            if (hasMetadata) {
              Logger.log('    매칭된 메타데이터: ' + hasMetadata[0]);
            }
          }

          if (isValidDescription) {
            isItemLine = true;
            parts = [line]; // QTY만 저장
          } else {
            Logger.log('  ⏭️ QTY 후보 제외 (유효한 Description 아님): ' + qty + ' → ' + nextLine.substring(0, 50));
          }
        }
      }
    }

    if (isItemLine) {
      debugLog('아이템 라인 감지', { line: i, vendor: vendor, parts: parts.length });

      var qtyOrdered = 0;
      var qtyShipped = 0;
      var itemId = '';
      var description = '';
      var descriptionBeforeCleanup = ''; // 원본 Description (cleanup 전)
      var unitPrice = 0;
      var extPrice = 0;

      if (vendor === 'SNG') {
        qtyOrdered = parseInt(parts[1]) || 0;
        qtyShipped = parseInt(parts[2]) || 0;
        itemId = parts[3] ? parts[3].trim() : '';

        // Description은 4번째 컬럼
        description = parts[4] ? parts[4].trim() : '';

        // 첫 번째 라인에서 Unit Price (5번째 컬럼)와 Ext Price (6번째 컬럼)
        unitPrice = parseAmount(parts[5]);
        extPrice = parseAmount(parts[6]);

        debugLog('SNG 1행 파싱', {
          description: description,
          unitPrice: unitPrice,
          extPrice: extPrice
        });

        // 두 번째 라인 확인 (할인된 가격)
        if (i + 1 < lines.length) {
          var nextLine = lines[i + 1];
          var nextParts = nextLine.split('\t');

          // 두 번째 라인이 "\t4.00\t160.00\t80.00" 형식인지 확인
          if (nextParts.length >= 4 &&
              nextParts[0].trim() === '' &&
              !isNaN(parseFloat(nextParts[1])) &&
              !isNaN(parseFloat(nextParts[2])) &&
              !isNaN(parseFloat(nextParts[3]))) {

            // 할인된 가격 사용
            unitPrice = parseAmount(nextParts[1]);
            extPrice = parseAmount(nextParts[2]);

            debugLog('SNG 2행 가격 적용', {
              unitPrice: unitPrice,
              extPrice: extPrice
            });
          }
        }

      } else if (vendor === 'OUTRE') {
        // OUTRE 다중 라인 파싱
        // parts[0] = QTY (라인 i)
        // 다음 라인들: DESCRIPTION (1-2줄), COLORS (다중 줄 가능), PRICES (3줄)

        qtyShipped = parseInt(parts[0]) || 0;
        qtyOrdered = qtyShipped;
        itemId = '';

        Logger.log('=== OUTRE 다중 라인 파싱 시작 (라인 ' + i + ', QTY=' + qtyShipped + ') ===');

        // 다음 15줄 안에서 DESCRIPTION, COLORS, PRICES 찾기
        var descriptionLines = [];
        var colorLinesArray = []; // 여러 줄 컬러 지원
        var priceLines = [];
        var foundFirstColor = false; // 첫 컬러 라인 발견 플래그

        for (var j = i + 1; j < Math.min(i + 15, lines.length); j++) {
          var nextLine = lines[j].trim();

          Logger.log('  [' + j + '] ' + nextLine.substring(0, 80));

          // 빈 줄 건너뛰기
          if (!nextLine) continue;

          // CRITICAL: 가격을 모두 찾았으면 즉시 종료 (최우선 체크)
          if (priceLines.length >= 3) {
            Logger.log('    ✅ 가격 3개 수집 완료, 즉시 파싱 종료');
            break;
          }

          // 다음 아이템 라인을 만나면 중단 (QTY만 있는 라인 + 뒤에 Description이 와야 함)
          if (nextLine.match(/^\d{1,3}$/)) {
            var possibleQty = parseInt(nextLine);
            // 수량 범위 검증 + 다음 줄이 Description인지 확인
            if (possibleQty >= 0 && possibleQty <= 700 && j + 1 < lines.length) {
              var nextNextLine = lines[j + 1].trim();

              // QTY 검증 로직과 동일하게 적용
              var hasProductKeywords = nextNextLine.match(/HAIR|WIG|LACE|WEAVE|CLIP|REMI|BATIK|SUGARPUNCH|X-PRESSION|BEAUTIFUL|MELTED|BRAID|CLOSURE|WAVE|CURL|STRAIGHT|BUNDLE|PONYTAIL|TARA|QW|BIG|BOHEMIAN|HD|PERUVIAN|TWIST|FEED|LOOKS|PASSION/i);
              var hasMetadata = nextNextLine.match(/\bSHIP\s+TO\b|\bSOLD\s+TO\b|\bWEIGHT\b|\bSUBTOTAL\b|\bRICHMOND\b|\bLLC\b|\bPKWAY\b|\bCOD\b|\bFee\b|\btag\b|\bDATE\s+SHIPPED\b|\bPAGE\b|\bSHIP\s+VIA\b|\bPAYMENT\b|\bTERMS\b|\bSALES\b|\bTOTAL\b/i);
              var startsWithUpperCase = nextNextLine.match(/^[A-Z]/);
              var hasMultipleUpperWords = nextNextLine.match(/[A-Z]{2,}.*[A-Z]{2,}/);

              var isValidDescription = startsWithUpperCase &&
                                      (hasProductKeywords || hasMultipleUpperWords) &&
                                      !hasMetadata;

              if (isValidDescription) {
                Logger.log('  ✋ 다음 아이템 라인 발견 (QTY + Description), 중단');
                break;
              }
            }
            // 그 외 단순 숫자는 넘어감 (컬러 라인의 일부일 수 있음)
          }

          // 소수점 2자리 금액 패턴 (18.00, 17.00, 85.00 등)
          if (nextLine.match(/^[\d,]+\.\d{2}$/)) {
            priceLines.push(parseAmount(nextLine));
            Logger.log('    ✓ 가격 라인: $' + priceLines[priceLines.length - 1]);
            // 가격 3개 수집 완료 시 즉시 루프 종료
            if (priceLines.length >= 3) {
              Logger.log('    ✅ 가격 3개 수집 완료 (별도 라인), 즉시 파싱 종료');
              break;
            }
            continue;
          }

          // 메타데이터 필터링 확장 (SHIP TO, SOLD TO, WEIGHT, 전화번호, 주소 등)
          if (nextLine.match(/SHIP\s+TO|SOLD\s+TO|WEIGHT\(S\)|SUBTOTAL|RICHMOND|LLC|PKWAY|DATE\s+SHIPPED|P\.O\.|SHIP\s+VIA|PAYMENT|TERMS|SHIPPING|Sales\s+Rep|PAGE|METHOD|Free\s+Shipment/i)) {
            Logger.log('    ⏭️ 메타데이터 라인 건너뜀: ' + nextLine.substring(0, 50));
            continue;
          }

          // 전화번호 패턴 필터링 (346/843-2709, 123-456-7890 등)
          if (nextLine.match(/^\d{3}[\/\-]\d{3}[\/\-]\d{4}$/)) {
            Logger.log('    ⏭️ 전화번호 라인 건너뜀: ' + nextLine.substring(0, 50));
            continue;
          }

          // 주소 패턴 필터링 (숫자로 시작하는 주소, "US", "TX" 등)
          if (nextLine.match(/^\d+\s+[A-Z].*(?:PKWAY|BLVD|AVE|ST|RD|DR)/i) || nextLine.match(/^US$/) || nextLine.match(/^[A-Z]{2}\s*$/)) {
            Logger.log('    ⏭️ 주소 라인 건너뜀: ' + nextLine.substring(0, 50));
            continue;
          }

          // Description 수집 플래그 체크 (컬러 발견 전까지만, 최대 3줄)
          var isDescriptionCandidate = false;
          if (!foundFirstColor && descriptionLines.length < 3) {
            // CRITICAL: 바로 이전 줄이 QTY 전용 라인(숫자만)이면, 현재 줄은 다음 아이템의 Description
            // 현재 아이템의 Description에 추가하면 안 됨!
            var isPreviousLineQty = (j >= i + 2) && lines[j - 1].trim().match(/^\d{1,3}$/);

            if (!isPreviousLineQty) {
              // Description은 제품명 패턴이어야 함
              var isDescriptionLine = nextLine.match(/^[A-Z]/) &&
                                     !nextLine.match(/^\d+$/) &&
                                     !nextLine.match(/SHIP\s+TO|SOLD\s+TO|WEIGHT|SUBTOTAL|RICHMOND|LLC|PKWAY|COD|\bFee\b|tag|DATE\s+SHIPPED|PAGE|SHIP\s+VIA|PAYMENT|TERMS|SALES|TOTAL|US$/i);

              // 또는 인치 표시만 있는 라인 (예: '10" 12" 14"')
              var isInchLine = nextLine.match(/^\d+["″'']/);

              // 1-2-3 스타일 또는 인치 리스트는 description으로 처리
              var hasThreeNumberPattern = nextLine.match(/\b\d+-\d+-\d+\b/);
              var hasMultipleInches = nextLine.match(/\d+["″'']\s+\d+["″'']/); // "10" 12" 같은 패턴

              // CRITICAL: 가격이 포함된 라인은 Description 후보에서 제외
              // 예: "NA- 2   NBLK- 2   	19.50	17.00	68.00"
              var hasPrices = nextLine.match(/\d+\.\d{2}/);

              // 디버깅 로그 추가
              if (j === i + 1) {
                Logger.log('    🔍 첫 Description 후보 검증: ' + nextLine.substring(0, 50));

                var startsWithUpper = nextLine.match(/^[A-Z]/);
                var notOnlyDigits = !nextLine.match(/^\d+$/);
                var metadataMatch = nextLine.match(/SHIP\s+TO|SOLD\s+TO|WEIGHT|SUBTOTAL|RICHMOND|LLC|PKWAY|COD|\bFee\b|tag|DATE\s+SHIPPED|PAGE|SHIP\s+VIA|PAYMENT|TERMS|SALES|TOTAL|US$/i);

                Logger.log('      startsWithUpper: ' + !!startsWithUpper);
                Logger.log('      notOnlyDigits: ' + !!notOnlyDigits);
                Logger.log('      metadataMatch: ' + (metadataMatch ? metadataMatch[0] : 'null'));
                Logger.log('      isDescriptionLine: ' + !!isDescriptionLine);
                Logger.log('      isInchLine: ' + !!isInchLine);
                Logger.log('      hasThreeNumberPattern: ' + !!hasThreeNumberPattern);
                Logger.log('      hasMultipleInches: ' + !!hasMultipleInches);
                Logger.log('      hasPrices: ' + !!hasPrices);
              }

              if ((isDescriptionLine || isInchLine || hasThreeNumberPattern || hasMultipleInches) && !hasPrices) {
                isDescriptionCandidate = true;
              }
            } else {
              Logger.log('    ⏭️ 이전 줄이 QTY, 다음 아이템의 Description으로 판단: ' + nextLine.substring(0, 50));
            }
          } else if (!foundFirstColor && descriptionLines.length >= 3) {
            Logger.log('    ⏭️ Description 3줄 도달, 추가 건너뜀: ' + nextLine.substring(0, 50));
          }

          // 숫자만 있는 라인 건너뛰기 (예: "265.00", "2387257")
          // 단, 컬러 라인의 연속일 수 있으므로 문맥 확인
          if (nextLine.match(/^[\d\s.,]+$/) && !foundFirstColor) {
            Logger.log('    ⏭️ 숫자 전용 라인 건너뜀 (컬러 전): ' + nextLine.substring(0, 50));
            continue;
          }

          // 컬러 패턴 매치 (일반적인 "COLOR- QTY" 형식)
          // 단, Description의 일부 (예: '18" - HT')는 제외
          //
          // 실제 컬러 라인 패턴:
          //   ✅ "1B- 2", "NA- 2", "NBLK- 2" (짧은 컬러)
          //   ✅ "DRFFCARMCH- 1", "M950/425/350/130S- 2" (긴 컬러, 최대 16글자)
          //   ❌ "REMI TARA 1-2-3" (Description + 숫자 패턴)
          //   ❌ "SUGARPUNCH - 4X4 HD..." (Description)
          //   ❌ '18" - HT' (인치 뒤 하이픈)

          var hasColorPattern = nextLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/);
          var isInchPattern = nextLine.match(/\d+["″'']\s*-/); // 인치 뒤 하이픈 (18" - HT)

          // Description 블랙리스트 (컬러가 아닌 제품명)
          // CRITICAL: 블랙리스트는 "컬러 라인처럼 보이지만 Description인 경우"를 걸러내기 위한 것
          // Description 수집 단계에서는 적용하지 않고, 컬러 라인 판별 시에만 사용
          var DESCRIPTION_BLACKLIST = [
            'SUGARPUNCH', 'HONEYPUNCH', 'REMI TARA', 'BATIK', 'X-PRESSION',
            'BEAUTIFUL HAIR', 'MELTED', 'SWOOP', 'PERFECT HAIR LINE',
            'LACE FRONT', 'LACE CLOSURE', 'HD LACE', 'BOHEMIAN', 'PERUVIAN',
            'UNPROCESSED', 'CLIP-IN', 'PONYTAIL', 'BUNDLE', 'WEAVE', 'WAVE',
            'CURL', 'STRAIGHT', 'BODY WAVE', 'BIG BEAUTIFUL', 'HD BOHEMIAN'
          ];

          var hasBlacklistedWord = false;
          var upperLine = nextLine.toUpperCase();

          // CRITICAL: 블랙리스트 체크는 컬러 패턴이 있을 때만 적용
          // 컬러 패턴이 없으면 어차피 컬러 라인이 아니므로 체크할 필요 없음
          if (hasColorPattern) {
            for (var bi = 0; bi < DESCRIPTION_BLACKLIST.length; bi++) {
              if (upperLine.indexOf(DESCRIPTION_BLACKLIST[bi]) > -1) {
                hasBlacklistedWord = true;
                Logger.log('    ⛔ 블랙리스트 매칭 (컬러 패턴 제외): "' + DESCRIPTION_BLACKLIST[bi] + '" in "' + nextLine.substring(0, 50) + '"');
                break;
              }
            }
          }

          var isColorLine = hasColorPattern && !isInchPattern && !hasBlacklistedWord;

          // foundFirstColor 플래그로 연속 컬러 라인 허용
          if (foundFirstColor && hasColorPattern && !isInchPattern && !hasBlacklistedWord) {
            isColorLine = true;
          }

          // CRITICAL: Description 후보 처리
          // - Description 후보이면서 컬러 라인이 아닌 경우: Description으로 추가하고 continue
          // - Description 후보이면서 컬러 라인인 경우: Description으로 추가하되 continue 하지 않음 (컬러 처리로 진행)
          // - 예외: 블랙리스트가 있어도 괄호 컬러 패턴 (P)COLOR- QTY가 있으면 컬러 라인으로 처리
          var hasParenColorPattern = nextLine.match(/\([A-Z]\)[A-Z0-9\-\/]+\s*-\s*\d+/);

          if (isDescriptionCandidate && !isColorLine && !hasParenColorPattern) {
            descriptionLines.push(nextLine);
            Logger.log('    ✓ Description 라인 추가 (' + descriptionLines.length + '/3): ' + nextLine.substring(0, 50));
            continue; // 다음 줄로 이동
          } else if (isDescriptionCandidate && (isColorLine || hasParenColorPattern)) {
            descriptionLines.push(nextLine);
            Logger.log('    ✓ Description 라인 추가 (컬러 포함, 컬러 처리 계속): ' + nextLine.substring(0, 50));
            // continue 하지 않음 - 아래 컬러 처리로 진행
            // hasParenColorPattern이 있으면 isColorLine을 강제로 true로 설정
            if (hasParenColorPattern) {
              isColorLine = true;
              Logger.log('    ✅ 괄호 컬러 패턴 발견, 블랙리스트 무시하고 컬러 라인으로 처리');
            }
          } else if (isDescriptionCandidate) {
            Logger.log('    ⏭️ Description 후보 제외 (메타데이터 또는 패턴 불일치): ' + nextLine.substring(0, 50));
          }

          if (isColorLine) {
            // 컬러 라인에 가격 정보가 포함되어 있는지 확인
            // 예: "NA- 2   NBLK- 2   	19.50	17.00	68.00"
            // 마지막에 소수점 2자리 숫자가 3개 있으면 가격으로 추출
            // 탭 또는 공백으로 구분될 수 있음
            var pricesInColorLine = nextLine.match(/([\d,]+\.\d{2})[\s\t]+([\d,]+\.\d{2})[\s\t]+([\d,]+\.\d{2})\s*$/);

            if (pricesInColorLine && priceLines.length === 0) {
              // 가격 추출
              priceLines.push(parseAmount(pricesInColorLine[1])); // UNIT PRICE
              priceLines.push(parseAmount(pricesInColorLine[2])); // DISC PRICE
              priceLines.push(parseAmount(pricesInColorLine[3])); // EXT PRICE

              Logger.log('    ✓ 컬러 라인에서 가격 추출: $' + pricesInColorLine[1] + ', $' + pricesInColorLine[2] + ', $' + pricesInColorLine[3]);

              // 가격 부분을 제거한 컬러 정보만 저장
              var colorOnly = nextLine.replace(pricesInColorLine[0], '').trim();
              colorLinesArray.push(colorOnly);
              Logger.log('    ✓ 컬러 라인 추가 (가격 제거됨): ' + colorOnly.substring(0, 50));

              // CRITICAL: 가격 3개 추출 완료 시 즉시 루프 종료
              foundFirstColor = true;
              Logger.log('    ✅ 컬러 라인에서 가격 3개 수집 완료, 즉시 파싱 종료');
              break;
            } else {
              // 가격 없는 일반 컬러 라인
              colorLinesArray.push(nextLine);
              Logger.log('    ✓ 컬러 라인 추가: ' + nextLine.substring(0, 50));
            }

            foundFirstColor = true;
            continue;
          }

          // 컬러 연속 라인: "(숫자)" 만 있는 경우 (backordered 정보)
          // 예: "S1B/BU- 0 \n(1)   "
          if (foundFirstColor && nextLine.match(/^\((\d+)\)\s*$/)) {
            // 이전 컬러 라인에 붙여서 추가
            if (colorLinesArray.length > 0) {
              var lastColorLine = colorLinesArray[colorLinesArray.length - 1];
              colorLinesArray[colorLinesArray.length - 1] = lastColorLine + ' ' + nextLine;
              Logger.log('    ✓ 컬러 라인에 backordered 추가: ' + nextLine.substring(0, 50));
            }
            continue;
          }

          // 컬러 라인 연속 중단 조건: 가격이 나오거나 메타데이터가 나옴
          if (foundFirstColor && priceLines.length > 0) {
            Logger.log('    ✋ 컬러 라인 수집 완료 (가격 시작)');
            // 더 이상 컬러 수집 안 함
          }
        }

        // Description 여러 줄을 공백으로 연결
        description = descriptionLines.join(' ');

        // CRITICAL: Description cleanup 전에 원본 저장
        // parseColorLinesImproved에서 제거할 때 필요
        descriptionBeforeCleanup = description;

        // Description 후처리: 컬러 패턴이 섞여 있으면 제거
        // 예 1: "X-PRESSION BRAID-PRE STRETCHED BRAID 52" 3X (P)M950/425/350/130S- 55"
        //   → "X-PRESSION BRAID-PRE STRETCHED BRAID 52" 3X"
        // 예 2: "BIG BEAUTIFUL HAIR CLIP-IN- 9PCS - PERUVIAN WAVE 18" - HT CBRN- 2   JBLK- 0 (2)"
        //   → "BIG BEAUTIFUL HAIR CLIP-IN- 9PCS - PERUVIAN WAVE 18""
        // 예 3: "LACE FRONT WIG-PERFECT HAIR LINE13X4-SWOOP SERIES-SWOOP1-HT DRFFAMSS- 1 DRFFCARMCH- 1"
        //   → "LACE FRONT WIG-PERFECT HAIR LINE13X4-SWOOP SERIES-SWOOP1-HT"

        Logger.log('  📝 Description 정리 전: ' + description);

        // 케이스 1: 괄호로 시작하는 컬러 패턴 제거
        // "52" 3X (P)M950..." → "52" 3X"
        var colorInDescMatch = description.match(/^(.+?)(\d+["″''])\s*(\d*X)?\s*\([A-Z0-9\/\-]+\)/i);
        if (colorInDescMatch) {
          // 기본: 인치 부분까지
          var cleanDesc = (colorInDescMatch[1] + colorInDescMatch[2]).trim();
          // 배수 표시가 있으면 공백 + 배수 추가
          if (colorInDescMatch[3]) {
            cleanDesc += ' ' + colorInDescMatch[3];
          }
          description = cleanDesc;
          Logger.log('  🔧 Description 정리 (괄호 컬러 패턴 제거): ' + description);
        }

        // 케이스 2: 일반 컬러 패턴 제거 (COLOR- QTY 형식)
        // 연속된 인치 패턴을 모두 유지하고, 컬러 패턴이 시작되기 직전까지만 유지
        // 예: "10" 12" 14"" → 전체 유지, "18" - HT" → 전체 유지
        // 예: "10" 12" 14" NA- 2" → "10" 12" 14""만 유지

        // 먼저 인치 패턴이 있는지 확인
        var hasInch = description.match(/\d+["″'']/);
        if (hasInch) {
          // 연속된 인치 패턴 매칭 (공백 또는 공백 없이)
          // 패턴: 10" 12" 14" 또는 10"12"14" 또는 18" - HT 또는 10" 3X
          // 마지막 인치 이후에 " - HT" 또는 " 3X" 같은 suffix 허용
          var allInchesPattern = description.match(/^(.+?)(\d+["″''](?:\s*\d+["″''])*(?:\s*(?:-\s*[A-Z]{2,3}|\d*X))?)/);

          if (allInchesPattern) {
            var beforeCleanup = description;
            // 텍스트 부분 + 모든 인치 + suffix
            description = (allInchesPattern[1] + allInchesPattern[2]).trim();

            if (beforeCleanup !== description) {
              Logger.log('  🔧 Description 정리 (연속 인치 유지): ' + description);
            }
          }
        } else {
          // 인치가 없으면 첫 번째 COLOR- QTY 패턴 직전까지만 유지
          // COLOR- QTY 패턴: 2글자 이상 대문자/숫자/하이픈/슬래시 + " - " + 숫자
          // 예외: "- HT", "- 9PCS" 같은 단어는 제외 (숫자만 있어야 컬러 패턴)
          var firstColorPattern = description.match(/^(.+?)\s+([A-Z0-9\/\-]{2,})\s*-\s*\d+/);
          if (firstColorPattern) {
            var beforeCleanup = description;
            description = firstColorPattern[1].trim();
            if (beforeCleanup !== description) {
              Logger.log('  🔧 Description 정리 (컬러 패턴 절단): ' + description);
            }
          }
        }

        Logger.log('  📝 최종 Description: ' + description.substring(0, 80));

        // 가격 정보 (최소 3개 필요: UNIT, DISC, EXT)
        if (priceLines.length >= 3) {
          var regularPrice = priceLines[0];  // UNIT PRICE (정가)
          unitPrice = priceLines[1];  // DISC PRICE (할인가) - 이것을 사용
          extPrice = priceLines[2];   // EXT PRICE

          Logger.log('  ✅ 가격 추출: REGULAR=$' + regularPrice + ', DISC(사용)=$' + unitPrice + ', EXT=$' + extPrice);
        } else {
          Logger.log('  ⚠️ 가격 정보 부족: ' + priceLines.length + '개만 발견');
          unitPrice = 0;
          extPrice = 0;
        }

        // 컬러 정보 처리 (다중 라인 결합)
        if (colorLinesArray.length > 0) {
          colorLines = colorLinesArray;
          Logger.log('  ✅ 컬러 라인 설정: ' + colorLinesArray.length + '줄');
          for (var clIdx = 0; clIdx < colorLinesArray.length; clIdx++) {
            Logger.log('    [' + clIdx + '] ' + colorLinesArray[clIdx].substring(0, 50));
          }
        } else {
          colorLines = [];
          Logger.log('  ⚠️ 컬러 라인 없음');
        }
      }

      debugLog('아이템 파싱 결과', {
        itemId: itemId,
        description: description,
        qtyOrdered: qtyOrdered,
        qtyShipped: qtyShipped,
        unitPrice: unitPrice,
        extPrice: extPrice
      });

      // 길이 추출 (예: 10"12"14" → 그대로 유지)
      // 복수 길이 패턴: 10"12"14" 또는 단일 길이: 18"
      // 공백으로 나뉘어 있을 수도 있음: "10" 12" 14"" → "10"12"14""로 합침
      var sizeMatch = description.match(/(\d+["″'']\s*)+/);
      var size = '';
      if (sizeMatch) {
        // 공백 제거하고 합치기
        size = sizeMatch[0].replace(/\s+/g, '');
      }

      // OUTRE의 경우 colorLines가 이미 설정되어 있으므로, 조건부로 초기화
      if (typeof colorLines === 'undefined') {
        var colorLines = [];
      }
      var priceInfo = { unitPrice: unitPrice, extPrice: extPrice }; // OUTRE에서 사용
      var searchLog = {
        itemId: itemId,
        searchRange: Math.min(i + 50, lines.length) - (i + 1),
        linesChecked: 0,
        linesFiltered: [],
        linesCollected: []
      };

      // OUTRE는 이미 같은 라인에서 모든 정보를 파싱했으므로 다음 라인 검색 건너뛰기
      if (vendor === 'OUTRE' && colorLines.length > 0) {
        debugLog('OUTRE: 같은 라인에서 컬러 정보 이미 파싱됨, 다음 라인 검색 건너뛰기', {
          colorCount: colorLines.length
        });
        // 바로 컬러 데이터 처리로 건너뜀
      } else {
        // SNG 또는 OUTRE에서 컬러를 못 찾은 경우, 다음 라인 검색
        for (var j = i + 1; j < Math.min(i + 50, lines.length); j++) {
        var nextLine = lines[j].trim();
        searchLog.linesChecked++;

        // 다음 아이템 라인을 만나면 중단
        if (vendor === 'SNG' && nextLine.match(/^[A-Z]\d+\t/)) {
          searchLog.linesFiltered.push({ line: j, reason: '다음 아이템 라인', text: nextLine.substring(0, 50) });
          break;
        }
        if (vendor === 'OUTRE' && nextLine.match(/^\d+[\t\s]+[A-Z]/)) {
          searchLog.linesFiltered.push({ line: j, reason: '다음 아이템 라인', text: nextLine.substring(0, 50) });
          break;
        }

        if (!nextLine) continue;

        // 페이지 헤더/푸터 패턴 무시 (확장)
        if (nextLine.match(/^Page \d+/i) || nextLine.match(/PAGE \d+ of \d+/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'Page 번호', text: nextLine });
          continue;
        }
        if (nextLine.match(/SHAKE-N-GO/i) || nextLine.match(/OUTRE/i)) {
          searchLog.linesFiltered.push({ line: j, reason: '회사명', text: nextLine });
          continue;
        }
        if (nextLine.match(/^INVOICE/i) && nextLine.length < 50) {
          searchLog.linesFiltered.push({ line: j, reason: 'INVOICE 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^[\-=]+$/)) {
          searchLog.linesFiltered.push({ line: j, reason: '구분선', text: nextLine });
          continue;
        }
        // OUTRE 특수 헤더
        if (nextLine.match(/QTY\s+SHIPPED.*DESCRIPTION/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'OUTRE 테이블 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/UNIT\s+PRICE.*DISC.*PRICE.*EXT.*PRICE/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'OUTRE 가격 헤더', text: nextLine });
          continue;
        }

        // 헤더 패턴 필터링 (추가)
        if (nextLine.match(/^\s*QTY\s+.*\s+ITEM/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'QTY...ITEM 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*ORDERED\s+SHIPPED/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'ORDERED SHIPPED 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*ITEM\s+NUMBER/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'ITEM NUMBER 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*DESCRIPTION/i) && nextLine.length < 50) {
          searchLog.linesFiltered.push({ line: j, reason: 'DESCRIPTION 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*UNIT\s+PRICE/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'UNIT PRICE 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*EXT\.?\s+PRICE/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'EXT PRICE 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*ORDER\s+NUMBER/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'ORDER NUMBER 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*CUSTOMER/i) && nextLine.length < 50) {
          searchLog.linesFiltered.push({ line: j, reason: 'CUSTOMER 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*SHIP\s+TO/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'SHIP TO 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*SOLD\s+TO/i)) {
          searchLog.linesFiltered.push({ line: j, reason: 'SOLD TO 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*DATE/i) && nextLine.length < 30) {
          searchLog.linesFiltered.push({ line: j, reason: 'DATE 헤더', text: nextLine });
          continue;
        }
        if (nextLine.match(/^\s*TERMS/i) && nextLine.length < 30) {
          searchLog.linesFiltered.push({ line: j, reason: 'TERMS 헤더', text: nextLine });
          continue;
        }

        // 언더스코어가 있는 컬러 라인 (주로 SNG)
        if (nextLine.indexOf('_') > -1) {
          colorLines.push(nextLine);
          searchLog.linesCollected.push({ line: j, type: '언더스코어', text: nextLine });
          continue;
        }

        // 컬러 패턴 매치
        if (nextLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/)) {
          // OUTRE의 경우, 컬러 라인에서 가격 정보도 추출
          if (vendor === 'OUTRE') {
            var priceMatch = nextLine.match(/([\d.]+)\s+([\d.]+)\s+([\d.]+)\s*$/);
            if (priceMatch) {
              // 마지막 3개 숫자: Unit Price, Disc Price, Ext Price
              priceInfo.unitPrice = parseAmount(priceMatch[2]); // Disc Price 사용
              priceInfo.extPrice = parseAmount(priceMatch[3]);

              debugLog('OUTRE 가격 정보 추출', {
                unitPrice: priceInfo.unitPrice,
                extPrice: priceInfo.extPrice,
                line: nextLine
              });
            }
          }

          colorLines.push(nextLine);
          searchLog.linesCollected.push({ line: j, type: '컬러 패턴', text: nextLine });
        } else {
          // 매치되지 않은 라인도 기록 (디버깅용)
          if (nextLine.length > 0 && nextLine.length < 100) {
            searchLog.linesFiltered.push({ line: j, reason: '패턴 불일치', text: nextLine });
          }
        }
        }
      } // else 블록 종료 (다음 라인 검색)

      // OUTRE의 경우 추출된 가격 정보 적용 (다음 라인에서 찾았을 경우만)
      if (vendor === 'OUTRE' && priceInfo.unitPrice > 0) {
        unitPrice = priceInfo.unitPrice;
        extPrice = priceInfo.extPrice;
      }

      Logger.log('=== 컬러 라인 검색: ' + itemId + ' ===');
      Logger.log('검색 범위: ' + searchLog.searchRange + '라인');
      Logger.log('확인한 라인 수: ' + searchLog.linesChecked);
      Logger.log('필터링된 라인 수: ' + searchLog.linesFiltered.length);
      Logger.log('수집된 컬러 라인 수: ' + searchLog.linesCollected.length);

      if (searchLog.linesCollected.length > 0) {
        Logger.log('수집된 컬러 라인:');
        for (var logIdx = 0; logIdx < searchLog.linesCollected.length; logIdx++) {
          Logger.log('  [' + searchLog.linesCollected[logIdx].line + '] ' + searchLog.linesCollected[logIdx].text);
        }
      }

      if (searchLog.linesCollected.length === 0 && searchLog.linesFiltered.length > 0) {
        Logger.log('❌ 컬러 라인을 찾을 수 없음. 필터링된 라인들:');
        for (var logIdx = 0; logIdx < Math.min(10, searchLog.linesFiltered.length); logIdx++) {
          Logger.log('  [' + searchLog.linesFiltered[logIdx].line + '] (' + searchLog.linesFiltered[logIdx].reason + ') ' + searchLog.linesFiltered[logIdx].text);
        }
      }

      debugLog('컬러 라인 수집', { count: colorLines.length, lines: colorLines });

      if (colorLines.length > 0) {
        // CRITICAL: parseColorLinesImproved에 원본 Description (cleanup 전)을 전달
        // colorLines에는 원본 Description 텍스트가 포함되어 있기 때문
        var colorData = parseColorLinesImproved(colorLines, descriptionBeforeCleanup || description);

        debugLog('컬러 파싱 결과', { count: colorData.length, data: colorData });

        if (colorData.length > 0) {
          var totalShipped = 0;
          for (var k = 0; k < colorData.length; k++) {
            totalShipped += colorData[k].shipped;
          }

          debugLog('총 shipped 수량', { total: totalShipped, original: qtyShipped });

          for (var k = 0; k < colorData.length; k++) {
            var cd = colorData[k];

            var itemExtPrice = 0;
            if (totalShipped > 0) {
              itemExtPrice = Number((extPrice * (cd.shipped / totalShipped)).toFixed(2));
            }

            // ExtPrice 검증: qtyShipped × unitPrice = extPrice
            var calculatedExtPrice = Number((cd.shipped * unitPrice).toFixed(2));
            var priceDiff = Math.abs(itemExtPrice - calculatedExtPrice);

            var memoText = cd.backordered > 0 ? 'Backordered: ' + cd.backordered : '';

            // 차이가 $0.50 이상이면 메모에 표시
            if (priceDiff >= 0.50) {
              debugLog('⚠️ ExtPrice 불일치', {
                itemId: itemId,
                color: cd.color,
                calculated: calculatedExtPrice,
                parsed: itemExtPrice,
                quantity: cd.shipped,
                unitPrice: unitPrice,
                difference: priceDiff
              });

              if (memoText) {
                memoText += ' | ExtPrice 차이: $' + priceDiff.toFixed(2);
              } else {
                memoText = 'ExtPrice 차이: $' + priceDiff.toFixed(2);
              }
            }

            var item = {
              lineNo: lineNo++,
              itemId: itemId,
              upc: '',
              description: description,
              brand: CONFIG.INVOICE.BRANDS[vendor],
              color: cd.color,
              sizeLength: size,
              qtyOrdered: cd.shipped + cd.backordered,
              qtyShipped: cd.shipped,
              unitPrice: unitPrice,
              extPrice: itemExtPrice,
              memo: memoText
            };

            items.push(item);

            debugLog('아이템 추가', item);
          }

          continue;
        }
      }

      // 컬러 정보가 없으면 경고하고 메모와 함께 추가
      debugLog('경고: 컬러 정보 없음', {
        itemId: itemId,
        description: description,
        qtyShipped: qtyShipped,
        extPrice: extPrice
      });

      // 컬러 정보를 찾을 수 없어도 반드시 리스트에 포함 (메모로 표시)
      if (qtyShipped > 0 || extPrice > 0) {
        // ExtPrice 검증: qtyShipped × unitPrice = extPrice
        var calculatedExtPrice = Number((qtyShipped * unitPrice).toFixed(2));
        var priceDiff = Math.abs(extPrice - calculatedExtPrice);

        var memoText = '⚠️ 컬러 정보 찾을 수 없음';

        if (priceDiff > 0.01) {
          debugLog('⚠️ ExtPrice 불일치 (컬러 없음)', {
            itemId: itemId,
            calculated: calculatedExtPrice,
            parsed: extPrice,
            quantity: qtyShipped,
            unitPrice: unitPrice,
            difference: priceDiff
          });

          // 계산된 값을 사용하고 메모에 표시
          extPrice = calculatedExtPrice;
          memoText += ' | ExtPrice 수정됨';
        }

        var item = {
          lineNo: lineNo++,
          itemId: itemId,
          upc: '',
          description: description,
          brand: CONFIG.INVOICE.BRANDS[vendor],
          color: '',
          sizeLength: size,
          qtyOrdered: qtyOrdered,
          qtyShipped: qtyShipped,
          unitPrice: unitPrice,
          extPrice: extPrice,
          memo: memoText
        };

        items.push(item);

        debugLog('컬러 없는 아이템 추가 (메모 표시)', item);
      }
    }
  }

  debugLog('라인 아이템 파싱 완료', { totalItems: items.length });

  return items;
}

/**
 * 컬러 라인 파싱 (개선 버전)
 * @param {Array} colorLines - 컬러 라인 배열
 * @param {string} description - Description 텍스트 (제외용)
 */
function parseColorLinesImproved(colorLines, description) {
  var colorData = [];

  var fullText = colorLines.join(' ');

  // 언더스코어를 공백으로 변환
  fullText = fullText.replace(/_+/g, ' ');
  fullText = fullText.replace(/\s+/g, ' ').trim();

  debugLog('컬러 라인 전처리', { original: colorLines, processed: fullText });

  // CRITICAL: Description 텍스트가 포함되어 있으면 제거
  // 예: "REMI TARA 1-2-3" → "1-2-3"이 컬러로 인식되는 것을 방지
  // 예: "SUGARPUNCH - 4X4 HD..." → "SUGARPUNCH - 4"가 컬러로 인식되는 것을 방지
  if (description) {
    var descClean = description.trim();

    // 방법 1: 정확히 일치하면 제거 (기존 로직)
    if (fullText.indexOf(descClean) === 0) {
      fullText = fullText.substring(descClean.length).trim();
      debugLog('Description 제거 (정확 매칭)', { removed: descClean, remaining: fullText });
    } else {
      // 방법 2: 단어 기반 매칭 (인코딩 차이 대응)
      // Description의 주요 단어들을 추출 (짧은 단어, 숫자, 따옴표 제외)
      var descWords = descClean.split(/[\s\-]+/).filter(function(word) {
        return word.length > 2 && !word.match(/^\d+$/) && !word.match(/^["″'']+$/);
      });

      if (descWords.length > 0) {
        // fullText에서 Description의 주요 단어들이 순서대로 나타나는지 확인
        var wordsToCheck = descWords.slice(0, Math.min(3, descWords.length));
        var allWordsFound = true;
        var lastIndex = 0;

        for (var i = 0; i < wordsToCheck.length; i++) {
          var wordIndex = fullText.indexOf(wordsToCheck[i], lastIndex);
          if (wordIndex === -1) {
            allWordsFound = false;
            break;
          }
          lastIndex = wordIndex + wordsToCheck[i].length;
        }

        if (allWordsFound) {
          // Description 끝 지점 찾기: 인치 마커 또는 X 패턴까지
          var descEndMatch = fullText.match(/^.+?(\d+["″'']|X)\s*/);
          if (descEndMatch) {
            var removedPart = fullText.substring(0, descEndMatch[0].length);
            fullText = fullText.substring(descEndMatch[0].length).trim();
            debugLog('Description 제거 (단어 기반)', {
              removed: removedPart,
              remaining: fullText,
              matchedWords: wordsToCheck
            });
          }
        }
      }
    }
  }

  // OUTRE의 경우: 마지막에 가격 정보가 있을 수 있으므로 제거
  // 예: "CBRN- 2   JBLK- 0 (2)   NBLK- 1 (1)   NBRN- 2   18.00  17.00  85.00"
  // 마지막 3개 숫자 패턴 제거: \d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}\s*$
  fullText = fullText.replace(/\d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}\s*$/g, '');

  debugLog('가격 제거 후', { processed: fullText });

  // 개선된 정규식: 숫자, 하이픈, 슬래시 뿐만 아니라 알파벳 텍스트도 매치
  // 패턴: [컬러명] - [shipped 수량] 또는 [컬러명] - [shipped 수량] (backorder 수량)
  // 컬러명은 영문자, 숫자, 하이픈, 슬래시 조합 (예: 1, 2, 30, GINGER, BLD-CRUSH, OM27, T30, CBRN, JBLK, NBLK, NBRN)
  var regex = /([A-Z0-9\-\/]+)\s*-\s*(\d+)\s*(?:\((\d+)\))?/gi;
  var match;

  while ((match = regex.exec(fullText)) !== null) {
    var color = match[1].trim();
    var shipped = parseInt(match[2]) || 0;
    var backordered = match[3] ? parseInt(match[3]) : 0;

    debugLog('컬러 매치', {
      color: color,
      shipped: shipped,
      backordered: backordered,
      fullMatch: match[0]
    });

    if (color && color.length > 0 && (shipped > 0 || backordered > 0)) {
      colorData.push({
        color: color,
        shipped: shipped,
        backordered: backordered
      });
    }
  }

  return colorData;
}

/**
 * 금액 파싱
 */
function parseAmount(amountStr) {
  if (!amountStr) return 0;
  var cleanedStr = String(amountStr).replace(/[^\d.]/g, '');
  var amount = parseFloat(cleanedStr);
  return isNaN(amount) ? 0 : amount;
}

/**
 * 날짜 파싱
 */
function parseDate(dateStr) {
  if (!dateStr) return '';
  
  var parts = dateStr.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!parts) return '';
  
  var month = parts[1].length === 1 ? '0' + parts[1] : parts[1];
  var day = parts[2].length === 1 ? '0' + parts[2] : parts[2];
  var year = parts[3];
  
  if (year.length === 2) {
    year = '20' + year;
  }
  
  return year + '-' + month + '-' + day;
}