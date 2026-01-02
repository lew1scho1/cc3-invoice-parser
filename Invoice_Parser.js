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

    if (result.vendor === 'OUTRE' && result.subtotal > 0 && result.discount > 0) {
      var expectedTotal = Number((result.subtotal - result.discount).toFixed(2));
      var subtotalDiff = Math.abs(result.totalAmount - expectedTotal);

      if (subtotalDiff > 1.0) {
        validationMsg += '\n\n⚠️ SUBTOTAL-DISCOUNT 검증 실패\n' +
                         'Expected Total: $' + expectedTotal.toFixed(2) + '\n' +
                         'Parsed Total: $' + result.totalAmount.toFixed(2) + '\n' +
                         '차이: $' + subtotalDiff.toFixed(2);
      } else {
        validationMsg += '\n\n✅ SUBTOTAL-DISCOUNT 검증 통과\n' +
                         'Expected Total: $' + expectedTotal.toFixed(2) + '\n' +
                         'Parsed Total: $' + result.totalAmount.toFixed(2);
      }
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
 * PARSING 시트에 저장 (배치 쓰기로 성능 개선)
 */
function saveToParsingSheet(data) {
  try {
    var sheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);

    // CRITICAL: 배치 쓰기를 위해 모든 행을 배열로 준비
    var rows = [];

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

      rows.push(row);
    }

    // CRITICAL: setValues로 한 번에 쓰기 (appendRow보다 100배 이상 빠름)
    if (rows.length > 0) {
      var lastRow = sheet.getLastRow();
      var targetRange = sheet.getRange(lastRow + 1, 1, rows.length, rows[0].length);
      targetRange.setValues(rows);
    }

    debugLog('PARSING 시트 저장 완료 (배치)', { savedCount: rows.length });

    return rows.length;

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
    
    // 데이터 복사 (성능 개선: 배치 쓰기 사용)
    // moveDataToSheet()는 이미 헤더를 제외하고 배치로 쓰기 때문에
    // data (헤더 포함)를 그대로 전달
    moveDataToSheet(data, targetSheetName);

    var dataRows = data.slice(1); // 로그용
    debugLog('확정 완료 (배치)', {
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

    var summaryTotalFound = false;

    // 2.5 summary block fallback: labels and amounts can be split across lines
    for (var li = 0; li < Math.min(lines.length, 200); li++) {
      var summaryLine = lines[li];
      if (!summaryLine) continue;

      var upperSummary = summaryLine.toUpperCase();
      if (upperSummary.indexOf('SUBTOTAL') > -1) {
        var summaryAmounts = [];
        var hasSummaryLabels = false;
        var scanLimit = Math.min(li + 20, lines.length);

        for (var sj = li; sj < scanLimit; sj++) {
          var blockLine = lines[sj];
          if (!blockLine) continue;

          if (blockLine.match(/TOTAL\s+CARTON|TOTAL\s+LB|AR\s+BALANCE|AGING\s+AS/i)) {
            break;
          }

          if (blockLine.match(/SUBTOTAL|DISCOUNT|TAX|COD|S\s*&\s*H|TOTAL/i)) {
            hasSummaryLabels = true;
          }

          var blockAmounts = blockLine.match(/-?[\d,]+\.\d{2}/g);
          if (blockAmounts) {
            for (var ai = 0; ai < blockAmounts.length; ai++) {
              summaryAmounts.push(parseAmount(blockAmounts[ai]));
            }
          }
        }

        if (hasSummaryLabels && summaryAmounts.length >= 2) {
          var subtotalCandidate = summaryAmounts[0];
          var totalCandidate = summaryAmounts[summaryAmounts.length - 1];

          var pairedAbs = {};
          for (var pi = 0; pi < summaryAmounts.length; pi++) {
            var amt = summaryAmounts[pi];
            var absKey = Math.abs(amt);
            if (!pairedAbs[absKey]) {
              pairedAbs[absKey] = { pos: false, neg: false };
            }
            if (amt > 0) {
              pairedAbs[absKey].pos = true;
            } else if (amt < 0) {
              pairedAbs[absKey].neg = true;
            }
          }

          var pairedValues = {};
          for (var key in pairedAbs) {
            if (pairedAbs[key].pos && pairedAbs[key].neg) {
              pairedValues[key] = true;
            }
          }

          var maxPositive = 0;
          for (var mi = 0; mi < summaryAmounts.length; mi++) {
            var candidate = summaryAmounts[mi];
            if (candidate > 0 &&
                subtotalCandidate > 0 &&
                candidate < subtotalCandidate &&
                !pairedValues[Math.abs(candidate)]) {
              if (candidate > maxPositive) {
                maxPositive = candidate;
              }
            }
          }

          var derivedTotal = 0;
          if (subtotalCandidate > 0 && maxPositive > 0) {
            derivedTotal = Number((subtotalCandidate - maxPositive).toFixed(2));
          }

          if (subtotalCandidate > 0 && data.subtotal === 0) {
            data.subtotal = subtotalCandidate;
            Logger.log('SUBTOTAL parsed from summary block: $' + data.subtotal);
          }

          var totalSelected = totalCandidate;
          var totalIsPaired = pairedValues[Math.abs(totalSelected)] === true;
          var totalTooSmall = subtotalCandidate > 0 &&
            totalSelected > 0 &&
            totalSelected < subtotalCandidate * 0.5;

          if (totalSelected <= 0 || totalIsPaired) {
            totalSelected = 0;
          }

          if (derivedTotal > 0 && (totalSelected === 0 || totalTooSmall)) {
            totalSelected = derivedTotal;
            Logger.log('TOTAL derived from subtotal block: $' + totalSelected);

            if (data.discount === 0 && maxPositive > 0) {
              data.discount = maxPositive;
              Logger.log('DISCOUNT inferred from subtotal block: $' + data.discount);
            }
          }

          if (totalSelected > 0) {
            if (Math.abs(data.totalAmount - totalSelected) > 0.01) {
              data.totalAmount = totalSelected;
              Logger.log('TOTAL parsed from summary block: $' + data.totalAmount);
            }
            summaryTotalFound = true;
          }
        }

        break;
      }
    }

    // 3. 여전히 못 찾으면 SUBTOTAL 근처 찾기 (SUBTOTAL 바로 아래에 TOTAL이 있는 경우)
    if ((!totalMatch || data.totalAmount === 0) && !summaryTotalFound) {
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
 * 라인 아이템 파싱 라우터 (SNG/OUTRE 분기)
 * - SNG → Invoice_Parser_SNG.js의 parseSNGLineItems()
 * - OUTRE → Invoice_Parser_OUTRE.js의 parseOUTRELineItems()
 */
function parseLineItems(lines, vendor) {
  debugLog('라인 아이템 파싱 시작 (라우터)', { vendor: vendor, totalLines: lines.length });

  if (vendor === 'SNG') {
    return parseSNGLineItems(lines);
  } else if (vendor === 'OUTRE') {
    return parseOUTRELineItems(lines);
  } else {
    debugLog('알 수 없는 vendor', { vendor: vendor });
    return [];
  }
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
