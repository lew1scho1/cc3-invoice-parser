// ============================================================================
// DEBUG_SNG.GS - SNG 파싱 디버그
// ============================================================================

/**
 * SNG 파싱 디버그 (DEBUG_OUTPUT 시트 출력)
 * - Invoice_Parser_SNG.js 최신 로직 기준
 */
function debugSNGParsing() {
  DEBUG_LOGS = [];

  try {
    collectDebugLog('=== SNG 파싱 디버그 시작 ===');

    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      collectDebugLog('오류', '폴더가 설정되지 않았습니다.');
      outputDebugLogsToSheet();
      return;
    }

    var fileInfo = pickFirstSNGFile(folderId);
    if (!fileInfo) {
      collectDebugLog('오류', 'SNG 파일이 없습니다.');
      outputDebugLogsToSheet();
      return;
    }

    collectDebugLog('선택된 파일', fileInfo);

    var file = DriveApp.getFileById(fileInfo.id);
    var text = '';
    if (fileInfo.mimeType === MimeType.PDF) {
      text = extractTextFromPdf(file);
    } else {
      text = extractTextFromDocx(file.getBlob());
    }

    collectDebugLog('텍스트 추출 완료', { length: text.length });

    var vendor = detectVendor(text, fileInfo.name);
    collectDebugLog('Vendor 감지', vendor);

    if (vendor !== 'SNG') {
      collectDebugLog('오류', 'SNG 파일이 아닙니다: ' + vendor);
      outputDebugLogsToSheet();
      return;
    }

    var lines = text.split('\n');
    collectDebugLog('총 라인 수', lines.length);

    debugAnalyzeSNG(lines);
  } catch (error) {
    collectDebugLog('에러', error.toString());
    collectDebugLog('스택', error.stack);
  }

  outputDebugLogsToSheet();
}

/**
 * SNG 특정 아이템만 빠르게 디버그 (전체 파싱 방지)
 * - itemId/description/color/qtyShipped만 로그 출력
 */
function debugSNGParsingTopItems() {
  DEBUG_LOGS = [];

  try {
    collectDebugLog('=== SNG 특정 아이템 디버그 시작 ===');

    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      collectDebugLog('오류', '폴더가 설정되지 않았습니다.');
      outputDebugLogsToSheet();
      return;
    }

    var fileInfo = pickFirstSNGFile(folderId);
    if (!fileInfo) {
      collectDebugLog('오류', 'SNG 파일이 없습니다.');
      outputDebugLogsToSheet();
      return;
    }

    collectDebugLog('선택된 파일', fileInfo);

    var file = DriveApp.getFileById(fileInfo.id);
    var text = '';
    if (fileInfo.mimeType === MimeType.PDF) {
      text = extractTextFromPdf(file);
    } else {
      text = extractTextFromDocx(file.getBlob());
    }

    collectDebugLog('텍스트 추출 완료', { length: text.length });

    var vendor = detectVendor(text, fileInfo.name);
    collectDebugLog('Vendor 감지', vendor);

    if (vendor !== 'SNG') {
      collectDebugLog('오류', 'SNG 파일이 아닙니다: ' + vendor);
      outputDebugLogsToSheet();
      return;
    }

    var lines = text.split('\n');
    collectDebugLog('총 라인 수', lines.length);

    // CRITICAL: 문제 아이템 리스트 (Description 복사 문제 대상)
    var targetItems = [
      'SKBTX28',  // 원본 Description
      'SKFCX18',  // ← SKBTX28의 Description 복사
      'SWDSWSX',  // 컬러 파싱 문제
      'SHBN36X',  // ← SHBN34X의 Description 복사
      'SOB4A12',  // ← SOATX30의 Description 복사
      'SOBDX24',  // ← SOBCX24의 Description 복사
      'SOGBBM3',  // 추가 테스트
      'SOHWX18',  // ← SOGBBM3의 Description 복사
      'SOLDX36',  // ← SOLDX24의 Description 복사
      'SGOXX12'   // ← SPRWX24의 Description 복사
    ];

    debugSNGSelectedItems(lines, targetItems);
  } catch (error) {
    collectDebugLog('에러', error.toString());
    collectDebugLog('스택', error.stack);
  }

  outputDebugLogsToSheet();
}

/**
 * 지정된 itemId만 컬러/수량 로그 출력
 * @param {Array<string>} lines
 * @param {Array<string>} targetItemIds
 */
function debugSNGSelectedItems(lines, targetItemIds) {
  var targets = {};
  var found = {};
  for (var t = 0; t < targetItemIds.length; t++) {
    targets[targetItemIds[t]] = true;
  }

  var foundCount = 0;

  for (var i = 0; i < lines.length; i++) {
    var itemInfo = parseSNGItemLine(lines[i]);
    if (!itemInfo) {
      continue;
    }

    if (!targets[itemInfo.itemId]) {
      continue;
    }

    if (!found[itemInfo.itemId]) {
      found[itemInfo.itemId] = true;
      foundCount++;
    }

    collectDebugLog('--- 대상 아이템 (Line ' + i + ') ---', {
      itemId: itemInfo.itemId,
      description: itemInfo.description,
      qtyShipped: itemInfo.qtyShipped
    });

    var priceInfo = parseSNGPriceLine(lines[i + 1]);
    var scanStartIndex = priceInfo ? i + 2 : i + 1;
    var colorLines = [];
    if (priceInfo && priceInfo.colorText) {
      colorLines.push(priceInfo.colorText);
    }

    var colorScan = collectSNGColorLines(scanStartIndex, lines);
    colorLines = colorLines.concat(colorScan.lines);

    var colorData = parseSNGColorLines(colorLines);
    var outputRows = [];

    if (colorData.length > 0) {
      for (var c = 0; c < colorData.length; c++) {
        outputRows.push({
          itemId: itemInfo.itemId,
          description: itemInfo.description,
          color: colorData[c].color,
          shippedQty: colorData[c].shipped
        });
      }
    } else {
      outputRows.push({
        itemId: itemInfo.itemId,
        description: itemInfo.description,
        color: '',
        shippedQty: itemInfo.qtyShipped
      });
    }

    for (var r = 0; r < outputRows.length; r++) {
      collectDebugLog('결과', outputRows[r]);
    }

    i = colorScan.nextIndex - 1;

    if (foundCount >= targetItemIds.length) {
      break;
    }
  }

  collectDebugLog('대상 아이템 찾음', Object.keys(found));
}

/**
 * SNG 디버그 분석
 * @param {Array<string>} lines
 */
function debugAnalyzeSNG(lines) {
  var i = 0;
  var itemCount = 0;
  var mismatches = 0;

  while (i < lines.length) {
    var itemInfo = parseSNGItemLine(lines[i]);
    if (!itemInfo) {
      i++;
      continue;
    }

    itemCount++;
    collectDebugLog('--- 아이템 라인 #' + itemCount + ' (Line ' + i + ') ---');
    collectDebugLog('아이템 정보', itemInfo);
    collectDebugLog('원문', lines[i].trim());

    var priceInfo = parseSNGPriceLine(lines[i + 1]);
    if (priceInfo) {
      collectDebugLog('가격 라인', priceInfo);
      collectDebugLog('가격 원문', lines[i + 1].trim());
    } else {
      collectDebugLog('가격 라인', '없음');
    }

    var scanStartIndex = priceInfo ? i + 2 : i + 1;
    var colorLines = [];
    if (priceInfo && priceInfo.colorText) {
      colorLines.push(priceInfo.colorText);
    }

    var colorScan = collectSNGColorLines(scanStartIndex, lines);
    colorLines = colorLines.concat(colorScan.lines);

    collectDebugLog('컬러 라인 수', colorLines.length);
    if (colorLines.length > 0) {
      collectDebugLog('컬러 라인 샘플', colorLines.slice(0, 3));
    }

    var colorData = parseSNGColorLines(colorLines);
    collectDebugLog('컬러 파싱 결과', colorData);

    if (colorData.length > 0) {
      var totalShipped = 0;
      for (var c = 0; c < colorData.length; c++) {
        totalShipped += colorData[c].shipped;
      }

      if (itemInfo.qtyShipped > 0 && totalShipped !== itemInfo.qtyShipped) {
        mismatches++;
        collectDebugLog('수량 불일치', {
          itemId: itemInfo.itemId,
          itemQtyShipped: itemInfo.qtyShipped,
          colorQtyShipped: totalShipped
        });
      }
    }

    i = colorScan.nextIndex;
  }

  collectDebugLog('아이템 라인 수', itemCount);
  collectDebugLog('수량 불일치 건수', mismatches);

  var parsedItems = parseSNGLineItems(lines);
  collectDebugLog('파싱 결과 라인 수', parsedItems.length);

  var preview = parsedItems.slice(0, 30).map(function(item) {
    return {
      itemId: item.itemId,
      color: item.color,
      qtyShipped: item.qtyShipped,
      unitPrice: item.unitPrice,
      extPrice: item.extPrice
    };
  });

  collectDebugLog('파싱 결과 샘플 (최대 30)', preview);
}

/**
 * 특정 파일명으로 SNG 파싱 디버그 (텍스트 구조 분석)
 * CRITICAL: Apps Script에서 함수 실행 시 파라미터 전달 불가
 * → 파일명을 아래 변수에 직접 설정하고 실행하세요
 */
function debugSNGParsingByFileName() {
  DEBUG_LOGS = [];

  // ========================================
  // 여기에 디버그할 파일명 패턴을 설정하세요
  // ========================================
  var fileName = "3000445999";  // ← 이 값을 변경하세요
  // ========================================

  try {
    collectDebugLog('=== SNG 특정 파일 디버그 ===', { fileName: fileName });

    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      collectDebugLog('오류', '폴더가 설정되지 않았습니다.');
      outputDebugLogsToSheet();
      return;
    }

    var folder = DriveApp.getFolderById(folderId);
    var allFiles = folder.getFiles();

    var file = null;
    while (allFiles.hasNext()) {
      var f = allFiles.next();
      if (f.getName().indexOf(fileName) > -1) {
        file = f;
        break;
      }
    }

    if (!file) {
      collectDebugLog('오류', '파일을 찾을 수 없습니다: ' + fileName);
      outputDebugLogsToSheet();
      return;
    }

    var mimeType = file.getMimeType();

    collectDebugLog('파일 정보', {
      name: file.getName(),
      mimeType: mimeType,
      size: file.getSize()
    });

    var text = '';
    if (mimeType === MimeType.PDF) {
      text = extractTextFromPdf(file);
    } else {
      text = extractTextFromDocx(file.getBlob());
    }

    collectDebugLog('텍스트 추출 완료', { length: text.length });

    var lines = text.split('\n');
    collectDebugLog('총 라인 수', lines.length);

    // 처음 50줄 출력
    collectDebugLog('=== 처음 50줄 ===');
    for (var i = 0; i < Math.min(50, lines.length); i++) {
      collectDebugLog('Line ' + i, lines[i].substring(0, 150));
    }

    // CRITICAL: Line 50-150 추가 출력 (아이템 라인 구조 확인용)
    collectDebugLog('=== Line 50-150 (아이템 구간) ===');
    for (var i = 50; i < Math.min(150, lines.length); i++) {
      collectDebugLog('Line ' + i, lines[i].substring(0, 150));
    }

    // 아이템 라인 감지 테스트
    collectDebugLog('=== 아이템 라인 감지 테스트 ===');
    var detectedCount = 0;
    for (var i = 0; i < Math.min(200, lines.length); i++) {
      var itemInfo = parseSNGItemLine(lines[i]);
      if (itemInfo) {
        detectedCount++;
        collectDebugLog('아이템 감지 (Line ' + i + ')', {
          itemId: itemInfo.itemId,
          description: itemInfo.description,
          qtyShipped: itemInfo.qtyShipped,
          rawLine: lines[i].substring(0, 100)
        });
      }
    }

    collectDebugLog('감지된 아이템 수 (200줄 내)', detectedCount);

    // 전체 파싱 시도
    var parsedItems = parseSNGLineItems(lines);
    collectDebugLog('파싱 결과', { totalItems: parsedItems.length });

    // 결과 샘플 출력
    var preview = parsedItems.slice(0, 10).map(function(item) {
      return {
        itemId: item.itemId,
        description: item.description,
        color: item.color,
        qtyShipped: item.qtyShipped
      };
    });
    collectDebugLog('파싱 결과 샘플 (최대 10)', preview);

  } catch (error) {
    collectDebugLog('에러', error.toString());
    collectDebugLog('스택', error.stack);
  }

  outputDebugLogsToSheet();
}

/**
 * 폴더에서 첫 번째 SNG 파일 선택
 * @param {string} folderId
 * @return {Object|null}
 */
function pickFirstSNGFile(folderId) {
  var folder = DriveApp.getFolderById(folderId);
  var files = folder.getFiles();
  var fileList = [];

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    var mimeType = file.getMimeType();

    if ((mimeType === MimeType.PDF ||
         mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') &&
        (name.indexOf('3000') === 0 || name.match(/\d{10}/))) {
      fileList.push({
        id: file.getId(),
        name: name,
        mimeType: mimeType
      });
    }
  }

  if (fileList.length === 0) {
    return null;
  }

  fileList.sort(function(a, b) {
    return a.name.localeCompare(b.name);
  });

  return fileList[0];
}
