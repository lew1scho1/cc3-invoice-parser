// ============================================================================
// DOCUMENTAI.JS - Google Cloud Document AI 통합
// ============================================================================

/**
 * Document AI로 인보이스 PDF 파싱
 * @param {File} file - Google Drive 파일 객체
 * @return {Object} Document AI 응답
 */
function parseInvoiceWithDocumentAI(file) {
  try {
    debugLog('Document AI 파싱 시작', { fileName: file.getName() });

    // 1. 설정 가져오기
    var props = PropertiesService.getScriptProperties();
    var projectId = props.getProperty('DOCUMENT_AI_PROJECT_ID');
    var location = props.getProperty('DOCUMENT_AI_LOCATION');
    var processorId = props.getProperty('DOCUMENT_AI_PROCESSOR_ID');

    if (!projectId || !location || !processorId) {
      throw new Error('Document AI 설정이 누락되었습니다. 스크립트 속성을 확인하세요.');
    }

    debugLog('Document AI 설정', { projectId: projectId, location: location, processorId: processorId });

    // 2. OAuth 토큰 얻기
    var token = getDocumentAIAccessToken();

    debugLog('OAuth 토큰 획득 완료');

    // 3. PDF를 Base64로 인코딩
    var blob = file.getBlob();
    var bytes = blob.getBytes();
    var base64 = Utilities.base64Encode(bytes);

    debugLog('PDF Base64 인코딩 완료', { size: bytes.length });

    // 4. Document AI API 호출
    var url = 'https://' + location + '-documentai.googleapis.com/v1/projects/' +
              projectId + '/locations/' + location + '/processors/' +
              processorId + ':process';

    var payload = {
      rawDocument: {
        content: base64,
        mimeType: file.getMimeType()
      }
    };

    var options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    debugLog('Document AI API 호출 시작', { url: url });

    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();

    debugLog('Document AI API 응답', { code: responseCode });

    if (responseCode !== 200) {
      var errorText = response.getContentText();
      debugLog('Document AI API 오류', { code: responseCode, error: errorText });
      throw new Error('Document AI API 오류 (' + responseCode + '): ' + errorText);
    }

    var result = JSON.parse(response.getContentText());

    debugLog('Document AI 파싱 완료');

    return result;

  } catch (error) {
    debugLog('parseInvoiceWithDocumentAI 오류', { error: error.toString() });
    logError('parseInvoiceWithDocumentAI', error, { fileName: file.getName() });
    throw error;
  }
}

/**
 * OAuth2 액세스 토큰 얻기
 * @return {string} 액세스 토큰
 */
function getDocumentAIAccessToken() {
  try {
    var props = PropertiesService.getScriptProperties();
    var serviceAccountJson = props.getProperty('DOCUMENT_AI_SERVICE_ACCOUNT');

    if (!serviceAccountJson) {
      throw new Error('서비스 계정 키가 설정되지 않았습니다.');
    }

    var serviceAccount = JSON.parse(serviceAccountJson);

    // OAuth2 라이브러리 사용
    var service = OAuth2.createService('DocumentAI')
      .setTokenUrl('https://oauth2.googleapis.com/token')
      .setPrivateKey(serviceAccount.private_key)
      .setIssuer(serviceAccount.client_email)
      .setPropertyStore(PropertiesService.getScriptProperties())
      .setCache(CacheService.getScriptCache())
      .setScope('https://www.googleapis.com/auth/cloud-platform');

    if (!service.hasAccess()) {
      debugLog('OAuth2 액세스 없음, 새로 인증 시도');
      service.reset();
    }

    var token = service.getAccessToken();

    if (!token) {
      throw new Error('OAuth2 토큰을 얻을 수 없습니다.');
    }

    return token;

  } catch (error) {
    debugLog('getDocumentAIAccessToken 오류', { error: error.toString() });
    throw new Error('OAuth2 토큰 획득 실패: ' + error.toString());
  }
}

/**
 * Document AI 응답에서 엔티티 추출
 * @param {Object} aiResult - Document AI 응답
 * @param {string} entityType - 엔티티 타입 (예: 'invoice_id', 'invoice_date')
 * @return {string} 추출된 값
 */
function extractEntity(aiResult, entityType) {
  try {
    if (!aiResult || !aiResult.document || !aiResult.document.entities) {
      return '';
    }

    var entities = aiResult.document.entities;

    for (var i = 0; i < entities.length; i++) {
      if (entities[i].type === entityType) {
        return entities[i].mentionText || '';
      }
    }

    return '';

  } catch (error) {
    debugLog('extractEntity 오류', { entityType: entityType, error: error.toString() });
    return '';
  }
}

/**
 * Document AI 응답에서 특정 타입의 모든 엔티티 추출
 * @param {Object} aiResult - Document AI 응답
 * @param {string} entityType - 엔티티 타입
 * @return {Array} 엔티티 배열
 */
function extractEntities(aiResult, entityType) {
  try {
    if (!aiResult || !aiResult.document || !aiResult.document.entities) {
      return [];
    }

    var entities = aiResult.document.entities;
    var results = [];

    for (var i = 0; i < entities.length; i++) {
      if (entities[i].type === entityType) {
        results.push(entities[i]);
      }
    }

    return results;

  } catch (error) {
    debugLog('extractEntities 오류', { entityType: entityType, error: error.toString() });
    return [];
  }
}

/**
 * 엔티티의 속성 값 가져오기
 * @param {Object} entity - 엔티티 객체
 * @param {string} propertyType - 속성 타입 (예: 'line_item/description')
 * @return {string} 속성 값
 */
function getEntityProperty(entity, propertyType) {
  try {
    if (!entity || !entity.properties) {
      return '';
    }

    for (var i = 0; i < entity.properties.length; i++) {
      if (entity.properties[i].type === propertyType) {
        return entity.properties[i].mentionText || '';
      }
    }

    return '';

  } catch (error) {
    debugLog('getEntityProperty 오류', { propertyType: propertyType, error: error.toString() });
    return '';
  }
}

/**
 * Document AI 응답을 현재 인보이스 데이터 구조로 변환
 * @param {Object} aiResult - Document AI 응답
 * @param {string} filename - 파일명
 * @return {Object} 파싱된 인보이스 데이터
 */
function convertDocumentAIToInvoiceData(aiResult, filename) {
  try {
    debugLog('Document AI 응답 변환 시작', { filename: filename });

    // 전체 텍스트 추출
    var fullText = '';
    if (aiResult && aiResult.document && aiResult.document.text) {
      fullText = aiResult.document.text;
    }
    var allLines = fullText.split('\n');

    // Vendor 감지 (파일명 또는 invoice_id로 판단)
    var invoiceId = extractEntity(aiResult, 'invoice_id');
    var vendor = 'UNKNOWN';

    if (filename.indexOf('3000') === 0 || filename.match(/\d{10}/)) {
      vendor = 'SNG';
      // SNG는 파일명에서 Invoice Number 추출
      var invoiceMatch = filename.match(/(\d{10})/);
      if (invoiceMatch) {
        invoiceId = invoiceMatch[1];
      }
    } else if (filename.indexOf('SINV') > -1 || invoiceId.indexOf('SINV') === 0) {
      vendor = 'OUTRE';
    }

    debugLog('Vendor 감지', { vendor: vendor, invoiceId: invoiceId });

    // 헤더 정보 추출
    var data = {
      vendor: vendor,
      filename: filename,
      invoiceNo: invoiceId,
      invoiceDate: parseDate(extractEntity(aiResult, 'invoice_date')),
      totalAmount: 0,
      subtotal: 0,
      discount: 0,
      shipping: 0,
      tax: 0,
      lineItems: []
    };

    // Vendor별 헤더 파싱
    if (vendor === 'SNG') {
      // SNG: Invoice Amount 찾기
      var invoiceAmountPattern = /INVOICE\s+AMOUNT/gi;
      var invoiceAmountPositions = [];
      var match;

      while ((match = invoiceAmountPattern.exec(fullText)) !== null) {
        invoiceAmountPositions.push({
          text: match[0],
          index: match.index
        });
      }

      if (invoiceAmountPositions.length > 0) {
        var lastInvoiceAmount = invoiceAmountPositions[invoiceAmountPositions.length - 1];
        var searchStart = lastInvoiceAmount.index;
        var searchEnd = Math.min(searchStart + 200, fullText.length);
        var searchText = fullText.substring(searchStart, searchEnd);

        var amountPattern = /(\d{1,3}(?:,\d{3})*\.\d{2})/g;
        var amounts = [];

        while ((match = amountPattern.exec(searchText)) !== null) {
          var amount = parseFloat(match[1].replace(/,/g, ''));
          if (amount >= 49.99 && amount <= 100000.00) {
            amounts.push(amount);
          }
        }

        if (amounts.length > 0) {
          data.totalAmount = amounts[amounts.length - 1];
        }
      }
    } else if (vendor === 'OUTRE') {
      // OUTRE: TOTAL 찾기
      var totalMatch = fullText.match(/\bTOTAL\s+(?:US\$)?\s*([\d,\.]+)/i);
      if (totalMatch) {
        data.totalAmount = parseFloat(totalMatch[1].replace(/,/g, '')) || 0;
      }
    }

    debugLog('헤더 정보 추출', {
      invoiceNo: data.invoiceNo,
      invoiceDate: data.invoiceDate,
      totalAmount: data.totalAmount
    });

    // 라인 아이템 추출
    var lineItemEntities = extractEntities(aiResult, 'line_item');

    debugLog('라인 아이템 개수', { count: lineItemEntities.length });

    for (var i = 0; i < lineItemEntities.length; i++) {
      var entity = lineItemEntities[i];

      var description = getEntityProperty(entity, 'line_item/description');
      var quantity = getEntityProperty(entity, 'line_item/quantity');
      var productCode = getEntityProperty(entity, 'line_item/product_code');

      // Vendor별 가격 필드 추출
      var prices = getVendorSpecificPrices(entity, vendor);
      var unitPrice = prices.unitPrice;
      var amount = prices.amount;

      // Size 추출 (Description에서)
      var sizeMatch = description.match(/(\d+)["″'']/);
      var size = sizeMatch ? sizeMatch[1] + '"' : '';

      // Description에서 color line이 포함되어 있는지 확인하고 제거
      // OUTRE의 경우 description이 "BIG BEAUTIFUL HAIR...\nCBRN-2 JBLK-0 (2)..." 처럼 올 수 있음
      var descriptionLines = description.split('\n');
      var cleanDescription = descriptionLines[0].trim(); // 첫 번째 라인만 description으로

      // 나머지 라인들은 color line으로 처리
      var colorLinesFromDesc = [];
      for (var j = 1; j < descriptionLines.length; j++) {
        if (descriptionLines[j].trim()) {
          colorLinesFromDesc.push(descriptionLines[j].trim());
        }
      }

      debugLog('Description 분리', {
        original: description,
        cleanDescription: cleanDescription,
        colorLinesFromDesc: colorLinesFromDesc
      });

      // 만약 color line이 description에 없다면, 전체 텍스트에서 찾기
      if (colorLinesFromDesc.length === 0 && fullText) {
        // Description 위치 찾기
        var descIndex = fullText.indexOf(cleanDescription);
        if (descIndex > -1) {
          // Description 이후 50줄 내에서 color line 찾기
          var startLineIdx = -1;
          for (var lineIdx = 0; lineIdx < allLines.length; lineIdx++) {
            if (allLines[lineIdx].indexOf(cleanDescription) > -1) {
              startLineIdx = lineIdx;
              break;
            }
          }

          if (startLineIdx > -1) {
            for (var lineIdx = startLineIdx + 1; lineIdx < Math.min(startLineIdx + 50, allLines.length); lineIdx++) {
              var line = allLines[lineIdx].trim();

              // 다음 line item을 만나면 중단
              if (line.match(/^\d+\s+[A-Z]/)) {
                break;
              }

              // Color line 패턴 체크
              if (line.indexOf('_') > -1 || line.match(/[A-Z0-9\-\/]+\s*-\s*\d+/)) {
                colorLinesFromDesc.push(line);
              }
            }
          }
        }
      }

      debugLog('최종 color lines', {
        itemId: productCode,
        count: colorLinesFromDesc.length,
        lines: colorLinesFromDesc
      });

      // color line 파싱 (기존 로직 사용)
      var colorData = [];
      if (colorLinesFromDesc.length > 0) {
        colorData = parseColorLinesImproved(colorLinesFromDesc);
      }

      debugLog('Color 파싱 결과', {
        itemId: productCode,
        colorCount: colorData.length,
        colors: colorData
      });

      // Color가 있으면 각 color별로 line item 생성
      if (colorData.length > 0) {
        var totalShipped = 0;
        for (var k = 0; k < colorData.length; k++) {
          totalShipped += colorData[k].shipped;
        }

        for (var k = 0; k < colorData.length; k++) {
          var cd = colorData[k];

          var itemExtPrice = 0;
          if (totalShipped > 0) {
            itemExtPrice = Number((parseFloat(amount) * (cd.shipped / totalShipped)).toFixed(2));
          }

          var item = {
            lineNo: data.lineItems.length + 1,
            itemId: productCode || '',
            upc: '',
            description: cleanDescription,
            brand: CONFIG.INVOICE.BRANDS[vendor],
            color: cd.color,
            sizeLength: size,
            qtyOrdered: cd.shipped + cd.backordered,
            qtyShipped: cd.shipped,
            unitPrice: parseFloat(unitPrice) || 0,
            extPrice: itemExtPrice,
            memo: cd.backordered > 0 ? 'Backordered: ' + cd.backordered : ''
          };

          data.lineItems.push(item);

          debugLog('Color별 라인 아이템 추가', {
            lineNo: item.lineNo,
            itemId: item.itemId,
            color: item.color,
            description: item.description.substring(0, 50),
            quantity: item.qtyShipped,
            extPrice: item.extPrice
          });
        }
      } else {
        // Color가 없으면 그냥 하나의 item으로
        var item = {
          lineNo: data.lineItems.length + 1,
          itemId: productCode || '',
          upc: '',
          description: cleanDescription,
          brand: CONFIG.INVOICE.BRANDS[vendor],
          color: '',
          sizeLength: size,
          qtyOrdered: parseInt(quantity) || 0,
          qtyShipped: parseInt(quantity) || 0,
          unitPrice: parseFloat(unitPrice) || 0,
          extPrice: parseFloat(amount) || 0,
          memo: '⚠️ 컬러 정보 찾을 수 없음'
        };

        data.lineItems.push(item);

        debugLog('라인 아이템 추가 (컬러 없음)', {
          lineNo: item.lineNo,
          itemId: item.itemId,
          description: item.description.substring(0, 50),
          quantity: item.qtyShipped,
          extPrice: item.extPrice
        });
      }
    }

    debugLog('Document AI 응답 변환 완료', { lineItems: data.lineItems.length });

    return data;

  } catch (error) {
    debugLog('convertDocumentAIToInvoiceData 오류', { error: error.toString() });
    throw error;
  }
}

/**
 * 컬러 라인 파싱 (개선 버전)
 * Invoice_Parser.js의 parseColorLinesImproved()와 동일
 */
function parseColorLinesImproved(colorLines) {
  var colorData = [];

  var fullText = colorLines.join(' ');

  // 언더스코어를 공백으로 변환
  fullText = fullText.replace(/_+/g, ' ');
  fullText = fullText.replace(/\s+/g, ' ').trim();

  debugLog('컬러 라인 전처리', { original: colorLines, processed: fullText });

  // 개선된 정규식: 숫자, 하이픈, 슬래시 뿐만 아니라 알파벳 텍스트도 매치
  // 패턴: [컬러명] - [shipped 수량] 또는 [컬러명] - [shipped 수량] (backorder 수량)
  // 컬러명은 영문자, 숫자, 하이픈, 슬래시 조합 (예: 1, 2, 30, GINGER, BLD-CRUSH, OM27, T30)
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
 * Document AI 응답 디버깅 함수
 * PARSING 탭의 첫 번째 파일로 테스트
 */
function debugDocumentAIResponse() {
  try {
    Logger.log('=== Document AI 응답 디버깅 ===');

    // PARSING 탭에서 파일명 가져오기
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var parsingSheet = ss.getSheetByName(CONFIG.INVOICE.PARSING_SHEET);

    if (!parsingSheet) {
      Logger.log('❌ PARSING 시트를 찾을 수 없습니다.');
      return;
    }

    var data = parsingSheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('❌ PARSING 탭에 데이터가 없습니다. 먼저 파싱을 실행하세요.');
      return;
    }

    var filename = data[1][1]; // INVOICE_NO 컬럼에서 파일명 가져오기
    Logger.log('파일명: ' + filename);

    // 파일 찾기
    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFilesByName(filename);

    if (!files.hasNext()) {
      Logger.log('❌ 파일을 찾을 수 없습니다: ' + filename);
      return;
    }

    var file = files.next();
    Logger.log('파일 발견: ' + file.getName());

    // Document AI 호출
    var aiResult = parseInvoiceWithDocumentAI(file);

    Logger.log('\n=== Document AI 원본 응답 ===');
    Logger.log('Invoice ID: ' + extractEntity(aiResult, 'invoice_id'));
    Logger.log('Invoice Date: ' + extractEntity(aiResult, 'invoice_date'));
    Logger.log('Total Amount: ' + extractEntity(aiResult, 'total_amount'));
    Logger.log('Net Amount: ' + extractEntity(aiResult, 'net_amount'));

    // Line items 상세 정보
    var lineItems = extractEntities(aiResult, 'line_item');
    Logger.log('\nLine Items 개수: ' + lineItems.length);

    for (var i = 0; i < Math.min(5, lineItems.length); i++) {
      var item = lineItems[i];
      Logger.log('\n--- Line Item ' + (i + 1) + ' ---');

      // 모든 properties 출력 (가능한 모든 price 필드 확인)
      if (item.properties) {
        Logger.log('All Properties (' + item.properties.length + ' total):');
        for (var j = 0; j < item.properties.length; j++) {
          var prop = item.properties[j];
          Logger.log('  [' + j + '] ' + prop.type + ': ' + (prop.mentionText || ''));
        }
      } else {
        Logger.log('No properties found');
      }

      // 일반적인 필드들
      Logger.log('\nCommon Fields:');
      Logger.log('  description: ' + getEntityProperty(item, 'line_item/description'));
      Logger.log('  product_code: ' + getEntityProperty(item, 'line_item/product_code'));
      Logger.log('  quantity: ' + getEntityProperty(item, 'line_item/quantity'));
      Logger.log('  unit_price: ' + getEntityProperty(item, 'line_item/unit_price'));
      Logger.log('  amount: ' + getEntityProperty(item, 'line_item/amount'));
    }

    // 전체 텍스트의 일부 출력 (color line 파싱 확인용)
    if (aiResult.document && aiResult.document.text) {
      var fullText = aiResult.document.text;
      Logger.log('\n=== Full Text Sample (first 1000 chars) ===');
      Logger.log(fullText.substring(0, 1000));

      Logger.log('\n=== Full Text Sample (last 1000 chars) ===');
      Logger.log(fullText.substring(Math.max(0, fullText.length - 1000)));
    }

    Logger.log('\n✅ 디버깅 완료');

  } catch (error) {
    Logger.log('❌ 오류: ' + error.toString());
    Logger.log(error.stack);
  }
}

/**
 * Document AI 테스트 함수
 */
function testDocumentAI() {
  try {
    // 설정 확인
    var props = PropertiesService.getScriptProperties();
    Logger.log('=== Document AI 설정 확인 ===');
    Logger.log('PROJECT_ID: ' + props.getProperty('DOCUMENT_AI_PROJECT_ID'));
    Logger.log('LOCATION: ' + props.getProperty('DOCUMENT_AI_LOCATION'));
    Logger.log('PROCESSOR_ID: ' + props.getProperty('DOCUMENT_AI_PROCESSOR_ID'));
    Logger.log('SERVICE_ACCOUNT: ' + (props.getProperty('DOCUMENT_AI_SERVICE_ACCOUNT') ? '설정됨' : '미설정'));

    // OAuth 토큰 테스트
    Logger.log('\n=== OAuth 토큰 테스트 ===');
    var token = getDocumentAIAccessToken();
    Logger.log('토큰 획득 성공: ' + token.substring(0, 20) + '...');

    Logger.log('\n✅ Document AI 설정 완료!');

  } catch (error) {
    Logger.log('❌ 오류: ' + error.toString());
    Logger.log(error.stack);
  }
}

/**
 * Document AI 응답을 JSON 파일과 Excel로 저장
 * 폴더에서 파일을 선택하여 Document AI 호출 후 저장
 */
function saveDocumentAIResponseToFiles() {
  try {
    var ui = SpreadsheetApp.getUi();
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 폴더 ID 가져오기
    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      ui.alert('오류', '먼저 인보이스 폴더를 설정해주세요.\n\n메뉴: CC3 ORDER APP > 📄 인보이스 > 📁 폴더 설정', ui.ButtonSet.OK);
      return;
    }

    // 폴더에서 PDF 파일 목록 가져오기
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();

    var fileList = [];
    while (files.hasNext()) {
      var file = files.next();
      var mimeType = file.getMimeType();

      // PDF만 선택
      if (mimeType === MimeType.PDF) {
        fileList.push({
          id: file.getId(),
          name: file.getName(),
          date: file.getDateCreated()
        });
      }
    }

    if (fileList.length === 0) {
      ui.alert('오류', '폴더에 PDF 파일이 없습니다.', ui.ButtonSet.OK);
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
      '분석할 PDF 파일 번호를 입력하세요:\n\n' + fileNames,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      return;
    }

    var input = response.getResponseText().trim();
    var fileIndex = parseInt(input) - 1;

    if (isNaN(fileIndex) || fileIndex < 0 || fileIndex >= fileList.length) {
      ui.alert('오류', '올바른 번호를 입력해주세요.', ui.ButtonSet.OK);
      return;
    }

    var selectedFile = fileList[fileIndex];
    var file = DriveApp.getFileById(selectedFile.id);
    var filename = file.getName();

    ss.toast('Document AI 호출 중...', '분석 중', -1);

    // Document AI 호출
    var aiResult = parseInvoiceWithDocumentAI(file);

    ss.toast('결과 저장 중...', '분석 중', -1);

    // 1. JSON 파일로 저장
    var jsonFilename = filename.replace(/\.(pdf|docx)$/i, '') + '_DocumentAI.json';
    var jsonContent = JSON.stringify(aiResult, null, 2);
    var jsonBlob = Utilities.newBlob(jsonContent, 'application/json', jsonFilename);
    folder.createFile(jsonBlob);

    // 2. Excel 시트에 원본 PDF 레이아웃 그대로 재구성
    var sheetName = 'DocumentAI_Invoice';
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }

    // Document AI의 모든 텍스트를 위치 정보와 함께 추출
    var textElements = extractAllTextWithPositions(aiResult);

    debugLog('텍스트 요소 개수', { count: textElements.length });

    // 위치 기반으로 Excel 셀에 배치
    if (textElements.length > 0) {
      // Y 좌표 범위 계산 (페이지 높이)
      var minY = Infinity;
      var maxY = -Infinity;
      var minX = Infinity;
      var maxX = -Infinity;

      for (var i = 0; i < textElements.length; i++) {
        var elem = textElements[i];
        if (elem.y < minY) minY = elem.y;
        if (elem.y > maxY) maxY = elem.y;
        if (elem.x < minX) minX = elem.x;
        if (elem.x > maxX) maxX = elem.x;
      }

      debugLog('좌표 범위', { minX: minX, maxX: maxX, minY: minY, maxY: maxY });

      // Y 좌표를 행 번호로 변환 (픽셀 → 행)
      // 대략 15-20 픽셀당 1행으로 추정
      var pixelsPerRow = 15;
      var pixelsPerCol = 8;

      // 각 텍스트 요소를 적절한 셀에 배치
      for (var i = 0; i < textElements.length; i++) {
        var elem = textElements[i];

        // 좌표를 행/열로 변환
        var row = Math.floor((elem.y - minY) / pixelsPerRow) + 1;
        var col = Math.floor((elem.x - minX) / pixelsPerCol) + 1;

        // 범위 제한 (Google Sheets 최대값)
        if (row > 1000) row = 1000;
        if (col > 26) col = 26; // A-Z

        try {
          var currentValue = sheet.getRange(row, col).getValue();

          // 이미 값이 있으면 옆 셀에 배치
          if (currentValue) {
            col++;
            if (col > 26) continue; // 너무 오른쪽이면 스킵
          }

          sheet.getRange(row, col).setValue(elem.text);

          // 폰트 크기 적용 (추정)
          if (elem.fontSize) {
            sheet.getRange(row, col).setFontSize(elem.fontSize);
          }

        } catch (e) {
          debugLog('셀 배치 오류', { row: row, col: col, error: e.toString() });
        }
      }

      // 컬럼 너비 조정 (대략적으로)
      for (var col = 1; col <= 26; col++) {
        sheet.setColumnWidth(col, 100);
      }
    }

    ss.toast('', '', 1);

    var message = '✅ Document AI 분석 결과 저장 완료!\n\n' +
                  '1. JSON 파일: ' + jsonFilename + '\n' +
                  '   (Drive 폴더에 저장됨)\n\n' +
                  '2. Excel 시트: "' + sheetName + '"\n' +
                  '   (원본 PDF 레이아웃으로 재구성)\n\n' +
                  '텍스트 요소: ' + textElements.length + '개 배치';

    ui.alert('저장 완료', message, ui.ButtonSet.OK);

    // 인보이스 시트로 이동
    ss.setActiveSheet(sheet);

  } catch (error) {
    SpreadsheetApp.getUi().alert('오류', '저장 중 오류 발생:\n' + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    Logger.log('saveDocumentAIResponseToFiles 오류: ' + error.toString());
    Logger.log(error.stack);
  }
}

/**
 * Document AI 응답에서 모든 텍스트와 위치 정보 추출
 * @param {Object} aiResult - Document AI 응답
 * @return {Array} 텍스트 요소 배열 [{text, x, y, fontSize}]
 */
function extractAllTextWithPositions(aiResult) {
  var elements = [];

  try {
    if (!aiResult || !aiResult.document || !aiResult.document.pages) {
      debugLog('페이지 정보 없음');
      return elements;
    }

    var pages = aiResult.document.pages;

    // 각 페이지 처리
    for (var pageIdx = 0; pageIdx < pages.length; pageIdx++) {
      var page = pages[pageIdx];

      if (!page.tokens) {
        debugLog('페이지 토큰 없음', { pageIdx: pageIdx });
        continue;
      }

      // 각 토큰(단어) 처리
      for (var tokenIdx = 0; tokenIdx < page.tokens.length; tokenIdx++) {
        var token = page.tokens[tokenIdx];

        if (!token.layout || !token.layout.boundingPoly) {
          continue;
        }

        // 텍스트 추출
        var text = '';
        if (token.layout.textAnchor && token.layout.textAnchor.textSegments) {
          var segment = token.layout.textAnchor.textSegments[0];
          if (segment && aiResult.document.text) {
            var startIdx = parseInt(segment.startIndex) || 0;
            var endIdx = parseInt(segment.endIndex) || startIdx;
            text = aiResult.document.text.substring(startIdx, endIdx);
          }
        }

        if (!text) continue;

        // Bounding box에서 좌표 추출
        var vertices = token.layout.boundingPoly.normalizedVertices || token.layout.boundingPoly.vertices;

        if (!vertices || vertices.length === 0) {
          continue;
        }

        // 왼쪽 상단 좌표 사용
        var topLeft = vertices[0];
        var x = topLeft.x || 0;
        var y = topLeft.y || 0;

        // Normalized 좌표 (0-1 범위)인 경우 픽셀로 변환
        if (x < 1 && y < 1) {
          // A4 페이지 크기 가정: 595 x 842 pt
          x = x * 595;
          y = y * 842;
        }

        // 폰트 크기 추정 (bounding box 높이 사용)
        var fontSize = 10; // 기본값
        if (vertices.length >= 3) {
          var height = Math.abs((vertices[2].y || 0) - (vertices[0].y || 0));
          if (height < 1) {
            height = height * 842; // normalized인 경우
          }
          fontSize = Math.max(8, Math.min(18, Math.round(height)));
        }

        elements.push({
          text: text,
          x: x,
          y: y,
          fontSize: fontSize
        });
      }
    }

    debugLog('위치 정보 추출 완료', { count: elements.length });

  } catch (error) {
    debugLog('extractAllTextWithPositions 오류', { error: error.toString() });
  }

  return elements;
}

/**
 * Vendor별 price 필드 찾기 헬퍼
 * @param {Object} entity - line_item 엔티티
 * @param {string} vendor - 'SNG' 또는 'OUTRE'
 * @return {Object} { unitPrice, amount } - 파싱된 가격 정보
 */
function getVendorSpecificPrices(entity, vendor) {
  var result = {
    unitPrice: 0,
    amount: 0
  };

  if (!entity || !entity.properties) {
    return result;
  }

  if (vendor === 'SNG') {
    // SNG: "your price" 또는 "your_price" 찾기
    var yourPrice = '';
    var yourExtended = '';

    for (var i = 0; i < entity.properties.length; i++) {
      var prop = entity.properties[i];
      var propType = prop.type.toLowerCase();

      // "your price" 또는 유사한 패턴 찾기
      if (propType.indexOf('your') > -1 && propType.indexOf('price') > -1 && propType.indexOf('extended') === -1) {
        yourPrice = prop.mentionText || '';
      }
      // "your extended" 또는 유사한 패턴 찾기
      if (propType.indexOf('your') > -1 && propType.indexOf('extended') > -1) {
        yourExtended = prop.mentionText || '';
      }
    }

    // fallback: 일반 unit_price, amount 사용
    if (!yourPrice) {
      yourPrice = getEntityProperty(entity, 'line_item/unit_price');
    }
    if (!yourExtended) {
      yourExtended = getEntityProperty(entity, 'line_item/amount');
    }

    result.unitPrice = parseFloat(String(yourPrice).replace(/[,$]/g, '')) || 0;
    result.amount = parseFloat(String(yourExtended).replace(/[,$]/g, '')) || 0;

  } else if (vendor === 'OUTRE') {
    // OUTRE: "disc price" 또는 "disc_price" 찾기
    var discPrice = '';
    var amount = '';

    for (var i = 0; i < entity.properties.length; i++) {
      var prop = entity.properties[i];
      var propType = prop.type.toLowerCase();

      // "disc price" 또는 유사한 패턴 찾기
      if (propType.indexOf('disc') > -1 && propType.indexOf('price') > -1) {
        discPrice = prop.mentionText || '';
      }
      // amount는 일반 필드 사용
      if (propType === 'line_item/amount') {
        amount = prop.mentionText || '';
      }
    }

    // fallback: 일반 unit_price, amount 사용
    if (!discPrice) {
      discPrice = getEntityProperty(entity, 'line_item/unit_price');
    }
    if (!amount) {
      amount = getEntityProperty(entity, 'line_item/amount');
    }

    result.unitPrice = parseFloat(String(discPrice).replace(/[,$]/g, '')) || 0;
    result.amount = parseFloat(String(amount).replace(/[,$]/g, '')) || 0;

  } else {
    // Unknown vendor: 일반 필드 사용
    var unitPrice = getEntityProperty(entity, 'line_item/unit_price');
    var amount = getEntityProperty(entity, 'line_item/amount');

    result.unitPrice = parseFloat(String(unitPrice).replace(/[,$]/g, '')) || 0;
    result.amount = parseFloat(String(amount).replace(/[,$]/g, '')) || 0;
  }

  return result;
}
