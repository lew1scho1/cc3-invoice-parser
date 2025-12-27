// ============================================================================
// DEBUG.GS - 디버깅 전용 함수
// ============================================================================

/**
 * OUTRE 인보이스 파싱 디버그 (특정 제품만)
 * REMI TARA, SUGARPUNCH 제품의 파싱 과정을 상세히 추적
 */
function debugOutreParsingIssues() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 결과를 저장할 시트 생성/초기화
    var debugSheetName = 'DEBUG_LOG';
    var debugSheet = ss.getSheetByName(debugSheetName);

    if (debugSheet) {
      ss.deleteSheet(debugSheet);
    }

    debugSheet = ss.insertSheet(debugSheetName);
    debugSheet.appendRow(['DEBUG LOG']);
    debugSheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);

    var logRow = 2;

    function log(message) {
      debugSheet.getRange(logRow, 1).setValue(message);
      logRow++;
      Logger.log(message);
    }

    log('='.repeat(80));
    log('🔍 OUTRE 인보이스 파싱 디버그 시작');
    log('='.repeat(80));
    log('');

    // 폴더에서 OUTRE 인보이스 파일 찾기
    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      log('❌ 오류: 인보이스 폴더가 설정되지 않았습니다.');
      return;
    }

    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var outreFile = null;

    while (files.hasNext()) {
      var file = files.next();
      var filename = file.getName();

      // SINV로 시작하는 파일 찾기
      if (filename.indexOf('SINV') > -1) {
        outreFile = file;
        log('✅ OUTRE 인보이스 파일 발견: ' + filename);
        break;
      }
    }

    if (!outreFile) {
      log('❌ 오류: OUTRE 인보이스 파일(SINV...)을 찾을 수 없습니다.');
      return;
    }

    log('');
    log('📄 파일 정보:');
    log('  이름: ' + outreFile.getName());
    log('  MIME 타입: ' + outreFile.getMimeType());
    log('');

    // 텍스트 추출
    log('📝 텍스트 추출 중...');
    var text = '';

    if (outreFile.getMimeType() === MimeType.PDF) {
      text = extractTextFromPdf(outreFile);
    } else {
      text = extractTextFromDocx(outreFile.getBlob());
    }

    log('  추출된 텍스트 길이: ' + text.length + ' 문자');
    log('');

    // 라인으로 분할
    var lines = text.split('\n');
    log('  총 라인 수: ' + lines.length);
    log('');

    // REMI TARA 또는 SUGARPUNCH 제품 찾기
    log('🔍 문제 제품 검색 중...');
    log('');

    var targetProducts = ['REMI TARA', 'SUGARPUNCH'];
    var foundProducts = [];

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();

      for (var tp = 0; tp < targetProducts.length; tp++) {
        if (line.indexOf(targetProducts[tp]) > -1) {
          foundProducts.push({
            product: targetProducts[tp],
            lineIndex: i,
            lineText: line
          });
          log('✅ 발견: ' + targetProducts[tp] + ' (라인 ' + i + ')');
          log('   텍스트: ' + line.substring(0, 100));
        }
      }
    }

    if (foundProducts.length === 0) {
      log('❌ REMI TARA 또는 SUGARPUNCH 제품을 찾을 수 없습니다.');
      return;
    }

    log('');
    log('📊 발견된 제품 수: ' + foundProducts.length);
    log('');

    // 각 제품에 대해 상세 분석
    for (var fp = 0; fp < foundProducts.length; fp++) {
      var product = foundProducts[fp];

      log('═'.repeat(80));
      log('🎯 제품 #' + (fp + 1) + ': ' + product.product);
      log('═'.repeat(80));
      log('라인 인덱스: ' + product.lineIndex);
      log('');

      // 해당 제품의 다음 15줄 수집 (OUTRE 다중 라인 형식)
      log('📋 원본 라인 데이터 (다음 15줄):');
      log('');

      var startIdx = product.lineIndex;
      var collectedLines = [];

      for (var j = 0; j < 15 && (startIdx + j) < lines.length; j++) {
        var currentLine = lines[startIdx + j].trim();
        collectedLines.push(currentLine);
        log('  [' + (startIdx + j) + '] "' + currentLine + '"');
      }

      log('');
      log('🔧 파싱 시뮬레이션 시작');
      log('');

      // QTY 찾기
      var qtyLine = null;
      var qtyValue = 0;
      var descriptionLines = [];
      var colorLines = [];
      var priceLines = [];

      // 현재 라인 또는 이전 라인에서 QTY 찾기
      for (var ql = Math.max(0, startIdx - 3); ql <= startIdx + 1; ql++) {
        var testLine = lines[ql].trim();
        if (testLine.match(/^\d{1,3}$/)) {
          var qty = parseInt(testLine);
          if (qty >= 0 && qty <= 700) {
            qtyLine = ql;
            qtyValue = qty;
            log('✅ QTY 발견:');
            log('   라인 인덱스: ' + ql);
            log('   값: ' + qty);
            break;
          }
        }
      }

      if (!qtyLine) {
        log('❌ QTY를 찾을 수 없습니다.');
        continue;
      }

      log('');
      log('📝 Description 수집 중...');

      // Description: QTY 다음 라인부터 수집
      var descStartIdx = qtyLine + 1;
      var foundFirstColor = false;

      for (var dl = 0; dl < 10 && (descStartIdx + dl) < lines.length; dl++) {
        var testLine = lines[descStartIdx + dl].trim();

        if (!testLine) continue;

        // 메타데이터 필터링
        if (testLine.match(/SHIP\s+TO|SOLD\s+TO|WEIGHT|SUBTOTAL|RICHMOND|LLC|PKWAY|COD|Fee|tag|DATE\s+SHIPPED|PAGE|SHIP\s+VIA|PAYMENT|TERMS|SALES|TOTAL/i)) {
          log('   [' + (descStartIdx + dl) + '] 메타데이터 건너뜀: ' + testLine.substring(0, 50));
          continue;
        }

        // 컬러 패턴 체크
        var hasColorPattern = testLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/);
        var isInchPattern = testLine.match(/\d+["″'']\s*-/);
        var isColorLine = hasColorPattern && !isInchPattern;

        if (isColorLine) {
          foundFirstColor = true;
          colorLines.push(testLine);
          log('   [' + (descStartIdx + dl) + '] 컬러 라인: ' + testLine.substring(0, 50));
          continue;
        }

        // Description 라인 (컬러 발견 전까지만)
        if (!foundFirstColor && descriptionLines.length < 3) {
          var isDescriptionLine = testLine.match(/^[A-Z]/) &&
                                 !testLine.match(/^\d+$/);

          if (isDescriptionLine) {
            descriptionLines.push(testLine);
            log('   [' + (descStartIdx + dl) + '] Description: ' + testLine.substring(0, 50));
          }
        }

        // 가격 패턴 체크
        if (testLine.match(/^\d+\.\d{2}$/)) {
          priceLines.push(testLine);
          log('   [' + (descStartIdx + dl) + '] 가격: $' + testLine);

          if (priceLines.length >= 3) {
            log('   ✅ 가격 3개 수집 완료, 중단');
            break;
          }
        }
      }

      log('');
      log('📊 수집 결과:');
      log('   QTY: ' + qtyValue);
      log('   Description 라인 수: ' + descriptionLines.length);
      log('   컬러 라인 수: ' + colorLines.length);
      log('   가격 라인 수: ' + priceLines.length);
      log('');

      // Description 조합
      var description = descriptionLines.join(' ');
      log('📝 Description (조합 전): "' + description + '"');

      var descriptionBeforeCleanup = description;

      // Description cleanup
      log('');
      log('🔧 Description Cleanup 시작...');
      log('   Cleanup 전: "' + description + '"');

      // 괄호 컬러 패턴 제거
      var colorInDescMatch = description.match(/^(.+?)(\d+["″''])\s*(\d*X)?\s*\([A-Z0-9\/\-]+\).*/i);
      if (colorInDescMatch) {
        var cleanDesc = (colorInDescMatch[1] + colorInDescMatch[2]).trim();
        if (colorInDescMatch[3]) {
          cleanDesc += ' ' + colorInDescMatch[3];
        }
        description = cleanDesc;
        log('   ✅ 괄호 컬러 패턴 제거: "' + description + '"');
      }

      // 인치 기준 절단
      var hasInch = description.match(/\d+["″'']/);
      if (hasInch) {
        var inchWithSuffix = description.match(/^(.+?\d+["″''](?:\s*-\s*[A-Z]{2,3})?)/);
        if (inchWithSuffix) {
          var beforeCleanup = description;
          description = inchWithSuffix[1].trim();
          if (beforeCleanup !== description) {
            log('   ✅ 인치 기준 절단: "' + description + '"');
          }
        }
      } else {
        // 컬러 패턴 절단
        var firstColorPattern = description.match(/^(.+?)\s+([A-Z0-9\/\-]{2,})\s*-\s*\d+/);
        if (firstColorPattern) {
          var beforeCleanup = description;
          description = firstColorPattern[1].trim();
          if (beforeCleanup !== description) {
            log('   ✅ 컬러 패턴 절단: "' + description + '"');
          }
        }
      }

      log('   최종 Description: "' + description + '"');
      log('');

      // parseColorLinesImproved 시뮬레이션
      log('█'.repeat(80));
      log('🎯 parseColorLinesImproved 시뮬레이션');
      log('█'.repeat(80));
      log('');
      log('📥 입력 파라미터:');
      log('   colorLines: ' + JSON.stringify(colorLines));
      log('   description (cleaned): "' + description + '"');
      log('   descriptionBeforeCleanup: "' + descriptionBeforeCleanup + '"');
      log('');

      var fullText = colorLines.join(' ');
      log('📝 colorLines.join(" "): "' + fullText + '"');

      // 언더스코어 제거
      fullText = fullText.replace(/_+/g, ' ');
      fullText = fullText.replace(/\s+/g, ' ').trim();
      log('📝 언더스코어 제거 후: "' + fullText + '"');
      log('');

      // Description 제거 로직
      log('🔧 Description 제거 로직 시작');

      if (descriptionBeforeCleanup) {
        var descClean = descriptionBeforeCleanup.trim();
        log('   descClean: "' + descClean + '"');
        log('   fullText: "' + fullText + '"');
        log('   fullText.indexOf(descClean): ' + fullText.indexOf(descClean));
        log('');

        // 방법 1: 정확 매칭
        if (fullText.indexOf(descClean) === 0) {
          fullText = fullText.substring(descClean.length).trim();
          log('   ✅ 방법 1 적용: 정확 매칭으로 제거');
          log('   제거된 부분: "' + descClean + '"');
          log('   남은 부분: "' + fullText + '"');
        } else {
          log('   ⏩ 방법 1 실패: 정확 매칭 안됨, 방법 2 시도');

          // 방법 2: 단어 기반 매칭
          var descWords = descClean.split(/[\s\-]+/).filter(function(word) {
            return word.length > 2 && !word.match(/^\d+$/) && !word.match(/^["″'']+$/);
          });
          log('   추출된 단어들 (3글자 이상): ' + JSON.stringify(descWords));

          if (descWords.length > 0) {
            var wordsToCheck = descWords.slice(0, Math.min(3, descWords.length));
            log('   검증할 단어들 (첫 3개): ' + JSON.stringify(wordsToCheck));

            var allWordsFound = true;
            var lastIndex = 0;

            for (var wi = 0; wi < wordsToCheck.length; wi++) {
              var wordIndex = fullText.indexOf(wordsToCheck[wi], lastIndex);
              log('     단어 "' + wordsToCheck[wi] + '" 검색: indexOf=' + wordIndex);
              if (wordIndex === -1) {
                allWordsFound = false;
                break;
              }
              lastIndex = wordIndex + wordsToCheck[wi].length;
            }

            log('   모든 단어 발견: ' + allWordsFound);

            if (allWordsFound) {
              var descEndMatch = fullText.match(/^.+?(\d+["″'']|X)\s*/);
              log('   Description 끝 패턴 매치: ' + (descEndMatch ? '"' + descEndMatch[0] + '"' : 'null'));

              if (descEndMatch) {
                var removedPart = fullText.substring(0, descEndMatch[0].length);
                fullText = fullText.substring(descEndMatch[0].length).trim();
                log('   ✅ 방법 2 적용: 단어 기반 매칭으로 제거');
                log('   제거된 부분: "' + removedPart + '"');
                log('   남은 부분: "' + fullText + '"');
              } else {
                log('   ❌ 방법 2 실패: Description 끝 패턴을 찾을 수 없음');
              }
            } else {
              log('   ❌ 방법 2 실패: 모든 단어가 발견되지 않음');
            }
          } else {
            log('   ❌ 방법 2 실패: 추출된 단어가 없음');
          }
        }
      } else {
        log('   ❌ Description이 없음 (undefined 또는 빈 문자열)');
      }

      log('');
      log('🔧 가격 정보 제거');
      log('   제거 전: "' + fullText + '"');
      fullText = fullText.replace(/\d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}\s*$/g, '');
      log('   제거 후: "' + fullText + '"');
      log('');

      // 컬러 패턴 매칭
      log('🎨 컬러 패턴 매칭 시작');
      log('   정규식: /([A-Z0-9\\-\\/]+)\\s*-\\s*(\\d+)\\s*(?:\\((\\d+)\\))?/gi');
      log('   대상 텍스트: "' + fullText + '"');
      log('');

      var regex = /([A-Z0-9\-\/]+)\s*-\s*(\d+)\s*(?:\((\d+)\))?/gi;
      var match;
      var matchCount = 0;
      var colorData = [];

      while ((match = regex.exec(fullText)) !== null) {
        matchCount++;
        var color = match[1].trim();
        var shipped = parseInt(match[2]) || 0;
        var backordered = match[3] ? parseInt(match[3]) : 0;

        log('   매치 #' + matchCount + ':');
        log('     전체 매치: "' + match[0] + '"');
        log('     컬러: "' + color + '"');
        log('     shipped: ' + shipped);
        log('     backordered: ' + backordered);

        if (color && color.length > 0 && (shipped > 0 || backordered > 0)) {
          colorData.push({
            color: color,
            shipped: shipped,
            backordered: backordered
          });
          log('     ✅ colorData에 추가됨');
        } else {
          log('     ❌ 조건 불충족, 추가 안됨');
        }
      }

      log('');
      log('📊 최종 결과:');
      log('   총 매치 수: ' + matchCount);
      log('   colorData 개수: ' + colorData.length);
      for (var i = 0; i < colorData.length; i++) {
        log('     [' + i + '] color="' + colorData[i].color + '", shipped=' + colorData[i].shipped + ', backordered=' + colorData[i].backordered);
      }
      log('');
      log('█'.repeat(80));
      log('');
    }

    log('');
    log('✅ 디버깅 완료!');
    log('');
    log('결과는 "' + debugSheetName + '" 시트에 저장되었습니다.');

    // 시트로 이동
    ss.setActiveSheet(debugSheet);

    // 첫 번째 열 너비 자동 조정
    debugSheet.autoResizeColumn(1);

    SpreadsheetApp.getUi().alert(
      '디버깅 완료',
      '결과가 "' + debugSheetName + '" 시트에 저장되었습니다.\n\n' +
      '로그를 복사해서 Claude에게 보내주세요.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('❌ 오류: ' + error.toString());
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      '오류 발생',
      '디버깅 중 오류가 발생했습니다:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * OUTRE Multi-line Parsing 디버그 - Line 461 문제 추적
 * Line 461이 descriptionLines와 colorLinesArray 둘 다에 추가되는 문제를 추적
 */
function debugOutreMultilineParsing() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 결과를 저장할 시트 생성/초기화
    var debugSheetName = 'DEBUG_MULTILINE';
    var debugSheet = ss.getSheetByName(debugSheetName);

    if (debugSheet) {
      ss.deleteSheet(debugSheet);
    }

    debugSheet = ss.insertSheet(debugSheetName);
    debugSheet.appendRow(['OUTRE MULTI-LINE PARSING DEBUG']);
    debugSheet.getRange(1, 1).setFontWeight('bold').setFontSize(14);

    var logRow = 2;

    function log(message) {
      debugSheet.getRange(logRow, 1).setValue(message);
      logRow++;
      Logger.log(message);
    }

    log('='.repeat(100));
    log('🔍 OUTRE Multi-line Parsing 디버그 - Line 461 문제 추적');
    log('='.repeat(100));
    log('');
    log('🎯 목적: Line 461이 descriptionLines와 colorLinesArray 둘 다에 추가되는지 확인');
    log('       그리고 Product 2의 colors가 Product 1의 colorLinesArray에 추가되는지 확인');
    log('');

    // 폴더에서 OUTRE 인보이스 파일 찾기
    var folderId = PropertiesService.getDocumentProperties()
      .getProperty(CONFIG.INVOICE.FOLDER_ID_PROPERTY);

    if (!folderId) {
      log('❌ 오류: 인보이스 폴더가 설정되지 않았습니다.');
      return;
    }

    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles();
    var outreFile = null;

    while (files.hasNext()) {
      var file = files.next();
      var filename = file.getName();

      if (filename.indexOf('SINV') > -1) {
        outreFile = file;
        log('✅ OUTRE 인보이스 파일: ' + filename);
        break;
      }
    }

    if (!outreFile) {
      log('❌ 오류: OUTRE 인보이스 파일(SINV...)을 찾을 수 없습니다.');
      return;
    }

    log('');

    // 텍스트 추출
    log('📝 텍스트 추출 중...');
    var text = '';

    if (outreFile.getMimeType() === MimeType.PDF) {
      text = extractTextFromPdf(outreFile);
    } else {
      text = extractTextFromDocx(outreFile.getBlob());
    }

    var lines = text.split('\n');
    log('  총 라인 수: ' + lines.length);
    log('');

    // 라인 460-475 출력 (문제가 있는 두 제품)
    log('📋 라인 460-475 (문제 영역):');
    log('');
    for (var i = 460; i <= 475 && i < lines.length; i++) {
      log('  [' + i + '] "' + lines[i].trim() + '"');
    }
    log('');

    // Product 1과 Product 2 파싱 시뮬레이션
    log('═'.repeat(100));
    log('🎯 PRODUCT 1 파싱 시뮬레이션 (QTY=55, Line 461)');
    log('═'.repeat(100));
    log('');

    var qtyLine1 = 460;
    var qtyValue1 = 55;

    log('✅ QTY: ' + qtyValue1 + ' (라인 ' + qtyLine1 + ')');
    log('');

    // OUTRE Multi-line 파싱 로직 시뮬레이션
    log('🔧 Multi-line 파싱 시작 (다음 라인부터 스캔)');
    log('   시작 인덱스: ' + (qtyLine1 + 1) + ' (라인 ' + (qtyLine1 + 1) + ')');
    log('');

    var descriptionLines = [];
    var colorLinesArray = [];
    var nextProductIndex = -1;

    var startIdx = qtyLine1 + 1;

    for (var i = 0; i < 15 && (startIdx + i) < lines.length; i++) {
      var currentIdx = startIdx + i;
      var nextLine = lines[currentIdx].trim();

      log('━'.repeat(100));
      log('🔍 [라인 ' + currentIdx + '] "' + nextLine + '"');
      log('');

      if (!nextLine) {
        log('   ⏩ 빈 라인, 건너뜀');
        log('');
        continue;
      }

      // 다음 제품 QTY 체크
      var isNextProductQty = false;
      if (nextLine.match(/^\d{1,3}$/)) {
        var nextQty = parseInt(nextLine);
        if (nextQty >= 1 && nextQty <= 700) {
          // 다음 라인이 제품명인지 체크
          var lineAfter = (currentIdx + 1) < lines.length ? lines[currentIdx + 1].trim() : '';
          var hasProductKeywords = lineAfter.match(/HAIR|WIG|LACE|WEAVE|CLIP|REMI|BATIK|SUGARPUNCH|X-PRESSION|BEAUTIFUL|MELTED|BRAID|CLOSURE|WAVE|CURL|STRAIGHT|BUNDLE|PONYTAIL|TARA|QW|BIG|BOHEMIAN|HD|PERUVIAN|TWIST|FEED|PASSION|LOOKS/i);

          if (hasProductKeywords) {
            isNextProductQty = true;
            nextProductIndex = currentIdx;
            log('   🚨 다음 제품 QTY 발견!');
            log('      QTY: ' + nextQty);
            log('      다음 라인: "' + lineAfter + '"');
            log('      ➡️ Multi-line 파싱 중단');
            log('');
            break;
          }
        }
      }

      // 메타데이터 필터링
      if (nextLine.match(/SHIP\s+TO|SOLD\s+TO|WEIGHT|SUBTOTAL|RICHMOND|LLC|PKWAY|COD|Fee|tag|DATE\s+SHIPPED|PAGE|SHIP\s+VIA|PAYMENT|TERMS|SALES|TOTAL/i)) {
        log('   ⏩ 메타데이터 라인, 건너뜀');
        log('');
        continue;
      }

      // 가격 패턴 체크 (3개 가격 = 종료)
      if (nextLine.match(/^\d+\.\d{2}$/)) {
        log('   💰 가격 라인');
        log('');
        continue;
      }

      // Description 후보 체크
      var isDescriptionCandidate = false;
      var blacklistKeywords = ['X-PRESSION', 'BEAUTIFUL', 'MELTED', 'BATIK', 'SUGARPUNCH', 'REMI', 'PONYTAIL', 'TARA', 'QW', 'BIG', 'BOHEMIAN', 'HD', 'PERUVIAN'];

      for (var bk = 0; bk < blacklistKeywords.length; bk++) {
        if (nextLine.indexOf(blacklistKeywords[bk]) > -1) {
          isDescriptionCandidate = true;
          log('   📝 Description 후보 (블랙리스트 키워드: "' + blacklistKeywords[bk] + '")');
          break;
        }
      }

      // Color 라인 체크
      var colorPattern = nextLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/);
      var inchPattern = nextLine.match(/\d+["″'']\s*-/);
      var isColorLine = colorPattern && !inchPattern;

      if (isColorLine) {
        log('   🎨 Color 패턴 발견');
        log('      매치: "' + colorPattern[0] + '"');
      }

      // 🚨 핵심 로직: Description 후보 AND Color 라인?
      if (isDescriptionCandidate && isColorLine) {
        log('');
        log('   🚨🚨🚨 중요: 이 라인은 Description 후보이면서 동시에 Color 패턴을 포함!');
        log('');
      }

      // 실제 처리
      if (isDescriptionCandidate) {
        descriptionLines.push(nextLine);
        log('   ✅ descriptionLines에 추가 (현재 개수: ' + descriptionLines.length + ')');
        log('      배열 상태: ' + JSON.stringify(descriptionLines.map(function(l) { return l.substring(0, 50); })));
      }

      if (isColorLine) {
        colorLinesArray.push(nextLine);
        log('   ✅ colorLinesArray에 추가 (현재 개수: ' + colorLinesArray.length + ')');
        log('      배열 상태: ' + JSON.stringify(colorLinesArray.map(function(l) { return l.substring(0, 50); })));
      }

      log('');

      // Description 3개 수집하면 color만 수집
      if (descriptionLines.length >= 3) {
        log('   ℹ️ Description 3개 수집 완료, 이후로는 Color만 수집');
        log('');
      }
    }

    log('');
    log('📊 PRODUCT 1 최종 수집 결과:');
    log('   descriptionLines 개수: ' + descriptionLines.length);
    for (var i = 0; i < descriptionLines.length; i++) {
      log('     [' + i + '] "' + descriptionLines[i] + '"');
    }
    log('');
    log('   colorLinesArray 개수: ' + colorLinesArray.length);
    for (var i = 0; i < colorLinesArray.length; i++) {
      log('     [' + i + '] "' + colorLinesArray[i] + '"');
    }
    log('');
    log('   다음 제품 인덱스: ' + nextProductIndex);
    log('');

    // 문제 분석
    log('═'.repeat(100));
    log('🔎 문제 분석');
    log('═'.repeat(100));
    log('');

    var line461 = lines[461].trim();
    var line461InDesc = false;
    var line461InColor = false;

    for (var i = 0; i < descriptionLines.length; i++) {
      if (descriptionLines[i] === line461) {
        line461InDesc = true;
        break;
      }
    }

    for (var i = 0; i < colorLinesArray.length; i++) {
      if (colorLinesArray[i] === line461) {
        line461InColor = true;
        break;
      }
    }

    log('Line 461: "' + line461 + '"');
    log('');
    log('  descriptionLines에 포함? ' + (line461InDesc ? '✅ YES' : '❌ NO'));
    log('  colorLinesArray에 포함? ' + (line461InColor ? '✅ YES' : '❌ NO'));
    log('');

    if (line461InDesc && line461InColor) {
      log('🚨 문제 확인: Line 461이 descriptionLines와 colorLinesArray 둘 다에 추가되었습니다!');
      log('   이것이 phantom line 115-116을 만드는 근본 원인입니다.');
    } else if (line461InDesc) {
      log('✅ Line 461은 descriptionLines에만 있습니다.');
    } else if (line461InColor) {
      log('✅ Line 461은 colorLinesArray에만 있습니다.');
    }
    log('');

    // Product 2 체크
    if (nextProductIndex === -1) {
      log('⚠️ 다음 제품을 찾지 못했습니다. Product 2의 colors도 Product 1에 추가되었을 수 있습니다.');
      log('');

      // Line 470-471 체크 (Product 2의 colors)
      var line470 = lines[470] ? lines[470].trim() : '';
      var line471 = lines[471] ? lines[471].trim() : '';

      log('Line 470: "' + line470 + '"');
      log('Line 471: "' + line471 + '"');
      log('');

      var line471InColor = false;
      for (var i = 0; i < colorLinesArray.length; i++) {
        if (colorLinesArray[i] === line471) {
          line471InColor = true;
          break;
        }
      }

      if (line471InColor) {
        log('🚨 문제 확인: Line 471 (Product 2의 color line)이 Product 1의 colorLinesArray에 추가되었습니다!');
        log('   이것이 phantom line 115-116 (색상 1, 1B)을 만듭니다.');
      }
    } else {
      log('✅ 다음 제품을 라인 ' + nextProductIndex + '에서 찾았습니다.');
      log('   Product 2의 colors는 Product 1에 추가되지 않았어야 합니다.');
    }

    log('');
    log('═'.repeat(100));
    log('✅ 디버깅 완료!');
    log('═'.repeat(100));
    log('');
    log('결과는 "' + debugSheetName + '" 시트에 저장되었습니다.');
    log('로그를 복사해서 Claude에게 보내주세요.');

    // 시트로 이동
    ss.setActiveSheet(debugSheet);
    debugSheet.autoResizeColumn(1);

    SpreadsheetApp.getUi().alert(
      '디버깅 완료',
      '결과가 "' + debugSheetName + '" 시트에 저장되었습니다.\n\n' +
      '로그를 복사해서 Claude에게 보내주세요.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    Logger.log('❌ 오류: ' + error.toString());
    Logger.log(error.stack);

    SpreadsheetApp.getUi().alert(
      '오류 발생',
      '디버깅 중 오류가 발생했습니다:\n\n' + error.toString(),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}
