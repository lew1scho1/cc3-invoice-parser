// ============================================================================
// DebugOUTRE.js - OUTRE 파서 디버깅 전용 함수
// ============================================================================

/**
 * OUTRE 복합 컬러명 파싱 디버그
 * - 컬러 라인 수집 과정 추적
 * - parseOUTREColorLines() 내부 동작 상세 로그
 * - 복합 컬러명(M950/425/350/130S) 분리 여부 확인
 *
 * 사용법:
 * 1. Google Apps Script 편집기에서 debugOUTREComplexColors() 실행
 * 2. DEBUG_OUTPUT 시트에서 결과 확인
 *
 * CRITICAL: SINV1911616.docx 파일을 직접 읽어서 파싱 (PARSING 시트 불필요)
 */
function debugOUTREComplexColors() {
  // CRITICAL: OUTRE 테스트 파일 하드코딩
  var TEST_FILE_NAME = 'SINV1911616.docx';

  Logger.log('=== OUTRE 복합 컬러명 파싱 디버그 시작 ===');
  Logger.log('테스트 파일: ' + TEST_FILE_NAME);

  // Google Drive에서 파일 검색
  var files = DriveApp.getFilesByName(TEST_FILE_NAME);
  if (!files.hasNext()) {
    Logger.log('❌ 파일을 찾을 수 없음: ' + TEST_FILE_NAME);
    Logger.log('Google Drive에 파일을 업로드하세요.');
    return;
  }

  var file = files.next();
  Logger.log('✅ 파일 발견: ' + file.getName() + ' (ID: ' + file.getId() + ')');

  // 텍스트 추출
  var blob = file.getBlob();
  var text = extractTextFromDocx(blob);

  if (!text) {
    Logger.log('❌ 텍스트 추출 실패');
    return;
  }

  Logger.log('✅ 텍스트 추출 성공: ' + text.length + ' bytes');

  // 라인 분리
  var lines = text.split(/\r?\n/);
  Logger.log('총 라인 수: ' + lines.length);

  Logger.log('='.repeat(80));
  Logger.log('OUTRE 복합 컬러명 파싱 디버그 시작');
  Logger.log('='.repeat(80));
  Logger.log('총 라인 수: ' + lines.length);

  var debugOutput = [];
  debugOutput.push('='.repeat(80));
  debugOutput.push('OUTRE 복합 컬러명 파싱 디버그');
  debugOutput.push('='.repeat(80));
  debugOutput.push('');

  // 복합 컬러 패턴 검색 (슬래시 2개 이상 포함)
  var complexColorPattern = /[A-Z0-9]+\/[A-Z0-9]+\/[A-Z0-9]+/i;
  var complexColorItems = [];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (complexColorPattern.test(line)) {
      complexColorItems.push({ index: i, line: line });
    }
  }

  debugOutput.push('📊 복합 컬러 패턴 검색 결과 (슬래시 2개 이상)');
  debugOutput.push('발견된 라인 수: ' + complexColorItems.length);
  debugOutput.push('');

  if (complexColorItems.length === 0) {
    debugOutput.push('⚠️ 복합 컬러 패턴을 찾을 수 없음');
    debugOutput.push('인보이스에 M950/425/350/130S 같은 패턴이 없습니다.');
    writeDebugOutput(debugOutput.join('\n'));
    return;
  }

  // 각 복합 컬러 라인에 대해 상세 분석
  for (var ci = 0; ci < complexColorItems.length; ci++) {
    var item = complexColorItems[ci];
    var lineIndex = item.index;
    var lineText = item.line;

    debugOutput.push('━'.repeat(80));
    debugOutput.push('복합 컬러 라인 #' + (ci + 1) + ' (Line ' + lineIndex + ')');
    debugOutput.push('━'.repeat(80));
    debugOutput.push('원문: ' + lineText);
    debugOutput.push('');

    // 이 라인 주변 컨텍스트 출력 (앞뒤 5줄)
    debugOutput.push('[ 주변 컨텍스트 ]');
    for (var j = Math.max(0, lineIndex - 5); j < Math.min(lines.length, lineIndex + 6); j++) {
      var marker = j === lineIndex ? '>>> ' : '    ';
      debugOutput.push(marker + 'Line ' + j + ': ' + lines[j].substring(0, 100));
    }
    debugOutput.push('');

    // QTY 라인 찾기 (역방향 검색, 최대 10줄)
    var qtyLine = -1;
    for (var j = lineIndex - 1; j >= Math.max(0, lineIndex - 10); j--) {
      if (lines[j].trim().match(/^\d{1,3}$/)) {
        var qty = parseInt(lines[j].trim());
        if (qty >= 0 && qty <= 700) {
          qtyLine = j;
          break;
        }
      }
    }

    if (qtyLine === -1) {
      debugOutput.push('⚠️ QTY 라인을 찾을 수 없음 (라인 ' + lineIndex + ' 기준 위 10줄 검색)');
      debugOutput.push('');
      continue;
    }

    debugOutput.push('✅ QTY 라인 발견: Line ' + qtyLine + ' (QTY=' + lines[qtyLine].trim() + ')');
    debugOutput.push('');

    // parseOUTREItem() 시뮬레이션
    debugOutput.push('[ parseOUTREItem() 시뮬레이션 ]');

    var qtyShipped = parseInt(lines[qtyLine].trim());
    var descriptionLines = [];
    var colorLinesArray = [];
    var foundFirstColor = false;

    debugOutput.push('QTY: ' + qtyShipped);
    debugOutput.push('검색 범위: Line ' + (qtyLine + 1) + ' ~ Line ' + Math.min(qtyLine + 15, lines.length - 1));
    debugOutput.push('');

    for (var j = qtyLine + 1; j < Math.min(qtyLine + 15, lines.length); j++) {
      var nextLine = lines[j].trim();
      if (!nextLine) continue;

      var hasColorPattern = nextLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/);
      var isInchPattern = nextLine.match(/\d+["″'']/);

      var DESCRIPTION_BLACKLIST = ['X-PRESSION', 'SHAKE-N-GO', 'BATIK', 'SUGARPUNCH'];
      var hasBlacklistedWord = false;

      if (hasColorPattern) {
        var upperLine = nextLine.toUpperCase();
        for (var bi = 0; bi < DESCRIPTION_BLACKLIST.length; bi++) {
          if (upperLine.indexOf(DESCRIPTION_BLACKLIST[bi]) > -1) {
            hasBlacklistedWord = true;
            break;
          }
        }
      }

      var isColorLine = hasColorPattern && !isInchPattern && !hasBlacklistedWord;
      var isDescriptionCandidate = !foundFirstColor && nextLine.length > 5;

      var action = '';
      if (isDescriptionCandidate && !isColorLine) {
        descriptionLines.push(nextLine);
        action = '→ Description 라인 추가';
      } else if (isDescriptionCandidate && isColorLine) {
        foundFirstColor = true;
        colorLinesArray.push(nextLine);
        action = '→ 컬러 라인 추가 (Description 종료)';
      } else if (isColorLine) {
        colorLinesArray.push(nextLine);
        foundFirstColor = true;
        action = '→ 컬러 라인 추가';
      } else {
        action = '→ 건너뜀';
      }

      debugOutput.push('  Line ' + j + ': ' + nextLine.substring(0, 80));
      debugOutput.push('    hasColorPattern=' + !!hasColorPattern +
                       ', isInchPattern=' + !!isInchPattern +
                       ', hasBlacklist=' + hasBlacklistedWord +
                       ', isColorLine=' + isColorLine);
      debugOutput.push('    ' + action);

      if (j === lineIndex) {
        debugOutput.push('    ★ 이 라인이 복합 컬러 라인입니다');
      }
      debugOutput.push('');
    }

    debugOutput.push('[ 수집 결과 ]');
    debugOutput.push('Description 라인 수: ' + descriptionLines.length);
    debugOutput.push('컬러 라인 수: ' + colorLinesArray.length);
    debugOutput.push('');

    if (descriptionLines.length > 0) {
      debugOutput.push('Description:');
      for (var d = 0; d < descriptionLines.length; d++) {
        debugOutput.push('  ' + (d + 1) + '. ' + descriptionLines[d]);
      }
      debugOutput.push('');
    }

    if (colorLinesArray.length === 0) {
      debugOutput.push('⚠️ 컬러 라인이 수집되지 않음!');
      debugOutput.push('원인: 복합 컬러 라인이 Description으로 인식되었거나 건너뛰어짐');
      debugOutput.push('');
      continue;
    }

    debugOutput.push('컬러 라인:');
    for (var cl = 0; cl < colorLinesArray.length; cl++) {
      debugOutput.push('  ' + (cl + 1) + '. ' + colorLinesArray[cl]);
    }
    debugOutput.push('');

    // parseOUTREColorLines() 시뮬레이션
    debugOutput.push('[ parseOUTREColorLines() 시뮬레이션 ]');

    var rawDescription = descriptionLines.join(' ').trim();
    debugOutput.push('Description (제거용): ' + rawDescription.substring(0, 80));
    debugOutput.push('');

    var fullText = colorLinesArray.join(' ');
    debugOutput.push('Step 0 (원본): ' + fullText.substring(0, 150));

    // Step 1: Normalize
    fullText = normalizeOutreText(fullText);
    debugOutput.push('Step 1 (Normalize): ' + fullText.substring(0, 150));

    // Step 2: Description 제거
    if (rawDescription) {
      var descWords = rawDescription.split(/\s+/);
      for (var i = 0; i < descWords.length; i++) {
        var word = descWords[i].trim();
        if (word.length > 2) {
          var regex = new RegExp('\\b' + word + '\\b', 'gi');
          fullText = fullText.replace(regex, ' ');
        }
      }
      fullText = normalizeOutreText(fullText);
      debugOutput.push('Step 2 (Description 제거): ' + fullText.substring(0, 150));
    }

    // Step 3: 가격 제거
    fullText = fullText.replace(/\d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}/g, ' ');
    fullText = normalizeOutreText(fullText);
    debugOutput.push('Step 3 (가격 제거): ' + fullText.substring(0, 150));

    // Step 4: 컬러 패턴 매칭
    debugOutput.push('');
    debugOutput.push('Step 4 (컬러 패턴 추출):');

    var colorPattern = /(\([A-Z]\))?([A-Z0-9\/-]+)\s*-\s*(\d+)(?:\s*\((\d+)\))?/gi;
    var match;
    var matchCount = 0;
    var colorData = [];

    while ((match = colorPattern.exec(fullText)) !== null) {
      matchCount++;
      var prefix = match[1] || '';
      var colorToken = match[2].trim();
      var shipped = parseInt(match[3]);
      var backordered = match[4] ? parseInt(match[4]) : 0;

      debugOutput.push('');
      debugOutput.push('  매치 #' + matchCount + ':');
      debugOutput.push('    전체 매치: ' + match[0]);
      debugOutput.push('    괄호 접두사: ' + (prefix || '(없음)'));
      debugOutput.push('    컬러 토큰: ' + colorToken);
      debugOutput.push('    Shipped: ' + shipped);
      debugOutput.push('    Backordered: ' + backordered);

      // 토큰 길이 체크
      debugOutput.push('    토큰 길이: ' + colorToken.length + '자');

      // 슬래시 개수 체크
      var slashCount = (colorToken.match(/\//g) || []).length;
      debugOutput.push('    슬래시 개수: ' + slashCount);

      // validateOUTREColorToken() 검증
      var isValid = validateOUTREColorToken(colorToken);
      debugOutput.push('    validateOUTREColorToken(): ' + (isValid ? '✅ PASS' : '❌ FAIL'));

      if (isValid) {
        var finalColor = prefix ? colorToken : colorToken.toUpperCase();
        colorData.push({
          color: finalColor,
          shipped: shipped,
          backordered: backordered
        });
        debugOutput.push('    ✅ 컬러 데이터 추가: ' + finalColor);
      } else {
        debugOutput.push('    ❌ 유효하지 않은 컬러 토큰, 무시됨');
      }
    }

    debugOutput.push('');
    debugOutput.push('[ 최종 결과 ]');
    debugOutput.push('총 매치 수: ' + matchCount);
    debugOutput.push('유효한 컬러 수: ' + colorData.length);
    debugOutput.push('');

    if (colorData.length === 0) {
      debugOutput.push('❌ 최종 컬러 데이터 없음!');
    } else {
      debugOutput.push('컬러 데이터:');
      for (var cd = 0; cd < colorData.length; cd++) {
        var c = colorData[cd];
        debugOutput.push('  ' + (cd + 1) + '. ' + c.color + ' - Shipped: ' + c.shipped + ', Backordered: ' + c.backordered);
      }
    }

    // 복합 컬러명 분리 여부 판단
    debugOutput.push('');
    if (colorData.length > 1) {
      debugOutput.push('⚠️ 복합 컬러명이 ' + colorData.length + '개로 분리되었습니다!');
      debugOutput.push('예상: 1개 컬러 (M950/425/350/130S)');
      debugOutput.push('실제: ' + colorData.length + '개 컬러 (' + colorData.map(function(c) { return c.color; }).join(', ') + ')');
    } else if (colorData.length === 1) {
      debugOutput.push('✅ 복합 컬러명이 단일 컬러로 인식되었습니다.');
      debugOutput.push('컬러: ' + colorData[0].color);
    }

    debugOutput.push('');
  }

  debugOutput.push('='.repeat(80));
  debugOutput.push('디버그 완료');
  debugOutput.push('='.repeat(80));

  // DEBUG_OUTPUT 시트에 출력
  writeDebugOutput(debugOutput.join('\n'));

  Logger.log('✅ 디버그 완료 - DEBUG_OUTPUT 시트 확인');
}

/**
 * DEBUG_OUTPUT 시트에 디버그 로그 작성
 * @param {string} text - 출력할 텍스트
 */
function writeDebugOutput(text) {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName('DEBUG_OUTPUT');

  if (!sheet) {
    sheet = ss.insertSheet('DEBUG_OUTPUT');
  }

  sheet.clear();

  var lines = text.split('\n');
  var data = lines.map(function(line) { return [line]; });

  if (data.length > 0) {
    sheet.getRange(1, 1, data.length, 1).setValues(data);
  }

  // 열 너비 자동 조정
  sheet.setColumnWidth(1, 1200);

  // 글꼴 고정폭
  sheet.getRange(1, 1, data.length, 1).setFontFamily('Courier New');
}

/**
 * OUTRE 멀티라인 백오더 파싱 디버그
 * - S4/30- 0 다음 줄의 (1) 패턴 추적
 * - colorLinesArray 수집 여부 확인
 *
 * 사용법: debugOUTREMultilineBackorder() 실행
 *
 * CRITICAL: SINV1911616.docx 파일을 직접 읽어서 파싱 (PARSING 시트 불필요)
 */
function debugOUTREMultilineBackorder() {
  // CRITICAL: OUTRE 테스트 파일 하드코딩
  var TEST_FILE_NAME = 'SINV1911616.docx';

  Logger.log('=== OUTRE 멀티라인 백오더 파싱 디버그 시작 ===');
  Logger.log('테스트 파일: ' + TEST_FILE_NAME);

  // Google Drive에서 파일 검색
  var files = DriveApp.getFilesByName(TEST_FILE_NAME);
  if (!files.hasNext()) {
    Logger.log('❌ 파일을 찾을 수 없음: ' + TEST_FILE_NAME);
    Logger.log('Google Drive에 파일을 업로드하세요.');
    return;
  }

  var file = files.next();
  Logger.log('✅ 파일 발견: ' + file.getName() + ' (ID: ' + file.getId() + ')');

  // 텍스트 추출
  var blob = file.getBlob();
  var text = extractTextFromDocx(blob);

  if (!text) {
    Logger.log('❌ 텍스트 추출 실패');
    return;
  }

  Logger.log('✅ 텍스트 추출 성공: ' + text.length + ' bytes');

  // 라인 분리
  var lines = text.split(/\r?\n/);
  Logger.log('총 라인 수: ' + lines.length);

  Logger.log('='.repeat(80));
  Logger.log('OUTRE 멀티라인 백오더 파싱 디버그');
  Logger.log('='.repeat(80));

  var debugOutput = [];
  debugOutput.push('='.repeat(80));
  debugOutput.push('OUTRE 멀티라인 백오더 파싱 디버그');
  debugOutput.push('='.repeat(80));
  debugOutput.push('');

  // 멀티라인 백오더 패턴 검색: "- 0" 다음 줄에 "(\d+)"
  var backorderCandidates = [];

  for (var i = 0; i < lines.length - 1; i++) {
    var line = lines[i];
    var nextLine = lines[i + 1];

    // "- 0" 패턴 찾기
    if (line.match(/[A-Z0-9\-\/]+\s*-\s*0/) && nextLine.trim().match(/^\(\d+\)$/)) {
      backorderCandidates.push({
        lineIndex: i,
        colorLine: line,
        backorderLine: nextLine
      });
    }
  }

  debugOutput.push('📊 멀티라인 백오더 패턴 검색 결과');
  debugOutput.push('발견된 패턴 수: ' + backorderCandidates.length);
  debugOutput.push('');

  if (backorderCandidates.length === 0) {
    debugOutput.push('⚠️ 멀티라인 백오더 패턴을 찾을 수 없음');
    debugOutput.push('패턴: "COLOR- 0" 다음 줄에 "(\d+)"');
    writeDebugOutput(debugOutput.join('\n'));
    return;
  }

  for (var ci = 0; ci < backorderCandidates.length; ci++) {
    var candidate = backorderCandidates[ci];
    var lineIndex = candidate.lineIndex;

    debugOutput.push('━'.repeat(80));
    debugOutput.push('멀티라인 백오더 #' + (ci + 1) + ' (Line ' + lineIndex + ')');
    debugOutput.push('━'.repeat(80));
    debugOutput.push('컬러 라인 (Line ' + lineIndex + '): ' + candidate.colorLine);
    debugOutput.push('백오더 라인 (Line ' + (lineIndex + 1) + '): ' + candidate.backorderLine);
    debugOutput.push('');

    // QTY 라인 찾기
    var qtyLine = -1;
    for (var j = lineIndex - 1; j >= Math.max(0, lineIndex - 15); j--) {
      if (lines[j].trim().match(/^\d{1,3}$/)) {
        var qty = parseInt(lines[j].trim());
        if (qty >= 0 && qty <= 700) {
          qtyLine = j;
          break;
        }
      }
    }

    if (qtyLine === -1) {
      debugOutput.push('⚠️ QTY 라인을 찾을 수 없음');
      debugOutput.push('');
      continue;
    }

    debugOutput.push('✅ QTY 라인 발견: Line ' + qtyLine + ' (QTY=' + lines[qtyLine].trim() + ')');
    debugOutput.push('');

    // parseOUTREItem() 시뮬레이션
    debugOutput.push('[ parseOUTREItem() 컬러 라인 수집 시뮬레이션 ]');

    var colorLinesArray = [];
    var foundFirstColor = false;

    for (var j = qtyLine + 1; j < Math.min(qtyLine + 15, lines.length); j++) {
      var nextLine = lines[j].trim();
      if (!nextLine) continue;

      var hasColorPattern = nextLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/);
      var isInchPattern = nextLine.match(/\d+["″'']/);

      var DESCRIPTION_BLACKLIST = ['X-PRESSION', 'SHAKE-N-GO', 'BATIK', 'SUGARPUNCH'];
      var hasBlacklistedWord = false;

      if (hasColorPattern) {
        var upperLine = nextLine.toUpperCase();
        for (var bi = 0; bi < DESCRIPTION_BLACKLIST.length; bi++) {
          if (upperLine.indexOf(DESCRIPTION_BLACKLIST[bi]) > -1) {
            hasBlacklistedWord = true;
            break;
          }
        }
      }

      var isColorLine = hasColorPattern && !isInchPattern && !hasBlacklistedWord;

      // 단독 괄호 라인 체크
      var isOrphanBackorder = nextLine.match(/^\(\d+\)$/);

      var action = '';
      var collected = false;

      if (isColorLine) {
        colorLinesArray.push(nextLine);
        foundFirstColor = true;
        action = '→ 컬러 라인 추가';
        collected = true;
      } else if (isOrphanBackorder && foundFirstColor) {
        action = '→ 단독 괄호 라인 (현재: 수집 안 됨)';
        // CRITICAL: 여기서 수집되지 않음!
      } else {
        action = '→ 건너뜀';
      }

      debugOutput.push('  Line ' + j + ': ' + nextLine.substring(0, 80));
      debugOutput.push('    hasColorPattern=' + !!hasColorPattern +
                       ', isColorLine=' + isColorLine +
                       ', isOrphanBackorder=' + !!isOrphanBackorder);
      debugOutput.push('    ' + action);

      if (j === lineIndex) {
        debugOutput.push('    ★ 이 라인이 "- 0" 컬러 라인입니다');
      } else if (j === lineIndex + 1) {
        debugOutput.push('    ★ 이 라인이 백오더 라인입니다 (수집됨: ' + collected + ')');
      }
      debugOutput.push('');
    }

    debugOutput.push('[ 수집 결과 ]');
    debugOutput.push('컬러 라인 수: ' + colorLinesArray.length);
    debugOutput.push('');

    if (colorLinesArray.length > 0) {
      debugOutput.push('수집된 컬러 라인:');
      for (var cl = 0; cl < colorLinesArray.length; cl++) {
        debugOutput.push('  ' + (cl + 1) + '. ' + colorLinesArray[cl]);
      }
      debugOutput.push('');

      // 백오더 라인이 수집되었는지 확인
      var backorderLineText = candidate.backorderLine.trim();
      var backorderCollected = false;
      for (var cl = 0; cl < colorLinesArray.length; cl++) {
        if (colorLinesArray[cl] === backorderLineText) {
          backorderCollected = true;
          break;
        }
      }

      if (backorderCollected) {
        debugOutput.push('✅ 백오더 라인이 colorLinesArray에 수집되었습니다.');
      } else {
        debugOutput.push('❌ 백오더 라인이 colorLinesArray에 수집되지 않았습니다!');
        debugOutput.push('원인: 단독 괄호 라인이 컬러 패턴으로 인식되지 않음');
      }
    } else {
      debugOutput.push('⚠️ 컬러 라인이 수집되지 않음');
    }

    debugOutput.push('');
  }

  debugOutput.push('='.repeat(80));
  debugOutput.push('디버그 완료');
  debugOutput.push('='.repeat(80));

  writeDebugOutput(debugOutput.join('\n'));
  Logger.log('✅ 디버그 완료 - DEBUG_OUTPUT 시트 확인');
}

/**
 * Reference 폴더 내 OUTRE 인보이스 이상 항목 점검
 * - itemId/UPC 누락
 * - memo 경고 포함
 *
 * 사용법: debugOUTREReferenceIssues() 실행
 */
function debugOUTREReferenceIssues() {
  var REFERENCE_FOLDER_NAME = 'Reference';
  var FILE_NAME_FILTER = /^SINV\d+\.docx$/i;

  var folders = DriveApp.getFoldersByName(REFERENCE_FOLDER_NAME);
  if (!folders.hasNext()) {
    Logger.log('Reference 폴더를 찾을 수 없음: ' + REFERENCE_FOLDER_NAME);
    return;
  }

  var folder = folders.next();
  var filesIter = folder.getFiles();
  var files = [];

  while (filesIter.hasNext()) {
    var file = filesIter.next();
    if (FILE_NAME_FILTER.test(file.getName())) {
      files.push(file);
    }
  }

  files.sort(function(a, b) {
    return a.getName().localeCompare(b.getName());
  });

  var output = [];
  output.push('OUTRE Reference Debug');
  output.push('Folder: ' + folder.getName());
  output.push('Files: ' + files.length);
  output.push('');

  for (var i = 0; i < files.length; i++) {
    var file = files[i];
    output.push('--- ' + file.getName() + ' ---');

    try {
      var text = extractTextFromDocx(file.getBlob());
      if (!text) {
        output.push('ERROR: text extraction failed');
        output.push('');
        continue;
      }

      var data = parseInvoice(text, file.getName());
      output.push('Invoice: ' + (data.invoiceNo || '(none)') +
                  ' | Vendor: ' + data.vendor +
                  ' | Lines: ' + data.lineItems.length);

      if (data.vendor !== 'OUTRE') {
        output.push('SKIP: vendor is not OUTRE');
        output.push('');
        continue;
      }

      var issueCount = 0;
      for (var li = 0; li < data.lineItems.length; li++) {
        var item = data.lineItems[li];
        var notes = [];

        if (!item.itemId) {
          notes.push('NO_ITEM');
        }
        if (item.color && !item.upc) {
          notes.push('NO_UPC');
        }
        if (item.memo && item.memo.indexOf('⚠️') > -1) {
          notes.push('MEMO=' + item.memo);
        }

        if (notes.length > 0) {
          issueCount++;
          output.push('Line ' + item.lineNo + ': ' + item.description +
                      ' | Color: ' + (item.color || '-') +
                      ' | ' + notes.join(' | '));
        }
      }

      if (issueCount === 0) {
        output.push('OK: no issues');
      }
    } catch (error) {
      output.push('ERROR: ' + error.toString());
    }

    output.push('');
  }

  writeDebugOutput(output.join('\n'));
  Logger.log('Debug 완료 - DEBUG_OUTPUT 시트 확인');
}

/**
 * PARSING 탭 기준 문제 라인 로그 출력
 * - OUTRE 기준 경고만 추출
 * - itemId/UPC 누락, memo 경고 포함
 *
 * 사용법: debugOUTREParsingIssues() 실행
 */
function debugOUTREParsingIssues() {
  var sheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    Logger.log('PARSING 탭에 데이터가 없습니다.');
    return;
  }

  var output = [];
  output.push('OUTRE Parsing Debug (PARSING 탭)');
  output.push('Rows: ' + (data.length - 1));
  output.push('');

  var issueCount = 0;
  var invoiceGroups = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var vendor = row[0];

    if (vendor !== 'OUTRE') {
      continue;
    }

    var invoiceNo = row[1];
    var lineNo = row[8];
    var itemId = row[9];
    var upc = row[10];
    var description = row[11];
    var color = row[13];
    var qtyShipped = row[16];
    var unitPrice = row[17];
    var extPrice = row[18];
    var memo = row[19];

    var notes = [];

    if (!description) {
      notes.push('NO_DESC');
    }
    if (!itemId) {
      notes.push('NO_ITEM');
    }
    if (color && !upc) {
      notes.push('NO_UPC');
    }
    if (!unitPrice || unitPrice === 0) {
      notes.push('UNIT_0');
    }
    if (!extPrice || extPrice === 0) {
      notes.push('EXT_0');
    }
    if (!qtyShipped || qtyShipped === 0) {
      notes.push('QTY_0');
    }
    if (memo && memo.indexOf('⚠️') > -1) {
      notes.push('MEMO=' + memo);
    }

    if (notes.length === 0) {
      continue;
    }

    if (!invoiceGroups[invoiceNo]) {
      invoiceGroups[invoiceNo] = [];
    }

    invoiceGroups[invoiceNo].push({
      lineNo: lineNo,
      description: description,
      color: color,
      notes: notes
    });

    issueCount++;
  }

  var invoices = Object.keys(invoiceGroups);
  if (invoices.length === 0) {
    output.push('OK: no issues found');
    writeDebugOutput(output.join('\n'));
    Logger.log('Debug 완료 - DEBUG_OUTPUT 시트 확인');
    return;
  }

  invoices.sort();

  for (var ii = 0; ii < invoices.length; ii++) {
    var inv = invoices[ii];
    output.push('--- ' + inv + ' ---');

    var lines = invoiceGroups[inv];
    lines.sort(function(a, b) {
      return a.lineNo - b.lineNo;
    });

    for (var li = 0; li < lines.length; li++) {
      var item = lines[li];
      output.push('Line ' + item.lineNo + ': ' + item.description +
                  ' | Color: ' + (item.color || '-') +
                  ' | ' + item.notes.join(' | '));
    }

    output.push('');
  }

  output.unshift('Issues: ' + issueCount);
  writeDebugOutput(output.join('\n'));
  Logger.log('Debug 완료 - DEBUG_OUTPUT 시트 확인');
}

/**
 * Debug trailing quote matches using PARSING tab.
 * - Filters OUTRE rows whose description ends with a quote-like char.
 * - Shows matchOUTREDescriptionFromDB result and size tokens.
 *
 * Usage: debugOUTRETrailingQuoteFromParsing()
 */
function debugOUTRETrailingQuoteFromParsing() {
  var sheet = getSheet(CONFIG.INVOICE.PARSING_SHEET);
  var data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    Logger.log('PARSING sheet has no data.');
    return;
  }

  var output = [];
  output.push('OUTRE Trailing Quote Debug (PARSING)');
  output.push('Rows: ' + (data.length - 1));
  output.push('');

  var quoteChars = ['"', "'", '��', '��', '`', '��', '��', '��'];
  var issueCount = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var vendor = row[0];
    if (vendor !== 'OUTRE') continue;

    var invoiceNo = row[1];
    var lineNo = row[8];
    var description = row[11] || '';
    var memo = row[19] || '';

    var trimmed = description.replace(/[\s\u200B-\u200D\uFEFF]+$/, '');
    if (!trimmed) continue;

    var lastChar = trimmed.charAt(trimmed.length - 1);
    if (quoteChars.indexOf(lastChar) === -1) continue;

    issueCount++;

    var trailingInfo = getOUTRETrailingQuoteInfo(description);
    var sizeTokens = extractOUTRESizeTokens(description);
    var match = matchOUTREDescriptionFromDB(description);

    output.push('--- ' + invoiceNo + ' / Line ' + lineNo + ' ---');
    output.push('Desc: ' + description);
    output.push('Memo: ' + (memo || '-'));
    output.push('TrailingQuote: ' + (trailingInfo.has ? 'YES(' + trailingInfo.char + ')' : 'NO'));
    output.push('SizeTokens: ' + (sizeTokens.length ? sizeTokens.join(', ') : '-'));

    if (match && match.description) {
      output.push('Match: ' + match.description + ' | type=' + match.matchType +
                  ' | score=' + (match.score || 0));
    } else if (match && match.altDescription) {
      output.push('Match: NONE | alt=' + match.altDescription +
                  ' | reason=' + (match.altReason || '-') +
                  ' | score=' + (match.altScore || 0));
    } else {
      output.push('Match: NONE');
    }

    output.push('');
  }

  output.unshift('Issues: ' + issueCount);
  writeDebugOutput(output.join('\n'));
  Logger.log('Debug done - see DEBUG_OUTPUT sheet.');
}

/**
 * Debug only the WAVY BOMB TWIST line from a specific file.
 * Usage: debugOUTREWavyBombTwistQuote()
 */
function debugOUTREWavyBombTwistQuote() {
  var TARGET_FILE_NAME = 'SINV1903556.docx';
  var TARGET_PHRASE = 'X-PRESSION - TWISTED UP - WAVY BOMB TWIST';

  var files = DriveApp.getFilesByName(TARGET_FILE_NAME);
  if (!files.hasNext()) {
    Logger.log('File not found: ' + TARGET_FILE_NAME);
    return;
  }

  var file = files.next();
  var text = extractTextFromDocx(file.getBlob());
  if (!text) {
    Logger.log('Text extraction failed for: ' + TARGET_FILE_NAME);
    return;
  }

  var lines = text.split(/\r?\n/);
  var output = [];
  var matches = 0;

  output.push('OUTRE Wavy Bomb Twist Debug');
  output.push('File: ' + TARGET_FILE_NAME);
  output.push('');

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    if (!line || line.indexOf(TARGET_PHRASE) === -1) continue;

    matches++;

    var trimmed = line.replace(/[\s\u200B-\u200D\uFEFF]+$/, '');
    var lastChar = trimmed ? trimmed.charAt(trimmed.length - 1) : '';
    var lastCode = lastChar ? lastChar.charCodeAt(0) : -1;

    var tail = trimmed.slice(Math.max(0, trimmed.length - 8));
    var tailCodes = [];
    for (var t = 0; t < tail.length; t++) {
      tailCodes.push('0x' + tail.charCodeAt(t).toString(16));
    }

    var trailingInfo = getOUTRETrailingQuoteInfo(trimmed);
    var sizeTokensInput = extractOUTRESizeTokens(trimmed);
    var match = matchOUTREDescriptionFromDB(trimmed);

    output.push('Line ' + i + ': ' + line);
    output.push('Trimmed: ' + trimmed);
    output.push('LastChar: ' + (lastChar ? "'" + lastChar + "'" : '(none)') +
                ' code=' + lastCode + ' hex=' + (lastCode >= 0 ? '0x' + lastCode.toString(16) : '-'));
    output.push('TailCodes: ' + (tailCodes.length ? tailCodes.join(', ') : '-'));
    output.push('TrailingQuoteInfo: has=' + (trailingInfo && trailingInfo.has ? 'YES' : 'NO') +
                ' char=' + (trailingInfo ? trailingInfo.char : '-'));
    output.push('InputSizeTokens: ' + (sizeTokensInput.length ? sizeTokensInput.join(', ') : '-'));

    if (match && match.description) {
      var dbSizeTokens = extractOUTRESizeTokens(match.description);
      output.push('DBMatch: ' + match.description + ' | type=' + match.matchType +
                  ' | score=' + (match.score || 0));
      output.push('DBSizeTokens: ' + (dbSizeTokens.length ? dbSizeTokens.join(', ') : '-'));
    } else if (match && match.altDescription) {
      var altSizeTokens = extractOUTRESizeTokens(match.altDescription);
      output.push('DBMatch: NONE | alt=' + match.altDescription +
                  ' | reason=' + (match.altReason || '-') +
                  ' | score=' + (match.altScore || 0));
      output.push('AltSizeTokens: ' + (altSizeTokens.length ? altSizeTokens.join(', ') : '-'));
    } else {
      output.push('DBMatch: NONE');
    }

    output.push('');
  }

  if (matches === 0) {
    output.push('No matching lines found.');
  }

  output.unshift('Matches: ' + matches);
  writeDebugOutput(output.join('\n'));
  Logger.log('Debug done - see DEBUG_OUTPUT sheet.');
}
