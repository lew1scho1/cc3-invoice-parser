// ============================================================================
// Invoice_Parser_OUTRE.js - OUTRE Invoice Parser
// ============================================================================
//
// OUTRE 인보이스 전용 파서
// - 다중 라인 구조: QTY, DESCRIPTION (1-2줄), COLORS (다중 줄), PRICES (3줄)
// - 컬러별 수량 분리 및 라인 아이템 생성
// - CRITICAL: 현재 완벽하게 작동 중 - 수정 금지!
//
// ============================================================================

/**
 * OUTRE 인보이스 라인 아이템 파싱
 * @param {Array<string>} lines - 인보이스 텍스트 라인 배열
 * @return {Array<Object>} 파싱된 라인 아이템 배열
 */
function parseOUTRELineItems(lines) {
  var items = [];
  var lineNo = 1;

  debugLog('OUTRE 라인 아이템 파싱 시작', { totalLines: lines.length });

  // CRITICAL: DB 캐시 초기화 (배치 파싱 성능 개선)
  initOUTREDBCache();

  var extractPriceValues = function(line) {
    if (!line) return [];

    var trimmed = line.trim();
    if (!trimmed) return [];

    // Accept comma-formatted prices like "1,375.00"
    var normalized = trimmed.replace(/\$/g, '');
    if (/[^0-9,.\s]/.test(normalized)) return [];

    var matches = normalized.match(/\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2}/g);
    if (!matches || matches.length === 0) return [];

    var values = [];
    for (var i = 0; i < matches.length; i++) {
      var value = parseFloat(matches[i].replace(/,/g, ''));
      if (!isNaN(value)) values.push(value);
    }

    return values;
  };

  var isPriceLine = function(line) {
    return extractPriceValues(line).length > 0;
  };

  var isQtyLine = function(line) {
    return /^\d{1,4}$/.test(line);
  };

  var isLikelyDescription = function(line) {
    var hasProductKeywords = line.match(/HAIR|WIG|LACE|WEAVE|CLIP|REMI|BATIK|SUGARPUNCH|X-PRESSION|BEAUTIFUL|MELTED|BRAID|CLOSURE|WAVE|CURL|STRAIGHT|BUNDLE|PONYTAIL|TARA|QW|BIG|BOHEMIAN|HD|PERUVIAN|TWIST|FEED|LOOKS|PASSION/i);
    var hasMetadata = line.match(/\bSHIP\s+TO\b|\bSOLD\s+TO\b|\bWEIGHT\b|\bSUBTOTAL\b|\bRICHMOND\b|\bLLC\b|\bPKWAY\b|\bCOD\b|\bFee\b|\btag\b|\bDATE\s+SHIPPED\b|\bPAGE\b|\bSHIP\s+VIA\b|\bPAYMENT\b|\bTERMS\b/i);
    var hasUpperCase = line.match(/[A-Z]/);
    var hasMinLength = line.length >= 3;

    return hasMinLength &&
      hasUpperCase &&
      (hasProductKeywords || line.length >= 5) &&
      !hasMetadata;
  };

  var isColorLine = function(line) {
    var hasColorPattern = line.match(/[A-Z0-9\-\/+]+\s*-\s*\d+/);
    var isInchPattern = line.match(/\d+["″'']/);
    var DESCRIPTION_BLACKLIST = ['X-PRESSION', 'SHAKE-N-GO', 'BATIK', 'SUGARPUNCH'];
    var hasBlacklistedWord = false;

    if (hasColorPattern) {
      var upperLine = line.toUpperCase();
      for (var bi = 0; bi < DESCRIPTION_BLACKLIST.length; bi++) {
        if (upperLine.indexOf(DESCRIPTION_BLACKLIST[bi]) > -1) {
          hasBlacklistedWord = true;
          break;
        }
      }
    }

    return hasColorPattern && !isInchPattern && !hasBlacklistedWord;
  };

  // OUTRE의 경우: 테이블 헤더를 찾아서 그 이후부터만 파싱
  var startLine = 0;

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
        Logger.log('=== 헤더 발견 후 첫 20줄 검사 시작 (라인 ' + i + ' 이후) ===');

        for (var k = i + 1; k < Math.min(i + 30, lines.length); k++) {
          var testLine = lines[k].trim();

          Logger.log('  [' + k + '] 길이=' + testLine.length + ' | ' + testLine.substring(0, 100));

          // OUTRE 다중 라인 형식: QTY만 있는 라인 찾기 (1~4자리 숫자만)
          if (isQtyLine(testLine)) {
            var qty = parseInt(testLine);

            Logger.log('    QTY 전용 라인 발견: ' + qty);

            // 헤더 이후 30줄 내에서 "Description + Color + Price(3)" 블록 확인
            var foundDescription = false;
            var foundColor = false;
            var priceCount = 0;

            for (var t = k + 1; t < Math.min(k + 20, lines.length); t++) {
              var blockLine = lines[t].trim();
              if (!blockLine) continue;

              if (isQtyLine(blockLine)) {
                break; // 다음 QTY로 넘어감
              }

              if (!foundDescription) {
                if (isLikelyDescription(blockLine)) {
                  foundDescription = true;
                }
                continue;
              }

              if (!foundColor) {
                if (isColorLine(blockLine)) {
                  foundColor = true;
                }
                continue;
              }

              if (isPriceLine(blockLine)) {
                priceCount += extractPriceValues(blockLine).length;
                if (priceCount >= 3) {
                  startLine = k;
                  Logger.log('  ✅ 테이블 시작 라인 확정 (패턴 매칭): ' + k);
                  debugLog('OUTRE 테이블 시작 라인 찾음 (패턴 매칭)', {
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

  // 이제 실제 아이템 파싱
  for (var i = startLine; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;

  // OUTRE 다중 라인 형식: QTY만 있는 라인 감지 (1~4자리 숫자만)
    if (line.match(/^\d{1,4}$/)) {
      var qty = parseInt(line);

      // 수량 범위 검증 (0-2000) + Description 검증
      if (qty >= 0 && qty <= 2000 && i + 1 < lines.length) {
        var nextLine = lines[i + 1].trim();

        // 다음 줄이 유효한 제품 Description인지 확인 (완화된 검증)
        var hasProductKeywords = nextLine.match(/HAIR|WIG|LACE|WEAVE|CLIP|REMI|BATIK|SUGARPUNCH|X-PRESSION|BEAUTIFUL|MELTED|BRAID|CLOSURE|WAVE|CURL|STRAIGHT|BUNDLE|PONYTAIL|TARA|QW|BIG|BOHEMIAN|HD|PERUVIAN|TWIST|FEED|LOOKS|PASSION/i);
        var hasMetadata = nextLine.match(/\bSHIP\s+TO\b|\bSOLD\s+TO\b|\bWEIGHT\b|\bSUBTOTAL\b|\bRICHMOND\b|\bLLC\b|\bPKWAY\b|\bCOD\b|\bFee\b|\btag\b|\bDATE\s+SHIPPED\b|\bPAGE\b|\bSHIP\s+VIA\b|\bPAYMENT\b|\bTERMS\b/i);

        // CRITICAL: 소문자 허용, 길이 체크 완화
        var hasUpperCase = nextLine.match(/[A-Z]/);  // 최소 1개 대문자만 있으면 OK
        var hasMinLength = nextLine.length >= 3;     // 최소 3자

        var isValidDescription = hasMinLength &&
                                hasUpperCase &&
                                (hasProductKeywords || nextLine.length >= 5) &&  // 5자 이상이면 키워드 불필요
                                !hasMetadata;

        if (isValidDescription) {
          // 아이템 파싱
          var result = parseOUTREItem(i, lines);

          if (result && result.items) {
            var itemsArray = Array.isArray(result.items) ? result.items : [result.items];

            for (var m = 0; m < itemsArray.length; m++) {
              itemsArray[m].lineNo = lineNo++;
              items.push(itemsArray[m]);
            }

            // CRITICAL: 처리한 라인 건너뛰기 (중복 방지)
            i = result.nextLineIndex - 1; // -1은 for 루프의 i++를 위함

            Logger.log('  ✅ 다음 파싱 시작 라인: ' + (i + 1));
          }
        }
      }
    }
  }

  debugLog('OUTRE 라인 아이템 파싱 완료', { totalItems: items.length });

  // CRITICAL: DB 캐시 리셋 (메모리 절약)
  resetOUTREDBCache();

  return items;
}

/**
 * OUTRE 개별 아이템 파싱 (개선 버전 v3 - DB 검증 추가)
 * CRITICAL: DB 검증 최우선, 중복 방지를 위한 nextLineIndex 반환
 *
 * @param {number} lineIndex - QTY 라인 인덱스
 * @param {Array<string>} lines - 전체 라인 배열
 * @return {Object} {items: Array<Object>|Object, nextLineIndex: number}
 */
function parseOUTREItem(lineIndex, lines) {
  var qtyShipped = parseInt(lines[lineIndex].trim()) || 0;
  var qtyOrdered = qtyShipped;
  var itemId = '';
  var isOUTREMetaLine = function(line) {
    if (!line) return false;

    var normalized = normalizeOutreText(line);
    if (!normalized) return false;

    var upper = normalized.toUpperCase();
    var hasPhoneKeyword = upper.indexOf('FAX') > -1 ||
      upper.indexOf('PHONE') > -1 ||
      upper.indexOf('TOLL FREE') > -1;
    var hasPhoneNumber = /\b\d{3}\s*[-.\s]?\s*\d{3}\s*[-.\s]?\s*\d{4}\b/.test(normalized);

    if (hasPhoneKeyword && hasPhoneNumber) {
      return true;
    }

    var hasHeaderKeyword = upper.indexOf('INVOICE DATE') > -1 ||
      upper.indexOf('ORDER DATE') > -1 ||
      upper.indexOf('SHIP VIA') > -1 ||
      upper.indexOf('TERMS') > -1 ||
      upper.indexOf('BILL TO') > -1 ||
      upper.indexOf('SHIP TO') > -1 ||
      upper.indexOf('SALES REP') > -1;
    var hasColon = normalized.indexOf(':') > -1;
    var hasDate = /\b\d{1,2}\/\d{1,2}(?:\/\d{2,4})?\b/.test(normalized);

    return hasHeaderKeyword && (hasColon || hasDate);
  };
  var extractPriceValues = function(line) {
    if (!line) return [];

    var trimmed = line.trim();
    if (!trimmed) return [];

    // Accept comma-formatted prices like "1,375.00"
    var normalized = trimmed.replace(/\$/g, '');
    if (/[^0-9,.\s]/.test(normalized)) return [];

    var matches = normalized.match(/\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2}/g);
    if (!matches || matches.length === 0) return [];

    var values = [];
    for (var i = 0; i < matches.length; i++) {
      var value = parseFloat(matches[i].replace(/,/g, ''));
      if (!isNaN(value)) values.push(value);
    }

    return values;
  };

  Logger.log('=== OUTRE 다중 라인 파싱 시작 (라인 ' + lineIndex + ', QTY=' + qtyShipped + ') ===');

  // 다음 15줄 안에서 DESCRIPTION, COLORS, PRICES 찾기
  var descriptionLines = [];
  var colorLinesArray = [];
  var priceLines = [];
  var foundFirstColor = false;
  var lastProcessedLine = lineIndex; // 마지막 처리 라인 추적

  for (var j = lineIndex + 1; j < Math.min(lineIndex + 15, lines.length); j++) {
    var nextLine = lines[j].trim();

    lastProcessedLine = j; // 현재 처리 중인 라인 기록

    if (!nextLine) {
      Logger.log('[' + j + '] (빈 줄)');
      continue;
    }

    Logger.log('[' + j + '] ' + nextLine.substring(0, 80));

    // Price line detection (supports comma-formatted prices)
    var priceValues = extractPriceValues(nextLine);
    var isPriceLine = priceValues.length > 0;

    if (isPriceLine) {
      for (var pv = 0; pv < priceValues.length; pv++) {
        priceLines.push(priceValues[pv]);
      }
      if (priceValues.length === 1) {
        Logger.log('  → 가격 라인 감지: $' + priceValues[0]);
      } else {
        Logger.log('  price line (multi): ' + priceValues.join(', '));
      }
      continue;
    }

    if (isOUTREMetaLine(nextLine)) {
      Logger.log('  skip meta line: ' + nextLine.substring(0, 80));
      continue;
    }

    // 다음 아이템 라인을 만나면 중단 (숫자만 있는 라인)
    if (nextLine.match(/^\d{1,4}$/)) {
      var nextQty = parseInt(nextLine);
      if (nextQty >= 0 && nextQty <= 2000) {
        Logger.log('  ✋ 다음 아이템 감지 (QTY=' + nextQty + '), 현재 아이템 파싱 중단');
        lastProcessedLine = j - 1; // 다음 아이템 라인 직전까지만 처리
        break;
      }
    }

    // ========================================
    // CRITICAL: Description 예외 패턴 처리 (최우선)
    // ========================================
    // REMI TARA 1-2-3 / 2-4-6 / 4-6-8 등은 숫자-숫자-숫자 패턴이지만
    // 컬러가 아닌 Description의 일부임
    // 컬러 판정보다 먼저 처리하여 오인식 방지
    var DESCRIPTION_EXCEPTION_PATTERNS = [
      { pattern: /REMI[\s\-]*TARA[\s\-]*\d+[\-\/]\d+[\-\/]\d+/i, name: 'REMI TARA' }
      // 향후 유사 케이스 추가 가능
    ];

    var isExceptionPattern = false;
    var exceptionName = '';

    for (var ei = 0; ei < DESCRIPTION_EXCEPTION_PATTERNS.length; ei++) {
      if (nextLine.match(DESCRIPTION_EXCEPTION_PATTERNS[ei].pattern)) {
        isExceptionPattern = true;
        exceptionName = DESCRIPTION_EXCEPTION_PATTERNS[ei].name;
        Logger.log('  ✅ Description 예외 패턴 감지: ' + exceptionName);
        break;
      }
    }

    // 예외 패턴인 경우 Description으로 확정
    if (isExceptionPattern) {
      // REMI TARA 전용: 숫자-숫자-숫자 패턴은 Description으로 고정
      var remiTaraMatch = nextLine.match(/^(REMI[\s\-]*TARA[\s\-]*\d+[\-\/]\d+[\-\/]\d+)\s+(.*)$/i);
      if (remiTaraMatch) {
        var remiDesc = remiTaraMatch[1].trim();
        var remiColors = remiTaraMatch[2].trim();
        descriptionLines.push(remiDesc);
        Logger.log('    → Description 추가 (' + exceptionName + '): ' + remiDesc.substring(0, 50));
        if (remiColors) {
          colorLinesArray.push(remiColors);
          foundFirstColor = true;
          Logger.log('    → 컬러 라인 추가: ' + remiColors.substring(0, 50));
        }
        continue; // 다음 라인으로
      }

      // 같은 라인에 컬러가 붙어 있는지 확인
      // 예: "REMI TARA 1-2-3 T30- 10 1B- 20"
      // → Description: "REMI TARA 1-2-3"
      // → 컬러: "T30- 10 1B- 20"
      var split = splitDescriptionAndColor(nextLine, false);

      if (split.color) {
        // Description + 컬러 혼재
        descriptionLines.push(split.description);
        colorLinesArray.push(split.color);
        foundFirstColor = true;
        Logger.log('    → Description 추가 (' + exceptionName + '): ' + split.description.substring(0, 50));
        Logger.log('    → 컬러 라인 추가: ' + split.color.substring(0, 50));
      } else {
        // Description만
        descriptionLines.push(nextLine);
        Logger.log('    → Description 추가 (' + exceptionName + '): ' + nextLine.substring(0, 50));
      }

      continue; // 다음 라인으로
    }

    // ========================================
    // 일반 컬러 패턴 감지
    // ========================================
    var hasColorPattern = nextLine.match(/[A-Z0-9\-\/+]+\s*-\s*\d+/);
    var isInchPattern = nextLine.match(/\d+["″'']/);

    // 블랙리스트: Description인데 컬러 패턴처럼 보이는 경우 제외
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

    // 짧은 Description 보조 라인 (예: 10X, 3X, 18", 3PCS)
    var isShortDescContinuation = !foundFirstColor &&
      !isPriceLine &&
      nextLine.match(/^(?:\d{1,2}X|\d{1,2}"|\d+\s*PCS?|\d{1,2}\s*IN(?:CH)?)$/i);

    if (isShortDescContinuation) {
      descriptionLines.push(nextLine);
      Logger.log('  → Description 보조 라인 추가');
      continue;
    }

    // Description 후보 판단
    var isDescriptionCandidate = !foundFirstColor && !isPriceLine && nextLine.length > 5;

    // STAGE 1: 괄호 컬러 패턴 우선 인식 (전용 분리 함수 사용)
    // 예: "X-PRESSION BRAID 52" 3X (P)M950/425/350/130S- 55"
    // → Description: "X-PRESSION BRAID 52" 3X"
    // → 컬러: "(P)M950/425/350/130S- 55"
    var hasParenColorPattern = nextLine.match(/\([A-Z]\)[A-Z0-9\-\/+]+\s*-\s*\d+/);

    if (isDescriptionCandidate && hasParenColorPattern) {
      Logger.log('  ✅ STAGE 1: 괄호 컬러 패턴 감지, Description/컬러 분리 진행');

      // CRITICAL: 괄호 컬러 전용 함수 사용 (일반 함수와 완전 분리)
      var split = extractColorsFromParenthesizedLine(nextLine);

      if (split.description) {
        descriptionLines.push(split.description);
        Logger.log('    → Description 추가: ' + split.description.substring(0, 50));
      }

      if (split.color) {
        colorLinesArray.push(split.color);
        foundFirstColor = true;
        Logger.log('    → 컬러 라인 추가: ' + split.color.substring(0, 50));
      }

      continue; // 다음 라인으로 (분리 완료)
    }

    if (isDescriptionCandidate && !isColorLine) {
      // Description만 추가
      descriptionLines.push(nextLine);
      Logger.log('  → Description 라인 추가');
      continue;
    } else if (isDescriptionCandidate && isColorLine) {
      // CRITICAL: 컬러 라인은 절대 Description에 넣지 않음 (DB 검증 오염 방지)
      // 혼재 라인은 splitDescriptionAndColor()로 분리됨
      foundFirstColor = true;
      colorLinesArray.push(nextLine);
      Logger.log('  → 컬러 라인 추가 (Description 종료)');
      continue;
    }

    // 컬러 라인 수집
    if (isColorLine) {
      colorLinesArray.push(nextLine);
      foundFirstColor = true;
      Logger.log('  → 컬러 라인 추가');
      continue;
    }

    // CRITICAL: 단독 괄호 라인 수집 (멀티라인 백오더)
    // foundFirstColor=true 이후 ^\(\d+\)$ 패턴은 이전 컬러의 백오더로 수집
    // 예: S4/30- 0 다음 줄의 (1)
    if (foundFirstColor && nextLine.match(/^\(\d+\)$/)) {
      colorLinesArray.push(nextLine);
      Logger.log('  → 단독 괄호 라인 추가 (멀티라인 백오더)');
    }
  }

  // Description 결합
  var rawDescription = descriptionLines.join(' ').trim();

  Logger.log('수집 완료:');
  Logger.log('  Description 라인: ' + descriptionLines.length);
  Logger.log('  컬러 라인: ' + colorLinesArray.length);
  Logger.log('  가격 라인: ' + priceLines.length);

  // STAGE 3: 사후 보정 안전망 (colorLinesArray === 0일 때만 실행)
  // STAGE 1에서 처리 실패 시 Description에서 괄호 컬러 추출 시도
  if (colorLinesArray.length === 0 && descriptionLines.length > 0) {
    var lastDescLine = descriptionLines[descriptionLines.length - 1];
    var parenColorMatch = lastDescLine.match(/\([A-Z]\)[A-Z0-9\-\/+]+\s*-\s*\d+/);

    if (parenColorMatch) {
      Logger.log('  ⚠️ STAGE 3: 사후 보정 - Description에서 괄호 컬러 추출 시도');
      Logger.log('    마지막 Description 라인: ' + lastDescLine.substring(0, 80));

      // CRITICAL: 괄호 컬러 전용 함수 사용 (일반 함수와 완전 분리)
      var split = extractColorsFromParenthesizedLine(lastDescLine);

      if (split.description && split.color) {
        // Description 라인 교체
        descriptionLines[descriptionLines.length - 1] = split.description;
        Logger.log('    → Description 업데이트: ' + split.description.substring(0, 50));

        // 컬러 라인 추가
        colorLinesArray.push(split.color);
        Logger.log('    → 컬러 복구: ' + split.color.substring(0, 50));
      }
    }
  }

  // Description 재결합 (STAGE 3에서 수정되었을 수 있음)
  var rawDescription = descriptionLines.join(' ').trim();

  if (!rawDescription) {
    Logger.log('  ⚠️ Description 없음');
    return { items: null, nextLineIndex: lastProcessedLine + 1 };
  }

  Logger.log('  📝 원본 Description: ' + rawDescription.substring(0, 80));

  var description = '';
  var descriptionMatchType = 'none';
  var descriptionMatchScore = 0;
  var descriptionSizeMismatch = false;

  // ✅ STEP 1: DB 검증 최우선
  var dbMatch = matchOUTREDescriptionFromDB(rawDescription);

  if (dbMatch && dbMatch.description) {
    description = dbMatch.description;
    descriptionMatchType = dbMatch.matchType || 'exact';
    descriptionMatchScore = dbMatch.score || (descriptionMatchType === 'exact' ? 1 : 0);
    descriptionSizeMismatch = !!dbMatch.sizeMismatch;

    if (descriptionMatchType === 'fuzzy') {
      Logger.log('  ✅ DB 유사 매칭: ' + description.substring(0, 80) +
                 ' (score=' + descriptionMatchScore.toFixed(2) + ')');
    } else {
      Logger.log('  ✅ DB 검증 성공: ' + description.substring(0, 80));
    }

    // DB 매칭 성공 시 컬러 분리하지 않음 (DB Description 그대로 사용)
    // 컬러 라인은 별도 수집
  } else {
    // ✅ STEP 2: DB 미매칭 시 파싱 로직으로 Description 처리
    Logger.log('  ⚠️ DB 미매칭, 파싱 로직으로 Description 처리');

    var split = splitDescriptionAndColor(rawDescription);
    description = split.description;
    var colorInDescLine = split.color;

    Logger.log('  🔧 Description 분리 결과:');
    Logger.log('    Description: ' + description.substring(0, 80));
    if (colorInDescLine) {
      Logger.log('    라인 내 컬러: ' + colorInDescLine.substring(0, 80));
      // Description 라인에 포함된 컬러를 컬러 라인 배열에 추가
      colorLinesArray.unshift(colorInDescLine);
      Logger.log('    ✅ 컬러 라인 배열에 추가됨 (총 ' + colorLinesArray.length + '줄)');
    }

    // Description 끝부분 정리 (보수적)
    var descriptionBeforeCleanup = description;
    var preserveNumberPattern = /REMI[\s\-]*TARA/i.test(description);
    description = cleanDescriptionEnd(description, preserveNumberPattern);

    if (description !== descriptionBeforeCleanup) {
      Logger.log('  🔧 Description 끝부분 정리: ' + description.substring(0, 80));
    }
  }

  var descriptionMatchMemo = '';
  if (descriptionMatchType === 'none') {
    if (dbMatch && dbMatch.altDescription) {
      descriptionMatchMemo = '⚠️ DESC 미매칭 (DB 후보: ' + dbMatch.altDescription + ')';
    } else {
      descriptionMatchMemo = '⚠️ DESC 미매칭';
    }
  }

  Logger.log('  📝 최종 Description: ' + description.substring(0, 80));

  // 가격 정보 (최소 3개 필요: UNIT, DISC, EXT)
  var unitPrice = 0;
  var extPrice = 0;

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

  // 컬러 정보 처리
  var colorLines = [];
  if (colorLinesArray.length > 0) {
    colorLines = colorLinesArray;
    Logger.log('  ✅ 컬러 라인 설정: ' + colorLinesArray.length + '줄');
  } else {
    colorLines = [];
    Logger.log('  ⚠️ 컬러 라인 없음');
  }

  debugLog('OUTRE 아이템 파싱 결과', {
    line: lineIndex,
    itemId: itemId,
    description: description,
    qtyOrdered: qtyOrdered,
    qtyShipped: qtyShipped,
    unitPrice: unitPrice,
    extPrice: extPrice,
    colorCount: colorLines.length
  });

  // 컬러 라인이 있으면 parseOUTREColorLines로 파싱하여 개별 아이템 생성
  if (colorLines.length > 0) {
    var colorData = parseOUTREColorLines(colorLines, description);

    Logger.log('컬러 파싱 결과: ' + colorData.length + '개 컬러');

    if (colorData.length > 0) {
      var totalShipped = 0;
      for (var m = 0; m < colorData.length; m++) {
        totalShipped += colorData[m].shipped;
      }

      var items = [];
      var sumExtPrice = 0; // ExtPrice 합계 추적

      for (var m = 0; m < colorData.length; m++) {
        var cd = colorData[m];

        var itemExtPrice = 0;

        // CRITICAL: 마지막 컬러는 나머지 할당 (반올림 오차 제거)
        if (m === colorData.length - 1) {
          itemExtPrice = Number((extPrice - sumExtPrice).toFixed(2));
          Logger.log('  마지막 컬러 ' + cd.color + ' ExtPrice 나머지 할당: $' + itemExtPrice);
        } else {
          if (totalShipped > 0) {
            itemExtPrice = Number((extPrice * (cd.shipped / totalShipped)).toFixed(2));
          }
          sumExtPrice += itemExtPrice;
        }

        var memoText = cd.backordered > 0 ? 'Backordered: ' + cd.backordered : '';
        memoText = appendOUTREMemo(memoText, descriptionMatchMemo);

        var item = {
          lineNo: 0, // 나중에 설정
          itemId: itemId,
          upc: '',
          description: description,
          brand: CONFIG.INVOICE.BRANDS['OUTRE'],
          color: cd.color,
          sizeLength: '',
          qtyOrdered: cd.shipped + cd.backordered,
          qtyShipped: cd.shipped,
          unitPrice: unitPrice,
          extPrice: itemExtPrice,
          memo: memoText
        };

        // CRITICAL: Item Number 보강 (Description만)
        item = enrichOUTREItemNumber(item);

        // CRITICAL: UPC 보강 (Description + Color, Item Number와 독립)
        item = enrichOUTREUPC(item);

        if (!item.itemId) {
          item.memo = appendOUTREMemo(item.memo, '⚠️ ITEM NO 미매칭');
        }
        if (item.color && !item.upc) {
          item.memo = appendOUTREMemo(item.memo, '⚠️ UPC 미매칭');
        }

        items.push(item);
      }

      // 반환값 구조 변경: {items, nextLineIndex}
      return { items: items, nextLineIndex: lastProcessedLine + 1 };
    }
  }

  // 컬러 라인이 없으면 단일 아이템으로 추가
  var memoText = colorLines.length === 0 ? '⚠️ 컬러 정보 찾을 수 없음' : '';
  memoText = appendOUTREMemo(memoText, descriptionMatchMemo);

  var item = {
    lineNo: 0, // 나중에 설정
    itemId: itemId,
    upc: '',
    description: description,
    brand: CONFIG.INVOICE.BRANDS['OUTRE'],
    color: '',
    sizeLength: '',
    qtyOrdered: qtyOrdered,
    qtyShipped: qtyShipped,
    unitPrice: unitPrice,
    extPrice: extPrice,
    memo: memoText
  };

  // CRITICAL: Item Number 보강 (Description만)
  item = enrichOUTREItemNumber(item);

  // CRITICAL: UPC 보강 (Color 없으므로 스킵됨)
  item = enrichOUTREUPC(item);

  if (!item.itemId) {
    item.memo = appendOUTREMemo(item.memo, '⚠️ ITEM NO 미매칭');
  }

  // 반환값 구조 변경: {items, nextLineIndex}
  return { items: item, nextLineIndex: lastProcessedLine + 1 };
}

/**
 * OUTRE Item Number 보강 (Description만 필요)
 * CRITICAL: Description만 매칭, Color 무관
 *
 * @param {Object} item - 라인 아이템
 * @return {Object} Item Number가 보강된 아이템
 */
function enrichOUTREItemNumber(item) {
  if (!item.description) return item;

  // 캐시 초기화 확인
  if (OUTRE_DB_CACHE === null) {
    initOUTREDBCache();
  }

  // 캐시 오류 시 스킵
  if (OUTRE_DB_CACHE.error) {
    return item;
  }

  try {
    var dbMap = OUTRE_DB_CACHE.dbMap;

    // 정규화 함수
    var normalize = function(text) {
      if (!text) return '';
      return text.toString()
        .trim()
        .replace(/["″''`]/g, '"')
        .replace(/\s+/g, ' ')
        .replace(/\-+/g, '-')
        .replace(/\s*-\s*/g, '-')
        .toUpperCase();
    };

    var normalizedDesc = normalize(item.description);
    var matchedRecords = dbMap[normalizedDesc];

    if (matchedRecords && matchedRecords.length > 0) {
      // 첫 번째 레코드의 Item Number 사용
      item.itemId = matchedRecords[0].itemNumber || '';
      if (item.itemId) {
        Logger.log('  ✅ Item Number 보강: ' + item.itemId);
      }
    } else {
      Logger.log('  ⚠️ Item Number 없음 (Description 미매칭)');
    }

  } catch (error) {
    Logger.log('❌ Item Number 보강 오류: ' + error.toString());
  }

  return item;
}

/**
 * OUTRE UPC 보강 (Description + Color 필요)
 * CRITICAL: Description + Color 모두 매칭, Item Number와 독립적
 *
 * @param {Object} item - 라인 아이템
 * @return {Object} UPC가 보강된 아이템
 */
function enrichOUTREUPC(item) {
  if (!item.description || !item.color) {
    if (!item.color) {
      Logger.log('  ⚠️ COLOR 없음, UPC 보강 스킵');
    }
    return item;
  }

  // 캐시 초기화 확인
  if (OUTRE_DB_CACHE === null) {
    initOUTREDBCache();
  }

  // 캐시 오류 시 스킵
  if (OUTRE_DB_CACHE.error) {
    return item;
  }

  try {
    var dbMap = OUTRE_DB_CACHE.dbMap;

    // Description 정규화 함수
    var normalizeDesc = function(text) {
      if (!text) return '';
      return text.toString()
        .trim()
        .replace(/["″''`]/g, '"')
        .replace(/\s+/g, ' ')
        .replace(/\-+/g, '-')
        .replace(/\s*-\s*/g, '-')
        .toUpperCase();
    };

    // Color 정규화 함수 (슬래시 공백 제거 추가)
    var normalizeColor = function(text) {
      if (!text) return '';
      return text.toString()
        .trim()
        .replace(/["″''`]/g, '"')
        .replace(/\s+/g, ' ')
        .replace(/\s*\/\s*/g, '/')    // CRITICAL: 슬래시 앞뒤 공백 제거
        .replace(/\-+/g, '-')
        .replace(/\s*-\s*/g, '-')
        .toUpperCase();
    };

    // Step 1: Description 기준 레코드 조회
    var normalizedDesc = normalizeDesc(item.description);
    var matchedRecords = dbMap[normalizedDesc];

    if (!matchedRecords || matchedRecords.length === 0) {
      Logger.log('  ⚠️ UPC 없음 (Description 미매칭)');
      return item;
    }

    // Step 2: Color 기준 필터링
    var normalizedColor = normalizeColor(item.color);
    var colorMatchedRecords = [];

    for (var i = 0; i < matchedRecords.length; i++) {
      var dbColor = normalizeColor(matchedRecords[i].color);

      if (dbColor === normalizedColor) {
        colorMatchedRecords.push(matchedRecords[i]);
      }
    }

    // Step 3: 매칭 결과 처리
    if (colorMatchedRecords.length === 0) {
      // 매칭 실패 → 파싱 컬러 유지 + 경고
      Logger.log('  ⚠️ DB 미등록 컬러: ' + item.color);
      Logger.log('    DB에 있는 컬러 목록 (최대 5개):');
      for (var i = 0; i < Math.min(matchedRecords.length, 5); i++) {
        Logger.log('      - ' + matchedRecords[i].color);
      }
      item.memo = (item.memo ? item.memo + ' / ' : '') + '⚠️ DB 미등록 컬러';

    } else if (colorMatchedRecords.length === 1) {
      // 단일 매칭 → DB 컬러로 확정 + UPC/Item Number 보강
      item.color = colorMatchedRecords[0].color;         // CRITICAL: DB 컬러로 덮어쓰기
      item.upc = colorMatchedRecords[0].barcode || '';
      item.itemId = colorMatchedRecords[0].itemNumber || item.itemId;  // Item Number도 확정
      Logger.log('  ✅ 컬러 확정: ' + item.color);
      Logger.log('  ✅ UPC 보강: ' + item.upc);
      Logger.log('  ✅ Item Number 확정: ' + item.itemId);

    } else {
      // 복수 매칭 → 파싱 컬러 유지 + 경고
      Logger.log('  ⚠️ 컬러 다중 매칭: ' + colorMatchedRecords.length + '개');
      Logger.log('    파싱 컬러: ' + item.color);
      for (var i = 0; i < Math.min(colorMatchedRecords.length, 3); i++) {
        Logger.log('      - ' + colorMatchedRecords[i].color + ' (UPC: ' + colorMatchedRecords[i].barcode + ')');
      }
      item.memo = (item.memo ? item.memo + ' / ' : '') + '⚠️ 컬러 다중 매칭';
    }

  } catch (error) {
    Logger.log('❌ UPC 보강 오류: ' + error.toString());
  }

  return item;
}

/**
 * OUTRE DB 캐시 (전역 변수)
 * CRITICAL: 배치 파싱 성능 개선을 위해 DB 데이터를 한 번만 로드
 */
var OUTRE_DB_CACHE = null;

/**
 * OUTRE DB 캐시 초기화
 * CRITICAL: parseOUTRELineItems() 시작 시 한 번만 호출
 */
function initOUTREDBCache() {
  if (OUTRE_DB_CACHE !== null) {
    return; // 이미 초기화됨
  }

  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.COMPANIES.OUTRE.dbSheet);

    if (!sheet) {
      Logger.log('⚠️ OUTRE DB 시트 없음');
      OUTRE_DB_CACHE = { error: true };
      return;
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      Logger.log('⚠️ OUTRE DB 데이터 없음');
      OUTRE_DB_CACHE = { error: true };
      return;
    }

    // 컬럼 인덱스 찾기
    var headers = data[0];
    var colMap = {};

    for (var i = 0; i < headers.length; i++) {
      colMap[headers[i]] = i;
    }

    var itemNameCol = colMap[CONFIG.COMPANIES.OUTRE.columns.ITEM_NAME];
    var itemNumberCol = colMap[CONFIG.COMPANIES.OUTRE.columns.ITEM_NUMBER];
    var colorCol = colMap[CONFIG.COMPANIES.OUTRE.columns.COLOR];
    var barcodeCol = colMap[CONFIG.COMPANIES.OUTRE.columns.BARCODE];

    if (itemNameCol === undefined) {
      Logger.log('⚠️ ITEM NAME 컬럼 없음');
      OUTRE_DB_CACHE = { error: true };
      return;
    }

    // 정규화 함수
    var normalize = function(text) {
      if (!text) return '';
      return text.toString()
        .trim()
        .replace(/["″''`]/g, '"')   // 인치 기호 통일
        .replace(/\s+/g, ' ')       // 다중 공백 → 단일 공백
        .replace(/\-+/g, '-')       // 다중 하이픈 → 단일 하이픈
        .replace(/\s*-\s*/g, '-')   // 하이픈 앞뒤 공백 제거
        .toUpperCase();
    };

    // CRITICAL: DB를 Map으로 변환 (정규화된 Description → 레코드 배열)
    // 동일 Description에 여러 Color가 있을 수 있으므로 배열 구조 사용
    var dbMap = {};
    for (var i = 1; i < data.length; i++) {
      var dbItemName = data[i][itemNameCol];
      if (!dbItemName) continue;

      var normalizedDB = normalize(dbItemName);

      // 배열로 저장 (동일 Description, 다른 Color 지원)
      if (!dbMap[normalizedDB]) {
        dbMap[normalizedDB] = [];
      }

      dbMap[normalizedDB].push({
        description: dbItemName.toString().trim(),
        itemNumber: itemNumberCol !== undefined ? (data[i][itemNumberCol] || '') : '',
        color: colorCol !== undefined ? (data[i][colorCol] || '') : '',
        barcode: barcodeCol !== undefined ? (data[i][barcodeCol] || '') : ''
      });
    }

    OUTRE_DB_CACHE = {
      error: false,
      dbMap: dbMap,
      columnMap: colMap
    };

    Logger.log('✅ OUTRE DB 캐시 초기화 완료: ' + Object.keys(dbMap).length + '개 Description (' +
               (data.length - 1) + '개 레코드)');

  } catch (error) {
    Logger.log('❌ OUTRE DB 캐시 초기화 오류: ' + error.toString());
    OUTRE_DB_CACHE = { error: true };
  }
}

/**
 * OUTRE DB 캐시 리셋
 * CRITICAL: parseOUTRELineItems() 종료 시 호출 (메모리 절약)
 */
function resetOUTREDBCache() {
  OUTRE_DB_CACHE = null;
}

/**
 * OUTRE 사이즈/팩 토큰 추출
 * @param {string} text - 원본 Description
 * @return {Array<string>} 정규화된 사이즈 토큰 배열
 */
function extractOUTRESizeTokens(text) {
  if (!text) return [];

  var tokens = [];
  var match;
  var normalizedText = text.toString()
    .replace(/[\u201C\u201D\u2033]/g, '"');

  // 1) 인치 표기: 12", 12″, 12”
  var inchQuotePattern = /\b(\d{1,2})\s*"(?!\w)/g;
  while ((match = inchQuotePattern.exec(normalizedText)) !== null) {
    tokens.push(match[1] + '"');
  }

  // 2) 인치 표기: 12 IN, 12 INCH
  var inchWordPattern = /\b(\d{1,2})\s*IN(?:CH)?\b/gi;
  while ((match = inchWordPattern.exec(normalizedText)) !== null) {
    tokens.push(match[1] + '"');
  }

  // 3) 배수 표기: 3X
  var xPattern = /\b(\d{1,2})\s*X\b/gi;
  while ((match = xPattern.exec(normalizedText)) !== null) {
    tokens.push(match[1] + 'X');
  }

  // 4) 팩 수량: 3PCS, 3PC
  var pcsPattern = /\b(\d+)\s*PCS?\b/gi;
  while ((match = pcsPattern.exec(normalizedText)) !== null) {
    tokens.push(match[1] + 'PCS');
  }

  // 중복 제거
  var uniq = {};
  var result = [];
  for (var i = 0; i < tokens.length; i++) {
    if (!uniq[tokens[i]]) {
      uniq[tokens[i]] = true;
      result.push(tokens[i]);
    }
  }

  return result;
}

/**
 * OUTRE Description 토큰화 (유사 매칭용)
 * @param {string} text - 정규화된 Description
 * @param {string} rawForSize - 사이즈 토큰 추출용 원문
 * @return {Object} {tokens, baseTokens, sizeTokens}
 */
function tokenizeOUTREDescriptionForMatch(text, rawForSize) {
  if (!text) {
    return { tokens: [], baseTokens: [], sizeTokens: [] };
  }

  var sizeTokens = extractOUTRESizeTokens(rawForSize || text);
  var sizeTokenMap = {};
  var sizeNumberMap = {};

  for (var si = 0; si < sizeTokens.length; si++) {
    sizeTokenMap[sizeTokens[si]] = true;
    var numberMatch = sizeTokens[si].match(/^(\d{1,2})"/);
    if (numberMatch) {
      sizeNumberMap[numberMatch[1]] = true;
    }
  }

  var normalized = text.toString()
    .toUpperCase()
    .replace(/["'\u2018\u2019\u201C\u201D\u2033`]/g, '"')
    .replace(/\s+/g, ' ')
    .replace(/\s*-\s*/g, '-')
    .replace(/[^A-Z0-9+\/\-"\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  var rawTokens = normalized ? normalized.split(' ') : [];
  var tokens = [];
  var baseTokens = [];

  for (var i = 0; i < rawTokens.length; i++) {
    var token = rawTokens[i];
    if (!token || token.length < 2) continue;

    tokens.push(token);
    var isSizeToken = false;
    if (sizeTokenMap[token]) {
      isSizeToken = true;
    } else if (/^\d{1,2}$/.test(token) && sizeNumberMap[token]) {
      isSizeToken = true;
    } else if (isOUTRESizeToken(token)) {
      isSizeToken = true;
    }

    if (!isSizeToken) {
      baseTokens.push(token);
    }
  }

  return {
    tokens: tokens,
    baseTokens: baseTokens,
    sizeTokens: sizeTokens
  };
}

/**
 * Description 끝단의 불완전 따옴표 감지
 * @param {string} text - 원본 Description
 * @return {Object} {has: boolean, char: string}
 */
function getOUTRETrailingQuoteInfo(text) {
  if (!text) {
    return { has: false, char: '' };
  }

  var trimmed = text.replace(/[\s\u200B-\u200D\uFEFF]+$/, '');
  if (!trimmed) {
    return { has: false, char: '' };
  }

  var lastChar = trimmed.charAt(trimmed.length - 1);
  var quoteChars = ['"', "'", '\u2019', '\u2018', '`', '\u00B4', '\u201C', '\u201D'];

  if (quoteChars.indexOf(lastChar) === -1) {
    return { has: false, char: '' };
  }

  var prevChar = trimmed.length > 1 ? trimmed.charAt(trimmed.length - 2) : '';
  if (/\d/.test(prevChar)) {
    return { has: false, char: '' };
  }

  return { has: true, char: lastChar };
}

/**
 * OUTRE 사이즈/팩 토큰 판단
 * @param {string} token - 토큰 문자열
 * @return {boolean} 사이즈/팩 토큰 여부
 */
function isOUTRESizeToken(token) {
  if (!token) return false;
  return !!(token.match(/^\d{1,2}"$/) ||
            token.match(/^\d{1,2}X$/) ||
            token.match(/^\d+\s*PCS?$/) ||
            token.match(/^\d{1,2}\s*IN(?:CH)?$/));
}

/**
 * 토큰 유사도 계산 (Jaccard)
 * @param {Array<string>} tokensA
 * @param {Array<string>} tokensB
 * @return {number} 0~1
 */
function scoreOUTRETokenSimilarity(tokensA, tokensB) {
  if (!tokensA || !tokensB) return 0;

  var setA = {};
  var setB = {};

  for (var i = 0; i < tokensA.length; i++) {
    setA[tokensA[i]] = true;
  }
  for (var j = 0; j < tokensB.length; j++) {
    setB[tokensB[j]] = true;
  }

  var unionCount = 0;
  var intersectionCount = 0;

  for (var key in setA) {
    unionCount++;
    if (setB[key]) {
      intersectionCount++;
    }
  }

  for (var key in setB) {
    if (!setA[key]) {
      unionCount++;
    }
  }

  return unionCount === 0 ? 0 : (intersectionCount / unionCount);
}

/**
 * 토큰 세트 동일 여부 확인
 * @param {Array<string>} tokensA
 * @param {Array<string>} tokensB
 * @return {boolean} 동일 여부
 */
function areOUTRETokenSetsEqual(tokensA, tokensB) {
  if (!tokensA || !tokensB) return false;
  if (tokensA.length !== tokensB.length) return false;

  var setA = {};
  for (var i = 0; i < tokensA.length; i++) {
    setA[tokensA[i]] = true;
  }

  for (var j = 0; j < tokensB.length; j++) {
    if (!setA[tokensB[j]]) {
      return false;
    }
  }

  return true;
}

/**
 * OUTRE DB 유사 매칭 인덱스 준비
 * @return {Array<Object>|null}
 */
function ensureOUTREFuzzyIndex() {
  if (OUTRE_DB_CACHE === null) {
    initOUTREDBCache();
  }

  if (!OUTRE_DB_CACHE || OUTRE_DB_CACHE.error) {
    return null;
  }

  if (OUTRE_DB_CACHE.fuzzyIndex) {
    return OUTRE_DB_CACHE.fuzzyIndex;
  }

  var dbMap = OUTRE_DB_CACHE.dbMap;
  var index = [];

  for (var key in dbMap) {
    if (!dbMap.hasOwnProperty(key)) continue;

    var rawDesc = dbMap[key][0].description || key;
    var tokenData = tokenizeOUTREDescriptionForMatch(key, rawDesc);
    index.push({
      normalized: key,
      description: dbMap[key][0].description,
      tokens: tokenData.tokens,
      baseTokens: tokenData.baseTokens,
      sizeTokens: tokenData.sizeTokens
    });
  }

  OUTRE_DB_CACHE.fuzzyIndex = index;
  return index;
}

/**
 * OUTRE DB에서 Description 매칭 (정확 일치 + 유사 매칭)
 *
 * @param {string} rawDescription - 파싱된 원본 Description
 * @return {Object|null} {description, matchType, score, sizeMismatch}
 */
function matchOUTREDescriptionFromDB(rawDescription) {
  if (!rawDescription) return null;

  // 캐시 초기화 확인
  if (OUTRE_DB_CACHE === null) {
    initOUTREDBCache();
  }

  // 캐시 오류 시 null 반환
  if (OUTRE_DB_CACHE.error) {
    return null;
  }

  try {
    var dbMap = OUTRE_DB_CACHE.dbMap;
    var trailingQuoteInfo = getOUTRETrailingQuoteInfo(rawDescription);

    // CRITICAL: Step 1 - 컬러 패턴 제거 (안전장치)
    // REMI TARA 1-2-3 등은 제품명 자체가 숫자 패턴이므로 컬러 제거 스킵
    var skipColorStrip = /REMI[\s\-]*TARA[\s\-]*\d+[\-\/]\d+[\-\/]\d+/i.test(rawDescription);
    var descriptionForDB = rawDescription;
    if (!skipColorStrip) {
      var colorPattern = /(\([A-Z]\))?([A-Z0-9\/+]+)\s*-\s*\d+(?:\s*\(\d+\))?/gi;
      descriptionForDB = rawDescription.replace(colorPattern, function(match, _prefix, token) {
        // validateOUTREColorToken()로 유효한 컬러만 제거
        if (validateOUTREColorToken(token)) {
          return ' ';  // 공백으로 대체
        }
        return match;  // 그 외는 유지
      });
    }

    // 정규화 함수 (캐시 초기화와 동일)
    var normalize = function(text) {
      if (!text) return '';
      return text.toString()
        .trim()
        .replace(/["″''`]/g, '"')   // 인치 기호 통일
        .replace(/\s+/g, ' ')       // 다중 공백 → 단일 공백
        .replace(/\-+/g, '-')       // 다중 하이픈 → 단일 하이픈
        .replace(/\s*-\s*/g, '-')   // 하이픈 앞뒤 공백 제거
        .toUpperCase();
    };

    var normalizedInput = normalize(descriptionForDB);

    // CRITICAL: Map 조회 (O(1) 시간 복잡도)
    // 배열 구조로 변경됨: dbMap[key] = [{description, itemNumber, color, barcode}, ...]
    var matchedRecords = dbMap[normalizedInput];

    if (matchedRecords && matchedRecords.length > 0) {
      // 첫 번째 레코드의 Description 반환
      var matchedDescription = matchedRecords[0].description;
      Logger.log('✅ DB 캐시 매칭 성공: ' + matchedDescription + ' (' + matchedRecords.length + '개 레코드)');
      return {
        description: matchedDescription,
        matchType: 'exact',
        score: 1,
        sizeMismatch: false
      };
    }

    // 유사 매칭 시도
    var fuzzyIndex = ensureOUTREFuzzyIndex();
    if (!fuzzyIndex || fuzzyIndex.length === 0) {
      Logger.log('⚠️ DB 유사 매칭 인덱스 없음');
      return null;
    }

    var inputTokens = tokenizeOUTREDescriptionForMatch(normalizedInput, descriptionForDB);

    var best = {
      entry: null,
      baseScore: 0,
      fullScore: 0,
      sizeMismatch: false
    };
    var bestRejected = {
      entry: null,
      baseScore: 0,
      fullScore: 0,
      sizeMismatch: false
    };

    for (var i = 0; i < fuzzyIndex.length; i++) {
      var entry = fuzzyIndex[i];
      var baseScore = scoreOUTRETokenSimilarity(inputTokens.baseTokens, entry.baseTokens);
      if (baseScore < 0.7) continue;

      var fullScore = scoreOUTRETokenSimilarity(inputTokens.tokens, entry.tokens);
      var sizeMismatch = !areOUTRETokenSetsEqual(inputTokens.sizeTokens, entry.sizeTokens);
      var inputHasSize = inputTokens.sizeTokens.length > 0;
      var entryHasSize = entry.sizeTokens.length > 0;

      if (trailingQuoteInfo.has) {
        if ((entryHasSize && !inputHasSize) || sizeMismatch) {
          if (baseScore > bestRejected.baseScore ||
              (baseScore === bestRejected.baseScore && fullScore > bestRejected.fullScore)) {
            bestRejected = {
              entry: entry,
              baseScore: baseScore,
              fullScore: fullScore,
              sizeMismatch: sizeMismatch
            };
          }
          continue;
        }
      }

      if (baseScore > best.baseScore ||
          (baseScore === best.baseScore && fullScore > best.fullScore)) {
        best = {
          entry: entry,
          baseScore: baseScore,
          fullScore: fullScore,
          sizeMismatch: sizeMismatch
        };
      }
    }

    var MIN_BASE_SCORE = 0.75;
    var MIN_FULL_SCORE = 0.65;
    var SIZE_MISMATCH_MIN_BASE = 0.85;

    var sizeMismatchAccept = best.sizeMismatch && best.baseScore >= SIZE_MISMATCH_MIN_BASE;
    var acceptMatch = best.entry &&
      best.baseScore >= MIN_BASE_SCORE &&
      (best.fullScore >= MIN_FULL_SCORE || sizeMismatchAccept);

    if (acceptMatch) {
      Logger.log('⚠️ DB 유사 매칭: ' + best.entry.description +
                 ' (base=' + best.baseScore.toFixed(2) +
                 ', full=' + best.fullScore.toFixed(2) +
                 (best.sizeMismatch ? ', sizeMismatch' : '') + ')');
      return {
        description: best.entry.description,
        matchType: 'fuzzy',
        score: best.baseScore,
        sizeMismatch: best.sizeMismatch
      };
    }

    if (trailingQuoteInfo.has && bestRejected.entry) {
      Logger.log('⚠️ 후행 따옴표 감지로 유사 매칭 제외: ' + bestRejected.entry.description +
                 ' (base=' + bestRejected.baseScore.toFixed(2) + ')');
      return {
        description: null,
        matchType: 'none',
        altDescription: bestRejected.entry.description,
        altScore: bestRejected.baseScore,
        altReason: 'TRAILING_QUOTE_SIZE'
      };
    }

    Logger.log('⚠️ DB 캐시 매칭 실패: ' + normalizedInput.substring(0, 60));
    return null;

  } catch (error) {
    Logger.log('❌ DB 검증 오류: ' + error.toString());
    return null;
  }
}

/**
 * OUTRE 텍스트 정규화
 * - 다중 공백 → 단일 공백
 * - 언더스코어 제거
 * - 앞뒤 공백 제거
 *
 * @param {string} text - 정규화할 텍스트
 * @return {string} 정규화된 텍스트
 */
function normalizeOutreText(text) {
  if (!text) return '';

  // 1. 다중 공백 → 단일 공백
  text = text.replace(/\s+/g, ' ');

  // 2. 언더스코어 제거
  text = text.replace(/_/g, '');

  // 3. 앞뒤 공백 제거
  return text.trim();
}

/**
 * 괄호 접두사 컬러 라인 전용 분리 함수
 * CRITICAL: STAGE 1/3에서만 사용 - 일반 케이스는 splitDescriptionAndColor() 사용
 *
 * 사용 케이스:
 * - "(P)1B- 60  (P)M4/30- 55  (P)M27/613- 55  (P)M30/33- 55" → 컬러 4개 추출
 * - "X-PRESSION BRAID 52" 3X (P)M950/425/350/130S- 55" → Description + 컬러 1개 분리
 *
 * 일반 컬러 라인("T30- 10  1B- 20")은 splitDescriptionAndColor() 사용
 *
 * 특징:
 * - 괄호 컬러는 "명확한 신호"이므로 1개만 있어도 분리 허용
 * - 인치 패턴 무시 (괄호 컬러는 Description과 확실히 구분됨)
 * - validColors.length >= 1 체크 (일반 케이스의 >= 2와 다름)
 *
 * @param {string} line - 괄호 컬러 포함 라인
 * @return {Object} {description: string, color: string|null}
 */
function extractColorsFromParenthesizedLine(line) {
  var normalized = normalizeOutreText(line);

  Logger.log('  🔧 괄호 컬러 전용 분리 시작: ' + normalized.substring(0, 80));

  // 컬러 패턴: 괄호 접두사 필수 + 컬러명 - 수량 (backorder 포함)
  // 예: (P)M950/425-55, (S)1B-20, (P)T30-10
  var parenColorPattern = /\([A-Z]\)[A-Z0-9\/+]+\s*-\s*\d+(?:\s*\(\d+\))?/gi;

  // 모든 괄호 컬러 패턴 매칭
  var matches = normalized.match(parenColorPattern);

  if (!matches || matches.length === 0) {
    Logger.log('    ⚠️ 괄호 컬러 패턴 없음');
    return { description: normalized, color: null };
  }

  Logger.log('    📊 괄호 컬러 패턴 ' + matches.length + '개 발견');

  // 각 매치를 validateOUTREColorToken()으로 검증
  var validColors = [];
  for (var i = 0; i < matches.length; i++) {
    var match = matches[i].match(/\([A-Z]\)([A-Z0-9\/+]+)\s*-\s*\d+/i);
    if (match) {
      var colorToken = match[1]; // 괄호 제외한 컬러명
      if (validateOUTREColorToken(colorToken)) {
        validColors.push(matches[i]);
        Logger.log('      ✓ 유효한 괄호 컬러: ' + matches[i]);
      } else {
        Logger.log('      ✗ 유효하지 않은 컬러 토큰: ' + colorToken);
      }
    }
  }

  // CRITICAL: 괄호 컬러는 1개만 있어도 분리 허용
  // (일반 케이스의 validColors.length < 2와 다름)
  if (validColors.length === 0) {
    Logger.log('    ⚠️ 유효한 괄호 컬러 없음');
    return { description: normalized, color: null };
  }

  // 첫 번째 유효한 컬러 위치에서 분리
  var firstColorIdx = normalized.indexOf(validColors[0]);

  Logger.log('    ✅ 괄호 컬러 분리 완료: ' + validColors.length + '개 유효한 컬러');

  return {
    description: normalized.slice(0, firstColorIdx).trim(),
    color: normalized.slice(firstColorIdx).trim()
  };
}

/**
 * Description과 컬러 라인이 섞인 라인을 분리 (일반 케이스 전용)
 * CRITICAL: 일반 컬러 라인에만 사용 - 괄호 컬러는 extractColorsFromParenthesizedLine() 사용
 *
 * 예: "SOME PRODUCT NAME T30- 10  1B- 20  613- 15"
 *     → description: "SOME PRODUCT NAME"
 *        color: "T30- 10  1B- 20  613- 15"
 *
 * 예외:
 * 1. 인치 패턴 포함 시 분리하지 않음 (allowInch=false일 때만)
 * 2. 컬러 패턴 1개 이하 시 분리하지 않음 ("CLIP-IN- 9PCS" 오인식 방지)
 * 3. 유효한 컬러 토큰 2개 미만 시 분리하지 않음
 *
 * @param {string} line - 분리할 라인
 * @param {boolean} allowInch - 인치 패턴 허용 여부 (기본값: false)
 * @return {Object} {description: string, color: string|null}
 */
function splitDescriptionAndColor(line, allowInch) {
  allowInch = allowInch || false;
  var normalized = normalizeOutreText(line);

  // 컬러 패턴: 괄호 접두사(선택) + 컬러명 - 수량 (backorder 포함)
  // 괄호 컬러: (P)M950/425-55
  // 일반 컬러: T30-10, 1B-20, 613/30-15
  var colorPattern = /(\([A-Z]\))?[A-Z0-9\/+]+\s*-\s*\d+(?:\s*\(\d+\))?/gi;

  var hasInchPattern = normalized.match(/\d+["″'']/);

  // 모든 컬러 패턴 매칭
  var matches = normalized.match(colorPattern);

  // 예외 1: 컬러 패턴이 1개 이하면 분리하지 않음
  if (!matches || matches.length < 2) {
    Logger.log('  ⚠️ 컬러 패턴 ' + (matches ? matches.length : 0) + '개, 분리하지 않음');
    return { description: normalized, color: null };
  }

  // 예외 3: 각 매치를 validateOUTREColorToken()으로 검증
  var validColors = [];
  for (var i = 0; i < matches.length; i++) {
    var match = matches[i].match(/(\([A-Z]\))?([A-Z0-9\/+]+)\s*-\s*\d+/i);
    if (match) {
      var colorToken = match[2];
      if (validateOUTREColorToken(colorToken)) {
        validColors.push(matches[i]);
      } else {
        Logger.log('  ⚠️ 유효하지 않은 컬러 토큰: ' + colorToken);
      }
    }
  }

  // 유효한 컬러가 2개 미만이면 분리하지 않음
  if (validColors.length < 2) {
    Logger.log('  ⚠️ 유효한 컬러 ' + validColors.length + '개, 분리하지 않음');
    return { description: normalized, color: null };
  }

  if (!allowInch && hasInchPattern) {
    Logger.log('  ⚠️ 인치 패턴 포함, 컬러 2개 이상만 분리 진행');
  }

  // 첫 번째 유효한 컬러 위치에서 분리
  var firstColorIdx = normalized.indexOf(validColors[0]);

  Logger.log('  ✅ 컬러 분리: ' + validColors.length + '개 유효한 컬러 발견');

  return {
    description: normalized.slice(0, firstColorIdx).trim(),
    color: normalized.slice(firstColorIdx).trim()
  };
}

/**
 * Memo에 경고/메모 텍스트 추가
 *
 * @param {string} memo - 기존 메모
 * @param {string} note - 추가할 메모
 * @return {string} 합쳐진 메모
 */
function appendOUTREMemo(memo, note) {
  if (!note) return memo || '';
  if (!memo) return note;
  return memo + ' / ' + note;
}

/**
 * Description 끝부분에 붙은 컬러 패턴 제거 (보수적)
 * 예: "X-PRESSION BRAID 52" 3X (P)M950-55"
 *     → "X-PRESSION BRAID 52" 3X"
 *
 * @param {string} description - 정리할 Description
 * @param {boolean} preserveNumberPattern - 숫자-숫자-숫자 패턴 유지 여부
 * @return {string} 정리된 Description
 */
function cleanDescriptionEnd(description, preserveNumberPattern) {
  if (!description) return '';

  var cleaned = description;

  // 1. 끝부분 컬러 패턴 제거 (괄호 접두사 포함)
  //    예: " (P)M950/425-55" 제거
  cleaned = cleaned.replace(/\s+(\([A-Z]\))?[A-Z0-9\/+]+\s*-\s*\d+(?:\s*\(\d+\))?$/i, '');

  // 2. 끝부분 숫자-숫자-숫자 패턴 제거 (라인 번호 오인식)
  //    예: " 201-549" 제거
  if (!preserveNumberPattern) {
    cleaned = cleaned.replace(/\s+\d+(?:-\d+){1,2}$/, '');
  }

  return cleaned.trim();
}

/**
 * OUTRE 컬러 토큰 검증 (개선 버전)
 * CRITICAL: 숫자 컬러(1, 30, 613, 530 등) 지원 필수!
 *
 * 허용:
 * - T30, 1B, 613/30 (알파벳 포함)
 * - 1, 30, 613 (순수 숫자 1-3자리)
 *
 * 차단:
 * - 201-549 (숫자-숫자 패턴, 라인 번호 오인식)
 * - 346/843 (큰 숫자 조합)
 *
 * @param {string} colorToken - 검증할 컬러 토큰
 * @return {boolean} 유효한 컬러 토큰인지 여부
 */
function validateOUTREColorToken(colorToken) {
  if (!colorToken || colorToken.length === 0) return false;

  // 너무 긴 토큰 제외 (20자 제한)
  if (colorToken.length > 20) return false;

  // 메타데이터 키워드 제외
  var metadataKeywords = [
    'SHIP', 'SOLD', 'WEIGHT', 'SUBTOTAL', 'RICHMOND', 'LLC',
    'PKWAY', 'COD', 'FEE', 'TAG', 'DATE', 'PAGE', 'VIA',
    'PAYMENT', 'TERMS', 'TOTAL', 'PRICE', 'UNIT', 'DISC',
    'EXT', 'HAIR', 'WIG', 'LACE', 'WEAVE', 'CLOSURE'
  ];

  var upperToken = colorToken.toUpperCase();
  for (var i = 0; i < metadataKeywords.length; i++) {
    if (upperToken.indexOf(metadataKeywords[i]) > -1) {
      return false;
    }
  }

  // CRITICAL: Description 키워드 블랙리스트 (컬러로 오인식 방지)
  // "CLIP-IN- 9PCS", "PERUVIAN WAVE 18" 등 Description 내부 패턴 차단
  var descriptionBlacklist = [
    'CLIP', 'PCS', 'WAVE', 'PERUVIAN', 'BRAZILIAN',
    'STRAIGHT', 'CURLY', 'BUNDLE', 'FRONTAL',
    'PONYTAIL', 'CROCHET', 'PACK', 'INCH', 'IN'
  ];

  for (var i = 0; i < descriptionBlacklist.length; i++) {
    if (upperToken.indexOf(descriptionBlacklist[i]) > -1) {
      Logger.log('  ⚠️ Description 키워드 포함: ' + colorToken);
      return false;
    }
  }

  // 유효한 컬러 패턴: 숫자, 알파벳, 하이픈, 슬래시 조합
  // 예: 1, 30, 613, 530, GINGER, T30, 1B/30, M950, BLD-CRUSH
  if (!colorToken.match(/^[A-Z0-9\-\/+]+$/i)) {
    return false;
  }

  // CRITICAL: 숫자 컬러 검증 (라인 번호 오인식 방지)
  // 알파벳 포함 → OK
  if (/[A-Z]/i.test(colorToken)) {
    return true;
  }

  // 순수 숫자 1-3자리 → OK (1, 30, 613)
  if (/^\d{1,3}$/.test(colorToken)) {
    return true;
  }

  if (/^\d{1,3}(\+\d{1,3})+$/.test(colorToken)) {
    return true;
  }


  // 그 외 (201-549, 346/843 등) → 차단
  Logger.log('  ⚠️ 유효하지 않은 숫자 조합 컬러 토큰: ' + colorToken);
  return false;
}

/**
 * 괄호 접두사 제거 및 정규화
 * CRITICAL: (P), (S) 제거하고 컬러명만 반환
 *
 * @param {string} colorToken - 컬러 토큰 (예: (P)M950, 1B, 30)
 * @return {string} 정규화된 컬러명
 */
function normalizeOUTREColorToken(colorToken) {
  if (!colorToken) return '';

  // 괄호 접두사 제거: (P)M950 → M950, (S)30 → 30
  var normalized = colorToken.replace(/^\([A-Z]\)/i, '');

  return normalized.trim().toUpperCase();
}

/**
 * OUTRE 컬러 라인 파싱 (개선 버전 v2)
 * CRITICAL: 숫자 컬러 지원 + 괄호 접두사 처리 개선 + Description 분리
 *
 * @param {Array} colorLines - 컬러 라인 배열
 * @param {string} description - Description 텍스트 (제외용)
 * @return {Array} 컬러 데이터 배열 [{color, shipped, backordered}, ...]
 */
function parseOUTREColorLines(colorLines, description) {
  var colorData = [];

  var fullText = colorLines.join(' ');

  Logger.log('=== OUTRE 컬러 라인 파싱 시작 (개선 버전 v2) ===');
  Logger.log('원본 라인 수: ' + colorLines.length);
  Logger.log('원본 텍스트: ' + fullText.substring(0, 150));

  // Step 1: Normalize (normalizeOutreText 사용)
  fullText = normalizeOutreText(fullText);
  Logger.log('Step 1 (Normalize): ' + fullText.substring(0, 150));

  // Step 2: Description 제거 (단어 기반)
  if (description) {
    var descWords = description.split(/\s+/);
    for (var i = 0; i < descWords.length; i++) {
      var word = descWords[i].trim();
      if (word.length > 2) {
        var regex = new RegExp('\\b' + word + '\\b', 'gi');
        fullText = fullText.replace(regex, ' ');
      }
    }
    fullText = normalizeOutreText(fullText);
    Logger.log('Step 2 (Description 제거): ' + fullText.substring(0, 150));
  }

  // Step 3: 가격 패턴 제거 (3개 연속 숫자.숫자)
  fullText = fullText.replace(/\d+\.\d{2}\s+\d+\.\d{2}\s+\d+\.\d{2}/g, ' ');
  fullText = normalizeOutreText(fullText);
  Logger.log('Step 3 (가격 제거): ' + fullText.substring(0, 150));

  // Step 3.5: 전화번호/헤더 키워드 제거 (컬러 오인식 방지)
  fullText = fullText.replace(/\bTOLL\s+FREE\b|\bPHONE\b|\bFAX\b/gi, ' ');
  fullText = fullText.replace(/\b\d{3}[-.\s]?\d{3}[-.\s]?\d{4}\b/g, ' ');
  fullText = normalizeOutreText(fullText);
  Logger.log('Step 3.5 (전화번호 제거): ' + fullText.substring(0, 150));

  // Step 4: 컬러 패턴 처리 (괄호 접두사 + 일반 패턴 통합)
  // CRITICAL: 괄호 컬러와 일반 컬러를 한 번에 처리하여 중복 방지
  // 패턴: [괄호접두사(선택)]컬러명 - shipped (backordered)
  // 예: (P)M950/425-55, T30-10, 1B-20, 613/30-15
  var colorPattern = /(\([A-Z]\))?([A-Z0-9\/+]+)\s*-\s*(\d+)(?:\s*\((\d+)\))?/gi;
  var match;

  Logger.log('Step 4 (컬러 패턴 추출):');

  while ((match = colorPattern.exec(fullText)) !== null) {
    var prefix = match[1] || '';        // (P), (S) 등
    var colorToken = match[2].trim();   // M950/425/350/130S, T30, 1B 등
    var shipped = parseInt(match[3]);
    var backordered = match[4] ? parseInt(match[4]) : 0;

    Logger.log('  매치: ' + (prefix ? prefix : '') + colorToken + ' - ' + shipped + (backordered > 0 ? ' (' + backordered + ')' : ''));

    // 토큰 검증
    if (!validateOUTREColorToken(colorToken)) {
      Logger.log('    ⚠️ 유효하지 않은 컬러 토큰 무시: ' + colorToken);
      continue;
    }

    // 괄호 접두사 제거 (normalizeOUTREColorToken 사용하지 않음 - 이미 대문자 변환됨)
    var finalColor = prefix ? colorToken : colorToken.toUpperCase();

    Logger.log('    ✅ 컬러 추가: ' + finalColor + ' (Shipped: ' + shipped + ', Backordered: ' + backordered + ')');

    colorData.push({
      color: finalColor,
      shipped: shipped,
      backordered: backordered
    });
  }

  Logger.log('컬러 파싱 완료: ' + colorData.length + '개 컬러');

  // CRITICAL: 2-pass orphan backorder 연결 (이중 안전장치)
  // parseOUTREItem()에서 수집 실패한 경우 대비
  // 패턴: shipped=0 & backordered=0인 컬러 다음에 단독 (\d+) 찾기
  Logger.log('Step 5 (2-pass orphan backorder 연결):');

  var orphanBackorders = [];
  var orphanPattern = /\((\d+)\)/g;
  var orphanMatch;

  while ((orphanMatch = orphanPattern.exec(fullText)) !== null) {
    var backorderQty = parseInt(orphanMatch[1]);
    // 이미 매칭된 컬러의 백오더는 제외 (중복 방지)
    var alreadyMatched = false;
    for (var i = 0; i < colorData.length; i++) {
      if (colorData[i].backordered === backorderQty) {
        alreadyMatched = true;
        break;
      }
    }

    if (!alreadyMatched && backorderQty > 0) {
      orphanBackorders.push(backorderQty);
      Logger.log('  고아 백오더 발견: (' + backorderQty + ')');
    }
  }

  if (orphanBackorders.length > 0) {
    Logger.log('  총 ' + orphanBackorders.length + '개 고아 백오더 발견');

    // shipped=0 & backordered=0인 컬러에 순서대로 할당
    var orphanIndex = 0;
    for (var i = 0; i < colorData.length && orphanIndex < orphanBackorders.length; i++) {
      if (colorData[i].shipped === 0 && colorData[i].backordered === 0) {
        colorData[i].backordered = orphanBackorders[orphanIndex];
        Logger.log('  ✅ ' + colorData[i].color + ' 백오더 연결: 0 → ' + orphanBackorders[orphanIndex]);
        orphanIndex++;
      }
    }

    if (orphanIndex < orphanBackorders.length) {
      Logger.log('  ⚠️ 할당되지 않은 고아 백오더 ' + (orphanBackorders.length - orphanIndex) + '개 남음');
    }
  } else {
    Logger.log('  고아 백오더 없음');
  }

  return colorData;
}

/**
 * Log OUTRE UPC scan diagnostics to a spreadsheet tab.
 *
 * @param {Object} data
 * @param {string} data.source - Caller name or context
 * @param {string} data.inputUpc - Raw scanned UPC
 * @param {string} data.normalizedUpc - Normalized UPC used for lookup
 * @param {string} data.matchedUpc - UPC returned from DB
 * @param {string} data.matchedColor - Color matched in DB
 * @param {string} data.matchedItemNumber - Item number matched in DB
 * @param {string} data.description - Description context
 * @param {string} data.note - Extra note or warning
 */
function logOUTREUPCScanToSheet(data) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheetName = 'OUTRE_UPC_SCAN_LOG';
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp',
        'Source',
        'Input UPC',
        'Normalized UPC',
        'Matched UPC',
        'Matched Color',
        'Matched Item Number',
        'Description',
        'Note'
      ]);
    }

    sheet.appendRow([
      new Date(),
      (data && data.source) || '',
      (data && data.inputUpc) || '',
      (data && data.normalizedUpc) || '',
      (data && data.matchedUpc) || '',
      (data && data.matchedColor) || '',
      (data && data.matchedItemNumber) || '',
      (data && data.description) || '',
      (data && data.note) || ''
    ]);
  } catch (error) {
    Logger.log('❌ OUTRE UPC scan log error: ' + error.toString());
  }
}

/**
 * Debug helper: scan a UPC against OUTRE DB and log results.
 *
 * @param {string} inputUpc - Raw UPC to test
 */
function debugOUTREUPCScan(inputUpc) {
  if (!inputUpc) {
    Logger.log('⚠️ debugOUTREUPCScan called without inputUpc');
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScan',
      note: 'Missing inputUpc'
    });
    return;
  }
  var normalizedInput = normalizeOUTREUPCValue(inputUpc);
  Logger.log('=== OUTRE UPC DEBUG START ===');
  Logger.log('Input UPC: ' + inputUpc + ' | Normalized: ' + normalizedInput);

  initOUTREDBCache();
  if (!OUTRE_DB_CACHE || OUTRE_DB_CACHE.error) {
    Logger.log('⚠️ OUTRE DB cache unavailable');
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScan',
      inputUpc: inputUpc,
      normalizedUpc: normalizedInput,
      note: 'DB cache unavailable'
    });
    return;
  }

  var dbMap = OUTRE_DB_CACHE.dbMap;
  var matches = [];

  for (var key in dbMap) {
    if (!dbMap.hasOwnProperty(key)) continue;
    var records = dbMap[key];
    for (var i = 0; i < records.length; i++) {
      var recordUpc = normalizeOUTREUPCValue(records[i].barcode);
      if (recordUpc && recordUpc === normalizedInput) {
        matches.push({
          description: records[i].description || '',
          itemNumber: records[i].itemNumber || '',
          color: records[i].color || '',
          barcode: records[i].barcode || ''
        });
      }
    }
  }

  if (matches.length === 0) {
    Logger.log('⚠️ No DB match for UPC: ' + normalizedInput);
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScan',
      inputUpc: inputUpc,
      normalizedUpc: normalizedInput,
      note: 'No DB match'
    });
    return;
  }

  Logger.log('✅ DB matches: ' + matches.length);
  for (var m = 0; m < Math.min(matches.length, 5); m++) {
    Logger.log('  - ' + matches[m].barcode + ' | ' + matches[m].color + ' | ' + matches[m].itemNumber +
               ' | ' + matches[m].description.substring(0, 80));
  }

  logOUTREUPCScanToSheet({
    source: 'debugOUTREUPCScan',
    inputUpc: inputUpc,
    normalizedUpc: normalizedInput,
    matchedUpc: matches[0].barcode || '',
    matchedColor: matches[0].color || '',
    matchedItemNumber: matches[0].itemNumber || '',
    description: matches[0].description || '',
    note: matches.length > 1 ? ('Multiple matches: ' + matches.length) : 'Single match'
  });
}

/**
 * Debug helper with a hardcoded sample UPC.
 */
function debugOUTREUPCScanSample() {
  debugOUTREUPCScan('827298092940');
}

/**
 * Normalize UPC to digits only.
 *
 * @param {string} value
 * @return {string}
 */
function normalizeOUTREUPCValue(value) {
  if (!value) return '';
  return value.toString().replace(/[^0-9]/g, '');
}

/**
 * Debug helper: scan a UPC directly from the OUTRE DB sheet.
 *
 * @param {string} inputUpc - Raw UPC to test
 */
function debugOUTREUPCScanBySheet(inputUpc) {
  if (!inputUpc) {
    Logger.log('⚠️ debugOUTREUPCScanBySheet called without inputUpc');
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScanBySheet',
      note: 'Missing inputUpc'
    });
    return;
  }

  var normalizedInput = normalizeOUTREUPCValue(inputUpc);
  Logger.log('=== OUTRE UPC SHEET DEBUG START ===');
  Logger.log('Input UPC: ' + inputUpc + ' | Normalized: ' + normalizedInput);

  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var sheet = ss.getSheetByName(CONFIG.COMPANIES.OUTRE.dbSheet);

  if (!sheet) {
    Logger.log('⚠️ OUTRE DB sheet not found');
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScanBySheet',
      inputUpc: inputUpc,
      normalizedUpc: normalizedInput,
      note: 'DB sheet not found'
    });
    return;
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log('⚠️ OUTRE DB sheet empty');
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScanBySheet',
      inputUpc: inputUpc,
      normalizedUpc: normalizedInput,
      note: 'DB sheet empty'
    });
    return;
  }

  var headers = data[0];
  var colMap = {};
  for (var i = 0; i < headers.length; i++) {
    colMap[headers[i]] = i;
  }

  var itemNameCol = colMap[CONFIG.COMPANIES.OUTRE.columns.ITEM_NAME];
  var itemNumberCol = colMap[CONFIG.COMPANIES.OUTRE.columns.ITEM_NUMBER];
  var colorCol = colMap[CONFIG.COMPANIES.OUTRE.columns.COLOR];
  var barcodeCol = colMap[CONFIG.COMPANIES.OUTRE.columns.BARCODE];

  if (barcodeCol === undefined) {
    Logger.log('⚠️ BARCODE column not found');
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScanBySheet',
      inputUpc: inputUpc,
      normalizedUpc: normalizedInput,
      note: 'BARCODE column not found'
    });
    return;
  }

  var matches = [];
  for (var r = 1; r < data.length; r++) {
    var rawBarcode = data[r][barcodeCol];
    var normalizedBarcode = normalizeOUTREUPCValue(rawBarcode);
    if (normalizedBarcode && normalizedBarcode === normalizedInput) {
      matches.push({
        description: itemNameCol !== undefined ? (data[r][itemNameCol] || '') : '',
        itemNumber: itemNumberCol !== undefined ? (data[r][itemNumberCol] || '') : '',
        color: colorCol !== undefined ? (data[r][colorCol] || '') : '',
        barcode: rawBarcode || ''
      });
    }
  }

  if (matches.length === 0) {
    Logger.log('⚠️ No sheet match for UPC: ' + normalizedInput);
    logOUTREUPCScanToSheet({
      source: 'debugOUTREUPCScanBySheet',
      inputUpc: inputUpc,
      normalizedUpc: normalizedInput,
      note: 'No sheet match'
    });
    return;
  }

  Logger.log('✅ Sheet matches: ' + matches.length);
  for (var m = 0; m < Math.min(matches.length, 5); m++) {
    Logger.log('  - ' + matches[m].barcode + ' | ' + matches[m].color + ' | ' + matches[m].itemNumber +
               ' | ' + matches[m].description.substring(0, 80));
  }

  logOUTREUPCScanToSheet({
    source: 'debugOUTREUPCScanBySheet',
    inputUpc: inputUpc,
    normalizedUpc: normalizedInput,
    matchedUpc: matches[0].barcode || '',
    matchedColor: matches[0].color || '',
    matchedItemNumber: matches[0].itemNumber || '',
    description: matches[0].description || '',
    note: matches.length > 1 ? ('Multiple matches: ' + matches.length) : 'Single match'
  });
}
