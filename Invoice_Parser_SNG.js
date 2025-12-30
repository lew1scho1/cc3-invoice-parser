// ============================================================================
// Invoice_Parser_SNG.js - SNG (Shake-N-Go) Invoice Parser
// ============================================================================
//
// Parsing model (DOCX/PDF text):
// 1) Item line: [PACKED BY?] QTY_ORDERED QTY_SHIPPED ITEM_ID DESCRIPTION LIST_PRICE LIST_EXT
// 2) Price line: YOUR_PRICE YOUR_EXT DISCOUNTED (may include color data after prices)
// 3) Color lines: one or more lines with COLOR - QTY (BACKORDER)
//
// The parser only trusts explicit quantities from color lines.
// ============================================================================

var SNG_ITEM_LINE_REGEX = /^(?:[A-Z]\d{1,3}\s+)?(\d{1,3})\s+(\d{1,3})\s*([A-Z][A-Z0-9]+)\s+(.+?)\s*(\d+\.\d{2})\s*(\d+\.\d{2})\s*$/;
var SNG_ITEM_LINE_ALT_REGEX = /^(?:[A-Z]\d{1,3}\s+)?(\d{1,3})\s*([A-Z][A-Z0-9]+)\s+(.+?)\s*(\d+\.\d{2})\s*(\d+\.\d{2})\s*$/;
var SNG_PRICE_LINE_REGEX = /^(\d+\.\d{2})\s*(\d+\.\d{2})\s*(\d+\.\d{2})(?:\s+(.*))?$/;
var SNG_COLOR_REGEX = /([A-Z0-9][A-Z0-9\-\/]{0,20})\s*-\s*(\d+)\s*(?:\((\d+)\))?/g;
var SNG_HEADER_REGEX = /INVOICE DATE|DUE DATE|SHIP VIA|ORDER DATE|SALESPERSON|TERMS|C\.O\.D|PACKED BY|QTY ORDERED|QTY SHIPPED|ITEM NUMBER|DESCRIPTION|LIST PRICE|YOUR PRICE|LIST EXTENDED|YOUR EXTENDED|DISCOUNTED/i;
var SNG_ITEM_BOUNDARY_REGEX = /^(?:[A-Z]\d{1,3}\s+)?\d{1,3}(?:\s+\d{1,3})?\s*[A-Z][A-Z0-9]+\s+.+?\s*\d+\.\d{2}\s*\d+\.\d{2}\s*$/;
var SNG_COLOR_SCAN_LIMIT = 160;

/**
 * SNG 인보이스 라인 아이템 파싱
 * @param {Array<string>} lines - 인보이스 텍스트 라인 배열
 * @param {Sheet} dbSheet - DB_SNG 시트 (옵션, DB 검증 활성화)
 * @return {Array<Object>} 파싱된 라인 아이템 배열
 */
function parseSNGLineItems(lines, dbSheet) {
  var items = [];
  var lineNo = 1;

  debugLog('SNG 라인 아이템 파싱 시작', { totalLines: lines.length, dbValidation: !!dbSheet });

  // CRITICAL: DB 검증 활성화 시 배치 최적화 - DB를 한 번만 읽음
  var dbMap = null;
  if (dbSheet) {
    debugLog('SNG DB Map 생성 시작 (배치 최적화)');
    dbMap = buildSNGItemCodeMap(dbSheet);
    debugLog('SNG DB Map 생성 완료', { itemCount: Object.keys(dbMap.itemCodeMap).length });
  }

  for (var i = 0; i < lines.length; ) {
    var itemInfo = parseSNGItemLine(lines[i]);
    if (!itemInfo) {
      i++;
      continue;
    }

    var unitPrice = itemInfo.listPrice;
    var extPrice = itemInfo.listExtended;
    var colorLines = [];
    var scanStartIndex = i + 1;
    var nextLine = lines[i + 1] || '';

    var priceInfo = parseSNGPriceLine(lines[i + 1]);
    if (priceInfo) {
      unitPrice = priceInfo.yourPrice;
      extPrice = priceInfo.yourExtended;
      scanStartIndex = i + 2;

      // CRITICAL: 가격 라인에 컬러 텍스트가 있으면 추가
      if (priceInfo.colorText && hasSNGColorPattern(priceInfo.colorText)) {
        colorLines.push(priceInfo.colorText);
      } else if (!priceInfo.colorText) {
        // 가격 라인에 컬러가 없으면 다음 라인 확인
        // CRITICAL: 가격 라인 바로 아래 컬러가 붙는 경우 대응
        if (hasSNGColorPattern(lines[i + 2])) {
          colorLines.push(lines[i + 2]);
          scanStartIndex = i + 3;
        }
      }
    } else if (hasSNGColorPattern(nextLine)) {
      // 가격 라인이 없고 바로 컬러 라인인 경우
      colorLines.push(nextLine);
      scanStartIndex = i + 2;
    }

    var colorScan = collectSNGColorLines(scanStartIndex, lines);
    colorLines = colorLines.concat(colorScan.lines);

    var colorData = parseSNGColorLines(colorLines);

    if (colorData.length > 0) {
      var totalShipped = 0;
      for (var c = 0; c < colorData.length; c++) {
        totalShipped += colorData[c].shipped;
      }

      for (var m = 0; m < colorData.length; m++) {
        var cd = colorData[m];
        var itemExtPrice = 0;

        if (totalShipped > 0) {
          itemExtPrice = Number((extPrice * (cd.shipped / totalShipped)).toFixed(2));
        }

        var item = {
          lineNo: lineNo++,
          itemId: itemInfo.itemId,
          upc: '',
          description: itemInfo.description,
          brand: CONFIG.INVOICE.BRANDS['SNG'],
          color: cd.color,
          sizeLength: '',
          qtyOrdered: cd.shipped + cd.backordered,
          qtyShipped: cd.shipped,
          unitPrice: unitPrice,
          extPrice: itemExtPrice,
          memo: cd.backordered > 0 ? 'Backordered: ' + cd.backordered : ''
        };

        // DB 검증 (dbMap이 제공된 경우만)
        if (dbMap) {
          item = validateSNGItemWithDB(item, null, dbMap);
        }

        items.push(item);
      }
    } else {
      var item = {
        lineNo: lineNo++,
        itemId: itemInfo.itemId,
        upc: '',
        description: itemInfo.description,
        brand: CONFIG.INVOICE.BRANDS['SNG'],
        color: '',
        sizeLength: '',
        qtyOrdered: itemInfo.qtyOrdered,
        qtyShipped: itemInfo.qtyShipped,
        unitPrice: unitPrice,
        extPrice: extPrice,
        memo: '⚠️ 컬러 정보 없음'
      };

      // DB 검증 (dbMap이 제공된 경우만)
      if (dbMap) {
        item = validateSNGItemWithDB(item, null, dbMap);
      }

      items.push(item);
    }

    i = colorScan.nextIndex;
  }

  debugLog('SNG 라인 아이템 파싱 완료', { totalItems: items.length });

  return items;
}

/**
 * SNG prefix 토큰 파싱 (PACKED BY 처리 포함)
 * CRITICAL: 파싱/경계감지 모든 로직의 단일 진실 공급원 (Single Source of Truth)
 *
 * @param {string} prefix - 가격 앞부분 텍스트
 * @return {Object|null} { qtyOrdered, qtyShipped, itemId, description }
 */
function parseSNGPrefix(prefix) {
  if (!prefix) return null;

  // 토큰 분리
  var tokens = prefix.split(/\s+/);
  if (tokens.length < 3) return null;

  // PACKED BY 판정 및 제거
  var first = tokens[0];
  var isPackedByNumeric = /^[A-Z]\d{1,3}$/.test(first);  // T75 형태
  var isPackedByAlpha = /^[A-Z]{2,3}$/.test(first);      // BKJ 형태

  if (tokens.length >= 4 && (isPackedByNumeric || isPackedByAlpha)) {
    var secondIsNum = !isNaN(parseInt(tokens[1]));
    var thirdIsNum = !isNaN(parseInt(tokens[2]));

    // CRITICAL: 뒤에 숫자 2개가 있을 때만 PACKED BY로 인정 (오탐 방지)
    if (secondIsNum && thirdIsNum) {
      tokens.shift();  // PACKED BY 제거
    }
  }

  // QTY, ITEM ID, DESCRIPTION 추출
  if (tokens.length < 3) return null;

  var qtyOrdered, qtyShipped, itemIdIndex;

  // 패턴 1: QTY_ORDERED QTY_SHIPPED ITEM_ID DESCRIPTION
  if (tokens.length >= 4 && !isNaN(parseInt(tokens[0])) && !isNaN(parseInt(tokens[1]))) {
    qtyOrdered = parseInt(tokens[0], 10);
    qtyShipped = parseInt(tokens[1], 10);
    itemIdIndex = 2;
  }
  // 패턴 2: QTY ITEM_ID DESCRIPTION (QTY_ORDERED = QTY_SHIPPED)
  else if (!isNaN(parseInt(tokens[0]))) {
    qtyOrdered = qtyShipped = parseInt(tokens[0], 10);
    itemIdIndex = 1;
  } else {
    return null;
  }

  var itemId = tokens[itemIdIndex];
  var description = tokens.slice(itemIdIndex + 1).join(' ');

  // 검증: ITEM_ID는 최소 4글자 이상, 대문자+숫자 조합
  if (!itemId || !itemId.match(/^[A-Z][A-Z0-9]{3,}$/)) {
    return null;
  }

  // 검증: Description이 없거나 너무 짧으면 제외
  if (!description || description.length < 5) {
    return null;
  }

  return {
    qtyOrdered: qtyOrdered,
    qtyShipped: qtyShipped,
    itemId: itemId,
    description: description
  };
}

/**
 * SNG 아이템 라인 파싱 (가격 토큰 기반 + 토큰 파싱)
 * CRITICAL: parseSNGPrefix() 공통 로직 사용
 *
 * @param {string} rawLine - 라인 원문
 * @return {Object|null} 파싱된 아이템 정보
 */
function parseSNGItemLine(rawLine) {
  if (!rawLine) return null;

  var normalized = rawLine.replace(/\s+/g, ' ').trim();
  if (!normalized) return null;

  // 1단계: 모든 가격 토큰 찾기 (숫자.두자리)
  var priceMatches = [];
  var priceRegex = /\d+\.\d{2}/g;
  var match;
  while ((match = priceRegex.exec(normalized)) !== null) {
    priceMatches.push({ text: match[0], index: match.index });
  }

  // 가격 토큰이 최소 2개 필요 (list price, list extended)
  if (priceMatches.length < 2) {
    return null;
  }

  // 2단계: 마지막 2개를 list price/extended로 해석
  var listPriceMatch = priceMatches[priceMatches.length - 2];
  var listExtendedMatch = priceMatches[priceMatches.length - 1];

  // 가격 앞부분이 아이템 정보
  var prefix = normalized.slice(0, listPriceMatch.index).trim();

  // 아이템 번호 패턴이 없으면 제외 (빠른 필터링)
  if (!/[A-Z][A-Z0-9]+/.test(prefix)) {
    return null;
  }

  // 3단계: 공통 토큰 파싱 로직 사용 (PACKED BY 처리 포함)
  var parsed = parseSNGPrefix(prefix);
  if (!parsed) return null;

  return {
    qtyOrdered: parsed.qtyOrdered,
    qtyShipped: parsed.qtyShipped,
    itemId: parsed.itemId,
    description: parsed.description,
    listPrice: parseAmount(listPriceMatch.text),
    listExtended: parseAmount(listExtendedMatch.text)
  };
}

/**
 * SNG 가격 라인 파싱 (가격 토큰 기반)
 * CRITICAL: 가격 토큰 3개를 찾아서 yourPrice/yourExtended/discounted로 처리
 *
 * @param {string} rawLine - 라인 원문
 * @return {Object|null} 가격 정보
 */
function parseSNGPriceLine(rawLine) {
  if (!rawLine) return null;

  var normalized = rawLine.replace(/\s+/g, ' ').trim();
  if (!normalized) return null;

  // 1단계: 모든 가격 토큰 찾기
  var priceMatches = [];
  var priceRegex = /\d+\.\d{2}/g;
  var match;
  while ((match = priceRegex.exec(normalized)) !== null) {
    priceMatches.push({ text: match[0], index: match.index });
  }

  // 가격 토큰 3개 필요 (yourPrice, yourExtended, discounted)
  if (priceMatches.length < 3) {
    return null;
  }

  // 2단계: 처음 3개를 가격으로 해석
  var thirdMatch = priceMatches[2];

  // 3단계: 세 번째 가격 뒤의 텍스트를 컬러 텍스트로 사용
  var colorText = normalized.slice(thirdMatch.index + thirdMatch.text.length).trim();

  return {
    yourPrice: parseAmount(priceMatches[0].text),
    yourExtended: parseAmount(priceMatches[1].text),
    discounted: parseAmount(thirdMatch.text),
    colorText: colorText
  };
}

/**
 * 아이템 이후 컬러 라인 수집 (강화된 종료 조건)
 * CRITICAL: 다음 아이템 감지, 가격 라인, 아이템 시작 라인을 만나면 즉시 중단
 *
 * @param {number} startIndex - 컬러 라인 시작 인덱스
 * @param {Array<string>} lines - 전체 라인 배열
 * @return {Object} { lines: Array<string>, nextIndex: number }
 */
function collectSNGColorLines(startIndex, lines) {
  var collected = [];
  var i;

  for (i = startIndex; i < lines.length; i++) {
    // 1단계: 스캔 범위 제한 (성능)
    if (i - startIndex >= SNG_COLOR_SCAN_LIMIT) {
      break;
    }

    var rawLine = lines[i];
    if (!rawLine || !rawLine.trim()) {
      continue;
    }

    // 2단계: 다음 아이템 라인 감지 → 즉시 중단
    if (parseSNGItemLine(rawLine)) {
      break;
    }

    // 3단계: 아이템 시작 라인 감지 → 즉시 중단
    if (isSNGItemStartLine(rawLine)) {
      break;
    }

    // 4단계: 아이템 경계 라인 감지 (가격 2개 이상) → 즉시 중단
    if (isSNGItemBoundaryLine(rawLine)) {
      break;
    }

    // 5단계: 컬러 텍스트 정규화
    var normalized = normalizeSNGColorText(rawLine);
    if (!normalized) {
      continue;
    }

    // 6단계: 컬러 패턴 확인
    var colorRegex = new RegExp(SNG_COLOR_REGEX.source);
    var hasColorPattern = colorRegex.test(normalized);

    // 7단계: 가격 라인 감지 (컬러 패턴 없으면) → 즉시 중단
    var isPriceLine = parseSNGPriceLine(rawLine);
    if (isPriceLine && !hasColorPattern) {
      break;
    }

    // 8단계: 헤더 라인 처리 (컬러 패턴 유무에 따라 다르게 처리)
    var hasHeader = SNG_HEADER_REGEX.test(normalized);

    if (hasHeader && hasColorPattern) {
      // 헤더+컬러 혼재 라인 → 컬러 라인으로 수집
      collected.push(rawLine);
      continue;
    } else if (hasHeader && !hasColorPattern) {
      // 헤더만 있는 라인 → 건너뛰기
      continue;
    }

    // 9단계: 컬러 패턴이 없으면 건너뛰기
    if (!hasColorPattern) {
      continue;
    }

    // 10단계: 컬러 라인 수집
    collected.push(rawLine);
  }

  return { lines: collected, nextIndex: i };
}

/**
 * SNG 아이템 시작 라인 감지 (가격 없이도 감지)
 * @param {string} rawLine
 * @return {boolean}
 */
function isSNGItemStartLine(rawLine) {
  if (!rawLine) return false;

  var normalized = rawLine.replace(/\s+/g, ' ').trim();
  if (!normalized) return false;
  if (SNG_HEADER_REGEX.test(normalized)) return false;

  // 가격 토큰이 2개 이상 있으면 아이템 라인
  var priceMatches = normalized.match(/\d+\.\d{2}/g);
  if (!priceMatches || priceMatches.length < 2) return false;

  // 가격 앞부분 추출
  var priceIndex = normalized.indexOf(priceMatches[priceMatches.length - 2]);
  var prefix = normalized.slice(0, priceIndex).trim();

  // 아이템 번호 패턴이 없으면 false
  if (!/[A-Z][A-Z0-9]+/.test(prefix)) return false;

  // CRITICAL: parseSNGItemLine()과 동일한 로직 사용 (경계 일치 보장)
  var parsed = parseSNGPrefix(prefix);
  return parsed !== null;
}

/**
 * SNG 아이템 경계 라인 감지
 * CRITICAL: parseSNGItemLine()과 동일한 로직 사용 (경계 일치 보장)
 * @param {string} rawLine
 * @return {boolean}
 */
function isSNGItemBoundaryLine(rawLine) {
  if (!rawLine) return false;

  var normalized = rawLine.replace(/\s+/g, ' ').trim();
  if (!normalized) return false;

  if (SNG_HEADER_REGEX.test(normalized)) return false;

  var priceMatches = normalized.match(/\d+\.\d{2}/g);
  if (!priceMatches || priceMatches.length < 2) return false;

  // 가격 앞부분 추출
  var priceIndex = normalized.indexOf(priceMatches[priceMatches.length - 2]);
  var prefix = normalized.slice(0, priceIndex).trim();

  // 아이템 번호 패턴이 없으면 false
  if (!/[A-Z][A-Z0-9]+/.test(prefix)) return false;

  // CRITICAL: parseSNGItemLine()과 동일한 로직 사용
  var parsed = parseSNGPrefix(prefix);
  return parsed !== null;
}

/**
 * 컬러 패턴 여부 확인
 * @param {string} text
 * @return {boolean}
 */
function hasSNGColorPattern(text) {
  if (!text) return false;
  var normalized = normalizeSNGColorText(text);
  if (!normalized) return false;
  var colorRegex = new RegExp(SNG_COLOR_REGEX.source);
  return colorRegex.test(normalized);
}

/**
 * SNG 컬러 라인 파싱
 * @param {Array<string>} colorLines - 컬러 라인 배열
 * @return {Array} [{color, shipped, backordered}, ...]
 */
function parseSNGColorLines(colorLines) {
  var colorData = [];
  if (!colorLines || colorLines.length === 0) {
    return colorData;
  }

  var fullText = normalizeSNGColorText(colorLines.join(' '));
  if (!fullText) {
    return colorData;
  }

  var regex = new RegExp(SNG_COLOR_REGEX);
  var match;

  while ((match = regex.exec(fullText)) !== null) {
    var color = match[1].trim();
    var shipped = parseInt(match[2], 10) || 0;
    var backordered = match[3] ? parseInt(match[3], 10) : 0;

    if (color && (shipped > 0 || backordered > 0)) {
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
 * 컬러 라인 정규화
 * @param {string} text - 컬러 라인 텍스트
 * @return {string} 정규화된 텍스트
 */
function normalizeSNGColorText(text) {
  if (!text) return '';

  var normalized = text.replace(/_+/g, ' ');
  normalized = normalized.replace(/\s+/g, ' ').trim();

  // 가격 3개가 라인 앞/뒤에 붙는 경우 제거
  normalized = normalized.replace(/^\d+\.\d{2}\s*\d+\.\d{2}\s*\d+\.\d{2}\s*/, '');
  normalized = normalized.replace(/\d+\.\d{2}\s*\d+\.\d{2}\s*\d+\.\d{2}\s*$/, '');

  return normalized.trim();
}

// ============================================================================
// DB 검증 관련 함수
// ============================================================================

/**
 * DESCRIPTION 정규화 (비교용)
 * - 대문자 변환
 * - 공백/하이픈 통일 (단일 공백)
 * - 따옴표 통일 (")
 *
 * @param {string} desc - Description 텍스트
 * @return {string} 정규화된 텍스트
 */
function normalizeSNGDescription(desc) {
  if (!desc) return '';

  return desc.toUpperCase()
             .replace(/[\s\-]+/g, ' ')      // 공백/하이픈 → 단일 공백
             .replace(/["″'']/g, '"')       // 따옴표 통일
             .trim();
}

/**
 * SNG 라인 아이템 DB 검증
 *
 * 규칙:
 * 1. ITEM CODE 정규화 (대문자만)
 * 2. ITEM CODE 형식 검증 (7자리 문자+숫자)
 * 3. DB 조회 (ITEM CODE 완전 일치)
 * 4. DB 매칭 성공 시:
 *    - DB DESCRIPTION으로 확정
 *    - 파싱 DESCRIPTION과 비교 (정규화 후)
 *    - 불일치 시 경고 로그만 (오류 아님)
 * 5. DB 미존재 시:
 *    - 파싱 DESCRIPTION 유지
 *    - Memo에 "⚠️ DB 미등록 제품" 추가
 *
 * @param {Object} item - 파싱된 라인 아이템
 * @param {Sheet} dbSheet - DB 시트 (DB_SNG) - 단일 검색 시 사용
 * @param {Object} dbMap - DB Map (배치 검색 시 사용) - buildSNGItemCodeMap() 결과
 * @return {Object} 검증 완료된 라인 아이템
 */
function validateSNGItemWithDB(item, dbSheet, dbMap) {
  // 1. ITEM CODE 정규화 (대문자만)
  var normalizedItemId = item.itemId.toUpperCase();

  // 2. ITEM CODE 형식 검증 (7자리 문자+숫자)
  if (!normalizedItemId.match(/^[A-Z0-9]{7}$/)) {
    Logger.log('❌ 잘못된 ITEM CODE 형식: ' + normalizedItemId + ' (길이: ' + normalizedItemId.length + ')');
    item.memo = (item.memo ? item.memo + ' / ' : '') + '❌ ITEM CODE 오류';
    return item; // DB 조회 생략, 파싱값 유지
  }

  // 3. DB 조회 (Map 우선, 없으면 직접 검색)
  var dbRecord = null;

  if (dbMap && dbMap.itemCodeMap) {
    // 배치 검색: Map에서 직접 조회 (O(1))
    dbRecord = dbMap.itemCodeMap[normalizedItemId];
  } else if (dbSheet) {
    // 단일 검색: DB 직접 검색 (O(n))
    var dbRecords = searchSNGDatabase(dbSheet, normalizedItemId);
    dbRecord = (dbRecords && dbRecords.length > 0) ? dbRecords[0] : null;
  }

  // 4a. DB 미존재
  if (!dbRecord) {
    Logger.log('⚠️ DB 미등록 제품: ' + normalizedItemId);
    Logger.log('  Parsed DESCRIPTION: ' + item.description);
    item.memo = (item.memo ? item.memo + ' / ' : '') + '⚠️ DB 미등록 제품';
    return item; // 파싱값 유지
  }

  // 4b. DB 매칭 성공
  var dbDescription = dbRecord.description || '';

  // 5. DESCRIPTION 더블체크 (정규화 후 비교)
  var parsedNorm = normalizeSNGDescription(item.description);
  var dbNorm = normalizeSNGDescription(dbDescription);

  if (parsedNorm !== dbNorm) {
    Logger.log('⚠️ DESCRIPTION 불일치 감지');
    Logger.log('  ITEM CODE: ' + normalizedItemId);
    Logger.log('  Parsed:    ' + item.description);
    Logger.log('  DB:        ' + dbDescription);
    Logger.log('  → DB 값으로 확정');
  } else {
    Logger.log('✅ DB 매칭 성공: ' + normalizedItemId);
  }

  // 6. DB DESCRIPTION으로 확정
  item.description = dbDescription;

  return item;
}

/**
 * SNG DB에서 ITEM CODE Map 생성 (배치 최적화)
 * - DB 전체를 한 번만 읽어서 Map 구조로 변환
 * - 여러 아이템 검증 시 성능 향상
 *
 * @param {Sheet} dbSheet - DB_SNG 시트
 * @return {Object} { itemCodeMap: {ITEMCODE: {description, color, barcode}}, columnMap: {} }
 */
function buildSNGItemCodeMap(dbSheet) {
  try {
    if (!dbSheet) {
      return { itemCodeMap: {}, columnMap: {} };
    }

    var data = dbSheet.getDataRange().getValues();

    if (data.length < 2) {
      debugLog('SNG DB 데이터 없음');
      return { itemCodeMap: {}, columnMap: {} };
    }

    var headers = data[0];
    var colMap = {};

    // 헤더 인덱스 매핑
    for (var i = 0; i < headers.length; i++) {
      colMap[headers[i]] = i;
    }

    // 필수 컬럼 확인
    var itemCodeCol = colMap['Item Code'];
    var descriptionCol = colMap['Description'];

    if (itemCodeCol === undefined || descriptionCol === undefined) {
      debugLog('SNG DB 필수 컬럼 누락', { itemCode: itemCodeCol, description: descriptionCol });
      return { itemCodeMap: {}, columnMap: colMap };
    }

    // ITEM CODE Map 생성
    var itemCodeMap = {};

    for (var i = 1; i < data.length; i++) {
      var rowItemCode = data[i][itemCodeCol];

      if (!rowItemCode) continue;

      var rowItemCodeNorm = rowItemCode.toString().toUpperCase();

      // CRITICAL: 동일 ITEM CODE는 첫 번째 레코드만 사용 (1:1 매핑)
      if (!itemCodeMap[rowItemCodeNorm]) {
        itemCodeMap[rowItemCodeNorm] = {
          itemCode: rowItemCode,
          description: data[i][descriptionCol] || '',
          color: data[i][colMap['Color']] || '',
          barcode: data[i][colMap['Barcode']] || ''
        };
      }
    }

    debugLog('SNG DB Map 생성 완료', { itemCount: Object.keys(itemCodeMap).length });

    return {
      itemCodeMap: itemCodeMap,
      columnMap: colMap
    };

  } catch (error) {
    debugLog('buildSNGItemCodeMap 오류', { error: error.toString() });
    logError('buildSNGItemCodeMap', error);
    return { itemCodeMap: {}, columnMap: {} };
  }
}

/**
 * SNG DB 검색 (ITEM CODE 기준)
 * DEPRECATED: buildSNGItemCodeMap + 직접 Map 조회 방식 권장 (배치 검색 시)
 *
 * @param {Sheet} dbSheet - DB_SNG 시트
 * @param {string} itemCode - 검색할 ITEM CODE (정규화 완료)
 * @return {Array<Object>} 매칭된 레코드 배열 [{description, color, barcode}, ...]
 */
function searchSNGDatabase(dbSheet, itemCode) {
  try {
    if (!dbSheet || !itemCode) {
      return [];
    }

    // 배치 검색용 Map 생성
    var dbMap = buildSNGItemCodeMap(dbSheet);

    // Map에서 조회
    var record = dbMap.itemCodeMap[itemCode];

    if (!record) {
      debugLog('SNG DB 검색 결과 없음', { itemCode: itemCode });
      return [];
    }

    debugLog('SNG DB 검색 결과', { itemCode: itemCode, count: 1 });

    return [record];

  } catch (error) {
    debugLog('searchSNGDatabase 오류', { error: error.toString() });
    logError('searchSNGDatabase', error, { itemCode: itemCode });
    return [];
  }
}

// ============================================================================
// 테스트 함수
// ============================================================================

/**
 * SNG DB 검증 테스트
 * - 실제 DB_SNG 시트를 사용하여 검증 로직 테스트
 * - Logger 출력으로 결과 확인
 */
function testSNGDBValidation() {
  Logger.log('========================================');
  Logger.log('SNG DB 검증 테스트 시작');
  Logger.log('========================================');

  try {
    // DB_SNG 시트 가져오기
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var dbSheet = ss.getSheetByName(CONFIG.COMPANIES.SNG.dbSheet);

    if (!dbSheet) {
      Logger.log('❌ DB_SNG 시트를 찾을 수 없습니다.');
      return;
    }

    Logger.log('✅ DB_SNG 시트 로드 성공');
    Logger.log('');

    // ========================================
    // 테스트 케이스 1: DB에 존재하는 ITEM CODE (정상 케이스)
    // ========================================
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('테스트 1: DB 매칭 성공 (정상 케이스)');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    var testItem1 = {
      lineNo: 1,
      itemId: 'SBGBOHX',  // 실제 DB에 존재하는 ITEM CODE (BG BOHO BANG)
      description: 'BG BOHO BANG',  // 파싱된 Description (임시)
      brand: 'Shake-N-Go',
      color: '1B',
      qtyShipped: 10,
      unitPrice: 4.00,
      extPrice: 40.00,
      memo: ''
    };

    var validatedItem1 = validateSNGItemWithDB(testItem1, dbSheet, null);

    Logger.log('파싱 DESCRIPTION: ' + testItem1.description);
    Logger.log('검증 후 DESCRIPTION: ' + validatedItem1.description);
    Logger.log('Memo: ' + validatedItem1.memo);
    Logger.log('');

    // ========================================
    // 테스트 케이스 2: DB에 없는 ITEM CODE (신제품)
    // ========================================
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('테스트 2: DB 미등록 제품 (신제품)');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    var testItem2 = {
      lineNo: 2,
      itemId: 'NEWPROD',  // 7자리, DB에 없음
      description: 'SHAKE-N-GO NEW PRODUCT 24"',
      brand: 'Shake-N-Go',
      color: '1B',
      qtyShipped: 5,
      unitPrice: 5.00,
      extPrice: 25.00,
      memo: ''
    };

    var validatedItem2 = validateSNGItemWithDB(testItem2, dbSheet, null);

    Logger.log('파싱 DESCRIPTION: ' + testItem2.description);
    Logger.log('검증 후 DESCRIPTION: ' + validatedItem2.description);
    Logger.log('Memo: ' + validatedItem2.memo);
    Logger.log('');

    // ========================================
    // 테스트 케이스 3: ITEM CODE 형식 오류 (5자리)
    // ========================================
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('테스트 3: ITEM CODE 형식 오류 (7자리 아님)');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    var testItem3 = {
      lineNo: 3,
      itemId: 'SKB12',  // 5자리 - 오류
      description: 'SHAKE-N-GO SOMETHING',
      brand: 'Shake-N-Go',
      color: 'BLACK',
      qtyShipped: 3,
      unitPrice: 3.00,
      extPrice: 9.00,
      memo: ''
    };

    var validatedItem3 = validateSNGItemWithDB(testItem3, dbSheet, null);

    Logger.log('파싱 DESCRIPTION: ' + testItem3.description);
    Logger.log('검증 후 DESCRIPTION: ' + validatedItem3.description);
    Logger.log('Memo: ' + validatedItem3.memo);
    Logger.log('');

    // ========================================
    // 테스트 케이스 4: DESCRIPTION 불일치 (정규화 후 비교)
    // ========================================
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
    Logger.log('테스트 4: DESCRIPTION 불일치 (정규화 테스트)');
    Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');

    var testItem4 = {
      lineNo: 4,
      itemId: 'SBGBOHX',  // 실제 DB에 존재하는 ITEM CODE
      description: 'BG  BOHO-BANG',  // 공백/하이픈 차이 (정규화 테스트용)
      brand: 'Shake-N-Go',
      color: '1',
      qtyShipped: 2,
      unitPrice: 4.00,
      extPrice: 8.00,
      memo: ''
    };

    var validatedItem4 = validateSNGItemWithDB(testItem4, dbSheet, null);

    Logger.log('파싱 DESCRIPTION: ' + testItem4.description);
    Logger.log('검증 후 DESCRIPTION: ' + validatedItem4.description);
    Logger.log('Memo: ' + validatedItem4.memo);
    Logger.log('');

    // ========================================
    // 최종 요약
    // ========================================
    Logger.log('========================================');
    Logger.log('테스트 완료');
    Logger.log('========================================');
    Logger.log('✅ 모든 테스트 케이스 실행 완료');
    Logger.log('');
    Logger.log('확인 사항:');
    Logger.log('1. 테스트 1: DB DESCRIPTION으로 교체되었는가? (BG BOHO BANG)');
    Logger.log('2. 테스트 2: 파싱 DESCRIPTION 유지 + Memo에 경고 표시?');
    Logger.log('3. 테스트 3: 파싱 DESCRIPTION 유지 + Memo에 오류 표시?');
    Logger.log('4. 테스트 4: 정규화 후 불일치 로그 + DB DESCRIPTION으로 교체?');

  } catch (error) {
    Logger.log('❌ 테스트 실행 오류: ' + error.toString());
    Logger.log(error.stack);
  }
}

/**
 * SNG 정규화 함수 단위 테스트
 */
function testSNGDescriptionNormalization() {
  Logger.log('========================================');
  Logger.log('DESCRIPTION 정규화 테스트');
  Logger.log('========================================');

  var testCases = [
    {
      input: 'SHAKE-N-GO FREETRESS',
      expected: 'SHAKE N GO FREETRESS'
    },
    {
      input: 'SHAKE  N  GO   FREETRESS',
      expected: 'SHAKE N GO FREETRESS'
    },
    {
      input: 'SHAKE-N-GO  FREETRESS--EQUAL',
      expected: 'SHAKE N GO FREETRESS EQUAL'
    },
    {
      input: 'BRAID 10″',
      expected: 'BRAID 10"'
    },
    {
      input: 'shake-n-go freetress',
      expected: 'SHAKE N GO FREETRESS'
    }
  ];

  var passed = 0;
  var failed = 0;

  for (var i = 0; i < testCases.length; i++) {
    var tc = testCases[i];
    var result = normalizeSNGDescription(tc.input);

    if (result === tc.expected) {
      Logger.log('✅ PASS: "' + tc.input + '" → "' + result + '"');
      passed++;
    } else {
      Logger.log('❌ FAIL: "' + tc.input + '"');
      Logger.log('   Expected: "' + tc.expected + '"');
      Logger.log('   Got:      "' + result + '"');
      failed++;
    }
  }

  Logger.log('');
  Logger.log('========================================');
  Logger.log('테스트 결과: ' + passed + '/' + testCases.length + ' 통과');
  Logger.log('========================================');
}
