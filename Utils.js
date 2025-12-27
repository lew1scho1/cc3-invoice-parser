// ============================================================================
// UTILS.GS - 유틸리티 함수
// ============================================================================

/**
 * 시트 가져오기
 */
function getSheet(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(sheetName + ' 시트를 찾을 수 없습니다.');
    }
    
    return sheet;
  } catch (error) {
    debugLog('getSheet 오류', { sheet: sheetName, error: error.toString() });
    throw error;
  }
}

/**
 * 스프레드시트 가져오기
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (error) {
    debugLog('getSpreadsheet 오류', { error: error.toString() });
    throw error;
  }
}

/**
 * 컬럼 인덱스 맵 생성
 */
function getColumnMap(headers) {
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    if (headers[i]) {
      map[headers[i]] = i;
    }
  }
  return map;
}

/**
 * 바코드 유효성 검사
 */
function isValidBarcode(barcode) {
  if (!barcode) return false;
  var str = barcode.toString().trim();
  return str.length === CONFIG.BARCODE_LENGTH && /^\d+$/.test(str);
}

/**
 * 인치 패턴 제거
 */
function removeInchPattern(itemName) {
  if (!itemName || typeof itemName !== 'string') {
    return '';
  }
  return itemName.replace(/\s*\d+\s*[""]?\s*/g, ' ').replace(/\s+/g, ' ').trim();
}

/**
 * 문자열에서 인치(길이) 추출
 */
function extractInches(itemName) {
  if (!itemName || typeof itemName !== 'string') {
    return 0;
  }
  var match = itemName.match(/(\d+)\s*[""]/);
  return match ? parseInt(match[1]) : 0;
}

/**
 * 숫자를 2자리로 패딩
 */
function padZero(num) {
  return num < 10 ? '0' + num : num.toString();
}

/**
 * COLOR 파싱 및 정렬 키 생성
 */
function parseColorForSorting(color) {
  if (!color) {
    return {
      priority: 999,
      prefix: 'ZZZ',
      priorityNums: [],
      otherNums: [],
      text: '',
      original: ''
    };
  }
  
  var colorStr = String(color).toUpperCase().trim();
  var original = colorStr;
  
  var priorityNumbers = ['1B', '1', '2', '4', '27', '30', '530', '613', '130', '350', '33'];
  
  if (colorStr.indexOf('+') > -1) {
    return {
      priority: 50,
      prefix: 'COMBO',
      priorityNums: [],
      otherNums: [],
      text: colorStr,
      original: original
    };
  }
  
  var prefixMatch = colorStr.match(/^(T|P|M|OT|DR|OM|OF|FS|GF|SOM)/);
  var prefix = prefixMatch ? prefixMatch[1] : '';
  var remaining = prefix ? colorStr.substring(prefix.length) : colorStr;
  
  var prefixPriority = {
    'T': 1,
    'P': 2,
    'M': 3,
    'OT': 4
  };
  
  var prefixOrder = prefixPriority[prefix] || 99;
  
  var foundPriorityNums = [];
  var tempRemaining = remaining;
  
  for (var i = 0; i < priorityNumbers.length; i++) {
    var prioNum = priorityNumbers[i];
    if (tempRemaining.indexOf(prioNum) === 0) {
      foundPriorityNums.push(prioNum);
      tempRemaining = tempRemaining.substring(prioNum.length);
      i = -1;
    }
  }
  
  var otherNums = [];
  var numMatches = tempRemaining.match(/\d+/g);
  if (numMatches) {
    for (var i = 0; i < numMatches.length; i++) {
      otherNums.push(parseInt(numMatches[i]));
    }
  }
  
  var textOnly = tempRemaining.replace(/\d+/g, '').trim();
  
  var overallPriority = 10;
  
  if (foundPriorityNums.length > 0 && prefix === '' && otherNums.length === 0 && textOnly === '') {
    overallPriority = 1;
  } else if (prefix === '' && textOnly === '' && otherNums.length > 0) {
    overallPriority = 2;
  } else if (prefix !== '' && (foundPriorityNums.length > 0 || otherNums.length > 0)) {
    overallPriority = 3;
  } else if (textOnly !== '' && otherNums.length === 0 && foundPriorityNums.length === 0) {
    overallPriority = 4;
  }
  
  return {
    priority: overallPriority,
    prefix: prefix || 'ZZZ',
    prefixOrder: prefixOrder,
    priorityNums: foundPriorityNums,
    otherNums: otherNums,
    text: textOnly,
    original: original
  };
}

/**
 * COLOR 비교 함수
 */
function compareColors(colorA, colorB) {
  var a = parseColorForSorting(colorA);
  var b = parseColorForSorting(colorB);
  
  if (a.priority !== b.priority) {
    return a.priority - b.priority;
  }
  
  if (a.priority === 1) {
    var priorityOrder = ['1', '1B', '2', '4', '27', '30', '530', '613', '130', '350', '33'];
    var aIndex = priorityOrder.indexOf(a.original);
    var bIndex = priorityOrder.indexOf(b.original);
    return aIndex - bIndex;
  }
  
  if (a.priority === 2) {
    var aNum = a.otherNums.length > 0 ? a.otherNums[0] : 0;
    var bNum = b.otherNums.length > 0 ? b.otherNums[0] : 0;
    return aNum - bNum;
  }
  
  if (a.priority === 3) {
    if (a.prefixOrder !== b.prefixOrder) {
      return a.prefixOrder - b.prefixOrder;
    }
    
    var priorityOrder = ['1', '1B', '2', '4', '27', '30', '530', '613', '130', '350', '33'];
    
    for (var i = 0; i < Math.max(a.priorityNums.length, b.priorityNums.length); i++) {
      var aPrio = a.priorityNums[i] || '';
      var bPrio = b.priorityNums[i] || '';
      
      if (aPrio !== bPrio) {
        var aIdx = priorityOrder.indexOf(aPrio);
        var bIdx = priorityOrder.indexOf(bPrio);
        
        if (aIdx === -1 && bIdx === -1) {
          return aPrio.localeCompare(bPrio);
        }
        if (aIdx === -1) return 1;
        if (bIdx === -1) return -1;
        return aIdx - bIdx;
      }
    }
    
    for (var i = 0; i < Math.max(a.otherNums.length, b.otherNums.length); i++) {
      var aNum = a.otherNums[i] || 0;
      var bNum = b.otherNums[i] || 0;
      if (aNum !== bNum) {
        return aNum - bNum;
      }
    }
    
    return a.text.localeCompare(b.text);
  }
  
  if (a.priority === 4) {
    return a.text.localeCompare(b.text);
  }
  
  return a.original.localeCompare(b.original);
}

/**
 * 제품 정렬 함수 (길이 → 컬러 → 제품명)
 */
function sortProducts(products) {
  return products.sort(function(a, b) {
    var inchA = extractInches(a.itemName);
    var inchB = extractInches(b.itemName);
    
    if (inchA !== inchB) {
      return inchA - inchB;
    }
    
    var colorCompare = compareColors(a.color, b.color);
    if (colorCompare !== 0) {
      return colorCompare;
    }
    
    var nameA = String(a.itemName || '');
    var nameB = String(b.itemName || '');
    return nameA.localeCompare(nameB);
  });
}