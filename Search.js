// ============================================================================
// SEARCH.GS - 검색 관련 함수
// ============================================================================

/**
 * 특정 회사 DB에서 바코드 검색
 */
function searchBarcodeInCompany(barcode, companyKey) {
  try {
    debugLog('searchBarcodeInCompany 시작', { barcode: barcode, company: companyKey });
    
    var company = CONFIG.COMPANIES[companyKey];
    if (!company) {
      throw new Error('Company not found: ' + companyKey);
    }
    
    var barcodeStr = barcode.toString().trim();
    
    var sheet = getDbSheet(company.dbSheet);
    var data = sheet.getDataRange().getValues();
    
    debugLog('시트 데이터 로드', { sheet: company.dbSheet, rows: data.length });
    
    if (data.length < 2) {
      debugLog('데이터 없음');
      return null;
    }
    
    var headers = data[0];
    var colMap = getColumnMap(headers);
    
    debugLog('컬럼 맵', colMap);
    
    var requiredCols = [
      company.columns.ITEM_NUMBER,
      company.columns.ITEM_NAME,
      company.columns.COLOR,
      company.columns.BARCODE
    ];
    
    for (var i = 0; i < requiredCols.length; i++) {
      if (colMap[requiredCols[i]] === undefined) {
        debugLog('필수 컬럼 누락', { column: requiredCols[i] });
        return null;
      }
    }
    
    var foundProduct = null;
    for (var i = 1; i < data.length; i++) {
      var rowBarcode = data[i][colMap[company.columns.BARCODE]];
      if (rowBarcode && rowBarcode.toString().trim() === barcodeStr) {
        foundProduct = {
          itemNumber: data[i][colMap[company.columns.ITEM_NUMBER]],
          itemName: data[i][colMap[company.columns.ITEM_NAME]],
          color: data[i][colMap[company.columns.COLOR]],
          barcode: rowBarcode,
          company: companyKey,
          status: data[i][colMap['status']] || CONFIG.STATUS.ACTIVE
        };
        break;
      }
    }
    
    if (!foundProduct) {
      debugLog('바코드 제품 찾을 수 없음');
      return null;
    }
    
    debugLog('제품 발견', foundProduct);
    
    var baseItemName = foundProduct.itemName;
    var products = [];
    var barcodeMap = {};
    
    for (var i = 1; i < data.length; i++) {
      var currentItemName = data[i][colMap[company.columns.ITEM_NAME]];
      var currentBarcode = data[i][colMap[company.columns.BARCODE]];
      
      if (currentItemName === baseItemName && currentBarcode) {
        var key = currentBarcode.toString();
        if (!barcodeMap[key]) {
          barcodeMap[key] = true;
          products.push({
            itemNumber: data[i][colMap[company.columns.ITEM_NUMBER]],
            itemName: currentItemName,
            color: data[i][colMap[company.columns.COLOR]] || '',
            barcode: currentBarcode,
            quantity: 0,
            company: companyKey,
            status: data[i][colMap['status']] || CONFIG.STATUS.ACTIVE
          });
        }
      }
    }
    
    debugLog('동일 ITEM NAME 제품', { count: products.length });
    
    var uniqueColors = {};
    for (var i = 0; i < products.length; i++) {
      var colorKey = String(products[i].color || '').toUpperCase().trim();
      if (colorKey) {
        uniqueColors[colorKey] = true;
      }
    }
    var colorCount = Object.keys(uniqueColors).length;
    
    debugLog('고유 COLOR 개수', { count: colorCount });
    
    var skipExpansion = false;
    if (companyKey === 'OUTRE' && baseItemName) {
      var baseNameUpper = baseItemName.toString().toUpperCase();
      if (baseNameUpper.indexOf('X-PRESSION BRAID-PRE STRETCHED BRAID') !== -1) {
        skipExpansion = true;
        debugLog('확장 로직 스킵 (X-PRESSION BRAID-PRE STRETCHED BRAID)', { itemName: baseItemName });
      }
    }
    
    if (!skipExpansion && colorCount <= CONFIG.MAX_COLORS_FOR_EXPANSION) {
      var simplifiedBase = removeInchPattern(baseItemName);
      debugLog('확장 검색 시작', { simplified: simplifiedBase });
      
      if (simplifiedBase && simplifiedBase !== baseItemName) {
        for (var i = 1; i < data.length; i++) {
          var currentItemName = data[i][colMap[company.columns.ITEM_NAME]];
          var currentBarcode = data[i][colMap[company.columns.BARCODE]];
          
          if (!currentItemName || !currentBarcode) continue;
          
          var currentSimplified = removeInchPattern(currentItemName);
          var key = currentBarcode.toString();
          
          if (currentSimplified === simplifiedBase && !barcodeMap[key]) {
            barcodeMap[key] = true;
            products.push({
              itemNumber: data[i][colMap[company.columns.ITEM_NUMBER]],
              itemName: currentItemName,
              color: data[i][colMap[company.columns.COLOR]] || '',
              barcode: currentBarcode,
              quantity: 0,
              company: companyKey,
              status: data[i][colMap['status']] || CONFIG.STATUS.ACTIVE
            });
          }
        }
        
        debugLog('확장 검색 후', { count: products.length });
      }
    }
    
    products = sortProducts(products);
    products = attachInvoiceMetricsToProducts(products, companyKey);
    products = buildGroupedProducts(products, baseItemName, companyKey);

    // 통계 계산
    var result = calculateColorStatistics(products);

    debugLog('searchBarcodeInCompany 완료', { resultCount: result.products.length });

    return result;
    
  } catch (error) {
    debugLog('searchBarcodeInCompany 오류', { error: error.toString() });
    logError('searchBarcodeInCompany', error.toString(), { barcode: barcode, company: companyKey });
    return null;
  }
}

/**
 * 바코드로 제품 검색
 */
/**
 * Insert group header rows when multiple item names are present.
 *
 * @param {Array<Object>} products
 * @param {string} baseItemName
 * @param {string} companyKey
 * @return {Array<Object>}
 */
function buildGroupedProducts(products, baseItemName, companyKey) {
  if (!products || products.length === 0) return products;
  
  var groupMap = {};
  var order = [];
  
  for (var i = 0; i < products.length; i++) {
    var itemName = products[i].itemName || '';
    if (!groupMap[itemName]) {
      groupMap[itemName] = [];
      order.push(itemName);
    }
    groupMap[itemName].push(products[i]);
  }
  
  if (order.length <= 1) return products;
  
  if (baseItemName && groupMap[baseItemName]) {
    var filtered = [];
    for (var j = 0; j < order.length; j++) {
      if (order[j] !== baseItemName) filtered.push(order[j]);
    }
    filtered.unshift(baseItemName);
    order = filtered;
  }
  
  var grouped = [];
  for (var k = 0; k < order.length; k++) {
    grouped.push({
      isGroupHeader: true,
      itemName: order[k],
      company: companyKey
    });
    grouped = grouped.concat(groupMap[order[k]]);
  }
  
  return grouped;
}

function normalizeInvoiceUpc(value) {
  if (value === null || value === undefined) return '';
  var str = value.toString().trim();
  if (str === '') return '';
  return str.replace(/^0+/, '');
}

function formatPriceValue(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'number') {
    if (isNaN(value)) return '';
    if (Math.floor(value) === value) return value.toString();
    return value.toFixed(2).replace(/0+$/, '').replace(/\.$/, '');
  }
  return String(value);
}

function buildPriceHistoryString(priceEntries) {
  if (!priceEntries || priceEntries.length === 0) return '';
  
  priceEntries.sort(function(a, b) {
    return a.date.getTime() - b.date.getTime();
  });
  
  var uniquePrices = [];
  var lastPrice = null;
  
  for (var i = 0; i < priceEntries.length; i++) {
    var currentPrice = priceEntries[i].price;
    if (lastPrice === null || currentPrice !== lastPrice) {
      uniquePrices.push(currentPrice);
      lastPrice = currentPrice;
    }
  }
  
  if (uniquePrices.length === 1) {
    return formatPriceValue(uniquePrices[0]);
  }
  
  var parts = [];
  for (var j = 0; j < uniquePrices.length; j++) {
    parts.push(formatPriceValue(uniquePrices[j]));
  }

  return parts.join(' -> ');
}

/**
 * 컬러별 상태 통계 계산
 * @param {Array} products - 제품 배열
 * @return {Object} { products: [...], colorStats: {...} }
 */
function calculateColorStatistics(products) {
  if (!products || products.length === 0) {
    return { products: products, colorStats: null };
  }

  var stats = {
    total: 0,
    active: 0,
    new: 0,
    discontinueUnknown: 0,
    discontinued: 0
  };

  for (var i = 0; i < products.length; i++) {
    if (products[i].isGroupHeader) continue;

    stats.total++;

    var status = products[i].status || CONFIG.STATUS.ACTIVE;
    if (status === CONFIG.STATUS.NEW) {
      stats.new++;
      stats.active++; // NEW도 활성으로 간주
    } else if (status === CONFIG.STATUS.ACTIVE) {
      stats.active++;
    } else if (status === CONFIG.STATUS.DISCONTINUE_UNKNOWN) {
      stats.discontinueUnknown++;
    } else if (status === CONFIG.STATUS.DISCONTINUED) {
      stats.discontinued++;
    }
  }

  return { products: products, colorStats: stats };
}

function buildInvoiceMetricsMap(companyKey) {
  var sheetName = companyKey === 'OUTRE' ? CONFIG.INVOICE.OUTRE_SHEET : CONFIG.INVOICE.SNG_SHEET;
  var sheet = getSheet(sheetName);
  var data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    return {};
  }
  
  var headers = data[0];
  var colMap = getColumnMap(headers);
  
  var dateCol = colMap['Invoice Date'];
  var upcCol = colMap['UPC'];
  var qtyCol = colMap['Qty Shipped'];
  var priceCol = colMap['Unit Price'];
  var orderedCol = colMap['Qty Ordered'];
  
  if (dateCol === undefined || upcCol === undefined || qtyCol === undefined || priceCol === undefined) {
    debugLog('invoice columns missing', { sheet: sheetName });
    return {};
  }
  
  var now = new Date();
  var shippedStart = new Date(now.getTime());
  shippedStart.setMonth(shippedStart.getMonth() - 12);
  
  var backorderStart = new Date(now.getTime());
  backorderStart.setMonth(backorderStart.getMonth() - 2);
  
  var metrics = {};
  
  for (var i = 1; i < data.length; i++) {
    var invoiceDate = data[i][dateCol];
    if (!(invoiceDate instanceof Date) || isNaN(invoiceDate.getTime())) continue;
    
    var withinShippedWindow = invoiceDate >= shippedStart && invoiceDate <= now;
    var withinBackorderWindow = invoiceDate >= backorderStart && invoiceDate <= now;
    
    if (!withinShippedWindow && !withinBackorderWindow) continue;
    
    var upcRaw = data[i][upcCol];
    if (!upcRaw) continue;
    
    var upcKey = normalizeInvoiceUpc(upcRaw);
    if (!upcKey) continue;
    
    var qtyValue = data[i][qtyCol];
    var qty = parseFloat(qtyValue);
    if (isNaN(qty)) qty = 0;
    
    if (!metrics[upcKey]) {
      metrics[upcKey] = { shipped: 0, prices: [], backorder: 0, backorderDate: null };
    }
    
    if (withinShippedWindow) {
      var priceValue = data[i][priceCol];
      var price = parseFloat(priceValue);
      
      metrics[upcKey].shipped += qty;
      
      if (!isNaN(price)) {
        metrics[upcKey].prices.push({ date: invoiceDate, price: price });
      }
    }
    
    if (withinBackorderWindow && orderedCol !== undefined) {
      var orderedValue = data[i][orderedCol];
      var orderedQty = parseFloat(orderedValue);
      if (isNaN(orderedQty)) orderedQty = 0;
      
      var backorderQty = orderedQty - qty;
      if (backorderQty > 0) {
        if (!metrics[upcKey].backorderDate || invoiceDate > metrics[upcKey].backorderDate) {
          metrics[upcKey].backorder = backorderQty;
          metrics[upcKey].backorderDate = invoiceDate;
        }
      }
    }
  }
  
  var result = {};
  for (var key in metrics) {
    if (!metrics.hasOwnProperty(key)) continue;
    result[key] = {
      shipped12mo: metrics[key].shipped,
      priceHistory: buildPriceHistoryString(metrics[key].prices),
      backorder2mo: metrics[key].backorder
    };
  }
  
  return result;
}

function buildOrderQuantityMap(companyKey) {
  var company = CONFIG.COMPANIES[companyKey];
  if (!company) return {};
  
  var sheet = getSheet(company.orderSheet);
  var data = sheet.getDataRange().getValues();
  
  if (data.length < 2) return {};
  
  var headers = data[0];
  var colMap = getColumnMap(headers);
  var upcCol = colMap['UPC'];
  var qtyCol = colMap['수량'];
  if (qtyCol === undefined) {
    qtyCol = colMap['?˜ëŸ‰'];
  }
  
  if (upcCol === undefined || qtyCol === undefined) return {};
  
  var orderMap = {};
  for (var i = 1; i < data.length; i++) {
    var upcRaw = data[i][upcCol];
    if (!upcRaw) continue;
    
    var upcKey = normalizeInvoiceUpc(upcRaw);
    if (!upcKey) continue;
    
    var qtyValue = data[i][qtyCol];
    var qty = parseFloat(qtyValue);
    if (isNaN(qty)) qty = 0;
    
    if (!orderMap[upcKey]) {
      orderMap[upcKey] = 0;
    }
    orderMap[upcKey] += qty;
  }
  
  return orderMap;
}

function attachInvoiceMetricsToProducts(products, companyKey) {
  if (!products || products.length === 0) return products;
  
  var metricsMap = buildInvoiceMetricsMap(companyKey);
  var orderMap = buildOrderQuantityMap(companyKey);
  
  for (var i = 0; i < products.length; i++) {
    var upcKey = normalizeInvoiceUpc(products[i].barcode);
    var metrics = metricsMap[upcKey];
    var orderQty = orderMap[upcKey] || 0;
    
    products[i].shipped12mo = metrics ? metrics.shipped12mo : 0;
    products[i].priceHistory = metrics ? metrics.priceHistory : '';
    products[i].backorder2mo = metrics ? metrics.backorder2mo : 0;
    
    if (orderQty > 0) {
      products[i].quantity = orderQty;
      products[i].cartQuantity = orderQty;
    } else if (products[i].backorder2mo > 0) {
      products[i].quantity = products[i].backorder2mo;
      products[i].backorderApplied = true;
    }
  }
  
  return products;
}

function searchBarcode(barcode) {
  try {
    debugLog('searchBarcode 시작', { barcode: barcode });
    
    if (!isValidBarcode(barcode)) {
      return {
        success: false,
        error: 'Please enter a valid 12-digit barcode.'
      };
    }
    
    var barcodeStr = barcode.toString().trim();
    
    var outreResult = searchBarcodeInCompany(barcodeStr, 'OUTRE');
    var sngResult = searchBarcodeInCompany(barcodeStr, 'SNG');

    if (outreResult && outreResult.products && outreResult.products.length > 0) {
      debugLog('OUTRE에서 발견');
      return buildSearchResponse(barcodeStr, 'OUTRE', outreResult);
    }

    if (sngResult && sngResult.products && sngResult.products.length > 0) {
      debugLog('SNG에서 발견');
      return buildSearchResponse(barcodeStr, 'SNG', sngResult);
    }

    return {
      success: false,
      error: 'Not found in OUTRE or SNG databases.'
    };
    
  } catch (error) {
    debugLog('searchBarcode 오류', { error: error.toString() });
    logError('searchBarcode', error.toString(), { barcode: barcode });
    
    return {
      success: false,
      error: 'An error occurred while searching.\n' + error.toString()
    };
  }
}

/**
 * 텍스트로 제품 검색
 */
function searchByText(searchText, companyKey) {
  try {
    debugLog('searchByText 시작', { text: searchText, company: companyKey });
    
    if (!searchText || searchText.trim() === '') {
      return {
        success: false,
        error: 'Please enter a search term.'
      };
    }
    
    var company = CONFIG.COMPANIES[companyKey];
    if (!company) {
      throw new Error('Company not found: ' + companyKey);
    }
    
    // 토큰 AND 매칭: 검색어를 공백 기준으로 나눠 모든 단어가 포함된 결과 찾기
    var normalizedTokens = normalizeSearchTokens(searchText);
    debugLog('검색 토큰', normalizedTokens);

    var sheet = getDbSheet(company.dbSheet);
    var data = sheet.getDataRange().getValues();

    debugLog('시트 데이터 로드', { sheet: company.dbSheet, rows: data.length });

    if (data.length < 2) {
      return {
        success: false,
        error: 'No data available.'
      };
    }

    var headers = data[0];
    var colMap = getColumnMap(headers);

    var requiredCols = [
      company.columns.ITEM_NUMBER,
      company.columns.ITEM_NAME,
      company.columns.COLOR,
      company.columns.BARCODE
    ];

    for (var i = 0; i < requiredCols.length; i++) {
      if (colMap[requiredCols[i]] === undefined) {
        debugLog('필수 컬럼 누락', { column: requiredCols[i] });
        return {
          success: false,
          error: 'Invalid data format.'
        };
      }
    }

    var foundProducts = [];
    var barcodeMap = {};

    for (var i = 1; i < data.length; i++) {
      var itemNumber = data[i][colMap[company.columns.ITEM_NUMBER]];
      var itemName = data[i][colMap[company.columns.ITEM_NAME]];
      var color = data[i][colMap[company.columns.COLOR]];
      var barcode = data[i][colMap[company.columns.BARCODE]];

      if (!barcode) continue;

      // 검색 대상 정규화 (공백/하이픈 제거, 대문자 변환)
      var normalizedTarget = normalizeSearchQuery(
        [itemNumber || '', itemName || '', color || ''].join(' ')
      );

      // 모든 토큰이 포함되는지 확인 (AND 매칭)
      var allTokensFound = true;
      for (var k = 0; k < normalizedTokens.length; k++) {
        if (normalizedTarget.indexOf(normalizedTokens[k]) === -1) {
          allTokensFound = false;
          break;
        }
      }

      if (allTokensFound) {
        var key = barcode.toString();
        if (!barcodeMap[key]) {
          barcodeMap[key] = true;
          foundProducts.push({
            itemNumber: itemNumber,
            itemName: itemName,
            color: color,
            barcode: barcode,
            quantity: 0,
            company: companyKey
          });
        }
      }
    }
    
    debugLog('검색 결과', { count: foundProducts.length });
    
    if (foundProducts.length === 0) {
      return {
        success: false,
        error: '"' + searchText + '" was not found.'
      };
    }
    
    foundProducts = sortProducts(foundProducts);
    foundProducts = attachInvoiceMetricsToProducts(foundProducts, companyKey);
    
    return {
      success: true,
      products: foundProducts,
      company: companyKey
    };
    
  } catch (error) {
    debugLog('searchByText 오류', { error: error.toString() });
    logError('searchByText', error.toString(), { text: searchText, company: companyKey });
    
    return {
      success: false,
      error: 'An error occurred while searching.\n' + error.toString()
    };
  }
}

/**
 * 통합 검색 (바코드 + 텍스트)
 */
function searchProduct(searchInput) {
  try {
    debugLog('searchProduct 시작', { input: searchInput });
    
    if (!searchInput || searchInput.trim() === '') {
      return {
        success: false,
        error: 'Please enter a search term.'
      };
    }
    
    var input = searchInput.toString().trim();
    
    if (input.length === 12 && /^\d{12}$/.test(input)) {
      debugLog('바코드 검색 모드');
      return searchBarcode(input);
    } else {
      debugLog('텍스트 검색 모드');
      
      var outreResult = searchByText(input, 'OUTRE');
      if (outreResult.success) {
        return outreResult;
      }
      
      var sngResult = searchByText(input, 'SNG');
      if (sngResult.success) {
        return sngResult;
      }
      
      return {
        success: false,
        error: '"' + input + '" was not found.\n\nSearched in OUTRE and SNG databases.'
      };
    }
    
  } catch (error) {
    debugLog('searchProduct 오류', { error: error.toString() });
    logError('searchProduct', error.toString(), { input: searchInput });

    return {
      success: false,
      error: 'An error occurred while searching.\n' + error.toString()
    };
  }
}

// ============================================================================
// 응답 구조 빌더 (스키마 표준화)
// ============================================================================

/**
 * 검색 응답을 표준 스키마로 변환
 * @param {string} scannedBarcode - 스캔한 바코드
 * @param {string} company - 회사 키 (OUTRE, SNG)
 * @param {object} searchResult - searchBarcodeInCompany 결과 {products, colorStats}
 * @return {object} 표준 응답 스키마
 */
function buildSearchResponse(scannedBarcode, company, searchResult) {
  var products = searchResult.products || [];
  var colorStats = searchResult.colorStats || null;

  // 스캔한 바코드와 일치하는 제품 찾기 (scannedProduct)
  var scannedProduct = null;
  for (var i = 0; i < products.length; i++) {
    if (products[i].isGroupHeader) continue;
    if (products[i].barcode === scannedBarcode) {
      scannedProduct = {
        barcode: products[i].barcode,
        status: products[i].status || CONFIG.STATUS.ACTIVE,
        itemName: products[i].itemName,
        color: products[i].color
      };
      break;
    }
  }

  // 알림 판단 (서버에서 결정)
  var alert = null;
  if (scannedProduct) {
    if (scannedProduct.status === CONFIG.STATUS.DISCONTINUED) {
      alert = {
        type: 'error',
        title: 'Discontinued product',
        message: 'This product is discontinued and cannot be ordered.\n\nPlease choose another color.'
      };
    } else if (scannedProduct.status === CONFIG.STATUS.DISCONTINUE_UNKNOWN) {
      alert = {
        type: 'warning',
        title: 'Discontinue soon',
        message: 'This product is marked as discontinue soon.\nPlease check stock before ordering.'
      };
    } else if (scannedProduct.status === CONFIG.STATUS.NEW) {
      alert = {
        type: 'info',
        title: 'New product',
        message: 'This is a newly added product.'
      };
    }
  }

  // 표준 응답 스키마 반환
  return {
    success: true,
    company: company,
    products: products,
    meta: {
      scannedBarcode: scannedBarcode,
      scannedProduct: scannedProduct,
      colorStats: colorStats,
      alert: alert
    }
  };
}

// ============================================================================
// Keyword Search Index (Hidden Sheet) + Keyword Search APIs
// ============================================================================

function getSearchIndexSheetName(companyKey) {
  return CONFIG.SEARCH_INDEX[companyKey];
}

function getOrCreateSearchIndexSheet(companyKey) {
  var sheetName = getSearchIndexSheetName(companyKey);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.hideSheet();
  }
  return sheet;
}

function getSearchIndexSheet(companyKey) {
  var sheetName = getSearchIndexSheetName(companyKey);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('Search index sheet not found: ' + sheetName);
  }
  return sheet;
}

function setSearchIndexDate(companyKey, uploadDate) {
  var props = PropertiesService.getScriptProperties();
  var key = companyKey === 'OUTRE' ? 'SEARCH_INDEX_DATE_OUTRE' : 'SEARCH_INDEX_DATE_SNG';
  props.setProperty(key, uploadDate);
}

function getSearchIndexDate(companyKey) {
  var props = PropertiesService.getScriptProperties();
  var key = companyKey === 'OUTRE' ? 'SEARCH_INDEX_DATE_OUTRE' : 'SEARCH_INDEX_DATE_SNG';
  return props.getProperty(key) || '';
}

function normalizeSearchQuery(text) {
  if (!text) return '';
  return text.toString()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, '')
    .trim();
}

function normalizeSearchTokens(text) {
  if (!text) return [];
  var rawTokens = text.toString().toUpperCase().split(/[^A-Z0-9]+/);
  var tokens = [];
  for (var i = 0; i < rawTokens.length; i++) {
    var token = normalizeSearchQuery(rawTokens[i]);
    if (token) tokens.push(token);
  }
  return tokens;
}


function normalizeColor(color) {
  if (!color) return '';
  return color.toString()
    .toUpperCase()
    .replace(/[\s\-_]+/g, '')
    .trim();
}

function formatYm(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getFullYear().toString() + padZero(value.getMonth() + 1);
  }
  return value.toString().trim();
}

function selectPrimaryBarcode(colors) {
  if (!colors || colors.length === 0) return '';

  var filtered = colors.filter(function(c) {
    return c.status === CONFIG.STATUS.ACTIVE || c.status === CONFIG.STATUS.NEW;
  });
  if (filtered.length === 0) {
    filtered = colors.filter(function(c) {
      return c.status === CONFIG.STATUS.DISCONTINUE_UNKNOWN;
    });
  }
  if (filtered.length === 0) {
    filtered = colors.slice();
  }

  var priorityColors = ['1B', '1', '2', '4', '27', '30', 'NA', 'NATURALBLACK'];
  for (var i = 0; i < priorityColors.length; i++) {
    var priority = normalizeColor(priorityColors[i]);
    var matches = filtered.filter(function(c) {
      return normalizeColor(c.color) === priority;
    });
    if (matches.length > 0) {
      var best = selectBestFromMatches(matches);
      return best ? best.barcode : '';
    }
  }

  var fallback = selectBestFromMatches(filtered);
  return fallback ? fallback.barcode : '';
}

function selectBestFromMatches(colors) {
  if (!colors || colors.length === 0) return null;

  var maxShipped = 0;
  for (var i = 0; i < colors.length; i++) {
    var shipped = colors[i].shipped12mo || 0;
    if (shipped > maxShipped) {
      maxShipped = shipped;
    }
  }

  if (maxShipped > 0) {
    colors = colors.filter(function(c) {
      return (c.shipped12mo || 0) === maxShipped;
    });
  }

  colors.sort(function(a, b) {
    var aLast = parseInt(a.lastSeenYm, 10) || 0;
    var bLast = parseInt(b.lastSeenYm, 10) || 0;
    if (bLast !== aLast) return bLast - aLast;

    var aFirst = parseInt(a.firstSeenYm, 10) || 0;
    var bFirst = parseInt(b.firstSeenYm, 10) || 0;
    if (aFirst !== bFirst) return aFirst - bFirst;

    return (a.barcode || '').toString().localeCompare((b.barcode || '').toString());
  });

  return colors[0];
}

function rebuildSearchIndex(companyKey, uploadDate) {
  var company = CONFIG.COMPANIES[companyKey];
  if (!company) {
    throw new Error('Unknown company: ' + companyKey);
  }

  var activeSheet = getDbSheet(company.dbSheet);
  var data = activeSheet.getDataRange().getValues();
  if (data.length < 2) {
    return;
  }

  var headers = data[0];
  var colMap = getColumnMap(headers);
  var itemNumberCol = colMap[company.columns.ITEM_NUMBER];
  var itemNameCol = colMap[company.columns.ITEM_NAME];
  var colorCol = colMap[company.columns.COLOR];
  var barcodeCol = colMap[company.columns.BARCODE];
  var statusCol = colMap['status'];
  var firstSeenCol = colMap['first_seen_date'];
  var lastSeenCol = colMap['last_seen_date'];

  if (itemNumberCol === undefined || itemNameCol === undefined || colorCol === undefined || barcodeCol === undefined) {
    throw new Error('Active DB columns missing for ' + companyKey);
  }

  var metricsMap = buildInvoiceMetricsMap(companyKey);
  var rows = data.slice(1);

  var descriptionMap = {};
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var description = row[itemNameCol];
    var barcode = row[barcodeCol];
    if (!description || !barcode) continue;

    var descriptionKey = description.toString().trim();
    if (!descriptionMap[descriptionKey]) {
      descriptionMap[descriptionKey] = {
        itemNumbers: {},
        colors: []
      };
    }

    var itemNumber = itemNumberCol !== undefined ? row[itemNumberCol] : '';
    if (itemNumber) {
      descriptionMap[descriptionKey].itemNumbers[itemNumber.toString().trim()] = true;
    }

    var upcKey = normalizeInvoiceUpc(barcode);
    var metrics = metricsMap[upcKey];

    descriptionMap[descriptionKey].colors.push({
      barcode: barcode,
      color: colorCol !== undefined ? row[colorCol] : '',
      itemNumber: itemNumber,
      status: statusCol !== undefined ? (row[statusCol] || CONFIG.STATUS.ACTIVE) : CONFIG.STATUS.ACTIVE,
      lastSeenYm: lastSeenCol !== undefined ? formatYm(row[lastSeenCol]) : '',
      firstSeenYm: firstSeenCol !== undefined ? formatYm(row[firstSeenCol]) : '',
      shipped12mo: metrics ? metrics.shipped12mo : 0
    });
  }

  var indexRows = [];
  for (var desc in descriptionMap) {
    if (!descriptionMap.hasOwnProperty(desc)) continue;

    var entry = descriptionMap[desc];
    var colors = entry.colors;
    if (!colors || colors.length === 0) continue;

    var primaryBarcode = selectPrimaryBarcode(colors);
    if (!primaryBarcode) continue;

    var primaryColor = null;
    for (var c = 0; c < colors.length; c++) {
      if (colors[c].barcode.toString() === primaryBarcode.toString()) {
        primaryColor = colors[c];
        break;
      }
    }
    if (!primaryColor) {
      primaryColor = colors[0];
    }

    var colorCounts = {
      total: colors.length,
      active: 0,
      discontinueUnknown: 0,
      discontinued: 0
    };

    for (var j = 0; j < colors.length; j++) {
      var status = colors[j].status || CONFIG.STATUS.ACTIVE;
      if (status === CONFIG.STATUS.NEW || status === CONFIG.STATUS.ACTIVE) {
        colorCounts.active++;
      } else if (status === CONFIG.STATUS.DISCONTINUE_UNKNOWN) {
        colorCounts.discontinueUnknown++;
      } else if (status === CONFIG.STATUS.DISCONTINUED) {
        colorCounts.discontinued++;
      } else {
        colorCounts.active++;
      }
    }

    var dominantStatus = colorCounts.active > 0 ? CONFIG.STATUS.ACTIVE :
      (colorCounts.discontinueUnknown > 0 ? CONFIG.STATUS.DISCONTINUE_UNKNOWN : CONFIG.STATUS.DISCONTINUED);

    var itemNumbers = Object.keys(entry.itemNumbers).join(' ');

    indexRows.push([
      desc,
      itemNumbers,
      primaryBarcode,
      colorCounts.total,
      colorCounts.active,
      colorCounts.discontinueUnknown,
      colorCounts.discontinued,
      dominantStatus,
      primaryColor.shipped12mo || 0
    ]);
  }

  var indexSheet = getOrCreateSearchIndexSheet(companyKey);
  indexSheet.clearContents();
  indexSheet.getRange(1, 1, 1, 9).setValues([[
    'description',
    'itemNumber',
    'primaryBarcode',
    'colorCounts_total',
    'colorCounts_active',
    'colorCounts_discontinueUnknown',
    'colorCounts_discontinued',
    'dominantStatus',
    'primaryShipped12mo'
  ]]);

  if (indexRows.length > 0) {
    indexSheet.getRange(2, 1, indexRows.length, 9).setValues(indexRows);
  }

  if (uploadDate) {
    setSearchIndexDate(companyKey, uploadDate);
  }
}

function searchByKeyword(query, lastSelectedCompany) {
  try {
    if (!query || query.toString().trim() === '') {
      return { success: false, error: 'Please enter a search term.' };
    }

    var normalizedQuery = normalizeSearchQuery(query);
    if (normalizedQuery.length < 2) {
      return { success: false, error: 'Please enter at least 2 characters.' };
    }
    var normalizedTokens = normalizeSearchTokens(query);

    var results = {
      OUTRE: searchInIndex('OUTRE', normalizedQuery, normalizedTokens),
      SNG: searchInIndex('SNG', normalizedQuery, normalizedTokens)
    };

    var meta = {
      totalCount: results.OUTRE.length + results.SNG.length,
      outreCount: results.OUTRE.length,
      sngCount: results.SNG.length,
      indexDate: getSearchIndexDate('OUTRE') || getSearchIndexDate('SNG'),
      lastSelectedCompany: lastSelectedCompany || ''
    };

    if (meta.totalCount === 0) {
      meta.suggestions = generateKeywordSuggestions(query);
      meta.tips = [
        'You can also search by item number.',
        'Partial matching is supported.'
      ];
    } else {
      meta.suggestions = [];
      meta.tips = [];
    }

    return {
      success: true,
      query: query,
      results: results,
      meta: meta
    };
  } catch (error) {
    debugLog('searchByKeyword 오류', { error: error.toString(), query: query });
    logError('searchByKeyword', error.toString(), { query: query });
    return { success: false, error: 'An error occurred while searching.' };
  }
}

function searchInIndex(companyKey, normalizedQuery, normalizedTokens) {
  var sheet;
  try {
    sheet = getSearchIndexSheet(companyKey);
  } catch (error) {
    return [];
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var rows = data.slice(1);
  var matches = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var description = row[0];
    var itemNumber = row[1];

    var normalizedDesc = normalizeSearchQuery(description);
    var normalizedItem = normalizeSearchQuery(itemNumber);

    var normalizedCombined = normalizedDesc + ' ' + normalizedItem;

    var exactMatch = normalizedDesc.indexOf(normalizedQuery) > -1 ||
      normalizedItem.indexOf(normalizedQuery) > -1;

    var tokenMatch = true;
    if (normalizedTokens && normalizedTokens.length > 0) {
      for (var t = 0; t < normalizedTokens.length; t++) {
        if (normalizedCombined.indexOf(normalizedTokens[t]) === -1) {
          tokenMatch = false;
          break;
        }
      }
    }

    if (exactMatch || tokenMatch) {
      matches.push({
        description: description,
        itemNumber: itemNumber,
        company: companyKey,
        primaryBarcode: row[2],
        colorCounts: {
          total: row[3],
          active: row[4],
          discontinueUnknown: row[5],
          discontinued: row[6]
        },
        dominantStatus: row[7],
        primaryShipped12mo: row[8]
      });
    }
  }

  matches.sort(function(a, b) {
    return (b.primaryShipped12mo || 0) - (a.primaryShipped12mo || 0);
  });

  return matches.slice(0, 100);
}

function generateKeywordSuggestions(query) {
  var normalized = normalizeSearchQuery(query);
  if (normalized.length < 2) return [];

  var suggestions = {};
  var companies = ['OUTRE', 'SNG'];

  for (var i = 0; i < companies.length; i++) {
    var sheet;
    try {
      sheet = getSearchIndexSheet(companies[i]);
    } catch (error) {
      continue;
    }

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    for (var j = 1; j < data.length; j++) {
      var desc = data[j][0];
      if (!desc) continue;
      var normalizedDesc = normalizeSearchQuery(desc);
      if (normalizedDesc.indexOf(normalized.slice(0, 2)) > -1) {
        suggestions[desc] = true;
        if (Object.keys(suggestions).length >= 5) {
          return Object.keys(suggestions);
        }
      }
    }
  }

  return Object.keys(suggestions);
}

function getQuantityViewByDescription(description, companyKey, primaryBarcode) {
  try {
    if (!description || !companyKey) {
      return { success: false, error: 'No results found.' };
    }

    var products = buildProductsByDescription(description, companyKey);
    if (!products || products.length === 0) {
      if (primaryBarcode) {
        return searchBarcode(primaryBarcode);
      }
      return { success: false, error: 'No results found.' };
    }

    products = sortProducts(products);
    products = attachInvoiceMetricsToProducts(products, companyKey);
    var result = calculateColorStatistics(products);

    return buildSearchResponse(primaryBarcode || '', companyKey, result);
  } catch (error) {
    debugLog('getQuantityViewByDescription 오류', { error: error.toString() });
    logError('getQuantityViewByDescription', error.toString(), { description: description, company: companyKey });
    return { success: false, error: 'An error occurred while searching.' };
  }
}

function buildProductsByDescription(description, companyKey) {
  var company = CONFIG.COMPANIES[companyKey];
  if (!company) return [];

  var target = description.toString().trim();
  var sheet = getDbSheet(company.dbSheet);
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  var headers = data[0];
  var colMap = getColumnMap(headers);
  var itemNumberCol = colMap[company.columns.ITEM_NUMBER];
  var itemNameCol = colMap[company.columns.ITEM_NAME];
  var colorCol = colMap[company.columns.COLOR];
  var barcodeCol = colMap[company.columns.BARCODE];
  var statusCol = colMap['status'];

  if (itemNameCol === undefined || barcodeCol === undefined) return [];

  var products = [];
  for (var i = 1; i < data.length; i++) {
    var itemName = data[i][itemNameCol];
    if (!itemName || itemName.toString().trim() !== target) continue;

    var barcode = data[i][barcodeCol];
    if (!barcode) continue;

    products.push({
      itemNumber: itemNumberCol !== undefined ? data[i][itemNumberCol] : '',
      itemName: itemName,
      color: colorCol !== undefined ? data[i][colorCol] : '',
      barcode: barcode,
      quantity: 0,
      company: companyKey,
      status: statusCol !== undefined ? (data[i][statusCol] || CONFIG.STATUS.ACTIVE) : CONFIG.STATUS.ACTIVE
    });
  }

  return products;
}

// ============================================================================
// History & Favorites (Server Backup)
// ============================================================================

function getUserKey() {
  var email = '';
  try {
    email = Session.getActiveUser().getEmail();
  } catch (e) {
    email = '';
  }
  return email || 'anonymous';
}

function getOrCreateHistorySheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.HISTORY_SHEET);
    sheet.appendRow(['email', 'timestamp', 'type', 'description', 'company', 'primaryBarcode']);
    sheet.hideSheet();
  }
  return sheet;
}

function getOrCreateFavoritesSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.FAVORITES_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.FAVORITES_SHEET);
    sheet.appendRow(['email', 'timestamp', 'description', 'company', 'primaryBarcode']);
    sheet.hideSheet();
  }
  return sheet;
}

function saveHistoryToServer(item) {
  if (!item || !item.description || !item.company || !item.primaryBarcode) return;

  var sheet = getOrCreateHistorySheet();
  var email = getUserKey();
  var data = sheet.getDataRange().getValues();

  var userRows = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      userRows.push({ rowIndex: i + 1, timestamp: data[i][1] });
    }
  }

  if (userRows.length >= 10) {
    userRows.sort(function(a, b) {
      return new Date(a.timestamp).getTime() - new Date(b.timestamp).getTime();
    });
    sheet.deleteRow(userRows[0].rowIndex);
  }

  sheet.appendRow([email, new Date(), item.type || 'search', item.description, item.company, item.primaryBarcode]);
}

function getHistoryFromServer() {
  var sheet = getOrCreateHistorySheet();
  var email = getUserKey();
  var data = sheet.getDataRange().getValues();
  var items = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== email) continue;
    items.push({
      type: data[i][2],
      description: data[i][3],
      company: data[i][4],
      primaryBarcode: data[i][5],
      timestamp: data[i][1]
    });
  }

  items.sort(function(a, b) {
    return new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime();
  });

  return items.slice(0, 10);
}

function saveFavoriteToServer(item, action) {
  if (!item || !item.description || !item.company || !item.primaryBarcode) return;

  var sheet = getOrCreateFavoritesSheet();
  var email = getUserKey();
  var data = sheet.getDataRange().getValues();

  var existingRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === email &&
        data[i][2] === item.description &&
        data[i][3] === item.company &&
        data[i][4] === item.primaryBarcode) {
      existingRow = i + 1;
      break;
    }
  }

  if (action === 'remove') {
    if (existingRow > -1) {
      sheet.deleteRow(existingRow);
    }
    return;
  }

  if (existingRow === -1) {
    sheet.appendRow([email, new Date(), item.description, item.company, item.primaryBarcode]);
  }
}

function getFavoritesFromServer() {
  var sheet = getOrCreateFavoritesSheet();
  var email = getUserKey();
  var data = sheet.getDataRange().getValues();
  var items = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] !== email) continue;
    items.push({
      description: data[i][2],
      company: data[i][3],
      primaryBarcode: data[i][4],
      timestamp: data[i][1]
    });
  }

  items.sort(function(a, b) {
    return new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime();
  });

  return items.slice(0, 20);
}
