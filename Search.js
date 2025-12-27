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
      throw new Error('회사 정보를 찾을 수 없습니다: ' + companyKey);
    }
    
    var barcodeStr = barcode.toString().trim();
    
    var sheet = getSheet(company.dbSheet);
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
          company: companyKey
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
            company: companyKey
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
    
    if (colorCount <= CONFIG.MAX_COLORS_FOR_EXPANSION) {
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
              company: companyKey
            });
          }
        }
        
        debugLog('확장 검색 후', { count: products.length });
      }
    }
    
    products = sortProducts(products);
    
    debugLog('searchBarcodeInCompany 완료', { resultCount: products.length });
    
    return products;
    
  } catch (error) {
    debugLog('searchBarcodeInCompany 오류', { error: error.toString() });
    logError('searchBarcodeInCompany', error.toString(), { barcode: barcode, company: companyKey });
    return null;
  }
}

/**
 * 바코드로 제품 검색
 */
function searchBarcode(barcode) {
  try {
    debugLog('searchBarcode 시작', { barcode: barcode });
    
    if (!isValidBarcode(barcode)) {
      return {
        success: false,
        error: '❌ 올바른 12자리 바코드를 입력해주세요.'
      };
    }
    
    var barcodeStr = barcode.toString().trim();
    
    var outreProducts = searchBarcodeInCompany(barcodeStr, 'OUTRE');
    var sngProducts = searchBarcodeInCompany(barcodeStr, 'SNG');
    
    if (outreProducts && outreProducts.length > 0) {
      debugLog('OUTRE에서 발견');
      return {
        success: true,
        products: outreProducts,
        company: 'OUTRE'
      };
    }
    
    if (sngProducts && sngProducts.length > 0) {
      debugLog('SNG에서 발견');
      return {
        success: true,
        products: sngProducts,
        company: 'SNG'
      };
    }
    
    return {
      success: false,
      error: '❌ OUTRE DB와 SNG DB 모두에서 찾을 수 없습니다.'
    };
    
  } catch (error) {
    debugLog('searchBarcode 오류', { error: error.toString() });
    logError('searchBarcode', error.toString(), { barcode: barcode });
    
    return {
      success: false,
      error: '❌ 검색 중 오류가 발생했습니다.\n' + error.toString()
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
        error: '❌ 검색어를 입력해주세요.'
      };
    }
    
    var company = CONFIG.COMPANIES[companyKey];
    if (!company) {
      throw new Error('회사 정보를 찾을 수 없습니다: ' + companyKey);
    }
    
    var keywords = searchText.toString().trim().toUpperCase().split(/\s+/);
    debugLog('검색 키워드', keywords);
    
    var sheet = getSheet(company.dbSheet);
    var data = sheet.getDataRange().getValues();
    
    debugLog('시트 데이터 로드', { sheet: company.dbSheet, rows: data.length });
    
    if (data.length < 2) {
      return {
        success: false,
        error: '❌ 데이터가 없습니다.'
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
          error: '❌ 데이터 형식 오류'
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
      
      var searchTarget = [
        itemNumber ? itemNumber.toString().toUpperCase() : '',
        itemName ? itemName.toString().toUpperCase() : '',
        color ? color.toString().toUpperCase() : ''
      ].join(' ');
      
      var allKeywordsFound = true;
      for (var k = 0; k < keywords.length; k++) {
        if (searchTarget.indexOf(keywords[k]) === -1) {
          allKeywordsFound = false;
          break;
        }
      }
      
      if (allKeywordsFound) {
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
        error: '❌ "' + searchText + '"에 해당하는 제품을 찾을 수 없습니다.'
      };
    }
    
    foundProducts = sortProducts(foundProducts);
    
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
      error: '❌ 검색 중 오류가 발생했습니다.\n' + error.toString()
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
        error: '❌ 검색어를 입력해주세요.'
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
        error: '❌ "' + input + '"에 해당하는 제품을 찾을 수 없습니다.\n\nOUTRE DB와 SNG DB 모두에서 검색했습니다.'
      };
    }
    
  } catch (error) {
    debugLog('searchProduct 오류', { error: error.toString() });
    logError('searchProduct', error.toString(), { input: searchInput });
    
    return {
      success: false,
      error: '❌ 검색 중 오류가 발생했습니다.\n' + error.toString()
    };
  }
}