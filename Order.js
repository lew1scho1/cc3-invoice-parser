// ============================================================================
// ORDER.GS - 주문 관련 함수
// ============================================================================

/**
 * ORDER 시트에 제품 추가 (회사별)
 */
function addToOrder(products, companyKey) {
  try {
    debugLog('addToOrder 시작', { productCount: products.length, company: companyKey });
    
    if (!products || !Array.isArray(products)) {
      return {
        success: false,
        error: '❌ 올바른 제품 데이터가 아닙니다.'
      };
    }
    
    if (!companyKey || !CONFIG.COMPANIES[companyKey]) {
      return {
        success: false,
        error: '❌ 올바른 회사 정보가 아닙니다.'
      };
    }
    
    var company = CONFIG.COMPANIES[companyKey];
    
    var validProducts = [];
    for (var i = 0; i < products.length; i++) {
      var qty = parseInt(products[i].quantity);
      if (!isNaN(qty) && qty > 0) {
        validProducts.push({
          itemName: products[i].itemName || '',
          color: products[i].color || '',
          quantity: qty,
          barcode: products[i].barcode || ''
        });
      }
    }
    
    debugLog('유효한 제품', { count: validProducts.length });
    
    if (validProducts.length === 0) {
      return {
        success: false,
        error: '⚠️ 수량이 입력된 제품이 없습니다.'
      };
    }
    
    var sheet = getSheet(company.orderSheet);
    
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      sheet.appendRow(['ITEM NAME', 'COLOR', '수량', 'UPC']);
      debugLog('ORDER 시트 헤더 생성');
    }
    
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (headerRow.indexOf('UPC') === -1) {
      headerRow.push('UPC');
      sheet.getRange(1, headerRow.length).setValue('UPC');
    }
    
    var headerMap = getColumnMap(headerRow);
    var itemCol = headerMap['ITEM NAME'];
    var colorCol = headerMap['COLOR'];
    var qtyCol = headerMap['수량'];
    if (qtyCol === undefined) {
      qtyCol = headerMap['?˜ëŸ‰'];
    }
    var upcCol = headerMap['UPC'];
    
    var addedCount = 0;
    for (var i = 0; i < validProducts.length; i++) {
      try {
        var row = new Array(headerRow.length);
        row[itemCol] = validProducts[i].itemName;
        row[colorCol] = validProducts[i].color;
        row[qtyCol] = validProducts[i].quantity;
        row[upcCol] = validProducts[i].barcode;
        sheet.appendRow(row);
        addedCount++;
      } catch (rowError) {
        debugLog('행 추가 실패', { error: rowError.toString() });
      }
    }
    
    if (addedCount === 0) {
      return {
        success: false,
        error: '❌ 주문 추가에 실패했습니다.'
      };
    }
    
    debugLog('주문 추가 완료', { addedCount: addedCount });
    
    return {
      success: true,
      message: '✅ ' + addedCount + '개 제품이 주문에 추가되었습니다.',
      addedCount: addedCount
    };
    
  } catch (error) {
    debugLog('addToOrder 오류', { error: error.toString() });
    logError('addToOrder', error.toString(), { 
      productCount: products ? products.length : 0,
      company: companyKey 
    });
    
    return {
      success: false,
      error: '❌ 주문 추가 중 오류가 발생했습니다.\n' + error.toString()
    };
  }
}

/**
 * ORDER 시트 데이터 가져오기 (회사별)
 */
function getOrderData(companyKey) {
  try {
    debugLog('getOrderData 시작', { company: companyKey });
    
    if (!companyKey || !CONFIG.COMPANIES[companyKey]) {
      return {
        success: false,
        error: '❌ 올바른 회사 정보가 아닙니다.'
      };
    }
    
    var company = CONFIG.COMPANIES[companyKey];
    var sheet = getSheet(company.orderSheet);
    var data = sheet.getDataRange().getValues();
    
    debugLog('ORDER 데이터 로드', { rows: data.length });
    
    if (data.length === 0) {
      return {
        success: true,
        data: [],
        message: '주문 내역이 없습니다.'
      };
    }
    
    var headers = data[0];
    var rows = data.slice(1);
    
    var colMap = getColumnMap(headers);
    var itemNameIdx = colMap['ITEM NAME'];
    var colorIdx = colMap['COLOR'];
    var upcIdx = colMap['UPC'];
    var qtyIdx = colMap['수량'];
    if (qtyIdx === undefined) {
      qtyIdx = colMap['?˜ëŸ‰'];
    }
    
    if (itemNameIdx === undefined || colorIdx === undefined) {
      return {
        success: true,
        data: data
      };
    }
    
    var products = [];
    for (var i = 0; i < rows.length; i++) {
      products.push({
        itemName: rows[i][itemNameIdx] || '',
        color: rows[i][colorIdx] || '',
        quantity: rows[i][qtyIdx] || '',
        rowData: rows[i]
      });
    }
    
    products = sortProducts(products);
    
    var displayHeaders = headers;
    if (upcIdx !== undefined) {
      displayHeaders = headers.filter(function(_, index) {
        return index !== upcIdx;
      });
    }
    
    var sortedData = [displayHeaders];
    for (var i = 0; i < products.length; i++) {
      var rowData = products[i].rowData;
      if (upcIdx !== undefined) {
        rowData = rowData.filter(function(_, index) {
          return index !== upcIdx;
        });
      }
      sortedData.push(rowData);
    }
    
    debugLog('getOrderData 완료', { sortedRows: sortedData.length });
    
    return {
      success: true,
      data: sortedData
    };
    
  } catch (error) {
    debugLog('getOrderData 오류', { error: error.toString() });
    logError('getOrderData', error.toString(), { company: companyKey });
    
    return {
      success: false,
      error: '❌ 주문 데이터 조회 중 오류가 발생했습니다.\n' + error.toString()
    };
  }
}

/**
 * ORDER 리스트 비우기 (백업 후) - 회사별
 */
function clearOrderList(companyKey) {
  try {
    debugLog('clearOrderList 시작', { company: companyKey });
    
    if (!companyKey || !CONFIG.COMPANIES[companyKey]) {
      return {
        success: false,
        error: '❌ 올바른 회사 정보가 아닙니다.'
      };
    }
    
    var company = CONFIG.COMPANIES[companyKey];
    var ss = getSpreadsheet();
    var orderSheet = ss.getSheetByName(company.orderSheet);
    
    if (!orderSheet) {
      return {
        success: false,
        error: '❌ ' + company.orderSheet + ' 시트를 찾을 수 없습니다.'
      };
    }
    
    var data = orderSheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return {
        success: true,
        message: '비울 데이터가 없습니다.'
      };
    }
    
    var now = new Date();
    var year = now.getFullYear();
    var month = padZero(now.getMonth() + 1);
    var day = padZero(now.getDate());
    var hour = padZero(now.getHours());
    var minute = padZero(now.getMinutes());
    var second = padZero(now.getSeconds());
    
    var backupName = company.name + '_' + year + '-' + month + '-' + day + '_' + hour + '-' + minute + '-' + second;
    
    debugLog('백업 탭 생성', { name: backupName });
    
    var backupSheet = ss.insertSheet(backupName);
    
    var range = orderSheet.getDataRange();
    var values = range.getValues();
    var formats = range.getTextStyles();
    
    if (values.length > 0) {
      backupSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
      if (formats && formats.length > 0) {
        backupSheet.getRange(1, 1, formats.length, formats[0].length).setTextStyles(formats);
      }
    }
    
    backupSheet.hideSheet();
    
    debugLog('백업 완료');
    
    if (orderSheet.getLastRow() > 1) {
      orderSheet.deleteRows(2, orderSheet.getLastRow() - 1);
    }
    
    debugLog('ORDER 시트 비우기 완료');
    
    return {
      success: true,
      message: '✅ 주문 내역이 "' + backupName + '" 탭에 백업되고 비워졌습니다.',
      backupName: backupName
    };
    
  } catch (error) {
    debugLog('clearOrderList 오류', { error: error.toString() });
    logError('clearOrderList', error.toString(), { company: companyKey });
    
    return {
      success: false,
      error: '❌ 리스트 비우기 중 오류가 발생했습니다.\n' + error.toString()
    };
  }
}
