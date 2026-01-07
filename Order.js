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

    function extractLatestPrice(value) {
      if (value === null || value === undefined) return null;
      var str = value.toString();
      var matches = str.match(/\$?[\d,]+(?:\.\d{1,2})?/g);
      if (!matches || matches.length === 0) return null;
      var last = matches[matches.length - 1];
      var num = parseFloat(last.replace(/[^0-9.]/g, ''));
      return isNaN(num) ? null : num;
    }
    
    var validProducts = [];
    for (var i = 0; i < products.length; i++) {
      // ========================================
      // CRITICAL: 그룹 헤더(구분자) 제외
      // ========================================
      if (products[i].isGroupHeader) {
        debugLog('그룹 헤더 제외', { itemName: products[i].itemName });
        continue;
      }

      var qty = parseInt(products[i].quantity);
      if (!isNaN(qty) && qty > 0) {
        var unitPrice = extractLatestPrice(products[i].priceHistory);
        var extPrice = unitPrice !== null ? unitPrice * qty : null;
        validProducts.push({
          itemName: products[i].itemName || '',
          color: products[i].color || '',
          quantity: qty,
          barcode: products[i].barcode || '',
          itemNumber: products[i].itemNumber || '',
          unitPrice: unitPrice,
          extPrice: extPrice
        });
      }
    }
    
    debugLog('유효한 제품', { count: validProducts.length });

    // ========================================
    // CRITICAL: 수량 0인 제품만 있어도 처리 진행
    // (기존 장바구니에서 삭제해야 하므로)
    // ========================================
    var hasAnyProducts = false;
    for (var i = 0; i < products.length; i++) {
      if (!products[i].isGroupHeader && products[i].barcode) {
        hasAnyProducts = true;
        break;
      }
    }

    if (!hasAnyProducts) {
      return {
        success: false,
        error: '⚠️ 처리할 제품이 없습니다.'
      };
    }

    var sheet = getSheet(company.orderSheet);
    
    var lastRow = sheet.getLastRow();
    if (lastRow === 0) {
      var initialHeaders = ['ITEM NAME', 'COLOR', '수량'];
      if (companyKey === 'SNG') {
        initialHeaders.push('ITEM CODE', 'UPC');
      } else {
        initialHeaders.push('UPC');
      }
      sheet.appendRow(initialHeaders);
      debugLog('ORDER 시트 헤더 생성');
    }

    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Ensure necessary columns exist
    if (companyKey === 'SNG') {
      if (headerRow.indexOf('ITEM CODE') === -1) {
        sheet.getRange(1, headerRow.length + 1).setValue('ITEM CODE');
        headerRow.push('ITEM CODE');
      }
    }
    if (headerRow.indexOf('UPC') === -1) {
      sheet.getRange(1, headerRow.length + 1).setValue('UPC');
      headerRow.push('UPC');
    }
    if (headerRow.indexOf('UNIT PRICE') === -1) {
      sheet.getRange(1, headerRow.length + 1).setValue('UNIT PRICE');
      headerRow.push('UNIT PRICE');
    }
    if (headerRow.indexOf('EXT PRICE') === -1) {
      sheet.getRange(1, headerRow.length + 1).setValue('EXT PRICE');
      headerRow.push('EXT PRICE');
    }
    
    var headerMap = getColumnMap(headerRow);
    var itemCol = headerMap['ITEM NAME'];
    var colorCol = headerMap['COLOR'];
    var qtyCol = headerMap['수량'];
    if (qtyCol === undefined) {
      qtyCol = headerMap['?˜ëŸ‰'];
    }
    var upcCol = headerMap['UPC'];
    var itemCodeCol = companyKey === 'SNG' ? headerMap['ITEM CODE'] : undefined;
    var unitPriceCol = headerMap['UNIT PRICE'];
    var extPriceCol = headerMap['EXT PRICE'];

    // ========================================
    // CRITICAL: 모든 제품(수량 0 포함)의 바코드 수집
    // 수량 0인 제품도 기존 장바구니에서 삭제해야 함
    // ========================================
    var allBarcodeSet = {};
    for (var i = 0; i < products.length; i++) {
      if (products[i].isGroupHeader) continue;

      var rawUpc = products[i].barcode;
      var normalizedUpc = normalizeInvoiceUpc(rawUpc);
      if (normalizedUpc) {
        allBarcodeSet[normalizedUpc] = true;
      }
    }

    // De-duplication: 기존 장바구니에서 현재 제품들의 바코드 모두 삭제
    if (upcCol !== undefined && Object.keys(allBarcodeSet).length > 0) {
      var lastDataRow = sheet.getLastRow();
      if (lastDataRow > 1) {
        var existingData = sheet.getRange(2, 1, lastDataRow - 1, sheet.getLastColumn()).getValues();
        for (var r = existingData.length - 1; r >= 0; r--) {
          var existingUpc = normalizeInvoiceUpc(existingData[r][upcCol]);
          if (existingUpc && allBarcodeSet[existingUpc]) {
            debugLog('기존 장바구니 행 삭제', { barcode: existingUpc });
            sheet.deleteRow(r + 2);
          }
        }
      }
    }
    
    var addedCount = 0;
    for (var i = 0; i < validProducts.length; i++) {
      try {
        var row = new Array(headerRow.length);
        row[itemCol] = validProducts[i].itemName;
        row[colorCol] = validProducts[i].color;
        row[qtyCol] = validProducts[i].quantity;
        
        if (companyKey === 'SNG') {
          if (itemCodeCol !== undefined) {
            row[itemCodeCol] = validProducts[i].itemNumber;
          }
        }
        if (upcCol !== undefined) {
          row[upcCol] = validProducts[i].barcode;
        }

        if (unitPriceCol !== undefined) {
          row[unitPriceCol] = validProducts[i].unitPrice !== null ?
            validProducts[i].unitPrice.toFixed(2) : '';
        }
        if (extPriceCol !== undefined) {
          row[extPriceCol] = validProducts[i].extPrice !== null ?
            validProducts[i].extPrice.toFixed(2) : '';
        }
        sheet.appendRow(row);
        addedCount++;
      } catch (rowError) {
        debugLog('행 추가 실패', { error: rowError.toString() });
      }
    }
    
    debugLog('주문 처리 완료', { addedCount: addedCount, validProducts: validProducts.length });

    // ========================================
    // 수량 0인 제품만 처리한 경우도 성공으로 처리
    // ========================================
    if (addedCount === 0) {
      return {
        success: true,
        message: '✅ 장바구니가 업데이트되었습니다.',
        addedCount: 0
      };
    }

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
function removeOrderItem(companyKey, upc, color) {
  try {
    debugLog('removeOrderItem ?œìž‘', { company: companyKey, upc: upc });

    if (!companyKey || !CONFIG.COMPANIES[companyKey]) {
      return {
        success: false,
        error: '???¬ë°”ë¥??Œì‚¬ ?•ë³´ê°€ ?„ë‹™?ˆë‹¤.'
      };
    }

    var normalizedUpc = normalizeInvoiceUpc(upc);
    if (!normalizedUpc) {
      return {
        success: false,
        error: 'UPCê°€ ìœ íš¨í•˜ì§€ ì•ŠìŠµ?ˆë‹¤.'
      };
    }

    var company = CONFIG.COMPANIES[companyKey];
    var sheet = getSheet(company.orderSheet);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, removed: 0 };
    }

    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var headerMap = getColumnMap(headerRow);
    var upcCol = headerMap['UPC'];
    if (upcCol === undefined) {
      return { success: false, error: 'UPC 열을 찾을 수 없습니다.' };
    }
    var colorCol = headerMap['COLOR'];

    var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    var removed = 0;
    var normalizedColor = color ? normalizeColor(color) : '';
    for (var i = data.length - 1; i >= 0; i--) {
      var rowUpc = normalizeInvoiceUpc(data[i][upcCol]);
      if (rowUpc !== normalizedUpc) continue;

      var rowColor = normalizeColor(data[i][colorCol]);
      if (rowColor !== normalizedColor) continue;
      
      sheet.deleteRow(i + 2);
      removed++;
      break; 
    }

    return {
      success: true,
      removed: removed
    };
  } catch (error) {
    debugLog('removeOrderItem ?¤ë¥˜', { error: error.toString() });
    logError('removeOrderItem', error.toString(), { company: companyKey, upc: upc });
    return {
      success: false,
      error: 'ì‚­ì œ ì¤??¤ë¥˜ê°€ ë°œìƒ?ˆìŠµ?ˆë‹¤.\n' + error.toString()
    };
  }
}

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

    // UPC 컬럼을 제거하지 않고 그대로 유지 (프론트엔드 가격 계산을 위해)
    var sortedData = [headers];
    var upcList = [];
    for (var i = 0; i < products.length; i++) {
      var rowData = products[i].rowData;
      var rowUpc = upcIdx !== undefined ? rowData[upcIdx] : '';
      upcList.push(rowUpc || '');
      sortedData.push(rowData);
    }
    
    debugLog('getOrderData 완료', { sortedRows: sortedData.length });
    
    return {
      success: true,
      data: sortedData,
      upcList: upcList
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
