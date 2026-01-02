// ============================================================================
// FillOUTREOrderUPC.js - OUTRE 오더 시트에 UPC 자동 입력
// ============================================================================
//
// 사용법:
// 1. Google Sheets에서 "outre 251230" 탭 생성
// 2. A열: ITEM NAME, B열: COLOR, C열: 수량 입력
// 3. fillOUTREOrderUPC() 함수 실행
// 4. D열에 UPC 자동 입력됨
//
// ============================================================================

/**
 * OUTRE 오더 시트에 UPC 입력
 * - A열: ITEM NAME (Description)
 * - B열: COLOR
 * - C열: 수량
 * - D열: UPC (자동 입력)
 */
function fillOUTREOrderUPC() {
  try {
    Logger.log('=== OUTRE 오더 UPC 입력 시작 ===');

    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var orderSheet = ss.getSheetByName('outre 251230');

    if (!orderSheet) {
      throw new Error('❌ "outre 251230" 시트를 찾을 수 없습니다.');
    }

    // 오더 데이터 읽기
    var orderData = orderSheet.getDataRange().getValues();

    if (orderData.length < 2) {
      throw new Error('❌ 오더 데이터가 없습니다.');
    }

    Logger.log('오더 데이터 로드: ' + (orderData.length - 1) + '개 라인');

    // OUTRE DB 캐시 초기화
    initOUTREDBCache();

    if (OUTRE_DB_CACHE.error) {
      throw new Error('❌ OUTRE DB 캐시 초기화 실패');
    }

    var dbMap = OUTRE_DB_CACHE.dbMap;

    Logger.log('OUTRE DB 캐시 로드: ' + Object.keys(dbMap).length + '개 Description');

    // 정규화 함수 (enrichOUTREUPC()와 동일)
    var normalize = function(text) {
      if (!text) return '';
      return text.toString()
        .trim()
        .replace(/["″''`]/g, '"')
        .replace(/\s+/g, ' ')
        .replace(/\-+/g, '-')
        .replace(/\s*-\s*/g, '-')
        .replace(/\s*\/\s*/g, '/')  // 슬래시 앞뒤 공백 제거
        .toUpperCase();
    };

    // 결과 배열 (D열에 입력할 UPC)
    var upcResults = [];
    var stats = {
      total: 0,
      matched: 0,
      descNotFound: 0,
      colorNotFound: 0
    };

    // 헤더 행 건너뛰기 (row 1)
    upcResults.push(['UPC']); // 헤더

    // 각 라인 처리
    for (var i = 1; i < orderData.length; i++) {
      var description = orderData[i][0]; // A열: ITEM NAME
      var color = orderData[i][1];       // B열: COLOR
      var quantity = orderData[i][2];    // C열: 수량

      stats.total++;

      Logger.log('');
      Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      Logger.log('[' + i + '] ' + description);
      Logger.log('  Color: ' + color);
      Logger.log('  수량: ' + quantity);

      var upc = '';

      if (!description || !color) {
        Logger.log('  ❌ Description 또는 Color 없음');
        upcResults.push(['']);
        continue;
      }

      // Step 1: Description 정규화 및 DB 조회
      var normalizedDesc = normalize(description);
      var matchedRecords = dbMap[normalizedDesc];

      if (!matchedRecords || matchedRecords.length === 0) {
        Logger.log('  ❌ Description DB 미매칭');
        Logger.log('    정규화된 Description: ' + normalizedDesc.substring(0, 60));
        upcResults.push(['⚠️ DB 미등록 제품']);
        stats.descNotFound++;
        continue;
      }

      Logger.log('  ✅ Description 매칭: ' + matchedRecords.length + '개 레코드');

      // Step 2: Color 매칭
      var normalizedColor = normalize(color);
      var found = false;

      for (var j = 0; j < matchedRecords.length; j++) {
        var dbColor = normalize(matchedRecords[j].color);

        if (dbColor === normalizedColor) {
          upc = matchedRecords[j].barcode || '';
          found = true;
          Logger.log('  ✅ Color 매칭 성공');
          Logger.log('    DB Color: ' + matchedRecords[j].color);
          Logger.log('    UPC: ' + upc);
          break;
        }
      }

      if (!found) {
        Logger.log('  ❌ Color 미매칭');
        Logger.log('    요청 Color: ' + color);
        Logger.log('    정규화: ' + normalizedColor);
        Logger.log('    DB에 있는 Color 목록:');
        for (var j = 0; j < Math.min(matchedRecords.length, 5); j++) {
          Logger.log('      - ' + matchedRecords[j].color + ' (정규화: ' + normalize(matchedRecords[j].color) + ')');
        }
        upcResults.push(['⚠️ DB 미등록 컬러']);
        stats.colorNotFound++;
        continue;
      }

      if (upc) {
        stats.matched++;
        upcResults.push([upc]);
      } else {
        Logger.log('  ⚠️ UPC 없음 (Barcode 컬럼 비어있음)');
        upcResults.push(['⚠️ UPC 없음']);
      }
    }

    // D열에 UPC 입력
    if (upcResults.length > 0) {
      var range = orderSheet.getRange(1, 4, upcResults.length, 1); // D열 (4번째 컬럼)
      range.setValues(upcResults);
      Logger.log('');
      Logger.log('✅ D열에 UPC 입력 완료');
    }

    // 통계 출력
    Logger.log('');
    Logger.log('========================================');
    Logger.log('📊 최종 통계');
    Logger.log('========================================');
    Logger.log('전체: ' + stats.total + '개');
    Logger.log('✅ 매칭 성공: ' + stats.matched + '개 (' + (stats.matched / stats.total * 100).toFixed(1) + '%)');
    Logger.log('❌ Description 미매칭: ' + stats.descNotFound + '개');
    Logger.log('❌ Color 미매칭: ' + stats.colorNotFound + '개');
    Logger.log('========================================');

    // 캐시 리셋
    resetOUTREDBCache();

    // 사용자 알림
    SpreadsheetApp.getUi().alert(
      '✅ UPC 입력 완료\n\n' +
      '전체: ' + stats.total + '개\n' +
      '성공: ' + stats.matched + '개\n' +
      'Description 미매칭: ' + stats.descNotFound + '개\n' +
      'Color 미매칭: ' + stats.colorNotFound + '개'
    );

  } catch (error) {
    Logger.log('❌ 오류 발생: ' + error.toString());
    Logger.log(error.stack);
    SpreadsheetApp.getUi().alert('❌ 오류: ' + error.toString());
  }
}

/**
 * OUTRE 오더 시트 템플릿 생성
 * - 헤더만 있는 빈 시트 생성
 */
function createOUTREOrderTemplate() {
  try {
    var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

    // 기존 시트 삭제 (있으면)
    var existingSheet = ss.getSheetByName('outre 251230');
    if (existingSheet) {
      ss.deleteSheet(existingSheet);
      Logger.log('기존 "outre 251230" 시트 삭제');
    }

    // 새 시트 생성
    var newSheet = ss.insertSheet('outre 251230');

    // 헤더 입력
    var headers = ['ITEM NAME', 'COLOR', '수량', 'UPC'];
    newSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // 헤더 서식
    newSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');

    // 컬럼 너비 조정
    newSheet.setColumnWidth(1, 400); // ITEM NAME
    newSheet.setColumnWidth(2, 100); // COLOR
    newSheet.setColumnWidth(3, 80);  // 수량
    newSheet.setColumnWidth(4, 150); // UPC

    // 고정 행
    newSheet.setFrozenRows(1);

    Logger.log('✅ "outre 251230" 템플릿 시트 생성 완료');
    SpreadsheetApp.getUi().alert('✅ "outre 251230" 템플릿 생성 완료\n\nA열: ITEM NAME\nB열: COLOR\nC열: 수량\n\n데이터 입력 후 fillOUTREOrderUPC() 실행');

  } catch (error) {
    Logger.log('❌ 템플릿 생성 오류: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ 오류: ' + error.toString());
  }
}

/**
 * 메뉴에 추가
 */
function addOutreMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔧 OUTRE 오더 UPC')
    .addItem('📝 템플릿 시트 생성', 'createOUTREOrderTemplate')
    .addItem('✅ UPC 자동 입력', 'fillOUTREOrderUPC')
    .addToUi();
}
// sng + outre rocks!!
