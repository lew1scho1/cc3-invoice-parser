# CC3 Invoice Parser

## 프로젝트 개요
Google Apps Script 기반의 인보이스 파싱 시스템으로, SNG와 OUTRE 브랜드의 PDF 인보이스를 자동으로 분석하고 Google Sheets에 정리합니다.

## 기술 스택
- **플랫폼**: Google Apps Script (V8 런타임)
- **AI 엔진**: Google Cloud Document AI (Invoice Parser)
- **스토리지**: Google Drive
- **데이터베이스**: Google Sheets
- **인증**: OAuth2 (서비스 계정)

## 주요 기능
1. Google Drive 폴더에서 PDF 인보이스 자동 감지
2. Document AI를 통한 인보이스 데이터 추출
3. Vendor별 맞춤형 파싱 (SNG, OUTRE)
4. 컬러별 수량 자동 분리 및 라인 아이템 생성
5. 데이터베이스 연동 (바코드, 아이템 정보 조회)
6. 디버그 모드 및 에러 로깅

## 파일 구조

### 핵심 파일
- **Invoice_Parser.js** (28K+ tokens)
  - 메인 파싱 로직
  - UI 인터페이스 (폴더 설정, 파일 선택)
  - 데이터 변환 및 시트 작성

- **DocumentAI.js** (1009 lines)
  - Google Cloud Document AI 통합
  - OAuth2 인증
  - 엔티티 추출 및 변환
  - Vendor별 price 필드 처리
  - PDF 레이아웃 재구성 (Excel)

- **Config.js** (82 lines)
  - 설정 및 상수 정의
  - Spreadsheet ID 및 시트 이름
  - Vendor별 컬럼 매핑
  - 브랜드 정보

### 유틸리티 파일
- **Search.js**: 데이터베이스 검색 로직
- **Utilss.js**: 공통 유틸리티 함수
- **Error_Log.js**: 에러 로깅
- **Debug.js**: 디버그 로깅
- **Order.js**: 주문 처리 관련
- **Code.js**: 기타 코드 (용도 불명)

### 설정 파일
- **appsscript.json**: Apps Script 매니페스트
  - Drive API v2 활성화
  - 웹앱 설정 (익명 접근 허용)
  - 시간대: America/Chicago

- **.clasp.json**: Clasp 배포 설정

## Vendor 정보

### SNG (Shake-N-Go)
- **Invoice Number 패턴**: 10자리 숫자 (예: 3000123456)
- **파일명 패턴**: `3000*.pdf` 또는 10자리 숫자
- **시트**: DB_SNG, ORDER_SNG, INVOICE_SNG
- **가격 필드**: "Your Price", "Your Extended"
- **특징**: Invoice Amount는 마지막 발생 위치에서 추출

### OUTRE
- **Invoice Number 패턴**: SINV로 시작 (예: SINV123456)
- **파일명 패턴**: `SINV*.pdf`
- **시트**: DB_OUTRE, ORDER_OUTRE, INVOICE_OUTRE
- **가격 필드**: "Disc Price", Amount
- **특징**: TOTAL 필드에서 총액 추출

## 데이터베이스 구조

### 공통 컬럼
- ITEM NUMBER / Item Code
- ITEM NAME / Description
- COLOR / Color
- BARCODE / Barcode

### 인보이스 시트 헤더
```
VENDOR, Invoice No, Invoice Date, Total Amount, Subtotal, Discount,
Shipping, Tax, Line No, Item ID, UPC, Description, Brand, Color,
Size/Length, Qty Ordered, Qty Shipped, Unit Price, Ext Price, Memo
```

## 워크플로우

### 1. 초기 설정
```
메뉴: CC3 ORDER APP > 📄 인보이스 > 📁 폴더 설정
```
- Google Drive 폴더 URL 입력
- 폴더 ID가 Document Properties에 저장됨

### 2. 파싱 실행
```
메뉴: CC3 ORDER APP > 📄 인보이스 > ▶️ 파싱 시작
```
1. 폴더에서 PDF 파일 목록 표시
2. 사용자가 파일 선택
3. Document AI 호출
4. PARSING 시트에 임시 결과 저장
5. 사용자 검토 후 확정

### 3. 확정
```
메뉴: CC3 ORDER APP > 📄 인보이스 > ✅ 확정
```
- PARSING 시트 → INVOICE_SNG 또는 INVOICE_OUTRE로 이동
- 누적 데이터 저장

## Document AI 설정

### 필수 Script Properties
```javascript
DOCUMENT_AI_PROJECT_ID       // GCP 프로젝트 ID
DOCUMENT_AI_LOCATION         // 리전 (예: us)
DOCUMENT_AI_PROCESSOR_ID     // Processor ID
DOCUMENT_AI_SERVICE_ACCOUNT  // 서비스 계정 JSON (전체)
```

### 엔티티 타입
- `invoice_id`: Invoice Number
- `invoice_date`: Invoice Date
- `total_amount`: 총액
- `line_item`: 라인 아이템
  - Properties: description, product_code, quantity, unit_price, amount

## 컬러 라인 파싱

### 패턴
```
[컬러명] - [shipped 수량] ([backordered 수량])
```

### 예시
```
GINGER - 12 (3)
T30 - 6
1B/30 - 24 (6)
```

### 정규식
```javascript
/([A-Z0-9\-\/]+)\s*-\s*(\d+)\s*(?:\((\d+)\))?/gi
```

## 코딩 컨벤션

### 스타일
- **언어**: JavaScript (Google Apps Script)
- **인코딩**: UTF-8 (한글 주석 사용)
- **들여쓰기**: 2 스페이스
- **따옴표**: 싱글 쿼트 (')

### 주석
```javascript
// ============================================================================
// 파일명.GS - 설명
// ============================================================================

/**
 * 함수 설명
 * @param {Type} paramName - 파라미터 설명
 * @return {Type} 반환값 설명
 */
```

### 디버그 로깅
```javascript
debugLog('메시지', { key: value });  // CONFIG.DEBUG가 true일 때만 실행
```

### 에러 처리
```javascript
try {
  // 코드
} catch (error) {
  debugLog('함수명 오류', { error: error.toString() });
  logError('함수명', error, { context: data });
  throw error;
}
```

## 개발 워크플로우

### 1. 로컬 개발
```bash
# Clasp으로 코드 pull/push
clasp pull
clasp push
```

### 2. 테스트
```javascript
// Document AI 설정 테스트
testDocumentAI()

// Document AI 응답 디버깅
debugDocumentAIResponse()

// JSON + Excel 저장 테스트
saveDocumentAIResponseToFiles()
```

### 3. 디버깅
- Apps Script 편집기: `실행 로그` 또는 `Logger` 확인
- `CONFIG.DEBUG = true` 설정 시 상세 로그 출력

## 최근 수정 사항

### 2025-12-27: OUTRE Description 검증 키워드 추가
**문제**: 일부 제품명이 유효한 Description으로 인식되지 않음
- 예: `X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED`
- `LOOKS`, `PASSION` 등의 키워드가 누락되어 QTY 검증 실패

**수정 위치**:
- [Invoice_Parser.js:1202](Invoice_Parser.js:1202) - 메인 QTY 검증
- [Invoice_Parser.js:1323](Invoice_Parser.js:1323) - 다음 아이템 감지

**수정 내용**:
```javascript
// 기존 키워드에 LOOKS, PASSION 추가
var hasProductKeywords = nextLine.match(/...FEED|LOOKS|PASSION/i);
```

**효과**: 더 많은 제품명 패턴을 정확히 인식

### 2025-12-27: OUTRE 컬러+가격 통합 라인 파싱 개선
**문제**: 컬러 정보와 가격이 한 줄에 있는 경우 파싱 실패
- 예: `1- 0 (10)   1B- 10   	4.00	3.50	35.00`
- 탭 문자로 구분된 가격을 인식하지 못함

**수정 위치**: [Invoice_Parser.js:1453](Invoice_Parser.js:1453)
```javascript
// 수정 전
var pricesInColorLine = nextLine.match(/([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$/);

// 수정 후
var pricesInColorLine = nextLine.match(/([\d,]+\.\d{2})[\s\t]+([\d,]+\.\d{2})[\s\t]+([\d,]+\.\d{2})\s*$/);
```

**효과**: 탭 또는 공백으로 구분된 가격을 모두 인식

### 2025-12-27: 메타데이터 정규식 단어 경계 추가
**문제**: 메타데이터 정규식이 제품명과 부분 일치하여 QTY 검증 실패
- 예: `X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED`
- 기존 regex가 `PASSION`을 메타데이터로 오인식 (false positive)
- `hasMetadata: true`로 인해 유효한 Description이 거부됨

**근본 원인**:
```javascript
// 기존 (문제): 부분 문자열 매칭
var hasMetadata = nextLine.match(/SHIP\s+TO|SOLD\s+TO|...|PAYMENT|TERMS/i);
// "PASSION" 내의 "PASS" 등이 일부 메타데이터 패턴과 매칭될 가능성
```

**수정 위치**:
- [Invoice_Parser.js:1203](Invoice_Parser.js:1203) - 메인 QTY 검증
- [Invoice_Parser.js:1327](Invoice_Parser.js:1327) - 다음 아이템 감지

**수정 내용**:
```javascript
// 수정 후: 단어 경계(\b) 추가로 정확한 단어 매칭만 허용
var hasMetadata = nextLine.match(/\bSHIP\s+TO\b|\bSOLD\s+TO\b|\bWEIGHT\b|\bSUBTOTAL\b|\bRICHMOND\b|\bLLC\b|\bPKWAY\b|\bCOD\b|\bFee\b|\btag\b|\bDATE\s+SHIPPED\b|\bPAGE\b|\bSHIP\s+VIA\b|\bPAYMENT\b|\bTERMS\b/i);

// 디버깅 로그 추가 (line 1220-1222)
if (hasMetadata) {
  Logger.log('    매칭된 메타데이터: ' + hasMetadata[0]);
}
```

**효과**:
- 제품명이 메타데이터로 오인식되는 것 방지
- "X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED"와 같은 제품명이 정상적으로 인식됨

### 2025-12-27: Description 블랙리스트 로직 개선
**문제**: Description이 불완전하게 파싱됨
- 예상: `X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED TWIST 10" 3X`
- 실제: `TWIST 10"` (첫 줄과 3X 누락)
- 블랙리스트 단어(X-PRESSION)가 포함된 Description 라인이 제외됨

**근본 원인**:
- 블랙리스트는 "컬러 라인처럼 보이지만 실제로는 Description인 경우"를 거르기 위한 것
- 하지만 컬러 패턴이 없는 순수 Description 라인도 블랙리스트 체크를 받아 제외됨

**수정 위치**: [Invoice_Parser.js:1435-1443](Invoice_Parser.js:1435-1443)

**수정 내용**:
```javascript
// 수정 후: 컬러 패턴이 있을 때만 블랙리스트 체크
if (hasColorPattern) {  // ← CRITICAL: 조건 추가
  for (var bi = 0; bi < DESCRIPTION_BLACKLIST.length; bi++) {
    if (upperLine.indexOf(DESCRIPTION_BLACKLIST[bi]) > -1) {
      hasBlacklistedWord = true;
      break;
    }
  }
}
```

**효과**:
- 컬러 패턴(`COLOR- QTY`)이 있는 경우에만 블랙리스트로 제품명과 구분
- 순수 Description 라인은 블랙리스트 단어가 있어도 정상 수집

### 2025-12-27: Fee 패턴 단어 경계 추가 (FEED 오인식 방지)
**문제**: "BOHEMIAN FEED" 제품명이 메타데이터 "Fee"로 오인식됨
- 예: `X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED`
- "FEED" 단어가 "Fee" 패턴과 부분 매칭되어 유효한 Description으로 인식 안 됨
- `metadataMatch: FEE` → `isDescriptionLine: false`

**근본 원인**:
- Description 수집 시 메타데이터 필터링에서 "Fee"가 단어 경계 없이 매칭
- "FEED"의 앞 3글자 "FEE"가 메타데이터 패턴에 걸림

**수정 위치**:
- [Invoice_Parser.js:1384](Invoice_Parser.js:1384) - Description 수집 시 메타데이터 체크
- [Invoice_Parser.js:1403](Invoice_Parser.js:1403) - 디버깅 로그용 메타데이터 체크

**수정 내용**:
```javascript
// 수정 전
!nextLine.match(/...COD|Fee|tag.../i)

// 수정 후: Fee에 단어 경계 추가
!nextLine.match(/...COD|\bFee\b|tag.../i)
```

**효과**:
- "FEED" 단어가 "Fee" 패턴에 매칭되지 않음 (단어 경계로 구분)
- "COD Fee" 같은 실제 메타데이터만 걸러짐
- "X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED" 정상 인식

### 2025-12-27: 괄호 컬러 패턴 블랙리스트 예외 처리
**문제**: Description과 컬러가 같은 줄에 있을 때 컬러 라인 수집 실패
- 예: `X-PRESSION BRAID-PRE STRETCHED BRAID 52" 3X (P)M950/425/350/130S- 55`
- "X-PRESSION"이 블랙리스트에 있어서 컬러 라인으로 인식 안 됨
- 결과: colorLinesArray가 비어있어 다음 라인 검색 로직으로 진입
- 잘못된 컬러 라인(Line 470)을 수집하여 Lines 115-116 생성

**근본 원인**:
```javascript
// Line 1471
var isColorLine = hasColorPattern && !isInchPattern && !hasBlacklistedWord;
// hasBlacklistedWord: true → isColorLine: false
```

**수정 위치**: [Invoice_Parser.js:1478-1499](Invoice_Parser.js:1478-1499)

**수정 내용**:
```javascript
// 괄호 컬러 패턴 감지 (예: (P)M950/425/350/130S- 55)
var hasParenColorPattern = nextLine.match(/\([A-Z]\)[A-Z0-9\-\/]+\s*-\s*\d+/);

if (isDescriptionCandidate && !isColorLine && !hasParenColorPattern) {
  // Description만 추가
  descriptionLines.push(nextLine);
  continue;
} else if (isDescriptionCandidate && (isColorLine || hasParenColorPattern)) {
  // Description + 컬러 라인으로 처리
  descriptionLines.push(nextLine);
  if (hasParenColorPattern) {
    isColorLine = true; // 블랙리스트 무시
  }
  // continue 하지 않음 - 컬러 처리로 진행
}
```

**효과**:
- 괄호 컬러 패턴이 있으면 블랙리스트를 무시하고 컬러 라인으로 처리
- Line 461이 colorLinesArray에 추가됨
- 잘못된 컬러 라인(Line 470) 수집 방지
- Lines 115-116이 정확한 컬러로 생성됨

### 2025-12-27: Description 정리 로직 3X 보존
**문제**: Description에서 "3X" suffix가 제거됨
- 예상: `X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED TWIST 10" 3X`
- 실제: `X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED TWIST 10"`
- "3X"가 누락됨

**근본 원인**:
- Line 1596의 정규식이 인치 뒤의 "3X"를 캡처하지 못함
```javascript
// 기존
var allInchesPattern = description.match(/^(.+?)(\d+["″''](?:\s*\d+["″''])*(?:\s*-\s*[A-Z]{2,3})?)/);
// "10" - HT"는 매칭, "10" 3X"는 매칭 안 됨
```

**수정 위치**: [Invoice_Parser.js:1596](Invoice_Parser.js:1596)

**수정 내용**:
```javascript
// 수정 후: " - HT" 또는 " 3X" 같은 suffix 허용
var allInchesPattern = description.match(/^(.+?)(\d+["″''](?:\s*\d+["″''])*(?:\s*(?:-\s*[A-Z]{2,3}|\d*X))?)/);
//                                                                                     ^^^^^ 추가
// \d*X: 숫자(0~n개) + X (예: 3X, 4X, X)
```

**효과**:
- "TWIST 10" 3X"에서 "3X"가 보존됨
- 최종 Description: "X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED TWIST 10" 3X"

### 2025-12-27: 다음 아이템 Description 혼입 방지 (최종 수정)
**문제**: 이전 아이템 파싱 중 다음 아이템의 Description이 혼입됨
- Line 460 (QTY=55) 파싱 중 Line 467 (QTY=10)의 Description까지 검색
- Line 468 "X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED"가 Line 460의 Description에 추가될 위험

**근본 원인**:
- Description 수집 시 15줄 범위 내 모든 라인 검사
- QTY 라인(숫자만) 다음의 Description을 구분하지 못함

**수정 위치**: [Invoice_Parser.js:1376-1402](Invoice_Parser.js:1376-1402)

**수정 내용**:
```javascript
// CRITICAL: 바로 이전 줄이 QTY 전용 라인(숫자만)이면, 현재 줄은 다음 아이템의 Description
// 현재 아이템의 Description에 추가하면 안 됨!
// BUG FIX: j > i + 1 → j >= i + 2 (루프가 j = i + 1부터 시작하므로)
// - j = i + 1일 때: j-1 = i (현재 아이템의 QTY) → 체크 안 함
// - j = i + 2일 때: j-1 = i + 1 (다음 아이템의 QTY 가능) → 체크 함
var isPreviousLineQty = (j >= i + 2) && lines[j - 1].trim().match(/^\d{1,3}$/);

if (!isPreviousLineQty) {
  // Description 수집 로직...
  if ((isDescriptionLine || isInchLine || hasThreeNumberPattern || hasMultipleInches) && !hasPrices) {
    isDescriptionCandidate = true;
  }
} else {
  Logger.log('    ⏭️ 이전 줄이 QTY, 다음 아이템의 Description으로 판단: ' + nextLine.substring(0, 50));
}
```

**효과**:
- QTY 라인 직후의 Description은 다음 아이템으로 인식
- "X-PRESSION-LIL LOOKS-PASSION BOHEMIAN FEED"가 Line 460이 아닌 Line 467의 Description으로 파싱됨
- Description이 여러 줄인 경우 모든 줄이 올바르게 수집됨
- Lines 115-116 파싱 문제 완전 해결

### 2025-12-27: SNG 컬러 라인 파싱 로직 추가
**문제**: SNG 인보이스에서 컬러 정보가 파싱되지 않음
- 예: `SWDSWSX WD PREMIUM WIG S-WAVE (S)` 제품에 `NATURAL - 1` 컬러 라인 존재
- 컬러 라인이 다음 제품의 description으로 잘못 인식되어 100-200개 라인이 3300개로 증가
- SNG는 3줄 구조: Line 1 (기본 정보), Line 2 (할인가), Line 3+ (컬러)

**근본 원인**:
- SNG 파싱 로직에 컬러 라인 처리 코드가 전혀 없음
- Line 2 (할인가) 처리 후 바로 다음 아이템으로 이동
- Line 3 (컬러 라인)이 새로운 아이템으로 감지되어 무한 증식

**수정 위치**: [Invoice_Parser.js:1288-1407](Invoice_Parser.js:1288-1407)

**수정 내용**:
```javascript
// Line 2 (할인가) 처리 후 Line 3부터 컬러 라인 검색
var colorLinesArray = [];
for (var k = priceLineIndex + 1; k < Math.min(priceLineIndex + 10, lines.length); k++) {
  var colorLine = lines[k].trim();

  // 다음 아이템 라인을 만나면 중단 (T## 형식)
  if (colorLine.match(/^[A-Z]\d+\t/)) break;

  // 컬러 패턴: "NATURAL - 1", "1B - 2" 등
  if (colorLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/)) {
    colorLinesArray.push(colorLine);
  }
}

// 컬러 라인이 있으면 parseColorLinesImproved로 파싱
if (colorLinesArray.length > 0) {
  var colorData = parseColorLinesImproved(colorLinesArray, description);

  // 컬러별로 개별 라인 아이템 생성 (OUTRE와 동일한 로직)
  for (var m = 0; m < colorData.length; m++) {
    var cd = colorData[m];
    var itemExtPrice = Number((extPrice * (cd.shipped / totalShipped)).toFixed(2));

    items.push({
      lineNo: lineNo++,
      itemId: itemId,
      description: description,
      color: cd.color,
      qtyShipped: cd.shipped,
      unitPrice: unitPrice,
      extPrice: itemExtPrice,
      memo: cd.backordered > 0 ? 'Backordered: ' + cd.backordered : ''
    });
  }
  continue; // 다음 아이템으로
}
```

**효과**:
- SNG 인보이스에서 컬러별로 정확한 수량과 가격 파싱
- SWDSWSX 제품의 "NATURAL - 1" 컬러 정상 인식
- 3300개 라인 → 정상 100-200개 라인으로 감소
- OUTRE와 동일한 컬러 파싱 품질 제공

### 2025-12-27: Vendor별 컬러 파싱 함수 분리 (안정성 강화)
**목적**: OUTRE와 SNG 컬러 파싱 로직을 완전히 분리하여 상호 영향 방지

**문제**:
- OUTRE 파싱 로직이 완벽하게 작동 중
- SNG에서 동일한 함수(`parseColorLinesImproved`) 공유
- SNG 수정 시 OUTRE에 영향을 줄 위험 존재

**해결 방안**:
- `parseColorLinesImproved` → `parseOUTREColorLines` 리네임
- `parseSNGColorLines` 함수 새로 생성
- 각 vendor는 독립적인 함수 호출

**수정 위치**:
- [Invoice_Parser.js:2107-2202](Invoice_Parser.js:2107-2202) - `parseOUTREColorLines` (OUTRE 전용)
- [Invoice_Parser.js:2213-2301](Invoice_Parser.js:2213-2301) - `parseSNGColorLines` (SNG 전용)
- [Invoice_Parser.js:1320](Invoice_Parser.js:1320) - SNG에서 `parseSNGColorLines` 호출
- [Invoice_Parser.js:1971](Invoice_Parser.js:1971) - OUTRE에서 `parseOUTREColorLines` 호출

**함수 구조**:
```javascript
// OUTRE 전용 - 절대 수정 금지!
function parseOUTREColorLines(colorLines, description) {
  // OUTRE 컬러 파싱 로직 (완벽하게 작동 중)
  // 괄호 패턴, 블랙리스트, Description 제거 등 OUTRE 특화 로직 포함
}

// SNG 전용 - SNG 수정은 이 함수만!
function parseSNGColorLines(colorLines, description) {
  // 현재는 OUTRE와 동일한 로직
  // 향후 SNG 특수 요구사항 발생 시 이 함수만 수정
}
```

**효과**:
- ✅ OUTRE 로직 완전 보호 (SNG 수정 시 영향 없음)
- ✅ Vendor별 특수 로직 추가 용이
- ✅ 디버깅 시 vendor 구분 명확
- ✅ 향후 다른 vendor 추가 시 확장 용이
- ✅ 코드 안정성 및 유지보수성 향상

**주의사항**:
- `parseOUTREColorLines` 함수는 절대 수정 금지 (OUTRE 파싱 완벽 작동 중)
- SNG 수정 필요 시 `parseSNGColorLines` 함수만 수정
- 공통 버그 발견 시 두 함수 모두 수정 필요

### 2025-12-28: SNG 아이템 라인 감지 로직 개선 (중복 및 Description 혼입 문제 해결)
**문제**: 중복 라인 생성 및 다음 아이템의 Description 혼입
- 가격 라인(`\t4.00\t160.00\t80.00`)이 별도 아이템으로 인식됨
- 컬러 라인(`_ _ _ 1 - 10\t2 - 10...`)이 별도 아이템으로 인식됨
- 혼합 라인(`3.00\t120.00\t80.00 BLD-CRUSH - 10\t...`)에서 컬러 부분이 itemId로 파싱됨
- 결과: 100-200개 라인이 3300개로 증가, description 중복

**근본 원인**:
- SNG 아이템 라인 감지 시 `[숫자, 숫자, Item Number]` 패턴만 체크
- 가격 라인, 컬러 라인도 탭으로 구분되어 있어 패턴에 일부 매칭됨
- 예: `3.00\t120.00\t80.00 BLD-CRUSH - 10` → col2 = `80.00 BLD-CRUSH - 10` → 공백 뒤 `BLD-CRUSH`가 Item Number로 인식

**디버그 결과 분석**:
```
아이템 #1: SKBEX12 (정상)
  Line 68: T75\t40\t40\tSKBEX12\t...
  Line 69: \t4.00\t160.00\t80.00 (할인가)
  Line 70: _ _ _ 1 - 10 2 - 10... (컬러)
  ✅ 4개 컬러로 정상 분리

아이템 #2: "" (잘못 인식!)
  Line 69: \t4.00\t160.00\t80.00
  itemId: "", qtyOrdered: 160, qtyShipped: 80
  Line 70 컬러를 또 수집 → 중복 생성

아이템 #3: "_ _ _ _ _" (잘못 인식!)
  Line 70: _ _ _ 1 - 10\t2 - 10...
  itemId: "_ _ _ _ _"

아이템 #8: "BRN-CRUSH - 10" (잘못 인식!)
  Line 75: 3.00\t120.00\t80.00 BLD-CRUSH - 10\tBRN-CRUSH - 10...
  itemId: "BRN-CRUSH - 10", description: "OM27 - 10"
```

**수정 위치**: [Invoice_Parser.js:1184-1197](Invoice_Parser.js:1184-1197)

**수정 내용**:
```javascript
// CRITICAL: 가격 라인, 컬러 라인, 혼합 라인 제외
// 1. col0이 빈 문자열이면 제외 (가격 라인: "\t4.00\t160.00\t80.00")
// 2. col0이 소수점 숫자면 제외 (가격 라인: "3.00\t120.00\t80.00...")
// 3. col2가 언더스코어로 시작하면 제외 (컬러 라인: "_ _ _ 1 - 10\t...")
// 4. col2에 공백이 있으면 제외 (혼합 라인: "80.00 BLD-CRUSH - 10")
// 5. col2에 하이픈+공백+숫자 패턴이 있으면 제외 (컬러 패턴: "BLD-CRUSH - 10")

if (!col0 || col0.indexOf('.') > -1) {
  continue; // 가격 라인 제외
}

if (col2.indexOf('_') === 0 || col2.indexOf(' ') > -1 || col2.match(/\-\s*\d+/)) {
  continue; // 컬러 라인 및 혼합 라인 제외
}
```

**효과**:
- ✅ 가격 라인이 별도 아이템으로 인식되지 않음
- ✅ 컬러 라인이 별도 아이템으로 인식되지 않음
- ✅ 혼합 라인의 컬러 부분이 itemId로 파싱되지 않음
- ✅ 중복 라인 완전 제거 (3300개 → 정상 100-200개)
- ✅ Description 혼입 문제 해결

**주의사항**:
- 실제 아이템 번호가 하이픈+숫자를 포함하는 경우 문제 발생 가능 (예: `ABC-123`)
- 하지만 SNG 아이템 번호는 `SKBEX12`, `SWDSWSX` 같은 형식이므로 문제 없음

### 2025-12-28: SNG 상세 디버그 함수 최적화 (아이템 감지 라인만 로깅)
**목적**: 가격/컬러 라인 필터가 작동하지 않는 근본 원인 파악 (로그 최소화)

**문제**:
- Lines 1191-1197에 필터 로직 추가했으나 여전히 가격/컬러 라인이 아이템으로 감지됨
- 디버그 출력에서 Line 69 (`4.00\t160.00\t80.00`)가 아이템으로 감지됨
- 필터가 실행되는지, 조건이 올바른지 확인 필요
- **모든 라인 로깅 시 1600+ 줄로 로그 과다 발생** → 아이템 감지 라인만 로깅으로 최적화

**수정 위치**: [DebugSNG.js:92-204](DebugSNG.js:92-204)

**핵심 개선사항**:
1. **필터링된 라인은 로그 출력 안 함** (가격/컬러 라인은 건너뜀)
2. **아이템으로 감지된 라인만** 상세 필터 평가 출력
3. 로그 분량 대폭 감소 (1600+ 줄 → ~100-200 줄)

**수정 내용**:
```javascript
for (var i = 0; i < lines.length && i < maxLinesToScan; i++) {
  var tabParts = line.split('\t');
  var qtyOrderedCol, qtyShippedCol, itemNumberCol;

  // 필터 검사 (로그 출력 안 함)
  for (var startCol = 0; startCol < Math.min(3, tabParts.length - 2); startCol++) {
    var col0 = tabParts[startCol] ? tabParts[startCol].trim() : '';
    var col1 = tabParts[startCol + 1] ? tabParts[startCol + 1].trim() : '';
    var col2 = tabParts[startCol + 2] ? tabParts[startCol + 2].trim() : '';

    // Filter 1: 가격 라인 제외
    if (!col0 || col0.indexOf('.') > -1) continue;

    // Filter 2: 컬러 라인 제외
    if (col2.indexOf('_') === 0 || col2.indexOf(' ') > -1 || col2.match(/\-\s*\d+/)) continue;

    // Pattern Match
    if (!isNaN(parseInt(col0)) && !isNaN(parseInt(col1)) && col2.match(/^[A-Z][A-Z0-9]+$/)) {
      qtyOrderedCol = startCol;
      // ... 매치됨
      break;
    }
  }

  // 필터링된 라인은 로그 출력 안 함
  if (qtyOrderedCol === undefined) {
    filteredLines++;
    continue;
  }

  // ========================================
  // ✅ 아이템으로 감지된 라인만 상세 로그 출력
  // ========================================
  collectDebugLog('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  collectDebugLog('✅ Line ' + i + ' 아이템 라인 감지됨');
  collectDebugLog('Raw: ' + line.substring(0, 100));
  collectDebugLog('Tab Parts: ' + JSON.stringify(tabParts));

  // 각 startCol 필터 평가 상세 출력
  for (var startCol = 0; startCol < Math.min(3, tabParts.length - 2); startCol++) {
    collectDebugLog('  [startCol=' + startCol + ']');
    collectDebugLog('    Filter1: empty=' + f1_empty + ', decimal=' + f1_decimal + ' → ' + (skip ? 'SKIP ✋' : 'PASS ✓'));
    collectDebugLog('    Filter2: underscore=' + f2_underscore + ', space=' + f2_space + ', color=' + f2_color + ' → ' + (skip ? 'SKIP ✋' : 'PASS ✓'));
    collectDebugLog('    Pattern: num=' + isNum0 + ', num=' + isNum1 + ', item=' + isItem + ' → ' + (match ? 'MATCH ✅' : 'FAIL ❌'));
  }
}
```

**출력 예시** (아이템 라인만):
```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
✅ Line 45 아이템 라인 감지됨
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Raw: 10	10	SKBEX12	SHAKE-N-GO FREETRESS EQUAL SYNTHETIC...
Tab Parts (6개): ["10","10","SKBEX12","SHAKE-N-GO..."]

필터 검사 상세:
  [startCol=0]
    col0="10", col1="10", col2="SKBEX12"
    Filter1: empty=false, decimal=false → PASS ✓
    Filter2: underscore=false, space=false, color=false → PASS ✓
    Pattern: num=true, num=true, item=true → MATCH ✅

Detection: startCol 0: MATCHED ✅
```

**최종 통계 출력**:
```
📊 최종 통계
전체 라인 수: 850
스캔한 라인 수: 120 (빈 줄 제외)
필터링된 라인 수: 110 (가격/컬러 라인) ← 로그 출력 안 함
아이템으로 감지된 라인: 10 ✅ ← 이것만 상세 로그 출력
생성된 라인 아이템 수: 45 (컬러 분리 포함)
```

**효과**:
- ✅ 로그 분량 90% 감소 (1600+ 줄 → ~100-200 줄)
- ✅ 아이템 감지 라인만 집중 분석 가능
- ✅ 필터가 제대로 작동하는지 즉시 확인
- ✅ 구글 시트 붙여넣기 가능한 분량

**사용법**:
1. Google Apps Script 편집기에서 `debugSNGParsing()` 실행
2. DEBUG_OUTPUT 시트에서 상세 로그 확인
3. **아이템으로 감지된 라인**의 필터 평가 결과 분석
4. 잘못 감지된 라인이 있다면 어느 필터가 통과했는지 확인

### 2025-12-28: SNG 헤더 구간 건너뛰기 로직 개선 (검색 범위 확장 + 헤더 끝 감지 강화)
**문제**: SWDSWSX 제품의 컬러 정보가 파싱되지 않음
- 컬러 라인이 line 553에 존재하지만 검색 범위(30줄)가 line 502까지만 도달
- 중간에 헤더 구간(lines 486-501)이 있어서 검색 범위 부족
- 헤더 끝 패턴(`List Extended`, `Your Extended` 등)이 이 인보이스에 없어서 `inHeaderSection` 플래그가 계속 true 상태로 유지됨
- 결과: "⚠️ 컬러 정보 찾을 수 없음" 오류

**근본 원인**:
1. 검색 범위 30줄은 헤더 구간(15-20줄) + 컬러 라인을 포함하기에 부족
2. 헤더 끝 감지가 특정 패턴에만 의존하여, 패턴이 없으면 영구히 헤더 모드
3. 헤더가 끝났어도 flag가 reset되지 않아 이후 모든 라인을 skip

**수정 위치**: [Invoice_Parser.js:1384-1432](Invoice_Parser.js:1384-1432)

**수정 내용**:
```javascript
// 1. 검색 범위 30 → 80으로 확장
for (var k = priceLineIndex; k < Math.min(priceLineIndex + 80, lines.length); k++) {

// 2. 헤더 끝 감지 방법 3가지로 확장:
// (a) 기존 패턴 매칭: List Extended, Your Extended, Discounted Amount
// (b) 빈 줄 감지: 마지막 헤더 키워드 이후 3줄 이내 빈 줄
if (inHeaderSection && lastHeaderLineIndex > -1 && (k - lastHeaderLineIndex) <= 3) {
  Logger.log('    📋 헤더 구간 종료 감지 (빈 줄): Line ' + k);
  inHeaderSection = false;
}

// (c) 아이템 라인 감지: T##\t 형식 라인이 나타나면 헤더 종료
if (inHeaderSection && colorLine.match(/^[A-Z]\d+\t/)) {
  Logger.log('    📋 헤더 구간 종료 감지 (아이템 라인 발견)');
  inHeaderSection = false;
}
```

**효과**:
- ✅ SWDSWSX 컬러 라인(line 553) 정상 도달
- ✅ 헤더 끝 패턴이 없어도 빈 줄 또는 다음 아이템으로 헤더 종료 감지
- ✅ `inHeaderSection` flag가 적절히 reset되어 컬러 수집 정상 작동
- ✅ "⚠️ 컬러 정보 찾을 수 없음" 오류 해결

**주의사항**:
- 헤더 구간이 매우 긴 경우(25줄 이상) 추가 확장 필요할 수 있음
- 빈 줄 감지는 lastHeaderLineIndex로부터 3줄 이내만 체크 (너무 멀리 떨어진 빈 줄은 무시)

### 2025-12-27: SNG 컬러 파싱 버그 수정 (단일 컬러명 + 다른 제품 컬러 혼입 방지)
**문제 1**: 단일 컬러명 패턴 미지원
- 예: `_ _ _ COPPER _ _` (언더스코어로 둘러싸인 컬러명)
- 기존 패턴 `[A-Z0-9\-\/]+\s*-\s*\d+`는 "COLOR - QTY" 형식만 인식
- 단일 컬러명은 `-` + 숫자가 없어서 매칭 실패

**문제 2**: 다른 제품의 컬러 라인 혼입
- 예: SWGRBAS (Line 554) 파싱 시 SWDSWSX의 컬러까지 수집
- SWGRBAS 자신의 컬러: `1 - 0 (2)   1B - 0 (2)   2 - 0 (2)   4 - 2` (QTY=2)
- 다음 라인 검색 로직에서 SWDSWSX의 컬러(`T30 - 10`, `T530 - 10`, ...) 추가 수집
- 결과: totalShipped = 42개 (실제 2개), ExtPrice 불일치 ($120 vs $17.14)
- 근본 원인: SNG 처리 완료 후 `continue` 누락으로 OUTRE 로직 및 다음 라인 검색 진입

**수정 위치**:
- [Invoice_Parser.js:1309-1323](Invoice_Parser.js:1309-1323) - 단일 컬러명 패턴 추가
- [Invoice_Parser.js:1420](Invoice_Parser.js:1420) - SNG 처리 완료 시 `continue` 추가

**수정 내용 (문제 1 해결)**:
```javascript
// 기존: 표준 패턴만
if (colorLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/)) {
  colorLinesArray.push(colorLine);
}

// 수정 후: 표준 + 단일 컬러명 패턴
var hasStandardColorPattern = colorLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/);
var hasSingleColorPattern = colorLine.match(/^[\s_]+([A-Z][A-Z0-9\-\/]{1,15})[\s_]+$/);

if (hasStandardColorPattern) {
  colorLinesArray.push(colorLine);
  Logger.log('    ✓ 컬러 라인 수집 (표준): ' + colorLine);
} else if (hasSingleColorPattern) {
  // 단일 컬러명 → "COLORNAME - 1" 형식으로 변환 (shipped=1 가정)
  var singleColor = hasSingleColorPattern[1];
  var convertedLine = singleColor + ' - 1';
  colorLinesArray.push(convertedLine);
  Logger.log('    ✓ 컬러 라인 수집 (단일): ' + singleColor + ' → ' + convertedLine);
}
```

**수정 내용 (문제 2 해결)**:
```javascript
// Line 1417 다음에 추가
items.push(item);
debugLog('SNG 단일 아이템 추가', item);

// CRITICAL: SNG 처리 완료 → 다음 아이템으로 (OUTRE 로직 및 다음 라인 검색 스킵)
continue;
```

**효과**:
- ✅ 단일 컬러명(`_ _ COPPER _ _`) 정상 인식
- ✅ SNG 처리 완료 시 다음 아이템으로 즉시 이동
- ✅ 다른 제품의 컬러 라인 혼입 완전 차단
- ✅ ExtPrice 계산 정확도 향상 (totalShipped 값 정상화)

**영향 범위**:
- SNG 인보이스 컬러 파싱 정확도 향상
- OUTRE 로직에는 영향 없음 (완전 분리됨)

## 현재 미해결 문제 (2025-12-28)

### 문제 1: PACKED BY 컬럼(T75 등) 처리 미흡 ✅ 해결됨
**증상**: `T75`, `T30` 등의 PACKED BY 데이터가 아이템 감지 로직에 간섭
**영향**: 아이템 라인 감지 시 startCol 오프셋 계산 오류 가능

**해결 방법**:
- 현재 코드는 이미 startCol 0, 1, 2를 스캔하여 `[숫자, 숫자, Item Number]` 패턴을 찾음
- PACKED BY가 있으면 startCol=1에서 매칭, 없으면 startCol=0에서 매칭
- **추가 조치 불필요** - 로직이 이미 PACKED BY를 올바르게 처리함

**문제 아이템**: 없음 (해결됨)

### 문제 2: Description 추출 실패 (이전 아이템 Description 복사)
**증상**: 일부 아이템에서 자신의 Description 대신 바로 이전 아이템의 Description을 가져옴

**문제 아이템 및 영향 관계**:
- SKFCX18 ← SKBTX28의 Description 복사
- SHBN36X ← SHBN34X의 Description 복사
- SOB4A12 ← SOATX30의 Description 복사
- SOBDX24 ← SOBCX24의 Description 복사
- SOHWX18 ← SOGBBM3의 Description 복사
- SOLDX36 ← SOLDX24의 Description 복사
- SGOXX12 ← SPRWX24의 Description 복사

**패턴**: 연속된 두 아이템에서 두 번째 아이템이 첫 번째 아이템의 Description을 복사

**근본 원인 (추정)**:
1. Description 검색 범위가 다음 아이템까지 확장됨
2. 현재 아이템의 Description이 없거나 찾지 못했을 때 이전 변수값 재사용
3. 아이템 라인 감지 후 Description 수집 전에 다음 아이템 라인을 만나는 경우

### 문제 3: SWDSWSX 컬러 파싱 실패
**증상**: SWDSWSX 제품의 컬러 정보가 지속적으로 파싱되지 않음

**예상 원인**:
- 헤더 구간 건너뛰기 로직 개선 후에도 여전히 컬러 라인 미발견
- 컬러 라인 패턴이 표준과 다를 가능성
- 검색 범위 80줄로도 컬러 라인에 도달하지 못할 가능성
- 컬러 라인이 특수 형식으로 되어있어 현재 정규식으로 매칭 실패

**다음 단계**: SWDSWSX 주변 라인 상세 분석 필요

**디버깅 개선** (2025-12-28):
- DebugSNG.js에 SWDSWSX 특별 디버깅 추가
- 컬러 라인이 없을 경우 다음 100줄 자동 출력
- Description Raw Value 로깅 추가
- Invoice_Parser.js에 Description 비어있을 때 경고 로그 추가

### 2025-12-29: OUTRE Description 예외 패턴 처리 (REMI TARA)
**문제**: REMI TARA 제품명의 숫자-숫자-숫자 패턴이 컬러로 오인식됨
- 예: "REMI TARA 1-2-3", "REMI TARA 2-4-6", "REMI TARA 4-6-8"
- "1-2-3" 패턴이 `[A-Z0-9\-\/]+\s*-\s*\d+` 컬러 정규식에 매칭됨
- Description이 비어있고 컬러 라인에 "1-2-3"만 수집되는 문제

**근본 원인**:
- 컬러 패턴 판정(isColorLine)이 Description 수집보다 우선 실행
- REMI TARA는 제품명 자체가 숫자-숫자-숫자 포함하는 예외 케이스
- 기존 로직: 컬러 패턴 → Description 제외 → Description 없음

**수정 위치**: [Invoice_Parser_OUTRE.js:201-246](Invoice_Parser_OUTRE.js:201-246)

**수정 내용**:
```javascript
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
var hasColorPattern = nextLine.match(/[A-Z0-9\-\/]+\s*-\s*\d+/);
// ... 기존 로직 계속
```

**효과**:
- ✅ REMI TARA 패턴이 Description으로 확정됨 (컬러 판정보다 우선)
- ✅ 같은 라인에 컬러가 있어도 분리하여 수집 가능
- ✅ 기존 컬러 판정 로직(isColorLine) 변경 없음 (회귀 방지)
- ✅ 향후 유사 예외 케이스 추가 용이 (패턴 배열 방식)

**기존 하드코딩 제거**:
- Lines 338-355의 REMI TARA 사후 처리 로직 제거 (중복)
- 예외 처리를 최전방(컬러 판정 전)으로 이동하여 근본 해결

**장기 안정성**:
- 예외 패턴이 일반 로직에 간섭하지 않음 (Early-exit)
- DESCRIPTION_EXCEPTION_PATTERNS 배열로 확장 가능
- DB 검증 로직과 독립적으로 작동

## 알려진 이슈 및 제한사항

### Document AI
- PDF 크기 제한: 최대 20MB
- 처리 시간: 파일당 5-10초
- 비용: 페이지당 과금

### Apps Script
- 실행 시간 제한: 6분
- 메모리 제한: 100MB
- 동시 실행: 제한적

### 컬러 파싱
- Description에 컬러 라인이 포함되지 않으면 감지 실패 가능
- 비표준 포맷은 수동 확인 필요
- Memo에 "⚠️ 컬러 정보 찾을 수 없음" 표시

## 유지보수 가이드

### 새로운 Vendor 추가
1. `Config.js`의 `CONFIG.COMPANIES`에 추가
2. `CONFIG.INVOICE.BRANDS`에 브랜드명 추가
3. `DocumentAI.js`의 `convertDocumentAIToInvoiceData()`에서 vendor 감지 로직 추가
4. `getVendorSpecificPrices()`에 가격 필드 매핑 추가

### Document AI 응답 필드 변경 시
1. `debugDocumentAIResponse()` 실행하여 실제 필드 확인
2. `extractEntity()` 호출 부분 수정
3. `getVendorSpecificPrices()` 업데이트

### 컬러 파싱 로직 개선 시
1. `parseColorLinesImproved()` 함수 수정
2. 정규식 패턴 조정
3. 테스트 데이터로 검증

## 참고 자료
- [Google Apps Script 문서](https://developers.google.com/apps-script)
- [Document AI 문서](https://cloud.google.com/document-ai/docs)
- [Clasp CLI](https://github.com/google/clasp)
