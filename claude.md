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
- **Utils.js**: 공통 유틸리티 함수
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
