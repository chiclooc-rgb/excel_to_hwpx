# Skill: 한글서식 분석

## 트리거
사용자가 다음과 같은 요청을 할 때 이 skill을 적용:
- "HWPX 양식에 Excel 데이터 넣어줘"
- "한글 문서 대량 생성해줘"
- "신청서/보고서 자동화해줘"
- 메일머지로 해결 안 되는 1:N 데이터 매핑 (한 사람 = 여러 행)

## 핵심 개념

**HWPX = ZIP + XML**

```
문서.hwpx (ZIP 압축파일)
├── Contents/
│   └── section0.xml  ← 본문 내용 (표, 텍스트)
├── Header/
│   └── header.xml    ← 스타일 정의 (폰트, 크기 등)
└── META-INF/
```

## 작업 프로세스

### 1단계: 구조 분석 (AI가 1회 수행)

```javascript
// analyze_structure.js - XML 테이블 구조 파악용
const AdmZip = require('adm-zip');
const zip = new AdmZip('Form.hwpx');
const xml = zip.getEntry('Contents/section0.xml').getData().toString('utf8');

// 테이블/행/셀 구조 출력
const tblRegex = /<hp:tbl[^>]*>([\s\S]*?)<\/hp:tbl>/g;
const trRegex = /<hp:tr[^>]*>([\s\S]*?)<\/hp:tr>/g;
const tcRegex = /<hp:tc[^>]*>([\s\S]*?)<\/hp:tc>/g;
const tRegex = /<hp:t[^>]*>([^<]*)<\/hp:t>/g;
```

### 2단계: 매핑 정의

분석 결과를 바탕으로 "어느 데이터 → 어느 셀" 매핑:

```javascript
// 예시: 신청서 양식
const CELL_MAPPING = {
    성명: { table: 0, row: 1, cell: 2 },
    생년월일: { table: 0, row: 1, cell: 4 },
    주소: { table: 0, row: 2, cell: 1 },
    연락처: { table: 0, row: 3, cell: 1 },
    필지시작: { table: 0, row: 6 },  // 반복 데이터
};
```

### 3단계: Node.js 스크립트 생성

```javascript
const fs = require('fs');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');

// Excel 읽기
const workbook = XLSX.readFile('데이터.xlsx');
const data = XLSX.utils.sheet_to_json(workbook.Sheets[시트명]);

// 각 건별 HWPX 생성
data.forEach(row => {
    const zip = new AdmZip('템플릿.hwpx');
    let xml = zip.getEntry('Contents/section0.xml').getData().toString('utf8');

    // XML 수정
    xml = modifyCell(xml, ...);

    // 저장
    zip.updateFile('Contents/section0.xml', Buffer.from(xml, 'utf8'));
    zip.writeZip(`output/결과_${row.이름}.hwpx`);
});
```

## 핵심 함수: 셀 내용 수정

```javascript
function modifyTextInCell(cellContent, value) {
    // 1. 기존 텍스트가 있는 경우
    const tRegex = /(<hp:t[^>]*>)([^<]*)(<\/hp:t>)/;
    if (tRegex.test(cellContent)) {
        return cellContent.replace(tRegex, `$1${value}$3`);
    }

    // 2. 빈 셀인 경우 (self-closing run)
    const runSelfClosingRegex = /<hp:run([^>]*)\/>/;
    if (runSelfClosingRegex.test(cellContent)) {
        return cellContent.replace(
            runSelfClosingRegex,
            `<hp:run$1><hp:t>${value}</hp:t></hp:run>`
        );
    }

    return cellContent;
}
```

## 필수 라이브러리

```bash
npm install xlsx adm-zip
```

| 라이브러리 | 용도 |
|-----------|------|
| `xlsx` | Excel 파일 읽기 |
| `adm-zip` | HWPX(ZIP) 압축/해제 |

## 주의사항

### MCP 우회
HWPX MCP가 `syntax error: line 1, column 0` 오류 발생 시 → 직접 ZIP/XML 조작으로 우회

### 빈 셀 처리
빈 셀은 `<hp:t>` 태그가 없음. 반드시 변환 필요:
```xml
<!-- Before (빈 셀) -->
<hp:run charPrIDRef="7"/>

<!-- After (값 입력) -->
<hp:run charPrIDRef="7"><hp:t>입력값</hp:t></hp:run>
```

### 스타일 유지
`charPrIDRef` 속성은 스타일 참조. 수정 시 보존해야 서식 유지됨.

### 다중 페이지 (별지)
필지 등 반복 데이터가 한 페이지 초과 시:
- 별지 포함 템플릿 별도 준비
- 테이블 인덱스로 구분 (Table 0 = 1페이지, Table 1 = 별지)

## 실제 적용 예시

**벼재배농가 신청서 자동화:**
- 입력: Excel 355명 x 다중 필지
- 출력: 개인별 HWPX 신청서
- 특이사항: 1인당 필지 수 가변 (1~39개)
- 해결: 필지 12개 초과 시 별지 템플릿 자동 선택

## 참고 저장소

https://github.com/chiclooc-rgb/excel_to_hwpx
