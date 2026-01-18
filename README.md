# HWPX 신청서 자동 생성기

Excel 데이터를 기반으로 HWPX(한글) 신청서를 자동 생성하는 Node.js 스크립트

## 핵심 아이디어

HWPX 파일은 실제로 **ZIP 압축 파일**입니다. 내부 XML을 직접 수정하여 데이터를 입력합니다.

```
Form.hwpx (ZIP)
├── Contents/
│   └── section0.xml  ← 문서 내용 (표, 텍스트)
├── META-INF/
└── ...
```

## 사용법

### 1. 의존성 설치
```bash
npm install xlsx adm-zip
```

### 2. 파일 준비
- `Form.hwpx` - 템플릿 파일 (빈 양식)
- `필지정보.xlsx` - Excel 데이터

### 3. 실행
```bash
node generate_hwpx_v5.js
```

### 4. 결과
`output/` 폴더에 신청서 파일들이 생성됩니다.

## 전체 데이터 처리

스크립트에서 한 줄 수정:
```javascript
// 5명 테스트
const entries = Object.entries(dataByBusiness).slice(0, 5);

// 전체 처리
const entries = Object.entries(dataByBusiness);
```

## 파일 구조

```
├── Form.hwpx                    # 템플릿 (빈 양식)
├── 필지정보.xlsx                 # Excel 원본 데이터
├── generate_hwpx_v5.js          # 메인 스크립트
├── analyze_structure.js         # XML 구조 분석 도구
├── check_generated.js           # 생성 결과 확인 도구
└── output/                      # 생성된 신청서들
    ├── 신청서_홍길동_1234567890.hwpx
    └── cell_mappings.json
```

## 테이블 셀 매핑 (Form.hwpx 기준)

| 항목 | Row | Cell |
|------|-----|------|
| 성명 | 1 | 2 |
| 생년월일 | 1 | 4 |
| 주소 | 2 | 1 |
| 연락처 | 3 | 1 |
| 필지 정보 | 6~17 | 0,1,2,6 |
| 합계 | 18 | 3 |
| 계좌번호 | 19 | 3 |
| 예금주명 | 19 | 5 |

## 기술 스택

- **Node.js** - 런타임
- **xlsx** - Excel 파일 읽기
- **adm-zip** - HWPX(ZIP) 압축/해제

## 왜 MCP 대신 직접 XML 수정?

HWPX MCP가 파일을 열 때 오류 발생:
```
syntax error: line 1, column 0
```

그래서 ZIP/XML 직접 조작 방식으로 우회했습니다. 이 방식이 더 빠르고 안정적입니다.
