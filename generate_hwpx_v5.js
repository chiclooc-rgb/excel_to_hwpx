/**
 * 벼재배농가 경영안정자금 지급대상자 신청서 자동 생성 스크립트 v5
 * 빈 셀에 <hp:t> 태그 추가하는 로직
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const AdmZip = require('adm-zip');

// ============================================
// 설정
// ============================================
const BASE_DIR = __dirname;
const EXCEL_FILE = path.join(BASE_DIR, '필지정보.xlsx');
const TEMPLATE_FILE = path.join(BASE_DIR, 'Form.hwpx');
const OUTPUT_DIR = path.join(BASE_DIR, 'output');
const SHEET_NAME = '메일머지 작업전(전략작물,타작물추가)';

// ============================================
// 유틸리티 함수
// ============================================

function formatBirthDate(birthStr) {
    if (!birthStr || String(birthStr).length !== 8) return String(birthStr || '');
    const s = String(birthStr);
    return `${s.slice(0, 4)}.${s.slice(4, 6)}.${s.slice(6)}`;
}

function formatPhone(phoneStr) {
    if (!phoneStr) return '';
    const s = String(phoneStr).replace(/-/g, '');
    if (s.length === 11) {
        return `${s.slice(0, 3)}-${s.slice(3, 7)}-${s.slice(7)}`;
    } else if (s.length === 10) {
        return `${s.slice(0, 3)}-${s.slice(3, 6)}-${s.slice(6)}`;
    }
    return s;
}

function formatJibun(bon, bu) {
    const bonStr = String(bon || '0').padStart(4, '0');
    const buStr = String(bu || '0').padStart(4, '0');
    return `${bonStr}-${buStr}`;
}

function fixRiName(riName) {
    if (riName === '마랑리') return '마룡리';
    return riName || '';
}

// ============================================
// Excel 데이터 읽기
// ============================================

function readExcelData() {
    console.log(`Excel 파일 읽는 중: ${EXCEL_FILE}`);

    const workbook = XLSX.readFile(EXCEL_FILE);
    const sheet = workbook.Sheets[SHEET_NAME];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const dataByBusiness = {};

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row[3]) continue;

        const businessId = String(row[3]);

        const parcelInfo = {
            읍면동: row[0] || '',
            마을명: row[1] || '',
            생년월일: row[2],
            경영체등록번호: businessId,
            성명: row[4] || '',
            주소: row[5] || '',
            읍면: row[6] || '',
            리: fixRiName(row[7]),
            본번: row[8],
            부번: row[9],
            벼경작면적: row[10] || 0,
            연락처: row[11],
        };

        if (!dataByBusiness[businessId]) {
            dataByBusiness[businessId] = [];
        }
        dataByBusiness[businessId].push(parcelInfo);
    }

    console.log(`총 ${Object.keys(dataByBusiness).length}명의 신청자 데이터 로드 완료`);
    return dataByBusiness;
}

// ============================================
// 개인 데이터 생성
// ============================================

function createPersonData(parcels) {
    const firstParcel = parcels[0];

    const personData = {
        성명: firstParcel.성명,
        생년월일: formatBirthDate(firstParcel.생년월일),
        주소: firstParcel.주소,
        연락처: formatPhone(firstParcel.연락처),
        경영체등록번호: firstParcel.경영체등록번호,
        필지목록: [],
        합계: 0,
    };

    let totalArea = 0;
    for (let i = 0; i < Math.min(parcels.length, 12); i++) {
        const parcel = parcels[i];
        const area = Number(parcel.벼경작면적) || 0;
        totalArea += area;

        personData.필지목록.push({
            읍면: parcel.읍면,
            리: parcel.리,
            지번: formatJibun(parcel.본번, parcel.부번),
            벼재배면적: area,
        });
    }

    if (parcels.length > 12) {
        console.log(`  경고: ${firstParcel.성명}님의 필지가 12개를 초과합니다.`);
    }

    personData.합계 = totalArea;
    return personData;
}

// ============================================
// XML 수정
// ============================================

/**
 * 특정 행과 셀의 내용을 수정
 * 빈 셀의 경우 <hp:run .../> 를 <hp:run ...><hp:t>값</hp:t></hp:run> 으로 변환
 */
function modifyXml(xml, personData) {
    let result = xml;

    // 수정할 셀 정보 (row, cell, value)
    const modifications = [
        // 신청인 정보
        { row: 1, cell: 2, value: personData.성명 },
        { row: 1, cell: 4, value: personData.생년월일 },
        { row: 2, cell: 1, value: personData.주소 },
        { row: 3, cell: 1, value: personData.연락처 },
    ];

    // 필지 정보 (Row 6~17)
    for (let i = 0; i < 12; i++) {
        const rowIdx = 6 + i;

        if (i < personData.필지목록.length) {
            const parcel = personData.필지목록[i];
            modifications.push({ row: rowIdx, cell: 0, value: parcel.읍면 });
            modifications.push({ row: rowIdx, cell: 1, value: parcel.리 });
            modifications.push({ row: rowIdx, cell: 2, value: parcel.지번 });
            modifications.push({ row: rowIdx, cell: 6, value: String(parcel.벼재배면적) });
        }
    }

    // 합계
    modifications.push({ row: 18, cell: 3, value: String(personData.합계) });

    // 계좌 정보
    modifications.push({ row: 19, cell: 3, value: '계좌번호 미기입' });
    modifications.push({ row: 19, cell: 5, value: personData.성명 });

    // 각 수정 사항을 역순으로 처리
    for (const mod of modifications) {
        result = modifyCell(result, mod.row, mod.cell, mod.value);
    }

    // 하단 서명란 "성명" → 실제 이름으로 변경
    // "신청인       성명      (인)" → "신청인       홍길동      (인)"
    result = result.replace(
        /신청인\s+성명\s+\(인\)/,
        `신청인       ${personData.성명}      (인)`
    );

    return result;
}

/**
 * 특정 행/셀의 내용 수정
 */
function modifyCell(xml, targetRow, targetCell, value) {
    // 모든 행 분리
    const rows = [];
    const trRegex = /(<hp:tr[^>]*>)([\s\S]*?)(<\/hp:tr>)/g;
    let lastIndex = 0;
    let match;

    while ((match = trRegex.exec(xml)) !== null) {
        // 이전 부분 저장
        if (match.index > lastIndex) {
            rows.push({ type: 'other', content: xml.substring(lastIndex, match.index) });
        }

        rows.push({
            type: 'row',
            start: match[1],
            content: match[2],
            end: match[3]
        });

        lastIndex = match.index + match[0].length;
    }

    // 나머지 부분
    if (lastIndex < xml.length) {
        rows.push({ type: 'other', content: xml.substring(lastIndex) });
    }

    // 해당 행 찾기 및 수정
    let rowIndex = 0;
    for (let i = 0; i < rows.length; i++) {
        if (rows[i].type === 'row') {
            if (rowIndex === targetRow) {
                rows[i].content = modifyCellInRow(rows[i].content, targetCell, value);
            }
            rowIndex++;
        }
    }

    // 다시 조합
    return rows.map(r => {
        if (r.type === 'row') {
            return r.start + r.content + r.end;
        }
        return r.content;
    }).join('');
}

/**
 * 행 내에서 특정 셀 수정
 */
function modifyCellInRow(rowContent, targetCell, value) {
    // 모든 셀 분리
    const cells = [];
    const tcRegex = /(<hp:tc[^>]*>)([\s\S]*?)(<\/hp:tc>)/g;
    let lastIndex = 0;
    let match;

    while ((match = tcRegex.exec(rowContent)) !== null) {
        if (match.index > lastIndex) {
            cells.push({ type: 'other', content: rowContent.substring(lastIndex, match.index) });
        }

        cells.push({
            type: 'cell',
            start: match[1],
            content: match[2],
            end: match[3]
        });

        lastIndex = match.index + match[0].length;
    }

    if (lastIndex < rowContent.length) {
        cells.push({ type: 'other', content: rowContent.substring(lastIndex) });
    }

    // 해당 셀 수정
    let cellIndex = 0;
    for (let i = 0; i < cells.length; i++) {
        if (cells[i].type === 'cell') {
            if (cellIndex === targetCell) {
                cells[i].content = modifyTextInCell(cells[i].content, value);
            }
            cellIndex++;
        }
    }

    return cells.map(c => {
        if (c.type === 'cell') {
            return c.start + c.content + c.end;
        }
        return c.content;
    }).join('');
}

/**
 * 셀 내용에서 텍스트 수정
 * 1. <hp:t>가 있으면 내용 교체
 * 2. <hp:run .../> (self-closing) 이면 <hp:run ...><hp:t>값</hp:t></hp:run>으로 변환
 */
function modifyTextInCell(cellContent, value) {
    // 1. 먼저 기존 <hp:t> 태그가 있는지 확인
    const tRegex = /(<hp:t[^>]*>)([^<]*)(<\/hp:t>)/;
    if (tRegex.test(cellContent)) {
        return cellContent.replace(tRegex, `$1${value}$3`);
    }

    // 2. Self-closing <hp:run .../> 태그를 찾아서 변환
    // 패턴: <hp:run charPrIDRef="7"/>
    const runSelfClosingRegex = /<hp:run([^>]*)\/>/;
    if (runSelfClosingRegex.test(cellContent)) {
        return cellContent.replace(runSelfClosingRegex, `<hp:run$1><hp:t>${value}</hp:t></hp:run>`);
    }

    // 3. <hp:run>...</hp:run> 형태이지만 <hp:t>가 없는 경우
    const runOpenCloseRegex = /(<hp:run[^>]*>)([\s\S]*?)(<\/hp:run>)/;
    const runMatch = runOpenCloseRegex.exec(cellContent);
    if (runMatch) {
        // run 태그 내부에 <hp:t>가 없으면 추가
        if (!/<hp:t/.test(runMatch[2])) {
            return cellContent.replace(runOpenCloseRegex, `$1<hp:t>${value}</hp:t>$2$3`);
        }
    }

    return cellContent;
}

// ============================================
// HWPX 파일 생성
// ============================================

function generateHwpxFile(personData, outputPath) {
    const zip = new AdmZip(TEMPLATE_FILE);

    const sectionEntry = zip.getEntry('Contents/section0.xml');
    if (!sectionEntry) {
        console.log('  경고: section0.xml을 찾을 수 없습니다.');
        return false;
    }

    let xml = sectionEntry.getData().toString('utf8');
    xml = modifyXml(xml, personData);

    zip.updateFile('Contents/section0.xml', Buffer.from(xml, 'utf8'));
    zip.writeZip(outputPath);

    return true;
}

// ============================================
// 메인 함수
// ============================================

function main() {
    console.log('='.repeat(50));
    console.log('벼재배농가 경영안정자금 지급대상자 신청서 자동 생성 v5');
    console.log('='.repeat(50));

    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }
    console.log(`출력 디렉토리: ${OUTPUT_DIR}`);

    const dataByBusiness = readExcelData();
    const allMappings = [];

    // 처음 5명만 테스트
    const entries = Object.entries(dataByBusiness).slice(0, 5);

    console.log(`\n${entries.length}명의 신청서 생성 시작...`);

    entries.forEach(([businessId, parcels], idx) => {
        const personData = createPersonData(parcels);

        const filename = `신청서_${personData.성명}_${businessId}.hwpx`;
        const outputPath = path.join(OUTPUT_DIR, filename);

        console.log(`[${idx + 1}/${entries.length}] ${personData.성명} (${parcels.length}개 필지) 생성 중...`);

        const success = generateHwpxFile(personData, outputPath);

        if (success) {
            console.log(`  -> ${filename} 생성 완료`);
        } else {
            console.log(`  -> ${filename} 생성 실패`);
        }

        allMappings.push({
            file: outputPath,
            businessId: businessId,
            data: personData
        });
    });

    const mappingsFile = path.join(OUTPUT_DIR, 'cell_mappings.json');
    fs.writeFileSync(mappingsFile, JSON.stringify(allMappings, null, 2), 'utf8');

    console.log(`\n매핑 정보 저장: ${mappingsFile}`);
    console.log('\n' + '='.repeat(50));
    console.log(`완료! ${entries.length}개의 신청서가 생성되었습니다.`);
    console.log(`출력 폴더: ${OUTPUT_DIR}`);
    console.log('='.repeat(50));
}

main();
