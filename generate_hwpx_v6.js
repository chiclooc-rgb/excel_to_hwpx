/**
 * 벼재배농가 경영안정자금 지급대상자 신청서 자동 생성 스크립트 v6
 * - 필지 12개 이하: Form.hwpx (1페이지)
 * - 필지 13개 이상: Form-backpage.hwpx (1페이지 + 별지)
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
const TEMPLATE_SINGLE = path.join(BASE_DIR, 'Form.hwpx');           // 1페이지용
const TEMPLATE_MULTI = path.join(BASE_DIR, 'Form-backpage.hwpx');   // 별지 포함
const OUTPUT_DIR = path.join(BASE_DIR, 'output');
const SHEET_NAME = '메일머지 작업전(전략작물,타작물추가)';

// 테이블 매핑 정보
const PAGE1_PARCEL_START_ROW = 6;   // 1페이지 필지 시작 행
const PAGE1_PARCEL_COUNT = 12;       // 1페이지 최대 필지 수
const PAGE2_PARCEL_START_ROW = 2;   // 별지 필지 시작 행 (테이블1 기준)
const PAGE2_PARCEL_COUNT = 27;       // 별지 최대 필지 수

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
    const maxParcels = PAGE1_PARCEL_COUNT + PAGE2_PARCEL_COUNT;  // 최대 39개

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
    for (let i = 0; i < Math.min(parcels.length, maxParcels); i++) {
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

    if (parcels.length > maxParcels) {
        console.log(`  경고: ${firstParcel.성명}님의 필지가 ${maxParcels}개를 초과합니다.`);
    }

    personData.합계 = totalArea;
    return personData;
}

// ============================================
// XML 수정 함수들
// ============================================

/**
 * 셀 내용에서 텍스트 수정
 */
function modifyTextInCell(cellContent, value) {
    const tRegex = /(<hp:t[^>]*>)([^<]*)(<\/hp:t>)/;
    if (tRegex.test(cellContent)) {
        return cellContent.replace(tRegex, `$1${value}$3`);
    }

    const runSelfClosingRegex = /<hp:run([^>]*)\/>/;
    if (runSelfClosingRegex.test(cellContent)) {
        return cellContent.replace(runSelfClosingRegex, `<hp:run$1><hp:t>${value}</hp:t></hp:run>`);
    }

    const runOpenCloseRegex = /(<hp:run[^>]*>)([\s\S]*?)(<\/hp:run>)/;
    const runMatch = runOpenCloseRegex.exec(cellContent);
    if (runMatch) {
        if (!/<hp:t/.test(runMatch[2])) {
            return cellContent.replace(runOpenCloseRegex, `$1<hp:t>${value}</hp:t>$2$3`);
        }
    }

    return cellContent;
}

/**
 * 특정 테이블의 특정 행/셀 수정
 */
function modifyCellInTable(xml, tableIndex, rowIndex, cellIndex, value) {
    // 모든 테이블 분리
    const tables = [];
    const tblRegex = /(<hp:tbl[^>]*>)([\s\S]*?)(<\/hp:tbl>)/g;
    let lastIndex = 0;
    let match;

    while ((match = tblRegex.exec(xml)) !== null) {
        if (match.index > lastIndex) {
            tables.push({ type: 'other', content: xml.substring(lastIndex, match.index) });
        }
        tables.push({
            type: 'table',
            start: match[1],
            content: match[2],
            end: match[3]
        });
        lastIndex = match.index + match[0].length;
    }
    if (lastIndex < xml.length) {
        tables.push({ type: 'other', content: xml.substring(lastIndex) });
    }

    // 해당 테이블 수정
    let tblIdx = 0;
    for (let i = 0; i < tables.length; i++) {
        if (tables[i].type === 'table') {
            if (tblIdx === tableIndex) {
                tables[i].content = modifyRowInTable(tables[i].content, rowIndex, cellIndex, value);
            }
            tblIdx++;
        }
    }

    return tables.map(t => t.type === 'table' ? t.start + t.content + t.end : t.content).join('');
}

function modifyRowInTable(tableContent, targetRow, targetCell, value) {
    const rows = [];
    const trRegex = /(<hp:tr[^>]*>)([\s\S]*?)(<\/hp:tr>)/g;
    let lastIndex = 0;
    let match;

    while ((match = trRegex.exec(tableContent)) !== null) {
        if (match.index > lastIndex) {
            rows.push({ type: 'other', content: tableContent.substring(lastIndex, match.index) });
        }
        rows.push({
            type: 'row',
            start: match[1],
            content: match[2],
            end: match[3]
        });
        lastIndex = match.index + match[0].length;
    }
    if (lastIndex < tableContent.length) {
        rows.push({ type: 'other', content: tableContent.substring(lastIndex) });
    }

    let rowIdx = 0;
    for (let i = 0; i < rows.length; i++) {
        if (rows[i].type === 'row') {
            if (rowIdx === targetRow) {
                rows[i].content = modifyCellInRow(rows[i].content, targetCell, value);
            }
            rowIdx++;
        }
    }

    return rows.map(r => r.type === 'row' ? r.start + r.content + r.end : r.content).join('');
}

function modifyCellInRow(rowContent, targetCell, value) {
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

    let cellIdx = 0;
    for (let i = 0; i < cells.length; i++) {
        if (cells[i].type === 'cell') {
            if (cellIdx === targetCell) {
                cells[i].content = modifyTextInCell(cells[i].content, value);
            }
            cellIdx++;
        }
    }

    return cells.map(c => c.type === 'cell' ? c.start + c.content + c.end : c.content).join('');
}

/**
 * XML 수정 메인 함수
 */
function modifyXml(xml, personData, useBackpage) {
    let result = xml;

    // ========== 테이블 0 (1페이지) ==========

    // 신청인 정보
    result = modifyCellInTable(result, 0, 1, 2, personData.성명);
    result = modifyCellInTable(result, 0, 1, 4, personData.생년월일);
    result = modifyCellInTable(result, 0, 2, 1, personData.주소);
    result = modifyCellInTable(result, 0, 3, 1, personData.연락처);

    // 1페이지 필지 정보 (최대 12개)
    const page1Parcels = personData.필지목록.slice(0, PAGE1_PARCEL_COUNT);
    for (let i = 0; i < PAGE1_PARCEL_COUNT; i++) {
        const rowIdx = PAGE1_PARCEL_START_ROW + i;

        if (i < page1Parcels.length) {
            const parcel = page1Parcels[i];
            result = modifyCellInTable(result, 0, rowIdx, 0, parcel.읍면);
            result = modifyCellInTable(result, 0, rowIdx, 1, parcel.리);
            result = modifyCellInTable(result, 0, rowIdx, 2, parcel.지번);
            result = modifyCellInTable(result, 0, rowIdx, 6, String(parcel.벼재배면적));
        }
    }

    // 합계
    result = modifyCellInTable(result, 0, 18, 3, String(personData.합계));

    // 계좌 정보
    result = modifyCellInTable(result, 0, 19, 3, '계좌번호 미기입');
    result = modifyCellInTable(result, 0, 19, 5, personData.성명);

    // 하단 서명란
    result = result.replace(
        /신청인\s+성명\s+\(인\)/,
        `신청인       ${personData.성명}      (인)`
    );

    // ========== 테이블 1 (별지) - 필지가 12개 초과인 경우 ==========
    if (useBackpage && personData.필지목록.length > PAGE1_PARCEL_COUNT) {
        const page2Parcels = personData.필지목록.slice(PAGE1_PARCEL_COUNT, PAGE1_PARCEL_COUNT + PAGE2_PARCEL_COUNT);

        for (let i = 0; i < page2Parcels.length; i++) {
            const rowIdx = PAGE2_PARCEL_START_ROW + i;
            const parcel = page2Parcels[i];

            result = modifyCellInTable(result, 1, rowIdx, 0, parcel.읍면);
            result = modifyCellInTable(result, 1, rowIdx, 1, parcel.리);
            result = modifyCellInTable(result, 1, rowIdx, 2, parcel.지번);
            result = modifyCellInTable(result, 1, rowIdx, 6, String(parcel.벼재배면적));
        }
    }

    return result;
}

// ============================================
// HWPX 파일 생성
// ============================================

function generateHwpxFile(personData, outputPath) {
    // 필지 개수에 따라 템플릿 선택
    const useBackpage = personData.필지목록.length > PAGE1_PARCEL_COUNT;
    const templateFile = useBackpage ? TEMPLATE_MULTI : TEMPLATE_SINGLE;

    const zip = new AdmZip(templateFile);

    const sectionEntry = zip.getEntry('Contents/section0.xml');
    if (!sectionEntry) {
        console.log('  경고: section0.xml을 찾을 수 없습니다.');
        return false;
    }

    let xml = sectionEntry.getData().toString('utf8');
    xml = modifyXml(xml, personData, useBackpage);

    zip.updateFile('Contents/section0.xml', Buffer.from(xml, 'utf8'));
    zip.writeZip(outputPath);

    return useBackpage ? '별지포함' : '1페이지';
}

// ============================================
// 메인 함수
// ============================================

function main() {
    console.log('='.repeat(50));
    console.log('벼재배농가 경영안정자금 지급대상자 신청서 자동 생성 v6');
    console.log('(별지 지원 버전)');
    console.log('='.repeat(50));

    if (!fs.existsSync(OUTPUT_DIR)) {
        fs.mkdirSync(OUTPUT_DIR, { recursive: true });
    }
    console.log(`출력 디렉토리: ${OUTPUT_DIR}`);

    const dataByBusiness = readExcelData();
    const allMappings = [];

    // 처음 5명만 테스트 (강호태 포함하도록)
    const entries = Object.entries(dataByBusiness).slice(0, 5);

    console.log(`\n${entries.length}명의 신청서 생성 시작...`);

    entries.forEach(([businessId, parcels], idx) => {
        const personData = createPersonData(parcels);

        const filename = `신청서_${personData.성명}_${businessId}.hwpx`;
        const outputPath = path.join(OUTPUT_DIR, filename);

        console.log(`[${idx + 1}/${entries.length}] ${personData.성명} (${parcels.length}개 필지) 생성 중...`);

        const result = generateHwpxFile(personData, outputPath);

        if (result) {
            console.log(`  -> ${filename} 생성 완료 (${result})`);
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
