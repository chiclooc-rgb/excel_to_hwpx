/**
 * HWPX XML 테이블 구조 상세 분석
 */

const fs = require('fs');
const path = require('path');

const xmlFile = path.join(__dirname, 'output', 'Form_extracted', 'Contents', 'section0.xml');
const content = fs.readFileSync(xmlFile, 'utf8');

// 테이블 내 행과 셀 구조 분석
console.log('=== 테이블 구조 분석 ===\n');

// hp:tr (행) 찾기
const trRegex = /<hp:tr[^>]*>([\s\S]*?)<\/hp:tr>/g;
let trMatch;
let rowIndex = 0;

while ((trMatch = trRegex.exec(content)) !== null) {
    const rowContent = trMatch[1];

    // 해당 행의 셀들 찾기
    const tcRegex = /<hp:tc[^>]*>([\s\S]*?)<\/hp:tc>/g;
    let tcMatch;
    let cellIndex = 0;

    console.log(`=== Row ${rowIndex} ===`);

    while ((tcMatch = tcRegex.exec(rowContent)) !== null) {
        const cellContent = tcMatch[1];

        // 셀 내의 텍스트 찾기
        const tRegex = /<hp:t[^>]*>([^<]*)<\/hp:t>/g;
        let tMatch;
        let texts = [];

        while ((tMatch = tRegex.exec(cellContent)) !== null) {
            texts.push(tMatch[1]);
        }

        const textStr = texts.join(' | ');
        if (textStr.trim()) {
            console.log(`  Cell ${cellIndex}: "${textStr}"`);
        } else {
            console.log(`  Cell ${cellIndex}: (빈 셀)`);
        }

        cellIndex++;
    }

    rowIndex++;
    console.log('');
}
