/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { ExcelRow, PivotResult, CurrencyUnit } from '../types';

/**
 * Parses raw sheet data from XLSX into structured ExcelRow items.
 * Robustly detects headers and filters out sub-total/total rows.
 */
export function parseUploadedExcel(sheetData: any[][]): ExcelRow[] {
  let headerIndex = -1;
  for (let i = 0; i < Math.min(sheetData.length, 15); i++) {
    const row = sheetData[i];
    if (!row) continue;
    const rowStr = row.map(cell => String(cell || '').trim()).join(' ');
    
    // Look for signature keywords of the columns
    if (
      (rowStr.includes('과목') || rowStr.includes('코드')) &&
      rowStr.includes('사업') &&
      (rowStr.includes('이월') || rowStr.includes('잔액'))
    ) {
      headerIndex = i;
      break;
    }
  }

  // Fallback to row 0 if no clear header was detected
  if (headerIndex === -1) {
    headerIndex = 0;
  }

  const headerRow = sheetData[headerIndex] || [];
  
  // Default positions (A-H)
  const colIndices = {
    subjectCode: 0,
    subject: 1,
    businessCode: 2,
    businessName: 3,
    priorCarryover: 4,
    currentIncrease: 5,
    currentDecrease: 6,
    endingBalance: 7
  };

  // Dynamically map columns based on exact text matches
  headerRow.forEach((cell, idx) => {
    const s = String(cell || '').replace(/\s+/g, '');
    if (s === '과목코드') colIndices.subjectCode = idx;
    else if (s === '과목') colIndices.subject = idx;
    else if (s === '사업코드') colIndices.businessCode = idx;
    else if (s === '사업명') colIndices.businessName = idx;
    else if (s.includes('전기이월')) colIndices.priorCarryover = idx;
    else if (s.includes('당기증가')) colIndices.currentIncrease = idx;
    else if (s.includes('당기감소')) colIndices.currentDecrease = idx;
    else if (s.includes('기말잔액')) colIndices.endingBalance = idx;
  });

  const parsedRows: ExcelRow[] = [];
  for (let i = headerIndex + 1; i < sheetData.length; i++) {
    const row = sheetData[i];
    if (!row || row.length === 0) continue;

    const subCode = String(row[colIndices.subjectCode] ?? '').trim();
    const subName = String(row[colIndices.subject] ?? '').trim();
    const bizCode = String(row[colIndices.businessCode] ?? '').trim();
    const bizName = String(row[colIndices.businessName] ?? '').trim();

    // Skip fully empty rows
    if (!subCode && !subName && !bizCode && !bizName) continue;

    // Filter out rows representing summaries/totals
    const isSummaryRow = 
      subName.includes('합계') || subName.includes('소계') || subName.includes('총계') || subName === '계' ||
      bizName.includes('합계') || bizName.includes('소계') || bizName.includes('총계') || bizName === '계' ||
      subCode.includes('합계') || bizCode.includes('합계') ||
      (!subCode && !bizCode && (subName.includes('합') || bizName.includes('합')));

    if (isSummaryRow) continue;

    const parseNum = (val: any): number => {
      if (val === null || val === undefined) return 0;
      if (typeof val === 'number') return val;
      const clean = String(val).replace(/,/g, '').trim();
      const num = parseFloat(clean);
      return isNaN(num) ? 0 : num;
    };

    parsedRows.push({
      subjectCode: subCode,
      subject: subName,
      businessCode: bizCode,
      businessName: bizName,
      priorCarryover: parseNum(row[colIndices.priorCarryover]),
      currentIncrease: parseNum(row[colIndices.currentIncrease]),
      currentDecrease: parseNum(row[colIndices.currentDecrease]),
      endingBalance: parseNum(row[colIndices.endingBalance]),
    });
  }

  return parsedRows;
}

/**
 * Builds the Pivot table on "Business Name" for the specified subjects.
 * If multiple subjects are requested, it collapses the "Subject" dimension
 * and aggregates priorCarryover, endingBalance, and calculates change (ending - prior).
 */
export function generatePivotData(rows: ExcelRow[], targetSubjects: string[]): PivotResult[] {
  // If no specific subjects are requested, we pivot on the entire dataset
  const filteredRows = targetSubjects.length > 0 
    ? rows.filter(row => targetSubjects.includes(row.subject))
    : rows;

  const grouped: { [key: string]: { bizCode: string; prior: number; ending: number; inc: number; dec: number } } = {};

  filteredRows.forEach(row => {
    const key = row.businessName || '기타사업';
    if (!grouped[key]) {
      grouped[key] = {
        bizCode: row.businessCode || '',
        prior: 0,
        ending: 0,
        inc: 0,
        dec: 0
      };
    }
    grouped[key].prior += row.priorCarryover;
    grouped[key].ending += row.endingBalance;
    grouped[key].inc += row.currentIncrease;
    grouped[key].dec += row.currentDecrease;
  });

  return Object.keys(grouped).map(businessName => {
    const g = grouped[businessName];
    return {
      businessCode: g.bizCode,
      businessName,
      priorCarryover: g.prior,
      endingBalance: g.ending,
      change: g.ending - g.prior, // 증감 = 기말잔액 - 전기이월
      currentIncrease: g.inc,
      currentDecrease: g.dec,
    };
  });
}

/**
 * Formats standard KRW amount according to selected CurrencyUnit
 */
export function formatAmount(value: number, unit: CurrencyUnit): string {
  let converted = value;
  if (unit === 'MILLION_KRW') {
    converted = value / 1000000;
  } else if (unit === 'HUNDRED_MILLION_KRW') {
    converted = value / 100000000;
  }

  // Use commas, keep 0 decimal for all units as requested
  const formatted = converted.toLocaleString('ko-KR', {
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  });

  return formatted;
}

/**
 * Returns currency unit label
 */
export function getUnitLabel(unit: CurrencyUnit): string {
  switch (unit) {
    case 'MILLION_KRW':
      return '백만원';
    case 'HUNDRED_MILLION_KRW':
      return '억원';
    case 'KRW':
    default:
      return '원';
  }
}
