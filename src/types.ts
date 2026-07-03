/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

export interface ExcelRow {
  subjectCode: string;     // 과목코드 (Column A)
  subject: string;         // 과목 (Column B)
  businessCode: string;    // 사업코드 (Column C)
  businessName: string;    // 사업명 (Column D)
  priorCarryover: number;  // 전기이월(원) (Column E)
  currentIncrease: number; // 당기증가(원) (Column F)
  currentDecrease: number; // 당기감소(원) (Column G)
  endingBalance: number;   // 기말잔액(원) (Column H)
}

export interface PivotResult {
  businessCode: string;
  businessName: string;
  priorCarryover: number;  // 전기이월(원)
  endingBalance: number;   // 기말잔액(원)
  change: number;          // 증감 (기말잔액 - 전기이월)
  currentIncrease: number; // 당기증가 (for reference / detail)
  currentDecrease: number; // 당기감소 (for reference / detail)
}

export type CurrencyUnit = 'KRW' | 'MILLION_KRW' | 'HUNDRED_MILLION_KRW';

export interface SubjectRequest {
  id: string;
  name: string;
}
