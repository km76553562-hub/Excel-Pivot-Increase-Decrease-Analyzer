/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  FileSpreadsheet, 
  Trash2, 
  RefreshCw, 
  Download, 
  HelpCircle,
  FileCheck, 
  TrendingUp, 
  TrendingDown, 
  DollarSign, 
  PieChart as PieIcon, 
  ChevronRight,
  Info
} from 'lucide-react';
import { ExcelRow, PivotResult, CurrencyUnit } from './types';
import { generatePivotData, formatAmount, getUnitLabel } from './utils/analysis';
import { sampleExcelRows } from './utils/sampleData';
import FileUploader from './components/FileUploader';
import SubjectSelector from './components/SubjectSelector';

// Interactive charts or visual progress bars for financial distribution
export default function App() {
  const [excelData, setExcelData] = useState<ExcelRow[]>([]);
  const [fileName, setFileName] = useState<string>('');
  const [pendingSubjects, setPendingSubjects] = useState<string[]>([]);
  const [selectedSubjects, setSelectedSubjects] = useState<string[]>([]);
  const [currencyUnit, setCurrencyUnit] = useState<CurrencyUnit>('KRW');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [showDetailTable, setShowDetailTable] = useState<boolean>(true);
  const [sortOrder, setSortOrder] = useState<'desc' | 'asc'>('desc');

  // Detect unapplied changes between current config panel state vs executed active analysis
  const hasUnappliedChanges = useMemo(() => {
    if (pendingSubjects.length !== selectedSubjects.length) return true;
    const sortedPending = [...pendingSubjects].sort();
    const sortedSelected = [...selectedSubjects].sort();
    return sortedPending.some((val, idx) => val !== sortedSelected[idx]);
  }, [pendingSubjects, selectedSubjects]);

  // Extract all unique subjects found in the uploaded file for easy selection
  const uniqueSubjects = useMemo(() => {
    const subjects = excelData.map(row => row.subject).filter(Boolean);
    return Array.from(new Set(subjects)).sort();
  }, [excelData]);

  // Load sample data instantly
  const handleLoadSample = () => {
    setExcelData(sampleExcelRows);
    setFileName('샘플_과목별_사업_결산_데이터.xlsx');
    // Set both pending and selected so it shows instantly on load
    setPendingSubjects(['일반행정비', '시설유지비']);
    setSelectedSubjects(['일반행정비', '시설유지비']);
  };

  const handleDataLoaded = (rows: ExcelRow[], name: string) => {
    setExcelData(rows);
    setFileName(name);
    // Auto select first 3 unique subjects as a gentle helper starting point
    const firstSubjects = Array.from(new Set(rows.map(r => r.subject)))
      .filter(Boolean)
      .slice(0, 3);
    setPendingSubjects(firstSubjects);
    setSelectedSubjects(firstSubjects);
  };

  // Filter raw data based on current requested subjects & search
  const filteredRawRows = useMemo(() => {
    return excelData.filter(row => {
      // If requested subjects are configured, filter by them
      const matchesSubject = selectedSubjects.length === 0 || selectedSubjects.includes(row.subject);
      
      // Search filter for business name or codes
      const matchesSearch = !searchTerm 
        ? true 
        : (row.businessName?.toLowerCase().includes(searchTerm.toLowerCase()) || 
           row.subject?.toLowerCase().includes(searchTerm.toLowerCase()) ||
           row.businessCode?.toLowerCase().includes(searchTerm.toLowerCase()));
           
      return matchesSubject && matchesSearch;
    });
  }, [excelData, selectedSubjects, searchTerm]);

  // Generate Pivot results aggregated strictly by Business Name (사업명)
  // If multiple subjects are requested, the pivot collapses the Subject dimension as requested:
  // "사업명 기준 피벗 산출 방식 : 과목을 한번에 여러개 요청한 경우 과목을 무시하고 사업명 기준 피벗 산출"
  const pivotResults = useMemo(() => {
    return generatePivotData(excelData, selectedSubjects);
  }, [excelData, selectedSubjects]);

  // Sort Pivot results by Net Increase/Decrease ('순 증감액' = change)
  const sortedPivotResults = useMemo(() => {
    const results = [...pivotResults];
    results.sort((a, b) => {
      if (sortOrder === 'desc') {
        return b.change - a.change;
      } else {
        return a.change - b.change;
      }
    });
    return results;
  }, [pivotResults, sortOrder]);

  // Totals for top statistics panel
  const totals = useMemo(() => {
    let priorSum = 0;
    let endingSum = 0;
    let increaseSum = 0;
    let decreaseSum = 0;

    filteredRawRows.forEach(row => {
      priorSum += row.priorCarryover;
      endingSum += row.endingBalance;
      increaseSum += row.currentIncrease;
      decreaseSum += row.currentDecrease;
    });

    return {
      prior: priorSum,
      ending: endingSum,
      change: endingSum - priorSum,
      increase: increaseSum,
      decrease: decreaseSum
    };
  }, [filteredRawRows]);

  const handleExportToExcel = () => {
    if (excelData.length === 0) return;

    const wb = XLSX.utils.book_new();

    // 1. Pivot Sheet (with active sorting)
    const pivotSheetData = sortedPivotResults.map(item => ({
      '사업코드': item.businessCode,
      '사업명': item.businessName,
      '전기이월(원)': item.priorCarryover,
      '기말잔액(원)': item.endingBalance,
      '증감(기말-전기)(원)': item.change,
      '당기증가(원)': item.currentIncrease,
      '당기감소(원)': item.currentDecrease,
    }));
    const wsPivot = XLSX.utils.json_to_sheet(pivotSheetData);
    XLSX.utils.book_append_sheet(wb, wsPivot, '사업명 기준 피벗 산출');

    // 2. Filtered Detail Sheet
    const detailSheetData = filteredRawRows.map(row => ({
      '과목코드': row.subjectCode,
      '과목': row.subject,
      '사업코드': row.businessCode,
      '사업명': row.businessName,
      '전기이월(원)': row.priorCarryover,
      '당기증가(원)': row.currentIncrease,
      '당기감소(원)': row.currentDecrease,
      '기말잔액(원)': row.endingBalance,
      '증감(원)': row.endingBalance - row.priorCarryover
    }));
    const wsDetail = XLSX.utils.json_to_sheet(detailSheetData);
    XLSX.utils.book_append_sheet(wb, wsDetail, '과목별 상세 증감 산출');

    // Save Workbook
    const downloadDate = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `결산_피벗_분석_결과_${downloadDate}.xlsx`);
  };

  const handleReset = () => {
    setExcelData([]);
    setFileName('');
    setPendingSubjects([]);
    setSelectedSubjects([]);
    setSearchTerm('');
    setCurrencyUnit('KRW');
  };

  // Find the top performing businesses by Net Growth
  const topGrowthBusinesses = useMemo(() => {
    return [...pivotResults]
      .sort((a, b) => b.change - a.change)
      .slice(0, 5);
  }, [pivotResults]);

  return (
    <div className="flex flex-col min-h-screen bg-[#E4E3E0] text-[#141414] font-sans antialiased selection:bg-[#141414] selection:text-white border-4 md:border-8 border-[#141414]" id="main-app-container">
      
      {/* HEADER SECTION */}
      <header className="flex flex-col sm:flex-row items-stretch sm:items-center justify-between px-6 py-4 border-b-2 border-[#141414] bg-[#F2F1EE] gap-4" id="app-header">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 bg-[#141414] text-white flex items-center justify-center font-bold text-xl shrink-0">
            <FileSpreadsheet className="w-6 h-6 text-[#E4E3E0]" />
          </div>
          <div>
            <h1 className="text-lg md:text-xl font-extrabold tracking-tight uppercase italic font-serif">
              Excel Financial Pivot & Analyzer
            </h1>
            <p className="text-[10px] text-slate-500 font-mono tracking-wider">
              PRECISION LEVEL 100% • DOUBLE ENTRY COMPLIANT
            </p>
          </div>
        </div>

        <div className="flex items-center gap-3 self-end sm:self-auto">
          {fileName ? (
            <div className="flex items-center gap-2">
              <span className="text-[11px] font-mono text-emerald-800 bg-emerald-100/80 border border-emerald-400 px-2.5 py-1 rounded font-bold uppercase shrink-0">
                Active File
              </span>
              <span className="text-xs font-mono font-bold max-w-[180px] md:max-w-xs truncate" title={fileName}>
                {fileName}
              </span>
            </div>
          ) : (
            <div className="flex items-center gap-2">
              <span className="w-2.5 h-2.5 bg-rose-500 rounded-full animate-ping"></span>
              <span className="text-xs font-mono opacity-70">STATUS: WAITING_FOR_FILE</span>
            </div>
          )}
        </div>
      </header>

      {/* CORE WORKFLOW AREA */}
      {excelData.length === 0 ? (
        <main className="flex-1 flex flex-col items-center justify-center p-6 md:p-12 bg-white" id="empty-state">
          <div className="max-w-2xl w-full text-center space-y-8">
            <div className="space-y-3">
              <div className="inline-flex items-center gap-1.5 px-3 py-1 bg-amber-50 border border-amber-300 text-amber-900 rounded-full text-xs font-bold font-mono">
                ✨ DESIGN THEME: HIGH DENSITY
              </div>
              <h2 className="text-3xl md:text-4xl font-black tracking-tight font-serif italic text-slate-800">
                엑셀 파일 증감 분석 및 피벗 대시보드
              </h2>
              <p className="text-sm text-slate-500 max-w-lg mx-auto">
                업로드하신 과목, 사업 데이터를 기반으로 전기이월과 기말잔액을 대조하여 실시간 사업별 증감 피벗 테이블을 생성합니다.
              </p>
            </div>

            <FileUploader 
              onDataLoaded={handleDataLoaded} 
              onLoadSample={handleLoadSample} 
            />

            {/* Template Column details */}
            <div className="bg-[#F2F1EE] border border-[#141414] p-5 text-left rounded-xl max-w-xl mx-auto space-y-2">
              <h4 className="text-xs font-bold uppercase tracking-wider flex items-center gap-1.5">
                <Info className="w-4 h-4 text-slate-600" />
                정확한 매핑을 위한 권장 엑셀 구성
              </h4>
              <p className="text-xs text-slate-600 leading-relaxed">
                업로드하는 엑셀 파일은 아래의 열 이름(또는 위치)을 포함해야 합니다. 순서는 달라도 이름 기준으로 자동 탐지하여 파싱합니다.
              </p>
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 pt-2 text-[11px] font-mono">
                <span className="p-1 bg-white border border-slate-200 rounded text-center">A: 과목코드</span>
                <span className="p-1 bg-white border border-slate-200 rounded text-center">B: 과목</span>
                <span className="p-1 bg-white border border-slate-200 rounded text-center">C: 사업코드</span>
                <span className="p-1 bg-white border border-slate-200 rounded text-center">D: 사업명</span>
                <span className="p-1 bg-white border border-slate-200 rounded text-center">E: 전기이월(원)</span>
                <span className="p-1 bg-white border border-slate-200 rounded text-center">F: 당기증가(원)</span>
                <span className="p-1 bg-white border border-slate-200 rounded text-center">G: 당기감소(원)</span>
                <span className="p-1 bg-white border border-slate-200 rounded text-center">H: 기말잔액(원)</span>
              </div>
            </div>
          </div>
        </main>
      ) : (
        <main className="flex-1 flex flex-col lg:flex-row overflow-hidden" id="dashboard-active-view">
          
          {/* SIDEBAR FILTERS AND SUBJECT MANIPULATOR */}
          <aside className="w-full lg:w-80 border-b lg:border-b-0 lg:border-r-2 border-[#141414] bg-[#F2F1EE] flex flex-col p-4 gap-4 overflow-y-auto shrink-0">
            <div>
              <label className="text-[10px] font-bold uppercase tracking-widest text-slate-500 block mb-1">
                Data Context
              </label>
              <div className="p-3 bg-white border border-[#141414] rounded shadow-sm space-y-1.5">
                <div className="text-xs font-mono truncate font-bold text-slate-700">
                  📁 {fileName}
                </div>
                <div className="text-[11px] text-slate-500 font-mono flex justify-between">
                  <span>전체 레코드 수:</span>
                  <span className="font-bold text-slate-800">{excelData.length}개 행</span>
                </div>
                <button
                  onClick={handleReset}
                  className="w-full mt-2 py-1.5 bg-zinc-100 hover:bg-zinc-200 text-slate-700 border border-slate-300 text-xs font-bold uppercase tracking-wider rounded transition-colors flex items-center justify-center gap-1.5"
                >
                  <RefreshCw className="w-3.5 h-3.5" />
                  새 파일 업로드
                </button>
              </div>
            </div>

            {/* Subject Selector panel */}
            <div className="space-y-1">
              <label className="text-[10px] font-bold uppercase tracking-widest text-slate-500 block">
                Subject filtering
              </label>
              <SubjectSelector
                selectedSubjects={pendingSubjects}
                onSelectedSubjectsChange={setPendingSubjects}
              />
            </div>

            {/* Run Analysis Action Trigger */}
            <div className="p-3 bg-white border-2 border-[#141414] rounded shadow-[3px_3px_0px_0px_#141414] space-y-2">
              <div className="flex items-center justify-between">
                <span className="text-[10px] font-extrabold uppercase tracking-wider text-slate-600 font-mono">
                  Analysis Controller
                </span>
                {hasUnappliedChanges && (
                  <span className="inline-flex items-center px-1.5 py-0.5 rounded text-[9px] font-bold bg-amber-100 text-amber-800 animate-pulse border border-amber-300">
                    ⚠️ 적용 필요
                  </span>
                )}
              </div>
              <button
                onClick={() => setSelectedSubjects(pendingSubjects)}
                className={`w-full py-2.5 px-4 font-black uppercase text-xs border border-[#141414] transition-all flex items-center justify-center gap-2 cursor-pointer ${
                  hasUnappliedChanges 
                    ? 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-[2px_2px_0px_0px_#141414]'
                    : 'bg-zinc-200 text-zinc-500 shadow-none'
                }`}
                id="btn-run-analysis"
              >
                <RefreshCw className="w-4 h-4 shrink-0 animate-spin-slow" />
                증감분석 실행 (Run Analysis)
              </button>
              <p className="text-[10px] text-slate-500 leading-normal text-center font-medium">
                {hasUnappliedChanges 
                  ? '⚠️ 과목이 변경되었습니다. 위 버튼을 누르면 새로운 결과가 산출됩니다.' 
                  : '✓ 모든 변경사항이 결과에 적용되어 있습니다.'}
              </p>
            </div>

            {/* Quick tips & constraints details */}
            <div className="p-3.5 bg-amber-50 border border-amber-200 text-amber-900 rounded-xl space-y-2 text-xs">
              <div className="font-bold flex items-center gap-1">
                <HelpCircle className="w-3.5 h-3.5" />
                사업명 기준 피벗 규칙
              </div>
              <p className="leading-relaxed text-[11px] text-amber-800">
                과목을 <strong>한번에 여러개 요청한 경우</strong>, 과목을 일괄 무시하고 <strong>사업명 기준</strong>으로 합산(피벗)하여 증감을 산출하도록 설계되었습니다.
              </p>
            </div>
          </aside>

          {/* MAIN RESULTS AND PIVOT ENGINE VIEW */}
          <section className="flex-1 flex flex-col bg-white overflow-hidden">
            
            {/* KPI STATS HEADER AND CURRENCY UNIT CONTROL */}
            <div className="flex flex-col xl:flex-row items-stretch xl:items-center justify-between p-4 border-b-2 border-[#141414] bg-[#E4E3E0] gap-4" id="stats-ribbon">
              
              {/* Financial KPIs */}
              <div className="grid grid-cols-2 md:grid-cols-4 gap-4 flex-1">
                <div className="flex flex-col bg-white p-2.5 border border-[#141414] rounded shadow-sm">
                  <span className="text-[10px] uppercase font-bold text-slate-500 font-mono">
                    전기이월 합계
                  </span>
                  <span className="font-mono text-sm md:text-base font-extrabold text-[#141414]">
                    ₩{formatAmount(totals.prior, currencyUnit)}
                  </span>
                </div>
                <div className="flex flex-col bg-white p-2.5 border border-[#141414] rounded shadow-sm">
                  <span className="text-[10px] uppercase font-bold text-slate-500 font-mono">
                    기말잔액 합계
                  </span>
                  <span className="font-mono text-sm md:text-base font-extrabold text-[#141414]">
                    ₩{formatAmount(totals.ending, currencyUnit)}
                  </span>
                </div>
                <div className="flex flex-col bg-white p-2.5 border border-[#141414] rounded shadow-sm">
                  <span className="text-[10px] uppercase font-bold text-slate-500 font-mono">
                    총 순증감액
                  </span>
                  <span className={`font-mono text-sm md:text-base font-extrabold flex items-center gap-1 ${
                    totals.change >= 0 ? 'text-emerald-700' : 'text-rose-600'
                  }`}>
                    {totals.change >= 0 ? '+' : ''}
                    ₩{formatAmount(totals.change, currencyUnit)}
                    {totals.change >= 0 ? (
                      <TrendingUp className="w-4 h-4 shrink-0" />
                    ) : (
                      <TrendingDown className="w-4 h-4 shrink-0" />
                    )}
                  </span>
                </div>
                <div className="flex flex-col bg-white p-2.5 border border-[#141414] rounded shadow-sm">
                  <span className="text-[10px] uppercase font-bold text-slate-500 font-mono">
                    활성 필터 상태
                  </span>
                  <span className="font-mono text-xs font-bold text-slate-700 truncate">
                    {selectedSubjects.length === 0 ? '전체 과목 분석 중' : `${selectedSubjects.length}개 과목 적용`}
                  </span>
                </div>
              </div>

              {/* Currency Unit selection buttons */}
              <div className="flex flex-col justify-center sm:flex-row xl:flex-col gap-2 shrink-0">
                <span className="text-[10px] font-extrabold uppercase tracking-widest text-slate-600 text-center xl:text-left">
                  금액 단위 변경
                </span>
                <div className="flex border-2 border-[#141414] rounded shadow-sm overflow-hidden bg-white">
                  <button
                    onClick={() => setCurrencyUnit('KRW')}
                    className={`flex-1 px-3 py-1.5 text-xs font-bold font-mono transition-colors ${
                      currencyUnit === 'KRW'
                        ? 'bg-[#141414] text-white'
                        : 'text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    원 (KRW)
                  </button>
                  <button
                    onClick={() => setCurrencyUnit('MILLION_KRW')}
                    className={`flex-1 px-3 py-1.5 text-xs font-bold font-mono border-l-2 border-[#141414] transition-colors ${
                      currencyUnit === 'MILLION_KRW'
                        ? 'bg-[#141414] text-white'
                        : 'text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    백만 원
                  </button>
                  <button
                    onClick={() => setCurrencyUnit('HUNDRED_MILLION_KRW')}
                    className={`flex-1 px-3 py-1.5 text-xs font-bold font-mono border-l-2 border-[#141414] transition-colors ${
                      currencyUnit === 'HUNDRED_MILLION_KRW'
                        ? 'bg-[#141414] text-white'
                        : 'text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    억 원
                  </button>
                </div>
              </div>

            </div>

            {/* PRIMARY DOUBLE-DECK PANELS (UPPER: DETAIL ROWS, LOWER: PIVOT) */}
            <div className={`flex-1 flex flex-col ${showDetailTable ? 'md:grid md:grid-rows-2' : ''} overflow-hidden`}>
              
              {/* TOP PANEL: DETAILED SUBJECT RECORDS */}
              {showDetailTable && (
                <div className="flex flex-col border-b-2 border-[#141414] overflow-hidden bg-white">
                  <div className="px-4 py-2 bg-[#F2F1EE] text-[10px] font-bold uppercase tracking-widest border-b border-[#141414] flex justify-between items-center shrink-0">
                    <span className="flex items-center gap-1 text-slate-700 font-mono">
                      <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                      과목별 세부 변동 내역 (Filtered Subject Records)
                    </span>
                    <div className="flex items-center gap-3">
                      {/* Inline Search inside filtered table */}
                      <input
                        type="text"
                        placeholder="결과 내 검색..."
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                        className="px-2 py-0.5 text-[10px] bg-white border border-slate-300 focus:border-[#141414] focus:outline-none rounded"
                      />
                      <span className="opacity-60 text-[9px] font-mono hidden sm:inline">
                        {filteredRawRows.length}개 검색됨 / 전체 {excelData.length}개
                      </span>
                      <button
                        onClick={() => setShowDetailTable(false)}
                        className="px-2 py-0.5 bg-white hover:bg-zinc-100 border border-[#141414] text-[9px] font-bold uppercase transition-colors rounded cursor-pointer shrink-0"
                        id="btn-hide-details"
                      >
                        감추기
                      </button>
                    </div>
                  </div>

                  <div className="flex-1 overflow-auto font-mono text-[11px] relative">
                    <table className="w-full border-collapse text-left">
                      <thead className="sticky top-0 bg-white shadow-sm z-10 border-b border-slate-200">
                        <tr className="bg-slate-50">
                          <th className="p-2 border-r border-slate-200 font-semibold text-slate-600">과목코드</th>
                          <th className="p-2 border-r border-slate-200 font-semibold text-slate-600">과목명</th>
                          <th className="p-2 border-r border-slate-200 font-semibold text-slate-600">사업코드</th>
                          <th className="p-2 border-r border-slate-200 font-semibold text-slate-600">사업명</th>
                          <th className="p-2 border-r border-slate-200 text-right font-semibold text-slate-600">전기이월 ({getUnitLabel(currencyUnit)})</th>
                          <th className="p-2 border-r border-slate-200 text-right font-semibold text-slate-600">당기증가 ({getUnitLabel(currencyUnit)})</th>
                          <th className="p-2 border-r border-slate-200 text-right font-semibold text-slate-600">당기감소 ({getUnitLabel(currencyUnit)})</th>
                          <th className="p-2 border-r border-slate-200 text-right font-semibold text-slate-600">기말잔액 ({getUnitLabel(currencyUnit)})</th>
                          <th className="p-2 text-right font-semibold text-slate-600">실증감액 ({getUnitLabel(currencyUnit)})</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredRawRows.length === 0 ? (
                          <tr>
                            <td colSpan={9} className="text-center p-12 text-slate-400 font-sans">
                              조건에 일치하는 데이터가 존재하지 않습니다. 과목 필터를 해제하거나 검색어를 변경해 보세요.
                            </td>
                          </tr>
                        ) : (
                          filteredRawRows.map((row, index) => {
                            const delta = row.endingBalance - row.priorCarryover;
                            return (
                              <tr key={index} className="hover:bg-slate-50/80 transition-colors">
                                <td className="p-2 border-r border-slate-100 font-medium text-slate-800">{row.subjectCode || '-'}</td>
                                <td className="p-2 border-r border-slate-100 text-slate-950 font-bold">{row.subject}</td>
                                <td className="p-2 border-r border-slate-100 text-slate-500">{row.businessCode || '-'}</td>
                                <td className="p-2 border-r border-slate-100 text-slate-700 font-medium max-w-[200px] truncate" title={row.businessName}>
                                  {row.businessName}
                                </td>
                                <td className="p-2 border-r border-slate-100 text-right">{formatAmount(row.priorCarryover, currencyUnit)}</td>
                                <td className="p-2 border-r border-slate-100 text-right text-emerald-600 bg-emerald-50/10">{formatAmount(row.currentIncrease, currencyUnit)}</td>
                                <td className="p-2 border-r border-slate-100 text-right text-rose-600 bg-rose-50/10">{formatAmount(row.currentDecrease, currencyUnit)}</td>
                                <td className="p-2 border-r border-slate-100 text-right font-bold">{formatAmount(row.endingBalance, currencyUnit)}</td>
                                <td className={`p-2 text-right font-bold ${
                                  delta >= 0 ? 'text-emerald-700 bg-emerald-50/20' : 'text-rose-600 bg-rose-50/20'
                                }`}>
                                  {delta >= 0 ? '+' : ''}
                                  {formatAmount(delta, currencyUnit)}
                                </td>
                              </tr>
                            );
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* BOTTOM PANEL: AGGREGATED BUSINESS UNIT PIVOT ANALYSIS */}
              <div className="flex flex-col bg-[#F9F9F7] overflow-hidden">
                <div className="px-4 py-2 bg-[#F2F1EE] text-[10px] font-bold uppercase tracking-widest border-b border-[#141414] flex flex-wrap gap-2 items-center justify-between shrink-0">
                  <div className="flex items-center gap-2">
                    <span className="flex items-center gap-1 text-slate-700 font-mono">
                      <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                      사업명 기준 증감 피벗 분석 (Business Unit Pivot Analysis)
                    </span>
                    {!showDetailTable && (
                      <button
                        onClick={() => setShowDetailTable(true)}
                        className="px-2 py-0.5 bg-white hover:bg-zinc-100 border border-[#141414] text-[9px] font-bold uppercase text-emerald-700 transition-colors rounded shadow-sm cursor-pointer"
                        id="btn-show-details"
                      >
                        + 세부내역 보이기
                      </button>
                    )}
                  </div>
                  
                  {/* Sorting controls */}
                  <div className="flex items-center gap-2">
                    <span className="text-[10px] text-slate-500 font-bold uppercase">순증감 정렬:</span>
                    <div className="flex border border-[#141414] rounded overflow-hidden bg-white shadow-sm">
                      <button
                        onClick={() => setSortOrder('desc')}
                        className={`px-2 py-0.5 text-[9px] font-bold uppercase transition-colors cursor-pointer ${
                          sortOrder === 'desc'
                            ? 'bg-[#141414] text-white font-black'
                            : 'text-slate-700 hover:bg-slate-100'
                        }`}
                        id="btn-sort-desc"
                      >
                        내림차순 (▼)
                      </button>
                      <button
                        onClick={() => setSortOrder('asc')}
                        className={`px-2 py-0.5 text-[9px] font-bold uppercase border-l border-[#141414] transition-colors cursor-pointer ${
                          sortOrder === 'asc'
                            ? 'bg-[#141414] text-white font-black'
                            : 'text-slate-700 hover:bg-slate-100'
                        }`}
                        id="btn-sort-asc"
                      >
                        오름차순 (▲)
                      </button>
                    </div>
                  </div>
                </div>

                <div className="flex-1 overflow-auto font-mono text-[11px]">
                  <table className="w-full border-collapse text-left">
                    <thead className="sticky top-0 bg-[#E4E3E0] z-10 border-b border-[#141414]">
                      <tr>
                        <th className="p-2 border-r border-slate-300 font-bold text-slate-700">사업코드 (Aggregated)</th>
                        <th className="p-2 border-r border-slate-300 font-bold text-slate-700">사업명 (Pivot Key)</th>
                        <th className="p-2 border-r border-slate-300 text-right font-bold text-slate-700">전기이월 합계 ({getUnitLabel(currencyUnit)})</th>
                        <th className="p-2 border-r border-slate-300 text-right font-bold text-slate-700">기말잔액 합계 ({getUnitLabel(currencyUnit)})</th>
                        <th className="p-2 text-right font-bold text-slate-700">순 증감액 (기말 - 전기) ({getUnitLabel(currencyUnit)})</th>
                      </tr>
                    </thead>
                    <tbody className="bg-white divide-y divide-slate-200">
                      {sortedPivotResults.length === 0 ? (
                        <tr>
                          <td colSpan={5} className="text-center p-12 text-slate-400 font-sans bg-white">
                            피벗 테이블 생성 대상 사업이 존재하지 않습니다.
                          </td>
                        </tr>
                      ) : (
                        <>
                          {sortedPivotResults.map((pivot, index) => {
                            return (
                              <tr key={index} className="hover:bg-slate-50 transition-colors">
                                <td className="p-2 border-r border-slate-200 text-slate-500 font-semibold">{pivot.businessCode || '-'}</td>
                                <td className="p-2 border-r border-slate-200 text-slate-900 font-bold flex items-center gap-2">
                                  <span className="w-1.5 h-1.5 bg-[#141414] rounded-full"></span>
                                  {pivot.businessName}
                                </td>
                                <td className="p-2 border-r border-slate-200 text-right text-slate-700 font-medium">
                                  {formatAmount(pivot.priorCarryover, currencyUnit)}
                                </td>
                                <td className="p-2 border-r border-slate-200 text-right text-slate-950 font-bold">
                                  {formatAmount(pivot.endingBalance, currencyUnit)}
                                </td>
                                <td className={`p-2 text-right font-extrabold ${
                                  pivot.change >= 0 ? 'text-emerald-700 bg-emerald-50/25' : 'text-rose-600 bg-rose-50/25'
                                }`}>
                                  <div className="flex items-center justify-end gap-1">
                                    {pivot.change >= 0 ? '+' : ''}
                                    {formatAmount(pivot.change, currencyUnit)}
                                    {pivot.change >= 0 ? (
                                      <TrendingUp className="w-3 h-3 text-emerald-600 shrink-0" />
                                    ) : (
                                      <TrendingDown className="w-3 h-3 text-rose-500 shrink-0" />
                                    )}
                                  </div>
                                </td>
                              </tr>
                            );
                          })}
                          {/* GRAND TOTAL ROW */}
                          <tr className="bg-slate-100/90 font-black border-t-2 border-[#141414]">
                            <td className="p-2 border-r border-slate-300 text-slate-500 font-bold">TOTAL</td>
                            <td className="p-2 border-r border-slate-300 text-slate-900 font-extrabold">합계 (Grand Total)</td>
                            <td className="p-2 border-r border-slate-300 text-right text-slate-800">
                              {formatAmount(totals.prior, currencyUnit)}
                            </td>
                            <td className="p-2 border-r border-slate-300 text-right text-slate-950">
                              {formatAmount(totals.ending, currencyUnit)}
                            </td>
                            <td className={`p-2 text-right text-sm font-black ${
                              totals.change >= 0 ? 'text-emerald-800 bg-emerald-100/30' : 'text-rose-700 bg-rose-100/30'
                            }`}>
                              {totals.change >= 0 ? '+' : ''}
                              {formatAmount(totals.change, currencyUnit)}
                            </td>
                          </tr>
                        </>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>

            {/* LOWER STATS SUMMARY DASHBOARD CHIPS (Top Businesses by growth) */}
            {pivotResults.length > 0 && (
              <div className="p-4 bg-slate-50 border-t border-slate-200 grid grid-cols-1 md:grid-cols-2 gap-4 shrink-0">
                <div>
                  <h4 className="text-xs font-bold text-slate-600 uppercase mb-2 flex items-center gap-1.5 font-mono">
                    <TrendingUp className="w-4 h-4 text-emerald-600" />
                    순증가 상위 사업 (Top Growth Units)
                  </h4>
                  <div className="space-y-1 text-xs">
                    {topGrowthBusinesses.slice(0, 3).map((item, i) => (
                      <div key={i} className="flex justify-between items-center p-2 bg-white border border-slate-200 rounded">
                        <span className="truncate max-w-[200px] font-medium text-slate-700">{item.businessName}</span>
                        <span className="font-mono text-emerald-700 font-bold bg-emerald-50 px-1.5 py-0.5 rounded">
                          +{formatAmount(item.change, currencyUnit)} {getUnitLabel(currencyUnit)}
                        </span>
                      </div>
                    ))}
                  </div>
                </div>

                <div>
                  <h4 className="text-xs font-bold text-slate-600 uppercase mb-2 flex items-center gap-1.5 font-mono">
                    <PieIcon className="w-4 h-4 text-slate-600" />
                    사업별 기말 비중 대비 (Ending Balance Proportions)
                  </h4>
                  <div className="space-y-1.5">
                    {pivotResults.slice(0, 3).map((item, i) => {
                      const totalEnding = totals.ending || 1;
                      const percentage = Math.max(0, Math.min(100, (item.endingBalance / totalEnding) * 100));
                      return (
                        <div key={i} className="space-y-0.5">
                          <div className="flex justify-between text-xs font-mono text-slate-600">
                            <span className="truncate max-w-[180px]">{item.businessName}</span>
                            <span>{percentage.toFixed(1)}%</span>
                          </div>
                          <div className="w-full bg-slate-200 h-2 rounded-full overflow-hidden border border-slate-300">
                            <div 
                              className="bg-[#141414] h-full" 
                              style={{ width: `${percentage}%` }}
                            />
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            )}

            {/* DASHBOARD ACTION FOOTER */}
            <footer className="px-6 py-3 border-t-2 border-[#141414] bg-[#F2F1EE] flex flex-col sm:flex-row items-center justify-between gap-3 shrink-0">
              <div className="text-[10px] font-mono text-slate-500 text-center sm:text-left">
                Calculation Engine: Financial Pivot v2.1 (Precision Mode) • Powered by SheetJS
              </div>
              <div className="flex gap-3 w-full sm:w-auto">
                <button 
                  onClick={handleReset}
                  className="flex-1 sm:flex-none px-5 py-2 bg-white text-slate-700 border-2 border-[#141414] text-xs font-black uppercase hover:bg-zinc-100 transition-colors"
                >
                  초기화 (Reset)
                </button>
                <button 
                  onClick={handleExportToExcel}
                  className="flex-1 sm:flex-none px-6 py-2 bg-[#141414] hover:bg-zinc-800 text-white text-xs font-black uppercase flex items-center justify-center gap-2 shadow-[2px_2px_0px_0px_#888888] transition-all"
                  id="btn-export-excel"
                >
                  <Download className="w-4 h-4 text-emerald-400" />
                  Excel 파일 다운로드
                </button>
              </div>
            </footer>

          </section>
        </main>
      )}

    </div>
  );
}
