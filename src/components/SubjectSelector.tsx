/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useRef } from 'react';
import { Plus, X, Upload, ListFilter } from 'lucide-react';
import * as XLSX from 'xlsx';

interface SubjectSelectorProps {
  selectedSubjects: string[];  // Currently filtered subjects
  onSelectedSubjectsChange: (subjects: string[]) => void;
}

export default function SubjectSelector({
  selectedSubjects,
  onSelectedSubjectsChange,
}: SubjectSelectorProps) {
  const [manualInput, setManualInput] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [uploadError, setUploadError] = useState<string | null>(null);

  const handleAddManualSubject = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const trimmed = manualInput.trim();
    if (!trimmed) return;

    if (!selectedSubjects.includes(trimmed)) {
      onSelectedSubjectsChange([...selectedSubjects, trimmed]);
    }
    setManualInput('');
  };

  const handleRemoveSubject = (sub: string) => {
    onSelectedSubjectsChange(selectedSubjects.filter(item => item !== sub));
  };

  const handleUploadSubjectList = (e: React.ChangeEvent<HTMLInputElement>) => {
    setUploadError(null);
    const file = e.target.files?.[0];
    if (!file) return;

    const fileExtension = file.name.split('.').pop()?.toLowerCase();
    const isExcel = fileExtension === 'xlsx' || fileExtension === 'xls';

    if (isExcel) {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const sheetData = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
          
          const parsedSubjects: string[] = [];
          sheetData.forEach((row) => {
            if (Array.isArray(row)) {
              row.forEach((cell) => {
                if (cell !== null && cell !== undefined) {
                  const strVal = String(cell).trim();
                  if (strVal && !parsedSubjects.includes(strVal)) {
                    if (isNaN(Number(strVal))) {
                      parsedSubjects.push(strVal);
                    }
                  }
                }
              });
            }
          });

          // Filter out typical table headers that might be parsed
          const headersToIgnore = ['과목', '과목명', '과목코드', 'subject', 'subjectname', 'subject_name'];
          const filteredParsed = parsedSubjects.filter(item => !headersToIgnore.includes(item.toLowerCase()));

          const newSubjects = filteredParsed.filter(item => !selectedSubjects.includes(item));

          if (newSubjects.length === 0) {
            setUploadError('새로운 과목이 발견되지 않았습니다. 파일 구성을 확인해 주세요.');
            return;
          }

          onSelectedSubjectsChange([...selectedSubjects, ...newSubjects]);
        } catch (err) {
          setUploadError('엑셀 파일을 파싱하는 데 실패했습니다.');
        }
      };
      reader.readAsArrayBuffer(file);
    } else {
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const text = event.target?.result as string;
          if (!text) return;

          // Split by lines, commas, or semicolons
          const rawItems = text.split(/[\r\n,;]+/);
          const newSubjects = rawItems
            .map(item => item.trim())
            .filter(item => item.length > 0 && !selectedSubjects.includes(item));

          if (newSubjects.length === 0) {
            setUploadError('새로운 과목이 발견되지 않았습니다. 파일 형식을 확인하세요.');
            return;
          }

          onSelectedSubjectsChange([...selectedSubjects, ...newSubjects]);
        } catch (err) {
          setUploadError('과목 목록 파일을 파싱하는 데 실패했습니다.');
        }
      };
      reader.readAsText(file);
    }

    if (fileInputRef.current) {
      fileInputRef.current.value = ''; // reset file input
    }
  };

  return (
    <div className="bg-white rounded-2xl shadow-sm border border-slate-100 p-6 flex flex-col gap-6" id="subject-selector-panel">
      
      {/* Subject Input & File Upload Section */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* Manual Input */}
        <form onSubmit={handleAddManualSubject} className="flex flex-col gap-2">
          <label className="text-sm font-medium text-slate-700">과목 직접 추가</label>
          <div className="flex gap-2">
            <input
              type="text"
              value={manualInput}
              onChange={(e) => setManualInput(e.target.value)}
              placeholder="예: 국어, 영어 (엔터 또는 추가 클릭)"
              className="flex-1 px-4 py-2 text-sm bg-slate-50 hover:bg-slate-100/50 focus:bg-white border border-slate-200 focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100 rounded-lg outline-none transition-all"
            />
            <button
              type="submit"
              className="px-4 py-2 text-sm font-medium text-white bg-slate-800 hover:bg-slate-900 rounded-lg transition-colors flex items-center gap-1 shrink-0"
              id="btn-add-subject-manual"
            >
              <Plus className="w-4 h-4" />
              추가
            </button>
          </div>
        </form>

        {/* Text/Excel Upload */}
        <div className="flex flex-col gap-2">
          <label className="text-sm font-medium text-slate-700">과목 목록 파일 업로드</label>
          <div className="flex gap-2">
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv,.txt"
              className="hidden"
              onChange={handleUploadSubjectList}
            />
            <button
              type="button"
              onClick={() => fileInputRef.current?.click()}
              className="w-full px-4 py-2 text-sm font-medium text-slate-600 bg-slate-50 hover:bg-slate-100 border border-slate-200 hover:border-slate-300 rounded-lg transition-all flex items-center justify-center gap-2"
              id="btn-upload-subject-file"
            >
              <Upload className="w-4 h-4 text-slate-400" />
              과목 목록 파일 (.xlsx, .xls, .csv, .txt) 선택
            </button>
          </div>
          {uploadError && (
            <p className="text-xs text-rose-500 mt-0.5">{uploadError}</p>
          )}
        </div>
      </div>

      {/* Selected Subjects tags */}
      <div>
        <div className="flex justify-between items-center mb-3">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium text-slate-700">요청 과목 목록</span>
            <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700 font-bold">
              {selectedSubjects.length}개 선택됨
            </span>
          </div>
          {selectedSubjects.length > 0 && (
            <button
              onClick={() => onSelectedSubjectsChange([])}
              className="text-xs text-slate-400 hover:text-slate-600 transition-colors"
            >
              전체 비우기
            </button>
          )}
        </div>

        {selectedSubjects.length === 0 ? (
          <div className="text-center py-6 px-4 bg-slate-50/50 rounded-xl border border-dashed border-slate-150">
            <ListFilter className="w-8 h-8 text-slate-300 mx-auto mb-2" />
            <p className="text-xs text-slate-400">
              선택하거나 추가한 과목이 없습니다. 위 입력창에서 수동으로 추가하거나 목록 파일을 업로드해 주세요.
            </p>
            <p className="text-[11px] text-slate-400 mt-1">
              (과목이 없으면 업로드된 파일 전체를 기준으로 자동 분석합니다)
            </p>
          </div>
        ) : (
          <div className="flex flex-wrap gap-1.5 max-h-36 overflow-y-auto p-1 border border-slate-100 rounded-xl bg-slate-50/20">
            {selectedSubjects.map((sub) => (
              <span
                key={sub}
                className="inline-flex items-center gap-1.5 pl-3 pr-2 py-1 text-xs font-medium bg-emerald-50 text-emerald-800 rounded-full border border-emerald-100 shadow-sm"
              >
                {sub}
                <button
                  onClick={() => handleRemoveSubject(sub)}
                  className="p-0.5 hover:bg-emerald-100 rounded-full text-emerald-600 hover:text-emerald-800 transition-colors"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              </span>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}
