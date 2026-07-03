/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, AlertCircle, Sparkles } from 'lucide-react';
import { ExcelRow } from '../types';
import { parseUploadedExcel } from '../utils/analysis';

interface FileUploaderProps {
  onDataLoaded: (rows: ExcelRow[], fileName: string) => void;
  onLoadSample: () => void;
}

export default function FileUploader({ onDataLoaded, onLoadSample }: FileUploaderProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processFile = (file: File) => {
    setError(null);
    const fileName = file.name;
    const fileExtension = fileName.split('.').pop()?.toLowerCase();
    
    if (fileExtension !== 'xlsx' && fileExtension !== 'xls') {
      setError('엑셀 파일(.xlsx, .xls)만 업로드할 수 있습니다.');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        
        if (workbook.SheetNames.length === 0) {
          throw new Error('엑셀 파일에 시트가 존재하지 않습니다.');
        }

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert sheet to 2D array representing raw rows and cells
        const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1 });
        
        if (rawRows.length < 2) {
          throw new Error('데이터가 부족하거나 비어 있는 엑셀 파일입니다.');
        }

        const parsed = parseUploadedExcel(rawRows);
        
        if (parsed.length === 0) {
          throw new Error('데이터를 분석하지 못했습니다. 열 구성이 올바른지 확인해 주세요.\n(과목코드, 과목, 사업코드, 사업명, 전기이월(원), 당기증가(원), 당기감소(원), 기말잔액(원))');
        }

        onDataLoaded(parsed, fileName);
      } catch (err: any) {
        setError(err.message || '엑셀 파일을 처리하는 동안 오류가 발생했습니다.');
      }
    };

    reader.onerror = () => {
      setError('파일을 읽는 중 에러가 발생했습니다.');
    };

    reader.readAsArrayBuffer(file);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFile(e.dataTransfer.files[0]);
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      processFile(e.target.files[0]);
    }
  };

  return (
    <div className="w-full max-w-2xl mx-auto" id="file-uploader-section">
      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={() => fileInputRef.current?.click()}
        className={`relative flex flex-col items-center justify-center border-2 border-dashed rounded-2xl p-10 cursor-pointer transition-all duration-300 ${
          isDragging
            ? 'border-emerald-500 bg-emerald-50/50 shadow-inner'
            : 'border-slate-300 bg-white hover:border-slate-400 hover:shadow-md'
        }`}
      >
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx, .xls"
          className="hidden"
          onChange={handleFileChange}
        />
        
        <div className="p-4 rounded-full bg-slate-50 text-slate-500 mb-4 group-hover:scale-110 transition-transform">
          <Upload className={`w-10 h-10 ${isDragging ? 'text-emerald-500 animate-bounce' : 'text-slate-400'}`} />
        </div>

        <h3 className="text-lg font-medium text-slate-800 mb-1">
          분석할 엑셀 파일을 업로드하세요
        </h3>
        <p className="text-sm text-slate-500 text-center mb-6 max-w-md">
          과목코드, 과목, 사업코드, 사업명, 전기이월(원), 당기증가(원), 당기감소(원), 기말잔액(원)
          형식의 시트를 드래그하거나 여기를 클릭하여 선택하세요.
        </p>

        <div className="flex flex-wrap gap-2 justify-center">
          <span className="inline-flex items-center gap-1 text-xs px-2.5 py-1 rounded bg-slate-100 text-slate-600">
            <FileSpreadsheet className="w-3.5 h-3.5" />
            Excel (.xlsx, .xls) 지원
          </span>
        </div>
      </div>

      {error && (
        <div className="mt-4 p-4 rounded-xl bg-rose-50 border border-rose-100 flex items-start gap-3 text-rose-800 text-sm">
          <AlertCircle className="w-5 h-5 shrink-0 mt-0.5 text-rose-500" />
          <div className="whitespace-pre-line">{error}</div>
        </div>
      )}

      <div className="mt-6 flex flex-col items-center gap-3">
        <div className="text-xs text-slate-400">또는 데이터가 준비되지 않으셨나요?</div>
        <button
          onClick={onLoadSample}
          className="inline-flex items-center gap-2 px-4 py-2 text-sm font-medium text-emerald-600 bg-emerald-50 hover:bg-emerald-100 hover:text-emerald-700 rounded-lg transition-colors border border-emerald-200"
          id="btn-load-sample"
        >
          <Sparkles className="w-4 h-4" />
          샘플 데이터로 즉시 시작하기
        </button>
      </div>
    </div>
  );
}
