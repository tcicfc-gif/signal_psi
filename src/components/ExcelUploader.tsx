"use client";

import React, { useState, useRef } from "react";
// Next.js Turbopack의 트리 쉐이킹 오작동으로 인한 파서 누락을 막기 위해, xlsx.mjs 번들을 직접 가리켜 임포트합니다.
// @ts-ignore
import { read, utils, set_cptable } from "xlsx/xlsx.mjs";
// @ts-ignore
import * as cptable from "xlsx/dist/cpexcel.full.mjs";
import { FileSpreadsheet } from "lucide-react";

// 브라우저 환경에서 한글 인코딩(EUC-KR / CP949 등) 디코딩을 보장하기 위해 코드페이지 테이블을 주입합니다.
set_cptable(cptable);

export interface SheetData {
  name: string;
  data: any[][];
}

interface ExcelUploaderProps {
  onDataParsed: (sheets: SheetData[], fileName: string) => void;
  onError: (errorMessage: string) => void;
}

export default function ExcelUploader({ onDataParsed, onError }: ExcelUploaderProps) {
  const [isDragging, setIsDragging] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const processFile = (file: File) => {
    const validExtensions = [".xlsx", ".xls", ".ods", ".csv"];
    const fileExtension = file.name.substring(file.name.lastIndexOf(".")).toLowerCase();

    if (!validExtensions.includes(fileExtension)) {
      onError("지원하지 않는 파일 형식입니다. .xlsx, .xls, .ods, .csv 파일만 업로드해 주세요.");
      return;
    }

    setIsLoading(true);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const arrayBuffer = e.target?.result;
        if (!arrayBuffer || !(arrayBuffer instanceof ArrayBuffer)) {
          throw new Error("파일 데이터를 정상적으로 읽지 못했습니다.");
        }

        const data = new Uint8Array(arrayBuffer);
        
        // [디버그 및 자가 진단] 파일 헤더 시그니처 (매직 넘버) 추출
        const headerBytes = data.slice(0, 8);
        const hexHeader = Array.from(headerBytes)
          .map((b) => b.toString(16).padStart(2, "0").toUpperCase())
          .join(" ");

        console.log("--- Excel Parsing Debug Log ---");
        console.log("File Name:", file.name);
        console.log("File Size:", file.size, "bytes");
        console.log("File Hex Signature (Magic Number):", hexHeader);

        // cellDates: 날짜 데이터를 JS Date 객체로 파싱
        // codepage 옵션을 949(EUC-KR / 한국어)로 추가 명시하여 레거시 인코딩 파싱 안정성을 극대화합니다.
        const workbook = read(data, { 
          type: "array", 
          cellDates: true,
          codepage: 949 
        });
        
        const sheetKeys = Object.keys(workbook.Sheets);
        console.log("SheetNames Array:", workbook.SheetNames);
        console.log("workbook.Sheets Keys:", sheetKeys);

        // 만약 Sheets 객체 자체가 완전히 비어있는 경우
        if (sheetKeys.length === 0) {
          // 파일 포맷 힌트 분석
          let formatHint = "알 수 없음 (바이너리 포맷)";
          if (hexHeader.startsWith("50 4B 03 04")) {
            formatHint = "표준 Zip 압축 파일 (.xlsx, .docx 등)";
          } else if (hexHeader.startsWith("D0 CF 11 E0")) {
            formatHint = "Microsoft OLE2 복합 문서 (.xls, .doc 등)";
          } else if (hexHeader.startsWith("EF BB BF") || hexHeader.startsWith("3C") || hexHeader.match(/^[2-7][0-9A-F]/)) {
            formatHint = "텍스트 또는 HTML/XML 텍스트 문서 (엑셀 확장자로 속여서 저장되었을 가능성 있음)";
          }

          throw new Error(
            `시트 이름(${workbook.SheetNames.join(", ")})은 인식했으나 시트 내부 데이터(Sheets)를 읽지 못했습니다. 
            [파일 상세 진단 정보]
            - 파일 크기: ${file.size.toLocaleString()} bytes
            - 파일 매직 넘버: [${hexHeader}]
            - 예상 파일 성격: ${formatHint}
            - 발견된 시트 수: ${workbook.SheetNames.length}
            (진짜 엑셀 포맷이 아니거나 텍스트 데이터를 파싱하지 못한 인코딩 오류일 수 있습니다.)`
          );
        }

        const parsedSheets: SheetData[] = [];

        workbook.SheetNames.forEach((sheetName: string) => {
          // 1. 기본 매칭 시도
          let worksheet = workbook.Sheets[sheetName];
          let matchedName = sheetName;

          // 2. 공백 제거 및 한글 자모 분리(NFD vs NFC) 인코딩 미스매칭 대응
          if (!worksheet) {
            const nfcName = sheetName.trim().normalize("NFC");
            const nfdName = sheetName.trim().normalize("NFD");
            
            const foundKey = sheetKeys.find((key) => {
              const normKeyNfc = key.trim().normalize("NFC");
              const normKeyNfd = key.trim().normalize("NFD");
              return (
                normKeyNfc === nfcName || 
                normKeyNfd === nfdName || 
                normKeyNfc === nfdName || 
                normKeyNfd === nfcName
              );
            });

            if (foundKey) {
              worksheet = workbook.Sheets[foundKey];
              matchedName = foundKey.trim().normalize("NFC");
              console.log(`Resolved sheet name mismatch with trim/normalize: "${sheetName}" mapped to key "${foundKey}"`);
            }
          }

          // 3. 인덱스 기준으로 매칭 시도
          if (!worksheet) {
            const sheetIndex = workbook.SheetNames.indexOf(sheetName);
            if (sheetIndex !== -1 && sheetKeys[sheetIndex]) {
              const fallbackKey = sheetKeys[sheetIndex];
              worksheet = workbook.Sheets[fallbackKey];
              matchedName = fallbackKey.trim().normalize("NFC");
              console.log(`Fallback mapping by index: "${sheetName}" mapped to index-matched key "${fallbackKey}"`);
            }
          }

          // 4. [극단적 폴백] 어떻게든 조회가 불가능하다면 Sheets 객체의 첫 번째 시트를 강제로 할당
          if (!worksheet && sheetKeys.length > 0) {
            const absoluteFallbackKey = sheetKeys[0];
            worksheet = workbook.Sheets[absoluteFallbackKey];
            matchedName = absoluteFallbackKey.trim().normalize("NFC");
            console.log(`Absolute fallback mapping: Using first available sheet key "${absoluteFallbackKey}"`);
          }

          if (!worksheet) {
            console.warn(`Sheet "${sheetName}" object is missing or null.`);
            return;
          }

          // [안전장치] 만약 !ref(셀 범위 정의)가 누락된 경우, 시트 내의 셀 좌표들을 스캔하여 !ref를 수동으로 설정합니다.
          if (!worksheet["!ref"]) {
            const cells = Object.keys(worksheet).filter(key => !key.startsWith("!"));
            if (cells.length > 0) {
              let minRow = Infinity, maxRow = -Infinity;
              let minCol = Infinity, maxCol = -Infinity;
              
              cells.forEach(cell => {
                const match = cell.match(/^([A-Z]+)([0-9]+)$/);
                if (match) {
                  const colStr = match[1];
                  const rowNum = parseInt(match[2], 10);
                  
                  let colNum = 0;
                  for (let i = 0; i < colStr.length; i++) {
                    colNum = colNum * 26 + (colStr.charCodeAt(i) - 64);
                  }
                  colNum = colNum - 1;

                  if (rowNum < minRow) minRow = rowNum;
                  if (rowNum > maxRow) maxRow = rowNum;
                  if (colNum < minCol) minCol = colNum;
                  if (colNum > maxCol) maxCol = colNum;
                }
              });

              if (minRow !== Infinity) {
                const startColStr = utils.encode_col(minCol);
                const endColStr = utils.encode_col(maxCol);
                worksheet["!ref"] = `${startColStr}${minRow}:${endColStr}${maxRow}`;
                console.log(`Recovered missing !ref for sheet "${matchedName}":`, worksheet["!ref"]);
              }
            }
          }

          // header: 1 옵션을 주어 2차원 배열 형태로 변환
          const sheetJson = utils.sheet_to_json<any[]>(worksheet, { header: 1 });
          
          console.log(`Parsed Sheet [${matchedName}] Row Count:`, sheetJson.length);

          // 유효 데이터 추출 (단순 배열 체크 및 1개 이상의 값 보유 여부 검증)
          const validData = sheetJson.filter((row: any) => {
            return (
              Array.isArray(row) && 
              row.length > 0 && 
              row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "")
            );
          });

          if (validData.length > 0) {
            parsedSheets.push({
              name: matchedName,
              data: validData,
            });
          } else {
            console.warn(`Sheet "${matchedName}" has no valid rows after filtering.`);
          }
        });

        if (parsedSheets.length === 0) {
          throw new Error(
            `엑셀 파일의 모든 시트가 비어있거나 유효한 테이블 데이터를 찾을 수 없습니다. (SheetNames: [${workbook.SheetNames.join(", ")}], Sheets Keys: [${sheetKeys.join(", ")}])`
          );
        }

        onDataParsed(parsedSheets, file.name);
      } catch (err: any) {
        console.error("Excel File Read/Parse Error:", err);
        onError(
          err.message || 
          "엑셀 파일을 분석하는 도중 오류가 발생했습니다. 파일 형식이 유효한지 확인해 주세요."
        );
      } finally {
        setIsLoading(false);
      }
    };

    reader.onerror = (e) => {
      console.error("FileReader Error:", e);
      onError("파일 리더(FileReader) 실행 중 에러가 발생했습니다.");
      setIsLoading(false);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
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

  const triggerFileInput = () => {
    fileInputRef.current?.click();
  };

  return (
    <div
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
      onClick={triggerFileInput}
      className={`flex flex-col items-center justify-center w-full h-64 border-2 border-dashed rounded-xl cursor-pointer transition-all duration-200 ${
        isDragging
          ? "border-blue-500 bg-blue-50/50 dark:bg-blue-950/20"
          : "border-slate-300 dark:border-slate-700 bg-slate-50 dark:bg-slate-900/50 hover:bg-slate-100 dark:hover:bg-slate-800/80"
      }`}
    >
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileChange}
        accept=".xlsx, .xls, .ods, .csv"
        className="hidden"
      />

      <div className="flex flex-col items-center justify-center pt-5 pb-6 text-center px-4">
        {isLoading ? (
          <div className="flex flex-col items-center space-y-4">
            <div className="w-12 h-12 border-4 border-blue-500 border-t-transparent rounded-full animate-spin"></div>
            <p className="text-sm text-slate-500 dark:text-slate-400 font-medium">
              엑셀 데이터를 불러오는 중...
            </p>
          </div>
        ) : (
          <>
            <div className="p-3 mb-4 rounded-full bg-blue-100 dark:bg-blue-950 text-blue-600 dark:text-blue-400">
              <FileSpreadsheet className="w-8 h-8" />
            </div>
            <p className="mb-2 text-base font-semibold text-slate-700 dark:text-slate-200">
              {isDragging ? "여기에 파일을 놓으세요!" : "클릭하거나 파일을 드래그하여 업로드하세요"}
            </p>
            <p className="text-xs text-slate-500 dark:text-slate-400">
              지원 형식: .xlsx, .xls, .ods, .csv (최대 50MB)
            </p>
          </>
        )}
      </div>
    </div>
  );
}
