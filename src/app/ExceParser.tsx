'use client';

import React, { useState } from 'react';
import * as XLSX from 'xlsx';

export default function ExcelParser() {
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>('');
  const [tableData, setTableData] = useState<any[]>([]);
  const [errorMsg, setErrorMsg] = useState<string>('');
  const [workbookData, setWorkbookData] = useState<XLSX.WorkBook | null>(null);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    setErrorMsg('');
    setTableData([]);
    setSheetNames([]);
    setSelectedSheet('');

    const files = e.target.files;
    if (!files || files.length === 0) return;

    const file = files[0];

    try {
      const data = await file.arrayBuffer();
      
      // WTF: true 옵션을 주어 숨겨진 파싱 에러 방지 및 상세 로그 활성화
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, WTF: true });
      
      const sheets = workbook.SheetNames;
      setSheetNames(sheets);

      if (sheets.length > 0) {
        setWorkbookData(workbook);
        const firstSheetName = sheets[0];
        setSelectedSheet(firstSheetName);
        parseSheetData(workbook, firstSheetName);
      } else {
        setErrorMsg('엑셀 파일에서 시트를 찾을 수 없습니다.');
      }
    } catch (error: any) {
      console.error('엑셀 파일 읽기 오류:', error);
      setErrorMsg(`파일을 읽는 중 오류가 발생했습니다: ${error.message}`);
    }
  };

  const parseSheetData = (workbook: XLSX.WorkBook, sheetName: string) => {
    try {
      const sheet = workbook.Sheets[sheetName];
      if (!sheet) {
        throw new Error(`"${sheetName}" 시트를 찾을 수 없습니다.`);
      }

      // [핵심 방어 코드] !ref 범위가 없거나 잘못 지정되어 데이터가 안 읽히는 현상 강제 복구
      if (!sheet['!ref']) {
        let maxRow = 0;
        let maxCol = 0;
        // 셀 주소들을 분석해 실제 데이터가 있는 최대 범위를 수동 계산
        Object.keys(sheet).forEach((key) => {
          if (key.startsWith('!')) return;
          const coord = XLSX.utils.decode_cell(key);
          if (coord.r > maxRow) maxRow = coord.r;
          if (coord.c > maxCol) maxCol = coord.c;
        });
        sheet['!ref'] = XLSX.utils.encode_range({
          s: { r: 0, c: 0 },
          e: { r: maxRow, c: maxCol }
        });
      }

      // JSON 변환
      const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: '' });
      
      console.log(`[디버깅] "${sheetName}" 시트 파싱 성공, 데이터 개수:`, jsonData.length);
      
      if (jsonData.length === 0) {
        setErrorMsg(`"${sheetName}" 시트에 데이터가 존재하지 않거나 빈 시트입니다.`);
      }

      setTableData(jsonData);
    } catch (error: any) {
      console.error(`시트(${sheetName}) 파싱 오류:`, error);
      setErrorMsg(`시트 데이터를 변환하는 중 오류가 발생했습니다: ${error.message}`);
    }
  };

  const handleSheetChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    const sheetName = e.target.value;
    setSelectedSheet(sheetName);
    if (workbookData) {
      parseSheetData(workbookData, sheetName);
    }
  };

  return (
    <div className="p-6 max-w-6xl mx-auto bg-white rounded-xl shadow-md space-y-6">
      <h2 className="text-2xl font-bold text-gray-800">엑셀 데이터 업로드 및 분석</h2>

      <div className="border-2 border-dashed border-gray-300 p-6 rounded-lg text-center">
        <input
          type="file"
          accept=".xlsx, .xls"
          onChange={handleFileUpload}
          className="block w-full text-sm text-gray-500
            file:mr-4 file:py-2 file:px-4
            file:rounded-full file:border-0
            file:text-sm file:font-semibold
            file:bg-blue-50 file:text-blue-700
            hover:file:bg-blue-100 cursor-pointer"
        />
      </div>

      {errorMsg && (
        <div className="p-4 bg-red-50 text-red-700 rounded-md border border-red-200">
          {errorMsg}
        </div>
      )}

      {sheetNames.length > 0 && (
        <div className="flex items-center space-x-4">
          <label className="font-semibold text-gray-700">시트 선택:</label>
          <select
            value={selectedSheet}
            onChange={handleSheetChange}
            className="border border-gray-300 rounded-md px-3 py-2 focus:outline-none focus:ring-2 focus:ring-blue-500"
          >
            {sheetNames.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </div>
      )}

      {tableData.length > 0 && (
        <div className="space-y-2">
          <h3 className="text-lg font-semibold text-gray-700">
            데이터 미리보기 ({tableData.length}행)
          </h3>
          <div className="overflow-x-auto max-h-96 border border-gray-200 rounded-lg">
            <table className="min-w-full divide-y divide-gray-200 text-sm">
              <thead className="bg-gray-50 sticky top-0">
                <tr>
                  {Object.keys(tableData[0]).map((key) => (
                    <th
                      key={key}
                      className="px-4 py-3 text-left font-medium text-gray-500 uppercase tracking-wider"
                    >
                      {key}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {tableData.slice(0, 100).map((row, index) => (
                  <tr key={index} className="hover:bg-gray-50">
                    {Object.values(row).map((val: any, vIdx) => (
                      <td key={vIdx} className="px-4 py-2 whitespace-nowrap text-gray-700">
                        {String(val)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <p className="text-xs text-gray-400">* 최대 100행까지 미리보기에 표시됩니다.</p>
        </div>
      )}
    </div>
  );
}