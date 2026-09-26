"use client";

import React, { useState, useMemo } from "react";
import { ChevronLeft, ChevronRight, Search } from "lucide-react";

interface DataTableProps {
  data: any[][];
}

export default function DataTable({ data }: DataTableProps) {
  const [searchTerm, setSearchTerm] = useState("");
  const [currentPage, setCurrentPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);

  // 첫 번째 행은 헤더로 간주, 나머지는 실제 데이터 행
  const headers = useMemo(() => {
    if (!data || data.length === 0) return [];
    return data[0] || [];
  }, [data]);

  const rawRows = useMemo(() => {
    if (!data || data.length <= 1) return [];
    return data.slice(1);
  }, [data]);

  // 검색 필터링 적용
  const filteredRows = useMemo(() => {
    if (!searchTerm.trim()) return rawRows;
    const term = searchTerm.toLowerCase();

    return rawRows.filter((row) =>
      row.some((cell) => {
        if (cell === null || cell === undefined) return false;
        return String(cell).toLowerCase().includes(term);
      })
    );
  }, [rawRows, searchTerm]);

  // 페이지네이션 계산
  const totalRows = filteredRows.length;
  const totalPages = Math.max(1, Math.ceil(totalRows / pageSize));

  // 현재 페이지에 해당하는 데이터
  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    const end = start + pageSize;
    return filteredRows.slice(start, end);
  }, [filteredRows, currentPage, pageSize]);

  // 페이지 변경 핸들러
  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };

  // 검색어 변경 시 페이지 초기화
  const handleSearchChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setSearchTerm(e.target.value);
    setCurrentPage(1);
  };

  // 페이지 크기 변경 시 페이지 초기화
  const handlePageSizeChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
    setPageSize(Number(e.target.value));
    setCurrentPage(1);
  };

  if (data.length === 0) {
    return (
      <div className="text-center py-8 text-slate-500 dark:text-slate-400">
        표시할 데이터가 없습니다.
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* 검색 및 제어 도구 */}
      <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3 bg-white dark:bg-slate-900 p-4 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800">
        <div className="relative flex-1 max-w-md">
          <span className="absolute inset-y-0 left-0 flex items-center pl-3 pointer-events-none text-slate-400">
            <Search className="w-5 h-5" />
          </span>
          <input
            type="text"
            placeholder="전체 데이터 검색..."
            value={searchTerm}
            onChange={handleSearchChange}
            className="w-full pl-10 pr-4 py-2 border border-slate-300 dark:border-slate-700 rounded-lg bg-slate-50 dark:bg-slate-950 text-slate-800 dark:text-slate-100 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 text-sm transition-all"
          />
        </div>

        <div className="flex items-center space-x-2 text-sm text-slate-600 dark:text-slate-300">
          <span>보기 설정:</span>
          <select
            value={pageSize}
            onChange={handlePageSizeChange}
            className="border border-slate-300 dark:border-slate-700 rounded px-2 py-1 bg-white dark:bg-slate-950 text-slate-800 dark:text-slate-100 focus:outline-none focus:ring-1 focus:ring-blue-500"
          >
            <option value={10}>10개씩</option>
            <option value={25}>25개씩</option>
            <option value={50}>50개씩</option>
            <option value={100}>100개씩</option>
          </select>
          <span className="hidden sm:inline">
            (총 {totalRows.toLocaleString()}개 데이터 행)
          </span>
        </div>
      </div>

      {/* 테이블 영역 */}
      <div className="bg-white dark:bg-slate-900 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800 overflow-hidden">
        <div className="overflow-x-auto max-h-[600px] overflow-y-auto">
          <table className="w-full text-left border-collapse">
            <thead className="bg-slate-50 dark:bg-slate-800/80 sticky top-0 z-10 border-b border-slate-200 dark:border-slate-700">
              <tr>
                <th className="px-6 py-3.5 text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider text-center w-12 bg-slate-50 dark:bg-slate-800">
                  No.
                </th>
                {headers.map((header, index) => (
                  <th
                    key={index}
                    className="px-6 py-3.5 text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase tracking-wider bg-slate-50 dark:bg-slate-800"
                  >
                    {header !== undefined && header !== null ? String(header) : `열 ${index + 1}`}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-800">
              {paginatedRows.length > 0 ? (
                paginatedRows.map((row, rowIndex) => {
                  const absoluteIndex = (currentPage - 1) * pageSize + rowIndex + 1;
                  return (
                    <tr
                      key={rowIndex}
                      className="hover:bg-slate-50/80 dark:hover:bg-slate-800/40 transition-colors"
                    >
                      <td className="px-6 py-3 text-sm text-slate-400 dark:text-slate-500 text-center font-medium">
                        {absoluteIndex}
                      </td>
                      {headers.map((_, colIndex) => {
                        const cellValue = row[colIndex];
                        return (
                          <td
                            key={colIndex}
                            className="px-6 py-3 text-sm text-slate-700 dark:text-slate-300 max-w-xs truncate"
                            title={cellValue !== undefined ? String(cellValue) : ""}
                          >
                            {cellValue !== undefined && cellValue !== null ? String(cellValue) : ""}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })
              ) : (
                <tr>
                  <td
                    colSpan={headers.length + 1}
                    className="px-6 py-10 text-center text-sm text-slate-400 dark:text-slate-500"
                  >
                    검색 조건에 맞는 데이터가 없습니다.
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* 페이지네이션 컨트롤 */}
      {totalPages > 1 && (
        <div className="flex items-center justify-between bg-white dark:bg-slate-900 p-4 rounded-xl shadow-sm border border-slate-200 dark:border-slate-800">
          <div className="text-sm text-slate-500 dark:text-slate-400">
            총 {totalRows.toLocaleString()}개 중{" "}
            <span className="font-semibold text-slate-700 dark:text-slate-200">
              {Math.min(totalRows, (currentPage - 1) * pageSize + 1)}
            </span>
            -
            <span className="font-semibold text-slate-700 dark:text-slate-200">
              {Math.min(totalRows, currentPage * pageSize)}
            </span>
            표시
          </div>

          <div className="flex items-center space-x-1">
            <button
              onClick={() => handlePageChange(1)}
              disabled={currentPage === 1}
              className="p-2 border border-slate-300 dark:border-slate-700 rounded-lg text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-xs transition-colors"
            >
              처음
            </button>
            <button
              onClick={() => handlePageChange(currentPage - 1)}
              disabled={currentPage === 1}
              className="p-2 border border-slate-300 dark:border-slate-700 rounded-lg text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            
            <span className="px-4 text-sm font-medium text-slate-700 dark:text-slate-300">
              {currentPage} / {totalPages}
            </span>

            <button
              onClick={() => handlePageChange(currentPage + 1)}
              disabled={currentPage === totalPages}
              className="p-2 border border-slate-300 dark:border-slate-700 rounded-lg text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
            <button
              onClick={() => handlePageChange(totalPages)}
              disabled={currentPage === totalPages}
              className="p-2 border border-slate-300 dark:border-slate-700 rounded-lg text-slate-600 dark:text-slate-300 hover:bg-slate-50 dark:hover:bg-slate-800 disabled:opacity-40 disabled:cursor-not-allowed text-xs transition-colors"
            >
              끝
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
