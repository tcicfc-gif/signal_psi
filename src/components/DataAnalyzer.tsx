"use client";

import React, { useState, useMemo, useEffect } from "react";
import { BarChart3, TrendingUp, Cpu, Hash, X } from "lucide-react";

interface DataAnalyzerProps {
  data: any[][];
  onSelectGroupFilter: (columnName: string | null, value: any | null) => void;
  selectedFilter: { columnName: string; value: any } | null;
}

interface GroupSummary {
  name: string;
  count: number;
  percentage: number;
  averageValue?: number | null;
}

export default function DataAnalyzer({ data, onSelectGroupFilter, selectedFilter }: DataAnalyzerProps) {
  const headers = useMemo(() => (data && data.length > 0 ? data[0] : []), [data]);
  const rows = useMemo(() => (data && data.length > 1 ? data.slice(1) : []), [data]);

  // 분석 대상 컬럼 설정 상태
  const [groupByIndex, setGroupByIndex] = useState<number>(-1);
  const [calcByIndex, setCalcByIndex] = useState<number>(-1);

  // 자동 컬럼 추천 휴리스틱
  useEffect(() => {
    if (headers.length === 0 || rows.length === 0) return;

    // 1. 그룹화 컬럼 자동 선택 (가장 먼저 만나는 '장치', '기기', '열차', '구분', '상태', '명' 등이 포함된 열 탐색)
    let bestGroupIdx = 0;
    const keywords = ["장치", "기기", "설비", "열차", "역명", "상태", "구분", "종류", "device", "type", "status", "train"];
    
    for (let i = 0; i < headers.length; i++) {
      const headerStr = String(headers[i]).toLowerCase();
      if (keywords.some(k => headerStr.includes(k))) {
        bestGroupIdx = i;
        break;
      }
    }
    setGroupByIndex(bestGroupIdx);

    // 2. 수치형 계산 컬럼 자동 선택 (숫자가 주로 많이 포함된 열 탐색)
    let bestCalcIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      if (i === bestGroupIdx) continue;
      
      // 샘플 행에서 숫자가 있는지 체크
      const numericCount = rows.slice(0, 10).filter(row => {
        const val = row[i];
        return val !== undefined && val !== null && val !== "" && !isNaN(Number(val));
      }).length;

      // 샘플 중 50% 이상이 숫자이면 수치형 열로 선정
      if (numericCount >= Math.min(5, rows.length * 0.5)) {
        bestCalcIdx = i;
        break;
      }
    }
    setCalcByIndex(bestCalcIdx);
  }, [headers, rows]);

  // 그룹화 데이터 연산
  const groupSummaries = useMemo((): GroupSummary[] => {
    if (groupByIndex === -1 || rows.length === 0) return [];

    const groupsMap = new Map<string, { count: number; sum: number; validNumCount: number }>();

    rows.forEach((row) => {
      let groupVal = row[groupByIndex];
      // 비어있거나 null 이면 '미지정' 처리
      if (groupVal === undefined || groupVal === null || String(groupVal).trim() === "") {
        groupVal = "(미지정)";
      }
      const groupKey = String(groupVal).trim();

      // 수치형 값 계산 준비
      let numVal = 0;
      let isNumeric = false;
      if (calcByIndex !== -1) {
        const rawVal = row[calcByIndex];
        if (rawVal !== undefined && rawVal !== null && rawVal !== "" && !isNaN(Number(rawVal))) {
          numVal = Number(rawVal);
          isNumeric = true;
        }
      }

      const existing = groupsMap.get(groupKey) || { count: 0, sum: 0, validNumCount: 0 };
      groupsMap.set(groupKey, {
        count: existing.count + 1,
        sum: existing.sum + numVal,
        validNumCount: existing.validNumCount + (isNumeric ? 1 : 0),
      });
    });

    const totalRows = rows.length;
    const summaries: GroupSummary[] = [];

    groupsMap.forEach((meta, name) => {
      const averageValue = meta.validNumCount > 0 ? meta.sum / meta.validNumCount : null;
      summaries.push({
        name,
        count: meta.count,
        percentage: (meta.count / totalRows) * 100,
        averageValue,
      });
    });

    // 건수(Count) 기준으로 내림차순 정렬
    return summaries.sort((a, b) => b.count - a.count);
  }, [rows, groupByIndex, calcByIndex]);

  // KPI 지표 요약
  const stats = useMemo(() => {
    if (groupSummaries.length === 0) return null;

    const totalGroups = groupSummaries.length;
    const dominant = groupSummaries[0]; // 정렬 상태이므로 첫 번째가 최다 빈도

    let overallAvg: number | null = null;
    if (calcByIndex !== -1 && rows.length > 0) {
      let sum = 0;
      let count = 0;
      rows.forEach(r => {
        const val = r[calcByIndex];
        if (val !== undefined && val !== null && val !== "" && !isNaN(Number(val))) {
          sum += Number(val);
          count++;
        }
      });
      if (count > 0) overallAvg = sum / count;
    }

    return {
      totalGroups,
      dominantName: dominant.name,
      dominantPct: dominant.percentage.toFixed(1),
      dominantCount: dominant.count,
      overallAvg: overallAvg !== null ? overallAvg.toFixed(2) : null,
    };
  }, [groupSummaries, rows, calcByIndex]);

  const handleGroupClick = (groupName: string) => {
    if (groupByIndex === -1) return;
    const colName = String(headers[groupByIndex]);
    
    // 이미 같은 필터가 지정되어 있으면 해제
    if (selectedFilter && selectedFilter.columnName === colName && selectedFilter.value === groupName) {
      onSelectGroupFilter(null, null);
    } else {
      onSelectGroupFilter(colName, groupName);
    }
  };

  const handleClearFilter = () => {
    onSelectGroupFilter(null, null);
  };

  if (headers.length === 0 || rows.length === 0) return null;

  return (
    <div className="space-y-6 bg-white dark:bg-slate-900 p-6 rounded-2xl shadow-sm border border-slate-200 dark:border-slate-800">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-100 dark:border-slate-800">
        <div className="flex items-center space-x-2">
          <BarChart3 className="w-5 h-5 text-blue-600 dark:text-blue-400" />
          <h3 className="text-lg font-bold text-slate-800 dark:text-slate-100">데이터 그룹 분석 및 요약</h3>
        </div>

        {/* 컬럼 가공 셀렉트 박스 */}
        <div className="flex flex-wrap items-center gap-3 text-sm">
          <div className="flex items-center space-x-1.5">
            <span className="text-slate-500 dark:text-slate-400 font-medium">분석 기준(그룹):</span>
            <select
              value={groupByIndex}
              onChange={(e) => {
                setGroupByIndex(Number(e.target.value));
                handleClearFilter(); // 기준 열이 바뀌면 필터 리셋
              }}
              className="border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 rounded px-2.5 py-1 text-slate-700 dark:text-slate-300 font-semibold focus:outline-none focus:ring-1 focus:ring-blue-500"
            >
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h !== undefined && h !== null ? String(h) : `열 ${i + 1}`}
                </option>
              ))}
            </select>
          </div>

          <div className="flex items-center space-x-1.5">
            <span className="text-slate-500 dark:text-slate-400 font-medium">수치 연산 대상:</span>
            <select
              value={calcByIndex}
              onChange={(e) => setCalcByIndex(Number(e.target.value))}
              className="border border-slate-200 dark:border-slate-700 bg-slate-50 dark:bg-slate-950 rounded px-2.5 py-1 text-slate-700 dark:text-slate-300 font-semibold focus:outline-none focus:ring-1 focus:ring-blue-500"
            >
              <option value={-1}>(없음)</option>
              {headers.map((h, i) => (
                <option key={i} value={i}>
                  {h !== undefined && h !== null ? String(h) : `열 ${i + 1}`}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* KPI 카드 섹션 */}
      {stats && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div className="p-4 bg-blue-50/50 dark:bg-blue-950/20 border border-blue-100/50 dark:border-blue-900/30 rounded-xl flex items-center space-x-4">
            <div className="p-3 bg-blue-100 dark:bg-blue-900 text-blue-600 dark:text-blue-400 rounded-lg">
              <Cpu className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">전체 그룹 개수</p>
              <p className="text-xl font-bold text-slate-800 dark:text-slate-100 mt-0.5">
                {stats.totalGroups}개
              </p>
            </div>
          </div>

          <div className="p-4 bg-emerald-50/50 dark:bg-emerald-950/20 border border-emerald-100/50 dark:border-emerald-900/30 rounded-xl flex items-center space-x-4">
            <div className="p-3 bg-emerald-100 dark:bg-emerald-900 text-emerald-600 dark:text-emerald-400 rounded-lg">
              <TrendingUp className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">최다 빈도 항목</p>
              <p className="text-xl font-bold text-slate-800 dark:text-slate-100 mt-0.5 truncate max-w-[180px]" title={stats.dominantName}>
                {stats.dominantName} <span className="text-xs font-normal text-slate-500">({stats.dominantPct}%)</span>
              </p>
            </div>
          </div>

          <div className="p-4 bg-amber-50/50 dark:bg-amber-950/20 border border-amber-100/50 dark:border-amber-900/30 rounded-xl flex items-center space-x-4">
            <div className="p-3 bg-amber-100 dark:bg-amber-900 text-amber-600 dark:text-amber-400 rounded-lg">
              <Hash className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs text-slate-500 dark:text-slate-400 font-medium">
                {calcByIndex !== -1 && headers[calcByIndex] ? `${String(headers[calcByIndex])} 평균` : "수치 데이터"}
              </p>
              <p className="text-xl font-bold text-slate-800 dark:text-slate-100 mt-0.5">
                {stats.overallAvg !== null ? stats.overallAvg : "N/A"}
              </p>
            </div>
          </div>
        </div>
      )}

      {/* 필터 활성화 배너 */}
      {selectedFilter && (
        <div className="flex items-center justify-between p-3 bg-blue-50 dark:bg-blue-950/40 border border-blue-200 dark:border-blue-900/60 rounded-xl">
          <div className="text-sm text-blue-700 dark:text-blue-300 font-medium">
            필터 적용 중: <span className="font-semibold">{selectedFilter.columnName}</span> = &quot;{selectedFilter.value}&quot;
          </div>
          <button
            onClick={handleClearFilter}
            className="p-1 text-blue-500 hover:text-blue-700 dark:hover:text-blue-300 hover:bg-blue-100 dark:hover:bg-blue-900 rounded-lg transition-colors"
            title="필터 해제"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      )}

      {/* 분석 요약 리스트 */}
      <div className="border border-slate-200 dark:border-slate-800 rounded-xl overflow-hidden">
        <div className="max-h-[300px] overflow-y-auto">
          <table className="w-full text-left text-sm border-collapse">
            <thead className="bg-slate-50 dark:bg-slate-800 text-xs font-semibold text-slate-500 dark:text-slate-400 uppercase border-b border-slate-200 dark:border-slate-700 sticky top-0 z-10">
              <tr>
                <th className="px-5 py-3">그룹명 ({groupByIndex !== -1 ? String(headers[groupByIndex]) : ""})</th>
                <th className="px-5 py-3 w-32 text-right">빈도 수 (행 수)</th>
                <th className="px-5 py-3 w-32 text-right">비율 (%)</th>
                {calcByIndex !== -1 && (
                  <th className="px-5 py-3 w-40 text-right">
                    {String(headers[calcByIndex])} 평균
                  </th>
                )}
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100 dark:divide-slate-800">
              {groupSummaries.map((group, i) => {
                const colName = groupByIndex !== -1 ? String(headers[groupByIndex]) : "";
                const isSelected = selectedFilter?.columnName === colName && selectedFilter?.value === group.name;

                return (
                  <tr
                    key={i}
                    onClick={() => handleGroupClick(group.name)}
                    className={`cursor-pointer hover:bg-blue-50/30 dark:hover:bg-blue-950/10 transition-colors ${
                      isSelected ? "bg-blue-50 dark:bg-blue-950/30 font-medium" : ""
                    }`}
                  >
                    <td className="px-5 py-3 text-slate-700 dark:text-slate-300 flex items-center space-x-2">
                      <div className={`w-2.5 h-2.5 rounded-full ${isSelected ? "bg-blue-600" : "bg-slate-300 dark:bg-slate-600"}`} />
                      <span>{group.name}</span>
                    </td>
                    <td className="px-5 py-3 text-right text-slate-600 dark:text-slate-400">
                      {group.count.toLocaleString()}개
                    </td>
                    <td className="px-5 py-3 text-right text-slate-600 dark:text-slate-400 font-mono">
                      {group.percentage.toFixed(1)}%
                    </td>
                    {calcByIndex !== -1 && (
                      <td className="px-5 py-3 text-right text-slate-700 dark:text-slate-300 font-mono">
                        {group.averageValue !== null && group.averageValue !== undefined
                          ? group.averageValue.toFixed(2)
                          : "-"}
                      </td>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
      <p className="text-xs text-slate-400 dark:text-slate-500">
        💡 리스트의 행을 클릭하면 해당 데이터를 기준으로 하단 테이블이 필터링됩니다.
      </p>
    </div>
  );
}
